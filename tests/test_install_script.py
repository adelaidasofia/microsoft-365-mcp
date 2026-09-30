"""install.sh is the path every new install takes, so it needs a gate.

Hermetic: `claude` and the venv interpreter are shims on PATH, so nothing here
touches the real Claude Code config or the network.

Four things are worth pinning, and nothing else:
  1. valid input produces the exact `claude mcp add` argv (a lost `--` or a
     dropped `-s user` is invisible until a person tries to use it)
  2. a client ID that cannot work registers nothing and exits non-zero. An
     Entra client ID is always a GUID, and pasting the app's display name or
     its Object ID instead is the common mistake -- it otherwise surfaces much
     later as an opaque AADSTS code with nothing pointing back at the typo
  3. re-running heals rather than fails
  4. inside the Claude desktop app's own Code tab there is often no `claude`
     on PATH, only CLAUDE_CODE_EXECPATH pointing at the app's bundled Claude
     Code -- the installer must find it there, PATH must still win when both
     exist, and an arbitrary binary must never be trusted just because the
     variable happens to be set

Subprocess calls pass argument lists, never shell strings.
"""

import os
import shutil
import stat
import subprocess
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
INSTALL_SH = REPO_ROOT / "install.sh"

CLIENT_ID = "1a2b3c4d-5e6f-7890-abcd-ef1234567890"

pytestmark = pytest.mark.skipif(not INSTALL_SH.exists(), reason="no install.sh here")


def _exe(path: Path, body: str) -> None:
    path.write_text(body)
    path.chmod(path.stat().st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)


@pytest.fixture
def box(tmp_path):
    repo = tmp_path / "repo"
    repo.mkdir()
    shutil.copy(INSTALL_SH, repo / "install.sh")
    (repo / "server.py").write_text("# stub\n")
    (repo / "requirements.txt").write_text("")

    bindir = tmp_path / "bin"
    bindir.mkdir()
    log = tmp_path / "calls.log"
    log.touch()
    # `mcp remove` exits 1 to mimic "not registered yet", which must be tolerated.
    _exe(
        bindir / "claude",
        f'#!/bin/sh\necho "$*" >> "{log}"\ncase "$2" in remove) exit 1 ;; esac\nexit 0\n',
    )
    # A pre-made venv interpreter skips venv creation, pip and the import check,
    # so this runs in milliseconds and installs nothing.
    venv_bin = repo / ".venv" / "bin"
    venv_bin.mkdir(parents=True)
    _exe(venv_bin / "python", "#!/bin/sh\nexit 0\n")

    return {"repo": repo, "bin": bindir, "log": log, "root": tmp_path}


def _run(box, **creds):
    env = dict(os.environ, PATH=f"{box['bin']}{os.pathsep}{os.environ['PATH']}", NO_COLOR="1")
    env.pop("M365_CLIENT_ID", None)
    # Never let this machine's own ambient CLAUDE_CODE_EXECPATH (set when this
    # very suite is run from inside Claude Code) leak into a test that isn't
    # exercising it -- each test below sets it explicitly if it wants it.
    env.pop("CLAUDE_CODE_EXECPATH", None)
    env.update(creds)
    proc = subprocess.run(
        ["bash", str(box["repo"] / "install.sh")],
        capture_output=True, text=True, env=env, input="", timeout=120,
    )
    return proc, box["log"].read_text()


def test_valid_client_id_registers_the_expected_command(box):
    proc, calls = _run(box, M365_CLIENT_ID=CLIENT_ID)
    assert proc.returncode == 0, proc.stdout + proc.stderr

    add = [ln for ln in calls.splitlines() if ln.startswith("mcp add")]
    assert len(add) == 1, f"expected one `mcp add`, got {calls!r}"
    argv = add[0]

    assert "mcp add microsoft-365" in argv
    assert "-s user" in argv, "must register at user scope, not just this project"
    assert f"-e M365_CLIENT_ID={CLIENT_ID}" in argv
    assert " -- " in argv, "without `--` the interpreter path parses as a flag"
    assert argv.rstrip().endswith("server.py")


@pytest.mark.parametrize(
    "value, why",
    [
        ("My AI Brain", "the app's display name"),
        ("not-a-guid", "a plausible-looking non-GUID"),
        ("1a2b3c4d5e6f7890abcdef1234567890", "a GUID with the dashes stripped"),
    ],
)
def test_a_client_id_that_cannot_work_registers_nothing(box, value, why):
    proc, calls = _run(box, M365_CLIENT_ID=value)
    assert proc.returncode != 0, f"{why}: should have failed loudly"
    assert "mcp add" not in calls, f"{why}: must not register an unusable client"


def test_rerunning_heals_instead_of_failing(box):
    _run(box, M365_CLIENT_ID=CLIENT_ID)
    proc, calls = _run(box, M365_CLIENT_ID=CLIENT_ID)

    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert calls.count("mcp remove microsoft-365") == 2
    assert calls.count("mcp add microsoft-365") == 2


# --------------------------------------------------------- CLAUDE_CODE_EXECPATH
#
# Inside the Claude desktop app's own Code tab, `claude` is frequently not on
# PATH at all -- only CLAUDE_CODE_EXECPATH, pointing at the app's own bundled
# Claude Code binary, somewhere with a space in the path (e.g. macOS's
# "Application Support"). These tests must run on a machine that *does* have
# a real `claude` on PATH (this one does), so PATH is sealed down to a
# handful of fixed system directories -- no plugin shims, no Homebrew, no
# node-managed installs -- none of which macOS ships a `claude` in.

REAL_PYTHON = sys.executable
_SEALED_SYSTEM_DIRS = ("/usr/bin", "/bin", "/usr/sbin", "/sbin")


def _sealed_env(bindir: Path, **extra: str) -> dict:
    env = {
        "PATH": os.pathsep.join([str(bindir), *_SEALED_SYSTEM_DIRS]),
        "HOME": os.environ.get("HOME", str(bindir)),
        "NO_COLOR": "1",
    }
    env.update(extra)
    return env


def _assert_sealed_path_has_no_claude(env: dict) -> None:
    # Positive control, run before every sealed test: prove THIS PATH, on
    # THIS machine, truly cannot resolve a real `claude`. Without this, a
    # sealed PATH that quietly still found one would pass every test below
    # for the wrong reason -- and this machine normally has a real `claude`.
    probe = subprocess.run(
        ["bash", "-c", "command -v claude"],
        capture_output=True, text=True, env=env, timeout=30,
    )
    assert probe.returncode != 0, f"sealed PATH still resolves a real claude: {probe.stdout!r}"


@pytest.fixture
def sealed_box(tmp_path):
    """Like `box`, but PATH cannot resolve a real `claude` anywhere.

    git and python3 must still work, since install.sh checks both before it
    ever gets to resolving Claude Code. Real git lives in the sealed system
    dirs already; python3 is shimmed to exec this very test run's own
    interpreter (>=3.11, this repo's own floor per pyproject.toml), so the
    fixture does not depend on which system python3 -- if any -- happens to
    satisfy install.sh's own >=3.10 gate on the machine running the tests.
    """
    repo = tmp_path / "repo"
    repo.mkdir()
    shutil.copy(INSTALL_SH, repo / "install.sh")
    (repo / "server.py").write_text("# stub\n")
    (repo / "requirements.txt").write_text("")

    venv_bin = repo / ".venv" / "bin"
    venv_bin.mkdir(parents=True)
    _exe(venv_bin / "python", "#!/bin/sh\nexit 0\n")

    bindir = tmp_path / "bin"
    bindir.mkdir()
    _exe(bindir / "python3", f'#!/bin/sh\nexec "{REAL_PYTHON}" "$@"\n')

    log = tmp_path / "calls.log"
    log.touch()

    return {"repo": repo, "bin": bindir, "log": log, "root": tmp_path}


def _desktop_shim(path: Path, log: Path, *, reports_claude_code: bool) -> None:
    """A recording stand-in for the binary CLAUDE_CODE_EXECPATH points at.

    Answers --version first, like the real Claude Code binary would, then
    falls through to the same call-recording behavior as the `claude` PATH
    shim in `box`, so a test can assert what -- if anything -- it was asked
    to do.
    """
    version_line = "2.1.281 (Claude Code)" if reports_claude_code else "impostor-tool 9.9.9"
    _exe(
        path,
        f'#!/bin/sh\n'
        f'if [ "$1" = "--version" ]; then echo "{version_line}"; exit 0; fi\n'
        f'echo "$*" >> "{log}"\n'
        f'case "$2" in remove) exit 1 ;; esac\n'
        f'exit 0\n',
    )


def _run_sealed(box, **env_extra):
    env = _sealed_env(box["bin"], **env_extra)
    _assert_sealed_path_has_no_claude(env)
    proc = subprocess.run(
        ["bash", str(box["repo"] / "install.sh")],
        capture_output=True, text=True, env=env, input="", timeout=120,
    )
    return proc, box["log"].read_text()


def test_execpath_registers_when_claude_not_on_path(sealed_box):
    # The desktop app's own bundled Claude Code lives somewhere with a SPACE
    # in the path (mirrors macOS's "Application Support") and is never
    # itself on PATH -- only reachable via CLAUDE_CODE_EXECPATH.
    desktop_dir = sealed_box["root"] / "Application Support" / "Claude"
    desktop_dir.mkdir(parents=True)
    execpath = desktop_dir / "claude"
    _desktop_shim(execpath, sealed_box["log"], reports_claude_code=True)

    proc, calls = _run_sealed(
        sealed_box, M365_CLIENT_ID=CLIENT_ID, CLAUDE_CODE_EXECPATH=str(execpath)
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr

    add = [ln for ln in calls.splitlines() if ln.startswith("mcp add")]
    assert len(add) == 1, f"expected one `mcp add`, got {calls!r}"
    argv = add[0]

    assert "mcp add microsoft-365" in argv
    assert "-s user" in argv, "must register at user scope, not just this project"
    assert f"-e M365_CLIENT_ID={CLIENT_ID}" in argv
    assert " -- " in argv, "without `--` the interpreter path parses as a flag"
    assert argv.rstrip().endswith("server.py")


def test_no_claude_anywhere_dies_loudly(sealed_box):
    """Negative control: neither PATH nor CLAUDE_CODE_EXECPATH has anything."""
    proc, calls = _run_sealed(sealed_box, M365_CLIENT_ID=CLIENT_ID)
    assert proc.returncode != 0
    assert "Claude Code is not installed" in proc.stdout + proc.stderr
    assert "mcp add" not in calls


def test_execpath_binary_that_is_not_claude_code_is_rejected(sealed_box):
    """An arbitrary executable at CLAUDE_CODE_EXECPATH must never be trusted
    just because the variable happens to be set -- only one that identifies
    itself as Claude Code in --version output is."""
    impostor = sealed_box["root"] / "impostor"
    _desktop_shim(impostor, sealed_box["log"], reports_claude_code=False)

    proc, calls = _run_sealed(
        sealed_box, M365_CLIENT_ID=CLIENT_ID, CLAUDE_CODE_EXECPATH=str(impostor)
    )
    assert proc.returncode != 0, "must never trust an arbitrary binary's --version claim"
    assert "mcp add" not in calls
    assert "Claude Code is not installed" in proc.stdout + proc.stderr


def test_path_claude_wins_over_execpath(box):
    """PATH wins even when CLAUDE_CODE_EXECPATH also points at something that
    would otherwise qualify -- the execpath binary must not even be asked."""
    execpath_log = box["root"] / "execpath_calls.log"
    execpath_log.touch()
    execpath = box["root"] / "desktop-execpath" / "claude"
    execpath.parent.mkdir()
    _desktop_shim(execpath, execpath_log, reports_claude_code=True)

    proc, calls = _run(box, M365_CLIENT_ID=CLIENT_ID, CLAUDE_CODE_EXECPATH=str(execpath))
    assert proc.returncode == 0, proc.stdout + proc.stderr

    add = [ln for ln in calls.splitlines() if ln.startswith("mcp add")]
    assert len(add) == 1, f"expected one `mcp add`, got {calls!r}"
    assert execpath_log.read_text() == "", "the execpath shim must never be invoked when claude is on PATH"
