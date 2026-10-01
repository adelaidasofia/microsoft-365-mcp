"""install.sh is the path every new install takes, so it needs a gate.

Hermetic: `claude` and the venv interpreter are shims on PATH, so nothing here
touches the real Claude Code config or the network. That is enforced, not hoped
for: every installer run gets a throwaway HOME and CLAUDE_CONFIG_DIR, and
refuses to start if a real `claude` could be reached from the PATH it was given
(see "The harness itself" at the bottom).

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
import time
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
INSTALL_SH = REPO_ROOT / "install.sh"
BASH = shutil.which("bash") or "bash"

CLIENT_ID = "1a2b3c4d-5e6f-7890-abcd-ef1234567890"

pytestmark = pytest.mark.skipif(not INSTALL_SH.exists(), reason="no install.sh here")


def _exe(path: Path, body: str) -> None:
    path.write_text(body)
    path.chmod(path.stat().st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)


# ---------------------------------------------------------------------------
# Keeping the installer away from the developer's real Claude Code
#
# install.sh ends by running `claude mcp add ... -s user`. If a real `claude` is
# the one that answers, that writes a connector into the developer's own
# ~/.claude.json, pointing at a pytest temp directory that is gone by the next
# run, and every Claude Code session on the machine then shows it as failing. A
# fixture that forgets to put its own fake `claude` first is all it takes.
#
# So two layers, neither of which depends on a fixture remembering anything:
#   * every run gets a HOME and CLAUDE_CONFIG_DIR inside its own tmp dir, so
#     even a real `claude` reached by mistake could only write a throwaway
#     config;
#   * every run first checks that `claude` cannot resolve to anything outside
#     its tmp dir, and refuses to start if it can.
# ---------------------------------------------------------------------------


def _throwaway_config(root: Path) -> dict:
    """HOME, CLAUDE_CONFIG_DIR and TMPDIR for one installer run, all inside `root`.

    TMPDIR is not about Claude Code: it keeps the installer's own scratch files
    inside the test's tmp dir too, where a test can see whether they are cleaned up.
    """
    home = root / "home"
    config = root / "claude-config"
    tmp = root / "tmp"
    for d in (home, config, tmp):
        d.mkdir(exist_ok=True)
    return {"HOME": str(home), "CLAUDE_CONFIG_DIR": str(config), "TMPDIR": str(tmp)}


def _command_found_on(env: dict, name: str) -> str:
    """Where `command -v <name>` lands under `env`: the lookup install.sh makes."""
    probe = subprocess.run(
        [BASH, "-c", 'command -v "$1"', "bash", name],
        capture_output=True, text=True, env=env, timeout=30,
    )
    return probe.stdout.strip() if probe.returncode == 0 else ""


def _assert_nothing_real_is_reachable(env: dict, root: Path, *, shim_expected: bool) -> None:
    """Refuse to run the installer unless it can only touch things inside `root`.

    `claude` must resolve to nothing at all (`shim_expected=False`: the tests that
    need `claude` to be absent from PATH also depend on this -- without it, a PATH
    that quietly still found one would pass every one of them for the wrong
    reason), or to a shim inside `root`. HOME and CLAUDE_CONFIG_DIR must point
    inside `root` as well, and so must CLAUDE_CODE_EXECPATH when it is set: it is
    the installer's second way of reaching a binary, and one outside `root` is run
    just as a PATH one would be.
    """
    root = root.resolve()
    found = _command_found_on(env, "claude")
    if found:
        assert shim_expected, f"expected no claude on this PATH, but it resolves to {found}"
        assert Path(found).resolve().is_relative_to(root), (
            f"a real claude is reachable on the fixture PATH: {found}. install.sh would "
            f"run `claude mcp add ... -s user` against it for real. PATH={env.get('PATH')}"
        )
    for key in ("HOME", "CLAUDE_CONFIG_DIR"):
        assert Path(env[key]).resolve().is_relative_to(root), (
            f"{key} must point inside this test's own tmp dir, got {env[key]!r}"
        )
    execpath = env.get("CLAUDE_CODE_EXECPATH")
    if execpath:
        assert Path(execpath).resolve().is_relative_to(root), (
            f"CLAUDE_CODE_EXECPATH must point inside this test's own tmp dir, got {execpath!r}: "
            f"install.sh runs whatever it names with --version, and as `claude mcp add` if it answers"
        )


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


def _box_env(box, **creds):
    env = dict(os.environ, PATH=f"{box['bin']}{os.pathsep}{os.environ['PATH']}", NO_COLOR="1")
    env.pop("M365_CLIENT_ID", None)
    # Never let this machine's own ambient CLAUDE_CODE_EXECPATH (set when this
    # very suite is run from inside Claude Code) leak into a test that isn't
    # exercising it -- each test below sets it explicitly if it wants it.
    env.pop("CLAUDE_CODE_EXECPATH", None)
    env.update(_throwaway_config(box["root"]))
    env.update(creds)
    return env


def _run_installer(fx, env, *, shim_expected):
    """The one place install.sh is started: the guard first, then the run.

    Both runners below go through here, so a run cannot skip the guard by
    forgetting it; test_every_runner_checks_the_guard_... shows that it cannot.
    """
    _assert_nothing_real_is_reachable(env, fx["root"], shim_expected=shim_expected)
    proc = subprocess.run(
        [BASH, str(fx["repo"] / "install.sh")],
        capture_output=True, text=True, env=env, input="", timeout=120,
    )
    return proc, fx["log"].read_text()


def _run(box, **creds):
    return _run_installer(box, _box_env(box, **creds), shim_expected=True)


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


def _sealed_env(box, *, system_dirs=_SEALED_SYSTEM_DIRS, **extra: str) -> dict:
    env = {
        "PATH": os.pathsep.join([str(box["bin"]), *system_dirs]),
        "NO_COLOR": "1",
        **_throwaway_config(box["root"]),
    }
    env.update(extra)
    return env


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


def _desktop_shim(
    path: Path,
    log: Path,
    *,
    reports_claude_code: bool,
    version_exit: int = 0,
    version_on: str = "stdout",
    add_exit: int = 0,
) -> None:
    """A recording stand-in for the binary CLAUDE_CODE_EXECPATH points at.

    Answers --version first, like the real Claude Code binary would, then
    falls through to the same call-recording behavior as the `claude` PATH
    shim in `box`, so a test can assert what -- if anything -- it was asked
    to do.

    `version_exit` and `version_on` shape how it answers --version (which stream
    it prints on, and what it exits with afterwards); `add_exit` is what
    `mcp add` exits with.
    """
    version_line = "2.1.281 (Claude Code)" if reports_claude_code else "impostor-tool 9.9.9"
    to_stderr = " >&2" if version_on == "stderr" else ""
    _exe(
        path,
        f'#!/bin/sh\n'
        f'if [ "$1" = "--version" ]; then echo "{version_line}"{to_stderr}; exit {version_exit}; fi\n'
        f'echo "$*" >> "{log}"\n'
        f'case "$2" in remove) exit 1 ;; add) exit {add_exit} ;; esac\n'
        f'exit 0\n',
    )


HANG_SECONDS = 60


def _hanging_claude(path: Path) -> None:
    """A binary that never answers --version (it sleeps for HANG_SECONDS).

    `exec`, so the binary is the sleeper itself. A wrapper that starts a child
    and leaves it holding the output open is a different case, with its own
    stand-in below.
    """
    sleeper = shutil.which("sleep")
    _exe(path, f'#!/bin/sh\nexec "{sleeper}" {HANG_SECONDS}\n')


def _claude_with_a_lingering_child(path: Path, log: Path, *, answers: bool) -> None:
    """A binary that starts a long-lived child when it is asked for its version.

    The child inherits the binary's stdout and stderr and keeps them open for
    HANG_SECONDS, as a wrapper script or a helper process that outlives its
    parent does. With `answers` the binary prints its banner and exits at once:
    it has said what it is, and only its child is slow. Without, it waits for the
    child, so it cannot answer inside the installer's ten seconds. Either way the
    installer must not wait for the child: a shell waits for the process it
    started, not for everything that inherited the output it was given.
    """
    sleeper = shutil.which("sleep")
    if answers:
        version = f'  echo "2.1.281 (Claude Code)"\n  "{sleeper}" {HANG_SECONDS} &\n  exit 0\n'
    else:
        version = f'  "{sleeper}" {HANG_SECONDS} &\n  wait\n  echo "2.1.281 (Claude Code)"\n  exit 0\n'
    _exe(
        path,
        f"""#!/bin/sh
echo "$*" >> "{log}"
if [ "$1" = "--version" ]; then
{version}fi
case "$2" in
  remove) exit 1 ;;
esac
exit 0
""",
    )


def _probe_leftovers(box) -> list:
    """Scratch files the version probe left behind in the run's TMPDIR."""
    return sorted(p.name for p in (box["root"] / "tmp").glob("claude-probe.*"))


def _run_sealed(box, *, system_dirs=_SEALED_SYSTEM_DIRS, **env_extra):
    env = _sealed_env(box, system_dirs=system_dirs, **env_extra)
    return _run_installer(box, env, shim_expected=False)


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
    # The heal-by-remove step goes through the same binary as the add.
    assert "mcp remove microsoft-365 -s user" in calls


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


def test_execpath_that_prints_its_banner_and_then_exits_non_zero_is_still_used(sealed_box):
    """What a binary says it is and how it exits are separate questions. A
    binary that prints "Claude Code" and then exits non-zero is still Claude
    Code; the exit status must not decide it."""
    execpath = sealed_box["root"] / "claude-exits-7"
    _desktop_shim(execpath, sealed_box["log"], reports_claude_code=True, version_exit=7)

    proc, calls = _run_sealed(
        sealed_box, M365_CLIENT_ID=CLIENT_ID, CLAUDE_CODE_EXECPATH=str(execpath)
    )

    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "mcp add microsoft-365" in calls


def test_execpath_that_says_what_it_is_on_stderr_is_still_used(sealed_box):
    """Some CLIs print their version banner on stderr. Reading only stdout would
    reject a real Claude Code for it."""
    execpath = sealed_box["root"] / "claude-says-it-on-stderr"
    _desktop_shim(execpath, sealed_box["log"], reports_claude_code=True, version_on="stderr")

    proc, calls = _run_sealed(
        sealed_box, M365_CLIENT_ID=CLIENT_ID, CLAUDE_CODE_EXECPATH=str(execpath)
    )

    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "mcp add microsoft-365" in calls


def test_execpath_that_hangs_on_version_is_given_up_on(sealed_box):
    """A binary that never answers --version must not hang the installer. It is
    cut off after ten seconds and counts as not being Claude Code.

    The binary here hangs for HANG_SECONDS, long enough that an installer with
    no bound finishes much later than one with it; the elapsed-time check is
    what tells the two apart, because a hang that eventually ends and prints
    nothing is rejected either way."""
    execpath = sealed_box["root"] / "claude-hangs"
    _hanging_claude(execpath)
    env = _sealed_env(sealed_box, M365_CLIENT_ID=CLIENT_ID, CLAUDE_CODE_EXECPATH=str(execpath))
    if not _command_found_on(env, "perl"):
        pytest.skip("no perl on the sealed PATH, so there is no bound to test")

    started = time.monotonic()
    proc, calls = _run_sealed(
        sealed_box, M365_CLIENT_ID=CLIENT_ID, CLAUDE_CODE_EXECPATH=str(execpath)
    )
    elapsed = time.monotonic() - started

    assert elapsed < HANG_SECONDS - 20, (
        f"the --version probe was not bounded: the installer took {elapsed:.0f}s "
        f"against a binary that hangs for {HANG_SECONDS}s"
    )
    assert proc.returncode != 0
    assert "mcp add" not in calls
    assert "Claude Code is not installed" in proc.stdout + proc.stderr


def test_a_child_that_keeps_the_output_open_does_not_stretch_the_wait(sealed_box):
    """The binary answered and exited; a helper it started still holds the output
    it was given. The installer waits for the binary it ran, not for everything
    that inherited the output, so the answer is used at once."""
    execpath = sealed_box["root"] / "claude-leaves-a-child"
    _claude_with_a_lingering_child(execpath, sealed_box["log"], answers=True)

    started = time.monotonic()
    proc, calls = _run_sealed(
        sealed_box, M365_CLIENT_ID=CLIENT_ID, CLAUDE_CODE_EXECPATH=str(execpath)
    )
    elapsed = time.monotonic() - started

    assert elapsed < HANG_SECONDS - 20, (
        f"the installer waited for a child of the binary it asked: it took {elapsed:.0f}s, "
        f"and the child lives for {HANG_SECONDS}s"
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "mcp add microsoft-365" in calls
    assert _probe_leftovers(sealed_box) == []


def test_a_wrapper_whose_child_outlives_the_bound_is_still_cut_off_on_time(sealed_box):
    """The ten second bound has to hold for a binary that is a wrapper around a
    child it waits for. The alarm ends the wrapper; whatever the child still has
    open must not stretch the wait past it."""
    execpath = sealed_box["root"] / "claude-wrapper"
    _claude_with_a_lingering_child(execpath, sealed_box["log"], answers=False)
    env = _sealed_env(sealed_box, M365_CLIENT_ID=CLIENT_ID, CLAUDE_CODE_EXECPATH=str(execpath))
    if not _command_found_on(env, "perl"):
        pytest.skip("no perl on the sealed PATH, so there is no bound to test")

    started = time.monotonic()
    proc, calls = _run_sealed(
        sealed_box, M365_CLIENT_ID=CLIENT_ID, CLAUDE_CODE_EXECPATH=str(execpath)
    )
    elapsed = time.monotonic() - started

    assert elapsed < HANG_SECONDS - 20, (
        f"the --version probe was not bounded: the installer took {elapsed:.0f}s "
        f"against a wrapper whose child lives for {HANG_SECONDS}s"
    )
    assert proc.returncode != 0
    assert "mcp add" not in calls
    assert "Claude Code is not installed" in proc.stdout + proc.stderr
    assert _probe_leftovers(sealed_box) == []


def test_execpath_is_still_used_where_perl_is_missing(sealed_box):
    """The ten-second bound needs perl. Without it the question is asked
    unbounded rather than not asked: no perl must never mean no fallback."""
    # A PATH of its own: just the few tools install.sh needs, and not perl. The
    # version probe writes its answer to a scratch file, so mktemp and rm are among
    # them: the point is the probe without perl, not the probe without its tools.
    tools = sealed_box["root"] / "tools"
    tools.mkdir()
    for name in ("git", "grep", "dirname", "cat", "mktemp", "rm"):
        real = shutil.which(name)
        if real is None:
            pytest.skip(f"{name} unavailable, cannot build a PATH without perl")
        (tools / name).symlink_to(real)
    execpath = sealed_box["root"] / "claude"
    _desktop_shim(execpath, sealed_box["log"], reports_claude_code=True)
    extra = {"M365_CLIENT_ID": CLIENT_ID, "CLAUDE_CODE_EXECPATH": str(execpath)}
    # Control: this PATH really has no perl, so the unbounded branch is the one that runs.
    assert _command_found_on(_sealed_env(sealed_box, system_dirs=(str(tools),), **extra), "perl") == ""

    proc, calls = _run_sealed(sealed_box, system_dirs=(str(tools),), **extra)

    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "mcp add microsoft-365" in calls


def test_a_failed_registration_names_the_binary_that_was_run(sealed_box):
    """The hint for seeing the error has to be a command that works for the
    person reading it. The desktop app's copy of Claude Code is not on PATH, so
    a bare `claude` would not be; its path has a space in it, so it is quoted."""
    desktop_dir = sealed_box["root"] / "Application Support" / "Claude"
    desktop_dir.mkdir(parents=True)
    execpath = desktop_dir / "claude"
    _desktop_shim(execpath, sealed_box["log"], reports_claude_code=True, add_exit=1)

    proc, _ = _run_sealed(
        sealed_box, M365_CLIENT_ID=CLIENT_ID, CLAUDE_CODE_EXECPATH=str(execpath)
    )

    assert proc.returncode != 0
    assert "Could not register the connector" in proc.stderr, proc.stderr
    assert f'"{execpath}" mcp add microsoft-365' in proc.stderr, proc.stderr


def test_a_failed_registration_through_claude_on_path_still_says_claude(box):
    """Nothing changes for the common case: with `claude` on PATH the hint is
    the plain `claude mcp add ...` it always was."""
    _exe(box["bin"] / "claude", '#!/bin/sh\ncase "$2" in add) exit 1 ;; esac\nexit 0\n')

    proc, _ = _run(box, M365_CLIENT_ID=CLIENT_ID)

    assert proc.returncode != 0
    assert "     claude mcp add microsoft-365" in proc.stderr, proc.stderr


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


# ---------------------------------------------------------------------------
# The harness itself
#
# The guard that keeps the installer away from a real Claude Code is only worth
# having if it can fail. These show that every PATH the fixtures build passes it,
# that it fails on the two things it exists to catch, and that the `claude` the
# installer launches really does receive the throwaway config.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "fixture_name, env_of, shim_expected",
    [
        ("box", _box_env, True),
        ("sealed_box", _sealed_env, False),
    ],
)
def test_no_fixture_can_reach_a_real_claude(request, fixture_name, env_of, shim_expected):
    """Every PATH a fixture builds resolves `claude` to its own shim, or to
    nothing: never to whatever the host happens to have installed."""
    fx = request.getfixturevalue(fixture_name)

    _assert_nothing_real_is_reachable(env_of(fx), fx["root"], shim_expected=shim_expected)


@pytest.mark.parametrize(
    "fixture_name, runner",
    [("box", _run), ("sealed_box", _run_sealed)],
)
def test_every_runner_checks_the_guard_before_it_starts_the_installer(
    request, monkeypatch, fixture_name, runner
):
    """The guard protects nothing if a run can skip it. With the guard swapped
    for one that raises, each runner has to raise: take the guard call out of
    the shared runner, or let one of these start install.sh itself, and the
    installer runs instead and this fails."""

    class GuardReached(Exception):
        pass

    def tripwire(env, root, *, shim_expected):
        raise GuardReached

    monkeypatch.setattr(sys.modules[__name__], "_assert_nothing_real_is_reachable", tripwire)
    fx = request.getfixturevalue(fixture_name)

    # The client ID is passed so that a run which does get through never stops
    # to ask for it.
    with pytest.raises(GuardReached):
        runner(fx, M365_CLIENT_ID=CLIENT_ID)


def test_the_guard_fails_when_a_real_claude_is_reachable(tmp_path):
    """Negative control. On a machine with no `claude` at all (CI) every test
    above passes with the guard deleted, so the guard has to be shown failing."""
    root = tmp_path / "fixture"
    root.mkdir()
    host = tmp_path / "host-install"  # stands in for wherever a real claude lives
    host.mkdir()
    _exe(host / "claude", "#!/bin/sh\nexit 0\n")
    env = {"PATH": str(host), "NO_COLOR": "1", **_throwaway_config(root)}

    with pytest.raises(AssertionError, match="real claude is reachable"):
        _assert_nothing_real_is_reachable(env, root, shim_expected=True)
    with pytest.raises(AssertionError, match="expected no claude"):
        _assert_nothing_real_is_reachable(env, root, shim_expected=False)


@pytest.mark.parametrize("key", ["HOME", "CLAUDE_CONFIG_DIR"])
def test_the_guard_fails_when_the_config_would_not_be_a_throwaway(tmp_path, key):
    """Negative control for the other layer: a HOME or CLAUDE_CONFIG_DIR that
    points outside the test's own tmp dir is refused."""
    root = tmp_path / "fixture"
    root.mkdir()
    elsewhere = tmp_path / "somewhere-else"
    elsewhere.mkdir()
    env = {"PATH": str(root / "empty"), "NO_COLOR": "1", **_throwaway_config(root), key: str(elsewhere)}

    with pytest.raises(AssertionError, match=key):
        _assert_nothing_real_is_reachable(env, root, shim_expected=False)


def test_the_guard_fails_when_claude_code_execpath_points_outside_the_tmp_dir(tmp_path):
    """Negative control for the installer's second way of reaching a binary.
    CLAUDE_CODE_EXECPATH names something install.sh runs with --version, and as
    `claude mcp add` if it answers as Claude Code, so when it is set it has to
    resolve inside the test's own tmp dir, symlinks followed."""
    root = tmp_path / "fixture"
    root.mkdir()
    elsewhere = tmp_path / "somewhere-else"
    elsewhere.mkdir()
    real = elsewhere / "claude"
    _exe(real, "#!/bin/sh\nexit 0\n")
    inside = root / "claude"
    _exe(inside, "#!/bin/sh\nexit 0\n")
    disguised = root / "claude-link"
    disguised.symlink_to(real)
    env = {"PATH": str(root / "empty"), "NO_COLOR": "1", **_throwaway_config(root)}

    for outside in (real, disguised):
        with pytest.raises(AssertionError, match="CLAUDE_CODE_EXECPATH"):
            _assert_nothing_real_is_reachable(
                {**env, "CLAUDE_CODE_EXECPATH": str(outside)}, root, shim_expected=False
            )
    _assert_nothing_real_is_reachable({**env, "CLAUDE_CODE_EXECPATH": str(inside)}, root, shim_expected=False)
    # An empty value is "not set" to install.sh, so it is to the guard.
    _assert_nothing_real_is_reachable({**env, "CLAUDE_CODE_EXECPATH": ""}, root, shim_expected=False)


def test_the_claude_the_installer_runs_only_sees_a_throwaway_config(box):
    """What matters is not what the harness hands the installer but what the
    `claude` the installer launches actually receives."""
    seen = box["root"] / "seen-by-claude.log"
    _exe(
        box["bin"] / "claude",
        f'#!/bin/sh\necho "$HOME|$CLAUDE_CONFIG_DIR" >> "{seen}"\nexit 0\n',
    )

    proc, _ = _run(box, M365_CLIENT_ID=CLIENT_ID)

    assert proc.returncode == 0, proc.stdout + proc.stderr
    runs = [ln.split("|") for ln in seen.read_text().splitlines()]
    assert len(runs) == 2, f"expected the remove and the add, got {runs!r}"
    root = box["root"].resolve()
    for home, config in runs:
        assert Path(home).resolve().is_relative_to(root), home
        assert Path(config).resolve().is_relative_to(root), config
