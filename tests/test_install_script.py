"""install.sh is the path every new install takes, so it needs a gate.

Hermetic: `claude` and the venv interpreter are shims on PATH, so nothing here
touches the real Claude Code config or the network.

Three things are worth pinning, and nothing else:
  1. valid input produces the exact `claude mcp add` argv (a lost `--` or a
     dropped `-s user` is invisible until a person tries to use it)
  2. a client ID that cannot work registers nothing and exits non-zero. An
     Entra client ID is always a GUID, and pasting the app's display name or
     its Object ID instead is the common mistake -- it otherwise surfaces much
     later as an opaque AADSTS code with nothing pointing back at the typo
  3. re-running heals rather than fails

Subprocess calls pass argument lists, never shell strings.
"""

import os
import shutil
import stat
import subprocess
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

    return {"repo": repo, "bin": bindir, "log": log}


def _run(box, **creds):
    env = dict(os.environ, PATH=f"{box['bin']}{os.pathsep}{os.environ['PATH']}", NO_COLOR="1")
    env.pop("M365_CLIENT_ID", None)
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
