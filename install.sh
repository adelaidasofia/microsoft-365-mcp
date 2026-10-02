#!/usr/bin/env bash
#
# One-command install for microsoft-365-mcp.
#
#   cd ~ && git clone https://github.com/adelaidasofia/microsoft-365-mcp.git
#   bash ~/microsoft-365-mcp/install.sh
#
# Builds an isolated venv next to this script, installs the dependencies, asks
# for the Entra Application (client) ID, and registers the server with Claude
# Code. Nothing is written outside this directory and Claude Code's own config.
# Safe to re-run.
#
# Targets macOS's stock bash 3.2, so no 4.x-only syntax.

set -euo pipefail

SERVER_NAME="microsoft-365"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$SCRIPT_DIR/.venv"
VENV_PY="$VENV_DIR/bin/python"

if [ -t 1 ] && [ -z "${NO_COLOR:-}" ]; then
  B=$(printf '\033[1m'); DIM=$(printf '\033[2m'); R=$(printf '\033[0m')
  GRN=$(printf '\033[32m'); RED=$(printf '\033[31m'); YLW=$(printf '\033[33m')
else
  B=""; DIM=""; R=""; GRN=""; RED=""; YLW=""
fi

step() { printf '\n%s==>%s %s%s%s\n' "$GRN" "$R" "$B" "$1" "$R"; }
ok()   { printf '    %s+%s %s\n' "$GRN" "$R" "$1"; }
warn() { printf '    %s!%s %s\n' "$YLW" "$R" "$1"; }
die()  { printf '\n%sX  %s%s\n\n' "$RED" "$1" "$R" >&2; exit 1; }

# Prompts must read from the terminal, not stdin: stdin may be the script
# itself when this is piped, and then every read would silently consume the
# script's own remaining lines instead of waiting for the person.
if [ -r /dev/tty ]; then TTY=/dev/tty; else TTY=/dev/stdin; fi

ask() { # ask <prompt> -> echoes the entered value
  local prompt="$1" value=""
  while [ -z "$value" ]; do
    printf '\n    %s%s%s\n    > ' "$B" "$prompt" "$R" > /dev/tty
    IFS= read -r value < "$TTY" || die "No input received. Run the script from a terminal."
    value="$(printf '%s' "$value" | sed -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')"
    [ -z "$value" ] && printf '    %sThat was empty. Try again.%s\n' "$YLW" "$R" > /dev/tty
  done
  printf '%s' "$value"
}

confirm_yes() { # confirm_yes <question> -> 0 if yes
  local reply=""
  printf '\n    %s%s%s [y/N] ' "$B" "$1" "$R" > /dev/tty
  IFS= read -r reply < "$TTY" || reply=""
  case "$reply" in [yY]*) return 0 ;; *) return 1 ;; esac
}

printf '\n%s  Microsoft 365 MCP  %s\n' "$B" "$R"
printf '%s  Outlook mail, Calendar and OneDrive, connected to Claude Code.%s\n' "$DIM" "$R"

# ---------------------------------------------------------------- 1. tooling

step "Checking what you already have"

command -v git >/dev/null 2>&1 || die \
"git is not installed.
   Run  xcode-select --install  , let it finish, then run this script again."

command -v python3 >/dev/null 2>&1 || die \
"python3 is not installed.
   Run  xcode-select --install  , let it finish, then run this script again."

PY_OK=$(python3 -c 'import sys; print(1 if sys.version_info[:2] >= (3,10) else 0)' 2>/dev/null || echo 0)
[ "$PY_OK" = "1" ] || die \
"python3 is version $(python3 -c 'import sys;print("%d.%d"%sys.version_info[:2])' 2>/dev/null || echo unknown), but 3.10 or newer is needed.
   Install a newer Python from python.org, then run this script again."
ok "python3 $(python3 -c 'import sys;print("%d.%d"%sys.version_info[:2])')"

# Run a command for at most ten seconds, with nothing on its stdin, so a binary
# that hangs, or that waits for input, cannot hang the installer. Stock macOS has
# no `timeout`, but it does have perl, and perl's alarm survives the exec that
# follows it. Where there is no perl the command still runs, just unbounded:
# asked, rather than not asked.
run_bounded() {
  if command -v perl >/dev/null 2>&1; then
    perl -e 'alarm shift; exec @ARGV' 10 "$@" </dev/null
  else
    "$@" </dev/null
  fi
}

# The temporary file the probe below writes the binary's answer to. An interrupt
# while the probe waits on a binary that is not answering (Ctrl-C, the window
# closing, a TERM) would otherwise leave it in TMPDIR, or in /tmp. A trap runs
# after the function's locals are gone, so the name is a global, and an EXIT trap
# removes whatever it names. The probe removes the file itself as soon as it has
# read it, and then forgets the name.
PROBE_OUT=""
trap '[ -z "$PROBE_OUT" ] || rm -f "$PROBE_OUT"' EXIT

# Is the binary at $1 really Claude Code? Ask it, and trust only what it says.
#
# The answer goes to a temporary file, not through a pipe or a $(...): the shell
# then waits for the process it started and no longer, so a child that process
# leaves behind holding its output open (a wrapper script, a helper) cannot
# stretch the ten seconds. It is matched afterwards, never piped straight into
# grep: under `set -o pipefail` a binary that printed its banner and then exited
# non-zero would fail the whole pipeline and be thrown out as "not Claude Code",
# when what it said is the only thing being asked about. Both streams are read:
# a banner printed only on stderr is still a banner. A TMPDIR that names a
# directory that does not exist does not change any of that: the file is made in
# /tmp instead. Only where there is no usable mktemp at all is the answer read
# through $(...), and that waits for a child that keeps the output open for as
# long as the child does.
#
# What counts is a line shaped like what `claude --version` prints,
# "<version> (Claude Code)": it starts with a digit, has no space in the
# version, and ends with " (Claude Code)". It does not have to be the first
# line, so a warning printed ahead of the banner does not turn a real Claude
# Code away, but it has to be among the first 4096 characters of the answer. A
# different tool, or an error that merely mentions Claude Code, does not say
# that.
is_claude_code() { # is_claude_code <binary>
  local said="" rest="" line=""
  PROBE_OUT="$(mktemp "${TMPDIR:-/tmp}/claude-probe.XXXXXX" 2>/dev/null || mktemp /tmp/claude-probe.XXXXXX 2>/dev/null)" || PROBE_OUT=""
  if [ -n "$PROBE_OUT" ]; then
    run_bounded "$1" --version >"$PROBE_OUT" 2>&1 || true
    said="$(cat "$PROBE_OUT" 2>/dev/null || true)"
    rm -f "$PROBE_OUT"
    PROBE_OUT=""
  else
    said="$(run_bounded "$1" --version 2>&1 || true)"
  fi
  # Only the start of the answer is looked at. The banner is one short line, and
  # going through pages of output one line at a time takes bash far longer than
  # that is worth: the time grows much faster than the length does.
  said="${said:0:4096}"
  rest="$said"
  while [ -n "$rest" ]; do
    line="${rest%%$'\n'*}"
    case "$line" in
      *" "*" (Claude Code)") ;;
      [0-9]*" (Claude Code)") return 0 ;;
    esac
    [ "$line" = "$rest" ] && break
    rest="${rest#*$'\n'}"
  done
  return 1
}

# Two ways to reach Claude Code: on PATH (the common case), or via
# CLAUDE_CODE_EXECPATH when this script is run from inside the Claude desktop
# app's own Code tab, which bundles its own Claude Code binary and exports
# that variable to point at it -- often without ever putting `claude` on PATH.
CLAUDE_BIN=""
if command -v claude >/dev/null 2>&1; then
  CLAUDE_BIN="claude"
  ok "claude"
elif [ -n "${CLAUDE_CODE_EXECPATH:-}" ]; then
  if [ -f "$CLAUDE_CODE_EXECPATH" ] && [ -x "$CLAUDE_CODE_EXECPATH" ] \
     && is_claude_code "$CLAUDE_CODE_EXECPATH"; then
    CLAUDE_BIN="$CLAUDE_CODE_EXECPATH"
    ok "using the copy of Claude Code the Claude desktop app runs"
  else
    # It was set and was not usable: missing, not executable, a different
    # program, or silent past the ten second bound (which only exists where there
    # is perl). Saying Claude Code is not installed would be wrong news about a
    # copy that may be sitting right there, so say what was refused. Two things
    # the message keeps out: a time limit, since without perl there is none, and a
    # `claude --version` check in Terminal, which for the people this is for only
    # prints "command not found".
    die \
"Claude Code is not on your PATH, and CLAUDE_CODE_EXECPATH is set to something that could not be used instead:
     $CLAUDE_CODE_EXECPATH
   It has to be an executable file that answers  --version  with a line like
   \"2.1.281 (Claude Code)\".
   In the Claude desktop app: put the app's own copy of Claude Code on PATH, then run this script again.
   Anywhere else: install Claude Code first, quit and reopen Terminal, then run this script again."
  fi
fi

[ -n "$CLAUDE_BIN" ] || die \
"Claude Code is not installed, or its 'claude' command is not on your PATH.
   Install Claude Code first, quit and reopen Terminal, then run this again."

[ -f "$SCRIPT_DIR/server.py" ] || die \
"This script is not sitting next to server.py, so the clone looks incomplete.
   Delete the folder and clone it again."

# --------------------------------------------------------------- 2. install

step "Installing the connector (about a minute)"

# A venv is not optional here: Homebrew and python.org interpreters are marked
# externally-managed (PEP 668), so a plain `pip install` refuses outright.
if [ ! -x "$VENV_PY" ]; then
  python3 -m venv "$VENV_DIR" || die \
"Could not create the virtual environment in $VENV_DIR.
   If that folder half-exists, delete it and run this script again."
fi
ok "isolated environment ready"

"$VENV_PY" -m pip install --quiet --upgrade pip >/dev/null 2>&1 || true
"$VENV_PY" -m pip install --quiet -r "$SCRIPT_DIR/requirements.txt" || die \
"Could not install the dependencies.
   Check your internet connection and run this script again."
ok "dependencies installed"

# Import the real module. A dependency that resolves at install time but not at
# import time (a missing transitive, a version clash) otherwise shows up much
# later as a connector that is simply absent from /mcp, with nothing to read.
LOAD_ERR="$(cd "$SCRIPT_DIR" && "$VENV_PY" -c 'import server' 2>&1)" || die \
"The connector installed but did not load. Please send this to the cohort channel:

$LOAD_ERR"
ok "connector loads"

# ------------------------------------------------------------ 3. credentials

step "Your Microsoft app registration"

# Two ways in. A person running this by hand gets prompted. An agent (Claude
# Code) driving it passes the value in the environment and is never asked a
# question, because a prompt it cannot answer would hang the whole install.
CLIENT_ID="${M365_CLIENT_ID:-}"
NONINTERACTIVE=0
if [ -n "$CLIENT_ID" ]; then
  NONINTERACTIVE=1
  ok "using the client ID passed in the environment"
else
  cat <<EOF

    You need one value from the Microsoft Entra admin centre, on the
    Overview page of the app you registered:

      $B Application (client) ID $R   looks like 1a2b3c4d-5e6f-7890-abcd-ef1234567890

    Do not have it yet? Press Ctrl-C, finish the portal steps in the
    guide, then run this script again. Nothing done so far is lost.
EOF
  CLIENT_ID="$(ask 'Paste your Application (client) ID, then press Return')"
fi

# Shape check. An Entra client ID is always a GUID. The common mistakes are
# pasting the Directory (tenant) ID's neighbour field, the Object ID, or the
# app's display name, and every one of those fails much later with an opaque
# AADSTS code instead of here.
if printf '%s' "$CLIENT_ID" | grep -Eqi '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'; then
  ok "client ID looks right"
else
  warn "That does not look like an Application (client) ID."
  warn "It should be 36 characters of the form 1a2b3c4d-5e6f-7890-abcd-ef1234567890."
  warn "The app's name, and the Object ID, are the two things most often pasted here by mistake."
  if [ "$NONINTERACTIVE" = "1" ]; then
    die "Stopping rather than registering a client ID that cannot work. Check the Overview page and try again."
  fi
  confirm_yes "Use it anyway?" || die "Nothing was changed. Run the script again with the right value."
fi

# -------------------------------------------------------------- 4. register

step "Connecting it to Claude Code"

# Re-running should heal a bad value rather than fail on "already exists". The
# remove is unconditional and its failure ignored, so this does not depend on
# parsing `claude mcp list` output, which is a display format, not a contract.
"$CLAUDE_BIN" mcp remove "$SERVER_NAME" -s user >/dev/null 2>&1 || true

# If registering fails, the person is told what to run to see why. That has to be
# the binary that was just run: the desktop app's copy is not on PATH, so a bare
# `claude` would be a command that does not exist for them. Every path in it goes
# through shq, so one with a space in it ("Application Support"), a quote or a `$`
# can still be pasted as it is.
shq() { # shq <string> -> <string> as one shell word, read back exactly as it was
  case "$1" in
    ""|*[!_./:=@%+,[:alnum:]-]*) printf "'%s'" "$(printf '%s' "$1" | sed "s/'/'\\\\''/g")" ;;
    *) printf '%s' "$1" ;;
  esac
}

"$CLAUDE_BIN" mcp add "$SERVER_NAME" -s user \
  -e "M365_CLIENT_ID=$CLIENT_ID" \
  -- "$VENV_PY" "$SCRIPT_DIR/server.py" >/dev/null || die \
"Could not register the connector with Claude Code.
   Run this to see the error:
     $(shq "$CLAUDE_BIN") mcp add $SERVER_NAME -s user -e M365_CLIENT_ID=... -- $(shq "$VENV_PY") $(shq "$SCRIPT_DIR/server.py")"

ok "registered as \"$SERVER_NAME\""

cat <<EOF

$GRN  Installed.$R

  ${B}Two things left, both inside Claude Code:${R}

    1. Quit Claude Code completely and open it again.
       It only picks up new connectors when it starts.

    2. Send it this message:

           Call m365_account_add

       Your browser opens. Sign in as yourself and approve the access.
       If it says "Need admin approval", that is your company's policy,
       not a fault here. Use a personal outlook.com account for now and
       send your IT team the request from the guide.

  Then try:  Call outlook_search with query "" and limit 5

EOF
