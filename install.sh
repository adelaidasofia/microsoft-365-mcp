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

command -v claude >/dev/null 2>&1 || die \
"Claude Code is not installed, or its 'claude' command is not on your PATH.
   Install Claude Code first, quit and reopen Terminal, then run this again."
ok "claude"

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
claude mcp remove "$SERVER_NAME" -s user >/dev/null 2>&1 || true

claude mcp add "$SERVER_NAME" -s user \
  -e "M365_CLIENT_ID=$CLIENT_ID" \
  -- "$VENV_PY" "$SCRIPT_DIR/server.py" >/dev/null || die \
"Could not register the connector with Claude Code.
   Run this to see the error:
     claude mcp add $SERVER_NAME -s user -e M365_CLIENT_ID=... -- $VENV_PY $SCRIPT_DIR/server.py"

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
