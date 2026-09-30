# Guided setup: connecting someone's Outlook through Claude

**Who this is for.** An AI assistant (Claude Code, or any agent that can run a terminal) that was handed this repository's link by a person who wants their Outlook mail, Calendar and OneDrive connected. Assume the person is not technical and may not read English. If you are a person setting this up by hand, [SETUP.md](SETUP.md) is the full manual.

The person's side of this is three things: paste messages you give them into Claude in Chrome, type their own password when a page asks for it, and restart Claude when you tell them to. Everything else is yours.

## How to run it

- **Talk in the person's language.** Reply in the language they wrote in. Short sentences, no jargon, one step at a time.
- **You do everything in the terminal.** Never ask the person to type a command, edit a file, or open Terminal. When a command needs their approval, say in one plain sentence what it does.
- **Anything in a web browser goes to Claude in Chrome** (the Claude extension for Google Chrome, [claude.com/chrome](https://claude.com/chrome)). Give the person ONE message to paste there, introduced with the line "Copy this and paste it into Claude in Chrome:" in their language, followed by one line saying what to bring back to you.
- **Passwords and authenticator codes are theirs.** Never ask for them. Every message you write for Claude in Chrome must say: if a page asks for my password or a code, stop and let me type it.
- **Tell them early:** keep this window open until I say we are done. The setup is lost if they close the session halfway.
- **Stay on this connector.** The person chose this repository; do not steer them to a different Microsoft connector.

## Which Microsoft app

The connector signs in through a Microsoft Entra app, identified by its **Application (client) ID** (a GUID like `1a2b3c4d-5e6f-7890-abcd-ef1234567890`).

- **They gave you an ID** (their program, team or company registered the app): use it. This is the normal path and needs no Microsoft portal at all.
- **They did not:** ask one question, whether their program or company gave them an app ID. If not, they need their own app registration (SETUP.md, sections 1 to 3), which Claude in Chrome can do for them in the Entra admin center. This works with a work or school account. A personal outlook.com, hotmail.com or live.com account cannot create one: Microsoft requires a directory for new app registrations, and personal accounts do not have one. Those people need an ID from a program or team.

## Why this is safe (so you can reassure, not alarm)

- The app ID is a public identifier, not a secret. There is no app secret at all: this is a public desktop client.
- Each person signs in with their own account and gets their own tokens, stored on their own computer. The app grants no access to anyone else's mailbox.
- Every permission is delegated: the connector can only do what the person can already do, acting as them, in their own mailbox and their own OneDrive.

## Steps

### 1. Get the code (you)

```bash
cd ~ && if [ -d microsoft-365-mcp/.git ]; then git -C microsoft-365-mcp pull --ff-only; else git clone https://github.com/adelaidasofia/microsoft-365-mcp.git; fi
```

On Windows, see [Windows](#windows) below.

### 2. Install (you)

```bash
M365_CLIENT_ID='<application client id>' bash ~/microsoft-365-mcp/install.sh
```

With the ID set, the installer asks no questions: it builds its own environment, installs dependencies, checks the ID's shape and registers the connector with Claude Code.

If it stops with "Claude Code is not installed" while you are running inside the Claude desktop app, put the app's own copy of Claude Code on `PATH` and run it again:

```bash
PATH="$(dirname "$CLAUDE_CODE_EXECPATH"):$PATH" M365_CLIENT_ID='<application client id>' bash ~/microsoft-365-mcp/install.sh
```

### 3. Restart and connect their account (the person, helped by Claude in Chrome)

A running Claude only loads new connectors when it starts. Before they restart, tell them all of this, because the next session will not have this context:

1. Quit Claude completely (Cmd + Q on a Mac) and open it again.
2. Send: "Connect my Outlook account and show me my last 5 emails." (in their language)
3. A Microsoft page opens. They sign in themselves and accept the permissions. If they want help with that screen, they can tell Claude in Chrome: "Help me accept this screen, it's my program's app. I'll type my password myself."

## If something goes wrong

| What they see | What to do |
|---|---|
| "Need admin approval" on a work account | Their company requires an administrator to approve new apps. If the page offers "Request approval", they can send it with one line saying what it is for. Otherwise give them the IT message below. Connect a personal account in the meantime. |
| "AADSTS50011" (redirect URI mismatch) or "AADSTS7000218" (public client) | The app registration is missing the `http://localhost` public client redirect, or "Allow public client flows" is off. Whoever owns the app fixes it (SETUP.md sections 1 and 2). |
| "AADSTS65001" (not consented) | The permissions were not accepted, or admin consent is still pending. |
| A personal account fails to sign in | Some older personal accounts misroute on the default tenant. Set `M365_TENANT_ID=consumers` on the registration and connect again. |
| Worked, then stopped after about 90 days without use | Microsoft expires inactive sign-ins. Run `m365_account_add` again. |

**IT message** (translate to their language and fill in the ID):

> Hello. I am connecting my work Microsoft account to an AI assistant that runs on my own computer. Could you grant admin consent for the app with Application (client) ID `<the app ID>`? In the Microsoft Entra admin center it is under Enterprise applications, then the app, then Permissions, then Grant admin consent. It needs these delegated Microsoft Graph permissions: Mail.ReadWrite, Mail.Send, Calendars.ReadWrite, Files.ReadWrite, User.Read and offline_access. It is a public desktop client with no app secret; I sign in with my own account and access stays on my computer.

## Windows

There is no Windows installer yet, so do its steps yourself in PowerShell:

```powershell
cd ~; git clone https://github.com/adelaidasofia/microsoft-365-mcp.git
py -3 -m venv microsoft-365-mcp\.venv
.\microsoft-365-mcp\.venv\Scripts\python.exe -m pip install -r .\microsoft-365-mcp\requirements.txt
claude.cmd mcp add microsoft-365 -s user -e M365_CLIENT_ID=<application client id> -- "$HOME\microsoft-365-mcp\.venv\Scripts\python.exe" "$HOME\microsoft-365-mcp\server.py"
```

Call `claude.cmd`, not `claude`: PowerShell drops a bare `--` when it goes through the `claude.ps1` shim, and the server command then registers wrong. Tokens fall back to a file automatically on Windows (SETUP.md, Token storage). Everything else is the same.
