# microsoft-365-mcp

Multi-account, token-efficient MCP for **Outlook mail + Calendar + OneDrive**
via Microsoft Graph. The Microsoft sibling of
[google-workspace-mcp](https://github.com/adelaidasofia/google-workspace-mcp).
Built because the official Claude Microsoft 365 connector is read-only and
vendor-locked — this one is your own app, your own OAuth, read-write, and
works from any MCP client.

## Why this exists

- **Multi-account**: OAuth multiple Microsoft accounts (work + personal). Every
  tool takes an `account` email.
- **Token-efficient**: Search/list returns compact shapes (`{id, from, subject,
  snippet, ...}` for mail, `{id, name, mime, modified, size, ...}` for OneDrive).
  Bodies and file content are opt-in; Graph converts mail bodies to plain text
  server-side (`Prefer: outlook.body-content-type="text"`).
- **Keychain-backed**: Tokens live in the OS keychain on macOS (file fallback on
  Windows/Linux — see SETUP.md). No tokens in any repo.
- **Own-OAuth**: You (or your team) register a free Entra app; a corporate
  admin can allowlist a single Client ID. No third-party middleman touches mail.

## Tools (v0.1, 33 tools)

### Account management (4)
- `m365_account_add` — browser OAuth flow (or `method="device_code"` for headless)
- `m365_account_complete` — finish a device-code sign-in
- `m365_account_list` — list authorized accounts + default
- `m365_account_remove` — remove local tokens (doesn't revoke Microsoft-side)

### Outlook mail (9)
- `outlook_search` — KQL search or folder listing. Compact response.
- `outlook_read` — read one message or the whole conversation. Plain-text body.
- `outlook_send` — send mail (supports `dry_run=True`)
- `outlook_draft` — create a draft
- `outlook_reply` — reply / reply-all (Graph preserves the thread server-side)
- `outlook_folders_list` — folders with unread/total counts
- `outlook_move` — batch move to a folder (well-known names work: `archive`, `deleteditems`, ...)
- `outlook_archive` — batch archive
- `outlook_mark_read` — batch mark read/unread

### Calendar (7)
- `mscal_list_calendars` — list all calendars
- `mscal_list_events` — list events, recurrences expanded (compact by default, `verbose=True` for full)
- `mscal_create_event` — create event, optional Teams meeting link
- `mscal_update_event` — partial-update fields
- `mscal_delete_event` — delete/cancel
- `mscal_freebusy` — busy windows for one or more people (getSchedule)
- `mscal_respond` — accept/decline/tentative

### OneDrive (13)
- `onedrive_search` — search filenames + content. Metadata-only response.
- `onedrive_read_file` — metadata by default; `include_content=True` for body
- `onedrive_list_folder` — direct children (`'root'` for OneDrive root)
- `onedrive_create_folder` — create a folder (conflicts auto-rename)
- `onedrive_upload` — upload a local file; >4 MB goes through a chunked session
- `onedrive_move` — change parent folder
- `onedrive_rename` — rename
- `onedrive_delete` — recycle bin (recoverable), supports `dry_run=True`
- `onedrive_share` — grant a person read/write access by email
- `onedrive_share_link` — create a view/edit sharing link
- `onedrive_permission_list` — who has access
- `onedrive_permission_update` — change a grantee's role
- `onedrive_permission_delete` — revoke access

Outlook uses folders + categories rather than Gmail labels; folder tools cover
the label workflows. SharePoint sites, Teams chats, and mail attachments are
deliberate v2 scope (`Sites.Read.All`, `Chat.Read`).

## Install

```bash
git clone https://github.com/adelaidasofia/microsoft-365-mcp.git
cd microsoft-365-mcp
pip install -e .
```

Register in Claude Code — project scope (`.mcp.json` at your project root):

```json
{
  "mcpServers": {
    "microsoft-365": {
      "type": "stdio",
      "command": "python3",
      "args": ["/absolute/path/to/microsoft-365-mcp/server.py"]
    }
  }
}
```

or user scope:

```bash
claude mcp add -s user microsoft-365 python3 /absolute/path/to/microsoft-365-mcp/server.py
```

Then follow **[SETUP.md](SETUP.md)** for the Entra app (2 minutes with a shared
team app, ~15 minutes solo) and OAuth your accounts.

## Config

| Env var | Purpose |
|---|---|
| `M365_CLIENT_ID` | Entra Application (client) ID (or use `client_config.json`) |
| `M365_TENANT_ID` | `common` (default) / `consumers` / `organizations` / tenant GUID |
| `M365_DEFAULT_ACCOUNT` | Default account email (else first authorized) |
| `M365_TOKEN_FILE` | Force file token storage (auto on Windows) |
| `M365_AUDIT_IO` | `1` = log full tool input/output payloads (default: summarized) |
| `MYCELIUM_NO_PING` | `1` = disable the one-time anonymous install ping |

## Observability

Every tool call appends a JSONL record (`execution_time_ms`, summarized `io`,
`token_usage`, `error_class`) to `~/.claude/microsoft-365-mcp/audit.log.jsonl`.
Destructive OneDrive ops also land in a plain audit log next to it. Payload
contents are NOT logged unless you opt in with `M365_AUDIT_IO=1`.

## Telemetry

On first session start after install, the plugin sends a single anonymous
`{plugin, version}` ping to myceliumai.co (no PII, one sentinel file per
machine). Set `MYCELIUM_NO_PING=1` to disable.

## Tests

```bash
pip install -e ".[test]"
pytest tests/
```

52 tests run against a fake Graph transport — no network, no real tokens.

## License

MIT

---

Built by Adelaida Diaz-Roa. Full install or team version at diazroa.com.
