"""Microsoft 365 MCP — Outlook mail + Calendar + OneDrive, multi-account,
token-efficient. The Microsoft sibling of google-workspace-mcp.

Registered in your project's `.mcp.json` (or user scope) as `microsoft-365`.

Two design goals:
1. Multi-account. Every tool takes `account` as an optional email; default is
   taken from M365_DEFAULT_ACCOUNT or the first authorized account.
2. Token-efficient. Search/list tools return compact shapes. Bodies and file
   content are opt-in (Graph $select + Prefer: body-content-type="text").

Keychain service name: `microsoft-365-mcp` (see accounts.py KEYRING_SERVICE).
See SETUP.md for Entra ID app setup; see README.md for day-to-day usage.
"""

from __future__ import annotations

import functools
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from fastmcp import FastMCP

import accounts
import calendar_tools
import graph
import mail_tools
import onedrive_tools

mcp = FastMCP("microsoft-365")


def _tool(fn):
    """Register a tool with per-call observability (4-field JSONL audit)."""

    @functools.wraps(fn)
    def wrapper(**kwargs):
        started = time.monotonic()
        try:
            result = fn(**kwargs)
        except graph.GraphError as e:
            graph.audit(fn.__name__, kwargs, started, error_class=e.error_class)
            raise
        except Exception:
            graph.audit(fn.__name__, kwargs, started, error_class="internal_error")
            raise
        graph.audit(fn.__name__, kwargs, started, result=result)
        return result

    return mcp.tool()(wrapper)


# ---------------------------------------------------------------------------
# Account management
# ---------------------------------------------------------------------------


@_tool
def m365_account_add(method: str = "interactive") -> dict:
    """Authorize a new Microsoft account. Default opens a browser sign-in
    (interactive). Stores tokens in the macOS Keychain (service:
    microsoft-365-mcp) or the M365_TOKEN_FILE fallback.

    Args:
        method: 'interactive' (browser, default) or 'device_code' (headless /
            remote: returns a URL + code; then call m365_account_complete).

    Requires an Entra Application (client) ID via M365_CLIENT_ID or
    client_config.json. See SETUP.md.
    """
    return accounts.add_account(method=method)


@_tool
def m365_account_complete() -> dict:
    """Finish a device-code sign-in started by m365_account_add with
    method='device_code'. Blocks until the user completes the browser step."""
    return accounts.complete_device_flow()


@_tool
def m365_account_list() -> dict:
    """List all authorized Microsoft accounts and which one is the default."""
    all_accounts = accounts.list_accounts()
    default = ""
    try:
        default = accounts.default_account()
    except accounts.AccountError:
        pass
    return {"accounts": all_accounts, "default": default, "count": len(all_accounts)}


@_tool
def m365_account_remove(email: str) -> dict:
    """Remove an account's local tokens. (Does not revoke Microsoft-side; do
    that at https://myaccount.microsoft.com/ > Privacy > App access.)"""
    return accounts.remove_account(email)


# ---------------------------------------------------------------------------
# Outlook mail
# ---------------------------------------------------------------------------


@_tool
def outlook_search(
    query: str = "",
    account: str | None = None,
    limit: int = 10,
    folder: str | None = None,
    unread_only: bool = False,
) -> list[dict]:
    """Search Outlook messages. Returns compact shape (no body) — call
    outlook_read for the body.

    Args:
        query: KQL, e.g. 'from:alex@example.com subject:invoice received>=2026-06-01',
            or plain words. Empty query lists the folder newest-first.
        account: Email address; defaults to configured default account.
        limit: 1-50. Default 10.
        folder: Well-known name ('inbox', 'archive', 'sentitems', 'drafts',
            'deleteditems', 'junkemail') or a folder id from outlook_folders_list.
            Default: inbox for listing, all mail for query searches.
        unread_only: Only unread messages.
    """
    return mail_tools.search(
        query=query, account=account, limit=limit, folder=folder,
        unread_only=unread_only,
    )


@_tool
def outlook_read(
    message_id: str,
    account: str | None = None,
    keep_html: bool = False,
    full_thread: bool = False,
) -> dict:
    """Read a single message or its whole conversation. Plain-text body by
    default (Graph converts server-side).

    Args:
        message_id: Message id from outlook_search.
        keep_html: If True, return the raw HTML body.
        full_thread: If True, return every message in the conversation.
    """
    return mail_tools.read(
        message_id=message_id, account=account,
        keep_html=keep_html, full_thread=full_thread,
    )


@_tool
def outlook_send(
    to: list[str],
    subject: str,
    body: str,
    account: str | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    reply_to: str | None = None,
    dry_run: bool = False,
) -> dict:
    """Send an email. DESTRUCTIVE — irreversible once sent.

    Args:
        dry_run: If True, return what WOULD be sent without calling the API.
            Use this to verify recipient, subject, and body before committing.
    """
    return mail_tools.send(
        to=to, subject=subject, body=body, account=account,
        cc=cc, bcc=bcc, reply_to=reply_to, dry_run=dry_run,
    )


@_tool
def outlook_draft(
    to: list[str],
    subject: str,
    body: str,
    account: str | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
) -> dict:
    """Create a draft in the Drafts folder. Does not send."""
    return mail_tools.draft(
        to=to, subject=subject, body=body, account=account, cc=cc, bcc=bcc,
    )


@_tool
def outlook_reply(
    message_id: str,
    body: str,
    account: str | None = None,
    reply_all: bool = False,
    dry_run: bool = False,
) -> dict:
    """Reply to a message. DESTRUCTIVE — sends immediately, same blast radius
    as outlook_send. Graph preserves the thread and recipients server-side.

    Args:
        dry_run: If True, show what WOULD be sent without sending.
    """
    return mail_tools.reply(
        message_id=message_id, body=body, account=account,
        reply_all=reply_all, dry_run=dry_run,
    )


@_tool
def outlook_folders_list(account: str | None = None) -> list[dict]:
    """List mail folders with unread/total counts. Use folder ids (or
    well-known names) with outlook_search and outlook_move."""
    return mail_tools.folders_list(account=account)


@_tool
def outlook_move(
    message_ids: list[str],
    destination: str,
    account: str | None = None,
) -> dict:
    """Move messages to a folder. destination = well-known name ('archive',
    'deleteditems', ...) or a folder id from outlook_folders_list."""
    return mail_tools.move(
        message_ids=message_ids, destination=destination, account=account,
    )


@_tool
def outlook_archive(message_ids: list[str], account: str | None = None) -> dict:
    """Archive messages: moves them to the Archive folder, still searchable.
    No deletion. Use outlook_move with 'deleteditems' for soft-delete."""
    return mail_tools.archive(message_ids=message_ids, account=account)


@_tool
def outlook_mark_read(
    message_ids: list[str],
    read: bool = True,
    account: str | None = None,
) -> dict:
    """Mark messages as read (or unread with read=False)."""
    return mail_tools.mark_read(message_ids=message_ids, read=read, account=account)


# ---------------------------------------------------------------------------
# Calendar
# ---------------------------------------------------------------------------


@_tool
def mscal_list_calendars(account: str | None = None) -> list[dict]:
    """List all calendars visible to this account."""
    return calendar_tools.list_calendars(account=account)


@_tool
def mscal_list_events(
    account: str | None = None,
    calendar_id: str = "primary",
    time_min: str = "now",
    time_max: str | None = None,
    days_ahead: int = 7,
    query: str | None = None,
    max_results: int = 25,
    verbose: bool = False,
) -> list[dict]:
    """List events on a calendar (recurring events expanded to instances).

    Args:
        time_min: ISO 8601 or 'now'/'today'/'tomorrow'.
        time_max: ISO 8601. If omitted, uses time_min + days_ahead.
        query: Match within event titles.
        verbose: Include description preview, attendees, recurrence.
    """
    return calendar_tools.list_events(
        account=account, calendar_id=calendar_id, time_min=time_min,
        time_max=time_max, days_ahead=days_ahead, query=query,
        max_results=max_results, verbose=verbose,
    )


@_tool
def mscal_create_event(
    summary: str,
    start: str,
    end: str,
    account: str | None = None,
    calendar_id: str = "primary",
    description: str | None = None,
    location: str | None = None,
    attendees: list[str] | None = None,
    time_zone: str = "America/Bogota",
    add_teams_meeting: bool = False,
) -> dict:
    """Create a calendar event. Invites are sent to attendees automatically.

    Args:
        start, end: ISO 8601 (or 'now'/'today'/'tomorrow').
        add_teams_meeting: Attach a Microsoft Teams meeting link.
    """
    return calendar_tools.create_event(
        summary=summary, start=start, end=end, account=account,
        calendar_id=calendar_id, description=description, location=location,
        attendees=attendees, time_zone=time_zone,
        add_teams_meeting=add_teams_meeting,
    )


@_tool
def mscal_update_event(
    event_id: str,
    account: str | None = None,
    summary: str | None = None,
    start: str | None = None,
    end: str | None = None,
    description: str | None = None,
    location: str | None = None,
    attendees_add: list[str] | None = None,
    attendees_remove: list[str] | None = None,
    time_zone: str = "America/Bogota",
) -> dict:
    """Partial-update an event. Only pass fields you want to change."""
    return calendar_tools.update_event(
        event_id=event_id, account=account, summary=summary, start=start,
        end=end, description=description, location=location,
        attendees_add=attendees_add, attendees_remove=attendees_remove,
        time_zone=time_zone,
    )


@_tool
def mscal_delete_event(event_id: str, account: str | None = None) -> dict:
    """Delete (cancel) an event. Attendees are notified."""
    return calendar_tools.delete_event(event_id=event_id, account=account)


@_tool
def mscal_freebusy(
    time_min: str,
    time_max: str,
    account: str | None = None,
    emails: list[str] | None = None,
) -> dict:
    """Check busy windows for one or more people (for scheduling). Requires
    the target calendars to be visible to this account.

    Args:
        emails: People to check. Defaults to the authenticated account.
    """
    return calendar_tools.freebusy(
        time_min=time_min, time_max=time_max, account=account, emails=emails,
    )


@_tool
def mscal_respond(
    event_id: str,
    response: str,
    account: str | None = None,
    comment: str | None = None,
) -> dict:
    """Respond to an event invite.

    Args:
        response: 'accepted' | 'declined' | 'tentative'.
        comment: Optional note sent with the response.
    """
    return calendar_tools.respond(
        event_id=event_id, response=response, account=account, comment=comment,
    )


# ---------------------------------------------------------------------------
# OneDrive
# ---------------------------------------------------------------------------


@_tool
def onedrive_search(
    query: str,
    account: str | None = None,
    limit: int = 10,
) -> list[dict]:
    """Search OneDrive files by name and content. Returns compact metadata
    (no content) — call onedrive_read_file for content."""
    return onedrive_tools.search(query=query, account=account, limit=limit)


@_tool
def onedrive_read_file(
    item_id: str,
    account: str | None = None,
    include_content: bool = False,
    max_chars: int = 50_000,
) -> dict:
    """Read a OneDrive item. Metadata by default; content is opt-in.

    Args:
        include_content: If True, download and return content (UTF-8 text;
            binary files return a size placeholder — Office docs are binary,
            convert or export them first).
        max_chars: Cap on returned content. Default 50k.
    """
    return onedrive_tools.read_file(
        item_id=item_id, account=account,
        include_content=include_content, max_chars=max_chars,
    )


@_tool
def onedrive_list_folder(
    folder_id: str = "root",
    account: str | None = None,
    limit: int = 50,
) -> list[dict]:
    """List direct children of a folder. Use 'root' for the OneDrive root."""
    return onedrive_tools.list_folder(folder_id=folder_id, account=account, limit=limit)


@_tool
def onedrive_create_folder(
    name: str,
    parent_id: str = "root",
    account: str | None = None,
) -> dict:
    """Create a folder. Name conflicts auto-rename (folder (1))."""
    return onedrive_tools.create_folder(name=name, parent_id=parent_id, account=account)


@_tool
def onedrive_upload(
    local_path: str,
    name: str | None = None,
    parent_id: str = "root",
    account: str | None = None,
) -> dict:
    """Upload a local file to OneDrive. Large files (>4 MB) upload in chunks."""
    return onedrive_tools.upload(
        local_path=local_path, name=name, parent_id=parent_id, account=account,
    )


@_tool
def onedrive_move(item_id: str, new_parent_id: str, account: str | None = None) -> dict:
    """Move a file or folder to a different parent folder."""
    return onedrive_tools.move(item_id=item_id, new_parent_id=new_parent_id, account=account)


@_tool
def onedrive_rename(item_id: str, new_name: str, account: str | None = None) -> dict:
    """Rename a file or folder."""
    return onedrive_tools.rename(item_id=item_id, new_name=new_name, account=account)


@_tool
def onedrive_delete(
    item_id: str,
    account: str | None = None,
    dry_run: bool = False,
) -> dict:
    """Move an item to the recycle bin (recoverable ~30 days). DESTRUCTIVE.

    Args:
        dry_run: If True, show what would be deleted without deleting it.
    """
    return onedrive_tools.delete(item_id=item_id, account=account, dry_run=dry_run)


@_tool
def onedrive_share(
    item_id: str,
    email: str,
    role: str = "write",
    send_notification: bool = False,
    message: str | None = None,
    account: str | None = None,
) -> dict:
    """Grant a person access to a file/folder by email. Creates a permission.
    To change an existing grantee's role, use onedrive_permission_update.

    Args:
        role: 'read' | 'write'.
        send_notification: If True, Microsoft emails the grantee.
    """
    return onedrive_tools.share(
        item_id=item_id, email=email, role=role,
        send_notification=send_notification, message=message, account=account,
    )


@_tool
def onedrive_share_link(
    item_id: str,
    link_type: str = "view",
    scope: str = "organization",
    account: str | None = None,
) -> dict:
    """Create a sharing link for an item.

    Args:
        link_type: 'view' | 'edit'.
        scope: 'organization' (tenant only) | 'anonymous' (anyone with the
            link — corporate tenants often disable this; expect a policy error).
    """
    return onedrive_tools.share_link(
        item_id=item_id, link_type=link_type, scope=scope, account=account,
    )


@_tool
def onedrive_permission_list(item_id: str, account: str | None = None) -> list[dict]:
    """List all permissions (who has access) on a file or folder."""
    return onedrive_tools.permission_list(item_id=item_id, account=account)


@_tool
def onedrive_permission_update(
    item_id: str,
    permission_id: str,
    role: str,
    account: str | None = None,
) -> dict:
    """Change the role of someone who already has access. Get permission_id
    from onedrive_permission_list. To grant new access, use onedrive_share.

    Args:
        role: 'read' | 'write'.
    """
    return onedrive_tools.permission_update(
        item_id=item_id, permission_id=permission_id, role=role, account=account,
    )


@_tool
def onedrive_permission_delete(
    item_id: str,
    permission_id: str,
    account: str | None = None,
) -> dict:
    """Revoke a permission. DESTRUCTIVE — removes access immediately.
    Get permission_id from onedrive_permission_list first."""
    return onedrive_tools.permission_delete(
        item_id=item_id, permission_id=permission_id, account=account,
    )


if __name__ == "__main__":
    mcp.run()
