"""Outlook mail tool implementations — token-efficient Microsoft Graph wrappers.

Design rules (mirror google-workspace-mcp gmail_tools):
- Search returns minimal shape: id, conversation_id, from, subject, snippet,
  date, unread, folder. ~80% token savings vs. full messages.
- `outlook_read` is the only tool that returns bodies. Graph is asked for
  text server-side (Prefer: outlook.body-content-type="text") unless
  keep_html=True — no client-side HTML stripping needed.
- All tools take `account`; None means "use default".
- Outlook has folders, not labels. Well-known names (inbox, archive,
  sentitems, drafts, deleteditems, junkemail) work anywhere a folder is taken.
"""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Any

import graph

_SELECT = (
    "id,conversationId,subject,bodyPreview,from,toRecipients,receivedDateTime,"
    "isRead,hasAttachments,parentFolderId,webLink"
)

_TEXT_BODY = {"Prefer": 'outlook.body-content-type="text"'}


def _addr(recipient: dict) -> str:
    email = (recipient or {}).get("emailAddress", {})
    name, address = email.get("name", ""), email.get("address", "")
    return f"{name} <{address}>" if name and name != address else address


def _summary(msg: dict) -> dict:
    return {
        "id": msg["id"],
        "conversation_id": msg.get("conversationId"),
        "from": _addr(msg.get("from")),
        "from_email": (msg.get("from") or {}).get("emailAddress", {}).get("address", ""),
        "to": [_addr(r) for r in msg.get("toRecipients", [])],
        "subject": msg.get("subject") or "(no subject)",
        "date": msg.get("receivedDateTime", ""),
        "snippet": msg.get("bodyPreview", ""),
        "unread": not msg.get("isRead", True),
        "has_attachments": msg.get("hasAttachments", False),
        "folder_id": msg.get("parentFolderId"),
        "link": msg.get("webLink"),
    }


def _recipients(emails: list[str]) -> list[dict]:
    return [{"emailAddress": {"address": e}} for e in emails]


def search(
    query: str = "",
    account: str | None = None,
    limit: int = 10,
    folder: str | None = None,
    unread_only: bool = False,
) -> list[dict]:
    """Search messages. KQL query (from:x subject:y) or plain terms; empty
    query lists the folder (default inbox) newest-first."""
    limit = min(max(limit, 1), 50)
    base = f"/me/mailFolders/{folder}/messages" if folder else "/me/messages"
    params: dict[str, Any] = {"$select": _SELECT, "$top": limit}
    if query:
        # $search cannot combine with $orderby or $filter (Graph restriction).
        params["$search"] = f'"{query}"'
    else:
        params["$orderby"] = "receivedDateTime desc"
        if unread_only:
            params["$filter"] = "isRead eq false"
        if not folder:
            base = "/me/mailFolders/inbox/messages"
    msgs = graph.get_all(base, account=account, params=params, limit=limit)
    out = [_summary(m) for m in msgs]
    if query and unread_only:
        out = [m for m in out if m["unread"]]
    return out


def read(
    message_id: str,
    account: str | None = None,
    keep_html: bool = False,
    full_thread: bool = False,
) -> dict:
    """Read one message or its whole conversation. Body as plain text unless keep_html."""
    headers = None if keep_html else _TEXT_BODY
    msg = graph.request(
        "GET", f"/me/messages/{message_id}", account=account,
        params={"$select": _SELECT + ",body"}, headers=headers,
    )
    if not full_thread:
        return {**_summary(msg), "body": (msg.get("body") or {}).get("content", "")}

    conv_id = msg.get("conversationId")
    # Graph rejects $filter (conversationId) + $orderby (receivedDateTime) as
    # "too complex" when the orderby property isn't the filter property. Fetch
    # the conversation, sort oldest-first client-side.
    thread = graph.get_all(
        "/me/messages",
        account=account,
        params={
            "$select": _SELECT + ",body",
            "$filter": f"conversationId eq '{conv_id}'",
        },
        limit=50,
        headers=headers,
    )
    thread.sort(key=lambda m: m.get("receivedDateTime") or "")
    messages = [
        {**_summary(m), "body": (m.get("body") or {}).get("content", "")} for m in thread
    ]
    return {"conversation_id": conv_id, "message_count": len(messages), "messages": messages}


def _message_payload(
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    reply_to: str | None = None,
) -> dict:
    msg: dict[str, Any] = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": _recipients(to),
    }
    if cc:
        msg["ccRecipients"] = _recipients(cc)
    if bcc:
        msg["bccRecipients"] = _recipients(bcc)
    if reply_to:
        msg["replyTo"] = _recipients([reply_to])
    return msg


def send(
    to: list[str],
    subject: str,
    body: str,
    account: str | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    reply_to: str | None = None,
    dry_run: bool = False,
) -> dict:
    """Send mail. DESTRUCTIVE — irreversible once sent."""
    if dry_run:
        return {
            "dry_run": True,
            "to": to,
            "subject": subject,
            "cc": cc,
            "body_preview": body[:500],
            "status": "NOT SENT — dry_run=True",
        }
    graph.request(
        "POST", "/me/sendMail", account=account,
        json_body={
            "message": _message_payload(to, subject, body, cc=cc, bcc=bcc, reply_to=reply_to),
            "saveToSentItems": True,
        },
    )
    return {"to": to, "subject": subject, "status": "sent"}


def draft(
    to: list[str],
    subject: str,
    body: str,
    account: str | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
) -> dict:
    """Create a draft in Drafts. Does not send."""
    created = graph.request(
        "POST", "/me/messages", account=account,
        json_body=_message_payload(to, subject, body, cc=cc, bcc=bcc),
    )
    return {"draft_id": created.get("id"), "status": "draft"}


def reply(
    message_id: str,
    body: str,
    account: str | None = None,
    reply_all: bool = False,
    dry_run: bool = False,
) -> dict:
    """Reply to a message. DESTRUCTIVE — sends immediately. Graph preserves
    thread, headers, and recipients server-side."""
    if dry_run:
        original = graph.request(
            "GET", f"/me/messages/{message_id}", account=account,
            params={"$select": _SELECT},
        )
        s = _summary(original)
        return {
            "dry_run": True,
            "to": [s["from"]] + (s["to"] if reply_all else []),
            "subject": f"Re: {s['subject']}",
            "body_preview": body[:500],
            "in_reply_to_conversation": s["conversation_id"],
            "status": "NOT SENT — dry_run=True",
        }
    action = "replyAll" if reply_all else "reply"
    graph.request(
        "POST", f"/me/messages/{message_id}/{action}", account=account,
        json_body={"comment": body},
    )
    return {"message_id": message_id, "status": "sent"}


def folders_list(account: str | None = None) -> list[dict]:
    """List mail folders with unread/total counts."""
    folders = graph.get_all(
        "/me/mailFolders", account=account,
        params={
            "$select": "id,displayName,unreadItemCount,totalItemCount,childFolderCount",
            "$top": 100,
        },
        limit=100,
    )
    return [
        {
            "id": f["id"],
            "name": f.get("displayName"),
            "unread": f.get("unreadItemCount"),
            "total": f.get("totalItemCount"),
            "child_folders": f.get("childFolderCount"),
        }
        for f in folders
    ]


def move(message_ids: list[str], destination: str, account: str | None = None) -> dict:
    """Move messages to a folder (well-known name like 'archive' or a folder id)."""
    for mid in message_ids:
        graph.request(
            "POST", f"/me/messages/{mid}/move", account=account,
            json_body={"destinationId": destination},
        )
    return {"count": len(message_ids), "destination": destination, "status": "moved"}


def archive(message_ids: list[str], account: str | None = None) -> dict:
    """Move messages to the Archive folder. No deletion."""
    return move(message_ids, "archive", account=account)


def mark_read(message_ids: list[str], read: bool = True, account: str | None = None) -> dict:
    """Mark messages read/unread."""
    for mid in message_ids:
        graph.request(
            "PATCH", f"/me/messages/{mid}", account=account,
            json_body={"isRead": read},
        )
    return {"count": len(message_ids), "read": read, "status": "updated"}


_DOWNLOADS_DIR = Path.home() / ".claude" / "microsoft-365-mcp" / "downloads"
_ATTACH_SELECT = "id,name,contentType,size,isInline"


def attachments_list(message_id: str, account: str | None = None) -> list[dict]:
    """List a message's attachments: filename, mime type, size, and the
    attachment_id attachment_save needs. Metadata only — no bytes fetched."""
    items = graph.get_all(
        f"/me/messages/{message_id}/attachments",
        account=account, params={"$select": _ATTACH_SELECT}, limit=50,
    )
    return [
        {
            "attachment_id": a["id"],
            "filename": a.get("name", ""),
            "mime_type": a.get("contentType", "application/octet-stream"),
            "size": a.get("size", 0),
            "inline": a.get("isInline", False),
        }
        for a in items
    ]


def attachment_save(
    message_id: str,
    attachment_id: str,
    filename: str | None = None,
    dest_dir: str | None = None,
    account: str | None = None,
) -> dict:
    """Download one attachment to a local file and return its path.

    Call attachments_list first for attachment_id and the original filename.
    Saves under ~/.claude/microsoft-365-mcp/downloads by default; pass
    dest_dir to save elsewhere. The returned path can be read directly (PDF,
    image, docx, ...) — this tool does not parse the content, only fetches it.
    """
    # No $select here. contentBytes is declared on microsoft.graph.fileAttachment,
    # not on the microsoft.graph.attachment base type this path resolves to, so
    # naming it in $select fails the WHOLE request before Graph ever looks at the
    # attachment: 400 "Could not find a property named 'contentBytes' on type
    # 'microsoft.graph.attachment'". Fetching the resource plain returns the full
    # fileAttachment, contentBytes included.
    meta = graph.request(
        "GET", f"/me/messages/{message_id}/attachments/{attachment_id}",
        account=account,
    )
    content_bytes = meta.get("contentBytes")
    if content_bytes is not None:
        raw = base64.b64decode(content_bytes)
    else:
        # Graph omits contentBytes above ~3 MB — read the raw stream instead.
        raw = graph.request(
            "GET", f"/me/messages/{message_id}/attachments/{attachment_id}/$value",
            account=account, raw=True,
        )

    out_dir = Path(dest_dir) if dest_dir else _DOWNLOADS_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    name = filename or meta.get("name") or f"attachment-{attachment_id[:8]}"
    path = out_dir / name
    path.write_bytes(raw)
    return {"path": str(path), "filename": name, "bytes": len(raw)}
