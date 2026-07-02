"""OneDrive tool implementations — multi-account, token-efficient.

Design rules (mirror google-workspace-mcp drive_tools):
- Search returns metadata only: id, name, mime, modified, size, link. Never
  dumps content.
- `onedrive_read_file` downloads content only when explicitly asked.
- item_id "root" works anywhere an item id is taken.
- Deletes go to the OneDrive recycle bin (recoverable), mirroring drive_trash.
"""

from __future__ import annotations

import datetime
import json
import pathlib
from pathlib import Path
from typing import Any

import graph

_AUDIT_LOG = pathlib.Path.home() / ".claude" / "microsoft-365-mcp" / "audit.log"

_SELECT = "id,name,file,folder,size,lastModifiedDateTime,parentReference,webUrl"

# Simple-PUT ceiling. Larger files go through an upload session (chunked).
_SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024
_CHUNK = 5 * 1024 * 1024  # must be a multiple of 320 KiB; 5 MiB is safe

_SHARE_ROLES = {"read", "write"}


def _audit(action: str, detail: str) -> None:
    ts = datetime.datetime.now(datetime.timezone.utc).isoformat()
    _AUDIT_LOG.parent.mkdir(parents=True, exist_ok=True)
    with open(_AUDIT_LOG, "a") as f:
        f.write(f"{ts}\t{action}\t{detail}\n")


def _summary(item: dict) -> dict:
    return {
        "id": item["id"],
        "name": item.get("name"),
        "mime": (item.get("file") or {}).get("mimeType") or ("folder" if "folder" in item else None),
        "is_folder": "folder" in item,
        "modified": item.get("lastModifiedDateTime"),
        "size": item.get("size"),
        "parent_id": (item.get("parentReference") or {}).get("id"),
        "link": item.get("webUrl"),
    }


def _item_path(item_id: str) -> str:
    return "/me/drive/root" if item_id in ("root", "", None) else f"/me/drive/items/{item_id}"


def search(query: str, account: str | None = None, limit: int = 10) -> list[dict]:
    """Search across the user's OneDrive (filenames + content)."""
    safe = query.replace("'", "''")
    items = graph.get_all(
        f"/me/drive/root/search(q='{safe}')", account=account,
        params={"$select": _SELECT, "$top": min(max(limit, 1), 50)},
        limit=min(max(limit, 1), 50),
    )
    return [_summary(i) for i in items]


def read_file(
    item_id: str,
    account: str | None = None,
    include_content: bool = False,
    max_chars: int = 50_000,
) -> dict:
    """Read item metadata. Content is opt-in via include_content=True (text
    files decode to UTF-8; binaries return a size-only placeholder)."""
    meta = graph.request(
        "GET", _item_path(item_id), account=account, params={"$select": _SELECT}
    )
    if not include_content or "folder" in meta:
        return _summary(meta)

    raw = graph.request("GET", f"{_item_path(item_id)}/content", account=account, raw=True)
    try:
        text = raw.decode("utf-8")
        truncated = len(text) > max_chars
        if truncated:
            text = text[:max_chars]
    except UnicodeDecodeError:
        text = f"(binary, {len(raw)} bytes — download directly or convert first)"
        truncated = False
    return {**_summary(meta), "content": text, "truncated": truncated, "bytes": len(raw)}


def list_folder(
    folder_id: str = "root",
    account: str | None = None,
    limit: int = 50,
) -> list[dict]:
    """List direct children of a folder. 'root' = OneDrive root."""
    items = graph.get_all(
        f"{_item_path(folder_id)}/children", account=account,
        params={"$select": _SELECT, "$top": min(max(limit, 1), 200), "$orderby": "name"},
        limit=min(max(limit, 1), 200),
    )
    return [_summary(i) for i in items]


def create_folder(
    name: str,
    parent_id: str = "root",
    account: str | None = None,
) -> dict:
    created = graph.request(
        "POST", f"{_item_path(parent_id)}/children", account=account,
        json_body={"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"},
    )
    return _summary(created)


def upload(
    local_path: str,
    name: str | None = None,
    parent_id: str = "root",
    account: str | None = None,
) -> dict:
    """Upload a local file. Files over 4 MB go through a chunked upload session."""
    src = Path(local_path).expanduser()
    if not src.is_file():
        raise ValueError(f"Not a file: {local_path}")
    fname = name or src.name
    size = src.stat().st_size

    if size <= _SIMPLE_UPLOAD_MAX:
        created = graph.request(
            "PUT", f"{_item_path(parent_id)}:/{fname}:/content", account=account,
            data=src.read_bytes(), headers={"Content-Type": "application/octet-stream"},
        )
        return _summary(created)

    session = graph.request(
        "POST", f"{_item_path(parent_id)}:/{fname}:/createUploadSession", account=account,
        json_body={"item": {"@microsoft.graph.conflictBehavior": "rename", "name": fname}},
    )
    upload_url = session["uploadUrl"]
    result: dict = {}
    with open(src, "rb") as f:
        offset = 0
        while offset < size:
            chunk = f.read(_CHUNK)
            end = offset + len(chunk) - 1
            # The uploadUrl is pre-authenticated; Graph rejects a bearer token
            # on it (401). auth=False sends no Authorization header while still
            # using graph.request for uniform retry/error handling.
            result = graph.request(
                "PUT", upload_url, account=account, data=chunk, auth=False,
                headers={
                    "Content-Range": f"bytes {offset}-{end}/{size}",
                    "Content-Type": "application/octet-stream",
                },
            )
            offset += len(chunk)
    return _summary(result) if result.get("id") else {"name": fname, "status": "uploaded"}


def move(item_id: str, new_parent_id: str, account: str | None = None) -> dict:
    updated = graph.request(
        "PATCH", _item_path(item_id), account=account,
        json_body={"parentReference": {"id": new_parent_id}},
    )
    return _summary(updated)


def rename(item_id: str, new_name: str, account: str | None = None) -> dict:
    updated = graph.request(
        "PATCH", _item_path(item_id), account=account, json_body={"name": new_name}
    )
    return _summary(updated)


def delete(item_id: str, account: str | None = None, dry_run: bool = False) -> dict:
    """Move an item to the recycle bin (recoverable). DESTRUCTIVE."""
    meta = graph.request(
        "GET", _item_path(item_id), account=account, params={"$select": _SELECT}
    )
    if dry_run:
        return {**_summary(meta), "dry_run": True, "status": "NOT DELETED — dry_run=True"}
    graph.request("DELETE", _item_path(item_id), account=account)
    _audit("onedrive_delete", f"account={account or 'default'} item_id={item_id} name={meta.get('name')!r}")
    return {**_summary(meta), "status": "deleted (recycle bin)"}


def share(
    item_id: str,
    email: str,
    role: str = "write",
    send_notification: bool = False,
    message: str | None = None,
    account: str | None = None,
) -> dict:
    """Grant a person access by email. role: 'read' | 'write'."""
    if role not in _SHARE_ROLES:
        raise ValueError(f"role must be one of: {sorted(_SHARE_ROLES)}")
    resp = graph.request(
        "POST", f"{_item_path(item_id)}/invite", account=account,
        json_body={
            "recipients": [{"email": email}],
            "roles": [role],
            "requireSignIn": True,
            "sendInvitation": send_notification,
            "message": message or "",
        },
    )
    perms = resp.get("value", [])
    return {"item_id": item_id, "email": email, "role": role,
            "permission_id": perms[0].get("id") if perms else None, "status": "shared"}


def share_link(
    item_id: str,
    link_type: str = "view",
    scope: str = "organization",
    account: str | None = None,
) -> dict:
    """Create a sharing link. link_type: 'view' | 'edit'. scope:
    'organization' (people in the tenant) | 'anonymous' (anyone with link;
    corporate tenants often disable it)."""
    if link_type not in ("view", "edit"):
        raise ValueError("link_type must be 'view' or 'edit'")
    if scope not in ("organization", "anonymous"):
        raise ValueError("scope must be 'organization' or 'anonymous'")
    resp = graph.request(
        "POST", f"{_item_path(item_id)}/createLink", account=account,
        json_body={"type": link_type, "scope": scope},
    )
    return {
        "item_id": item_id,
        "link": ((resp.get("link") or {}).get("webUrl")),
        "type": link_type,
        "scope": scope,
        "permission_id": resp.get("id"),
    }


def permission_list(item_id: str, account: str | None = None) -> list[dict]:
    perms = graph.get_all(f"{_item_path(item_id)}/permissions", account=account, limit=100)
    out = []
    for p in perms:
        grantee = (
            ((p.get("grantedToV2") or {}).get("user") or {}).get("email")
            or ((p.get("grantedToV2") or {}).get("user") or {}).get("displayName")
            or ((p.get("link") or {}).get("scope"))
        )
        out.append({
            "id": p.get("id"),
            "roles": p.get("roles", []),
            "grantee": grantee,
            "is_link": "link" in p,
            "link": ((p.get("link") or {}).get("webUrl")),
        })
    return out


def permission_update(
    item_id: str,
    permission_id: str,
    role: str,
    account: str | None = None,
) -> dict:
    """Change an existing permission's role. role: 'read' | 'write'."""
    if role not in _SHARE_ROLES:
        raise ValueError(f"role must be one of: {sorted(_SHARE_ROLES)}")
    updated = graph.request(
        "PATCH", f"{_item_path(item_id)}/permissions/{permission_id}", account=account,
        json_body={"roles": [role]},
    )
    return {"item_id": item_id, "permission_id": permission_id, "roles": updated.get("roles", [role])}


def permission_delete(
    item_id: str,
    permission_id: str,
    account: str | None = None,
) -> dict:
    """Revoke a permission. DESTRUCTIVE — removes access immediately."""
    graph.request(
        "DELETE", f"{_item_path(item_id)}/permissions/{permission_id}", account=account
    )
    _audit("onedrive_permission_delete", f"account={account or 'default'} item_id={item_id} perm={permission_id}")
    return {"item_id": item_id, "permission_id": permission_id, "status": "deleted"}
