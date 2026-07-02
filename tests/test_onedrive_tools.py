"""OneDrive tools against the fake Graph transport."""

from __future__ import annotations

import pytest

import onedrive_tools

ITEM = {
    "id": "IT1",
    "name": "notes.txt",
    "file": {"mimeType": "text/plain"},
    "size": 42,
    "lastModifiedDateTime": "2026-07-01T15:00:00Z",
    "parentReference": {"id": "root0"},
    "webUrl": "https://onedrive.example/1",
}

FOLDER = {
    "id": "F1",
    "name": "Projects",
    "folder": {"childCount": 3},
    "size": 0,
    "lastModifiedDateTime": "2026-07-01T15:00:00Z",
    "parentReference": {"id": "root0"},
    "webUrl": "https://onedrive.example/f1",
}


def test_search_escapes_quotes_and_returns_summaries(fake_graph):
    fake_graph.queue("GET", "/me/drive/root/search(q='o''brien report')", {"value": [ITEM]})
    out = onedrive_tools.search("o'brien report")
    assert out == [{
        "id": "IT1", "name": "notes.txt", "mime": "text/plain", "is_folder": False,
        "modified": "2026-07-01T15:00:00Z", "size": 42, "parent_id": "root0",
        "link": "https://onedrive.example/1",
    }]


def test_read_file_metadata_only_by_default(fake_graph):
    fake_graph.queue("GET", "/me/drive/items/IT1", ITEM)
    out = onedrive_tools.read_file("IT1")
    assert "content" not in out
    assert len(fake_graph.calls) == 1


def test_read_file_content_optin_and_truncation(fake_graph):
    fake_graph.queue("GET", "/me/drive/items/IT1", ITEM)
    fake_graph.queue("GET", "/me/drive/items/IT1/content", b"hello world")
    out = onedrive_tools.read_file("IT1", include_content=True, max_chars=5)
    assert out["content"] == "hello"
    assert out["truncated"] is True
    assert out["bytes"] == 11


def test_read_file_binary_placeholder(fake_graph):
    fake_graph.queue("GET", "/me/drive/items/IT1", ITEM)
    fake_graph.queue("GET", "/me/drive/items/IT1/content", b"\xff\xfe\x00binary")
    out = onedrive_tools.read_file("IT1", include_content=True)
    assert "binary" in out["content"]
    assert out["truncated"] is False


def test_list_folder_root_shortcut(fake_graph):
    fake_graph.queue("GET", "/me/drive/root/children", {"value": [FOLDER]})
    out = onedrive_tools.list_folder()
    assert out[0]["is_folder"] is True
    assert out[0]["mime"] == "folder"


def test_create_folder_payload(fake_graph):
    fake_graph.queue("POST", "/me/drive/root/children", FOLDER)
    onedrive_tools.create_folder("Projects")
    body = fake_graph.calls[0]["json_body"]
    assert body["folder"] == {}
    assert body["@microsoft.graph.conflictBehavior"] == "rename"


def test_upload_small_file_simple_put(fake_graph, tmp_path):
    src = tmp_path / "small.txt"
    src.write_bytes(b"tiny")
    fake_graph.queue("PUT", "/me/drive/root:/small.txt:/content", ITEM)
    out = onedrive_tools.upload(str(src))
    assert out["id"] == "IT1"
    assert fake_graph.calls[0]["data"] == b"tiny"


def test_upload_large_file_uses_session_chunks(fake_graph, tmp_path, monkeypatch):
    monkeypatch.setattr(onedrive_tools, "_SIMPLE_UPLOAD_MAX", 4)
    monkeypatch.setattr(onedrive_tools, "_CHUNK", 4)
    src = tmp_path / "big.bin"
    src.write_bytes(b"123456")  # 6 bytes -> 2 chunks of 4 + 2
    fake_graph.queue("POST", "/me/drive/root:/big.bin:/createUploadSession",
                     {"uploadUrl": "https://up.example/session1"})
    fake_graph.queue("PUT", "https://up.example/session1", {})
    fake_graph.queue("PUT", "https://up.example/session1", ITEM)
    out = onedrive_tools.upload(str(src))
    assert out["id"] == "IT1"
    ranges = [c["headers"]["Content-Range"] for c in fake_graph.calls[1:]]
    assert ranges == ["bytes 0-3/6", "bytes 4-5/6"]
    # Graph rejects a bearer token on the pre-authenticated uploadUrl — the
    # chunk PUTs MUST pass auth=False (createUploadSession POST stays authed).
    assert fake_graph.calls[0]["auth"] is True  # createUploadSession
    assert all(c["auth"] is False for c in fake_graph.calls[1:])  # chunk PUTs


def test_upload_missing_file_raises():
    with pytest.raises(ValueError):
        onedrive_tools.upload("/nonexistent/file.txt")


def test_move_and_rename(fake_graph):
    fake_graph.queue("PATCH", "/me/drive/items/IT1", ITEM)
    onedrive_tools.move("IT1", "F1")
    assert fake_graph.calls[0]["json_body"] == {"parentReference": {"id": "F1"}}

    fake_graph.queue("PATCH", "/me/drive/items/IT1", ITEM)
    onedrive_tools.rename("IT1", "new.txt")
    assert fake_graph.calls[1]["json_body"] == {"name": "new.txt"}


def test_delete_dry_run_does_not_delete(fake_graph):
    fake_graph.queue("GET", "/me/drive/items/IT1", ITEM)
    out = onedrive_tools.delete("IT1", dry_run=True)
    assert out["dry_run"] is True
    assert all(c["method"] != "DELETE" for c in fake_graph.calls)


def test_delete_hits_recycle_bin(fake_graph, monkeypatch, tmp_path):
    monkeypatch.setattr(onedrive_tools, "_AUDIT_LOG", tmp_path / "audit.log")
    fake_graph.queue("GET", "/me/drive/items/IT1", ITEM)
    fake_graph.queue("DELETE", "/me/drive/items/IT1", {})
    out = onedrive_tools.delete("IT1")
    assert out["status"] == "deleted (recycle bin)"
    assert (tmp_path / "audit.log").exists()


def test_share_validates_role_and_builds_invite(fake_graph):
    with pytest.raises(ValueError):
        onedrive_tools.share("IT1", "a@example.com", role="owner")
    fake_graph.queue("POST", "/me/drive/items/IT1/invite", {"value": [{"id": "P1"}]})
    out = onedrive_tools.share("IT1", "a@example.com", role="read")
    assert out["permission_id"] == "P1"
    body = fake_graph.calls[0]["json_body"]
    assert body["recipients"] == [{"email": "a@example.com"}]
    assert body["roles"] == ["read"]
    assert body["requireSignIn"] is True


def test_share_link_validation_and_shape(fake_graph):
    with pytest.raises(ValueError):
        onedrive_tools.share_link("IT1", link_type="embed")
    with pytest.raises(ValueError):
        onedrive_tools.share_link("IT1", scope="world")
    fake_graph.queue("POST", "/me/drive/items/IT1/createLink",
                     {"id": "P2", "link": {"webUrl": "https://share.example/x"}})
    out = onedrive_tools.share_link("IT1", link_type="edit")
    assert out["link"] == "https://share.example/x"
    assert out["permission_id"] == "P2"


def test_permission_list_grantee_extraction(fake_graph):
    fake_graph.queue("GET", "/me/drive/items/IT1/permissions", {"value": [
        {"id": "P1", "roles": ["write"],
         "grantedToV2": {"user": {"email": "a@example.com"}}},
        {"id": "P2", "roles": ["read"], "link": {"scope": "organization",
                                                  "webUrl": "https://share.example/x"}},
    ]})
    out = onedrive_tools.permission_list("IT1")
    assert out[0] == {"id": "P1", "roles": ["write"], "grantee": "a@example.com",
                      "is_link": False, "link": None}
    assert out[1]["is_link"] is True
    assert out[1]["grantee"] == "organization"


def test_permission_update_and_delete(fake_graph, monkeypatch, tmp_path):
    monkeypatch.setattr(onedrive_tools, "_AUDIT_LOG", tmp_path / "audit.log")
    with pytest.raises(ValueError):
        onedrive_tools.permission_update("IT1", "P1", role="admin")
    fake_graph.queue("PATCH", "/me/drive/items/IT1/permissions/P1", {"roles": ["read"]})
    out = onedrive_tools.permission_update("IT1", "P1", role="read")
    assert out["roles"] == ["read"]

    fake_graph.queue("DELETE", "/me/drive/items/IT1/permissions/P1", {})
    out = onedrive_tools.permission_delete("IT1", "P1")
    assert out["status"] == "deleted"
