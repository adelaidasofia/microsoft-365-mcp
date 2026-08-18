"""Outlook attachment tools against the fake Graph transport.

No unit test in this repo has exercised binary download before this: every
other tool round-trips JSON. attachment_save's real risk is the base64 path —
Graph's small-attachment shape (contentBytes, standard b64) and the
large-attachment fallback ($value, raw bytes) decode differently, and only
one of the two is exercised if a test only covers the common case.
"""

from __future__ import annotations

import base64

import mail_tools

ATTACH_LIST_RESP = {
    "value": [
        {"id": "AAId1", "name": "invoice.pdf", "contentType": "application/pdf",
         "size": 40231, "isInline": False},
        {"id": "AAId2", "name": "logo.png", "contentType": "image/png",
         "size": 900, "isInline": True},
    ]
}


def test_attachments_list_compact_shape(fake_graph):
    fake_graph.queue("GET", "/me/messages/MSG1/attachments", ATTACH_LIST_RESP)
    out = mail_tools.attachments_list("MSG1")
    assert out == [
        {"attachment_id": "AAId1", "filename": "invoice.pdf",
         "mime_type": "application/pdf", "size": 40231, "inline": False},
        {"attachment_id": "AAId2", "filename": "logo.png",
         "mime_type": "image/png", "size": 900, "inline": True},
    ]


def test_attachment_save_small_file_decodes_content_bytes(fake_graph, tmp_path):
    raw = b"%PDF-1.4 fake invoice bytes"
    fake_graph.queue(
        "GET", "/me/messages/MSG1/attachments/AAId1",
        {"name": "invoice.pdf", "contentType": "application/pdf",
         "contentBytes": base64.b64encode(raw).decode(), "size": len(raw)},
    )
    result = mail_tools.attachment_save(
        "MSG1", "AAId1", dest_dir=str(tmp_path)
    )
    saved = tmp_path / "invoice.pdf"
    assert saved.read_bytes() == raw
    assert result == {"path": str(saved), "filename": "invoice.pdf", "bytes": len(raw)}


def test_attachment_save_large_file_falls_back_to_raw_value(fake_graph, tmp_path):
    # Graph omits contentBytes above ~3MB — no key at all, not an empty string.
    raw = b"large binary payload" * 1000
    fake_graph.queue(
        "GET", "/me/messages/MSG1/attachments/AAId3",
        {"name": "recording.mp4", "contentType": "video/mp4", "size": len(raw)},
    )
    fake_graph.queue(
        "GET", "/me/messages/MSG1/attachments/AAId3/$value", raw
    )
    result = mail_tools.attachment_save("MSG1", "AAId3", dest_dir=str(tmp_path))
    saved = tmp_path / "recording.mp4"
    assert saved.read_bytes() == raw
    assert result["bytes"] == len(raw)


def test_attachment_save_uses_explicit_filename_over_graph_name(fake_graph, tmp_path):
    raw = b"content"
    fake_graph.queue(
        "GET", "/me/messages/MSG1/attachments/AAId1",
        {"name": "invoice.pdf", "contentBytes": base64.b64encode(raw).decode()},
    )
    result = mail_tools.attachment_save(
        "MSG1", "AAId1", filename="renamed.pdf", dest_dir=str(tmp_path)
    )
    assert result["filename"] == "renamed.pdf"
    assert (tmp_path / "renamed.pdf").exists()
    assert not (tmp_path / "invoice.pdf").exists()


def test_attachment_save_falls_back_to_placeholder_name_with_no_name_anywhere(
    fake_graph, tmp_path
):
    raw = b"mystery bytes"
    fake_graph.queue(
        "GET", "/me/messages/MSG1/attachments/AAId9",
        {"contentBytes": base64.b64encode(raw).decode()},
    )
    result = mail_tools.attachment_save("MSG1", "AAId9", dest_dir=str(tmp_path))
    assert result["filename"] == "attachment-AAId9"
    assert (tmp_path / "attachment-AAId9").read_bytes() == raw
