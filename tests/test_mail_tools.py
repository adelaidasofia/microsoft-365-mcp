"""Outlook mail tools against the fake Graph transport."""

from __future__ import annotations

import mail_tools

MSG = {
    "id": "AAMk1",
    "conversationId": "CONV1",
    "subject": "Q3 invoice",
    "bodyPreview": "attached the invoice",
    "from": {"emailAddress": {"name": "Alex Doe", "address": "alex@example.com"}},
    "toRecipients": [{"emailAddress": {"name": "", "address": "me@example.com"}}],
    "receivedDateTime": "2026-07-01T15:00:00Z",
    "isRead": False,
    "hasAttachments": True,
    "parentFolderId": "inboxid",
    "webLink": "https://outlook.example/1",
}


def test_search_compact_shape_and_inbox_default(fake_graph):
    fake_graph.queue("GET", "/me/mailFolders/inbox/messages", {"value": [MSG]})
    out = mail_tools.search(limit=5)
    assert out == [
        {
            "id": "AAMk1",
            "conversation_id": "CONV1",
            "from": "Alex Doe <alex@example.com>",
            "from_email": "alex@example.com",
            "to": ["me@example.com"],
            "subject": "Q3 invoice",
            "date": "2026-07-01T15:00:00Z",
            "snippet": "attached the invoice",
            "unread": True,
            "has_attachments": True,
            "folder_id": "inboxid",
            "link": "https://outlook.example/1",
        }
    ]
    call = fake_graph.calls[0]
    assert call["params"]["$top"] == 5
    assert "body" not in call["params"]["$select"].split(",")  # bodyPreview only
    assert call["params"]["$orderby"] == "receivedDateTime desc"


def test_search_with_query_uses_search_not_orderby(fake_graph):
    fake_graph.queue("GET", "/me/messages", {"value": []})
    mail_tools.search(query="from:alex@example.com invoice")
    call = fake_graph.calls[0]
    assert call["params"]["$search"] == '"from:alex@example.com invoice"'
    # Graph rejects $search combined with $orderby / $filter.
    assert "$orderby" not in call["params"]
    assert "$filter" not in call["params"]


def test_read_requests_text_body_by_default(fake_graph):
    fake_graph.queue("GET", "/me/messages/AAMk1", {**MSG, "body": {"content": "plain text"}})
    out = mail_tools.read("AAMk1")
    assert out["body"] == "plain text"
    assert fake_graph.calls[0]["headers"] == {"Prefer": 'outlook.body-content-type="text"'}


def test_read_keep_html_skips_prefer_header(fake_graph):
    fake_graph.queue("GET", "/me/messages/AAMk1", {**MSG, "body": {"content": "<b>hi</b>"}})
    mail_tools.read("AAMk1", keep_html=True)
    assert fake_graph.calls[0]["headers"] is None


def test_read_full_thread_filters_by_conversation(fake_graph):
    fake_graph.queue("GET", "/me/messages/AAMk1", {**MSG, "body": {"content": "one"}})
    fake_graph.queue("GET", "/me/messages", {"value": [
        {**MSG, "body": {"content": "one"}},
        {**MSG, "id": "AAMk2", "body": {"content": "two"}},
    ]})
    out = mail_tools.read("AAMk1", full_thread=True)
    assert out["conversation_id"] == "CONV1"
    assert out["message_count"] == 2
    assert [m["body"] for m in out["messages"]] == ["one", "two"]
    assert "conversationId eq 'CONV1'" in fake_graph.calls[1]["params"]["$filter"]


def test_send_payload_shape(fake_graph):
    fake_graph.queue("POST", "/me/sendMail", {})
    out = mail_tools.send(
        to=["a@example.com"], subject="s", body="b",
        cc=["c@example.com"], reply_to="r@example.com",
    )
    assert out["status"] == "sent"
    payload = fake_graph.calls[0]["json_body"]
    msg = payload["message"]
    assert payload["saveToSentItems"] is True
    assert msg["toRecipients"] == [{"emailAddress": {"address": "a@example.com"}}]
    assert msg["ccRecipients"] == [{"emailAddress": {"address": "c@example.com"}}]
    assert msg["replyTo"] == [{"emailAddress": {"address": "r@example.com"}}]
    assert msg["body"] == {"contentType": "Text", "content": "b"}


def test_send_dry_run_makes_no_calls(fake_graph):
    out = mail_tools.send(to=["a@example.com"], subject="s", body="b", dry_run=True)
    assert out["dry_run"] is True
    assert "NOT SENT" in out["status"]
    assert fake_graph.calls == []


def test_reply_all_uses_replyall_action(fake_graph):
    fake_graph.queue("POST", "/me/messages/AAMk1/replyAll", {})
    out = mail_tools.reply("AAMk1", "thanks", reply_all=True)
    assert out["status"] == "sent"
    assert fake_graph.calls[0]["json_body"] == {"comment": "thanks"}


def test_reply_dry_run_fetches_original_only(fake_graph):
    fake_graph.queue("GET", "/me/messages/AAMk1", MSG)
    out = mail_tools.reply("AAMk1", "thanks", dry_run=True)
    assert out["dry_run"] is True
    assert out["subject"] == "Re: Q3 invoice"
    assert len(fake_graph.calls) == 1
    assert fake_graph.calls[0]["method"] == "GET"


def test_move_and_archive(fake_graph):
    fake_graph.queue("POST", "/me/messages/A/move", {})
    fake_graph.queue("POST", "/me/messages/B/move", {})
    out = mail_tools.move(["A", "B"], "deleteditems")
    assert out == {"count": 2, "destination": "deleteditems", "status": "moved"}

    fake_graph.queue("POST", "/me/messages/C/move", {})
    out = mail_tools.archive(["C"])
    assert out["destination"] == "archive"


def test_folders_list_shape(fake_graph):
    fake_graph.queue("GET", "/me/mailFolders", {"value": [
        {"id": "f1", "displayName": "Inbox", "unreadItemCount": 3,
         "totalItemCount": 10, "childFolderCount": 0},
    ]})
    out = mail_tools.folders_list()
    assert out == [{"id": "f1", "name": "Inbox", "unread": 3, "total": 10, "child_folders": 0}]


def test_mark_read_patches_each_message(fake_graph):
    fake_graph.queue("PATCH", "/me/messages/A", {})
    out = mail_tools.mark_read(["A"], read=False)
    assert out["read"] is False
    assert fake_graph.calls[0]["json_body"] == {"isRead": False}
