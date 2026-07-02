"""Calendar tools against the fake Graph transport."""

from __future__ import annotations

import pytest

import calendar_tools

EVENT = {
    "id": "EV1",
    "subject": "Sync",
    "start": {"dateTime": "2026-07-03T10:00:00", "timeZone": "America/Bogota"},
    "end": {"dateTime": "2026-07-03T10:30:00", "timeZone": "America/Bogota"},
    "showAs": "busy",
    "responseStatus": {"response": "organizer"},
    "organizer": {"emailAddress": {"address": "me@example.com"}},
    "location": {"displayName": "Room 1"},
    "attendees": [
        {"emailAddress": {"address": "a@example.com"}, "type": "required",
         "status": {"response": "accepted"}},
    ],
    "onlineMeeting": {"joinUrl": "https://teams.example/j/1"},
    "webLink": "https://outlook.example/ev1",
    "isAllDay": False,
    "bodyPreview": "agenda",
}


def test_parse_time_shortcuts_and_iso():
    assert calendar_tools._parse_time("2026-07-03T10:00:00Z").startswith("2026-07-03T10:00:00")
    assert calendar_tools._parse_time("now")
    with pytest.raises(ValueError):
        calendar_tools._parse_time("next tuesday-ish")


def test_list_events_uses_calendar_view(fake_graph):
    fake_graph.queue("GET", "/me/calendarView", {"value": [EVENT]})
    out = calendar_tools.list_events(time_min="2026-07-03T00:00:00Z", days_ahead=1)
    assert out[0]["summary"] == "Sync"
    assert out[0]["meeting_link"] == "https://teams.example/j/1"
    assert "attendees" not in out[0]  # compact by default
    call = fake_graph.calls[0]
    assert call["path"] == "/me/calendarView"
    assert call["params"]["startDateTime"].startswith("2026-07-03")


def test_list_events_verbose_includes_attendees(fake_graph):
    fake_graph.queue("GET", "/me/calendarView", {"value": [EVENT]})
    out = calendar_tools.list_events(verbose=True)
    assert out[0]["attendees"] == [
        {"email": "a@example.com", "response": "accepted", "type": "required"}
    ]


def test_create_event_payload(fake_graph):
    fake_graph.queue("POST", "/me/events", EVENT)
    calendar_tools.create_event(
        summary="Sync", start="2026-07-03T10:00:00", end="2026-07-03T10:30:00",
        attendees=["a@example.com"], add_teams_meeting=True, location="Room 1",
    )
    body = fake_graph.calls[0]["json_body"]
    assert body["subject"] == "Sync"
    assert body["start"]["timeZone"] == "America/Bogota"
    assert body["attendees"] == [
        {"emailAddress": {"address": "a@example.com"}, "type": "required"}
    ]
    assert body["isOnlineMeeting"] is True
    assert body["onlineMeetingProvider"] == "teamsForBusiness"
    assert body["location"] == {"displayName": "Room 1"}


def test_update_event_patches_only_changed_fields(fake_graph):
    fake_graph.queue("PATCH", "/me/events/EV1", EVENT)
    calendar_tools.update_event("EV1", summary="New title")
    call = fake_graph.calls[0]
    assert call["method"] == "PATCH"
    assert call["json_body"] == {"subject": "New title"}


def test_update_event_merges_attendees(fake_graph):
    fake_graph.queue("GET", "/me/events/EV1", {"attendees": EVENT["attendees"]})
    fake_graph.queue("PATCH", "/me/events/EV1", EVENT)
    calendar_tools.update_event(
        "EV1", attendees_add=["b@example.com"], attendees_remove=["a@example.com"],
    )
    patch = fake_graph.calls[1]["json_body"]
    emails = [a["emailAddress"]["address"] for a in patch["attendees"]]
    assert emails == ["b@example.com"]


def test_delete_event(fake_graph):
    fake_graph.queue("DELETE", "/me/events/EV1", {})
    assert calendar_tools.delete_event("EV1") == {"event_id": "EV1", "status": "deleted"}


def test_freebusy_maps_schedule_items_and_drops_free(fake_graph):
    fake_graph.queue("POST", "/me/calendar/getSchedule", {"value": [
        {"scheduleId": "a@example.com", "scheduleItems": [
            {"start": {"dateTime": "T1"}, "end": {"dateTime": "T2"}, "status": "busy"},
            {"start": {"dateTime": "T3"}, "end": {"dateTime": "T4"}, "status": "free"},
        ]},
    ]})
    out = calendar_tools.freebusy(
        "2026-07-03T00:00:00Z", "2026-07-04T00:00:00Z", emails=["a@example.com"],
    )
    assert out == {"a@example.com": [{"start": "T1", "end": "T2", "status": "busy"}]}


def test_respond_validates_and_maps_action(fake_graph):
    with pytest.raises(ValueError):
        calendar_tools.respond("EV1", "maybe")
    fake_graph.queue("POST", "/me/events/EV1/tentativelyAccept", {})
    out = calendar_tools.respond("EV1", "tentative", comment="might be late")
    assert out["status"] == "responded"
    assert fake_graph.calls[0]["json_body"] == {"sendResponse": True, "comment": "might be late"}
