"""Outlook Calendar tool implementations — Microsoft Graph wrappers.

Token-efficient by design: list_events returns compact event shape without
attendee lists, body, or recurrence detail unless verbose=True.

calendar_id "primary" maps to the user's default calendar. Event listing uses
/calendarView, which expands recurring events into instances (the Graph
equivalent of Google Calendar's singleEvents=True).
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from typing import Any

import graph

DEFAULT_TZ = "America/Bogota"

_RESPONSES = {"accepted": "accept", "declined": "decline", "tentative": "tentativelyAccept"}


def _parse_time(value: str) -> str:
    """Normalize a user-supplied time string to ISO 8601 for Graph."""
    value = value.strip()
    if not value:
        raise ValueError("Empty time value")

    shortcuts = {
        "now": datetime.now(timezone.utc),
        "today": datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0),
        "tomorrow": datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
        + timedelta(days=1),
    }
    if value.lower() in shortcuts:
        return shortcuts[value.lower()].isoformat()

    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt.isoformat()
    except ValueError as e:
        raise ValueError(
            f"Can't parse time '{value}'. Use ISO 8601 or 'now'/'today'/'tomorrow'."
        ) from e


def _events_base(calendar_id: str) -> str:
    if calendar_id in ("primary", "", None):
        return "/me"
    return f"/me/calendars/{calendar_id}"


def _summarize_event(event: dict, verbose: bool = False) -> dict:
    attendees = event.get("attendees", []) or []
    out: dict[str, Any] = {
        "id": event["id"],
        "summary": event.get("subject") or "(no title)",
        "start": (event.get("start") or {}).get("dateTime"),
        "end": (event.get("end") or {}).get("dateTime"),
        "time_zone": (event.get("start") or {}).get("timeZone"),
        "status": event.get("showAs"),
        "my_response": (event.get("responseStatus") or {}).get("response"),
        "organizer": ((event.get("organizer") or {}).get("emailAddress") or {}).get("address"),
        "location": (event.get("location") or {}).get("displayName"),
        "attendee_count": len(attendees),
        "meeting_link": (event.get("onlineMeeting") or {}).get("joinUrl"),
        "link": event.get("webLink"),
        "is_all_day": event.get("isAllDay", False),
    }
    if verbose:
        out["description"] = event.get("bodyPreview", "")
        out["attendees"] = [
            {
                "email": (a.get("emailAddress") or {}).get("address"),
                "response": (a.get("status") or {}).get("response"),
                "type": a.get("type"),
            }
            for a in attendees
        ]
        out["recurrence"] = event.get("recurrence")
    return out


_SELECT = (
    "id,subject,start,end,showAs,responseStatus,organizer,location,attendees,"
    "onlineMeeting,webLink,isAllDay,bodyPreview,recurrence"
)


def list_calendars(account: str | None = None) -> list[dict]:
    cals = graph.get_all(
        "/me/calendars", account=account,
        params={"$select": "id,name,isDefaultCalendar,canEdit,owner"},
        limit=100,
    )
    return [
        {
            "id": c["id"],
            "summary": c.get("name"),
            "primary": c.get("isDefaultCalendar", False),
            "can_edit": c.get("canEdit"),
            "owner": ((c.get("owner") or {}).get("address")),
        }
        for c in cals
    ]


def list_events(
    account: str | None = None,
    calendar_id: str = "primary",
    time_min: str = "now",
    time_max: str | None = None,
    days_ahead: int = 7,
    query: str | None = None,
    max_results: int = 25,
    verbose: bool = False,
) -> list[dict]:
    t_min = _parse_time(time_min)
    if time_max:
        t_max = _parse_time(time_max)
    else:
        t_max = (
            datetime.fromisoformat(t_min.replace("Z", "+00:00")) + timedelta(days=days_ahead)
        ).isoformat()

    cap = min(max(max_results, 1), 100)
    # calendarView rejects contains()/$filter on subject; over-fetch the window
    # when a query is present and match the title client-side (case-insensitive).
    fetch = 100 if query else cap
    params: dict[str, Any] = {
        "startDateTime": t_min,
        "endDateTime": t_max,
        "$select": _SELECT,
        "$orderby": "start/dateTime",
        "$top": fetch,
    }

    events = graph.get_all(
        f"{_events_base(calendar_id)}/calendarView", account=account,
        params=params, limit=fetch,
    )
    if query:
        q = query.lower()
        events = [e for e in events if q in (e.get("subject") or "").lower()][:cap]
    return [_summarize_event(e, verbose=verbose) for e in events]


def create_event(
    summary: str,
    start: str,
    end: str,
    account: str | None = None,
    calendar_id: str = "primary",
    description: str | None = None,
    location: str | None = None,
    attendees: list[str] | None = None,
    time_zone: str = DEFAULT_TZ,
    add_teams_meeting: bool = False,
) -> dict:
    body: dict[str, Any] = {
        "subject": summary,
        "start": {"dateTime": _parse_time(start), "timeZone": time_zone},
        "end": {"dateTime": _parse_time(end), "timeZone": time_zone},
    }
    if description:
        body["body"] = {"contentType": "Text", "content": description}
    if location:
        body["location"] = {"displayName": location}
    if attendees:
        body["attendees"] = [
            {"emailAddress": {"address": a}, "type": "required"} for a in attendees
        ]
    if add_teams_meeting:
        body["isOnlineMeeting"] = True
        body["onlineMeetingProvider"] = "teamsForBusiness"

    path = "/me/events" if calendar_id in ("primary", "", None) else f"/me/calendars/{calendar_id}/events"
    event = graph.request("POST", path, account=account, json_body=body)
    return _summarize_event(event, verbose=True)


def update_event(
    event_id: str,
    account: str | None = None,
    summary: str | None = None,
    start: str | None = None,
    end: str | None = None,
    description: str | None = None,
    location: str | None = None,
    attendees_add: list[str] | None = None,
    attendees_remove: list[str] | None = None,
    time_zone: str = DEFAULT_TZ,
) -> dict:
    """Partial-update: only fields passed are changed. Attendee changes are
    merged against the event's current attendee list."""
    patch: dict[str, Any] = {}
    if summary is not None:
        patch["subject"] = summary
    if start is not None:
        patch["start"] = {"dateTime": _parse_time(start), "timeZone": time_zone}
    if end is not None:
        patch["end"] = {"dateTime": _parse_time(end), "timeZone": time_zone}
    if description is not None:
        patch["body"] = {"contentType": "Text", "content": description}
    if location is not None:
        patch["location"] = {"displayName": location}

    if attendees_add or attendees_remove:
        event = graph.request(
            "GET", f"/me/events/{event_id}", account=account,
            params={"$select": "attendees"},
        )
        current = {
            ((a.get("emailAddress") or {}).get("address") or "").lower(): a
            for a in event.get("attendees", []) or []
        }
        for email in attendees_add or []:
            current.setdefault(
                email.lower(), {"emailAddress": {"address": email}, "type": "required"}
            )
        for email in attendees_remove or []:
            current.pop(email.lower(), None)
        patch["attendees"] = list(current.values())

    updated = graph.request("PATCH", f"/me/events/{event_id}", account=account, json_body=patch)
    return _summarize_event(updated, verbose=True)


def delete_event(event_id: str, account: str | None = None) -> dict:
    graph.request("DELETE", f"/me/events/{event_id}", account=account)
    return {"event_id": event_id, "status": "deleted"}


def freebusy(
    time_min: str,
    time_max: str,
    account: str | None = None,
    emails: list[str] | None = None,
) -> dict:
    """Busy windows for one or more people via getSchedule. emails defaults to
    the authenticated account itself."""
    from accounts import default_account

    schedules = emails or [account or default_account()]
    resp = graph.request(
        "POST", "/me/calendar/getSchedule", account=account,
        json_body={
            "schedules": schedules,
            "startTime": {"dateTime": _parse_time(time_min), "timeZone": "UTC"},
            "endTime": {"dateTime": _parse_time(time_max), "timeZone": "UTC"},
            "availabilityViewInterval": 30,
        },
    )
    out: dict[str, list] = {}
    for sched in resp.get("value", []):
        out[sched.get("scheduleId", "")] = [
            {
                "start": (i.get("start") or {}).get("dateTime"),
                "end": (i.get("end") or {}).get("dateTime"),
                "status": i.get("status"),
            }
            for i in sched.get("scheduleItems", [])
            if i.get("status") != "free"
        ]
    return out


def respond(
    event_id: str,
    response: str,
    account: str | None = None,
    comment: str | None = None,
) -> dict:
    """Respond accepted / declined / tentative to an event you're invited to."""
    if response not in _RESPONSES:
        raise ValueError("response must be one of: accepted, declined, tentative")
    body: dict[str, Any] = {"sendResponse": True}
    if comment:
        body["comment"] = comment
    graph.request(
        "POST", f"/me/events/{event_id}/{_RESPONSES[response]}", account=account,
        json_body=body,
    )
    return {"event_id": event_id, "response": response, "status": "responded"}
