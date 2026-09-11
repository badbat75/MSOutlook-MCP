"""Unit tests for all-day events and "Show as" in the calendar tools.

Graph takes an all-day event only as midnight to midnight, the end a day or more
after the start and both in one time zone, and on an update it refuses isAllDay
unless start comes with it, whichever way the flag flips. None of that shows in
a tool schema, so a tool that merely forwarded is_all_day would still fail on
the one call a user makes: "make this all day, and show me as free". These
tests pin what the tools send, against a fake Graph that answers the way the
real one did when probed.
"""

import re

import anyio
import pytest
from pydantic import ValidationError

from outlook_mcp.app import mcp
from outlook_mcp.events import all_day_span, describe_all_day, read_event_times
from outlook_mcp.models import CreateEventInput, UpdateEventInput
from outlook_mcp.tools import calendar

EVENT = "AAMkAGEvent=="


class FakeGraph:
    """Records every call, and knows one event's times in the zones asked for.

    The default is the event a date-only create leaves behind: midnight to
    midnight on 2 October, in Rome. Asked without a Prefer header, Graph answers
    in UTC, where that event starts the evening before.
    """

    def __init__(self, own_zone="Europe/Rome", times=None):
        self.own_zone = own_zone
        self.times = times or {
            "Europe/Rome": ("2026-10-02T00:00:00.0000000", "2026-10-02T00:00:00.0000000"),
            "UTC": ("2026-10-01T22:00:00.0000000", "2026-10-01T22:00:00.0000000"),
        }
        self.calls = []

    async def get(self, endpoint, params=None, headers=None):
        self.calls.append(("GET", endpoint, params, headers))
        if params and params.get("$select") == "originalStartTimeZone":
            return {"originalStartTimeZone": self.own_zone}
        prefer = (headers or {}).get("Prefer", "")
        match = re.fullmatch(r'outlook\.timezone="([^"]+)"', prefer)
        zone = match.group(1) if match else "UTC"
        start, end = self.times[zone]
        return {
            "start": {"dateTime": start, "timeZone": zone},
            "end": {"dateTime": end, "timeZone": zone},
        }

    async def post(self, endpoint, json_data=None):
        self.calls.append(("POST", endpoint, json_data))
        return {"id": "new-event-id", **(json_data or {})}

    async def patch(self, endpoint, json_data=None):
        self.calls.append(("PATCH", endpoint, json_data))
        return {"id": EVENT, **(json_data or {})}

    def sent(self, method):
        """The body of the one call made with this method."""
        bodies = [call[2] for call in self.calls if call[0] == method]
        assert len(bodies) == 1, f"expected one {method}, got {self.calls}"
        return bodies[0]

    def reads(self):
        return [call for call in self.calls if call[0] == "GET"]


@pytest.fixture
def graph(monkeypatch):
    fake = FakeGraph()
    monkeypatch.setattr(calendar, "get_graph", lambda ctx: fake)
    return fake


def update(**fields):
    return anyio.run(calendar.outlook_update_event, UpdateEventInput(event_id=EVENT, **fields), None)


def create(**fields):
    fields = {"subject": "Rata condominiale", **fields}
    return anyio.run(calendar.outlook_create_event, CreateEventInput(**fields), None)


class TestAllDaySpan:
    def test_a_date_alone_is_one_day(self):
        assert all_day_span("2026-10-01", None) == ("2026-10-01T00:00:00", "2026-10-02T00:00:00")

    def test_the_same_date_twice_is_one_day(self):
        # The natural way to ask for one day, and Graph refuses it as it is:
        # "The duration of an event marked as All day must be at least 24 hours."
        assert all_day_span("2026-10-01", "2026-10-01") == ("2026-10-01T00:00:00", "2026-10-02T00:00:00")

    def test_an_end_given_as_a_date_is_the_last_day(self):
        assert all_day_span("2026-10-01", "2026-10-02") == ("2026-10-01T00:00:00", "2026-10-03T00:00:00")

    def test_an_end_at_midnight_is_exclusive(self):
        # How Graph returns a one-day all-day event: passing it back in must
        # not grow it by a day.
        span = ("2026-10-01T00:00:00.0000000", "2026-10-02T00:00:00.0000000")
        assert all_day_span(*span) == ("2026-10-01T00:00:00", "2026-10-02T00:00:00")

    def test_a_timed_event_becomes_the_day_it_is_on(self):
        # Graph refuses the times as they are: "The Event.Start property for an
        # all-day event needs to be set to midnight."
        assert all_day_span("2026-10-01T10:00:00", "2026-10-01T11:00:00") == (
            "2026-10-01T00:00:00", "2026-10-02T00:00:00",
        )

    def test_every_day_a_timed_event_touches_is_in(self):
        assert all_day_span("2026-10-01T22:00:00", "2026-10-02T01:00:00") == (
            "2026-10-01T00:00:00", "2026-10-03T00:00:00",
        )

    def test_midnight_to_midnight_on_one_day_is_that_day(self):
        # What the date-only workaround in the bug report produced: a
        # zero-length event at 00:00.
        span = ("2026-10-02T00:00:00.0000000", "2026-10-02T00:00:00.0000000")
        assert all_day_span(*span) == ("2026-10-02T00:00:00", "2026-10-03T00:00:00")

    def test_an_end_before_the_start_still_leaves_one_day(self):
        assert all_day_span("2026-10-05", "2026-10-01") == ("2026-10-05T00:00:00", "2026-10-06T00:00:00")

    def test_an_offset_does_not_hide_a_midnight(self):
        assert all_day_span("2026-10-01", "2026-10-02T00:00:00Z")[1] == "2026-10-02T00:00:00"

    def test_something_that_is_not_a_date_says_so(self):
        with pytest.raises(ValueError, match="not an ISO date"):
            all_day_span("tomorrow", None)


class TestDescribingTheDays:
    def test_one_day(self):
        assert describe_all_day("2026-10-01T00:00:00", "2026-10-02T00:00:00") == "all day on 2026-10-01"

    def test_several_days_count_the_last_one_in(self):
        assert describe_all_day("2026-10-01T00:00:00", "2026-10-04T00:00:00") == (
            "all day, 2026-10-01 to 2026-10-03 (3 days)"
        )


class TestReadingTheCurrentTimes:
    def test_they_are_read_in_the_event_s_own_zone(self):
        # In UTC this Rome-midnight event starts on the 1st, and making it
        # all-day there would move it a day early.
        fake = FakeGraph()
        start, end, zone = anyio.run(read_event_times, fake, EVENT)
        assert (start[:10], zone) == ("2026-10-02", "Europe/Rome")
        assert fake.reads()[-1][3] == {"Prefer": 'outlook.timezone="Europe/Rome"'}

    def test_a_zone_the_caller_names_wins_and_saves_a_request(self):
        fake = FakeGraph()
        _, _, zone = anyio.run(read_event_times, fake, EVENT, "UTC")
        assert zone == "UTC"
        assert len(fake.reads()) == 1

    def test_an_event_without_a_zone_is_read_in_utc(self):
        fake = FakeGraph(own_zone=None)
        _, _, zone = anyio.run(read_event_times, fake, EVENT)
        assert zone == "UTC"


class TestUpdatingToAllDay:
    def test_the_call_from_the_bug_report_sends_what_graph_needs(self, graph):
        # update(is_all_day=True, show_as="free") and nothing else: Graph
        # answers "Missing parameters: Event.Start" unless start and end go
        # with the flag, at midnight, a day apart, in one zone.
        result = update(is_all_day=True, show_as="free")

        assert graph.sent("PATCH") == {
            "start": {"dateTime": "2026-10-02T00:00:00", "timeZone": "Europe/Rome"},
            "end": {"dateTime": "2026-10-03T00:00:00", "timeZone": "Europe/Rome"},
            "isAllDay": True,
            "showAs": "free",
        }
        assert "all day on 2026-10-02" in result

    def test_a_given_start_is_used_and_nothing_is_read(self, graph):
        update(is_all_day=True, start="2026-11-15", timezone="Europe/Rome")

        assert graph.reads() == []
        body = graph.sent("PATCH")
        assert body["start"] == {"dateTime": "2026-11-15T00:00:00", "timeZone": "Europe/Rome"}
        assert body["end"] == {"dateTime": "2026-11-16T00:00:00", "timeZone": "Europe/Rome"}

    def test_a_given_end_is_the_last_day(self, graph):
        update(is_all_day=True, start="2026-11-15", end="2026-11-16")
        assert graph.sent("PATCH")["end"]["dateTime"] == "2026-11-17T00:00:00"

    def test_the_caller_s_zone_is_where_the_days_are_read(self, graph):
        update(is_all_day=True, timezone="UTC")

        body = graph.sent("PATCH")
        assert body["start"] == {"dateTime": "2026-10-01T00:00:00", "timeZone": "UTC"}
        assert body["end"]["timeZone"] == "UTC"

    def test_start_and_end_always_share_one_zone(self, graph):
        # A given end with the start read from the event: Graph refuses an
        # all-day event whose two ends are in different zones.
        update(is_all_day=True, end="2026-10-03")

        body = graph.sent("PATCH")
        assert body["start"]["timeZone"] == body["end"]["timeZone"] == "Europe/Rome"
        assert body["end"]["dateTime"] == "2026-10-04T00:00:00"


class TestUpdatingBackToTimed:
    def test_the_current_times_go_with_the_flag(self, graph):
        # Graph refuses isAllDay=false alone too: "Missing parameters: Event.Start".
        update(is_all_day=False)

        body = graph.sent("PATCH")
        assert body["isAllDay"] is False
        assert body["start"] == {"dateTime": "2026-10-02T00:00:00.0000000", "timeZone": "Europe/Rome"}
        assert body["end"] == {"dateTime": "2026-10-02T00:00:00.0000000", "timeZone": "Europe/Rome"}

    def test_given_times_are_sent_as_they_are(self, graph):
        update(is_all_day=False, start="2026-10-02T09:00:00", end="2026-10-02T10:00:00", timezone="Europe/Rome")

        assert graph.reads() == []
        body = graph.sent("PATCH")
        assert body["start"] == {"dateTime": "2026-10-02T09:00:00", "timeZone": "Europe/Rome"}
        assert body["end"] == {"dateTime": "2026-10-02T10:00:00", "timeZone": "Europe/Rome"}


class TestOtherUpdates:
    def test_show_as_alone_touches_nothing_else(self, graph):
        update(show_as="free")

        assert graph.reads() == []
        assert graph.sent("PATCH") == {"showAs": "free"}

    def test_times_without_the_flag_are_as_before(self, graph):
        update(start="2026-10-02T09:00:00", end="2026-10-02T10:00:00")

        assert graph.reads() == []
        assert graph.sent("PATCH") == {
            "start": {"dateTime": "2026-10-02T09:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-10-02T10:00:00", "timeZone": "UTC"},
        }

    def test_cancelling_reads_nothing_and_patches_nothing(self, graph):
        update(is_cancelled=True, is_all_day=True)

        assert [call[:2] for call in graph.calls] == [("POST", f"/me/events/{EVENT}/cancel")]

    def test_nothing_to_do_is_said(self, graph):
        assert update() == "No updates specified."
        assert graph.calls == []


class TestCreating:
    def test_show_as_is_sent(self, graph):
        create(start="2026-10-01T09:00:00", end="2026-10-01T10:00:00", show_as="free")
        assert graph.sent("POST")["showAs"] == "free"

    def test_no_show_as_leaves_graph_its_default(self, graph):
        create(start="2026-10-01T09:00:00", end="2026-10-01T10:00:00")
        assert "showAs" not in graph.sent("POST")

    def test_a_one_day_all_day_event_from_dates(self, graph):
        result = create(
            start="2026-10-01", end="2026-10-01", is_all_day=True, show_as="free",
            timezone="Europe/Rome",
        )

        body = graph.sent("POST")
        assert body["isAllDay"] is True
        assert body["start"] == {"dateTime": "2026-10-01T00:00:00", "timeZone": "Europe/Rome"}
        assert body["end"] == {"dateTime": "2026-10-02T00:00:00", "timeZone": "Europe/Rome"}
        assert "all day on 2026-10-01" in result

    def test_a_timed_event_is_sent_as_given(self, graph):
        create(start="2026-10-01T09:00:00", end="2026-10-01T10:30:00")

        body = graph.sent("POST")
        assert body["start"]["dateTime"] == "2026-10-01T09:00:00"
        assert body["end"]["dateTime"] == "2026-10-01T10:30:00"


class TestShowAsValues:
    @pytest.mark.parametrize("given, canonical", [
        ("free", "free"),
        ("FREE", "free"),
        (" Tentative ", "tentative"),
        ("oof", "oof"),
        ("workingelsewhere", "workingElsewhere"),
        ("WorkingElsewhere", "workingElsewhere"),
    ])
    def test_any_casing_is_graph_s_spelling(self, given, canonical):
        assert UpdateEventInput(event_id=EVENT, show_as=given).show_as == canonical

    def test_an_unknown_status_is_refused_before_graph_sees_it(self):
        with pytest.raises(ValidationError):
            UpdateEventInput(event_id=EVENT, show_as="available")

    def test_create_takes_the_same_values(self):
        event = CreateEventInput(subject="x", start="2026-10-01", end="2026-10-01", show_as="OOF")
        assert event.show_as == "oof"


class TestWhatTheClientSees:
    """The bug as reported: the fields were missing from the tool schema."""

    @staticmethod
    def schema_of(tool_name):
        tools = {tool.name: tool for tool in anyio.run(mcp.list_tools)}
        schema = tools[tool_name].input_schema
        params = schema["properties"]["params"]
        if "$ref" in params:
            params = schema["$defs"][params["$ref"].rsplit("/", 1)[-1]]
        return params["properties"]

    @staticmethod
    def enum_of(field):
        return next(option["enum"] for option in field["anyOf"] if "enum" in option)

    def test_update_takes_is_all_day_and_show_as(self):
        properties = self.schema_of("outlook_update_event")
        assert "is_all_day" in properties
        assert self.enum_of(properties["show_as"]) == [
            "free", "tentative", "busy", "oof", "workingElsewhere",
        ]

    def test_create_takes_show_as(self):
        properties = self.schema_of("outlook_create_event")
        assert "is_all_day" in properties
        assert "free" in self.enum_of(properties["show_as"])
