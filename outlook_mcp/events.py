"""All-day calendar events, the way Graph insists on them.

Graph accepts an all-day event only as midnight to midnight, the end exclusive
and at least a day after the start, both in one time zone. And on an update it
refuses isAllDay, whichever way it flips, unless start comes with it ("Missing
parameters: Event.Start"), even when the event already sits on midnight. None
of that is something a caller should have to know to say "make it all day":
this module turns what a caller has, a date or the times of the event being
converted, into what Graph takes.
"""

import re
from datetime import date, timedelta
from typing import Optional, Tuple

from .auth import GraphClient

# What follows the date in a value that ends exactly at midnight: the time,
# Graph's seven fractional digits included, and an offset or Z if a caller
# wrote one. A bare date does not match, on purpose: see all_day_span().
_MIDNIGHT = re.compile(r"[T ]00(?::00(?::00(?:\.0+)?)?)?(?:Z|[+-]\d\d(?::?\d\d)?)?")


def _day(value: str) -> date:
    try:
        return date.fromisoformat(value.strip()[:10])
    except ValueError:
        raise ValueError(
            f"'{value}' is not an ISO date or datetime: write '2026-10-01' or "
            f"'2026-10-01T09:00:00'."
        ) from None


def _midnight(day: date) -> str:
    return f"{day.isoformat()}T00:00:00"


def all_day_span(start: str, end: Optional[str]) -> Tuple[str, str]:
    """The whole days an all-day event from `start` to `end` covers, as Graph wants them.

    Every day the interval touches is in, which is what Outlook does when the
    "All day" box is ticked on a timed event. A date names a whole day, so an
    end given as one is the last day ('2026-10-02' ends after the 2nd). A time
    names an instant, so an end at midnight is exclusive, which is how Graph
    itself returns an all-day event: reading one back and passing its times in
    again changes nothing. No end, or one that leaves nothing, means one day.

    Returns start and end as midnights, the end exclusive.
    """
    first = _day(start)
    stop = first + timedelta(days=1)
    if end:
        last = _day(end)
        if not _MIDNIGHT.fullmatch(end.strip()[10:]):
            last += timedelta(days=1)
        stop = max(stop, last)
    return _midnight(first), _midnight(stop)


def describe_all_day(start: str, end: str) -> str:
    """An all_day_span() result in words, counting the days it covers."""
    first, stop = _day(start), _day(end)
    days = (stop - first).days
    if days <= 1:
        return f"all day on {first}"
    return f"all day, {first} to {stop - timedelta(days=1)} ({days} days)"


async def read_event_times(
    graph: GraphClient, event_id: str, zone: Optional[str] = None
) -> Tuple[str, str, str]:
    """An event's current start and end as wall-clock times in one time zone.

    Without a Prefer header Graph answers in UTC, and midnight in Rome is 22:00
    the day before in UTC: the wrong day to make an event all-day on. So the
    times are read in `zone`, or when that is not given in the event's own,
    the one it was scheduled in. A zone Graph does not know, such as the custom
    one an external invitation can carry, is its 400 to report.

    Returns (start, end, zone), the zone being the one the times are in.
    """
    if not zone:
        meta = await graph.get(
            f"/me/events/{event_id}", params={"$select": "originalStartTimeZone"}
        )
        zone = meta.get("originalStartTimeZone") or "UTC"
    data = await graph.get(
        f"/me/events/{event_id}",
        params={"$select": "start,end"},
        headers={"Prefer": f'outlook.timezone="{zone}"'},
    )
    start, end = data["start"], data["end"]
    return start["dateTime"], end["dateTime"], start.get("timeZone") or zone
