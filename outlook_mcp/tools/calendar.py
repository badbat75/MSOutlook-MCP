"""Calendar tools: events and calendars."""

from datetime import datetime, timedelta, timezone
from typing import Any, Dict

from mcp.server.mcpserver import Context

from ..app import mcp
from ..credentials import get_graph
from ..helpers import (
    format_event_summary,
    format_graph_datetime,
    get_day_of_week,
    handle_graph_error,
)
from ..models import (
    CreateEventInput,
    DeleteEventInput,
    GetEventInput,
    ListCalendarsInput,
    ListEventsInput,
    RespondEventInput,
    UpdateEventInput,
)


@mcp.tool(
    name="outlook_list_events",
    annotations={
        "title": "List Calendar Events",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_list_events(params: ListEventsInput, ctx: Context = None) -> str:
    """List calendar events within a date range.

    Uses the calendarView endpoint for accurate recurring event expansion.
    Defaults to the next 7 days if no dates specified.

    Returns:
        str: Formatted list of calendar events with details.
    """
    try:
        graph = get_graph(ctx)
        now = datetime.now(timezone.utc)
        start = params.start_date or now.strftime("%Y-%m-%dT00:00:00")
        end = params.end_date or (now + timedelta(days=7)).strftime("%Y-%m-%dT23:59:59")

        # Ensure proper ISO format
        if "T" not in start:
            start += "T00:00:00"
        if "T" not in end:
            end += "T23:59:59"

        query = {
            "startDateTime": start,
            "endDateTime": end,
            "$top": params.top,
            "$orderby": "start/dateTime",
            "$select": "id,subject,start,end,location,organizer,attendees,isOnlineMeeting,showAs,isCancelled,recurrence",
        }

        if params.calendar_id:
            # Single calendar: query directly.
            data = await graph.get(
                f"/me/calendars/{params.calendar_id}/calendarView", params=query
            )
            events = [e for e in data.get("value", []) if not e.get("isCancelled")]
        else:
            # No calendar specified: aggregate across ALL calendars (e.g. the
            # default "Calendar", "Birthdays", "Your Family", etc.). The plain
            # /me/calendarView endpoint only returns the default calendar.
            cal_data = await graph.get(
                "/me/calendars", params={"$select": "id,name", "$top": 50}
            )
            calendars = cal_data.get("value", [])

            events = []
            for cal in calendars:
                cal_view = await graph.get(
                    f"/me/calendars/{cal['id']}/calendarView", params=query
                )
                for event in cal_view.get("value", []):
                    if event.get("isCancelled"):
                        continue
                    event["calendarName"] = cal.get("name", "")
                    events.append(event)

            # Merge and sort all calendars by start time, then cap at top.
            events.sort(key=lambda e: e.get("start", {}).get("dateTime", ""))
            events = events[: params.top]

        if not events:
            return f"No events found between {start[:10]} and {end[:10]}"

        result = f"📅 **Calendar Events** ({start[:10]} → {end[:10]})\n\n"
        for event in events:
            result += format_event_summary(event) + "\n\n---\n\n"

        return result
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_get_event",
    annotations={
        "title": "Get Event Details",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_get_event(params: GetEventInput, ctx: Context = None) -> str:
    """Get full details of a specific calendar event.

    Returns:
        str: Complete event details including body, attendees, and online meeting info.
    """
    try:
        graph = get_graph(ctx)
        data = await graph.get(f"/me/events/{params.event_id}")

        result = f"# {data.get('subject', '(no subject)')}\n\n"
        result += f"**Start:** {format_graph_datetime(data.get('start', {}))}\n"
        result += f"**End:** {format_graph_datetime(data.get('end', {}))}\n"
        result += f"**Location:** {data.get('location', {}).get('displayName', 'None')}\n"
        result += f"**Status:** {data.get('showAs', 'busy')}\n"
        result += f"**All Day:** {'Yes' if data.get('isAllDay') else 'No'}\n"

        organizer = data.get("organizer", {}).get("emailAddress", {})
        result += f"**Organizer:** {organizer.get('name', '')} <{organizer.get('address', '')}>\n"

        if data.get("isOnlineMeeting"):
            join_url = data.get("onlineMeeting", {}).get("joinUrl", "N/A")
            result += f"**Teams Meeting:** [Join]({join_url})\n"

        attendees = data.get("attendees", [])
        if attendees:
            result += "\n**Attendees:**\n"
            for a in attendees:
                email = a["emailAddress"]
                status = a.get("status", {}).get("response", "none")
                result += f"- {email['name']} <{email['address']}>: {status}\n"

        body = data.get("body", {})
        if body.get("content"):
            result += f"\n---\n\n**Description:**\n\n{body['content']}"

        return result
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_create_event",
    annotations={
        "title": "Create Calendar Event",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def outlook_create_event(params: CreateEventInput, ctx: Context = None) -> str:
    """Create a new calendar event with optional attendees and Teams meeting.

    Supports setting location, body, reminders, recurrence, and online meeting creation.

    Returns:
        str: Confirmation with the new event ID and details.
    """
    try:
        graph = get_graph(ctx)
        event_body: Dict[str, Any] = {
            "subject": params.subject,
            "start": {"dateTime": params.start, "timeZone": params.timezone},
            "end": {"dateTime": params.end, "timeZone": params.timezone},
            "isOnlineMeeting": params.is_online_meeting,
            "isAllDay": params.is_all_day,
            "reminderMinutesBeforeStart": params.reminder_minutes,
        }

        if params.body:
            event_body["body"] = {"contentType": "HTML", "content": params.body}
        if params.location:
            event_body["location"] = {"displayName": params.location}
        if params.attendees:
            event_body["attendees"] = [
                {
                    "emailAddress": {"address": email},
                    "type": "required",
                }
                for email in params.attendees
            ]
        if params.is_online_meeting:
            event_body["onlineMeetingProvider"] = "teamsForBusiness"

        if params.recurrence:
            pattern_map = {
                "daily": {"type": "daily", "interval": 1},
                "weekly": {"type": "weekly", "interval": 1, "daysOfWeek": [get_day_of_week(params.start)]},
                "monthly": {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": int(params.start[8:10])},
            }
            if params.recurrence in pattern_map:
                event_body["recurrence"] = {
                    "pattern": pattern_map[params.recurrence],
                    "range": {
                        "type": "noEnd",
                        "startDate": params.start[:10],
                    },
                }

        base = f"/me/calendars/{params.calendar_id}/events" if params.calendar_id else "/me/events"
        data = await graph.post(base, json_data=event_body)

        result = f"✅ Event created!\n"
        result += f"**Subject:** {params.subject}\n"
        result += f"**When:** {params.start} → {params.end} ({params.timezone})\n"
        if params.location:
            result += f"**Location:** {params.location}\n"
        if params.is_online_meeting:
            join_url = data.get("onlineMeeting", {}).get("joinUrl", "")
            result += f"**Teams Meeting:** {join_url}\n"
        result += f"**Event ID:** `{data.get('id', 'N/A')}`"
        return result
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_update_event",
    annotations={
        "title": "Update Calendar Event",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_update_event(params: UpdateEventInput, ctx: Context = None) -> str:
    """Update properties of an existing calendar event.

    Returns:
        str: Confirmation of applied changes.
    """
    try:
        graph = get_graph(ctx)
        updates: Dict[str, Any] = {}

        if params.subject:
            updates["subject"] = params.subject
        if params.start:
            tz = params.timezone or "UTC"
            updates["start"] = {"dateTime": params.start, "timeZone": tz}
        if params.end:
            tz = params.timezone or "UTC"
            updates["end"] = {"dateTime": params.end, "timeZone": tz}
        if params.location:
            updates["location"] = {"displayName": params.location}
        if params.body:
            updates["body"] = {"contentType": "HTML", "content": params.body}
        if params.is_cancelled:
            await graph.post(f"/me/events/{params.event_id}/cancel", json_data={})
            return f"✅ Event `{params.event_id}` has been cancelled."

        if not updates:
            return "No updates specified."

        await graph.patch(f"/me/events/{params.event_id}", json_data=updates)
        changes = ", ".join(updates.keys())
        return f"✅ Event updated ({changes}). ID: `{params.event_id}`"
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_delete_event",
    annotations={
        "title": "Delete Calendar Event",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_delete_event(params: DeleteEventInput, ctx: Context = None) -> str:
    """Permanently delete a calendar event.

    Returns:
        str: Confirmation of deletion.
    """
    try:
        graph = get_graph(ctx)
        await graph.delete(f"/me/events/{params.event_id}")
        return f"✅ Event `{params.event_id}` has been deleted."
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_respond_event",
    annotations={
        "title": "Respond to Event Invitation",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_respond_event(params: RespondEventInput, ctx: Context = None) -> str:
    """Accept, tentatively accept, or decline a calendar event invitation.

    Returns:
        str: Confirmation of response.
    """
    try:
        graph = get_graph(ctx)
        payload: Dict[str, Any] = {"sendResponse": params.send_response}
        if params.comment:
            payload["comment"] = params.comment

        await graph.post(
            f"/me/events/{params.event_id}/{params.response}",
            json_data=payload,
        )
        return f"✅ Event `{params.event_id}`: response '{params.response}' sent."
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_list_calendars",
    annotations={
        "title": "List Calendars",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_list_calendars(params: ListCalendarsInput, ctx: Context = None) -> str:
    """List all calendars in the user's account.

    Returns:
        str: List of calendars with names, IDs, and properties.
    """
    try:
        graph = get_graph(ctx)
        data = await graph.get(
            "/me/calendars",
            params={"$top": params.top, "$select": "id,name,color,isDefaultCalendar,canEdit,owner"},
        )
        calendars = data.get("value", [])
        if not calendars:
            return "No calendars found."

        result = "📅 **Your Calendars**\n\n"
        for cal in calendars:
            default_badge = " ⭐" if cal.get("isDefaultCalendar") else ""
            owner = cal.get("owner", {})
            result += (
                f"- **{cal['name']}**{default_badge}\n"
                f"  Color: {cal.get('color', 'auto')} | "
                f"Can Edit: {'Yes' if cal.get('canEdit') else 'No'} | "
                f"Owner: {owner.get('name', 'N/A')}\n"
                f"  ID: `{cal['id']}`\n"
            )
        return result
    except Exception as e:
        return handle_graph_error(e)
