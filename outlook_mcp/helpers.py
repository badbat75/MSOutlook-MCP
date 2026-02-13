"""Formatting helpers and utility functions."""

from datetime import datetime
from typing import List

import httpx


def make_recipients(addresses: List[str]) -> list:
    """Convert a list of email addresses to Graph API recipient format."""
    return [{"emailAddress": {"address": addr}} for addr in addresses]


def format_email_summary(msg: dict) -> str:
    """Format an email message for display."""
    sender = msg.get("from", {}).get("emailAddress", {})
    sender_str = f"{sender.get('name', 'Unknown')} <{sender.get('address', '')}>"
    received = msg.get("receivedDateTime", "")
    if received:
        dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
        received = dt.strftime("%Y-%m-%d %H:%M UTC")

    importance = msg.get("importance", "normal")
    is_read = "✓ Read" if msg.get("isRead") else "● Unread"
    has_attachments = " 📎" if msg.get("hasAttachments") else ""

    return (
        f"**{msg.get('subject', '(no subject)')}**{has_attachments}\n"
        f"From: {sender_str}\n"
        f"Date: {received} | {is_read} | Importance: {importance}\n"
        f"ID: `{msg.get('id', '')}`"
    )


def format_event_summary(event: dict) -> str:
    """Format a calendar event for display."""
    start = event.get("start", {})
    end = event.get("end", {})
    start_str = format_graph_datetime(start)
    end_str = format_graph_datetime(end)

    location = event.get("location", {}).get("displayName", "No location")
    organizer = event.get("organizer", {}).get("emailAddress", {})
    organizer_str = f"{organizer.get('name', '')} <{organizer.get('address', '')}>"
    is_online = " 🎥" if event.get("isOnlineMeeting") else ""
    status = event.get("showAs", "busy")

    attendees = event.get("attendees", [])
    attendee_list = ", ".join(
        f"{a['emailAddress']['name']} ({a.get('status', {}).get('response', 'none')})"
        for a in attendees[:5]
    )
    if len(attendees) > 5:
        attendee_list += f" +{len(attendees) - 5} more"

    result = (
        f"**{event.get('subject', '(no subject)')}**{is_online}\n"
        f"When: {start_str} → {end_str} | Status: {status}\n"
        f"Location: {location}\n"
        f"Organizer: {organizer_str}\n"
    )
    if attendees:
        result += f"Attendees: {attendee_list}\n"
    result += f"ID: `{event.get('id', '')}`"
    return result


def format_graph_datetime(dt_obj: dict) -> str:
    """Format Graph API datetime object."""
    dt_str = dt_obj.get("dateTime", "")
    tz = dt_obj.get("timeZone", "UTC")
    if dt_str:
        try:
            dt = datetime.fromisoformat(dt_str)
            return f"{dt.strftime('%Y-%m-%d %H:%M')} ({tz})"
        except ValueError:
            return f"{dt_str} ({tz})"
    return "Unknown"


def handle_graph_error(e: Exception) -> str:
    """Format Graph API errors into actionable messages."""
    if isinstance(e, httpx.HTTPStatusError):
        status = e.response.status_code
        try:
            error_body = e.response.json()
            error_msg = error_body.get("error", {}).get("message", str(e))
        except Exception:
            error_msg = str(e)

        if status == 401:
            return (
                f"Error 401: Authentication failed. Token may be expired. "
                f"Re-run auth setup: python outlook_mcp_auth.py\n"
                f"Detail: {error_msg}"
            )
        elif status == 403:
            return f"Error 403: Insufficient permissions. Check app registration scopes.\nDetail: {error_msg}"
        elif status == 404:
            return f"Error 404: Resource not found. Verify the ID is correct.\nDetail: {error_msg}"
        elif status == 429:
            retry_after = e.response.headers.get("Retry-After", "60")
            return f"Error 429: Rate limited. Retry after {retry_after} seconds."
        else:
            return f"Error {status}: {error_msg}"
    elif isinstance(e, httpx.TimeoutException):
        return "Error: Request timed out. The Graph API may be slow. Please retry."
    return f"Error: {type(e).__name__}: {str(e)}"


def get_day_of_week(iso_date: str) -> str:
    """Get day of week name from ISO date string."""
    try:
        dt = datetime.fromisoformat(iso_date.replace("Z", "+00:00"))
        days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
        return days[dt.weekday()]
    except Exception:
        return "monday"
