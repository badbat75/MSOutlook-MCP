"""Formatting helpers and utility functions."""

import base64
import os
from datetime import datetime
from pathlib import Path
from typing import List

import httpx

from .auth import CredentialsError


# Reusable guidance for building a valid OData $filter (shared by the pre-flight
# validation in server.py and the 400-error handler below).
FILTER_SYNTAX_HINT = (
    "Build $filter with comparison operators (eq, ne, gt, ge, lt, le) joined by "
    "and/or, using ISO-8601 datetimes with a 'Z' suffix. Examples: "
    "\"isRead eq false\", "
    "\"receivedDateTime ge 2026-05-26T00:00:00Z and receivedDateTime lt 2026-05-27T00:00:00Z\", "
    "\"from/emailAddress/address eq 'john@example.com'\"."
)


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
    calendar_name = event.get("calendarName")
    if calendar_name:
        result += f"Calendar: {calendar_name}\n"
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
        elif status == 400 and "filter" in error_msg.lower():
            return f"Error 400: {error_msg}\n\n{FILTER_SYNTAX_HINT}"
        else:
            return f"Error {status}: {error_msg}"
    elif isinstance(e, httpx.TimeoutException):
        return "Error: Request timed out. The Graph API may be slow. Please retry."
    elif isinstance(e, CredentialsError):
        # The message already says which headers / variables are missing.
        return f"Error: {e}"
    return f"Error: {type(e).__name__}: {str(e)}"


def get_day_of_week(iso_date: str) -> str:
    """Get day of week name from ISO date string."""
    try:
        dt = datetime.fromisoformat(iso_date.replace("Z", "+00:00"))
        days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
        return days[dt.weekday()]
    except Exception:
        return "monday"


# Attachment handling constants
DEFAULT_ATTACHMENT_DIR = Path.home() / "Downloads" / "outlook_attachments"


def attachment_download_dir() -> Path:
    """Directory attachments are written to: OUTLOOK_DOWNLOAD_PATH, or a default.

    Resolved on every call rather than at import time, because the entry point
    merges the project .env into the environment and may do so after this
    module has been imported. Server-side on purpose: in HTTP mode a remote
    caller must never choose where the server writes files.
    """
    configured = os.environ.get("OUTLOOK_DOWNLOAD_PATH", "").strip()
    return Path(configured).expanduser() if configured else DEFAULT_ATTACHMENT_DIR


MAX_INLINE_SIZE_MB = 10
VIEWABLE_IMAGE_TYPES = {
    "image/png", "image/jpeg", "image/jpg", "image/gif",
    "image/bmp", "image/webp", "image/svg+xml"
}
ANALYZABLE_TYPES = {
    "application/pdf", "text/plain", "text/html", "text/csv",
    "application/json", "application/xml", "text/xml"
}


def format_attachment_summary(attachment: dict) -> str:
    """Format an attachment metadata for display.

    Args:
        attachment: Graph API attachment object

    Returns:
        str: Formatted attachment summary with type, name, size
    """
    att_type = attachment.get("@odata.type", "")
    name = attachment.get("name", "(unnamed)")
    size_bytes = attachment.get("size", 0)
    content_type = attachment.get("contentType", "unknown")
    att_id = attachment.get("id", "")

    # Format size
    if size_bytes < 1024:
        size_str = f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        size_str = f"{size_bytes / 1024:.1f} KB"
    else:
        size_str = f"{size_bytes / (1024 * 1024):.1f} MB"

    # Icon based on type
    icon = "📎"
    if "image" in content_type:
        icon = "🖼️"
    elif "pdf" in content_type:
        icon = "📄"
    elif "zip" in content_type or "compressed" in content_type:
        icon = "🗜️"
    elif "word" in content_type or "document" in content_type:
        icon = "📝"
    elif "excel" in content_type or "spreadsheet" in content_type:
        icon = "📊"

    # Type indicator
    type_label = ""
    if att_type == "#microsoft.graph.fileAttachment":
        type_label = "File"
    elif att_type == "#microsoft.graph.itemAttachment":
        type_label = "Item (email/meeting)"
    elif att_type == "#microsoft.graph.referenceAttachment":
        type_label = "Reference (cloud)"

    return (
        f"{icon} **{name}**\n"
        f"   Type: {type_label} ({content_type})\n"
        f"   Size: {size_str}\n"
        f"   ID: `{att_id}`"
    )


def should_save_to_disk(content_type: str, size_bytes: int, force_disk: bool) -> bool:
    """Determine if an attachment should be saved to disk vs returned as base64.

    Args:
        content_type: MIME type of the attachment
        size_bytes: Size in bytes
        force_disk: User-requested force save to disk

    Returns:
        bool: Always True - base64 streaming is too heavy for MCP, save everything to disk
    """
    # Always save to disk - base64 data URLs are too heavy for MCP protocol
    return True


def create_data_url(content_type: str, base64_content: str) -> str:
    """Create a data URL from base64 content.

    Args:
        content_type: MIME type
        base64_content: Base64-encoded content

    Returns:
        str: data URL that Claude can render
    """
    return f"data:{content_type};base64,{base64_content}"


def save_attachment_to_disk(filename: str, content_bytes: bytes) -> str:
    """Save attachment bytes to disk.

    Args:
        filename: Original filename
        content_bytes: Raw file bytes

    Returns:
        str: Absolute path to saved file

    Raises:
        OSError: If save fails
    """
    # Create directory if needed
    download_dir = attachment_download_dir()
    download_dir.mkdir(parents=True, exist_ok=True)

    # Sanitize filename (remove path traversal attempts)
    safe_filename = Path(filename).name
    if not safe_filename:
        safe_filename = "attachment"

    # Handle duplicates
    target_path = download_dir / safe_filename
    counter = 1
    while target_path.exists():
        stem = Path(safe_filename).stem
        suffix = Path(safe_filename).suffix
        target_path = download_dir / f"{stem}_{counter}{suffix}"
        counter += 1

    # Write file
    target_path.write_bytes(content_bytes)
    return str(target_path.absolute())
