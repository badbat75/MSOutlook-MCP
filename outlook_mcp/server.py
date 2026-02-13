"""
Outlook MCP Server - Tool definitions and server lifecycle.

Provides 16 MCP tools for Outlook email and calendar operations
via Microsoft Graph API.
"""

import json
import os
import sys
from datetime import datetime, timedelta, timezone
from typing import Dict, Any
from contextlib import asynccontextmanager

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.server import Context

from .auth import AuthManager, GraphClient
from .models import (
    ListMailInput, GetMailInput, SendMailInput, CreateDraftInput,
    ReplyMailInput, MoveMailInput, UpdateMailInput, ListMailFoldersInput,
    ListEventsInput, GetEventInput, CreateEventInput, UpdateEventInput,
    DeleteEventInput, RespondEventInput, ListCalendarsInput,
)
from .helpers import (
    make_recipients, format_email_summary, format_event_summary,
    format_graph_datetime, handle_graph_error, get_day_of_week,
)


# =============================================================================
# MCP Server Setup
# =============================================================================

@asynccontextmanager
async def app_lifespan(app):
    """Initialize Graph client on startup, clean up on shutdown."""
    client_id = os.environ.get("OUTLOOK_CLIENT_ID", "")
    client_secret = os.environ.get("OUTLOOK_CLIENT_SECRET", "")
    tenant_id = os.environ.get("OUTLOOK_TENANT_ID", "common")

    if not client_id or not client_secret:
        import logging
        logging.getLogger("outlook_mcp").warning(
            "OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET must be set. "
            "The server will start but all tools will fail until configured."
        )

    auth = AuthManager(client_id, client_secret, tenant_id)
    graph = GraphClient(auth)

    yield {"graph": graph}

    await graph.close()


mcp = FastMCP("MS_Outlook_MCP", lifespan=app_lifespan)


def _get_graph(ctx: Context) -> GraphClient:
    """Extract GraphClient from context."""
    return ctx.request_context.lifespan_context["graph"]


# =============================================================================
# EMAIL TOOLS
# =============================================================================

@mcp.tool(
    name="outlook_list_mail",
    annotations={
        "title": "List Outlook Emails",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_list_mail(params: ListMailInput, ctx: Context = None) -> str:
    """List emails from an Outlook mailbox folder with filtering and search.

    Retrieves messages from the specified folder (default: inbox) with support for
    OData filters, full-text search, field selection, and pagination.

    Returns:
        str: Formatted list of email summaries with subject, sender, date, and IDs.
    """
    try:
        graph = _get_graph(ctx)
        folder_map = {
            "inbox": "inbox",
            "sentitems": "sentItems",
            "sent": "sentItems",
            "drafts": "drafts",
            "deleteditems": "deletedItems",
            "trash": "deletedItems",
            "junkemail": "junkEmail",
            "junk": "junkEmail",
            "archive": "archive",
        }
        folder = folder_map.get(params.folder.lower(), params.folder)
        endpoint = f"/me/mailFolders/{folder}/messages"

        query_params = {
            "$top": params.top,
            "$skip": params.skip,
            "$orderby": "receivedDateTime desc",
            "$select": params.select or "id,subject,from,receivedDateTime,isRead,importance,hasAttachments,bodyPreview",
        }
        if params.filter:
            query_params["$filter"] = params.filter
        if params.search:
            query_params["$search"] = f'"{params.search}"'

        data = await graph.get(endpoint, params=query_params)
        messages = data.get("value", [])

        if not messages:
            return f"No messages found in '{params.folder}'"

        total = data.get("@odata.count", "unknown")
        result = f"📬 **{params.folder.title()}** — {len(messages)} messages (skip: {params.skip})\n\n"
        for msg in messages:
            result += format_email_summary(msg) + "\n\n---\n\n"

        if data.get("@odata.nextLink"):
            result += f"\n*More messages available. Use skip={params.skip + params.top} for next page.*"

        return result
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_get_mail",
    annotations={
        "title": "Get Email Details",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_get_mail(params: GetMailInput, ctx: Context = None) -> str:
    """Get the full details of a specific email by its ID.

    Returns complete message content including body, headers, attachments info,
    and all metadata.

    Returns:
        str: Full email details in formatted text.
    """
    try:
        graph = _get_graph(ctx)
        select = "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,importance,isRead,hasAttachments,categories,flag,internetMessageHeaders"
        if params.include_body:
            select += ",body,bodyPreview"
        data = await graph.get(f"/me/messages/{params.message_id}", params={"$select": select})

        sender = data.get("from", {}).get("emailAddress", {})
        to_list = ", ".join(
            f"{r['emailAddress']['name']} <{r['emailAddress']['address']}>"
            for r in data.get("toRecipients", [])
        )
        cc_list = ", ".join(
            f"{r['emailAddress']['name']} <{r['emailAddress']['address']}>"
            for r in data.get("ccRecipients", [])
        )

        result = f"# {data.get('subject', '(no subject)')}\n\n"
        result += f"**From:** {sender.get('name', '')} <{sender.get('address', '')}>\n"
        result += f"**To:** {to_list}\n"
        if cc_list:
            result += f"**CC:** {cc_list}\n"
        result += f"**Date:** {data.get('receivedDateTime', '')}\n"
        result += f"**Importance:** {data.get('importance', 'normal')}\n"
        result += f"**Read:** {'Yes' if data.get('isRead') else 'No'}\n"
        result += f"**Has Attachments:** {'Yes' if data.get('hasAttachments') else 'No'}\n"

        categories = data.get("categories", [])
        if categories:
            result += f"**Categories:** {', '.join(categories)}\n"

        flag = data.get("flag", {}).get("flagStatus", "notFlagged")
        result += f"**Flag:** {flag}\n"

        if params.include_body:
            body = data.get("body", {})
            content_type = body.get("contentType", "text")
            content = body.get("content", "")
            result += f"\n---\n\n**Body** ({content_type}):\n\n{content}"

        return result
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_send_mail",
    annotations={
        "title": "Send Email",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def outlook_send_mail(params: SendMailInput, ctx: Context = None) -> str:
    """Send an email through Outlook.

    Composes and sends an email with support for HTML/text body, CC/BCC,
    importance level, and optional save to Sent Items.

    Returns:
        str: Confirmation message with details.
    """
    try:
        graph = _get_graph(ctx)

        payload = {
            "message": {
                "subject": params.subject,
                "body": {
                    "contentType": "HTML" if params.is_html else "Text",
                    "content": params.body,
                },
                "toRecipients": make_recipients(params.to),
                "importance": params.importance,
            },
            "saveToSentItems": params.save_to_sent,
        }

        if params.cc:
            payload["message"]["ccRecipients"] = make_recipients(params.cc)
        if params.bcc:
            payload["message"]["bccRecipients"] = make_recipients(params.bcc)

        await graph.post("/me/sendMail", json_data=payload)

        recipients = ", ".join(params.to)
        return f"✅ Email sent successfully!\n**To:** {recipients}\n**Subject:** {params.subject}"
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_create_draft",
    annotations={
        "title": "Create Draft Email",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def outlook_create_draft(params: CreateDraftInput, ctx: Context = None) -> str:
    """Create a draft email in the Drafts folder without sending it.

    Creates a message via POST /me/messages. The draft can later be edited
    in Outlook or sent via the Graph API.

    Returns:
        str: Confirmation with draft ID for later reference.
    """
    try:
        graph = _get_graph(ctx)

        payload = {
            "subject": params.subject,
            "body": {
                "contentType": "HTML" if params.is_html else "Text",
                "content": params.body,
            },
            "toRecipients": make_recipients(params.to),
            "importance": params.importance,
        }

        if params.cc:
            payload["ccRecipients"] = make_recipients(params.cc)
        if params.bcc:
            payload["bccRecipients"] = make_recipients(params.bcc)

        result = await graph.post("/me/messages", json_data=payload)

        draft_id = result.get("id", "unknown")
        recipients = ", ".join(params.to)
        return (
            f"📝 Draft created successfully!\n"
            f"**To:** {recipients}\n"
            f"**Subject:** {params.subject}\n"
            f"**Draft ID:** `{draft_id}`"
        )
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_reply_mail",
    annotations={
        "title": "Reply to Email",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def outlook_reply_mail(params: ReplyMailInput, ctx: Context = None) -> str:
    """Reply to an email or reply-all.

    Args:
        params: Message ID, reply text, and whether to reply to all.

    Returns:
        str: Confirmation of reply sent.
    """
    try:
        graph = _get_graph(ctx)
        endpoint_suffix = "replyAll" if params.reply_all else "reply"
        await graph.post(
            f"/me/messages/{params.message_id}/{endpoint_suffix}",
            json_data={"comment": params.comment},
        )
        mode = "Reply All" if params.reply_all else "Reply"
        return f"✅ {mode} sent successfully for message `{params.message_id}`"
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_move_mail",
    annotations={
        "title": "Move Email to Folder",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_move_mail(params: MoveMailInput, ctx: Context = None) -> str:
    """Move an email to a different folder (archive, trash, etc.).

    Returns:
        str: Confirmation of the move.
    """
    try:
        graph = _get_graph(ctx)
        folder_map = {
            "inbox": "inbox",
            "archive": "archive",
            "deleteditems": "deleteditems",
            "trash": "deleteditems",
            "junkemail": "junkemail",
            "junk": "junkemail",
            "drafts": "drafts",
            "sentitems": "sentitems",
        }
        dest = folder_map.get(params.destination_folder.lower(), params.destination_folder)
        data = await graph.post(
            f"/me/messages/{params.message_id}/move",
            json_data={"destinationId": dest},
        )
        return f"✅ Message moved to '{params.destination_folder}'. New ID: `{data.get('id', 'N/A')}`"
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_update_mail",
    annotations={
        "title": "Update Email Properties",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_update_mail(params: UpdateMailInput, ctx: Context = None) -> str:
    """Update email properties: read status, categories, or flag.

    Returns:
        str: Confirmation of updates applied.
    """
    try:
        graph = _get_graph(ctx)
        updates: Dict[str, Any] = {}
        if params.is_read is not None:
            updates["isRead"] = params.is_read
        if params.categories is not None:
            updates["categories"] = params.categories
        if params.flag_status is not None:
            updates["flag"] = {"flagStatus": params.flag_status}

        if not updates:
            return "No updates specified. Provide at least one property to update."

        await graph.patch(f"/me/messages/{params.message_id}", json_data=updates)

        changes = ", ".join(f"{k}={v}" for k, v in updates.items())
        return f"✅ Message updated: {changes}"
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_list_folders",
    annotations={
        "title": "List Mail Folders",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_list_folders(params: ListMailFoldersInput, ctx: Context = None) -> str:
    """List all mail folders in the mailbox.

    Returns:
        str: List of folders with names, IDs, and message counts.
    """
    try:
        graph = _get_graph(ctx)
        data = await graph.get(
            "/me/mailFolders",
            params={"$top": params.top, "$select": "id,displayName,totalItemCount,unreadItemCount"},
        )
        folders = data.get("value", [])
        if not folders:
            return "No mail folders found."

        result = "📁 **Mail Folders**\n\n"
        for f in folders:
            unread = f.get("unreadItemCount", 0)
            unread_badge = f" (📬 {unread} unread)" if unread > 0 else ""
            result += (
                f"- **{f['displayName']}**{unread_badge} — "
                f"{f.get('totalItemCount', 0)} items | ID: `{f['id']}`\n"
            )
        return result
    except Exception as e:
        return handle_graph_error(e)


# =============================================================================
# CALENDAR TOOLS
# =============================================================================

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
        graph = _get_graph(ctx)
        now = datetime.now(timezone.utc)
        start = params.start_date or now.strftime("%Y-%m-%dT00:00:00")
        end = params.end_date or (now + timedelta(days=7)).strftime("%Y-%m-%dT23:59:59")

        # Ensure proper ISO format
        if "T" not in start:
            start += "T00:00:00"
        if "T" not in end:
            end += "T23:59:59"

        base = f"/me/calendars/{params.calendar_id}" if params.calendar_id else "/me"
        endpoint = f"{base}/calendarView"

        data = await graph.get(
            endpoint,
            params={
                "startDateTime": start,
                "endDateTime": end,
                "$top": params.top,
                "$orderby": "start/dateTime",
                "$select": "id,subject,start,end,location,organizer,attendees,isOnlineMeeting,showAs,isCancelled,recurrence",
            },
        )
        events = data.get("value", [])

        if not events:
            return f"No events found between {start[:10]} and {end[:10]}"

        result = f"📅 **Calendar Events** ({start[:10]} → {end[:10]})\n\n"
        for event in events:
            if event.get("isCancelled"):
                continue
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
        graph = _get_graph(ctx)
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
                result += f"- {email['name']} <{email['address']}> — {status}\n"

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
        graph = _get_graph(ctx)
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
        graph = _get_graph(ctx)
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
        graph = _get_graph(ctx)
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
        graph = _get_graph(ctx)
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
        graph = _get_graph(ctx)
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


# =============================================================================
# USER INFO TOOL
# =============================================================================

@mcp.tool(
    name="outlook_get_profile",
    annotations={
        "title": "Get User Profile",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_get_profile(ctx: Context = None) -> str:
    """Get the authenticated user's profile information.

    Returns:
        str: User profile with name, email, job title, etc.
    """
    try:
        graph = _get_graph(ctx)
        data = await graph.get(
            "/me",
            params={"$select": "displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone"},
        )
        result = "👤 **User Profile**\n\n"
        result += f"**Name:** {data.get('displayName', 'N/A')}\n"
        result += f"**Email:** {data.get('mail') or data.get('userPrincipalName', 'N/A')}\n"
        result += f"**Job Title:** {data.get('jobTitle', 'N/A')}\n"
        result += f"**Department:** {data.get('department', 'N/A')}\n"
        result += f"**Office:** {data.get('officeLocation', 'N/A')}\n"
        result += f"**Phone:** {data.get('mobilePhone', 'N/A')}\n"
        return result
    except Exception as e:
        return handle_graph_error(e)


# =============================================================================
# Entry Point
# =============================================================================

def main():
    """Run the MCP server (stdio or HTTP transport)."""
    if "--http" in sys.argv:
        port = 8000
        for i, arg in enumerate(sys.argv):
            if arg == "--port" and i + 1 < len(sys.argv):
                port = int(sys.argv[i + 1])
        print(f"Starting Outlook MCP server on http://localhost:{port}")
        mcp.run(transport="streamable_http", port=port)
    else:
        mcp.run()  # stdio transport (default for Claude Desktop)


if __name__ == "__main__":
    main()
