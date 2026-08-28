"""Email tools: listing, reading, sending, filing and attachments."""

import base64
import json
from typing import Any, Dict

from mcp.server.mcpserver import Context

from ..app import mcp
from ..attachments import attach_files
from ..credentials import get_graph
from ..folders import format_folder_tree, resolve_folder
from ..helpers import (
    create_data_url,
    format_attachment_summary,
    format_email_summary,
    handle_graph_error,
    make_recipients,
    save_attachment_to_disk,
    should_save_to_disk,
    validate_odata_filter,
)
from ..models import (
    CreateDraftInput,
    DeleteMailInput,
    GetAttachmentInput,
    GetMailInput,
    ListAttachmentsInput,
    ListMailFoldersInput,
    ListMailInput,
    MoveMailInput,
    ReplyMailInput,
    SendMailInput,
    UpdateMailInput,
)


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
        graph = get_graph(ctx)
        # '*'/'all' searches the whole mailbox via /me/messages (every folder),
        # so a message can be found without knowing which folder it lives in.
        all_folders = params.folder.strip().lower() in ("*", "all")
        if all_folders:
            endpoint = "/me/messages"
            folder_label = "All folders"
        else:
            folder = await resolve_folder(graph, params.folder)
            endpoint = f"/me/mailFolders/{folder}/messages"
            folder_label = params.folder.title()

        query_params = {
            "$top": params.top,
            "$select": params.select or "id,subject,from,receivedDateTime,isRead,importance,hasAttachments,bodyPreview",
        }
        if params.filter:
            validate_odata_filter(params.filter)
            query_params["$filter"] = params.filter
        if params.search:
            # Graph forbids combining $search with $orderby or $skip: search
            # sorts by relevance and does not support offset pagination, so
            # adding either returns 400. Only set them when not searching.
            query_params["$search"] = f'"{params.search}"'
        else:
            query_params["$skip"] = params.skip
            # Graph rejects $orderby combined with an arbitrary $filter as
            # "The restriction or sort order is too complex for this operation".
            # Only sort by date when unfiltered; with a filter, use default order.
            if not params.filter:
                query_params["$orderby"] = "receivedDateTime desc"

        data = await graph.get(endpoint, params=query_params)
        messages = data.get("value", [])

        if not messages:
            return f"No messages found in '{params.folder}'"

        total = data.get("@odata.count", "unknown")
        header = f"📬 **{folder_label}**: {len(messages)} messages"
        if not params.search:
            header += f" (skip: {params.skip})"
        result = header + "\n\n"
        for msg in messages:
            result += format_email_summary(msg) + "\n\n---\n\n"

        # $skip-based pagination only applies when not searching (search has no offset).
        if data.get("@odata.nextLink") and not params.search:
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
        graph = get_graph(ctx)
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
        graph = get_graph(ctx)

        message = {
            "subject": params.subject,
            "body": {
                "contentType": "HTML" if params.is_html else "Text",
                "content": params.body,
            },
            "toRecipients": make_recipients(params.to),
            "importance": params.importance,
        }
        if params.cc:
            message["ccRecipients"] = make_recipients(params.cc)
        if params.bcc:
            message["bccRecipients"] = make_recipients(params.bcc)

        recipients = ", ".join(params.to)

        if params.attachments:
            # sendMail can't carry large attachments in a single request, so build
            # a draft, attach files (inline or via upload session), then send it.
            draft = await graph.post("/me/messages", json_data=message)
            message_id = draft["id"]
            attached = await attach_files(graph, message_id, params.attachments)
            await graph.post(f"/me/messages/{message_id}/send")
            return (
                f"✅ Email sent successfully!\n**To:** {recipients}\n"
                f"**Subject:** {params.subject}\n"
                f"**Attachments:** {', '.join(attached)}"
            )

        await graph.post(
            "/me/sendMail",
            json_data={"message": message, "saveToSentItems": params.save_to_sent},
        )
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
        graph = get_graph(ctx)

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
        attached = []
        if params.attachments:
            attached = await attach_files(graph, draft_id, params.attachments)
        message = (
            f"📝 Draft created successfully!\n"
            f"**To:** {recipients}\n"
            f"**Subject:** {params.subject}\n"
            f"**Draft ID:** `{draft_id}`"
        )
        if attached:
            message += f"\n**Attachments:** {', '.join(attached)}"
        return message
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
        graph = get_graph(ctx)
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
        graph = get_graph(ctx)
        dest = await resolve_folder(graph, params.destination_folder)
        data = await graph.post(
            f"/me/messages/{params.message_id}/move",
            json_data={"destinationId": dest},
        )
        return f"✅ Message moved to '{params.destination_folder}'. New ID: `{data.get('id', 'N/A')}`"
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_delete_mail",
    annotations={
        "title": "Delete Email or Draft",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_delete_mail(params: DeleteMailInput, ctx: Context = None) -> str:
    """Delete an email or draft.

    By default this is a soft delete: the message is moved to Deleted Items and
    can be recovered. Set `permanent=True` to permanently delete it (cannot be
    recovered). Works for ordinary messages and for drafts.

    Returns:
        str: Confirmation of the deletion.
    """
    try:
        graph = get_graph(ctx)
        if params.permanent:
            await graph.post(f"/me/messages/{params.message_id}/permanentDelete")
            return f"🗑️ Message `{params.message_id}` permanently deleted (not recoverable)."
        await graph.delete(f"/me/messages/{params.message_id}")
        return f"🗑️ Message `{params.message_id}` moved to Deleted Items."
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
        graph = get_graph(ctx)
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
        graph = get_graph(ctx)
        tree = await format_folder_tree(
            graph, "/me/mailFolders", params.top, params.include_subfolders
        )
        if not tree:
            return "No mail folders found."
        return "📁 **Mail Folders**\n\n" + tree
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_list_attachments",
    annotations={
        "title": "List Email Attachments",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_list_attachments(params: ListAttachmentsInput, ctx: Context = None) -> str:
    """List all attachments for a specific email message.

    Returns metadata for each attachment including name, type, size, and ID.
    Use the attachment ID with outlook_get_attachment to download.

    Args:
        params: Message ID to retrieve attachments from

    Returns:
        str: Formatted list of attachment metadata
    """
    try:
        graph = get_graph(ctx)
        endpoint = f"/me/messages/{params.message_id}/attachments"

        # Get all attachments (Graph API auto-paginates for small result sets)
        # Note: @odata.type is automatically included by Graph API, don't specify in $select
        data = await graph.get(endpoint, params={
            "$select": "id,name,contentType,size,isInline,lastModifiedDateTime"
        })

        attachments = data.get("value", [])

        if not attachments:
            return f"No attachments found for message `{params.message_id}`"

        result = f"📎 **Attachments** ({len(attachments)} total)\n\n"

        for att in attachments:
            result += format_attachment_summary(att) + "\n\n---\n\n"

        # Hint for next steps
        result += "\n*Use `outlook_get_attachment` with message_id + attachment_id to download.*"

        return result

    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_get_attachment",
    annotations={
        "title": "Download Email Attachment",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_get_attachment(params: GetAttachmentInput, ctx: Context = None) -> str:
    """Download a specific attachment from an email.

    Handles three types of Graph API attachments:
    - fileAttachment: Regular files (most common) → saved to configured download path
    - itemAttachment: Embedded emails or calendar items → metadata only
    - referenceAttachment: Cloud file links (OneDrive, SharePoint) → URL provided

    All file attachments are saved to disk (base64 streaming is too heavy for MCP).
    Download path can be customized via OUTLOOK_DOWNLOAD_PATH env var
    (default: ~/Downloads/outlook_attachments/).

    Args:
        params: Message ID and attachment ID

    Returns:
        str: File path (for fileAttachment) or metadata (for other types)
    """
    try:
        graph = get_graph(ctx)
        endpoint = f"/me/messages/{params.message_id}/attachments/{params.attachment_id}"

        # Get attachment metadata and content
        data = await graph.get(endpoint)

        att_type = data.get("@odata.type", "")
        name = data.get("name", "attachment")
        content_type = data.get("contentType", "application/octet-stream")
        size_bytes = data.get("size", 0)

        result = f"# Attachment: {name}\n\n"
        result += f"**Type:** {att_type}\n"
        result += f"**Content-Type:** {content_type}\n"
        result += f"**Size:** {size_bytes:,} bytes ({size_bytes / 1024 / 1024:.2f} MB)\n\n"

        # Handle different attachment types
        if att_type == "#microsoft.graph.fileAttachment":
            # Regular file attachment
            content_bytes_b64 = data.get("contentBytes")
            if not content_bytes_b64:
                return result + "Error: No content available for this file attachment."

            # Decode base64
            try:
                content_bytes = base64.b64decode(content_bytes_b64)
            except Exception as e:
                return result + f"Error decoding base64 content: {e}"

            # Decide: disk or inline?
            if should_save_to_disk(content_type, size_bytes, params.save_to_disk):
                # Save to disk
                try:
                    file_path = save_attachment_to_disk(name, content_bytes)
                    result += f"✅ **Saved to disk:**\n`{file_path}`\n\n"
                    result += "*File is ready to access on your local system.*"
                    return result
                except Exception as e:
                    return result + f"Error saving to disk: {e}"
            else:
                # Return as base64 data URL
                data_url = create_data_url(content_type, content_bytes_b64)

                # For images, Claude can render them directly
                if content_type in {"image/png", "image/jpeg", "image/jpg", "image/gif", "image/bmp", "image/webp"}:
                    result += "✅ **Image ready for viewing:**\n\n"
                    result += f"![{name}]({data_url})\n\n"
                    result += f"*Data URL: `{data_url[:80]}...` ({len(data_url)} chars)*"
                else:
                    # For PDFs and text files, provide data URL
                    result += f"✅ **Content available as base64 data URL:**\n\n"
                    result += f"```\n{data_url[:200]}...\n```\n\n"
                    result += f"*Full data URL length: {len(data_url)} characters*\n"
                    result += "*You can analyze this content or ask to save it to disk.*"

                return result

        elif att_type == "#microsoft.graph.itemAttachment":
            # Embedded email or calendar item
            item = data.get("item", {})
            item_type = item.get("@odata.type", "unknown")
            result += f"**Item Type:** {item_type}\n\n"

            if "#microsoft.graph.message" in item_type:
                # Embedded email
                result += "**This is an embedded email message:**\n\n"
                from_addr = item.get("from", {}).get("emailAddress", {}).get("address", "N/A")
                result += f"- Subject: {item.get('subject', 'N/A')}\n"
                result += f"- From: {from_addr}\n"
                result += f"- Received: {item.get('receivedDateTime', 'N/A')}\n\n"
                result += "*Item attachments cannot be downloaded as files. Use the metadata above.*"
            elif "#microsoft.graph.event" in item_type:
                # Embedded calendar event
                result += "**This is an embedded calendar event:**\n\n"
                result += f"- Subject: {item.get('subject', 'N/A')}\n"
                start_dt = item.get("start", {}).get("dateTime", "N/A")
                end_dt = item.get("end", {}).get("dateTime", "N/A")
                result += f"- Start: {start_dt}\n"
                result += f"- End: {end_dt}\n\n"
                result += "*Item attachments cannot be downloaded as files. Use the metadata above.*"
            else:
                result += "*Unknown item attachment type. Metadata only.*"

            return result

        elif att_type == "#microsoft.graph.referenceAttachment":
            # Cloud file reference (OneDrive, SharePoint)
            result += "**This is a cloud file reference (link):**\n\n"

            source_url = data.get("sourceUrl")
            permission_type = data.get("permission", "unknown")

            if source_url:
                result += f"**URL:** {source_url}\n"
            result += f"**Permission:** {permission_type}\n\n"
            result += "*Reference attachments are links to cloud files. Open the URL to access.*"

            return result

        else:
            # Unknown type
            result += "**Unknown attachment type.**\n"
            result += f"Raw data:\n```json\n{json.dumps(data, indent=2)[:500]}\n```"
            return result

    except Exception as e:
        return handle_graph_error(e)
