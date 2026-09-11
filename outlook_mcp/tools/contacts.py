"""Contact tools: finding, creating, reading, correcting, moving and removing
contacts, and the folders they are filed in.

Enough to merge duplicates: find them with outlook_list_contacts, see all a
duplicate holds with outlook_get_contact, carry it over to the contact being
kept with outlook_update_contact, then outlook_delete_contact the rest. And to
fold one contact folder into another: outlook_move_contacts, then
outlook_delete_contact_folder once it is empty.
"""

from typing import Any, Dict

import httpx
from mcp.server.mcpserver import Context

from ..app import mcp
from ..contacts import (
    ALREADY_THERE,
    HAS_PHOTO,
    INCOMPLETE_COPY,
    MOVED,
    PAGE_SIZE,
    HalfMoved,
    birthday_value,
    born_on,
    contact_folders,
    find_folder,
    folder_name,
    in_deleted_items,
    item_count,
    matches,
    month_day,
    move_contact,
    read_address_book,
    read_all,
    sort_key,
)
from ..credentials import get_graph
from ..helpers import format_contact_details, format_contact_summary, handle_graph_error
from ..models import (
    CreateContactInput,
    DeleteContactFolderInput,
    DeleteContactInput,
    GetContactInput,
    ListContactsInput,
    MoveContactsInput,
    UpdateContactInput,
)

# What a create or an update may set, by the name the tool takes and Graph's.
_GRAPH_NAMES = {
    "display_name": "displayName",
    "given_name": "givenName",
    "surname": "surname",
    "nickname": "nickName",
    "company_name": "companyName",
    "mobile_phone": "mobilePhone",
    "home_phones": "homePhones",
    "business_phones": "businessPhones",
    "personal_notes": "personalNotes",
}


def _graph_fields(params) -> Dict[str, Any]:
    """What a create or an update was given, as Graph properties; None is not given.

    Email addresses are left to each tool: an update keeps the names the
    contact's addresses already had, which a new contact has none of.
    """
    fields = {
        graph_name: getattr(params, name)
        for name, graph_name in _GRAPH_NAMES.items()
        if getattr(params, name) is not None
    }
    if params.birthday is not None:
        fields["birthday"] = birthday_value(params.birthday)
    return fields


def _plural(count: int, noun: str) -> str:
    return f"{count} {noun}" if count == 1 else f"{count} {noun}s"


@mcp.tool(
    name="outlook_list_contacts",
    annotations={
        "title": "List Contacts",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_list_contacts(params: ListContactsInput, ctx: Context = None) -> str:
    """List contacts from every contact folder: name, email, phone, birthday, folder and ID.

    Reads the whole address book, the default Contacts folder and every other
    one (such as a folder a phone syncs into), so a duplicate filed elsewhere
    shows up beside its original. `search` narrows it to a name or an email
    address; `birthday` to one day of the year, which is how to find the
    contacts behind an entry of the Birthdays calendar: that calendar shows only
    the day and month, and a contact with a birthday is what generates it.
    Sorted by display name.

    Returns:
        str: Formatted list of contact summaries with their IDs.
    """
    try:
        graph = get_graph(ctx)
        folders = await contact_folders(graph)
        place = "in all folders"
        if params.folder:
            folders = [find_folder(folders, params.folder)]
            place = f"in {folders[0].name}"
        contacts = await read_address_book(graph, folders)

        criteria = []
        if params.search:
            contacts = [c for c in contacts if matches(c, params.search)]
            criteria.append(f"matching '{params.search}'")
        if params.birthday:
            month, day = month_day(params.birthday)
            contacts = [c for c in contacts if born_on(c, month, day)]
            criteria.append(f"born on {day:02d}/{month:02d} (day/month)")
        described = f" {' and '.join(criteria)}" if criteria else ""

        if not contacts:
            return f"No contacts{described}{' ' + place if params.folder else ''}."

        contacts.sort(key=sort_key)
        page = contacts[params.skip: params.skip + params.top]
        if not page:
            return f"{len(contacts)} contacts{described}, none after skip={params.skip}."

        result = f"👥 **Contacts**{described}: {len(contacts)} {place}"
        if len(page) < len(contacts):
            result += f", showing {params.skip + 1}-{params.skip + len(page)}"
        result += "\n\n"
        for contact in page:
            result += format_contact_summary(contact) + "\n\n---\n\n"

        if params.skip + len(page) < len(contacts):
            result += f"\n*More contacts available. Use skip={params.skip + len(page)} for the next page.*"
        return result
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_get_contact",
    annotations={
        "title": "Get Contact Details",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_get_contact(params: GetContactInput, ctx: Context = None) -> str:
    """Get everything one contact holds: every name, email, phone, address, date and note.

    Read a duplicate with this before deleting it: whatever it has that the
    contact being kept lacks is lost with it unless carried over first.

    Returns:
        str: Complete contact details.
    """
    try:
        graph = get_graph(ctx)
        data = await graph.get(f"/me/contacts/{params.contact_id}")
        folder = await folder_name(graph, data.get("parentFolderId"))
        return format_contact_details(data, folder)
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_create_contact",
    annotations={
        "title": "Create Contact",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def outlook_create_contact(params: CreateContactInput, ctx: Context = None) -> str:
    """Create a contact, in the default Contacts folder or in the folder named.

    Needs at least a name, an email address or a phone number. A birthday puts
    the contact in the Birthdays calendar, as Outlook does for any contact that
    has one.

    Returns:
        str: Confirmation, with the contact as created and its ID.
    """
    try:
        graph = get_graph(ctx)
        folder = find_folder(await contact_folders(graph), params.folder)
        body = _graph_fields(params)
        if params.email_addresses:
            body["emailAddresses"] = [{"address": a, "name": a} for a in params.email_addresses]
        data = await graph.post(f"/me/contactFolders/{folder.id}/contacts", json_data=body)
        data["folderName"] = folder.name
        return f"✅ Contact created in **{folder.name}**.\n\n" + format_contact_summary(data)
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_update_contact",
    annotations={
        "title": "Update Contact",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_update_contact(params: UpdateContactInput, ctx: Context = None) -> str:
    """Update a contact's names, email addresses, phone numbers, birthday or notes.

    Fields left out are not touched. A list (email addresses, home or business
    phones) replaces what the contact has, so to merge a duplicate into this
    contact pass both contacts' entries together. The display name stays as it
    is unless display_name is given, even when the name parts change.

    Returns:
        str: Confirmation, with the contact as it now stands.
    """
    try:
        graph = get_graph(ctx)
        changes = _graph_fields(params)
        if params.email_addresses is not None:
            changes["emailAddresses"] = params.email_addresses
        if not changes:
            return "No updates specified."
        changed = ", ".join(changes)

        current = await graph.get(
            f"/me/contacts/{params.contact_id}",
            params={"$select": "displayName,emailAddresses"},
        )
        if params.email_addresses is not None:
            # An address the contact already had keeps the name it had.
            names = {
                (e.get("address") or "").casefold(): e.get("name")
                for e in current.get("emailAddresses") or []
            }
            changes["emailAddresses"] = [
                {"address": address, "name": names.get(address.casefold()) or address}
                for address in params.email_addresses
            ]
        if "displayName" not in changes and current.get("displayName"):
            # Graph may regenerate the display name from the other properties of
            # an update that leaves it out, as its own documentation warns, and
            # the display name is what the contact is known by: it goes along.
            changes["displayName"] = current["displayName"]

        data = await graph.patch(f"/me/contacts/{params.contact_id}", json_data=changes)
        return f"✅ Contact updated ({changed}).\n\n" + format_contact_summary(data)
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_delete_contact",
    annotations={
        "title": "Delete Contact",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_delete_contact(params: DeleteContactInput, ctx: Context = None) -> str:
    """Delete a contact, in whichever folder it is filed.

    It is moved to Deleted Items, where it can be recovered from Outlook. Its
    entry in the Birthdays calendar, if it had a birthday, goes with it.

    Returns:
        str: Confirmation naming the contact deleted.
    """
    try:
        graph = get_graph(ctx)
        # Named in the answer, so a transcript of a clean-up says who went
        # rather than which opaque ID. The creation time is what recognises it
        # in Deleted Items if Graph misreports the delete.
        contact = await graph.get(
            f"/me/contacts/{params.contact_id}",
            params={"$select": "displayName,createdDateTime"},
        )
        name = contact.get("displayName") or "(no name)"
        try:
            await graph.delete(f"/me/contacts/{params.contact_id}")
        except httpx.HTTPStatusError as e:
            # The contact was there a moment ago, so a 404 now may be Graph
            # misreporting a delete it carried out (see in_deleted_items).
            if e.response.status_code != 404 or not await in_deleted_items(graph, contact):
                raise
        return f"🗑️ Contact **{name}** moved to Deleted Items. ID: `{params.contact_id}`"
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_list_contact_folders",
    annotations={
        "title": "List Contact Folders",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_list_contact_folders(ctx: Context = None) -> str:
    """List the contact folders, nested as they are, with how many contacts each holds.

    The default Contacts folder comes first, and a folder inside it, such as
    one a phone syncs into, is indented beneath it. A folder's name or ID is
    what outlook_list_contacts, outlook_create_contact, outlook_move_contacts
    and outlook_delete_contact_folder take.

    Returns:
        str: Every contact folder with its contact count and ID.
    """
    try:
        graph = get_graph(ctx)
        folders = await contact_folders(graph)
        lines = []
        total = 0
        for folder in folders:
            listed = len(await read_all(
                graph, f"/me/contactFolders/{folder.id}/contacts", {"$top": PAGE_SIZE, "$select": "id"}
            ))
            counted = await item_count(graph, folder.id) or 0
            total += listed
            line = f"{'  ' * folder.depth}- **{folder.name}**"
            if folder.depth == 0:
                line += " (default)"
            line += f": {_plural(listed, 'contact')}"
            if counted > listed:
                line += f" and {_plural(counted - listed, 'other item')} not listed as contacts"
            lines.append(f"{line} | ID: `{folder.id}`")
        return (
            f"📇 **Contact folders**: {len(folders)}, {_plural(total, 'contact')} in all\n\n"
            + "\n".join(lines)
        )
    except Exception as e:
        return handle_graph_error(e)


@mcp.tool(
    name="outlook_move_contacts",
    annotations={
        "title": "Move Contacts to a Folder",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_move_contacts(params: MoveContactsInput, ctx: Context = None) -> str:
    """Move contacts into another contact folder, e.g. out of one a phone no longer syncs into.

    Graph cannot move a contact, so each is copied into the folder with every
    property it has (names, addresses, phones, birthday, notes, categories),
    and the original goes to Deleted Items once the copy is confirmed. A moved
    contact is therefore a new item, with a new ID and today's creation date;
    its Birthdays calendar entry follows it. A contact with a photo is left
    where it is, and the answer says so: move that one in Outlook.

    At most 25 per call, one after the other. An unexpected error stops the
    call, and the answer lists what was not attempted.

    Returns:
        str: What became of each contact, with the new IDs.
    """
    try:
        graph = get_graph(ctx)
        destination = find_folder(await contact_folders(graph), params.destination_folder)
    except Exception as e:
        return handle_graph_error(e)

    lines = []
    moved = 0
    for index, contact_id in enumerate(params.contact_ids):
        try:
            outcome = await move_contact(graph, contact_id, destination)
        except Exception as e:
            lines.append(f"❌ {e}" if isinstance(e, HalfMoved) else f"❌ `{contact_id}`: {handle_graph_error(e)}")
            rest = params.contact_ids[index + 1:]
            if rest:
                lines.append("Stopped there. Not attempted: " + ", ".join(f"`{c}`" for c in rest))
            break
        if outcome.status == MOVED:
            moved += 1
            lines.append(f"✅ **{outcome.name}**: new ID `{outcome.detail}`")
        elif outcome.status == ALREADY_THERE:
            lines.append(f"⏭️ **{outcome.name}**: already in {destination.name}")
        elif outcome.status == HAS_PHOTO:
            lines.append(
                f"⚠️ **{outcome.name}**: not moved, it has a photo. A copy carrying "
                f"the same photo was seen to vanish once its original was deleted: "
                f"move this one in Outlook."
            )
        elif outcome.status == INCOMPLETE_COPY:
            lines.append(
                f"⚠️ **{outcome.name}**: not moved, the copy came back without "
                f"{outcome.detail}. The copy was deleted; the original is untouched."
            )
        else:
            lines.append(f"❌ `{contact_id}`: no such contact")

    result = f"📇 **Moved {moved} of {len(params.contact_ids)}** to **{destination.name}**\n\n"
    result += "\n".join(lines)
    if moved:
        result += "\n\n*The originals are in Deleted Items.*"
    return result


@mcp.tool(
    name="outlook_delete_contact_folder",
    annotations={
        "title": "Delete Contact Folder",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_delete_contact_folder(params: DeleteContactFolderInput, ctx: Context = None) -> str:
    """Delete an empty contact folder, such as one a phone no longer syncs into.

    Only a folder with nothing left in it: no contacts, no other items and no
    subfolders. Move its contacts out first with outlook_move_contacts, or
    delete them. It goes to Deleted Items, where it can be recovered from
    Outlook. The default Contacts folder cannot be deleted.

    Returns:
        str: Confirmation naming the folder, or why it was left alone.
    """
    try:
        graph = get_graph(ctx)
        folders = await contact_folders(graph)
        folder = find_folder(folders, params.folder)
        if folder.depth == 0:
            return f"Not deleted: **{folder.name}** is the default contact folder, which cannot be deleted."
        children = [f.name for f in folders if f.parent_id == folder.id]
        if children:
            return (
                f"Not deleted: **{folder.name}** has subfolders ({', '.join(children)}). "
                f"Empty and delete those first."
            )
        listed = len(await read_all(
            graph, f"/me/contactFolders/{folder.id}/contacts", {"$top": PAGE_SIZE, "$select": "id"}
        ))
        # Graph's own count as well: an item it does not list as a contact would
        # still go to Deleted Items with the folder.
        counted = await item_count(graph, folder.id) or 0
        if listed or counted:
            held = _plural(listed, "contact")
            if counted > listed:
                held += f" and {_plural(counted - listed, 'other item')} not listed as contacts"
            return (
                f"Not deleted: **{folder.name}** still holds {held}. Move them out "
                f"with outlook_move_contacts, or delete them, first."
            )
        await graph.delete(f"/me/contactFolders/{folder.id}")
        return f"🗑️ Contact folder **{folder.name}** moved to Deleted Items."
    except Exception as e:
        return handle_graph_error(e)
