"""Contact tools: finding, reading, correcting and removing contacts.

Enough to merge duplicates: find them with outlook_list_contacts, see all a
duplicate holds with outlook_get_contact, carry it over to the contact being
kept with outlook_update_contact, then outlook_delete_contact the rest.
"""

from typing import Any, Dict

import httpx
from mcp.server.mcpserver import Context

from ..app import mcp
from ..contacts import (
    birthday_value,
    born_on,
    folder_name,
    in_deleted_items,
    matches,
    month_day,
    read_address_book,
    sort_key,
)
from ..credentials import get_graph
from ..helpers import format_contact_details, format_contact_summary, handle_graph_error
from ..models import (
    DeleteContactInput,
    GetContactInput,
    ListContactsInput,
    UpdateContactInput,
)


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
        contacts = await read_address_book(graph)

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
            return f"No contacts{described}."

        contacts.sort(key=sort_key)
        page = contacts[params.skip: params.skip + params.top]
        if not page:
            return f"{len(contacts)} contacts{described}, none after skip={params.skip}."

        result = f"👥 **Contacts**{described}: {len(contacts)} in all folders"
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
        changes: Dict[str, Any] = {}
        if params.display_name is not None:
            changes["displayName"] = params.display_name
        if params.given_name is not None:
            changes["givenName"] = params.given_name
        if params.surname is not None:
            changes["surname"] = params.surname
        if params.nickname is not None:
            changes["nickName"] = params.nickname
        if params.mobile_phone is not None:
            changes["mobilePhone"] = params.mobile_phone
        if params.home_phones is not None:
            changes["homePhones"] = params.home_phones
        if params.business_phones is not None:
            changes["businessPhones"] = params.business_phones
        if params.birthday is not None:
            changes["birthday"] = birthday_value(params.birthday)
        if params.personal_notes is not None:
            changes["personalNotes"] = params.personal_notes
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
