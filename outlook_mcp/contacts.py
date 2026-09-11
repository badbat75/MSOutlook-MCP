"""Contacts: the whole address book, matched the way a person looks for someone.

Graph's /me/contacts is the default Contacts folder only. A contact filed in any
other folder is at /me/contactFolders/{id}/contacts, and folders nest; the
mailbox this was written against keeps a folder a phone syncs into inside the
default one. A duplicate is exactly the contact that tends to live elsewhere, so
reading the address book means walking every folder.

Nor can Graph find a contact by name: on a contact, $filter reaches only
emailAddresses/any(a:a/address eq '...'), with no contains() and no startswith().
Matching is therefore done here, over every contact, which a personal address
book keeps to a few hundred.
"""

import re
import unicodedata
from datetime import date
from typing import List, Optional, Tuple

import anyio

from .auth import GraphClient

# Big enough that a folder of a personal address book arrives in one page.
PAGE_SIZE = 500

# Between two looks for a deleted contact in Deleted Items; see in_deleted_items().
RECHECK_SECONDS = 1.5

# What a listing needs of each contact. get_contact reads everything instead.
LIST_FIELDS = (
    "id,displayName,givenName,middleName,surname,nickName,fileAs,"
    "emailAddresses,mobilePhone,homePhones,businessPhones,birthday,parentFolderId"
)

# The fields a search looks at: every name the contact goes by, and its addresses.
_NAME_FIELDS = ("displayName", "givenName", "middleName", "surname", "nickName", "fileAs")

# Outlook's own time of day for a birthday, the one every birthday in the
# mailbox this was written against carries (21 of 21, in both folders). Close to
# noon UTC, the instant falls on the same calendar day from UTC-11 to UTC+12;
# the Birthdays calendar shows its UTC date.
BIRTHDAY_TIME = "T11:59:00Z"

_ISO_DATE = re.compile(r"\d{4}-\d{2}-\d{2}")
_MONTH_DAY = re.compile(r"--(\d{2})-(\d{2})")


async def read_all(graph: GraphClient, endpoint: str, params: dict) -> List[dict]:
    """Every item of a collection, following @odata.nextLink to the end."""
    data = await graph.get(endpoint, params=params)
    items = list(data.get("value", []))
    while data.get("@odata.nextLink"):
        # The link carries the query already, and it is absolute: GraphClient
        # passes it through as it is.
        data = await graph.get(data["@odata.nextLink"])
        items.extend(data.get("value", []))
    return items


async def contact_folders(graph: GraphClient) -> List[Tuple[str, str]]:
    """Every contact folder as (id, display name), the default one first.

    /me/contactFolders lists the folders inside the default one, not the
    default itself, which answers to the well-known name "contacts".
    """
    default = await graph.get("/me/contactFolders/contacts", params={"$select": "id,displayName"})
    folders = [(default["id"], default.get("displayName") or "Contacts")]

    async def walk(endpoint: str) -> None:
        for folder in await read_all(graph, endpoint, {"$top": 100, "$select": "id,displayName"}):
            folders.append((folder["id"], folder.get("displayName") or "(unnamed folder)"))
            await walk(f"/me/contactFolders/{folder['id']}/childFolders")

    await walk("/me/contactFolders")
    return folders


async def read_address_book(graph: GraphClient) -> List[dict]:
    """Every contact in every folder, each tagged with its folder's name as folderName."""
    contacts: List[dict] = []
    seen = set()
    for folder_id, folder_name in await contact_folders(graph):
        items = await read_all(
            graph,
            f"/me/contactFolders/{folder_id}/contacts",
            {"$top": PAGE_SIZE, "$select": LIST_FIELDS},
        )
        for contact in items:
            if contact["id"] in seen:
                continue
            seen.add(contact["id"])
            contact["folderName"] = folder_name
            contacts.append(contact)
    return contacts


async def folder_name(graph: GraphClient, folder_id: Optional[str]) -> Optional[str]:
    """The display name of one contact folder, or None if it cannot be read."""
    if not folder_id:
        return None
    try:
        folder = await graph.get(f"/me/contactFolders/{folder_id}", params={"$select": "displayName"})
    except Exception:
        # Only a label: a contact whose folder cannot be named is still shown.
        return None
    return folder.get("displayName")


async def in_deleted_items(graph: GraphClient, contact: dict, attempts: int = 3) -> bool:
    """Whether a contact that was just deleted is in Deleted Items.

    Graph can answer 404 to a DELETE it has carried out. Seen three times on a
    personal mailbox, on contacts given a birthday seconds before, each found in
    Deleted Items afterwards; not reproduced in a dozen tries since. The move
    changes the ID, so the contact is recognised by what a move leaves alone,
    its name and creation time. It is looked for a few times, in case the
    folder is slow to show an item just moved into it.
    """
    for attempt in range(attempts):
        if attempt:
            await anyio.sleep(RECHECK_SECONDS)
        items = await read_all(
            graph,
            "/me/contactFolders/deleteditems/contacts",
            {"$top": PAGE_SIZE, "$select": "displayName,createdDateTime"},
        )
        if any(
            item.get("displayName") == contact.get("displayName")
            and item.get("createdDateTime") == contact.get("createdDateTime")
            for item in items
        ):
            return True
    return False


def _fold(text: str) -> str:
    """Text as a search compares it: no case, no accents ("Papà" is "papa")."""
    decomposed = unicodedata.normalize("NFKD", text)
    return "".join(c for c in decomposed if not unicodedata.combining(c)).casefold()


def matches(contact: dict, query: str) -> bool:
    """Whether any name the contact goes by, or any of its addresses, contains `query`."""
    wanted = _fold(query.strip())
    haystack = [contact.get(field) or "" for field in _NAME_FIELDS]
    for email in contact.get("emailAddresses") or []:
        haystack += [email.get("address") or "", email.get("name") or ""]
    return any(wanted in _fold(value) for value in haystack)


def month_day(value: str) -> Tuple[int, int]:
    """(month, day) of a day of the year, given as 'YYYY-MM-DD' or '--MM-DD'.

    Never 'MM-DD' or 'DD/MM': "07-06" is 7 June to half the people who would
    write it and 6 July to the other half. A full date is unambiguous, and the
    date of an entry in the Birthdays calendar is exactly what a caller holds.
    """
    text = value.strip()
    match = _MONTH_DAY.fullmatch(text)
    try:
        if match:
            # A leap year, so that 29 February is a day a birthday can be on.
            day = date(2000, int(match.group(1)), int(match.group(2)))
        elif _ISO_DATE.fullmatch(text):
            day = date.fromisoformat(text)
        else:
            raise ValueError
    except ValueError:
        raise ValueError(
            f"'{value}' is not a day of the year: write an ISO date such as "
            f"'2026-06-07' (the year is ignored) or '--06-07'."
        ) from None
    return day.month, day.day


def birthday_date(value: Optional[str]) -> Optional[date]:
    """The day a stored birthday stands for: the UTC date of its instant.

    That is the date the Birthdays calendar shows for it, measured on a real
    mailbox, where every birthday sits at BIRTHDAY_TIME. None for no birthday.
    """
    if not value:
        return None
    try:
        return date.fromisoformat(value[:10])
    except ValueError:
        return None


def born_on(contact: dict, month: int, day: int) -> bool:
    birthday = birthday_date(contact.get("birthday"))
    return birthday is not None and (birthday.month, birthday.day) == (month, day)


def birthday_value(value: str) -> Optional[str]:
    """What to send Graph for a birthday given as 'YYYY-MM-DD'; '' removes it.

    A birthday is a DateTimeOffset to Graph, so the day is sent at
    BIRTHDAY_TIME, the time Outlook itself writes.
    """
    text = value.strip()
    if not text:
        return None
    if not _ISO_DATE.fullmatch(text):
        raise ValueError(f"'{value}' is not a date: write the birthday as 'YYYY-MM-DD', e.g. '2016-06-07'.")
    try:
        day = date.fromisoformat(text)
    except ValueError:
        raise ValueError(f"'{value}' is not a real date.") from None
    return f"{day.isoformat()}{BIRTHDAY_TIME}"


def sort_key(contact: dict) -> Tuple[bool, str]:
    """By name, ignoring case and accents; the nameless (a bare phone number) last."""
    name = _fold(contact.get("displayName") or contact.get("fileAs") or "")
    return not name, name
