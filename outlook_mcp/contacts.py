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
from typing import List, NamedTuple, Optional, Tuple

import anyio
import httpx

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


class ContactFolder(NamedTuple):
    """A contact folder: what it is called and where it sits in the tree."""

    id: str
    name: str
    # 0 for the default folder, 1 for a folder inside it, and so on.
    depth: int = 0
    parent_id: Optional[str] = None


# What a caller writes for the default folder, whatever its display name:
# Graph's own well-known name for it.
DEFAULT_FOLDER = "contacts"


async def contact_folders(graph: GraphClient) -> List[ContactFolder]:
    """Every contact folder, the default one first and each folder before its children.

    /me/contactFolders lists the folders inside the default one, not the
    default itself, which answers to the well-known name "contacts".
    """
    default = await graph.get("/me/contactFolders/contacts", params={"$select": "id,displayName"})
    folders = [ContactFolder(default["id"], default.get("displayName") or "Contacts")]

    async def walk(endpoint: str, parent_id: str, depth: int) -> None:
        for folder in await read_all(graph, endpoint, {"$top": 100, "$select": "id,displayName"}):
            name = folder.get("displayName") or "(unnamed folder)"
            folders.append(ContactFolder(folder["id"], name, depth, parent_id))
            await walk(f"/me/contactFolders/{folder['id']}/childFolders", folder["id"], depth + 1)

    await walk("/me/contactFolders", default["id"], 1)
    return folders


def find_folder(folders: List[ContactFolder], ref: str) -> ContactFolder:
    """The folder a caller names: its ID, its display name, or 'contacts' for the default one.

    A name has to match whole, ignoring case and accents. Contacts are moved
    into a folder found this way, and folders deleted, so a fragment that
    happens to name a folder nobody meant is not good enough.
    """
    wanted = ref.strip()
    if wanted.casefold() == DEFAULT_FOLDER:
        return folders[0]
    for folder in folders:
        if folder.id == wanted:
            return folder
    named = [folder for folder in folders if _fold(folder.name) == _fold(wanted)]
    if len(named) == 1:
        return named[0]
    if named:
        raise ValueError(
            f"{len(named)} contact folders are called '{ref}': name the one meant by "
            f"its ID, as outlook_list_contact_folders shows it."
        )
    listed = ", ".join(f"'{folder.name}'" for folder in folders)
    raise ValueError(f"No contact folder '{ref}'. The contact folders are: {listed}.")


async def item_count(graph: GraphClient, folder_id: str) -> Optional[int]:
    """How many items Graph counts in a contact folder, or None if it does not say.

    It can be more than the contacts the folder lists: on the mailbox this was
    written against, the default folder counted 164 items and listed 162
    contacts. Whatever the other two are, deleting the folder would take them
    along, so the delete looks at this number as well as at the contacts.
    """
    data = await graph.get(
        f"/me/contactFolders/{folder_id}/contacts",
        params={"$count": "true", "$top": 1, "$select": "id"},
    )
    return data.get("@odata.count")


async def read_address_book(
    graph: GraphClient, folders: Optional[List[ContactFolder]] = None
) -> List[dict]:
    """Every contact in `folders`, or in every folder when None.

    Each contact is tagged with its folder's name as folderName.
    """
    if folders is None:
        folders = await contact_folders(graph)
    contacts: List[dict] = []
    seen = set()
    for folder in folders:
        items = await read_all(
            graph,
            f"/me/contactFolders/{folder.id}/contacts",
            {"$top": PAGE_SIZE, "$select": LIST_FIELDS},
        )
        for contact in items:
            if contact["id"] in seen:
                continue
            seen.add(contact["id"])
            contact["folderName"] = folder.name
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


# Set by Graph on every contact, never by a caller.
_READ_ONLY = frozenset({"id", "changeKey", "createdDateTime", "lastModifiedDateTime", "parentFolderId"})

# A personal mailbox also returns a contact's first three email addresses a
# second time, one property each, which the v1.0 contact resource does not
# document. They repeated emailAddresses exactly in all 277 contacts of the
# mailbox this was written against, so a copy carries them there alone.
_ADDRESS_ECHOES = frozenset({"primaryEmailAddress", "secondaryEmailAddress", "tertiaryEmailAddress"})


def copy_of(contact: dict) -> dict:
    """What to create a copy of `contact` from: every property it has that a caller can write.

    Graph can neither move nor copy a contact: POST .../contacts/{id}/move is
    "Resource not found for the segment 'move'", and so is copy. Moving one
    means creating it anew elsewhere, and a copy made from a full GET this way
    came back equal to its original in every property, on a real mailbox.
    """
    return {
        key: value for key, value in contact.items()
        if not key.startswith("@") and key not in _READ_ONLY and key not in _ADDRESS_ECHOES
    }


def _has_value(value) -> bool:
    """Whether a property holds anything: an address made of empty strings does not."""
    if isinstance(value, dict):
        return any(_has_value(v) for v in value.values())
    return value not in (None, "", [])


def lost_fields(original: dict, copy: dict) -> List[str]:
    """The properties the original has a value for that the copy does not hold the same."""
    return sorted(
        key for key, value in copy_of(original).items()
        if _has_value(value) and copy.get(key) != value
    )


# What became of one contact asked to move; see move_contact().
MOVED = "moved"
ALREADY_THERE = "already there"
HAS_PHOTO = "has a photo"
NOT_FOUND = "not found"
INCOMPLETE_COPY = "incomplete copy"


class MoveOutcome(NamedTuple):
    status: str
    name: str
    # The new ID once moved; the properties a copy lacked when incomplete.
    detail: str = ""


class HalfMoved(Exception):
    """A step after the copy failed, so the contact now exists twice. The message says which is which."""


async def move_contact(graph: GraphClient, contact_id: str, destination: ContactFolder) -> MoveOutcome:
    """Move one contact into `destination`: copy it there, then delete the original.

    The original goes to Deleted Items only once the copy is known to hold all
    it held. A contact with a photo is not touched (HAS_PHOTO): measured on a
    real mailbox, a copy carrying the original's photo vanished, from every
    folder and from Deleted Items alike, within 10 to 45 seconds of the
    original being deleted. Seen three times, with nothing but the photo in
    common, while copies without one (categories, birthdays, email addresses)
    survived every time.

    Raises HalfMoved when a step after the copy fails; anything else it raises
    has changed nothing.
    """
    try:
        original = await graph.get(f"/me/contacts/{contact_id}")
    except httpx.HTTPStatusError as e:
        # 404 for an ID that is gone, 400 for a string that never was one.
        if e.response.status_code in (400, 404):
            return MoveOutcome(NOT_FOUND, contact_id)
        raise
    name = original.get("displayName") or "(no name)"
    if original.get("parentFolderId") == destination.id:
        return MoveOutcome(ALREADY_THERE, name)
    if await graph.get_bytes(f"/me/contacts/{contact_id}/photo/$value") is not None:
        return MoveOutcome(HAS_PHOTO, name)

    copy = await graph.post(f"/me/contactFolders/{destination.id}/contacts", json_data=copy_of(original))
    lost = lost_fields(original, copy)
    if lost:
        try:
            await graph.delete(f"/me/contacts/{copy['id']}")
        except Exception as e:
            raise HalfMoved(
                f"The copy of {name} made in {destination.name} lacks {', '.join(lost)}, "
                f"and could not be removed: delete it, `{copy['id']}`, with "
                f"outlook_delete_contact. The original is untouched."
            ) from e
        return MoveOutcome(INCOMPLETE_COPY, name, ", ".join(lost))

    try:
        await graph.delete(f"/me/contacts/{contact_id}")
    except Exception as e:
        # Graph has answered 404 to a delete it carried out (see in_deleted_items).
        missing = isinstance(e, httpx.HTTPStatusError) and e.response.status_code == 404
        if not (missing and await in_deleted_items(graph, original)):
            raise HalfMoved(
                f"{name} was copied into {destination.name} (new ID `{copy['id']}`), "
                f"but the original could not be deleted, so it is in both folders: "
                f"delete the original, `{contact_id}`, with outlook_delete_contact."
            ) from e
    return MoveOutcome(MOVED, name, copy["id"])


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
