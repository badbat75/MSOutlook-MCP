"""Pydantic input models for all MCP tools."""

from typing import Any, Literal, Optional, List, get_args

from pydantic import BaseModel, Field, field_validator, model_validator, ConfigDict

from .contacts import birthday_value, month_day

# An event's "Show as": Graph's freeBusyStatus, less "unknown", which is what
# Graph reports when it has nothing to say rather than a status anybody sets.
ShowAs = Literal["free", "tentative", "busy", "oof", "workingElsewhere"]
_SHOW_AS_BY_LOWER = {value.lower(): value for value in get_args(ShowAs)}

SHOW_AS_DESCRIPTION = (
    "How the event shows on the calendar (Outlook's 'Show as'): 'free', "
    "'tentative', 'busy', 'oof' (out of office) or 'workingElsewhere'."
)


def canonical_show_as(value: Any) -> Any:
    """Graph's spelling of a status given in any casing ('WORKINGELSEWHERE')."""
    if isinstance(value, str):
        return _SHOW_AS_BY_LOWER.get(value.strip().lower(), value)
    return value


class ListMailInput(BaseModel):
    """Input for listing emails."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    folder: str = Field(
        default="inbox",
        description=(
            "Mail folder: 'inbox', 'sentitems', 'drafts', 'deleteditems', 'junkemail', "
            "a subfolder display name (e.g. 'Centri Estivi'), or a folder ID. "
            "Use '*' (or 'all') to search/list across the entire mailbox, ignoring folders; "
            "useful with `search` when you don't know which folder a message is in."
        )
    )
    top: int = Field(default=10, description="Number of messages to return", ge=1, le=50)
    skip: int = Field(default=0, description="Number of messages to skip (pagination)", ge=0)
    filter: Optional[str] = Field(
        default=None,
        description="OData filter, e.g. 'isRead eq false' or \"from/emailAddress/address eq 'john@example.com'\""
    )
    search: Optional[str] = Field(
        default=None,
        description="Search query string to search across subject, body, and sender"
    )
    select: Optional[str] = Field(
        default=None,
        description="Comma-separated fields to return, e.g. 'subject,from,receivedDateTime'"
    )


class GetMailInput(BaseModel):
    """Input for getting a specific email."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    message_id: str = Field(..., description="The message ID to retrieve", min_length=1)
    include_body: bool = Field(default=True, description="Whether to include the full email body")


class SendMailInput(BaseModel):
    """Input for sending an email."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    to: List[str] = Field(..., description="List of recipient email addresses", min_length=1)
    subject: str = Field(..., description="Email subject line", min_length=1, max_length=500)
    body: str = Field(..., description="Email body content (HTML supported)")
    cc: Optional[List[str]] = Field(default=None, description="CC recipients")
    bcc: Optional[List[str]] = Field(default=None, description="BCC recipients")
    importance: str = Field(default="normal", description="'low', 'normal', or 'high'")
    is_html: bool = Field(default=True, description="Whether body is HTML (True) or plain text (False)")
    save_to_sent: bool = Field(default=True, description="Save a copy in Sent Items")
    attachments: Optional[List[str]] = Field(
        default=None,
        description=(
            "Local file paths to attach. Files up to 3MB are sent inline; larger "
            "files are uploaded via an upload session. Note: when attachments are "
            "present, the message is always saved to Sent Items (save_to_sent is ignored)."
        ),
    )

    @field_validator("importance")
    @classmethod
    def validate_importance(cls, v: str) -> str:
        if v.lower() not in ("low", "normal", "high"):
            raise ValueError("importance must be 'low', 'normal', or 'high'")
        return v.lower()


class CreateDraftInput(BaseModel):
    """Input for creating a draft email."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    to: List[str] = Field(..., description="List of recipient email addresses", min_length=1)
    subject: str = Field(..., description="Email subject line", min_length=1, max_length=500)
    body: str = Field(..., description="Email body content (HTML supported)")
    cc: Optional[List[str]] = Field(default=None, description="CC recipients")
    bcc: Optional[List[str]] = Field(default=None, description="BCC recipients")
    importance: str = Field(default="normal", description="'low', 'normal', or 'high'")
    is_html: bool = Field(default=True, description="Whether body is HTML (True) or plain text (False)")
    attachments: Optional[List[str]] = Field(
        default=None,
        description=(
            "Local file paths to attach. Files up to 3MB are sent inline; larger "
            "files are uploaded via an upload session."
        ),
    )

    @field_validator("importance")
    @classmethod
    def validate_importance(cls, v: str) -> str:
        if v.lower() not in ("low", "normal", "high"):
            raise ValueError("importance must be 'low', 'normal', or 'high'")
        return v.lower()


class ReplyMailInput(BaseModel):
    """Input for replying to an email."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    message_id: str = Field(..., description="ID of the message to reply to")
    comment: str = Field(..., description="Reply body text (HTML supported)")
    reply_all: bool = Field(default=False, description="Reply to all recipients")


class MoveMailInput(BaseModel):
    """Input for moving an email to a folder."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    message_id: str = Field(..., description="ID of the message to move")
    destination_folder: str = Field(
        ...,
        description="Target folder: 'inbox', 'archive', 'deleteditems', 'junkemail', or folder ID"
    )


class DeleteMailInput(BaseModel):
    """Input for deleting an email or draft."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    message_id: str = Field(..., description="ID of the message (or draft) to delete", min_length=1)
    permanent: bool = Field(
        default=False,
        description=(
            "If False (default), the message is moved to Deleted Items and can be "
            "recovered. If True, it is permanently deleted and cannot be recovered."
        ),
    )


class UpdateMailInput(BaseModel):
    """Input for updating email properties."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    message_id: str = Field(..., description="ID of the message to update")
    is_read: Optional[bool] = Field(default=None, description="Mark as read/unread")
    categories: Optional[List[str]] = Field(default=None, description="Set categories/labels")
    flag_status: Optional[str] = Field(
        default=None,
        description="'notFlagged', 'flagged', or 'complete'"
    )


class ListEventsInput(BaseModel):
    """Input for listing calendar events."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    start_date: Optional[str] = Field(
        default=None,
        description="Start date in ISO format (YYYY-MM-DD). Defaults to today."
    )
    end_date: Optional[str] = Field(
        default=None,
        description="End date in ISO format (YYYY-MM-DD). Defaults to 7 days from start."
    )
    top: int = Field(default=20, description="Max events to return", ge=1, le=50)
    calendar_id: Optional[str] = Field(
        default=None,
        description="Specific calendar ID. Omit to aggregate events across ALL "
                    "calendars (e.g. Calendar, Birthdays, Your Family)."
    )


class GetEventInput(BaseModel):
    """Input for getting a specific calendar event."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    event_id: str = Field(..., description="The event ID to retrieve")


class CreateEventInput(BaseModel):
    """Input for creating a calendar event."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    subject: str = Field(..., description="Event title/subject", min_length=1)
    start: str = Field(
        ...,
        description=(
            "Start datetime in ISO format, e.g. '2025-06-15T10:00:00'. "
            "For an all-day event a date is enough: '2025-06-15'."
        )
    )
    end: str = Field(
        ...,
        description=(
            "End datetime in ISO format, e.g. '2025-06-15T11:00:00'. For an "
            "all-day event, the last day as a date: '2025-06-15' for one day, "
            "'2025-06-16' for two."
        )
    )
    timezone: str = Field(
        default="UTC",
        description="Timezone for start/end, e.g. 'UTC', 'Europe/Rome', 'America/New_York'"
    )
    body: Optional[str] = Field(default=None, description="Event description/body (HTML supported)")
    location: Optional[str] = Field(default=None, description="Event location name")
    attendees: Optional[List[str]] = Field(default=None, description="List of attendee email addresses")
    is_online_meeting: bool = Field(default=False, description="Create as Teams meeting")
    reminder_minutes: int = Field(default=15, description="Reminder before event in minutes", ge=0)
    is_all_day: bool = Field(
        default=False,
        description="All-day event: start and end are widened to whole days, midnight to midnight",
    )
    show_as: Optional[ShowAs] = Field(
        default=None,
        description=SHOW_AS_DESCRIPTION + " Omit for Outlook's default, busy.",
    )
    recurrence: Optional[str] = Field(
        default=None,
        description="Recurrence pattern: 'daily', 'weekly', 'monthly', or null for none"
    )
    calendar_id: Optional[str] = Field(default=None, description="Target calendar ID (omit for default)")

    @field_validator("show_as", mode="before")
    @classmethod
    def validate_show_as(cls, v: Any) -> Any:
        return canonical_show_as(v)


class UpdateEventInput(BaseModel):
    """Input for updating a calendar event."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    event_id: str = Field(..., description="ID of the event to update")
    subject: Optional[str] = Field(default=None, description="New subject")
    start: Optional[str] = Field(
        default=None,
        description="New start datetime (ISO format). With is_all_day=true a date is enough.",
    )
    end: Optional[str] = Field(
        default=None,
        description=(
            "New end datetime (ISO format). With is_all_day=true, the last day "
            "as a date; omit it for a one-day event."
        ),
    )
    timezone: Optional[str] = Field(
        default=None,
        description=(
            "Timezone for start/end, e.g. 'Europe/Rome' (default UTC). When "
            "is_all_day is set without start, the event's current days are read "
            "in this zone, or in the event's own zone if omitted."
        ),
    )
    location: Optional[str] = Field(default=None, description="New location")
    body: Optional[str] = Field(default=None, description="New body content")
    is_all_day: Optional[bool] = Field(
        default=None,
        description=(
            "true makes it an all-day event, widened to the whole days it covers, "
            "midnight to midnight: the days of start/end when given, otherwise "
            "the days it is on now. false makes it a timed event again: pass the "
            "new start and end, or it keeps its current times. Omit to leave it."
        ),
    )
    show_as: Optional[ShowAs] = Field(
        default=None,
        description=SHOW_AS_DESCRIPTION + " Omit to leave it unchanged.",
    )
    is_cancelled: bool = Field(default=False, description="Cancel the event")

    @field_validator("show_as", mode="before")
    @classmethod
    def validate_show_as(cls, v: Any) -> Any:
        return canonical_show_as(v)


class DeleteEventInput(BaseModel):
    """Input for deleting a calendar event."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    event_id: str = Field(..., description="ID of the event to delete")


class RespondEventInput(BaseModel):
    """Input for responding to a calendar event invitation."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    event_id: str = Field(..., description="ID of the event to respond to")
    response: str = Field(
        ...,
        description="Response: 'accept', 'tentativelyAccept', or 'decline'"
    )
    comment: Optional[str] = Field(default=None, description="Optional message with your response")
    send_response: bool = Field(default=True, description="Send response to organizer")

    @field_validator("response")
    @classmethod
    def validate_response(cls, v: str) -> str:
        valid = ("accept", "tentativelyaccept", "decline")
        if v.lower() not in valid:
            raise ValueError(f"response must be one of: {', '.join(valid)}")
        return v.lower()


class ListMailFoldersInput(BaseModel):
    """Input for listing mail folders."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    top: int = Field(default=20, description="Max folders to return per level", ge=1, le=50)
    include_subfolders: bool = Field(
        default=True,
        description=(
            "Recurse into child folders and show the full nested hierarchy "
            "(e.g. subfolders under Inbox). Set False for top-level folders only."
        ),
    )


class ListCalendarsInput(BaseModel):
    """Input for listing calendars."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    top: int = Field(default=10, description="Max calendars to return", ge=1, le=50)


class ListAttachmentsInput(BaseModel):
    """Input for listing email attachments."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    message_id: str = Field(..., description="The message ID to retrieve attachments from", min_length=1)


class GetAttachmentInput(BaseModel):
    """Input for downloading a specific attachment."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    message_id: str = Field(..., description="The message ID containing the attachment", min_length=1)
    attachment_id: str = Field(..., description="The attachment ID to download", min_length=1)


class ListContactsInput(BaseModel):
    """Input for listing contacts across every contact folder."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    search: Optional[str] = Field(
        default=None,
        description=(
            "Only contacts with this text in a name or an email address, ignoring "
            "case and accents: 'gabriele', 'de simoni', 'papa' (finds 'Papà'), "
            "'@example.com'. Looks at display name, given name, middle name, "
            "surname, nickname, 'file as' and every email address."
        ),
    )
    birthday: Optional[str] = Field(
        default=None,
        description=(
            "Only contacts whose birthday falls on this day of the year: an ISO "
            "date, 'YYYY-MM-DD', whose year is ignored (e.g. the date of an entry "
            "in the Birthdays calendar, '2026-06-07' for 7 June), or '--MM-DD'."
        ),
    )
    folder: Optional[str] = Field(
        default=None,
        description=(
            "Only the contacts filed in this folder, not in its subfolders: its "
            "name as outlook_list_contact_folders shows it (case and accents "
            "ignored), its ID, or 'contacts' for the default folder. Omit for "
            "every folder."
        ),
    )
    top: int = Field(default=50, description="Max contacts to return", ge=1, le=200)
    skip: int = Field(default=0, description="Number of matching contacts to skip (pagination)", ge=0)

    @field_validator("birthday")
    @classmethod
    def validate_birthday(cls, v: Optional[str]) -> Optional[str]:
        if v:
            month_day(v)
        return v or None


class GetContactInput(BaseModel):
    """Input for reading one contact."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    contact_id: str = Field(..., description="The contact ID, as outlook_list_contacts shows it", min_length=1)


class UpdateContactInput(BaseModel):
    """Input for updating a contact. Omitted fields are left as they are."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    contact_id: str = Field(..., description="ID of the contact to update", min_length=1)
    display_name: Optional[str] = Field(
        default=None,
        description=(
            "New display name, the one the contact is listed and its birthday "
            "shown under. Omit to keep the current one, even when changing the "
            "name parts: it is never regenerated from them."
        ),
    )
    given_name: Optional[str] = Field(default=None, description="New given (first) name")
    surname: Optional[str] = Field(default=None, description="New surname (last name)")
    nickname: Optional[str] = Field(default=None, description="New nickname")
    company_name: Optional[str] = Field(default=None, description="New company name; '' removes it")
    email_addresses: Optional[List[str]] = Field(
        default=None,
        max_length=3,
        description=(
            "The contact's email addresses, replacing all the current ones: pass "
            "the full list, existing addresses included, to add one (Outlook keeps "
            "at most three). An empty list removes them all."
        ),
    )
    mobile_phone: Optional[str] = Field(default=None, description="New mobile phone number; '' removes it")
    home_phones: Optional[List[str]] = Field(
        default=None,
        max_length=2,
        description="Home phone numbers, replacing the current ones (at most two); [] removes them",
    )
    business_phones: Optional[List[str]] = Field(
        default=None,
        max_length=2,
        description="Business phone numbers, replacing the current ones (at most two); [] removes them",
    )
    birthday: Optional[str] = Field(
        default=None,
        description=(
            "Birthday as 'YYYY-MM-DD', e.g. '2016-06-07'. An empty string removes "
            "it, which also removes the contact's entry from the Birthdays calendar."
        ),
    )
    personal_notes: Optional[str] = Field(
        default=None,
        description="The notes on the contact, replacing the current ones; '' removes them",
    )

    @field_validator("birthday")
    @classmethod
    def validate_birthday(cls, v: Optional[str]) -> Optional[str]:
        if v is not None:
            birthday_value(v)
        return v


class DeleteContactInput(BaseModel):
    """Input for deleting a contact. It goes to Deleted Items, recoverable.

    There is deliberately no permanent option: on a personal mailbox, Graph's
    permanentDelete twice left a contact with a birthday in Deleted Items
    instead of purging it, once while answering 404.
    """
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    contact_id: str = Field(..., description="ID of the contact to delete", min_length=1)


class CreateContactInput(BaseModel):
    """Input for creating a contact."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    folder: str = Field(
        default="contacts",
        min_length=1,
        description=(
            "The contact folder to create it in: its name as "
            "outlook_list_contact_folders shows it (case and accents ignored), "
            "its ID, or 'contacts', the default, for the main Contacts folder."
        ),
    )
    display_name: Optional[str] = Field(
        default=None,
        description=(
            "The name the contact is listed and its birthday shown under. Omit "
            "to let Outlook compose it from the given name and surname."
        ),
    )
    given_name: Optional[str] = Field(default=None, description="Given (first) name")
    surname: Optional[str] = Field(default=None, description="Surname (last name)")
    nickname: Optional[str] = Field(default=None, description="Nickname")
    company_name: Optional[str] = Field(default=None, description="Company name")
    email_addresses: Optional[List[str]] = Field(
        default=None, max_length=3, description="Email addresses, at most three (Outlook's limit)"
    )
    mobile_phone: Optional[str] = Field(default=None, description="Mobile phone number")
    home_phones: Optional[List[str]] = Field(
        default=None, max_length=2, description="Home phone numbers, at most two"
    )
    business_phones: Optional[List[str]] = Field(
        default=None, max_length=2, description="Business phone numbers, at most two"
    )
    birthday: Optional[str] = Field(
        default=None,
        description=(
            "Birthday as 'YYYY-MM-DD', e.g. '2016-06-07'. It puts the contact in "
            "the Birthdays calendar."
        ),
    )
    personal_notes: Optional[str] = Field(default=None, description="Notes on the contact")

    @field_validator("birthday")
    @classmethod
    def validate_birthday(cls, v: Optional[str]) -> Optional[str]:
        if v:
            birthday_value(v)
        return v or None

    @model_validator(mode="after")
    def says_who(self) -> "CreateContactInput":
        if not any((
            self.display_name, self.given_name, self.surname, self.nickname,
            self.company_name, self.email_addresses, self.mobile_phone,
            self.home_phones, self.business_phones,
        )):
            raise ValueError("A contact needs at least a name, an email address or a phone number.")
        return self


# Contacts one outlook_move_contacts call takes. Each costs four sequential
# requests, so a full call answers in some twenty seconds.
MOVE_BATCH = 25


class MoveContactsInput(BaseModel):
    """Input for moving contacts into another contact folder."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    contact_ids: List[str] = Field(
        ...,
        min_length=1,
        max_length=MOVE_BATCH,
        description=(
            f"IDs of the contacts to move, as outlook_list_contacts shows them, at "
            f"most {MOVE_BATCH} per call. outlook_list_contacts with `folder` lists "
            f"the contacts of one folder."
        ),
    )
    destination_folder: str = Field(
        ...,
        min_length=1,
        description=(
            "The folder to move them into: its name as outlook_list_contact_folders "
            "shows it (case and accents ignored), its ID, or 'contacts' for the "
            "default folder."
        ),
    )

    @field_validator("contact_ids")
    @classmethod
    def validate_contact_ids(cls, v: List[str]) -> List[str]:
        if any(not contact_id for contact_id in v):
            raise ValueError("A contact ID cannot be empty.")
        return v


class DeleteContactFolderInput(BaseModel):
    """Input for deleting an empty contact folder. It goes to Deleted Items, recoverable."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    folder: str = Field(
        ...,
        min_length=1,
        description=(
            "The folder to delete: its name as outlook_list_contact_folders shows "
            "it (case and accents ignored), or its ID. It has to be empty."
        ),
    )


class DeleteAttachmentFilesInput(BaseModel):
    """Input for deleting the attachment files downloaded from one message."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    message_id: str = Field(
        ...,
        description="The message whose downloaded attachment files should be removed from the server's filesystem",
        min_length=1,
    )
    filename: Optional[str] = Field(
        default=None,
        description=(
            "Delete only this file, named as outlook_get_attachment reported it. "
            "Omit to delete everything downloaded from this message."
        ),
    )
