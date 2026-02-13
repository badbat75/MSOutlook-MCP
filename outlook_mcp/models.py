"""Pydantic input models for all MCP tools."""

from typing import Optional, List

from pydantic import BaseModel, Field, field_validator, ConfigDict


class ListMailInput(BaseModel):
    """Input for listing emails."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    folder: str = Field(
        default="inbox",
        description="Mail folder: 'inbox', 'sentitems', 'drafts', 'deleteditems', 'junkemail', or folder ID"
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
        description="Specific calendar ID. Omit for default calendar."
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
        description="Start datetime in ISO format, e.g. '2025-06-15T10:00:00'"
    )
    end: str = Field(
        ...,
        description="End datetime in ISO format, e.g. '2025-06-15T11:00:00'"
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
    is_all_day: bool = Field(default=False, description="All-day event")
    recurrence: Optional[str] = Field(
        default=None,
        description="Recurrence pattern: 'daily', 'weekly', 'monthly', or null for none"
    )
    calendar_id: Optional[str] = Field(default=None, description="Target calendar ID (omit for default)")


class UpdateEventInput(BaseModel):
    """Input for updating a calendar event."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    event_id: str = Field(..., description="ID of the event to update")
    subject: Optional[str] = Field(default=None, description="New subject")
    start: Optional[str] = Field(default=None, description="New start datetime (ISO format)")
    end: Optional[str] = Field(default=None, description="New end datetime (ISO format)")
    timezone: Optional[str] = Field(default=None, description="Timezone for start/end")
    location: Optional[str] = Field(default=None, description="New location")
    body: Optional[str] = Field(default=None, description="New body content")
    is_cancelled: bool = Field(default=False, description="Cancel the event")


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

    top: int = Field(default=20, description="Max folders to return", ge=1, le=50)


class ListCalendarsInput(BaseModel):
    """Input for listing calendars."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    top: int = Field(default=10, description="Max calendars to return", ge=1, le=50)
