"""Helpers for extracting and formatting Outlook item data."""
import re

from outlook_desktop_mcp.tools._folder_constants import (
    BUSY_STATUS_NAMES,
    MEETING_STATUS_NAMES,
    RESPONSE_NAMES,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
    FLAG_STATUS_NAMES,
)


def truncate(text: str, max_length: int = 2000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


def strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", "", html)
    text = re.sub(r"\s+", " ", text).strip()
    return text


# Field name -> reader, applied lazily on purpose: every property access
# crosses the COM process boundary, so a field that was not requested must not
# be read at all. Field selection is therefore a latency win, not only a
# payload win. Note has_attachments and attachment_count share a single
# Attachments.Count call instead of making two.
EMAIL_SUMMARY_FIELDS = (
    "entry_id",
    "subject",
    "sender",
    "sender_name",
    "received_time",
    "unread",
    "flag_status",
    "has_attachments",
    "attachment_count",
)


def format_email_summary(item, fields=None) -> dict:
    """Extract key fields from an Outlook MailItem into a dict.

    fields: optional iterable of field names to include. None (the default)
        returns every field, matching the historical behaviour.
    """
    wanted = EMAIL_SUMMARY_FIELDS if fields is None else frozenset(fields)
    out = {}

    if "entry_id" in wanted:
        out["entry_id"] = item.EntryID
    if "subject" in wanted:
        out["subject"] = item.Subject or "(no subject)"
    if "sender" in wanted:
        out["sender"] = getattr(item, "SenderEmailAddress", "unknown")
    if "sender_name" in wanted:
        out["sender_name"] = getattr(item, "SenderName", "unknown")
    if "received_time" in wanted:
        out["received_time"] = str(item.ReceivedTime)
    if "unread" in wanted:
        out["unread"] = bool(item.UnRead)
    if "flag_status" in wanted:
        out["flag_status"] = FLAG_STATUS_NAMES.get(getattr(item, "FlagStatus", 0), "none")
    if "has_attachments" in wanted or "attachment_count" in wanted:
        n = item.Attachments.Count
        if "has_attachments" in wanted:
            out["has_attachments"] = bool(n > 0)
        if "attachment_count" in wanted:
            out["attachment_count"] = n

    # Keep the declared order regardless of the order the caller asked in.
    return {k: out[k] for k in EMAIL_SUMMARY_FIELDS if k in out}


def parse_summary_fields(spec: str, valid=EMAIL_SUMMARY_FIELDS):
    """Turn a comma-separated field spec into a validated tuple.

    Returns (fields, error). An empty spec yields (None, None), i.e. all
    fields. An unknown name is a hard error rather than a silent drop, so a
    typo cannot quietly remove data the caller expected.

    `valid` selects the vocabulary: mail summaries by default, or
    DRAFT_SUMMARY_FIELDS for the Drafts listing, which has its own columns.
    """
    if not spec or not spec.strip():
        return None, None
    names = [p.strip() for p in spec.split(",") if p.strip()]
    if not names:
        return None, None
    unknown = [n for n in names if n not in valid]
    if unknown:
        return None, (
            f"Unknown field(s): {', '.join(unknown)}. "
            f"Valid: {', '.join(valid)}"
        )
    return tuple(dict.fromkeys(names)), None


# Drafts list from their own columns: there is no sender (it is you) and no
# ReceivedTime (an unsent item was never received), while the recipients and a
# body preview are what a caller needs in order to pick one.
DRAFT_SUMMARY_FIELDS = (
    "entry_id",
    "subject",
    "to",
    "cc",
    "bcc",
    "last_modified",
    "has_attachments",
    "attachment_count",
    "body_preview",
)

DRAFT_BODY_PREVIEW_CHARS = 200


def format_draft_summary(item, fields=None) -> dict:
    """Extract key fields from a draft MailItem into a dict.

    Same lazy-read contract as format_email_summary: a field that was not
    requested is never read. That matters most for body_preview, which pulls
    the entire message body across the COM boundary just to keep its first
    200 characters.
    """
    wanted = DRAFT_SUMMARY_FIELDS if fields is None else frozenset(fields)
    out = {}

    if "entry_id" in wanted:
        out["entry_id"] = item.EntryID
    if "subject" in wanted:
        out["subject"] = item.Subject or "(no subject)"
    if "to" in wanted:
        out["to"] = item.To or ""
    if "cc" in wanted:
        out["cc"] = item.CC or ""
    if "bcc" in wanted:
        out["bcc"] = item.BCC or ""
    if "last_modified" in wanted:
        out["last_modified"] = str(item.LastModificationTime)
    if "has_attachments" in wanted or "attachment_count" in wanted:
        n = item.Attachments.Count
        if "has_attachments" in wanted:
            out["has_attachments"] = bool(n > 0)
        if "attachment_count" in wanted:
            out["attachment_count"] = n
    if "body_preview" in wanted:
        body = item.Body or ""
        preview = body[:DRAFT_BODY_PREVIEW_CHARS]
        if len(body) > DRAFT_BODY_PREVIEW_CHARS:
            preview += "..."
        out["body_preview"] = preview

    # Keep the declared order regardless of the order the caller asked in.
    return {k: out[k] for k in DRAFT_SUMMARY_FIELDS if k in out}


def format_email_full(item, body_max_length: int = 5000) -> dict:
    """Extract full email details including body."""
    result = format_email_summary(item)
    result["to"] = item.To or ""
    result["cc"] = item.CC or ""
    result["body"] = truncate(item.Body or "", body_max_length)
    return result


# --- Calendar formatting ---


def format_event_summary(item) -> dict:
    """Extract key fields from an Outlook AppointmentItem."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "start": str(item.Start),
        "end": str(item.End),
        "duration": item.Duration,
        "location": item.Location or "",
        "organizer": item.Organizer or "",
        "is_recurring": bool(item.IsRecurring),
        "all_day": bool(item.AllDayEvent),
        "busy_status": BUSY_STATUS_NAMES.get(item.BusyStatus, "unknown"),
        "meeting_status": MEETING_STATUS_NAMES.get(item.MeetingStatus, "unknown"),
        "required_attendees": item.RequiredAttendees or "",
        "optional_attendees": item.OptionalAttendees or "",
        "categories": item.Categories or "",
    }


def format_event_full(item, body_max_length: int = 5000) -> dict:
    """Full event details including body."""
    result = format_event_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["reminder_minutes"] = (
        item.ReminderMinutesBeforeStart if item.ReminderSet else None
    )
    result["response_status"] = RESPONSE_NAMES.get(item.ResponseStatus, "unknown")
    return result


# --- Task formatting ---


def format_task_summary(item) -> dict:
    """Extract key fields from an Outlook TaskItem."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "status": TASK_STATUS_NAMES.get(item.Status, "unknown"),
        "percent_complete": item.PercentComplete,
        "due_date": str(item.DueDate) if str(item.DueDate) != "01/01/4501" else None,
        "start_date": str(item.StartDate) if str(item.StartDate) != "01/01/4501" else None,
        "importance": IMPORTANCE_NAMES.get(item.Importance, "normal"),
        "complete": bool(item.Complete),
        "categories": item.Categories or "",
        "owner": item.Owner or "",
    }


def format_task_full(item, body_max_length: int = 5000) -> dict:
    """Full task details including body."""
    result = format_task_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["date_completed"] = (
        str(item.DateCompleted) if item.Complete else None
    )
    return result
