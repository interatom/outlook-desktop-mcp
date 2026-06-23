"""
Outlook Desktop MCP Server
===========================
Exposes Microsoft Outlook Desktop (Classic) as an MCP server over stdio.
Uses COM automation — no Microsoft Graph, no Entra app registration.
Just run this on Windows with Outlook open and you have a full email MCP server.

Entry point: python -m outlook_desktop_mcp.server
"""
import sys
import json
import logging
import re

from mcp.server.fastmcp import FastMCP

from outlook_desktop_mcp.com_bridge import OutlookBridge
from datetime import datetime, timedelta
import locale as _locale

import os

from outlook_desktop_mcp.tools._folder_constants import (
    FOLDER_NAME_TO_ENUM,
    OL_MAIL_ITEM,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_DRAFTS,
    OL_FOLDER_TASKS,
    OL_MEETING,
    OL_MEETING_CANCELED,
    OL_RESPONSE_TENTATIVE,
    OL_RESPONSE_ACCEPTED,
    OL_RESPONSE_DECLINED,
    OL_REQUIRED,
    OL_OPTIONAL,
    OL_TASK_ITEM,
    OL_TASK_COMPLETE,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
    OL_FLAG_MARKED,
)
from outlook_desktop_mcp.utils.formatting import (
    format_email_summary,
    format_email_full,
    format_event_summary,
    format_event_full,
    format_task_summary,
    format_task_full,
)
from outlook_desktop_mcp.utils.errors import format_com_error

# --- Logging (all to stderr, stdout is reserved for MCP JSON-RPC) ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("outlook_desktop_mcp")


# --- Security helpers ---

def _safe_dasl(query: str) -> str:
    """Sanitize a string for use in a DASL LIKE filter value.
    Escapes SQL wildcards (% and _) so user input is treated as literals,
    then escapes quote characters required by DASL syntax.
    """
    query = query.replace("%", "[%]").replace("_", "[_]")
    return query.replace("'", "''").replace('"', '""')


# Outlook item Class constants (olObjectClass — distinct from olItemType used in CreateItem)
_OL_CLASS_MAIL = 43
_OL_CLASS_APPOINTMENT = 26
_OL_CLASS_TASK = 48


def _check_item_class(item, expected_class: int, label: str) -> str | None:
    """Return an error string if item is the wrong type, else None."""
    if item.Class != expected_class:
        return f"Error: Entry ID does not refer to a {label}."
    return None


# --- MCP Server ---

mcp = FastMCP(
    "outlook-desktop-mcp",
    instructions=(
        "This MCP server gives you full access to Microsoft Outlook Desktop on "
        "Windows via COM automation. It can send emails, read inbox messages, "
        "search across folders, mark messages as read/unread, move messages "
        "between folders (including archive), reply to emails, and list the "
        "complete folder hierarchy.\n\n"
        "All operations use the locally authenticated Outlook profile — no "
        "Microsoft Graph API, no Entra app registration, no OAuth tokens needed. "
        "The user's existing Outlook session handles all authentication.\n\n"
        "PREREQUISITE: Outlook Desktop (Classic) must be running. The new/modern "
        "Outlook (olk.exe) is NOT supported — only the classic OUTLOOK.EXE.\n\n"
        "AVAILABLE TOOL CATEGORIES:\n"
        "- Email: send, list, read, search, reply, mark read/unread, move, attachments\n"
        "- Calendar: list events, create appointments/meetings, update, delete, "
        "respond to invites, search events\n"
        "- Tasks: create, list, complete, update, delete to-do items\n"
        "- Categories: list and set color categories on any item\n"
        "- Rules: list and manage mail rules\n"
        "- Out of Office: check auto-reply status\n"
        "- Folders: list folder hierarchy with item counts"
    ),
)

bridge = OutlookBridge()


# --- Helper: resolve store by account name ---

def _resolve_store(namespace, account: str = ""):
    """Resolve an account name to an Outlook Store object.

    If account is empty, returns DefaultStore.
    Otherwise does a case-insensitive substring match on Store.DisplayName.
    """
    if not account:
        return namespace.DefaultStore

    account_lower = account.lower().strip()
    for i in range(namespace.Stores.Count):
        store = namespace.Stores.Item(i + 1)
        if account_lower in store.DisplayName.lower():
            return store

    return None


def _require_store(namespace, account: str = ""):
    """Resolve store, raising ValueError if not found."""
    store = _resolve_store(namespace, account)
    if store is None:
        raise ValueError(f"Account '{account}' not found. Use list_accounts to see available accounts.")
    return store


# --- Helper: resolve folder by name ---

def _walk_folders(parent, name_lower: str):
    """Recursively search subfolders of parent for a folder matching name_lower."""
    try:
        # Some Exchange/shared-mailbox folder types raise COMError on .Count;
        # treat them as empty to avoid crashing the traversal.
        count = parent.Folders.Count
    except Exception:
        return None
    for i in range(count):
        try:
            f = parent.Folders.Item(i + 1)
            if f.Name.lower() == name_lower:
                return f
            found = _walk_folders(f, name_lower)
            if found:
                return found
        except Exception:
            continue
    return None


def _resolve_folder(namespace, folder_name: str, store=None):
    """Resolve a folder name to an Outlook MAPIFolder object.

    Resolution order:
    1. Slash-delimited path (e.g. "Inbox/Receipts") — traverse segment by segment
    2. Built-in Outlook folder enum (inbox, sent, deleted, etc.)
    3. Root-level folder name match (fast path)
    4. Recursive depth-first search of entire folder tree (fallback)
    5. Search folders (virtual/search folders not in regular tree)
    """
    folder_name = folder_name.strip()
    store = store or namespace.DefaultStore

    # Slash-delimited path: traverse segment by segment
    if "/" in folder_name:
        parts = [p.strip() for p in folder_name.split("/")]
        current = _resolve_folder(namespace, parts[0], store)
        if current is None:
            return None
        for part in parts[1:]:
            part_lower = part.lower()
            found = None
            try:
                count = current.Folders.Count
            except Exception:
                return None
            for i in range(count):
                try:
                    f = current.Folders.Item(i + 1)
                    if f.Name.lower() == part_lower:
                        found = f
                        break
                except Exception:
                    continue
            if found is None:
                return None
            current = found
        return current

    folder_lower = folder_name.lower()

    # Built-in Outlook folders
    if folder_lower in FOLDER_NAME_TO_ENUM:
        return store.GetDefaultFolder(FOLDER_NAME_TO_ENUM[folder_lower])

    # Root-level search (fast path)
    root = store.GetRootFolder()
    for i in range(root.Folders.Count):
        try:
            f = root.Folders.Item(i + 1)
            if f.Name.lower() == folder_lower:
                return f
        except Exception:
            continue

    # Recursive fallback: search entire folder tree
    result = _walk_folders(root, folder_lower)
    if result:
        return result

    # Search folders fallback: virtual folders (e.g. "Flagged", "To-Do") are
    # not in the regular folder tree — access via Store.GetSearchFolders()
    try:
        search_folders = store.GetSearchFolders()
        for i in range(search_folders.Count):
            try:
                f = search_folders.Item(i + 1)
                if f.Name.lower() == folder_lower:
                    return f
            except Exception:
                continue
    except Exception:
        pass

    return None


# =====================================================================
# TOOL: list_accounts
# =====================================================================

@mcp.tool()
async def list_accounts() -> str:
    """List all Outlook accounts (stores) configured in the profile.

    Returns a JSON array of account objects with display_name, store_id,
    and is_default. Use the display_name (or a unique substring) as the
    'account' parameter in other tools to target a specific account.

    Returns:
        JSON array of account objects.
    """
    def _list(outlook, namespace):
        default_id = namespace.DefaultStore.StoreID
        results = []
        for i in range(namespace.Stores.Count):
            store = namespace.Stores.Item(i + 1)
            results.append({
                "display_name": store.DisplayName,
                "store_id": store.StoreID,
                "is_default": store.StoreID == default_id,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list)
    except Exception as e:
        return f"Error listing accounts: {format_com_error(e)}"


# =====================================================================
# TOOL 1: send_email
# =====================================================================

@mcp.tool()
async def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    account: str = "",
) -> str:
    """Send an email using the user's Outlook account.

    Creates and sends an email immediately through the default Outlook profile.
    The email will appear in the user's Sent Items folder after sending.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.
        account: Optional. Account display name (or substring) to send from.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        A confirmation message with subject and recipients, or an error.
    """
    def _send(outlook, namespace, to, subject, body, cc, bcc, html_body, account):
        store = _require_store(namespace, account)
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        # Set the sending account
        for acc in outlook.Session.Accounts:
            if acc.DeliveryStore.StoreID == store.StoreID:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))  # SendUsingAccount
                break
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        if html_body:
            mail.HTMLBody = html_body
        mail.Send()
        return f"Email sent: '{subject}' to {to}"

    try:
        return await bridge.call(_send, to, subject, body, cc, bcc, html_body, account)
    except Exception as e:
        return f"Error sending email: {format_com_error(e)}"


# =====================================================================
# TOOL 2: list_emails
# =====================================================================

@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
) -> str:
    """List recent emails from a specified Outlook folder.

    Returns a JSON array of email summaries sorted by received time (newest
    first). Each summary includes entry_id, subject, sender, sender_name,
    received_time, unread status, and attachment info.

    Use the entry_id from results to read full content with read_email,
    or to perform actions like mark_as_read, move_email, or reply_email.

    Args:
        folder: The folder to list. Case-insensitive names: "inbox" (default),
            "sent"/"sentmail", "drafts", "deleted"/"trash", "junk"/"spam",
            "outbox", "archive", or any custom folder name visible in
            list_folders output.
        count: Maximum number of emails to return. Default 10, max recommended 50.
        unread_only: If true, only return unread emails. Default false.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of email summary objects.
    """
    def _list(outlook, namespace, folder, count, unread_only, start_date, end_date, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        items = target.Items
        items.Sort("[ReceivedTime]", True)

        # Build restriction filters
        restrictions = []
        if unread_only:
            restrictions.append("[UnRead] = True")
        if start_date:
            start = _parse_date(start_date)
            restrictions.append(f"[ReceivedTime] >= '{start.strftime('%m/%d/%Y %H:%M')}'")
        if end_date:
            end = _parse_date(end_date)
            restrictions.append(f"[ReceivedTime] <= '{end.strftime('%m/%d/%Y %H:%M')}'")
        elif start_date:
            # Default end to now when start is specified
            restrictions.append(f"[ReceivedTime] <= '{datetime.now().strftime('%m/%d/%Y %H:%M')}'")

        if restrictions:
            items = items.Restrict(" AND ".join(restrictions))

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_email_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, count, unread_only, start_date, end_date, account)
    except Exception as e:
        return f"Error listing emails: {format_com_error(e)}"


# =====================================================================
# TOOL 3: read_email
# =====================================================================

@mcp.tool()
async def read_email(
    entry_id: str = "",
    subject_search: str = "",
    folder: str = "inbox",
    account: str = "",
) -> str:
    """Read the full content of a specific email.

    Retrieves complete email details including body text, recipients, CC,
    and metadata. Provide EITHER entry_id (preferred, exact match) OR
    subject_search (finds most recent match by subject substring).

    Args:
        entry_id: The unique Outlook EntryID of the email. Most reliable way
            to identify a specific email. Get this from list_emails or
            search_emails results.
        subject_search: Alternative to entry_id. A case-insensitive substring
            to search for in email subjects. Returns the most recent match.
        folder: Folder to search when using subject_search. Ignored when
            entry_id is provided. Default "inbox".
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with full email details (entry_id, subject, sender,
        sender_name, received_time, unread, to, cc, body, attachment info).
    """
    def _read(outlook, namespace, entry_id, subject_search, folder, account):
        if entry_id:
            item = namespace.GetItemFromID(entry_id)
            return json.dumps(format_email_full(item), indent=2, default=str)

        if not subject_search:
            return json.dumps({"error": "Provide either entry_id or subject_search"})

        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        safe_query = _safe_dasl(subject_search)
        filter_str = (
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%'"
        )
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)
        if items.Count == 0:
            return json.dumps({"error": f"No email found matching '{subject_search}'"})

        return json.dumps(format_email_full(items.Item(1)), indent=2, default=str)

    try:
        return await bridge.call(_read, entry_id, subject_search, folder, account)
    except Exception as e:
        return f"Error reading email: {format_com_error(e)}"


# =====================================================================
# TOOL 4: mark_as_read
# =====================================================================

@mcp.tool()
async def mark_as_read(entry_id: str, account: str = "") -> str:
    """Mark a specific email as read in Outlook.

    Changes the unread status to read, same as clicking on an email in Outlook.
    The change is persisted immediately and synced to the server.

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        item.UnRead = False
        item.Save()
        return f"Marked as read: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id, account)
    except Exception as e:
        return f"Error marking email as read: {format_com_error(e)}"


# =====================================================================
# TOOL 5: mark_as_unread
# =====================================================================

@mcp.tool()
async def mark_as_unread(entry_id: str, account: str = "") -> str:
    """Mark a specific email as unread in Outlook.

    Restores a previously read email to unread status. Useful for flagging
    emails that need follow-up attention. Persisted immediately.

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        item.UnRead = True
        item.Save()
        return f"Marked as unread: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id, account)
    except Exception as e:
        return f"Error marking email as unread: {format_com_error(e)}"


# =====================================================================
# TOOL 6: set_flag / clear_flag
# =====================================================================

@mcp.tool()
async def set_flag(entry_id: str, account: str = "") -> str:
    """Flag an email for follow-up in Outlook.

    Sets the follow-up flag on the specified email (equivalent to clicking
    the flag icon in Outlook). Use clear_flag to remove it.

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _set(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        item.FlagStatus = OL_FLAG_MARKED
        item.Save()
        return f"Flagged: '{subject}'"

    try:
        return await bridge.call(_set, entry_id, account)
    except Exception as e:
        return f"Error setting flag: {format_com_error(e)}"


@mcp.tool()
async def clear_flag(entry_id: str, account: str = "") -> str:
    """Remove the follow-up flag from an email in Outlook.

    Clears the flag set on the specified email (equivalent to right-clicking
    the flag and selecting 'Clear Flag' in Outlook).

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _clear(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        # ClearTaskFlag() removes all follow-up flag properties in one call
        # (FlagStatus, FlagRequest, TaskDueDate, TaskStartDate).
        item.ClearTaskFlag()
        item.Save()
        return f"Flag cleared: '{subject}'"

    try:
        return await bridge.call(_clear, entry_id, account)
    except Exception as e:
        return f"Error clearing flag: {format_com_error(e)}"


# =====================================================================
# TOOL 7: move_email
# =====================================================================

@mcp.tool()
async def move_email(
    entry_id: str,
    target_folder: str = "archive",
    account: str = "",
) -> str:
    """Move an email to a different Outlook folder.

    Moves the specified email from its current location to the target folder.
    IMPORTANT: After moving, the email gets a NEW entry_id — the old one
    becomes invalid. Common use: archiving emails after processing.

    Args:
        entry_id: The unique Outlook EntryID of the email to move.
        target_folder: Destination folder name. Default is "archive". Supports
            same names as list_emails: "archive", "inbox", "sent", "deleted"/
            "trash", "drafts", "junk"/"spam", or any custom folder name.
        account: Optional. Account display name (or substring) to resolve
            the target folder in. Default: primary account.

    Returns:
        Confirmation with email subject and destination, or an error.
    """
    def _move(outlook, namespace, entry_id, target_folder, account):
        item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject

        store = _require_store(namespace, account)
        dest = _resolve_folder(namespace, target_folder, store)
        if not dest:
            return f"Error: Target folder '{target_folder}' not found. Use list_folders to see available folders."

        item.Move(dest)
        return f"Moved '{subject}' to {target_folder}"

    try:
        return await bridge.call(_move, entry_id, target_folder, account)
    except Exception as e:
        return f"Error moving email: {format_com_error(e)}"


# =====================================================================
# TOOL 7: reply_email
# =====================================================================

@mcp.tool()
async def reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    save_as_draft: bool = False,
    account: str = "",
) -> str:
    """Reply to an email in Outlook.

    Creates and sends a reply, preserving the original message thread.
    Use reply_all=True to reply to all recipients (sender + CC list).
    Use save_as_draft=True to save the reply as a draft instead of sending.

    Args:
        entry_id: The unique Outlook EntryID of the email to reply to.
        body: The reply message text. Prepended above the original message
            in the email thread.
        reply_all: If true, reply to all recipients (sender + all CC/To).
            If false (default), reply only to the sender.
        save_as_draft: If true, save the reply to Drafts instead of sending.
            Default false (sends immediately).
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation the reply was sent, or JSON with entry_id if saved as draft.
    """
    def _reply(outlook, namespace, entry_id, body, reply_all, save_as_draft, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        reply_item = item.ReplyAll() if reply_all else item.Reply()
        reply_item.Body = body + "\n\n" + reply_item.Body
        if save_as_draft:
            reply_item.Save()
            return json.dumps({"entry_id": reply_item.EntryID, "subject": reply_item.Subject})
        reply_item.Send()
        return f"Reply sent to '{subject}' (reply_all={reply_all})"

    try:
        return await bridge.call(_reply, entry_id, body, reply_all, save_as_draft, account)
    except Exception as e:
        return f"Error replying to email: {format_com_error(e)}"


# =====================================================================
# TOOL 8: list_folders
# =====================================================================

def _folder_recency(folder):
    """Return (last_received, oldest_received) datetimes for a mail folder, or
    (None, None) for an empty folder or a non-mail folder (calendar/contacts/...).

    Sorts the folder's items by ReceivedTime once (descending) and reads both
    ends — cheap even for thousands of items (Outlook sorts natively).
    """
    try:
        if folder.DefaultItemType != OL_MAIL_ITEM:
            return None, None
    except Exception:
        return None, None
    try:
        items = folder.Items
        if items.Count == 0:
            return None, None
        items.Sort("[ReceivedTime]", True)  # descending: GetFirst=newest, GetLast=oldest
        return (getattr(items.GetFirst(), "ReceivedTime", None),
                getattr(items.GetLast(), "ReceivedTime", None))
    except Exception:
        return None, None


@mcp.tool()
async def list_folders(folder: str = "", max_depth: int = 3,
                       include_dates: bool = False, account: str = "") -> str:
    """List mail folders in the user's Outlook mailbox.

    When called with no folder argument, lists top-level folders. Provide a
    folder name to drill into its subfolders — use this to browse the full
    folder tree step by step (e.g. first call with no folder to see top-level,
    then call with folder="Inbox" to see Inbox children, then
    folder="Inbox/Projects" to go deeper).

    Folder names from this output can be used directly in list_emails,
    move_email, search_emails, etc. Use slash-delimited paths for nested
    folders (e.g. "Inbox/Receipts/2026").

    Args:
        folder: Optional. Folder to list children of. Supports folder names
            ("Inbox"), slash paths ("Inbox/Receipts"), or built-in names
            ("sent", "drafts"). When empty, lists from the mailbox root.
        max_depth: How many levels deep to recurse below the starting folder.
            Default 3. Set to 1 to see only immediate children.
        include_dates: Optional. When True, each mail folder also gets
            `last_received` and `oldest_received` (newest/oldest message
            ReceivedTime, null for empty or non-mail folders). One recursive
            call then answers "which folders have gone quiet?" without a
            per-folder loop. Off by default — it sorts each folder's items, so
            only request it when you need the recency signal.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of folder objects with name, full_path, item_count,
        unread_count, subfolders (if any), and — when include_dates=True —
        last_received / oldest_received.
    """
    def _list(outlook, namespace, folder, max_depth, include_dates, account):
        max_depth = min(max(1, max_depth), 10)
        store = _require_store(namespace, account)

        if folder:
            start = _resolve_folder(namespace, folder, store)
            if not start:
                return json.dumps({"error": f"Folder '{folder}' not found"})
            base_path = folder
        else:
            start = store.GetRootFolder()
            base_path = ""

        def walk(f, depth, path_prefix):
            current_path = f"{path_prefix}/{f.Name}" if path_prefix else f.Name
            result = {
                "name": f.Name,
                "full_path": current_path,
                "item_count": f.Items.Count,
                "unread_count": f.UnReadItemCount,
            }
            if include_dates:
                last, oldest = _folder_recency(f)
                result["last_received"] = last
                result["oldest_received"] = oldest
            if depth < max_depth:
                children = []
                for i in range(f.Folders.Count):
                    try:
                        child = f.Folders.Item(i + 1)
                        children.append(walk(child, depth + 1, current_path))
                    except Exception:
                        continue
                if children:
                    result["subfolders"] = children
            return result

        folders = []
        for i in range(start.Folders.Count):
            try:
                child = start.Folders.Item(i + 1)
                folders.append(walk(child, 1, base_path))
            except Exception:
                continue
        return json.dumps(folders, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, max_depth, include_dates, account)
    except Exception as e:
        return f"Error listing folders: {format_com_error(e)}"


# =====================================================================
# TOOL 9: search_emails
# =====================================================================

@mcp.tool()
async def search_emails(
    query: str,
    folder: str = "inbox",
    count: int = 10,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
) -> str:
    """Search for emails in Outlook using text search.

    Searches email subjects and bodies using Outlook's DASL filter.
    Results are sorted by received time (newest first). Each result
    includes entry_id for further operations.

    Args:
        query: The search term (case-insensitive substring match).
            Examples: "budget report", "meeting notes", "quarterly".
        folder: Folder to search in. Default "inbox". Supports same
            names as list_emails.
        count: Maximum results to return. Default 10.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of matching email summaries, or an error.
    """
    def _search(outlook, namespace, query, folder, count, start_date, end_date, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        safe_query = _safe_dasl(query)
        dasl_parts = [
            f"(\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%' OR "
            f"\"urn:schemas:httpmail:textdescription\" LIKE '%{safe_query}%')"
        ]
        if start_date:
            start = _parse_date(start_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" >= '{start.strftime('%m/%d/%Y %H:%M')}'"
            )
        if end_date:
            end = _parse_date(end_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= '{end.strftime('%m/%d/%Y %H:%M')}'"
            )
        elif start_date:
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= '{datetime.now().strftime('%m/%d/%Y %H:%M')}'"
            )

        filter_str = "@SQL=" + " AND ".join(dasl_parts)
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_email_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_search, query, folder, count, start_date, end_date, account)
    except Exception as e:
        return f"Error searching emails: {format_com_error(e)}"


# =====================================================================
# CALENDAR TOOLS
# =====================================================================


# --- Helper: parse ISO date string ---

def _parse_date(date_str: str) -> datetime:
    """Parse ISO 8601 date string like '2026-02-25 14:00' or '2026-02-25T14:00:00'."""
    return datetime.fromisoformat(date_str)


def _date_to_restrict_str(dt: datetime) -> str:
    """Format datetime for Outlook Restrict() using the system locale's date order.

    Outlook's COM Restrict() parses date strings using the Windows system locale.
    Using %x (locale's preferred short date) ensures day/month order matches,
    avoiding the MM/DD vs DD/MM swap on non-US locales (e.g. de-DE).
    """
    saved = _locale.getlocale(_locale.LC_TIME)
    try:
        _locale.setlocale(_locale.LC_TIME, "")
        return dt.strftime("%x %H:%M")
    finally:
        _locale.setlocale(_locale.LC_TIME, saved)


# =====================================================================
# TOOL 10: list_events
# =====================================================================

@mcp.tool()
async def list_events(
    start_date: str = "",
    end_date: str = "",
    count: int = 20,
    account: str = "",
) -> str:
    """List upcoming calendar events from Outlook.

    Returns a JSON array of event summaries within a date range, sorted by
    start time. Includes recurring event occurrences. Each summary has
    entry_id, subject, start, end, duration, location, organizer, attendees,
    and status info.

    Use entry_id from results with get_event, update_event, delete_event,
    or respond_to_meeting.

    Args:
        start_date: Start of date range in ISO 8601 format (e.g. "2026-02-25"
            or "2026-02-25 09:00"). Default: now.
        end_date: End of date range. Default: 7 days from start_date.
        count: Maximum number of events to return. Default 20.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of event summary objects.
    """
    def _list(outlook, namespace, start_date, end_date, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        calendar = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
        items = calendar.Items

        # CRITICAL ORDER: Sort BEFORE IncludeRecurrences BEFORE Restrict
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now()
        end = _parse_date(end_date) if end_date else start + timedelta(days=7)

        restrict = (
            f"[Start] >= '{_date_to_restrict_str(start)}' "
            f"AND [Start] <= '{_date_to_restrict_str(end)}'"
        )
        filtered = items.Restrict(restrict)

        results = []
        n = 0
        for item in filtered:
            n += 1
            try:
                results.append(format_event_summary(item))
            except Exception:
                continue
            if n >= count:
                break

        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, start_date, end_date, count, account)
    except Exception as e:
        return f"Error listing events: {format_com_error(e)}"


# =====================================================================
# TOOL 11: get_event
# =====================================================================

@mcp.tool()
async def get_event(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific calendar event.

    Retrieves complete event information including body/description,
    attendees, recurrence status, reminders, and response status.

    Args:
        entry_id: The unique Outlook EntryID of the event. Get this from
            list_events or search_events results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full event details.
    """
    def _get(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        return json.dumps(format_event_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return f"Error reading event: {format_com_error(e)}"


# =====================================================================
# TOOL 12: create_event
# =====================================================================

@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    reminder_minutes: int = 15,
    account: str = "",
) -> str:
    """Create a personal calendar appointment (no attendees).

    Creates and saves an appointment on the user's calendar. This is a
    personal event — no meeting invitations are sent. Use create_meeting
    instead if you need to invite attendees.

    Args:
        subject: The event title.
        start: Start time in ISO 8601 format. Examples: "2026-02-25 14:00",
            "2026-02-25T14:00:00". For all-day events, use just the date:
            "2026-02-25".
        end: End time in ISO 8601 format. For all-day events, use the next
            day: "2026-02-26".
        location: Optional. Event location (e.g. "Conference Room A",
            "Microsoft Teams Meeting").
        body: Optional. Description or notes for the event.
        all_day: If true, creates an all-day event. Default false.
        reminder_minutes: Minutes before the event to show a reminder.
            Default 15. Set to 0 to disable reminder.
        account: Optional. Account display name (or substring) to create
            the event in. Default: primary account.

    Returns:
        Confirmation with event subject and entry_id, or an error.
    """
    def _create(outlook, namespace, subject, start, end, location, body,
                all_day, reminder_minutes, account):
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        # Move to correct store's calendar if account specified
        if account:
            store = _require_store(namespace, account)
            cal = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
            appt.Move(cal)
            appt = namespace.GetItemFromID(appt.EntryID)
        appt.Subject = subject
        appt.Start = start
        appt.End = end
        if location:
            appt.Location = location
        if body:
            appt.Body = body
        appt.AllDayEvent = all_day
        if reminder_minutes > 0:
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = reminder_minutes
        else:
            appt.ReminderSet = False
        appt.Save()
        return json.dumps({
            "status": "created",
            "subject": appt.Subject,
            "start": str(appt.Start),
            "end": str(appt.End),
            "entry_id": appt.EntryID,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, start, end, location, body, all_day,
            reminder_minutes, account,
        )
    except Exception as e:
        return f"Error creating event: {format_com_error(e)}"


# =====================================================================
# TOOL 13: create_meeting
# =====================================================================

@mcp.tool()
async def create_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
    account: str = "",
) -> str:
    """Create a meeting and send invitations to attendees.

    Creates a calendar meeting and immediately sends meeting requests to
    all specified attendees. The meeting will appear on the organizer's
    calendar and attendees will receive an invitation they can accept,
    decline, or tentatively accept.

    Args:
        subject: The meeting title.
        start: Start time in ISO 8601 format (e.g. "2026-02-25 14:00").
        end: End time in ISO 8601 format (e.g. "2026-02-25 15:00").
        required_attendees: Required attendee email addresses, separated by
            semicolons. Example: "alice@example.com; bob@example.com"
        location: Optional. Meeting location (e.g. "Teams", "Room 301").
        body: Optional. Meeting description or agenda.
        optional_attendees: Optional. Optional attendee emails, separated
            by semicolons.
        account: Optional. Account display name (or substring) to send from.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        Confirmation that the meeting was created and invitations sent.
    """
    def _create(outlook, namespace, subject, start, end, required_attendees,
                location, body, optional_attendees, account):
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        # Set sending account
        if account:
            store = _require_store(namespace, account)
            for acc in outlook.Session.Accounts:
                if acc.DeliveryStore.StoreID == store.StoreID:
                    appt._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
                    break
        appt.Subject = subject
        appt.Start = start
        appt.End = end
        appt.MeetingStatus = OL_MEETING
        if location:
            appt.Location = location
        if body:
            appt.Body = body

        for addr in required_attendees.split(";"):
            addr = addr.strip()
            if addr:
                recip = appt.Recipients.Add(addr)
                recip.Type = OL_REQUIRED

        if optional_attendees:
            for addr in optional_attendees.split(";"):
                addr = addr.strip()
                if addr:
                    recip = appt.Recipients.Add(addr)
                    recip.Type = OL_OPTIONAL

        appt.Recipients.ResolveAll()
        appt.Send()
        return (
            f"Meeting '{subject}' created and invitations sent to "
            f"{required_attendees}"
        )

    try:
        return await bridge.call(
            _create, subject, start, end, required_attendees, location, body,
            optional_attendees, account,
        )
    except Exception as e:
        return f"Error creating meeting: {format_com_error(e)}"


# =====================================================================
# TOOL 14: update_event
# =====================================================================

@mcp.tool()
async def update_event(
    entry_id: str,
    subject: str = "",
    start: str = "",
    end: str = "",
    location: str = "",
    body: str = "",
    account: str = "",
) -> str:
    """Update an existing calendar event.

    Modifies properties of an appointment or meeting. Only the fields you
    provide will be updated — omitted fields remain unchanged. For meetings
    you organize, attendees will receive an update notification.

    Args:
        entry_id: The unique Outlook EntryID of the event to update.
        subject: Optional. New event title.
        start: Optional. New start time in ISO 8601 format.
        end: Optional. New end time in ISO 8601 format.
        location: Optional. New location.
        body: Optional. New description/notes.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with updated event details, or an error.
    """
    def _update(outlook, namespace, entry_id, subject, start, end, location, body, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        if subject:
            item.Subject = subject
        if start:
            item.Start = start
        if end:
            item.End = end
        if location:
            item.Location = location
        if body:
            item.Body = body
        item.Save()
        return json.dumps({
            "status": "updated",
            "subject": item.Subject,
            "start": str(item.Start),
            "end": str(item.End),
            "location": item.Location or "",
            "entry_id": item.EntryID,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _update, entry_id, subject, start, end, location, body, account,
        )
    except Exception as e:
        return f"Error updating event: {format_com_error(e)}"


# =====================================================================
# TOOL 15: delete_event
# =====================================================================

@mcp.tool()
async def delete_event(entry_id: str, account: str = "") -> str:
    """Delete a calendar event or cancel a meeting.

    For personal appointments, the event is simply deleted. For meetings
    you organized, this cancels the meeting and sends cancellation notices
    to all attendees. For meetings you received, this declines and removes
    the event from your calendar.

    Args:
        entry_id: The unique Outlook EntryID of the event to delete/cancel.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the event subject, or an error.
    """
    def _delete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        subject = item.Subject
        meeting_status = item.MeetingStatus

        # If this is a meeting we organized, cancel it (sends notices)
        if meeting_status == OL_MEETING:
            item.MeetingStatus = OL_MEETING_CANCELED
            item.Send()
            return f"Meeting canceled: '{subject}' (cancellation sent to attendees)"

        # Otherwise just delete
        item.Delete()
        return f"Event deleted: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id, account)
    except Exception as e:
        return f"Error deleting event: {format_com_error(e)}"


# =====================================================================
# TOOL 16: respond_to_meeting
# =====================================================================

@mcp.tool()
async def respond_to_meeting(
    entry_id: str,
    response: str,
    account: str = "",
) -> str:
    """Respond to a meeting invitation (accept, decline, or tentative).

    Sends your response to the meeting organizer. The meeting will be
    added to (or updated on) your calendar accordingly.

    Args:
        entry_id: The unique Outlook EntryID of the meeting to respond to.
            Get this from list_events or search_events.
        response: Your response. Must be one of: "accept", "decline",
            or "tentative".
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation of your response, or an error.
    """
    def _respond(outlook, namespace, entry_id, response, account):
        response_map = {
            "accept": OL_RESPONSE_ACCEPTED,
            "decline": OL_RESPONSE_DECLINED,
            "tentative": OL_RESPONSE_TENTATIVE,
        }
        response_lower = response.lower().strip()
        if response_lower not in response_map:
            return f"Error: response must be 'accept', 'decline', or 'tentative'. Got: '{response}'"

        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        subject = item.Subject
        response_item = item.Respond(response_map[response_lower])
        response_item.Send()
        return f"Responded '{response_lower}' to meeting: '{subject}'"

    try:
        return await bridge.call(_respond, entry_id, response, account)
    except Exception as e:
        return f"Error responding to meeting: {format_com_error(e)}"


# =====================================================================
# TOOL 17: search_events
# =====================================================================

@mcp.tool()
async def search_events(
    query: str,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
    account: str = "",
) -> str:
    """Search for calendar events by keyword.

    Searches event subjects within a date range. Results are sorted by
    start time. Includes recurring event occurrences.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "standup", "review", "1:1".
        start_date: Start of search range in ISO 8601 format. Default: 30
            days ago.
        end_date: End of search range. Default: 30 days from now.
        count: Maximum results to return. Default 10.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of matching event summaries.
    """
    def _search(outlook, namespace, query, start_date, end_date, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        calendar = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now() - timedelta(days=30)
        end = _parse_date(end_date) if end_date else datetime.now() + timedelta(days=30)

        restrict = (
            f"[Start] >= '{_date_to_restrict_str(start)}' "
            f"AND [Start] <= '{_date_to_restrict_str(end)}'"
        )
        filtered = items.Restrict(restrict)

        query_lower = query.lower()
        results = []
        for item in filtered:
            if query_lower in (item.Subject or "").lower():
                try:
                    results.append(format_event_summary(item))
                except Exception:
                    continue
                if len(results) >= count:
                    break

        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_search, query, start_date, end_date, count, account)
    except Exception as e:
        return f"Error searching events: {format_com_error(e)}"


# =====================================================================
# TASK TOOLS
# =====================================================================

@mcp.tool()
async def list_tasks(
    include_completed: bool = False,
    count: int = 20,
    account: str = "",
) -> str:
    """List tasks from the Outlook Tasks folder.

    Returns a JSON array of task summaries sorted by due date. Each task
    includes entry_id, subject, status, percent_complete, due_date,
    importance, and categories.

    Args:
        include_completed: If true, include completed tasks. Default false
            (only pending/in-progress tasks).
        count: Maximum number of tasks to return. Default 20.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of task summary objects.
    """
    def _list(outlook, namespace, include_completed, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
        items = folder.Items
        items.Sort("[DueDate]")

        if not include_completed:
            items = items.Restrict("[Complete] = False")

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_task_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, include_completed, count, account)
    except Exception as e:
        return f"Error listing tasks: {format_com_error(e)}"


@mcp.tool()
async def get_task(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific task.

    Args:
        entry_id: The unique Outlook EntryID of the task.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full task details including body.
    """
    def _get(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        return json.dumps(format_task_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return f"Error reading task: {format_com_error(e)}"


@mcp.tool()
async def create_task(
    subject: str,
    body: str = "",
    due_date: str = "",
    importance: str = "normal",
    reminder_minutes: int = 0,
    account: str = "",
) -> str:
    """Create a new task in Outlook.

    Args:
        subject: The task title.
        body: Optional. Task description or notes.
        due_date: Optional. Due date in ISO 8601 format (e.g. "2026-03-01").
        importance: Optional. "low", "normal" (default), or "high".
        reminder_minutes: Optional. Minutes before due date to remind.
            Default 0 (no reminder).
        account: Optional. Account display name (or substring) to create
            the task in. Default: primary account.

    Returns:
        Confirmation with task subject and entry_id.
    """
    def _create(outlook, namespace, subject, body, due_date, importance,
                reminder_minutes, account):
        task = outlook.CreateItem(OL_TASK_ITEM)
        # Move to correct store's tasks folder if account specified
        if account:
            store = _require_store(namespace, account)
            tasks_folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
            task.Move(tasks_folder)
            task = namespace.GetItemFromID(task.EntryID)
        task.Subject = subject
        if body:
            task.Body = body
        if due_date:
            task.DueDate = due_date
        imp_map = {"low": 0, "normal": 1, "high": 2}
        task.Importance = imp_map.get(importance.lower(), 1)
        if reminder_minutes > 0:
            task.ReminderSet = True
            task.ReminderMinutesBeforeStart = reminder_minutes
        else:
            task.ReminderSet = False
        task.Save()
        return json.dumps({
            "status": "created",
            "subject": task.Subject,
            "entry_id": task.EntryID,
            "due_date": str(task.DueDate) if due_date else None,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, body, due_date, importance, reminder_minutes,
            account,
        )
    except Exception as e:
        return f"Error creating task: {format_com_error(e)}"


@mcp.tool()
async def complete_task(entry_id: str, account: str = "") -> str:
    """Mark a task as complete.

    Sets the task status to complete and percent to 100%.

    Args:
        entry_id: The unique Outlook EntryID of the task.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the task subject.
    """
    def _complete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_TASK, "task item"):
            return err
        item.Status = OL_TASK_COMPLETE
        item.PercentComplete = 100
        item.Save()
        return f"Task completed: '{item.Subject}'"

    try:
        return await bridge.call(_complete, entry_id, account)
    except Exception as e:
        return f"Error completing task: {format_com_error(e)}"


@mcp.tool()
async def delete_task(entry_id: str, account: str = "") -> str:
    """Delete a task from Outlook.

    Args:
        entry_id: The unique Outlook EntryID of the task to delete.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the task subject.
    """
    def _delete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_TASK, "task item"):
            return err
        subject = item.Subject
        item.Delete()
        return f"Task deleted: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id, account)
    except Exception as e:
        return f"Error deleting task: {format_com_error(e)}"


# =====================================================================
# ATTACHMENT TOOLS
# =====================================================================

@mcp.tool()
async def list_attachments(entry_id: str, account: str = "") -> str:
    """List all attachments on an email or calendar event.

    Args:
        entry_id: The EntryID of the email or event to check for attachments.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON array of attachment objects with index, filename, and size.
    """
    def _list(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        results = []
        for i in range(item.Attachments.Count):
            att = item.Attachments.Item(i + 1)
            results.append({
                "index": i + 1,
                "filename": att.FileName,
                "size": att.Size,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, entry_id, account)
    except Exception as e:
        return f"Error listing attachments: {format_com_error(e)}"


@mcp.tool()
async def save_attachment(
    entry_id: str,
    attachment_index: int = 1,
    save_directory: str = "",
    account: str = "",
) -> str:
    """Save an attachment from an email or event to disk.

    Downloads the specified attachment to a local directory.

    Args:
        entry_id: The EntryID of the email or event containing the attachment.
        attachment_index: Which attachment to save (1-based index). Default 1
            (first attachment). Use list_attachments to see available indices.
        save_directory: Directory to save the file to. Default: user's
            Downloads folder.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        The full file path where the attachment was saved, or an error.
    """
    def _save(outlook, namespace, entry_id, attachment_index, save_directory, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if attachment_index < 1 or item.Attachments.Count < attachment_index:
            return f"Error: Only {item.Attachments.Count} attachment(s), requested index {attachment_index}"

        att = item.Attachments.Item(attachment_index)
        if not save_directory:
            save_directory = os.path.join(os.path.expanduser("~"), "Downloads")

        # Resolve to real path before creating
        save_directory = os.path.realpath(save_directory)
        os.makedirs(save_directory, exist_ok=True)

        # Strip path separators and dangerous characters from filename
        safe_name = os.path.basename(att.FileName)
        safe_name = re.sub(r'[^\w\.\-_ ]', '_', safe_name)
        if not safe_name:
            safe_name = "attachment"

        save_path = os.path.join(save_directory, safe_name)

        # Ensure final path is still inside the intended directory
        if not os.path.realpath(save_path).startswith(save_directory + os.sep) and \
           os.path.realpath(save_path) != save_directory:
            return "Error: Attachment filename would escape the target directory."

        att.SaveAsFile(save_path)
        return json.dumps({
            "status": "saved",
            "filename": safe_name,
            "path": save_path,
            "size": att.Size,
        }, indent=2, default=str)

    try:
        return await bridge.call(_save, entry_id, attachment_index, save_directory, account)
    except Exception as e:
        return f"Error saving attachment: {format_com_error(e)}"


# =====================================================================
# CATEGORY TOOLS
# =====================================================================

@mcp.tool()
async def list_categories(account: str = "") -> str:
    """List all available Outlook categories.

    Returns the color categories configured in the user's Outlook profile.
    These can be applied to emails, events, tasks, and other items.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of category objects with name and color index.
    """
    def _list(outlook, namespace, account):
        # Categories are profile-wide, not per-store, but we accept the param for consistency
        results = []
        for i in range(namespace.Categories.Count):
            cat = namespace.Categories.Item(i + 1)
            results.append({"name": cat.Name, "color": cat.Color})
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return f"Error listing categories: {format_com_error(e)}"


@mcp.tool()
async def set_category(
    entry_id: str,
    categories: str,
    account: str = "",
) -> str:
    """Set categories on an email, event, or task.

    Replaces any existing categories on the item. Use comma-separated
    values for multiple categories.

    Args:
        entry_id: The EntryID of the item to categorize.
        categories: Category name(s), comma-separated. Example:
            "Important" or "Work, Follow-up". Use an empty string to
            clear all categories.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the item subject and applied categories.
    """
    def _set(outlook, namespace, entry_id, categories, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        item.Categories = categories
        item.Save()
        return (
            f"Categories set on '{item.Subject}': "
            f"'{item.Categories or '(none)'}'"
        )

    try:
        return await bridge.call(_set, entry_id, categories, account)
    except Exception as e:
        return f"Error setting categories: {format_com_error(e)}"


# =====================================================================
# RULES TOOLS
# =====================================================================

@mcp.tool()
async def list_rules(account: str = "") -> str:
    """List all mail rules in Outlook.

    Returns the configured inbox rules with their names and enabled status.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of rule objects with name, enabled status, and index.
    """
    def _list(outlook, namespace, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        results = []
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            results.append({
                "index": i + 1,
                "name": rule.Name,
                "enabled": bool(rule.Enabled),
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return f"Error listing rules: {format_com_error(e)}"


@mcp.tool()
async def get_rule(rule_name: str = "", account: str = "") -> str:
    """Read a mail rule's full structure (conditions, exceptions, actions).

    Fills the gap list_rules leaves: list_rules returns only name/enabled/index,
    and create_rule/update_rule can set conditions but not read them back. This
    serializes the rule(s) to JSON, resolving recipient conditions to primary
    SMTP addresses server-side (Exchange X.500 DNs are resolved for you, so the
    caller never sees an empty .Address or a raw DN).

    Args:
        rule_name: Exact rule name. Empty (default) returns ALL rules as an array.
        account: Optional. Account display name (or substring) to target.

    Returns:
        JSON — a single rule object (rule_name given) or an array (all rules).
        Each rule carries `fully_representable`: true when every condition/action
        is within the set update_rule can edit in place (false => only the
        Outlook Rules Wizard can safely edit it).
    """
    def _get(outlook, namespace, rule_name, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        if rule_name:
            for i in range(1, rules.Count + 1):
                if rules.Item(i).Name == rule_name:
                    return json.dumps(_serialize_rule(rules.Item(i)),
                                      indent=2, default=str)
            return (f"Error: Rule '{rule_name}' not found. "
                    "Use list_rules to see available rules.")
        out = [_serialize_rule(rules.Item(i)) for i in range(1, rules.Count + 1)]
        return json.dumps(out, indent=2, default=str)

    try:
        return await bridge.call(_get, rule_name, account)
    except Exception as e:
        return f"Error reading rule: {format_com_error(e)}"


@mcp.tool()
async def toggle_rule(
    rule_name: str,
    enabled: bool,
    account: str = "",
) -> str:
    """Enable or disable a mail rule by name.

    CAUTION: This modifies live mail rules immediately. Confirm the rule name
    with list_rules before calling.

    Args:
        rule_name: The exact name of the rule to toggle. Use list_rules
            to see available rule names.
        enabled: True to enable the rule, False to disable it.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        Confirmation with the rule name and new status.
    """
    def _toggle(outlook, namespace, rule_name, enabled, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            if rule.Name == rule_name:
                logger.warning(
                    "toggle_rule: setting rule '%s' enabled=%s", rule_name, enabled
                )
                rule.Enabled = enabled
                rules.Save()
                status = "enabled" if enabled else "disabled"
                return f"Rule '{rule_name}' {status}"
        return f"Error: Rule '{rule_name}' not found. Use list_rules to see available rules."

    try:
        return await bridge.call(_toggle, rule_name, enabled, account)
    except Exception as e:
        return f"Error toggling rule: {format_com_error(e)}"


# --- Rule helpers (shared by create_rule / update_rule) ---

# OlRuleConditionType / OlRuleActionType ints these tools can read AND write
# back losslessly (verified by round-trip spike). update_rule refuses to
# re-save a rule containing anything outside these sets, because Save()
# re-serializes the whole rule and an unrepresentable element could be dropped.
_SAFE_CONDITION_TYPES = {1, 2, 11, 12, 13, 15, 17, 23}
#   From=1 Subject=2 NotTo=11 SentTo=12 Body=13 MessageHeader=15
#   SenderAddress=17 FormName=23
_SAFE_ACTION_TYPES = {1, 2, 3, 21}
#   MoveToFolder=1 AssignToCategory=2 Delete=3 Stop=21


def _split_list(raw):
    """Split a ';'-separated tool argument into a clean list of fragments."""
    return [s.strip() for s in raw.split(";") if s.strip()]


def _force_folder_putref(action, folder):
    """Set a RuleAction's by-reference Folder property.

    A plain attribute assignment (PROPERTYPUT) silently no-ops for this
    by-reference property, leaving an invalid action that fails Rules.Save().
    Force PROPERTYPUTREF via a low-level Invoke.
    """
    import pythoncom
    dispid = action._oleobj_.GetIDsOfNames("Folder")
    action._oleobj_.Invoke(
        dispid, 0, pythoncom.DISPATCH_PROPERTYPUTREF, False, folder
    )


def _unsupported_rule_elements(rule):
    """List labels for enabled conditions/actions outside the round-trip-safe
    set. Empty => the rule is safe to edit in place; non-empty => a Save() might
    drop those elements, so update_rule refuses to touch it.
    """
    bad = []
    checks = (
        (rule.Conditions, _SAFE_CONDITION_TYPES, "condition", False),
        (rule.Exceptions, _SAFE_CONDITION_TYPES, "exception", False),
        (rule.Actions, _SAFE_ACTION_TYPES, "action", True),
    )
    for coll, safe, kind, is_action in checks:
        try:
            n = coll.Count
        except Exception:
            bad.append(f"<non-enumerable {kind} set>")
            continue
        for i in range(1, n + 1):
            try:
                it = coll.Item(i)
                if not it.Enabled:
                    continue
                t = int(it.ActionType) if is_action else int(it.ConditionType)
                if t not in safe:
                    bad.append(f"{kind} type {t}")
            except Exception:
                bad.append(f"<unreadable {kind}>")
    return bad


# OlRuleConditionType / OlRuleActionType -> friendly name (from the Outlook
# typelib; do NOT hand-guess these — a guessed map mislabeled Stop as
# "StartApplication" and SenderAddress as "Sensitivity").
_CONDITION_NAMES = {
    0: "Unknown", 1: "From", 2: "Subject", 3: "Account", 4: "OnlyToMe",
    5: "To", 6: "Importance", 7: "Sensitivity", 8: "FlaggedForAction",
    9: "Cc", 10: "ToOrCc", 11: "NotTo", 12: "SentTo", 13: "Body",
    14: "BodyOrSubject", 15: "MessageHeader", 16: "RecipientAddress",
    17: "SenderAddress", 18: "Category", 19: "OOF", 20: "HasAttachment",
    21: "SizeRange", 22: "DateRange", 23: "FormName", 24: "Property",
    25: "SenderInAddressBook", 26: "MeetingInviteOrUpdate",
    27: "LocalMachineOnly", 28: "OtherMachine", 29: "AnyCategory",
    30: "FromRssFeed", 31: "FromAnyRssFeed",
}
_ACTION_NAMES = {
    0: "Unknown", 1: "MoveToFolder", 2: "AssignToCategory", 3: "Delete",
    4: "DeletePermanently", 5: "CopyToFolder", 6: "Forward",
    7: "ForwardAsAttachment", 8: "Redirect", 9: "ServerReply", 10: "Template",
    11: "FlagForActionInDays", 12: "FlagColor", 13: "FlagClear",
    14: "Importance", 15: "Sensitivity", 16: "Print", 17: "PlaySound",
    18: "StartApplication", 19: "MarkRead", 20: "RunScript", 21: "Stop",
    22: "CustomAction", 23: "NewItemAlert", 24: "DesktopAlert",
    25: "NotifyRead", 26: "NotifyDelivery", 27: "CcMessage", 28: "Defer",
    29: "MarkAsTask", 30: "ClearCategories",
}

_PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"


def _recipient_smtp(recipient):
    """Resolve a rule recipient to its primary SMTP address.

    Rule-condition recipients report an empty .Address; the real value lives on
    the AddressEntry, where Exchange entries carry an X.500 DN ("/o=...") rather
    than SMTP. Resolve those via GetExchangeUser / GetExchangeDistributionList,
    fall back to the PR_SMTP_ADDRESS MAPI property, then to the raw DN. Callers
    never see an empty address or a bare DN.
    """
    try:
        ae = recipient.AddressEntry
    except Exception:
        return ""
    try:
        addr = ae.Address or ""
    except Exception:
        addr = ""
    if addr and not addr.startswith("/"):
        return addr  # already an SMTP address
    for getter in ("GetExchangeUser", "GetExchangeDistributionList"):
        try:
            obj = getattr(ae, getter)()
            if obj and obj.PrimarySmtpAddress:
                return obj.PrimarySmtpAddress
        except Exception:
            pass
    try:
        val = ae.PropertyAccessor.GetProperty(_PR_SMTP_ADDRESS)
        if val:
            return val
    except Exception:
        pass
    return addr


def _serialize_conditions(coll):
    """Serialize the enabled conditions of a RuleConditions collection.

    The enumerated Item(i) exposes only ConditionType/Enabled — the typed
    payload (text, recipients, ...) is reachable only via the matching NAMED
    property (coll.From, coll.Body, ...). So read the type from the item but the
    payload from getattr(coll, <name>).
    """
    out = []
    try:
        n = coll.Count
    except Exception:
        return out
    for i in range(1, n + 1):
        try:
            it = coll.Item(i)
            if not it.Enabled:
                continue
            t = int(it.ConditionType)
        except Exception:
            continue
        name = _CONDITION_NAMES.get(t)
        d = {"type": name or f"type{t}"}
        cond = it
        if name:
            try:
                cond = getattr(coll, name)
            except Exception:
                cond = it
        for attr, key in (("Text", "text"), ("Address", "address"),
                          ("Categories", "categories"), ("FormName", "form")):
            try:
                v = getattr(cond, attr)
                if v:
                    d[key] = list(v) if not isinstance(v, str) else v
            except Exception:
                pass
        try:
            reps = cond.Recipients
            rcs = [{"display": reps.Item(k).Name,
                    "smtp": _recipient_smtp(reps.Item(k))}
                   for k in range(1, reps.Count + 1)]
            if rcs:
                d["recipients"] = rcs
        except Exception:
            pass
        out.append(d)
    return out


def _relative_folder_path(folder):
    """Convert a MAPIFolder's FolderPath ("\\\\Store\\Inbox\\Sub") to a relative
    slash-path ("Inbox/Sub") — the same form create_rule / update_rule accept for
    move_to_folder, so get_rule output round-trips straight back as input.
    """
    try:
        path = folder.FolderPath
    except Exception:
        return None
    if not path:
        return None
    parts = path.lstrip("\\").split("\\")
    if len(parts) > 1:
        parts = parts[1:]  # drop the leading store/root segment
    return "/".join(parts)


def _serialize_actions(coll):
    """Serialize the enabled actions of a RuleActions collection.

    Same indirection as _serialize_conditions: the action's payload (target
    folder, categories) lives on the named property, not the enumerated item.
    """
    out = []
    try:
        n = coll.Count
    except Exception:
        return out
    for i in range(1, n + 1):
        try:
            it = coll.Item(i)
            if not it.Enabled:
                continue
            t = int(it.ActionType)
        except Exception:
            continue
        name = _ACTION_NAMES.get(t)
        d = {"type": name or f"type{t}"}
        act = it
        if name:
            try:
                act = getattr(coll, name)
            except Exception:
                act = it
        try:
            if act.Folder:
                d["folder"] = _relative_folder_path(act.Folder)
        except Exception:
            pass
        try:
            if act.Categories:
                d["categories"] = list(act.Categories)
        except Exception:
            pass
        out.append(d)
    return out


def _serialize_rule(rule):
    """Full structured view of a rule, with recipients resolved to SMTP and a
    `fully_representable` flag (= editable in place by update_rule)."""
    try:
        order = rule.ExecutionOrder
    except Exception:
        order = None
    return {
        "name": rule.Name,
        "enabled": bool(rule.Enabled),
        "execution_order": order,
        "conditions": _serialize_conditions(rule.Conditions),
        "exceptions": _serialize_conditions(rule.Exceptions),
        "actions": _serialize_actions(rule.Actions),
        "fully_representable": not _unsupported_rule_elements(rule),
    }


@mcp.tool()
async def create_rule(
    name: str,
    move_to_folder: str = "",
    from_addresses: str = "",
    subject_contains: str = "",
    body_contains: str = "",
    header_contains: str = "",
    sent_to: str = "",
    not_to: bool = False,
    assign_category: str = "",
    delete: bool = False,
    stop_processing: bool = False,
    enabled: bool = True,
    account: str = "",
) -> str:
    """Create a new receive rule in Outlook.

    CAUTION: This creates a live mail rule immediately. Rules run on incoming
    mail — confirm the conditions and the target folder before calling.

    At least one condition (from_addresses, subject_contains, body_contains,
    header_contains, sent_to, or not_to) AND at least one action
    (move_to_folder, assign_category, delete, or stop_processing) must be supplied.

    Args:
        name: Display name for the new rule. Must be unique.
        move_to_folder: Optional. Folder to move matching mail to. Accepts a
            built-in name ("inbox"), a root folder name, or a slash-path
            ("Inbox/Receipts") — resolved the same way as list_emails.
        from_addresses: Optional. Semicolon-separated sender-address fragments;
            matches when the sender SMTP address contains any of them
            (e.g. "github.com;noreply@gitlab.com").
        subject_contains: Optional. Semicolon-separated words/phrases; matches
            when the subject contains any of them.
        body_contains: Optional. Semicolon-separated words/phrases matched
            against the message body.
        header_contains: Optional. Semicolon-separated words/phrases matched
            against the raw message header (e.g. "X-GitHub-Reason").
        sent_to: Optional. Semicolon-separated recipient addresses; matches when
            the message was sent to any of them.
        not_to: Optional. Match mail where you are NOT a direct (To) recipient
            (e.g. you were only CC'd). Default False.
        assign_category: Optional. Color-category name to assign to matching mail.
        delete: Optional. Move matching mail to Deleted Items. Default False.
        stop_processing: Optional. Stop evaluating further rules after this one.
            Default False.
        enabled: Optional. Whether the rule is active. Default True.
        account: Optional. Account display name (or substring) to target.

    Returns:
        Confirmation with the rule name and a summary of its conditions/actions.

    Note:
        Outlook's COM object model exposes only a subset of the Rules Wizard's
        conditions/actions. The MoveToFolder action's Folder is a by-reference
        property: a plain assignment silently no-ops and the Save() then fails
        with "invalid actions or conditions"; this tool forces the required
        PROPERTYPUTREF internally.
    """
    def _create(outlook, namespace, name, move_to_folder, from_addresses,
                subject_contains, body_contains, header_contains, sent_to,
                not_to, assign_category, delete, stop_processing, enabled,
                account):
        store = _require_store(namespace, account)

        senders = _split_list(from_addresses)
        subjects = _split_list(subject_contains)
        bodies = _split_list(body_contains)
        headers = _split_list(header_contains)
        tos = _split_list(sent_to)

        if not (senders or subjects or bodies or headers or tos or not_to):
            return ("Error: at least one condition required (from_addresses, "
                    "subject_contains, body_contains, header_contains, sent_to, "
                    "or not_to).")
        if not (move_to_folder or assign_category or delete or stop_processing):
            return ("Error: at least one action required "
                    "(move_to_folder, assign_category, delete, or stop_processing).")

        # Resolve the target folder up front, so we fail before creating a
        # half-built rule rather than after.
        target = None
        if move_to_folder:
            target = _resolve_folder(namespace, move_to_folder, store)
            if target is None:
                return (f"Error: folder '{move_to_folder}' not found. "
                        "Use list_folders to see available folders.")

        rules = store.GetRules()
        for i in range(rules.Count):
            if rules.Item(i + 1).Name == name:
                return f"Error: a rule named '{name}' already exists."

        rule = rules.Create(name, 0)  # 0 = olRuleReceive

        # --- Conditions ---
        conds = []
        if senders:
            cond = rule.Conditions.SenderAddress
            cond.Address = senders
            cond.Enabled = True
            conds.append(f"sender address contains {senders}")
        if subjects:
            cond = rule.Conditions.Subject
            cond.Text = subjects
            cond.Enabled = True
            conds.append(f"subject contains {subjects}")
        if bodies:
            cond = rule.Conditions.Body
            cond.Text = bodies
            cond.Enabled = True
            conds.append(f"body contains {bodies}")
        if headers:
            cond = rule.Conditions.MessageHeader
            cond.Text = headers
            cond.Enabled = True
            conds.append(f"header contains {headers}")
        if tos:
            cond = rule.Conditions.SentTo
            for addr in tos:
                cond.Recipients.Add(addr)
            cond.Recipients.ResolveAll()
            cond.Enabled = True
            conds.append(f"sent to {tos}")
        if not_to:
            rule.Conditions.NotTo.Enabled = True
            conds.append("not sent directly to me")

        # --- Actions ---
        applied = []
        if target is not None:
            act = rule.Actions.MoveToFolder
            act.Enabled = True
            _force_folder_putref(act, target)
            applied.append(f"move to '{target.Name}'")
        if assign_category:
            act = rule.Actions.AssignToCategory
            act.Categories = [assign_category]
            act.Enabled = True
            applied.append(f"assign category '{assign_category}'")
        if delete:
            rule.Actions.Delete.Enabled = True
            applied.append("delete (move to Deleted Items)")
        if stop_processing:
            rule.Actions.Stop.Enabled = True
            applied.append("stop processing further rules")

        rule.Enabled = bool(enabled)

        logger.warning("create_rule: creating rule '%s'", name)
        rules.Save()

        status = "enabled" if enabled else "disabled"
        return (
            f"Rule '{name}' created ({status}).\n"
            f"Conditions: {'; '.join(conds)}\n"
            f"Actions: {'; '.join(applied)}"
        )

    try:
        return await bridge.call(
            _create, name, move_to_folder, from_addresses, subject_contains,
            body_contains, header_contains, sent_to, not_to, assign_category,
            delete, stop_processing, enabled, account,
        )
    except Exception as e:
        return f"Error creating rule: {format_com_error(e)}"


@mcp.tool()
async def delete_rule(rule_name: str, account: str = "") -> str:
    """Delete a mail rule by name.

    CAUTION: This permanently removes a live mail rule. Confirm the exact name
    with list_rules before calling.

    Args:
        rule_name: Exact name of the rule to delete.
        account: Optional. Account display name (or substring) to target.

    Returns:
        Confirmation, or an error if the rule was not found.
    """
    def _delete(outlook, namespace, rule_name, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        for i in range(rules.Count, 0, -1):
            if rules.Item(i).Name == rule_name:
                logger.warning("delete_rule: removing rule '%s'", rule_name)
                rules.Remove(i)
                rules.Save()
                return f"Rule '{rule_name}' deleted."
        return (f"Error: Rule '{rule_name}' not found. "
                "Use list_rules to see available rules.")

    try:
        return await bridge.call(_delete, rule_name, account)
    except Exception as e:
        return f"Error deleting rule: {format_com_error(e)}"


@mcp.tool()
async def rename_rule(rule_name: str, new_name: str, account: str = "") -> str:
    """Rename an existing mail rule.

    Changes only the rule's display name — its conditions and actions are left
    untouched. This is safe even for rules created in the Rules Wizard with
    conditions the COM object model can't fully represent: there is no edit of
    a rule's logic here, so nothing can be silently dropped on Save().

    (To change a rule's conditions or actions, use update_rule, which edits in
    place and preserves everything it doesn't touch.)

    Args:
        rule_name: Exact current name of the rule. Use list_rules to confirm.
        new_name: New display name. Must not collide with an existing rule.
        account: Optional. Account display name (or substring) to target.

    Returns:
        Confirmation, or an error if the rule was not found or the name is taken.
    """
    def _rename(outlook, namespace, rule_name, new_name, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        target = None
        for i in range(rules.Count):
            name = rules.Item(i + 1).Name
            if name == new_name:
                return f"Error: a rule named '{new_name}' already exists."
            if name == rule_name:
                target = rules.Item(i + 1)
        if target is None:
            return (f"Error: Rule '{rule_name}' not found. "
                    "Use list_rules to see available rules.")
        logger.warning("rename_rule: '%s' -> '%s'", rule_name, new_name)
        target.Name = new_name
        rules.Save()
        return f"Rule '{rule_name}' renamed to '{new_name}'."

    try:
        return await bridge.call(_rename, rule_name, new_name, account)
    except Exception as e:
        return f"Error renaming rule: {format_com_error(e)}"


@mcp.tool()
async def update_rule(
    rule_name: str,
    move_to_folder: str = "",
    from_addresses: str = "",
    subject_contains: str = "",
    body_contains: str = "",
    header_contains: str = "",
    sent_to: str = "",
    assign_category: str = "",
    not_to: bool | None = None,
    delete: bool | None = None,
    stop_processing: bool | None = None,
    account: str = "",
) -> str:
    """Modify an existing rule in place, changing only the facets you specify.

    CAUTION: This modifies a live mail rule immediately.

    Unlike delete + recreate, this PRESERVES every condition and action you do
    NOT name — including ones create_rule can't build — because Save() keeps the
    rest of the rule intact (verified by round-trip).

    SAFETY: before saving, the rule is scanned; if it contains any enabled
    condition or action of a type these tools can't read back losslessly, the
    update is REFUSED with the rule untouched (edit such a rule in Outlook).

    Facet semantics:
      - List args (from_addresses, subject_contains, body_contains,
        header_contains, sent_to): a non-empty value REPLACES that condition;
        empty (default) leaves it unchanged.
      - move_to_folder / assign_category: a non-empty value sets/changes that
        action.
      - not_to / delete / stop_processing: True enables, False disables, omitted
        (None) leaves unchanged.

    Args:
        rule_name: Exact name of the rule to modify.
        move_to_folder: New target folder (built-in name, root name, or slash-path).
        from_addresses: Semicolon-separated sender-address fragments (replaces).
        subject_contains: Semicolon-separated subject phrases (replaces).
        body_contains: Semicolon-separated body phrases (replaces).
        header_contains: Semicolon-separated message-header phrases (replaces).
        sent_to: Semicolon-separated recipient addresses (replaces).
        assign_category: Color-category name to assign.
        not_to: Enable/disable the "not sent directly to me" condition.
        delete: Enable/disable the delete (move to Deleted Items) action.
        stop_processing: Enable/disable the "stop processing more rules" action.
        account: Optional. Account display name (or substring) to target.

    Returns:
        Confirmation summarizing what changed, or an error.
    """
    def _update(outlook, namespace, rule_name, move_to_folder, from_addresses,
                subject_contains, body_contains, header_contains, sent_to,
                assign_category, not_to, delete, stop_processing, account):
        store = _require_store(namespace, account)

        # Resolve folder up front (fail before mutating anything).
        target = None
        if move_to_folder:
            target = _resolve_folder(namespace, move_to_folder, store)
            if target is None:
                return (f"Error: folder '{move_to_folder}' not found. "
                        "Use list_folders to see available folders.")

        rules = store.GetRules()
        rule = None
        for i in range(1, rules.Count + 1):
            if rules.Item(i).Name == rule_name:
                rule = rules.Item(i)
                break
        if rule is None:
            return (f"Error: Rule '{rule_name}' not found. "
                    "Use list_rules to see available rules.")

        # Refuse to re-save a rule with elements we can't round-trip safely.
        bad = _unsupported_rule_elements(rule)
        if bad:
            return ("Error: rule '%s' contains element(s) this tool can't safely "
                    "re-save (%s). Edit it manually in Outlook to avoid data loss."
                    % (rule_name, ", ".join(sorted(set(bad)))))

        changed = []
        froms = _split_list(from_addresses)
        if froms:
            c = rule.Conditions.SenderAddress
            c.Address = froms
            c.Enabled = True
            changed.append(f"sender address -> {froms}")
        subjects = _split_list(subject_contains)
        if subjects:
            c = rule.Conditions.Subject
            c.Text = subjects
            c.Enabled = True
            changed.append(f"subject -> {subjects}")
        bodies = _split_list(body_contains)
        if bodies:
            c = rule.Conditions.Body
            c.Text = bodies
            c.Enabled = True
            changed.append(f"body -> {bodies}")
        headers = _split_list(header_contains)
        if headers:
            c = rule.Conditions.MessageHeader
            c.Text = headers
            c.Enabled = True
            changed.append(f"header -> {headers}")
        tos = _split_list(sent_to)
        if tos:
            c = rule.Conditions.SentTo
            while c.Recipients.Count > 0:   # replace, don't append
                c.Recipients.Remove(1)
            for addr in tos:
                c.Recipients.Add(addr)
            c.Recipients.ResolveAll()
            c.Enabled = True
            changed.append(f"sent-to -> {tos}")
        if not_to is not None:
            rule.Conditions.NotTo.Enabled = bool(not_to)
            changed.append(f"not-to {'on' if not_to else 'off'}")

        if target is not None:
            act = rule.Actions.MoveToFolder
            act.Enabled = True
            _force_folder_putref(act, target)
            changed.append(f"move to '{target.Name}'")
        if assign_category:
            act = rule.Actions.AssignToCategory
            act.Categories = [assign_category]
            act.Enabled = True
            changed.append(f"category -> '{assign_category}'")
        if delete is not None:
            rule.Actions.Delete.Enabled = bool(delete)
            changed.append(f"delete {'on' if delete else 'off'}")
        if stop_processing is not None:
            rule.Actions.Stop.Enabled = bool(stop_processing)
            changed.append(f"stop {'on' if stop_processing else 'off'}")

        if not changed:
            return "Error: nothing to update — specify at least one facet to change."

        logger.warning("update_rule: modifying '%s' (%s)", rule_name,
                        "; ".join(changed))
        rules.Save()
        return f"Rule '{rule_name}' updated.\nChanged: {'; '.join(changed)}"

    try:
        return await bridge.call(
            _update, rule_name, move_to_folder, from_addresses, subject_contains,
            body_contains, header_contains, sent_to, assign_category, not_to,
            delete, stop_processing, account,
        )
    except Exception as e:
        return f"Error updating rule: {format_com_error(e)}"


@mcp.tool()
async def run_rule_now(
    rule_name: str,
    folder: str = "inbox",
    include_subfolders: bool = False,
    account: str = "",
) -> str:
    """Run an existing rule against mail ALREADY in a folder (backlog sweep).

    CAUTION: this MUTATES mail immediately — it applies the rule's actions
    (move/delete/categorize) to matching messages already sitting in the folder.
    Confirm the rule and folder before calling. Normal rules only fire on
    incoming mail; this is how you retroactively apply one to a backlog.

    Args:
        rule_name: Exact name of the rule to run. Use get_rule to inspect it first.
        folder: Folder to run against (built-in name, root name, or slash-path).
            Default "inbox".
        include_subfolders: Also process the folder's subfolders. Default False.
        account: Optional. Account display name (or substring) to target.

    Returns:
        Confirmation, or an error if the rule or folder was not found.
    """
    def _run(outlook, namespace, rule_name, folder, include_subfolders, account):
        store = _require_store(namespace, account)

        target = _resolve_folder(namespace, folder, store) if folder else None
        if target is None:
            return (f"Error: folder '{folder}' not found. "
                    "Use list_folders to see available folders.")

        rules = store.GetRules()
        rule = None
        for i in range(1, rules.Count + 1):
            if rules.Item(i).Name == rule_name:
                rule = rules.Item(i)
                break
        if rule is None:
            return (f"Error: Rule '{rule_name}' not found. "
                    "Use list_rules to see available rules.")

        logger.warning("run_rule_now: executing '%s' on '%s' (subfolders=%s)",
                       rule_name, target.Name, include_subfolders)
        # Execute(ShowProgress, Folder, IncludeSubfolders, RuleExecuteOption);
        # RuleExecuteOption 0 = all messages (OlRuleExecuteOption not in typelib).
        rule.Execute(False, target, bool(include_subfolders), 0)

        sub = " (incl. subfolders)" if include_subfolders else ""
        return f"Rule '{rule_name}' executed against '{target.Name}'{sub}."

    try:
        return await bridge.call(
            _run, rule_name, folder, include_subfolders, account,
        )
    except Exception as e:
        return f"Error running rule: {format_com_error(e)}"


@mcp.tool()
async def reorder_rule(rule_name: str, position: int, account: str = "") -> str:
    """Move a rule to a new position in the execution order.

    Rules run top-down by ExecutionOrder (1 = runs first); position decides which
    rule wins when several match (e.g. a specific rule before a catch-all, or
    before one that stops processing).

    CAUTION: modifies live rules immediately. Like every rule write this re-saves
    the WHOLE rule set, so do NOT run it while you are editing rules in the
    Outlook Rules Wizard — a concurrent edit could be clobbered.

    Args:
        rule_name: Exact name of the rule to move. Use list_rules to confirm.
        position: New 1-based position in the execution order (1 = runs first).
            Must be within 1..number-of-rules; other rules shift to make room.
        account: Optional. Account display name (or substring) to target.

    Returns:
        Confirmation with the old and new position, or an error.
    """
    def _reorder(outlook, namespace, rule_name, position, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        count = rules.Count
        # ExecutionOrder must be in 1..count; out-of-range raises a COM
        # "parameter is incorrect" error, so validate up front.
        if position < 1 or position > count:
            return f"Error: position {position} out of range (1..{count})."

        rule = None
        for i in range(1, count + 1):
            if rules.Item(i).Name == rule_name:
                rule = rules.Item(i)
                break
        if rule is None:
            return (f"Error: Rule '{rule_name}' not found. "
                    "Use list_rules to see available rules.")

        old = rule.ExecutionOrder
        if old == position:
            return f"Rule '{rule_name}' is already at position {position}."

        logger.warning("reorder_rule: '%s' %d -> %d", rule_name, old, position)
        rule.ExecutionOrder = position
        rules.Save()

        # Re-read to report the order actually applied.
        fresh = store.GetRules()
        new_pos = position
        for i in range(1, fresh.Count + 1):
            if fresh.Item(i).Name == rule_name:
                new_pos = fresh.Item(i).ExecutionOrder
                break
        return f"Rule '{rule_name}' moved from position {old} to {new_pos}."

    try:
        return await bridge.call(_reorder, rule_name, position, account)
    except Exception as e:
        return f"Error reordering rule: {format_com_error(e)}"


# =====================================================================
# OUT OF OFFICE TOOLS
# =====================================================================

@mcp.tool()
async def get_out_of_office(account: str = "") -> str:
    """Check the current Out of Office (auto-reply) status.

    Returns whether Out of Office is currently enabled.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with the OOF status.
    """
    def _get(outlook, namespace, account):
        store = _require_store(namespace, account)
        try:
            prop_tag = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
            oof_state = store.PropertyAccessor.GetProperty(prop_tag)
            return json.dumps({
                "out_of_office": bool(oof_state),
                "status": "on" if oof_state else "off",
            }, indent=2)
        except Exception:
            return json.dumps({
                "out_of_office": None,
                "status": "unknown",
                "note": "Could not read OOF property. Check Outlook settings directly.",
            }, indent=2)

    try:
        return await bridge.call(_get, account)
    except Exception as e:
        return f"Error checking OOF status: {format_com_error(e)}"


# =====================================================================
# Entry point
# =====================================================================

def main():
    logger.info("Starting Outlook Desktop MCP server...")
    bridge.start()
    logger.info("COM bridge ready. Starting MCP stdio transport...")
    try:
        mcp.run(transport="stdio")
    finally:
        bridge.stop()


if __name__ == "__main__":
    main()
