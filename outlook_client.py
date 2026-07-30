"""
outlook_client.py
Thin wrapper around Outlook COM automation (classic desktop Outlook, Windows only).
Requires Outlook to be installed and already signed in / running is not required,
but the profile must be configured.
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from typing import Any

import win32com.client

OL_FOLDER_INBOX = 6
RECIPIENT_TYPE_TO = 1
RECIPIENT_TYPE_CC = 2
RECIPIENT_TYPE_BCC = 3

OL_MAIL = 43
# Meeting requests/cancellations arrive in the Inbox as their own item types
# (not olMail), but carry a Body that often has the organizer's agenda/
# description text - worth scoring like a regular mail. Meeting *responses*
# (accept/decline/tentative, classes 55-57) are plain RSVP acks with no real
# content, so they're intentionally left out.
OL_MEETING_REQUEST = 53
OL_MEETING_CANCELLATION = 54
MEETING_ITEM_CLASSES = {OL_MEETING_REQUEST: "Request", OL_MEETING_CANCELLATION: "Cancellation"}


@dataclass
class MailInfo:
    entry_id: str
    subject: str
    sender_name: str
    sender_email: str
    received: dt.datetime
    importance: int  # 0=Low, 1=Normal, 2=High
    flagged: bool
    recipient_type: str  # "To", "CC", or "Unknown"
    recipients: list[str]  # display names of all To/CC recipients (excludes sender)
    body_preview: str
    categories: str
    item: Any = field(repr=False, compare=False)  # live COM MailItem, not for CSV export
    is_meeting: bool = False
    meeting_type: str = ""  # "Request" or "Cancellation" when is_meeting is True, else ""


def _connect():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    return outlook, namespace


def _get_smtp_address(mail_or_recipient, email_type_attr: str, email_addr_attr: str) -> str:
    """Resolve a proper SMTP address, handling Exchange DN-style addresses."""
    try:
        if getattr(mail_or_recipient, email_type_attr, None) == "EX":
            exch_user = mail_or_recipient.GetExchangeUser() if hasattr(mail_or_recipient, "GetExchangeUser") else None
            if exch_user is not None:
                return exch_user.PrimarySmtpAddress
    except Exception:
        pass
    return getattr(mail_or_recipient, email_addr_attr, "") or ""


def _current_user_address(namespace) -> str:
    try:
        current_user = namespace.CurrentUser
        exch_user = current_user.GetExchangeUser()
        if exch_user is not None:
            return (exch_user.PrimarySmtpAddress or "").lower()
        return (current_user.Address or "").lower()
    except Exception:
        return ""


def get_current_user_info() -> dict:
    """
    Best-effort detection of the signed-in Outlook user's display name and
    email address, used to prefill the setup wizard (my_names/trusted_domains)
    so a teammate doesn't have to type them manually. Returns {} if Outlook
    isn't available/signed in - callers should treat this as optional.
    """
    try:
        _outlook, namespace = _connect()
        name = namespace.CurrentUser.Name or ""
        email = _current_user_address(namespace)
        domain = email.split("@")[-1] if "@" in email else ""
        return {"name": name, "email": email, "domain": domain}
    except Exception:
        return {}


def _recipient_type_for_current_user(mail_item, current_user_address: str) -> str:
    if not current_user_address:
        return "Unknown"
    try:
        for recipient in mail_item.Recipients:
            addr = _get_smtp_address(recipient, "AddressEntry.Type", "Address")
            try:
                if recipient.AddressEntry and recipient.AddressEntry.Type == "EX":
                    exch_user = recipient.AddressEntry.GetExchangeUser()
                    if exch_user is not None:
                        addr = exch_user.PrimarySmtpAddress
            except Exception:
                pass
            if addr and addr.lower() == current_user_address:
                if recipient.Type == RECIPIENT_TYPE_TO:
                    return "To"
                if recipient.Type == RECIPIENT_TYPE_CC:
                    return "CC"
                if recipient.Type == RECIPIENT_TYPE_BCC:
                    return "BCC"
    except Exception:
        pass
    return "Unknown"


def _parse_mail_item(item, current_user_address: str, max_body_chars: int) -> MailInfo | None:
    item_class = item.Class
    meeting_type = MEETING_ITEM_CLASSES.get(item_class, "")
    is_meeting = bool(meeting_type)
    if item_class != OL_MAIL and not is_meeting:
        return None
    received = item.ReceivedTime
    received_naive = dt.datetime(received.year, received.month, received.day,
                                  received.hour, received.minute, received.second)

    sender_email = _get_smtp_address(item, "SenderEmailType", "SenderEmailAddress")
    body = (item.Body or "")[:max_body_chars]

    recipient_names: list[str] = []
    try:
        for recipient in item.Recipients:
            if recipient.Name:
                recipient_names.append(recipient.Name)
    except Exception:
        pass

    return MailInfo(
        entry_id=item.EntryID,
        subject=item.Subject or "",
        sender_name=item.SenderName or "",
        sender_email=sender_email,
        received=received_naive,
        importance=item.Importance,
        flagged=bool(item.FlagRequest) if hasattr(item, "FlagRequest") else False,
        recipient_type=_recipient_type_for_current_user(item, current_user_address),
        recipients=recipient_names,
        body_preview=body,
        categories=item.Categories or "",
        item=item,
        is_meeting=is_meeting,
        meeting_type=meeting_type,
    )


def _iter_folders_recursive(folder, include_subfolders: bool):
    yield folder
    if include_subfolders:
        try:
            for sub in folder.Folders:
                yield from _iter_folders_recursive(sub, include_subfolders)
        except Exception:
            pass


def _get_target_inbox_folders(namespace, scan_all_accounts: bool, include_subfolders: bool):
    """
    Yield Outlook Folder COM objects to scan for mail: the default Inbox
    (plus its subfolders if `include_subfolders` - e.g. ones an Outlook rule
    auto-files mail into), and optionally the Inbox of every other mail
    account/store on this Outlook profile (secondary accounts, shared
    mailboxes) if `scan_all_accounts`. De-duplicates by folder EntryID so the
    default account isn't scanned twice.
    """
    seen_entry_ids: set[str] = set()

    def _emit(folder):
        try:
            entry_id = folder.EntryID
        except Exception:
            entry_id = None
        if entry_id:
            if entry_id in seen_entry_ids:
                return
            seen_entry_ids.add(entry_id)
        yield from _iter_folders_recursive(folder, include_subfolders)

    yield from _emit(namespace.GetDefaultFolder(OL_FOLDER_INBOX))

    if scan_all_accounts:
        for root_folder in namespace.Folders:
            try:
                inbox = root_folder.Folders["Inbox"]
            except Exception:
                continue  # this account/store has no Inbox subfolder (e.g. a public folder)
            yield from _emit(inbox)


def get_unread_emails(
    days_lookback: int = 14,
    max_body_chars: int = 500,
    max_results: int = 200,
    scan_all_accounts: bool = True,
    include_subfolders: bool = True,
) -> list[MailInfo]:
    """
    Return unread emails, newest first, within `days_lookback` days, capped at
    `max_results` (default 200 most recent) so a huge unread backlog doesn't
    overload the scoring/summary step.

    By default this scans every account/mailbox on the Outlook profile
    (`scan_all_accounts`) and every subfolder under each Inbox
    (`include_subfolders`) - so mail an Outlook rule has filed into a
    subfolder, or mail in a secondary/shared mailbox, is included too. Set
    either to False (via config.json) to restrict back to just the default
    Inbox.
    """
    _outlook, namespace = _connect()
    current_user_address = _current_user_address(namespace)
    cutoff = dt.datetime.now() - dt.timedelta(days=days_lookback)

    candidates: list[tuple[dt.datetime, Any]] = []
    for folder in _get_target_inbox_folders(namespace, scan_all_accounts, include_subfolders):
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # newest first
            unread = items.Restrict("[Unread] = true")
        except Exception:
            continue  # folder doesn't support Items/Restrict (e.g. a non-mail folder)

        for item in unread:
            try:
                received = item.ReceivedTime
                received_naive = dt.datetime(received.year, received.month, received.day,
                                              received.hour, received.minute, received.second)
                if received_naive < cutoff:
                    continue
                candidates.append((received_naive, item))
            except Exception:
                continue

    candidates.sort(key=lambda pair: pair[0], reverse=True)

    results: list[MailInfo] = []
    skipped_non_mail = 0
    for _received, item in candidates:
        if len(results) >= max_results:
            break
        try:
            mail_info = _parse_mail_item(item, current_user_address, max_body_chars)
            if mail_info is not None:
                results.append(mail_info)
            else:
                skipped_non_mail += 1  # e.g. a meeting response (accept/decline/tentative), task, or other non-mail item
        except Exception:
            # Skip items that fail to parse (e.g. non-standard message classes)
            continue

    if skipped_non_mail:
        print(
            f"Note: {skipped_non_mail} unread item(s) were meeting responses (accept/decline/"
            f"tentative) or other non-email items, not scored by this tool - handle those "
            f"directly in Outlook. Meeting requests/cancellations *are* scored. This is why the "
            f"count here may be lower than Outlook's unread badge."
        )

    return results


def get_recent_emails(
    limit: int = 100,
    max_body_chars: int = 500,
    scan_all_accounts: bool = True,
    include_subfolders: bool = True,
) -> list[MailInfo]:
    """
    Return the most recent `limit` emails, newest first, regardless of
    read/unread status. Intended for testing the scorer against a broader,
    realistic sample rather than only today's unread mail. Same
    all-accounts/all-subfolders coverage as get_unread_emails() by default.
    """
    _outlook, namespace = _connect()
    current_user_address = _current_user_address(namespace)

    candidates: list[tuple[dt.datetime, Any]] = []
    for folder in _get_target_inbox_folders(namespace, scan_all_accounts, include_subfolders):
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # newest first
        except Exception:
            continue

        count = 0
        for item in items:
            if count >= limit:  # per-folder cap; the real cap is applied after merging below
                break
            try:
                received = item.ReceivedTime
                received_naive = dt.datetime(received.year, received.month, received.day,
                                              received.hour, received.minute, received.second)
                candidates.append((received_naive, item))
                count += 1
            except Exception:
                continue

    candidates.sort(key=lambda pair: pair[0], reverse=True)

    results: list[MailInfo] = []
    for _received, item in candidates:
        if len(results) >= limit:
            break
        try:
            mail_info = _parse_mail_item(item, current_user_address, max_body_chars)
            if mail_info is not None:
                results.append(mail_info)
        except Exception:
            # Skip items that fail to parse (e.g. non-standard message classes)
            continue

    return results


def set_category(mail_item, category: str) -> None:
    """Append a category label to a mail item and save it."""
    existing = mail_item.Categories or ""
    labels = [c.strip() for c in existing.split(",") if c.strip()]
    if category not in labels:
        labels.append(category)
    mail_item.Categories = ", ".join(labels)
    mail_item.Save()


def get_mail_item_by_entry_id(entry_id: str):
    """Reconnect to a live Outlook COM MailItem by its EntryID. Useful when a
    mail was fetched/assessed in one process and needs to be re-opened
    (e.g. via .Display()) from a separate process/run."""
    _outlook, namespace = _connect()
    return namespace.GetItemFromID(entry_id)
