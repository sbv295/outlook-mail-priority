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
    if item.Class != 43:  # olMail
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
    )


def get_unread_emails(days_lookback: int = 14, max_body_chars: int = 500) -> list[MailInfo]:
    """Return unread emails from the default Inbox, newest first."""
    _outlook, namespace = _connect()
    inbox = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    current_user_address = _current_user_address(namespace)

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # newest first
    unread = items.Restrict("[Unread] = true")

    cutoff = dt.datetime.now() - dt.timedelta(days=days_lookback)
    results: list[MailInfo] = []

    for item in unread:
        try:
            received = item.ReceivedTime
            received_naive = dt.datetime(received.year, received.month, received.day,
                                          received.hour, received.minute, received.second)
            if received_naive < cutoff:
                continue
            mail_info = _parse_mail_item(item, current_user_address, max_body_chars)
            if mail_info is not None:
                results.append(mail_info)
        except Exception:
            # Skip items that fail to parse (e.g. non-standard message classes)
            continue

    return results


def get_recent_emails(limit: int = 100, max_body_chars: int = 500) -> list[MailInfo]:
    """
    Return the most recent `limit` emails from the default Inbox, newest first,
    regardless of read/unread status. Intended for testing the scorer against a
    broader, realistic sample rather than only today's unread mail.
    """
    _outlook, namespace = _connect()
    inbox = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    current_user_address = _current_user_address(namespace)

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # newest first

    results: list[MailInfo] = []
    for item in items:
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
