"""
scorers.py
Pluggable priority scorers.

RuleBasedScorer is the only scorer here: deterministic, config-driven, no
external calls, zero setup. It's also used to produce `detect_signals()` -
human-readable hints handed to whichever LLM does the actual subjective
scoring.

This product's one intended AI path is GitHub Copilot Chat in VS Code, which
runs main.py's --export-json/--gui-from-json CLI flags itself (see
.github/copilot-instructions.md) - Copilot reads the raw mail + these
rule-based hints and assigns its own subjective score/label/reasons/summary.
There is no in-process LLM API call here by design, so this file has no
external dependencies beyond the stdlib.
"""
from __future__ import annotations

import re
from abc import ABC, abstractmethod
from dataclasses import dataclass

from outlook_client import MailInfo

HIGH_THRESHOLD = 40
MEDIUM_THRESHOLD = 15


@dataclass
class ScoreResult:
    score: int
    label: str  # "High", "Medium", "Low", "Automated", or "External"
    reasons: list[str]
    summary: str = ""  # short human-readable summary of the email's content


AUTOMATED_LABEL = "Automated"
EXTERNAL_LABEL = "External"


def _label_for_score(score: int) -> str:
    if score >= HIGH_THRESHOLD:
        return "High"
    if score >= MEDIUM_THRESHOLD:
        return "Medium"
    return "Low"


_QUOTE_HEADER_RE = re.compile(r"\n\s*(From:|Sent:|To:|Cc:|Subject:)\s", re.IGNORECASE)


_NAME_TOKEN_RE = re.compile(r"[a-z0-9]+")


def _name_tokens(text: str) -> set[str]:
    """Lowercase word tokens, stripping punctuation - used so name matching
    doesn't care about word order or separators (commas, periods, ...)."""
    return set(_NAME_TOKEN_RE.findall(text.lower()))


def _sender_matches_entry(entry: str, sender_email: str, sender_name: str) -> bool:
    """
    True if a configured sender entry (from priority_senders/low_priority_senders)
    matches this mail's sender. An entry containing '@' must match the sender's
    email exactly; otherwise it's treated as a name and matched by token subset
    rather than a literal substring - Outlook often displays names as
    "Last, First Middle" (e.g. "Dnv, Sarath B"), which a naive substring check
    against a naturally-typed "Sarath DNV" would miss entirely since the word
    order is reversed.
    """
    entry = entry.strip().lower()
    if not entry:
        return False
    if "@" in entry:
        return entry == sender_email
    entry_tokens = _name_tokens(entry)
    if not entry_tokens:
        return False
    return entry_tokens.issubset(_name_tokens(sender_name))


def _heuristic_summary(body_preview: str, max_chars: int = 160) -> str:
    """
    Cheap, no-LLM extractive summary: drop quoted reply-chain history (the
    'From:/Sent:/To:' block Outlook prepends to replies), collapse whitespace,
    and truncate to a word boundary. Used by RuleBasedScorer, which has no
    model to generate a real abstractive summary.
    """
    text = body_preview or ""
    match = _QUOTE_HEADER_RE.search(text)
    if match:
        text = text[: match.start()]
    text = " ".join(text.split())  # collapse all whitespace/newlines
    if len(text) <= max_chars:
        return text
    truncated = text[:max_chars].rsplit(" ", 1)[0]
    return truncated + "..."


class BaseScorer(ABC):
    @abstractmethod
    def score(self, mail: MailInfo) -> ScoreResult:
        ...


class RuleBasedScorer(BaseScorer):
    def __init__(self, config: dict):
        self.vip_domains = {d.lower() for d in config.get("vip_domains", [])}
        self.urgent_keywords = [k.lower() for k in config.get("urgent_keywords", [])]
        self.low_priority_keywords = [k.lower() for k in config.get("low_priority_keywords", [])]
        self.my_names = [n.strip() for n in config.get("my_names", []) if n.strip()]
        self.my_email = (config.get("my_email") or "").strip().lower()
        # Each entry may be an email address or a (sub)string of a display name.
        self.priority_senders = [p.lower() for p in config.get("priority_senders", [])]
        # Specific people (not automated systems) whose mail should always be
        # deprioritized - distinct from automated_senders (machine/system
        # notifications, detected on their own, never user-defined) and from
        # low_priority_keywords (content-based rather than sender-based).
        self.low_priority_senders = [p.lower() for p in config.get("low_priority_senders", [])]
        self.trusted_domains = {d.lower() for d in config.get("trusted_domains", [])}
        # Substrings matched against sender email/name that mark a mail as a machine-
        # generated notification (bug trackers, Jira, system alerts, ...). These are
        # triaged as their own "Automated" category, never High/Medium/Low - a bug
        # tracker update mentioning "Gen6"/"urgent" is not the same as a person asking
        # you for something, no matter how many keywords it happens to match.
        self.automated_senders = [s.lower() for s in config.get("automated_senders", [])]
        self.weights = config.get("weights", {})

        # Generic mechanism for any rule of the form "this text near one of my names
        # matters". Each entry is {template, weight, reason}; {name} in the template
        # is expanded to an alternation of all my_names. Add new name-based rules here
        # without touching any code.
        name_alternation = "|".join(re.escape(n) for n in self.my_names)
        self._name_patterns: list[tuple[re.Pattern, int, str]] = []
        if name_alternation:
            for entry in config.get("name_patterns", []):
                template = entry.get("template", "")
                if not template:
                    continue
                pattern_str = template.replace("{name}", f"(?:{name_alternation})")
                self._name_patterns.append((
                    re.compile(pattern_str, re.IGNORECASE),
                    entry.get("weight", 0),
                    entry.get("reason", "Name pattern matched"),
                ))

    def _evaluate(self, mail: MailInfo) -> list[tuple[int, str]]:
        """
        Evaluate every config-driven rule against a mail and return the list of
        (weight, reason) pairs that matched. This is the single source of truth
        for what the config *suggests* about a mail - deterministic scoring
        (score()) sums these; LLM-backed scorers instead pass the reasons along
        as hints for a subjective judgment call (detect_signals()).
        """
        signals: list[tuple[int, str]] = []
        text = f"{mail.subject} {mail.body_preview}".lower()
        sender_email = (mail.sender_email or "").lower()
        sender_domain = sender_email.split("@")[-1] if "@" in sender_email else ""

        if sender_domain in self.vip_domains:
            signals.append((25, "VIP domain"))

        # Rule 1: any configured name pattern (mentions, greetings, tags, ...) found
        # in the subject/body
        search_text = f"{mail.subject} {mail.body_preview}"
        for pattern, weight, reason in self._name_patterns:
            if pattern.search(search_text):
                signals.append((weight, reason))

        # Rule 1b: your own email address appearing in the body (e.g. explicitly
        # CC'd/forwarded inline, or quoted in a "To:" line) also implies you're
        # being addressed, same as a name mention.
        if self.my_email and self.my_email in text:
            signals.append((25, "Addressed by your email address in the body"))

        # Rule 2: mails from outside trusted domains (e.g. intel.com, incl. subdomains) lose points
        is_trusted = any(
            sender_domain == d or sender_domain.endswith(f".{d}")
            for d in self.trusted_domains
        )
        if self.trusted_domains and sender_domain and not is_trusted:
            signals.append((-self.weights.get("external_domain_penalty", 30), f"External domain ({sender_domain})"))

        # Rule 3: sender (matched by email or by name) is on the priority senders list
        is_priority_sender = any(
            _sender_matches_entry(entry, sender_email, mail.sender_name)
            for entry in self.priority_senders
        )
        if is_priority_sender:
            signals.append((self.weights.get("priority_sender", 35), "Priority sender"))

        # Rule 3b: sender is a specific person you've flagged as always lower priority
        # (distinct from automated/system senders, which are detected separately).
        is_low_priority_sender = any(
            _sender_matches_entry(entry, sender_email, mail.sender_name)
            for entry in self.low_priority_senders
        )
        if is_low_priority_sender:
            signals.append((-self.weights.get("low_priority_sender", 25), "Low-priority sender (per your settings)"))

        if mail.importance == 2:
            signals.append((20, "Marked High importance"))
        elif mail.importance == 0:
            signals.append((-5, "Marked Low importance"))

        if mail.recipient_type == "To":
            signals.append((10, "Directly addressed (To)"))
        elif mail.recipient_type == "CC":
            signals.append((-5, "CC only"))

        if mail.flagged:
            signals.append((15, "Flagged for follow-up"))

        matched_urgent = [k for k in self.urgent_keywords if k in text]
        if matched_urgent:
            bonus = min(15 * len(matched_urgent), 30)
            signals.append((bonus, f"Urgent keyword(s): {', '.join(matched_urgent[:3])}"))

        matched_low = [k for k in self.low_priority_keywords if k in text]
        if matched_low:
            signals.append((-20, f"Low-priority keyword(s): {', '.join(matched_low[:3])}"))

        return signals

    def is_automated_sender(self, mail: MailInfo) -> bool:
        """True if the sender matches a known automated/system notification source
        (bug trackers, Jira, IT alerts, ...) per config's automated_senders list."""
        sender_email = (mail.sender_email or "").lower()
        sender_name_lower = (mail.sender_name or "").lower()
        return any(
            entry in sender_email or entry in sender_name_lower
            for entry in self.automated_senders
        )

    def is_external_sender(self, mail: MailInfo) -> bool:
        """True if the sender's domain is outside config's trusted_domains
        (e.g. intel.com), including subdomains."""
        sender_email = (mail.sender_email or "").lower()
        sender_domain = sender_email.split("@")[-1] if "@" in sender_email else ""
        if not sender_domain or not self.trusted_domains:
            return False
        is_trusted = any(
            sender_domain == d or sender_domain.endswith(f".{d}")
            for d in self.trusted_domains
        )
        return not is_trusted

    def detect_signals(self, mail: MailInfo) -> list[str]:
        """
        Human-readable reasons for whatever config rules matched this mail,
        with no weights/scoring math applied. Intended to be handed to an LLM
        as a *starting point* / set of hints - the LLM decides how much (if
        at all) each one should matter, rather than these being summed into
        the final score.
        """
        reasons = [reason for _, reason in self._evaluate(mail)]
        if self.is_automated_sender(mail):
            reasons.insert(0, "Automated notification sender (bug tracker/Jira/system alert)")
        return reasons

    def score(self, mail: MailInfo) -> ScoreResult:
        signals = self._evaluate(mail)
        score = sum(weight for weight, _ in signals)
        reasons = [reason for _, reason in signals]
        summary = _heuristic_summary(mail.body_preview)

        # Precedence: Automated (machine-generated) > External (outside trusted
        # domains) > the normal score-based High/Medium/Low label.
        if self.is_automated_sender(mail):
            reasons.insert(0, "Automated notification sender (bug tracker/Jira/system alert)")
            return ScoreResult(score=score, label=AUTOMATED_LABEL, reasons=reasons, summary=summary)

        if self.is_external_sender(mail):
            return ScoreResult(score=score, label=EXTERNAL_LABEL, reasons=reasons, summary=summary)

        return ScoreResult(score=score, label=_label_for_score(score), reasons=reasons, summary=summary)

