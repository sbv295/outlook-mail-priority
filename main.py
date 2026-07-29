"""
main.py
Fetch unread Outlook emails, score them by priority, report results,
and optionally tag them back into Outlook as categories.

RuleBasedScorer (deterministic, config-driven, no external calls) is the only
scorer available directly from this CLI. This product's one intended AI path
is GitHub Copilot Chat in VS Code (see .github/copilot-instructions.md) -
Copilot runs this CLI's --export-json/--gui-from-json flags itself and reads
raw mail + this file's deterministic hints to assign its own subjective
score/label/reasons/summary.

Usage:
  python main.py                     # rule-based scoring, prints table, writes CSV
  python main.py --tag               # also write "Priority-High/Medium/Low" category onto each email
  python main.py --top 10            # only show/report the top 10 by score
  python main.py --recent 20 --export-json dump.json   # no scoring at all - just dump raw mail
                                                        # data to a JSON file for ad-hoc review
  python main.py --gui                                  # also pop up a categorized GUI window;
                                                        # double-click an email to open it in Outlook
  python main.py --gui-from-json dump.json              # pop up the GUI using priority_score/
                                                        # priority_label/priority_reasons fields
                                                        # added to a previously exported JSON file
  python main.py --setup                                # re-open the guided setup wizard for
                                                        # config.json/priority_profile.md at any time
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import sys
from pathlib import Path

from outlook_client import get_recent_emails, get_unread_emails, set_category
from scorers import RuleBasedScorer, ScoreResult

APP_DIR = Path(__file__).parent
USER_DATA_DIR = APP_DIR / "user_data"
USER_DATA_DIR.mkdir(exist_ok=True)

CONFIG_PATH = USER_DATA_DIR / "config.json"
DEFAULT_CSV = USER_DATA_DIR / "outlook_priority_report.csv"
PROFILE_PATH = USER_DATA_DIR / "priority_profile.md"


def load_config() -> dict:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def load_profile() -> str:
    if PROFILE_PATH.exists():
        return PROFILE_PATH.read_text(encoding="utf-8")
    return ""


def build_scorer(config: dict):
    return RuleBasedScorer(config)


def _write_csv(path: Path, scored: list) -> None:
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Score", "Label", "Received", "SenderName", "SenderEmail",
                          "RecipientType", "Subject", "Summary", "Reasons"])
        for mail, result in scored:
            writer.writerow([
                result.score, result.label, mail.received.isoformat(),
                mail.sender_name, mail.sender_email, mail.recipient_type,
                mail.subject, result.summary, "; ".join(result.reasons),
            ])


def _export_mails_json(path: Path, mails: list, config: dict) -> None:
    """
    Dump raw mail fields (no final scoring) so they can be reviewed ad hoc, e.g.
    by reading this file directly rather than calling an LLM API. Each entry
    also includes 'rule_hints' - RuleBasedScorer.detect_signals() output,
    computed once here with the same tested, deterministic logic as the CLI's
    own scoring path (correct sender-name matching regardless of Outlook's
    "Last, First" display order, automated-sender detection, etc.). The point
    is to remove any need for whoever/whatever reads this file to re-derive
    those checks by eye - the fact is already computed and sitting in the data.
    """
    scorer = RuleBasedScorer(config)
    data = [
        {
            "entry_id": mail.entry_id,
            "received": mail.received.isoformat(),
            "sender_name": mail.sender_name,
            "sender_email": mail.sender_email,
            "recipient_type": mail.recipient_type,
            "subject": mail.subject,
            "body_preview": mail.body_preview,
            "flagged": mail.flagged,
            "importance": mail.importance,
            "categories": mail.categories,
            "rule_hints": scorer.detect_signals(mail),
        }
        for mail in mails
    ]
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def _load_scored_from_json(path: Path) -> list:
    """
    Load a JSON file previously written by --export-json, optionally augmented
    with 'priority_score' (int), 'priority_label' (str), 'priority_reasons'
    (list[str]), and 'priority_summary' (str) fields per entry (e.g. added by
    an assistant after reading and assessing the raw dump). Reconnects each
    entry to its live Outlook item by entry_id so the GUI's click-to-open
    still works.
    """
    from outlook_client import MailInfo, get_mail_item_by_entry_id
    from scorers import _heuristic_summary

    with open(path, "r", encoding="utf-8-sig") as f:
        entries = json.load(f)

    scored = []
    for entry in entries:
        entry_id = entry.get("entry_id")
        if not entry_id:
            print(f"Skipping entry with no entry_id: {entry.get('subject')!r}", file=sys.stderr)
            continue
        try:
            item = get_mail_item_by_entry_id(entry_id)
        except Exception as exc:
            print(f"Could not reconnect to '{entry.get('subject')!r}': {exc}", file=sys.stderr)
            continue
        mail = MailInfo(
            entry_id=entry_id,
            subject=entry.get("subject", ""),
            sender_name=entry.get("sender_name", ""),
            sender_email=entry.get("sender_email", ""),
            received=dt.datetime.fromisoformat(entry["received"]),
            importance=entry.get("importance", 1),
            flagged=entry.get("flagged", False),
            recipient_type=entry.get("recipient_type", "Unknown"),
            recipients=[],
            body_preview=entry.get("body_preview", ""),
            categories=entry.get("categories", ""),
            item=item,
        )
        result = ScoreResult(
            score=entry.get("priority_score", 0),
            label=entry.get("priority_label", "Low"),
            reasons=entry.get("priority_reasons", []),
            summary=entry.get("priority_summary") or _heuristic_summary(entry.get("body_preview", "")),
        )
        scored.append((mail, result))

    scored.sort(key=lambda pair: pair[1].score, reverse=True)
    return scored


def _maybe_prompt_for_updates() -> None:
    """Rate-limited (see updater.CHECK_INTERVAL_SECONDS), silent no-op if not
    a git checkout / no remote / offline - only surfaces a dialog when there's
    a genuine update and it's been a while since the last check."""
    try:
        from updater import check_for_updates, prompt_and_apply
        info = check_for_updates(force=False)
        if info:
            prompt_and_apply(info)
    except Exception:
        pass  # update-checking must never block the actual priority check


def main() -> int:
    parser = argparse.ArgumentParser(description="Prioritize unread Outlook emails.")
    parser.add_argument("--tag", action="store_true", help="Write Priority-* category onto each email in Outlook")
    parser.add_argument("--top", type=int, default=None, help="Only process/report the top N by score")
    parser.add_argument("--csv", type=Path, default=DEFAULT_CSV, help="Output CSV path")
    parser.add_argument("--recent", type=int, default=None,
                         help="Testing: score the N most recent inbox emails (read or unread) "
                              "instead of only unread mail")
    parser.add_argument("--export-json", type=Path, default=None,
                         help="Skip scoring entirely; just dump fetched mail data to this JSON path "
                              "for ad-hoc review (e.g. by an assistant reading the file)")
    parser.add_argument("--gui", action="store_true",
                         help="Show a popup window with mail grouped by priority; "
                              "double-click an email to open it in Outlook")
    parser.add_argument("--gui-from-json", type=Path, default=None,
                         help="Skip fetching/scoring; show the GUI using priority_score/priority_label/"
                              "priority_reasons fields from a previously exported JSON file")
    parser.add_argument("--setup", action="store_true",
                         help="Open the guided setup wizard for config.json/priority_profile.md, then exit")
    parser.add_argument("--check-updates", action="store_true",
                         help="Check GitHub for updates now (bypassing the normal rate limit) and "
                              "prompt to pull them, then exit")
    args = parser.parse_args()

    from setup_wizard import is_first_run, run_wizard

    if args.setup:
        run_wizard()
        return 0

    if args.check_updates:
        from updater import check_for_updates, prompt_and_apply
        info = check_for_updates(force=True)
        if info:
            prompt_and_apply(info)
        else:
            print("No updates available (or not a git checkout / no remote / offline).")
        return 0

    if is_first_run():
        print("First run detected - opening the guided setup wizard...")
        run_wizard()

    config = load_config()

    if args.gui_from_json:
        scored = _load_scored_from_json(args.gui_from_json)
        _maybe_prompt_for_updates()
        from gui import show_priority_popup
        show_priority_popup(scored)
        return 0

    if args.export_json:
        if args.recent:
            print(f"Connecting to Outlook and fetching the {args.recent} most recent email(s)...")
            mails = get_recent_emails(
                limit=args.recent,
                max_body_chars=config.get("max_body_chars", 500),
                scan_all_accounts=config.get("scan_all_accounts", True),
                include_subfolders=config.get("include_subfolders", True),
            )
        else:
            print("Connecting to Outlook and fetching unread mail...")
            mails = get_unread_emails(
                days_lookback=config.get("days_lookback", 14),
                max_body_chars=config.get("max_body_chars", 500),
                max_results=config.get("max_unread_results", 200),
                scan_all_accounts=config.get("scan_all_accounts", True),
                include_subfolders=config.get("include_subfolders", True),
            )
        _export_mails_json(args.export_json, mails, config)
        print(f"Wrote {len(mails)} email(s) to {args.export_json}")
        return 0

    try:
        scorer = build_scorer(config)
    except RuntimeError as exc:
        print(f"Could not initialize the rule-based scorer: {exc}", file=sys.stderr)
        return 1

    if args.recent:
        print(f"Connecting to Outlook and fetching the {args.recent} most recent email(s)...")
        mails = get_recent_emails(
            limit=args.recent,
            max_body_chars=config.get("max_body_chars", 500),
            scan_all_accounts=config.get("scan_all_accounts", True),
            include_subfolders=config.get("include_subfolders", True),
        )
        print(f"Found {len(mails)} email(s).")
    else:
        print("Connecting to Outlook and fetching unread mail...")
        mails = get_unread_emails(
            days_lookback=config.get("days_lookback", 14),
            max_body_chars=config.get("max_body_chars", 500),
            max_results=config.get("max_unread_results", 200),
            scan_all_accounts=config.get("scan_all_accounts", True),
            include_subfolders=config.get("include_subfolders", True),
        )
        print(f"Found {len(mails)} unread email(s) within lookback window "
              f"(capped at {config.get('max_unread_results', 200)} most recent).")

    scored = []
    for mail in mails:
        result = scorer.score(mail)
        scored.append((mail, result))

    scored.sort(key=lambda pair: pair[1].score, reverse=True)
    if args.top:
        scored = scored[: args.top]

    header = f"{'Score':>5} {'Label':<7} {'Received':<17} {'Sender':<30} Subject"
    print(header)
    print("-" * len(header))
    for mail, result in scored:
        print(
            f"{result.score:>5} {result.label:<7} "
            f"{mail.received.strftime('%Y-%m-%d %H:%M'):<17} "
            f"{mail.sender_name[:29]:<30} {mail.subject[:60]}"
        )

    csv_path = args.csv
    try:
        _write_csv(csv_path, scored)
    except PermissionError:
        csv_path = csv_path.with_stem(csv_path.stem + "_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S"))
        print(f"'{args.csv}' is locked (likely open elsewhere). Writing to {csv_path} instead.", file=sys.stderr)
        _write_csv(csv_path, scored)
    print(f"\nCSV report written to {csv_path}")

    if args.tag:
        print("Tagging emails in Outlook with Priority categories...")
        for mail, result in scored:
            set_category(mail.item, f"Priority-{result.label}")
        print("Done tagging.")

    if args.gui:
        _maybe_prompt_for_updates()
        from gui import show_priority_popup
        show_priority_popup(scored)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
