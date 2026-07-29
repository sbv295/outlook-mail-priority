"""
setup_wizard.py
Guided, GUI-based onboarding for config.json + priority_profile.md - so a
teammate never has to hand-edit either file. Runs automatically on first use
(see main.py) and is also reachable any time via the "Customize" button in
gui.py's title bar, or `python main.py --setup` / `python setup_wizard.py`.

Two entry modes:
  run_wizard()               - standalone: creates its own Tk root, blocks
                                until the window is closed (used by main.py
                                before the very first fetch/GUI).
  run_wizard(master=<Tk>)    - embedded: opens as a modal Toplevel on top of
                                an already-running Tk app (used by gui.py's
                                "Customize" button), and returns control to
                                the caller once closed.

Returns True if the user saved changes, False if they cancelled/skipped.
"""
from __future__ import annotations

import json
import re
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

APP_DIR = Path(__file__).parent
USER_DATA_DIR = APP_DIR / "user_data"
USER_DATA_DIR.mkdir(exist_ok=True)

CONFIG_PATH = USER_DATA_DIR / "config.json"
CONFIG_DEFAULT_PATH = APP_DIR / "config.default.json"
PROFILE_PATH = USER_DATA_DIR / "priority_profile.md"
PROFILE_DEFAULT_PATH = APP_DIR / "priority_profile.default.md"

_BG = "#faf9f8"
_PANEL_BG = "#ffffff"
_BORDER = "#e1dfdd"
_HEADER_BG = "#ffffff"
_HEADER_FG = "#201f1e"
_HEADER_SUBFG = "#605e5c"
_TEXT = "#201f1e"
_TEXT_GRAY = "#605e5c"
_ACCENT = "#0078d4"
_ACCENT_FG = "#ffffff"

_FONT_TITLE = ("Segoe UI", 13, "bold")
_FONT_SUBTITLE = ("Segoe UI", 9)
_FONT_SECTION = ("Segoe UI", 10, "bold")
_FONT_LABEL = ("Segoe UI", 9)
_FONT_HINT = ("Segoe UI", 8)
_FONT_TEXT = ("Segoe UI", 10)
_FONT_BUTTON = ("Segoe UI", 10, "bold")

# Fixed, non-personal boilerplate re-used verbatim from priority_profile.default.md
# so a wizard-generated profile stays consistent with the shipped documentation.
_CORE_PRINCIPLE = """The central question for High/Medium/Low is: **does this email need MY attention or action?**
Not "is this topic important in general" or "does this mention a project I care about."
- If I need to read it, decide something, respond, or act soon \u2192 High/Medium (depending on urgency/deadline).
- If it's FYI, background noise, or something I can skim/ignore without consequence \u2192 Low, even if it's
  about a project I care about or from someone on my priority list.
- A thread I'm only passively CC'd on, with no question or action directed at me, is still Low.
  Project keywords are a hint, not a reason to inflate priority on their own.
- Conversely, a direct question, ask, blocker, or deadline addressed to me is High/Medium even on a topic
  that isn't one of my "current focus" areas."""

_AUTOMATED_BOILERPLATE = """Machine-generated notifications should be classified as **"Automated"**, a separate category from High/Medium/Low, regardless of how urgent their content sounds or how many project keywords they contain. A bug-tracker or system email that merely mentions urgent-sounding keywords is not the same as a person asking me to do something.
- See `automated_senders` in config.json for the exact matched sender patterns \u2014 add new ones there (or via the setup wizard) as you notice more automated senders, rather than trying to keyword-detect "this looks like a bot."
- **This list will never be complete.** If you're an LLM-backed scorer (e.g. GitHub Copilot triaging this inbox) and a sender isn't on the list, still use judgment: templated/ticket-system formatting, "unsubscribe"/"do not reply" boilerplate, no personalized greeting, or a system/service name as the sender (rather than a person) are all signs of an automated notification even without a config match. Conversely, a sender matching the list isn't an absolute guarantee either. The plain rule-based scorer (no LLM) can only go by the config list, so it will miss unlisted automated senders \u2014 that's an accepted limitation of the free path."""

_EXTERNAL_BOILERPLATE = """Any mail from a sender whose domain is outside `trusted_domains` in config.json should be labeled **"External"**, regardless of how urgent or relevant the content is \u2014 vendor replies, marketing, conference/webinar promos, etc. all land here so they're triaged separately from internal traffic.
- **Precedence: Automated > External > High/Medium/Low.** If a mail is both from an automated system *and* an external domain, it's "Automated", not "External" \u2014 the automated/system-notification distinction is more specific and useful than just "external."
- Genuine external collaborators with a direct, substantive ask can still be flagged as such in `priority_reasons` even while labeled "External" \u2014 the category is about triage grouping, not about ignoring the content."""

_SECTION_RE_TEMPLATE = r"##\s+{header}\s*\n(.*?)(?=\n##\s|\Z)"


def _split_list(raw: str) -> list[str]:
    """Split free-typed text on commas AND newlines, strip, dedupe (keep order), drop empties."""
    parts = re.split(r"[,\n]", raw)
    seen = set()
    out = []
    for p in parts:
        p = p.strip()
        if p and p not in seen:
            seen.add(p)
            out.append(p)
    return out


def _load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _extract_section(text: str, header: str) -> str:
    """Pull the placeholder-stripped body of a '## <header>' section out of a
    previously-generated priority_profile.md, for pre-filling the wizard on
    re-customization. Returns '' if not found or still just a placeholder."""
    pattern = _SECTION_RE_TEMPLATE.format(header=re.escape(header))
    match = re.search(pattern, text, re.DOTALL)
    if not match:
        return ""
    body = match.group(1).strip()
    if body.startswith("- (") or body.startswith("(") or not body:
        return ""
    # Strip a single leading "- " per line (the wizard writes bullets)
    lines = [re.sub(r"^- ?", "", line) for line in body.splitlines()]
    return "\n".join(lines).strip()


def _build_profile_markdown(focus: str, priority_senders: list[str], deprioritize: str, situational: str) -> str:
    def _bulleted(text: str, placeholder: str) -> str:
        text = text.strip()
        if not text:
            return f"- {placeholder}"
        return "\n".join(f"- {line.strip()}" for line in text.splitlines() if line.strip())

    people_section = (
        "\n".join(f"- {name}" for name in priority_senders)
        if priority_senders
        else "- (Add priority senders in the wizard's \"Priority senders\" section and they'll be listed here automatically.)"
    )

    return f"""# Priority Profile
<!-- Generated by the setup wizard. This file is context for GitHub Copilot when it triages
     mail, not meant to be read/edited by hand - use the "Customize" button (or
     `python main.py --setup`) to change it. -->

## Core principle
{_CORE_PRINCIPLE}

## Current focus / active projects
{_bulleted(focus, "(No active projects specified yet - re-run the setup wizard to add some.)")}

## People whose emails matter more than usual right now
{people_section}

## Automated notifications \u2014 always their own category, never High/Medium/Low
{_AUTOMATED_BOILERPLATE}

## External senders \u2014 a separate category from High/Medium/Low/Automated
{_EXTERNAL_BOILERPLATE}

## Things to actively deprioritize
{_bulleted(deprioritize, "(Nothing specified yet - re-run the setup wizard to add some.)")}

## Situational notes
{_bulleted(situational, "(None right now.)")}
"""


def is_first_run() -> bool:
    return not CONFIG_PATH.exists() or not PROFILE_PATH.exists()


def _detect_outlook_identity() -> dict:
    try:
        from outlook_client import get_current_user_info
        return get_current_user_info()
    except Exception:
        return {}


def _prompt_text(master: tk.Misc, title: str, label: str) -> str | None:
    """Small modal popup with a single Entry - used by the 'Add' buttons on the
    priority/low-priority sender list editors. Returns the typed value, or
    None if cancelled/empty."""
    popup = tk.Toplevel(master)
    popup.title(title)
    popup.configure(bg=_PANEL_BG)
    popup.transient(master.winfo_toplevel())
    popup.resizable(False, False)
    popup.grab_set()

    tk.Label(
        popup, text=label, bg=_PANEL_BG, fg=_TEXT, font=_FONT_LABEL, anchor="w",
    ).pack(fill="x", padx=16, pady=(16, 4))
    var = tk.StringVar()
    entry = tk.Entry(
        popup, textvariable=var, font=_FONT_TEXT, width=36, relief="solid", borderwidth=1,
        highlightthickness=1, highlightbackground=_BORDER,
    )
    entry.pack(fill="x", padx=16, pady=(0, 12), ipady=4)
    entry.focus_set()

    result = {"value": None}

    def _ok(_event=None) -> None:
        if var.get().strip():
            result["value"] = var.get().strip()
        popup.destroy()

    def _cancel() -> None:
        popup.destroy()

    entry.bind("<Return>", _ok)
    popup.bind("<Escape>", lambda _e: _cancel())

    btns = tk.Frame(popup, bg=_PANEL_BG)
    btns.pack(fill="x", padx=16, pady=(0, 16))
    tk.Button(
        btns, text="Add", command=_ok, font=_FONT_BUTTON, bg=_ACCENT, fg=_ACCENT_FG,
        activebackground=_ACCENT, activeforeground=_ACCENT_FG, relief="flat", padx=14, pady=4,
        cursor="hand2",
    ).pack(side="right")
    tk.Button(
        btns, text="Cancel", command=_cancel, font=_FONT_LABEL, bg=_PANEL_BG, fg=_TEXT_GRAY,
        relief="flat", padx=10, pady=4, cursor="hand2",
    ).pack(side="right", padx=(0, 8))

    popup.wait_window()
    return result["value"]


def _list_editor(parent: tk.Widget, initial_items: list[str], popup_title: str, popup_label: str):
    """
    A Listbox of string entries plus '+ Add' (opens a small popup Entry via
    _prompt_text) and 'Remove' buttons - used for priority/low-priority
    senders so the user never has to type a comma/newline-separated blob.
    Returns a zero-arg callable that returns the current list of items.
    """
    wrapper = tk.Frame(parent, bg=_PANEL_BG)
    wrapper.pack(fill="x", padx=16, pady=(0, 4))

    listbox = tk.Listbox(
        wrapper, height=4, font=_FONT_TEXT, relief="solid", borderwidth=1,
        highlightthickness=1, highlightbackground=_BORDER, activestyle="none",
        selectbackground=_ACCENT, selectforeground="#ffffff",
    )
    for item in initial_items:
        listbox.insert("end", item)
    listbox.pack(side="left", fill="both", expand=True)

    btns = tk.Frame(wrapper, bg=_PANEL_BG)
    btns.pack(side="left", padx=(8, 0), fill="y")

    def _add() -> None:
        value = _prompt_text(wrapper, popup_title, popup_label)
        if value:
            listbox.insert("end", value)

    def _remove() -> None:
        for index in reversed(listbox.curselection()):
            listbox.delete(index)

    tk.Button(
        btns, text="+ Add", command=_add, font=_FONT_LABEL, bg=_ACCENT, fg=_ACCENT_FG,
        activebackground=_ACCENT, activeforeground=_ACCENT_FG, relief="flat", padx=10, pady=3,
        cursor="hand2",
    ).pack(fill="x", pady=(0, 4))
    tk.Button(
        btns, text="Remove", command=_remove, font=_FONT_LABEL, bg=_BG, fg=_TEXT_GRAY,
        relief="flat", padx=10, pady=3, cursor="hand2",
    ).pack(fill="x")

    return lambda: list(listbox.get(0, "end"))


def run_wizard(master: tk.Misc | None = None) -> bool:
    first_run = is_first_run()
    base_config = _load_json(CONFIG_PATH) if CONFIG_PATH.exists() else _load_json(CONFIG_DEFAULT_PATH)
    base_profile_text = PROFILE_PATH.read_text(encoding="utf-8") if PROFILE_PATH.exists() else ""

    identity = {}
    if not base_config.get("my_names") or not base_config.get("my_email"):
        identity = _detect_outlook_identity()

    saved = {"value": False}

    owns_root = master is None
    win = tk.Tk() if owns_root else tk.Toplevel(master)
    win.title("Set Up Your Priorities")
    win.geometry("760x680")
    win.minsize(600, 420)
    win.configure(bg=_BG)
    if not owns_root:
        win.transient(master)
        win.grab_set()

    style = ttk.Style(win)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure("Wizard.Vertical.TScrollbar", background=_BORDER, troughcolor=_BG, arrowsize=14)

    header = tk.Frame(win, bg=_HEADER_BG)
    header.pack(fill="x", side="top")
    tk.Label(
        header, text="Set Up Your Priorities", bg=_HEADER_BG, fg=_HEADER_FG,
        font=_FONT_TITLE, anchor="w", padx=18, pady=12,
    ).pack(side="left")
    tk.Frame(win, bg=_BORDER, height=1).pack(fill="x", side="top")
    tk.Label(
        win,
        text="Answer what's relevant to you \u2014 leave anything blank to keep the default. "
             "You can re-run this any time from the \"Customize\" button.",
        bg=_BG, fg=_TEXT_GRAY, anchor="w", justify="left", padx=16, pady=8, font=_FONT_SUBTITLE,
        wraplength=720,
    ).pack(fill="x")

    # ---- Scrollable form ----
    panel = tk.Frame(win, bg=_BORDER)
    panel.pack(fill="both", expand=True, padx=14, pady=(0, 8))
    inner = tk.Frame(panel, bg=_PANEL_BG)
    inner.pack(fill="both", expand=True, padx=1, pady=1)

    canvas = tk.Canvas(inner, bg=_PANEL_BG, highlightthickness=0)
    vsb = ttk.Scrollbar(inner, orient="vertical", command=canvas.yview, style="Wizard.Vertical.TScrollbar")
    canvas.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    content = tk.Frame(canvas, bg=_PANEL_BG)
    content_window = canvas.create_window((0, 0), window=content, anchor="nw")
    content.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: canvas.itemconfigure(content_window, width=e.width))

    def _on_mousewheel(event) -> None:
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _section(title: str) -> tk.Frame:
        tk.Label(
            content, text=title, bg=_PANEL_BG, fg=_TEXT, font=_FONT_SECTION, anchor="w",
        ).pack(fill="x", padx=16, pady=(16, 4))
        return content

    def _hint(text: str) -> None:
        tk.Label(
            content, text=text, bg=_PANEL_BG, fg=_TEXT_GRAY, font=_FONT_HINT, anchor="w",
            justify="left", wraplength=680,
        ).pack(fill="x", padx=16, pady=(0, 4))

    def _entry(initial: str = "") -> tk.Entry:
        var = tk.StringVar(value=initial)
        e = tk.Entry(content, textvariable=var, font=_FONT_TEXT, relief="solid", borderwidth=1,
                      highlightthickness=1, highlightbackground=_BORDER)
        e.pack(fill="x", padx=16, pady=(0, 4), ipady=4)
        return e

    def _text(height: int, initial: str = "") -> tk.Text:
        t = tk.Text(content, height=height, font=_FONT_TEXT, wrap="word", relief="solid",
                     borderwidth=1, highlightthickness=1, highlightbackground=_BORDER)
        t.pack(fill="x", padx=16, pady=(0, 4))
        if initial:
            t.insert("1.0", initial)
        return t

    # -- Identity --
    _section("Your identity")
    _hint("Your full name, as it appears in your email signature/directory. Any occurrence of your "
          "name or email address in a message implies you're the one being addressed.")
    full_name_entry = _entry(base_config.get("my_full_name") or identity.get("name", ""))
    _hint("Your email address.")
    email_entry = _entry(base_config.get("my_email") or identity.get("email", ""))

    # -- Priority senders --
    _section("Priority senders")
    _hint("People whose emails should always get extra priority.")
    get_priority_senders = _list_editor(
        content, base_config.get("priority_senders", []),
        "Add priority sender", "Name or email address:",
    )

    # -- Low-priority senders --
    _section("Low-priority senders")
    _hint("Specific people whose emails should always be treated as lower priority. This is separate "
          "from automated/system senders (bots, bug trackers, Jira, ...), which are detected "
          "automatically - you don't need to define those here.")
    get_low_priority_senders = _list_editor(
        content, base_config.get("low_priority_senders", []),
        "Add low-priority sender", "Name or email address:",
    )

    # -- Keywords --
    _section("Keywords")
    _hint("Urgent keywords (comma-separated) \u2014 edit freely.")
    urgent_text = _text(3, ", ".join(base_config.get("urgent_keywords", [])))
    _hint("Low-priority keywords (comma-separated) \u2014 edit freely.")
    low_text = _text(2, ", ".join(base_config.get("low_priority_keywords", [])))

    # -- Free text (feeds priority_profile.md) --
    _section("Your priorities, in your own words")
    _hint("What are your current focus areas / active projects, and when are they High vs. Medium?")
    focus_text = _text(4, _extract_section(base_profile_text, "Current focus / active projects"))
    _hint("What should always be treated as low priority for you?")
    deprioritize_text = _text(3, _extract_section(base_profile_text, "Things to actively deprioritize"))
    _hint("Any temporary/situational notes? (e.g. \"I'm out this week\")")
    situational_text = _text(2, _extract_section(base_profile_text, "Situational notes (delete when stale)"))

    tk.Frame(content, bg=_PANEL_BG, height=8).pack(fill="x")

    # ---- Buttons ----
    footer = tk.Frame(win, bg=_BG)
    footer.pack(fill="x", side="bottom", padx=14, pady=(0, 14))

    def _close() -> None:
        canvas.unbind_all("<MouseWheel>")
        win.destroy()

    def _save() -> None:
        config = dict(base_config)

        full_name = full_name_entry.get().strip()
        config["my_full_name"] = full_name
        # Derive match variants automatically (full name + individual name parts,
        # e.g. "Swar Vaid" -> ["Swar Vaid", "Swar", "Vaid"]) so greeting patterns
        # like "Hi Swar" still match without asking the user to type each variant.
        name_words = [w for w in re.split(r"\s+", full_name) if len(w) > 1]
        config["my_names"] = list(dict.fromkeys(([full_name] if full_name else []) + name_words))
        config["my_email"] = email_entry.get().strip().lower()

        # trusted_domains is fixed to intel.com and not user-editable here.
        config["trusted_domains"] = base_config.get("trusted_domains") or ["intel.com"]

        config["priority_senders"] = get_priority_senders()
        config["low_priority_senders"] = get_low_priority_senders()
        # automated_senders is never user-edited - detected on its own (built-in
        # patterns + GitHub Copilot's own judgment), so it's carried over as-is.
        config["automated_senders"] = base_config.get(
            "automated_senders", _load_json(CONFIG_DEFAULT_PATH).get("automated_senders", [])
        )

        config["urgent_keywords"] = _split_list(urgent_text.get("1.0", "end"))
        config["low_priority_keywords"] = _split_list(low_text.get("1.0", "end"))

        CONFIG_PATH.write_text(json.dumps(config, indent=2), encoding="utf-8")
        PROFILE_PATH.write_text(
            _build_profile_markdown(
                focus_text.get("1.0", "end"),
                config["priority_senders"],
                deprioritize_text.get("1.0", "end"),
                situational_text.get("1.0", "end"),
            ),
            encoding="utf-8",
        )
        saved["value"] = True
        messagebox.showinfo("Saved", "Your priorities have been saved.", parent=win)
        _close()

    def _skip() -> None:
        if not CONFIG_PATH.exists():
            CONFIG_PATH.write_text(CONFIG_DEFAULT_PATH.read_text(encoding="utf-8"), encoding="utf-8")
        if not PROFILE_PATH.exists():
            PROFILE_PATH.write_text(PROFILE_DEFAULT_PATH.read_text(encoding="utf-8"), encoding="utf-8")
        _close()

    save_btn = tk.Button(
        footer, text="Save", command=_save, font=_FONT_BUTTON, bg=_ACCENT, fg=_ACCENT_FG,
        activebackground=_ACCENT, activeforeground=_ACCENT_FG, relief="flat", padx=18, pady=6,
        cursor="hand2",
    )
    save_btn.pack(side="right")

    if first_run:
        tk.Button(
            footer, text="Skip for now", command=_skip, font=_FONT_LABEL, bg=_BG, fg=_TEXT_GRAY,
            relief="flat", padx=12, pady=6, cursor="hand2",
        ).pack(side="right", padx=(0, 8))
    else:
        tk.Button(
            footer, text="Cancel", command=_close, font=_FONT_LABEL, bg=_BG, fg=_TEXT_GRAY,
            relief="flat", padx=12, pady=6, cursor="hand2",
        ).pack(side="right", padx=(0, 8))

    win.protocol("WM_DELETE_WINDOW", _skip if first_run else _close)

    if owns_root:
        win.mainloop()
    else:
        win.wait_window()

    return saved["value"]


if __name__ == "__main__":
    run_wizard()
