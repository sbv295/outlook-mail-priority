"""
gui.py
A Tkinter popup styled to resemble Outlook's own mail-list view: sender bold
on top, subject and preview text below, a colored category flag per row, real
1px divider lines between messages, collapsible group headers, and a hover
highlight. Double-click a row to open the actual email in Outlook (via COM).
"""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

LABEL_ORDER = ["High", "Medium", "Low", "Automated", "External"]

# Solid, saturated colors used for the small category flag/badge on each row
# (Outlook uses colored category tags rather than tinting the whole row).
LABEL_COLORS = {
    "High": "#d13438",       # red
    "Medium": "#b8860b",     # dark goldenrod
    "Low": "#8a8886",        # gray
    "Automated": "#0078d4",  # Outlook blue - machine-generated notifications
    "External": "#8764b8",   # purple - sender outside trusted domains (e.g. intel.com)
}

_BG = "#faf9f8"           # page background
_ROW_BG = "#ffffff"
_ROW_HOVER_BG = "#f3f2f1"
_BORDER = "#e1dfdd"        # real divider line color, and panel border
_HEADER_BG = "#ffffff"     # light title bar
_HEADER_FG = "#201f1e"
_HEADER_SUBFG = "#605e5c"
_GROUP_BG = "#f3f2f1"      # collapsible group header band
_GROUP_FG = "#201f1e"
_STATUS_BG = "#1f2937"     # dark slate status bar (bottom)
_STATUS_FG = "#9ca3af"
_TEXT = "#201f1e"          # near-black body text
_TEXT_GRAY = "#605e5c"     # secondary text (sender meta, section labels)
_DETAIL_BG = "#f5f6f8"     # tinted card behind Summary/Why - the actual payload of this tool
_DETAIL_HOVER_BG = "#eceef1"
_DETAIL_BORDER = "#e3e5e9"

_FONT_SENDER = ("Segoe UI", 10, "bold")
_FONT_SUBJECT = ("Segoe UI", 11, "bold")
_FONT_SECTION_LABEL = ("Segoe UI", 8, "bold")
_FONT_SUMMARY = ("Segoe UI", 10)
_FONT_REASONS = ("Segoe UI", 9)
_FONT_META = ("Segoe UI", 8)
_FONT_BADGE = ("Segoe UI", 8, "bold")
_FONT_TITLE = ("Segoe UI", 13, "bold")
_FONT_SUBTITLE = ("Segoe UI", 9)
_FONT_GROUP = ("Segoe UI", 9, "bold")
_FONT_ARROW = ("Segoe UI", 16, "bold")

_WRAP_PX = 760  # pixel wrap width for subject/summary/reasons text


def _bind_hover(pairs: list) -> None:
    """
    Highlight a whole row together on hover. `pairs` is a list of
    (widget, normal_bg, hover_bg) tuples - the detail card keeps its own
    tinted colors while the rest of the row uses the plain row colors, but
    they all highlight/un-highlight together as one visual unit.
    """
    widgets = [w for (w, _n, _h) in pairs]

    def on_enter(_event=None) -> None:
        for w, _n, h in pairs:
            w.configure(bg=h)

    def on_leave(_event=None) -> None:
        anchor = widgets[0]
        x, y = anchor.winfo_pointerxy()
        target = anchor.winfo_containing(x, y)
        if target not in widgets:
            for w, n, _h in pairs:
                w.configure(bg=n)

    for w in widgets:
        w.bind("<Enter>", on_enter)
        w.bind("<Leave>", on_leave)


def _bind_open(widgets: list, mail_item) -> None:
    def _open(_event=None) -> None:
        try:
            mail_item.Display()
        except Exception as exc:
            messagebox.showerror("Could not open email", str(exc))

    for w in widgets:
        w.bind("<Double-1>", _open)
        w.configure(cursor="hand2")


def show_priority_popup(scored: list) -> None:
    """
    scored: list of (MailInfo, ScoreResult) tuples, already sorted however the
    caller likes (this function re-groups by label, preserving relative order
    within each label).
    """
    grouped: dict[str, list] = {lbl: [] for lbl in LABEL_ORDER}
    for mail, result in scored:
        grouped.setdefault(result.label, []).append((mail, result))

    root = tk.Tk()
    root.title("Outlook Priority Report")
    root.geometry("1000x720")
    root.minsize(760, 480)
    root.configure(bg=_BG)

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure("Report.Vertical.TScrollbar", background=_BORDER, troughcolor=_BG, arrowsize=14)

    # ---- Title bar ----
    header = tk.Frame(root, bg=_HEADER_BG)
    header.pack(fill="x", side="top")
    tk.Label(
        header, text="Outlook Priority Report", bg=_HEADER_BG, fg=_HEADER_FG,
        font=_FONT_TITLE, anchor="w", padx=18, pady=12,
    ).pack(side="left")
    tk.Label(
        header, text=f"{len(scored)} email(s) analyzed", bg=_HEADER_BG, fg=_HEADER_SUBFG,
        font=_FONT_SUBTITLE, anchor="e", padx=18,
    ).pack(side="right")

    def _open_customize() -> None:
        from setup_wizard import run_wizard
        run_wizard(master=root)

    tk.Button(
        header, text="Customize", command=_open_customize, font=_FONT_SUBTITLE,
        bg=_GROUP_BG, fg=_HEADER_FG, activebackground=_BORDER, relief="flat",
        padx=12, pady=4, cursor="hand2",
    ).pack(side="right", padx=(0, 8))
    tk.Frame(root, bg=_BORDER, height=1).pack(fill="x", side="top")

    tk.Label(
        root,
        text="Double-click a message to open it in Outlook.  Click a group header to collapse/expand it.",
        bg=_BG, fg=_TEXT_GRAY, anchor="w", padx=16, pady=6, font=_FONT_SUBTITLE,
    ).pack(fill="x")

    # ---- Status bar (packed first so it stays pinned to the bottom) ----
    status = tk.Frame(root, bg=_STATUS_BG)
    status.pack(fill="x", side="bottom")
    counts_text = "    ".join(
        f"{lbl}: {len(grouped.get(lbl, []))}" for lbl in LABEL_ORDER if grouped.get(lbl)
    )
    tk.Label(
        status, text=counts_text or "No emails", bg=_STATUS_BG, fg=_STATUS_FG,
        font=_FONT_SUBTITLE, anchor="w", padx=18, pady=6,
    ).pack(side="left")

    # ---- Scrollable, bordered mail list (Outlook-style) ----
    panel = tk.Frame(root, bg=_BORDER)
    panel.pack(fill="both", expand=True, padx=14, pady=(0, 12))
    inner = tk.Frame(panel, bg=_ROW_BG)
    inner.pack(fill="both", expand=True, padx=1, pady=1)

    canvas = tk.Canvas(inner, bg=_ROW_BG, highlightthickness=0)
    vsb = ttk.Scrollbar(inner, orient="vertical", command=canvas.yview, style="Report.Vertical.TScrollbar")
    canvas.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    content = tk.Frame(canvas, bg=_ROW_BG)
    content_window = canvas.create_window((0, 0), window=content, anchor="nw")

    def _on_content_configure(_event=None) -> None:
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _on_canvas_configure(event) -> None:
        canvas.itemconfigure(content_window, width=event.width)

    content.bind("<Configure>", _on_content_configure)
    canvas.bind("<Configure>", _on_canvas_configure)

    def _on_mousewheel(event) -> None:
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _fmt_time(mail) -> str:
        return mail.received.strftime("%b %d, %I:%M %p")

    def _toggle_group(body: tk.Frame, arrow_label: tk.Label, anchor: tk.Widget) -> None:
        if body.winfo_ismapped():
            body.pack_forget()
            arrow_label.configure(text="\u25b8")  # collapsed
        else:
            # pack() with no position would append body to the END of the
            # packing order (after every other group) instead of putting it
            # back where it belongs, right after its own header's divider.
            body.pack(fill="x", after=anchor)
            arrow_label.configure(text="\u25be")  # expanded

    for lbl in LABEL_ORDER:
        entries = grouped.get(lbl, [])
        if not entries:
            continue

        group_header = tk.Frame(content, bg=_GROUP_BG, cursor="hand2")
        group_header.pack(fill="x")
        arrow = tk.Label(
            group_header, text="\u25be", bg=_GROUP_BG, fg=_GROUP_FG, font=_FONT_ARROW, width=2,
        )
        arrow.pack(side="left", padx=(10, 4), pady=4)
        flag = tk.Frame(group_header, bg=LABEL_COLORS[lbl], width=10, height=10)
        flag.pack(side="left", padx=(6, 8), pady=6)
        title_lbl = tk.Label(
            group_header, text=f"{lbl.upper()}  \u00b7  {len(entries)}", bg=_GROUP_BG, fg=_GROUP_FG,
            font=_FONT_GROUP, anchor="w",
        )
        title_lbl.pack(side="left", pady=6)
        header_divider = tk.Frame(content, bg=_BORDER, height=1)
        header_divider.pack(fill="x")

        group_body = tk.Frame(content, bg=_ROW_BG)
        group_body.pack(fill="x", after=header_divider)

        for widget in (group_header, arrow, flag, title_lbl):
            widget.bind(
                "<Button-1>",
                lambda _e, b=group_body, a=arrow, anchor=header_divider: _toggle_group(b, a, anchor),
            )

        for mail, result in entries:
            row = tk.Frame(group_body, bg=_ROW_BG)
            row.pack(fill="x")

            stripe = tk.Frame(row, bg=LABEL_COLORS[lbl], width=4)
            stripe.pack(side="left", fill="y")

            main = tk.Frame(row, bg=_ROW_BG)
            main.pack(side="left", fill="both", expand=True, padx=(12, 14), pady=10)

            top_line = tk.Frame(main, bg=_ROW_BG)
            top_line.pack(fill="x")
            sender_lbl = tk.Label(
                top_line, text=mail.sender_name, bg=_ROW_BG, fg=_TEXT,
                font=_FONT_SENDER, anchor="w",
            )
            sender_lbl.pack(side="left")
            meta_lbl = tk.Label(
                top_line, text=_fmt_time(mail),
                bg=_ROW_BG, fg=_TEXT_GRAY, font=_FONT_META, anchor="e",
            )
            meta_lbl.pack(side="right")
            badge_lbl = tk.Label(
                top_line, text=lbl, bg=LABEL_COLORS[lbl], fg="#ffffff",
                font=_FONT_BADGE, padx=6, pady=1,
            )
            badge_lbl.pack(side="right", padx=(0, 10))

            subject_lbl = tk.Label(
                main, text=mail.subject, bg=_ROW_BG, fg=_TEXT, font=_FONT_SUBJECT,
                anchor="w", justify="left", wraplength=_WRAP_PX,
            )
            subject_lbl.pack(fill="x", anchor="w", pady=(4, 0))

            # (widget, normal_bg, hover_bg) pairs - the detail card below keeps
            # its own tinted colors while the rest of the row uses plain white,
            # but the whole thing highlights together as one row on hover.
            pairs = [
                (row, _ROW_BG, _ROW_HOVER_BG),
                (main, _ROW_BG, _ROW_HOVER_BG),
                (top_line, _ROW_BG, _ROW_HOVER_BG),
                (sender_lbl, _ROW_BG, _ROW_HOVER_BG),
                (meta_lbl, _ROW_BG, _ROW_HOVER_BG),
                (subject_lbl, _ROW_BG, _ROW_HOVER_BG),
            ]

            # Summary/Why are the actual payload of this tool, so they get their
            # own labeled, tinted "detail card" instead of small Outlook-style
            # preview text - full-size readable font, not an afterthought.
            if result.summary or result.reasons:
                detail = tk.Frame(main, bg=_DETAIL_BG, highlightthickness=1, highlightbackground=_DETAIL_BORDER)
                detail.pack(fill="x", pady=(8, 0))
                pairs.append((detail, _DETAIL_BG, _DETAIL_HOVER_BG))

                if result.summary:
                    summary_caption = tk.Label(
                        detail, text="SUMMARY", bg=_DETAIL_BG, fg=_TEXT_GRAY,
                        font=_FONT_SECTION_LABEL, anchor="w",
                    )
                    summary_caption.pack(fill="x", padx=10, pady=(8, 0))
                    summary_lbl = tk.Label(
                        detail, text=result.summary, bg=_DETAIL_BG, fg=_TEXT, font=_FONT_SUMMARY,
                        anchor="w", justify="left", wraplength=_WRAP_PX,
                    )
                    summary_lbl.pack(fill="x", anchor="w", padx=10, pady=(2, 0))
                    pairs.append((summary_caption, _DETAIL_BG, _DETAIL_HOVER_BG))
                    pairs.append((summary_lbl, _DETAIL_BG, _DETAIL_HOVER_BG))

                if result.reasons:
                    reasons_caption = tk.Label(
                        detail, text="WHY", bg=_DETAIL_BG, fg=_TEXT_GRAY,
                        font=_FONT_SECTION_LABEL, anchor="w",
                    )
                    reasons_caption.pack(fill="x", padx=10, pady=(8, 0))
                    reasons_text = "\n".join(f"\u2022 {r}" for r in result.reasons)
                    reasons_lbl = tk.Label(
                        detail, text=reasons_text, bg=_DETAIL_BG, fg=_TEXT, font=_FONT_REASONS,
                        anchor="w", justify="left", wraplength=_WRAP_PX,
                    )
                    reasons_lbl.pack(fill="x", anchor="w", padx=10, pady=(2, 8))
                    pairs.append((reasons_caption, _DETAIL_BG, _DETAIL_HOVER_BG))
                    pairs.append((reasons_lbl, _DETAIL_BG, _DETAIL_HOVER_BG))
                else:
                    tk.Frame(detail, bg=_DETAIL_BG, height=8).pack(fill="x")

            _bind_hover(pairs)
            _bind_open([w for (w, _n, _h) in pairs] + [badge_lbl], mail.item)

            # Real 1px divider line between messages.
            tk.Frame(group_body, bg=_BORDER, height=1).pack(fill="x")

    root.mainloop()
