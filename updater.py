"""
updater.py
Lightweight "check for updates" using git fetch/pull against the repo's
configured remote - no custom update server, no download mechanism, just
git. Purely optional: silently does nothing if this isn't a git checkout,
has no remote configured, or the network is unavailable.

Rate-limited to avoid a network call on every single invocation (Copilot
often runs this CLI many times in one chat session) - see CHECK_INTERVAL_SECONDS.
Never touches user_data/ (it's gitignored, so a git pull can't affect it).
"""
from __future__ import annotations

import json
import subprocess
import time
from pathlib import Path

APP_DIR = Path(__file__).parent
USER_DATA_DIR = APP_DIR / "user_data"
LAST_CHECK_PATH = USER_DATA_DIR / ".last_update_check.json"

CHECK_INTERVAL_SECONDS = 6 * 60 * 60  # don't hit the network more than once every 6 hours unless forced


def _run_git(*args: str) -> str | None:
    try:
        result = subprocess.run(
            ["git", *args],
            cwd=APP_DIR,
            capture_output=True,
            text=True,
            timeout=15,
        )
        if result.returncode != 0:
            return None
        return result.stdout.strip()
    except Exception:
        return None


def _should_check(force: bool) -> bool:
    if force:
        return True
    if not LAST_CHECK_PATH.exists():
        return True
    try:
        last = json.loads(LAST_CHECK_PATH.read_text(encoding="utf-8")).get("last_check", 0)
    except Exception:
        return True
    return (time.time() - last) > CHECK_INTERVAL_SECONDS


def _record_check() -> None:
    USER_DATA_DIR.mkdir(exist_ok=True)
    LAST_CHECK_PATH.write_text(json.dumps({"last_check": time.time()}), encoding="utf-8")


def check_for_updates(force: bool = False) -> dict | None:
    """
    Returns None if there's nothing to report (not a git repo, no remote,
    offline, already up to date, or skipped due to rate limiting).
    Otherwise returns {"behind": int, "messages": list[str], "branch": str,
    "remote": str}.
    """
    if not _should_check(force):
        return None

    if _run_git("rev-parse", "--is-inside-work-tree") != "true":
        return None

    branch = _run_git("rev-parse", "--abbrev-ref", "HEAD")
    if not branch or branch == "HEAD":
        return None

    remote = _run_git("config", f"branch.{branch}.remote") or "origin"
    if _run_git("remote", "get-url", remote) is None:
        return None

    if _run_git("fetch", remote, "--quiet") is None:
        _record_check()
        return None  # offline or fetch failed - don't retry again for a while

    _record_check()

    remote_ref = f"{remote}/{branch}"
    behind_output = _run_git("rev-list", "--count", f"HEAD..{remote_ref}")
    if not behind_output or not behind_output.isdigit():
        return None
    behind = int(behind_output)
    if behind == 0:
        return None

    log_output = _run_git("log", "--oneline", f"HEAD..{remote_ref}") or ""
    messages = [
        line.split(" ", 1)[1] if " " in line else line
        for line in log_output.splitlines()
    ][:10]

    return {"behind": behind, "messages": messages, "branch": branch, "remote": remote}


def apply_update(branch: str, remote: str = "origin") -> tuple[bool, str]:
    """
    Fast-forward-only pull. Never touches user_data/ (gitignored) and never
    overwrites local commits/changes - if a fast-forward isn't possible
    (e.g. you've made local edits), this fails safely with no changes made.
    """
    output = _run_git("pull", "--ff-only", remote, branch)
    if output is None:
        return False, (
            "Update failed - you may have local changes that conflict. "
            "Run 'git pull' manually in the project folder to see details."
        )
    return True, "Updated successfully. Restart the tool to use the new version."


def prompt_and_apply(info: dict) -> None:
    """Show a small Tkinter Yes/No dialog for an update found by check_for_updates(),
    and apply it if the user accepts. Safe to call even with no other Tk window open."""
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()

    commit_list = "\n".join(f"\u2022 {m}" for m in info["messages"])
    remaining = info["behind"] - len(info["messages"])
    if remaining > 0:
        commit_list += f"\n... and {remaining} more"

    proceed = messagebox.askyesno(
        "Update available",
        f"{info['behind']} new commit(s) available:\n\n{commit_list}\n\nUpdate now?",
    )
    if proceed:
        success, message = apply_update(info["branch"], info["remote"])
        if success:
            messagebox.showinfo("Update complete", message)
        else:
            messagebox.showerror("Update failed", message)

    root.destroy()
