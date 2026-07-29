# Outlook Mail Priority

A Windows tool that triages your Outlook inbox by priority — High / Medium / Low / Automated / External —
and shows the results in a clean popup GUI. Double-click any email to open it in Outlook.

There's no external AI API, no API keys, and no subscription: the "AI" doing the actual prioritizing
is **GitHub Copilot Chat in VS Code**, using your own enterprise Copilot access. The tool itself just
fetches mail from Outlook and renders the GUI — everything subjective is done by Copilot, guided by
your own priorities.

> **New user?** You don't need to read this whole README. Just grab [GET_STARTED.md](GET_STARTED.md)
> (ask whoever shared this tool with you for that one file), open it in VS Code, and ask Copilot
> Chat to "do what is asked in this file." It clones the repo, installs the one dependency, and
> tells you the one manual click left. From then on, just open the **Mail Priority** agent and
> say "Hi".

## Requirements

- Windows, with desktop Outlook installed and configured
- Python 3.11+
- VS Code with GitHub Copilot Chat (enterprise or individual access)

## Getting started

1. **Clone the repo**
   ```powershell
   git clone <this-repo-url>
   cd outlook-mail-priority
   ```

2. **Install the one dependency**
   ```powershell
   pip install -r requirements.txt
   ```

3. **Open the folder in VS Code** and make sure Copilot Chat is signed in.

4. **Run the setup wizard** to tell it about yourself — your name/email, people whose mail matters
   more, low-priority senders, keywords, and your current focus areas:
   ```powershell
   python main.py --setup
   ```
   (If you skip this, it runs automatically the first time you use the tool.)

5. **Ask Copilot Chat to check your mail.** In VS Code's Copilot Chat, just type something like:
   > Check my email priority

   Copilot will fetch your recent/unread mail, score it itself (using your priorities as context),
   and pop up the GUI. You can also switch to the included **"Mail Priority"** custom agent
   (agent picker in Copilot Chat) and just say **"Hi"** to trigger the same check instantly.

## Re-customizing your priorities

Priorities drift over time. You can update them anytime:
- Click **Customize** in the top-right of the GUI, or
- Run `python main.py --setup` again

Your answers are saved locally to `user_data/config.json` and `user_data/priority_profile.md` —
this folder is git-ignored and never leaves your machine.

## Updates

Whenever you use the GUI, the tool does a quick (rate-limited to once every 6 hours) check for
new commits on GitHub. If any are found, you'll get a popup listing them with a Yes/No prompt to
pull the update - it's a plain `git pull --ff-only`, so it never touches `user_data/` and never
overwrites local changes. You can also check manually anytime:
```powershell
python main.py --check-updates
```

## Categories

| Label | Meaning |
|---|---|
| **High / Medium / Low** | Genuine human-sent mail, ranked by whether it needs *your* attention or action |
| **Automated** | Machine-generated notifications (bug trackers, Jira, IT alerts, bulk mail, ...) — triaged separately, regardless of how urgent the content sounds |
| **External** | Mail from outside your trusted domain(s) (e.g. not your company) |

Precedence: **Automated > External > High/Medium/Low**.

## CLI reference (optional, for manual/advanced use)

```powershell
python main.py --recent 50 --gui          # rule-based scoring (no Copilot) + GUI
python main.py --export-json out.json     # just dump raw mail data, no scoring
python main.py --gui-from-json out.json   # show the GUI from a scored JSON file
python main.py --tag                      # also tag emails in Outlook with Priority-* categories
python main.py --setup                    # re-open the guided setup wizard
python main.py --check-updates            # check GitHub for updates now and prompt to pull them
```

## Project structure

```
main.py                          CLI entry point
gui.py                           Tkinter popup GUI
scorers.py                       Deterministic rule-based scorer (hints only)
outlook_client.py                Outlook COM automation
setup_wizard.py                  Guided onboarding wizard
config.default.json              Clean-slate config template (tracked)
priority_profile.default.md      Clean-slate profile template (tracked)
.github/copilot-instructions.md  Tells Copilot Chat how to run the workflow
.github/agents/                  Custom "Mail Priority" agent mode
user_data/                       Your personal config/profile/mail dumps (git-ignored)
```

## Privacy

Nothing you enter (names, senders, keywords, focus areas) or any mail data is ever committed to
this repo — it all lives in `user_data/`, which is excluded via `.gitignore`. Only the tool's code
and clean-slate templates are tracked.
