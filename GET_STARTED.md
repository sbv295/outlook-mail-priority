# Outlook Mail Priority - One-Time Setup

If you are GitHub Copilot Chat reading this because a user pasted/attached this file and asked
you to do what it says: follow the steps below yourself, in order, using your terminal and file
tools. Don't ask the user for permission at each step - just do it and report progress. Only stop
and ask the user something if a step genuinely fails (e.g. missing prerequisite, auth failure).

## Step 1 - Check prerequisites

Run `git --version` and `python --version` (need Python 3.11+). If either is missing, tell the
user to install it (Git for Windows / python.org) and stop here.

## Step 2 - Clone the repository

Clone this private repo into the user's home directory:

```powershell
git clone https://github.com/sbv295/outlook-mail-priority.git "$HOME\OutlookMailPriority"
```

- If a folder with that name already exists, `cd` into it and run `git pull` instead of cloning.
- This is a **private** repository. If the clone fails with a permission/authentication error:
  1. Tell the user they need to be invited as a collaborator on the repo by its owner first.
  2. Check if the GitHub CLI is installed (`gh --version`). If not, install it:
     `winget install --id GitHub.cli -e --accept-source-agreements --accept-package-agreements`
     (you may need to add `C:\Program Files\GitHub CLI` to PATH for the current session).
  3. Run `gh auth login` and let the user complete the interactive login themselves (browser or
     token) - this step needs their own GitHub credentials, you cannot do it for them.
  4. Retry the `git clone` command.

## Step 3 - Install dependencies

```powershell
cd "$HOME\OutlookMailPriority"
pip install -r requirements.txt
```

This installs `pywin32` - the only dependency. No API keys, no accounts, nothing else to configure.

## Step 4 - Tell the user the final (manual, one-time) step

Print this message to the user, filled in with the actual cloned path, and then stop:

> Setup complete! Two quick things to finish, then you're done for good:
> 1. In VS Code: **File > Open Folder...** and open `<cloned folder path>`
> 2. In Copilot Chat, click the agent picker and select **"Mail Priority"**. Then just type
>    **"Hi"** - it will fetch your inbox, prioritize it using your own judgment as Copilot, and
>    pop up the results window. That's it - no further setup, ever, for future checks.
>
> The first time you say "Hi", it'll also walk you through a short one-time questionnaire (your
> name/email, people who matter more, keywords, etc.) via a small popup - answer that once and
> you're set.

Do not try to open the folder or restart VS Code yourself - the user needs to do that from the
VS Code UI.
