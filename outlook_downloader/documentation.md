# Outlook Email Downloader

Downloads emails from your local Microsoft Outlook and saves them as `.msg` files, organised into folders by date. Also includes a daily brief agent that summarises your inbox every morning using a local Ollama model.

---

## Requirements

- Windows PC with **Microsoft Outlook installed and signed in**
- PowerShell (built into Windows — no installs needed)
- [Ollama](https://ollama.com) installed and running (for the daily brief agent)

---

## Quick Start

Open a PowerShell window, navigate to this folder, and run:

```powershell
cd "$env:USERPROFILE\Desktop\outlook_downloader"
.\outlook_downloader.ps1 -Hours 18
```

If PowerShell blocks the script, run this once to allow it:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

---

## Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Hours N` | Number | Emails received in the **past N hours** |
| `-Count N` | Number | The **last N emails** (newest first) |
| `-Unread` | Flag | **Unread emails only** |
| `-Output PATH` | Path | Where to save files. Default: `.\emails` |
| `-Folder NAME` | Text | Which Outlook folder to pull from. Default: `Inbox` |
| `-Help` | Flag | Show usage information |

At least one of `-Hours`, `-Count`, or `-Unread` is required. The rest are optional.

---

## Examples

```powershell
# Emails from the past 18 hours
.\outlook_downloader.ps1 -Hours 18

# Last 20 emails
.\outlook_downloader.ps1 -Count 20

# All unread emails in the inbox
.\outlook_downloader.ps1 -Unread

# Unread emails from the past 24 hours
.\outlook_downloader.ps1 -Hours 24 -Unread

# Last 20 unread emails, saved to a custom folder
.\outlook_downloader.ps1 -Count 20 -Unread -Output C:\EmailBackup

# Last 50 emails from Sent Items
.\outlook_downloader.ps1 -Count 50 -Folder "Sent Items"

# Show help
.\outlook_downloader.ps1 -Help
```

---

## Output Structure

Emails are saved as `.msg` files inside dated subfolders:

```
emails/
  2026-06-19/
    100949_Nachamma Sockalingam_Week 3 June 2026.msg
    083729_Sports Booking_FW SUTD Booking Confirmation.msg
  2026-06-18/
    194211_Lymon Sim_RE Update on Cohort Instructor.msg
    191741_Koh Chye Soon_RE Follow up on AI Training.msg
```

Filename format: `HHMMSS_SenderName_Subject.msg`

`.msg` files are native Outlook format — double-click any file to open it directly in Outlook, with full formatting, attachments, and reply capability intact.

---

## Supported Folders

The `-Folder` parameter accepts these names (case-insensitive):

| Name | What it accesses |
|------|-----------------|
| `Inbox` | Your main inbox (default) |
| `Sent Items` or `Sent` | Emails you have sent |
| `Drafts` | Unsent draft emails |
| `Deleted` | Deleted items |
| `Junk` or `Junk Email` | Spam / junk folder |
| Any other name | Searches all mailboxes for a matching folder name |

---

## How Parameters Stack

Each parameter is independent and only applies when explicitly passed. Omitting a parameter imposes no hidden limit.

| Parameters used | Behaviour |
|----------------|-----------|
| `-Hours 18` | All emails from the past 18 hours. No count cap. |
| `-Count 20` | The 20 newest emails. No time cap. |
| `-Unread` | All unread emails. No time or count cap. |
| `-Hours 24 -Unread` | Unread emails from the past 24 hours. |
| `-Count 20 -Unread` | The 20 newest unread emails. |
| `-Hours 48 -Count 10` | Up to 10 emails, but only from the past 48 hours. |
| `-Hours 24 -Count 20 -Unread` | Up to 20 unread emails, all within the past 24 hours. |

When `-Hours` and `-Count` are both set, whichever limit is hit first stops the download. `-Unread` is always a filter on top of whatever else is specified, never a limit of its own.

---

## Daily Brief Agent

A script that runs automatically at 7am every day. It reads your inbox and sent emails from the past 24 hours, sends them to a local Ollama model, and saves a structured `Daily Brief.txt` to your Desktop.

### One-time setup

```powershell
cd "$env:USERPROFILE\Desktop\outlook_downloader"
.\setup.ps1
```

The setup script will:
1. Verify Ollama is running and the default model (`gemma4:latest`) is available — offering to pull it if not
2. Register a Windows Scheduled Task to run the agent daily at 7am

### Running manually

```powershell
.\email_task_agent.ps1
```

With a custom time window:

```powershell
.\email_task_agent.ps1 -Hours 48
```

With a different Ollama model:

```powershell
.\email_task_agent.ps1 -Model gemma3:27b
```

Force regeneration even if a brief was already produced today:

```powershell
.\email_task_agent.ps1 -Force
```

### Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-Hours N` | `24` | How many hours of email to review |
| `-BriefFile PATH` | `Desktop\Daily Brief.txt` | Where to save the output |
| `-Model NAME` | `gemma4:latest` | Ollama model to use |
| `-OllamaUrl URL` | `http://localhost:11434` | Ollama base URL |
| `-Force` | off | Run even if a brief was already generated today |

### What it produces

**`Daily Brief.txt`** on your Desktop — regenerated each morning:

```
DAILY BRIEF — 2026-06-24
=========================

OVERVIEW
--------
12 emails received, 2 sent. Dominant themes: project scheduling and
vendor follow-ups. One item is time-sensitive.

KEY THREADS
-----------
  [RE: Follow up on AI Training] (From: Koh Chye Soon)
  Awaiting revised slide deck before Thursday's session.

  ...

ACTION ITEMS
------------
  > Send revised slides to Koh Chye Soon (from: RE Follow up on AI Training)

FYI — NO ACTION NEEDED
-----------------------
  [Sports Booking Confirmation] — booking confirmed for Friday
```

**`logs/agent_YYYY-MM-DD.log`** — one log file per day with timestamps for every step.

### How it works

1. Connects to local Outlook via COM (same as the downloader)
2. Reads inbox and sent emails from the past 24 hours directly — no `.msg` files written
3. Sends the email summaries to a local Ollama model (`gemma4:latest` by default)
4. Saves the formatted brief to `Daily Brief.txt` on your Desktop
5. Shows a Windows notification when done

### Notes

- Outlook must be installed and signed in. The task runs using your interactive Windows session, so Outlook will open silently if not already running.
- Ollama must be running at the time the task fires. If it is not, the script exits with an error logged to `logs/`.
- No API key or internet connection required — the model runs entirely on your machine.

---

## Troubleshooting

**"Could not connect to Outlook"**
- Open Outlook and make sure you are fully signed in before running the script.
- If Outlook prompts for a password or shows a setup wizard, complete that first.

**"0 emails saved"**
- Your filter may be too narrow. Try relaxing it — remove `-Unread`, increase `-Hours`, or try `-Count 5` to confirm emails exist.
- Check you are looking in the right folder with `-Folder`.

**"Ollama not reachable"**
- Make sure the Ollama app is running. You can start it from the system tray or by running `ollama serve` in a terminal.

**Script is blocked by PowerShell**
- Run this once in PowerShell as your user (no admin required):
  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
  ```

**Outlook security prompt appears**
- Outlook may ask "Allow this program to access email?" — click **Allow**. This is a one-time prompt per session.
