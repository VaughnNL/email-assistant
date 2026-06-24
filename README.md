# Email Assistant

A personal email productivity system built on Microsoft Outlook and local AI. Generates a structured daily brief from your inbox every morning and provides an interactive AI chat client with live email context.

---

## Structure

```
EmailAssistant/
  outlook_downloader/     — Outlook email tools + daily brief agent
  managed-agent-client/   — Interactive AI chat client
  briefs/                 — Daily briefs (0624_Brief.txt, 0625_Brief.txt, ...)
```

---

## Tools

### Daily Brief Agent

Runs automatically at 7am. Reads your inbox and sent emails from the past 24 hours, sends them to a local [Ollama](https://ollama.com) model, and saves a structured plain-text brief to `briefs/MMDD_Brief.txt`.

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

ACTION ITEMS
------------
  > Send revised slides to Koh Chye Soon (from: RE Follow up on AI Training)

FYI — NO ACTION NEEDED
-----------------------
  [Sports Booking Confirmation] — booking confirmed for Friday
```

### Managed Agent Client

A terminal client for an Anthropic Managed Agent. Automatically injects live email context when your message contains email-related keywords (`brief`, `inbox`, `email`, `summary`, etc.).

```powershell
python managed-agent-client/talk_to_agent.py "What's in my inbox today?"
```

### Outlook Downloader

Downloads emails from Outlook as `.msg` files, organised into dated folders. Useful for bulk archiving or offline processing.

```powershell
.\outlook_downloader\outlook_downloader.ps1 -Hours 24
```

---

## Requirements

- Windows with Microsoft Outlook installed and signed in
- [Ollama](https://ollama.com) running locally with `gemma4:latest` (or another model of your choice)
- Python 3.8+ with `anthropic` SDK (`pip install anthropic`)
- Anthropic API key (for the managed agent client)

---

## Setup

```powershell
cd outlook_downloader
.\setup.ps1
```

The setup script:
1. Verifies Ollama is running and the model is available (offers to pull it if not)
2. Registers a Windows Scheduled Task to run the daily brief at 7am

---

## Running Manually

```powershell
# Generate today's brief now
.\outlook_downloader\email_task_agent.ps1 -Force

# Use a different model
.\outlook_downloader\email_task_agent.ps1 -Force -Model gemma3:27b

# Chat with the managed agent
python managed-agent-client\talk_to_agent.py "Give me a brief on my inbox"

# Download emails as .msg files
.\outlook_downloader\outlook_downloader.ps1 -Hours 24
```

---

## Notes

- `outlook_downloader/config.ps1` (API key) and the `briefs/` and `logs/` folders are excluded from this repo via `.gitignore`
- The daily brief agent requires Outlook to be open and signed in at 7am — it connects via COM and reuses the existing session
- The two tools reference each other via relative paths, so the folder structure must be kept intact
