# Email Assistant

A personal email productivity system built on Microsoft Outlook and local AI (Ollama). No cloud API required after setup.

---

## What It Does

**Every morning at 7am** — reads your inbox and sent emails, generates a structured daily brief using a local Ollama model, saves it to `briefs/MMDD_Brief.txt`, and opens it automatically.

**On demand** — chat with a local AI assistant that can pull live email context from Outlook mid-conversation.

---

## Structure

```
EmailAssistant/
  outlook_downloader/
    outlook_downloader.ps1    — save emails as .msg files
    fetch_emails_text.ps1     — read emails as plain text (used by chat)
    email_task_agent.ps1      — daily brief agent (runs at 7am)
    setup.ps1                 — first-time setup
  managed-agent-client/
    talk_to_agent.py          — interactive chat assistant
  briefs/
    0624_Brief.txt            — one brief per day
```

---

## Requirements

- Windows with Microsoft Outlook installed and signed in
- [Ollama](https://ollama.com) running locally (`gemma4:latest` by default)
- Python 3.8+ (for the chat assistant — no packages needed)

---

## Setup

```powershell
cd outlook_downloader
.\setup.ps1
```

Verifies Ollama is running, checks the model is available (offers to pull it if not), and registers a Windows Scheduled Task for the 7am daily brief.

---

## Usage

**Generate today's brief manually:**
```powershell
.\outlook_downloader\email_task_agent.ps1 -Force
```

**Chat with the assistant:**
```powershell
python managed-agent-client\talk_to_agent.py
```

**Download emails as .msg files:**
```powershell
.\outlook_downloader\outlook_downloader.ps1 -Hours 24
```

---

## Configuration

| What | How |
|------|-----|
| Change Ollama model for briefs | `-Model gemma3:27b` flag on `email_task_agent.ps1` |
| Change Ollama model for chat | `$env:OLLAMA_MODEL = "gemma3:27b"` |
| Change brief time window | `-Hours 48` flag on `email_task_agent.ps1` |
| Change Ollama URL | `-OllamaUrl` flag or `$env:OLLAMA_URL` |

---

## Notes

- All AI runs locally — no data leaves your machine
- `briefs/` and `logs/` are excluded from this repo via `.gitignore`
- The two tools reference each other via relative paths — keep the folder structure intact
- Outlook must be open and signed in at 7am for the scheduled brief to work
