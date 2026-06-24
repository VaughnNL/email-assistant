# Outlook Downloader

Two scripts for reading emails from local Microsoft Outlook.

---

## Requirements

- Windows with Microsoft Outlook installed and signed in
- PowerShell (built into Windows)

---

## outlook_downloader.ps1 — Save emails as .msg files

Downloads emails and saves them as native `.msg` files organised into dated subfolders.

```powershell
.\outlook_downloader.ps1 -Hours 24
.\outlook_downloader.ps1 -Count 20
.\outlook_downloader.ps1 -Unread
.\outlook_downloader.ps1 -Hours 24 -Unread
.\outlook_downloader.ps1 -Count 50 -Folder "Sent Items"
```

| Parameter | Description |
|-----------|-------------|
| `-Hours N` | Emails from the past N hours |
| `-Count N` | The last N emails (newest first) |
| `-Unread` | Unread emails only |
| `-Output PATH` | Where to save files. Default: `.\emails` |
| `-Folder NAME` | Outlook folder. Default: `Inbox`. Others: `Sent Items`, `Drafts`, `Deleted`, `Junk` |
| `-Help` | Show usage |

At least one of `-Hours`, `-Count`, or `-Unread` is required.

**Output structure:**
```
emails/
  2026-06-24/
    100949_Sender Name_Subject.msg
```

`.msg` files open directly in Outlook with full formatting and attachments.

---

## fetch_emails_text.ps1 — Return emails as plain text

Returns recent inbox and sent emails as formatted plain text. Used internally by the chat assistant to inject live email context.

```powershell
.\fetch_emails_text.ps1 -Hours 24
```

---

## email_task_agent.ps1 — Daily brief

Reads your inbox and sent emails, sends them to a local Ollama model, and saves a structured brief to `../briefs/MMDD_Brief.txt`. Runs daily at 7am via Windows Task Scheduler.

```powershell
.\email_task_agent.ps1                     # normal run (skips if already done today)
.\email_task_agent.ps1 -Force              # regenerate even if already done today
.\email_task_agent.ps1 -Hours 48           # cover the past 48 hours
.\email_task_agent.ps1 -Model gemma3:27b   # use a different Ollama model
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-Hours N` | `24` | Hours of email to review |
| `-Model NAME` | `gemma4:latest` | Ollama model |
| `-OllamaUrl URL` | `http://localhost:11434` | Ollama base URL |
| `-BriefFile PATH` | `../briefs/MMDD_Brief.txt` | Output path |
| `-Force` | off | Run even if a brief exists for today |

The brief opens automatically when generated.

---

## setup.ps1 — First-time setup

Run once to verify Ollama and register the 7am scheduled task:

```powershell
.\setup.ps1
```

Re-run any time you move the folder or want to re-register the task.

---

## Troubleshooting

**"Could not connect to Outlook"** — Open Outlook and sign in fully before running.

**"0 emails found"** — Widen your filter: increase `-Hours`, remove `-Unread`, or check `-Folder`.

**"Ollama not reachable"** — Start Ollama from the system tray or run `ollama serve`.

**Script blocked by PowerShell** — Run once:
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

**Outlook security prompt** — Click Allow. One-time per session.
