# Chat Assistant

A local Python chat client powered by Ollama. Runs interactively in the terminal. Automatically injects live Outlook email context when your message contains email-related keywords.

## Requirements

- Python 3.8+
- Ollama running locally with a model pulled (default: `gemma4:latest`)
- No external Python packages — uses stdlib only

## Usage

```powershell
# Interactive REPL
python talk_to_agent.py

# Single message
python talk_to_agent.py "What's in my inbox today?"
```

## Email Context Injection

If your message contains any of these keywords, recent emails are automatically fetched from Outlook and prepended to the message before sending to the model:

```
brief  email  inbox  mail  summary  summarise  summarize
```

Examples that trigger injection:
- "Give me a brief on my inbox"
- "Any urgent emails today?"
- "Summarize my mail from this morning"

## Configuration

Override defaults via environment variables:

```powershell
$env:OLLAMA_MODEL = "gemma3:27b"
$env:OLLAMA_URL   = "http://localhost:11434"
```

## How It Works

1. Starts a chat session with a system prompt
2. Detects email keywords in each message
3. If triggered, calls `../outlook_downloader/fetch_emails_text.ps1` as a subprocess and prepends the output
4. Sends the full message history to Ollama and streams the response
5. In REPL mode, keeps conversation history across turns
