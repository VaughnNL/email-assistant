# Managed Agent Client

A minimal Python client that connects to a pre-deployed Anthropic Managed Agent, streams its responses, and exits cleanly. When the user's message contains email-related keywords, it automatically fetches live inbox context from Outlook and injects it into the request.

## Files

| File | Description |
|---|---|
| `talk_to_agent.py` | Main script |
| `documentation.md` | This file |

## Requirements

- Python 3.8+
- `anthropic` SDK
- Microsoft Outlook (for automatic email context injection)
- `outlook_downloader/fetch_emails_text.ps1` must exist at `$env:USERPROFILE\Desktop\outlook_downloader\fetch_emails_text.ps1`

```powershell
pip install anthropic
```

## Setup

Set your Anthropic API key as an environment variable:

```powershell
$env:ANTHROPIC_API_KEY = "sk-ant-..."
```

## Usage

```powershell
# Default greeting
python talk_to_agent.py

# Custom message
python talk_to_agent.py "Summarize what you can do for me"

# Triggers automatic email context injection
python talk_to_agent.py "Give me a brief on my inbox"
```

## Email Context Injection

If the user message contains any of these keywords, the client automatically calls `fetch_emails_text.ps1` and prepends the last 24 hours of inbox and sent emails to the message before sending:

```
brief  email  inbox  mail  summary  summarise  summarize
```

This means asking "what's in my inbox?" or "give me a brief" will automatically include live email data — no manual copy-paste required.

The PowerShell script is invoked as a subprocess and must be reachable at:
```
$env:USERPROFILE\Desktop\outlook_downloader\fetch_emails_text.ps1
```

## Configuration

The agent and environment IDs are hardcoded at the top of `talk_to_agent.py`:

```python
AGENT_ID       = "agent_015VHQjSZein5oq1iUZr6wSX"
ENVIRONMENT_ID = "env_01KLJ9UFFW2otvWxy6GAjeyo"
```

Change these to point at a different agent or environment without altering any other logic.

## How It Works

1. **Create session** — calls `client.beta.sessions.create` with the agent ID and environment ID. The session is a single stateful run of the agent inside an Anthropic-hosted container.

2. **Inject email context (if triggered)** — detects keywords in the user message, calls `fetch_emails_text.ps1` as a subprocess, and prepends the output to the message.

3. **Open SSE stream first** — opens `client.beta.sessions.events.stream` *before* sending the message. This is required so that no early events (status transitions, agent output) are missed.

4. **Send user message** — calls `client.beta.sessions.events.send` with a `user.message` event while the stream is already open.

5. **Print streamed text** — iterates over the event stream and prints `agent.message` text blocks to stdout as they arrive.

6. **Finish on idle** — breaks out of the loop when a `session.status_idle` event arrives with a terminal `stop_reason` (anything other than `requires_action`, which would mean the agent is waiting on a tool result).

7. **Exit on error** — `session.error` events and SDK `APIError` exceptions both print to stderr and call `sys.exit(1)`.

## Event Flow

```
client                         Anthropic
  |                                |
  |-- sessions.create() --------->|
  |<-- session (id, status) ------|
  |                                |
  |-- events.stream() ----------->|  <- open stream first
  |-- events.send(user.message) ->|  <- message includes email context if triggered
  |                                |
  |<-- session.status_running ----|
  |<-- agent.message (text) ------|  <- printed to stdout
  |<-- agent.message (text) ------|
  |<-- session.status_idle -------|  <- loop exits here
```

## Error Handling

| Situation | Behaviour |
|---|---|
| Session creation fails | Prints error to stderr, exits with code 1 |
| `session.error` event received | Prints event to stderr, exits with code 1 |
| Any `anthropic.APIError` | Prints error to stderr, exits with code 1 |
| Email fetch fails | Logs warning, sends message without email context |
| `Ctrl+C` / `KeyboardInterrupt` | Prints "Interrupted." to stderr, exits with code 130 |
