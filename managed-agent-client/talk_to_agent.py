#!/usr/bin/env python3
"""Minimal client for Anthropic Managed Agent agent_015VHQjSZein5oq1iUZr6wSX."""

import os
import subprocess
import sys
import anthropic

AGENT_ID       = "agent_015VHQjSZein5oq1iUZr6wSX"
ENVIRONMENT_ID = "env_01KLJ9UFFW2otvWxy6GAjeyo"

EMAIL_FETCHER  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "outlook_downloader", "fetch_emails_text.ps1")
BRIEF_KEYWORDS = {"brief", "email", "inbox", "mail", "summary", "summarise", "summarize"}


def fetch_emails(hours: int = 24) -> str | None:
    """Run the local PowerShell email fetcher and return its output, or None on failure."""
    try:
        result = subprocess.run(
            ["powershell", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", EMAIL_FETCHER, "-Hours", str(hours)],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
        if result.stderr.strip():
            print(f"[email fetch warning] {result.stderr.strip()}", file=sys.stderr)
    except Exception as e:
        print(f"[email fetch warning] {e}", file=sys.stderr)
    return None


def wants_email_context(message: str) -> bool:
    words = set(message.lower().split())
    return bool(words & BRIEF_KEYWORDS)


def main() -> None:
    user_message = " ".join(sys.argv[1:]) or "Hello! What can you help me with?"

    if wants_email_context(user_message):
        print("Fetching emails from local Outlook...", file=sys.stderr)
        email_data = fetch_emails(hours=24)
        if email_data:
            print(f"Email data collected ({len(email_data)} chars). Sending to agent.", file=sys.stderr)
            user_message = email_data + "\n\n" + user_message
        else:
            print("Could not fetch emails — sending message without email context.", file=sys.stderr)

    client = anthropic.Anthropic()

    try:
        session = client.beta.sessions.create(
            agent=AGENT_ID,
            environment_id=ENVIRONMENT_ID,
        )
    except anthropic.APIConnectionError as e:
        print(f"Connection error (network/firewall?): {e}", file=sys.stderr)
        sys.exit(1)
    except anthropic.AuthenticationError as e:
        print(f"Authentication error (bad API key?): {e}", file=sys.stderr)
        sys.exit(1)
    except anthropic.PermissionDeniedError as e:
        print(f"Permission denied (key lacks access to this beta?): {e}", file=sys.stderr)
        sys.exit(1)
    except anthropic.NotFoundError as e:
        print(f"Not found (agent/environment ID wrong?): {e}", file=sys.stderr)
        sys.exit(1)
    except anthropic.APIStatusError as e:
        print(f"API error {e.status_code}: {e.message}", file=sys.stderr)
        sys.exit(1)
    except anthropic.APIError as e:
        print(f"Unexpected API error ({type(e).__name__}): {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Session {session.id} created.", file=sys.stderr)

    try:
        with client.beta.sessions.events.stream(session_id=session.id) as stream:
            client.beta.sessions.events.send(
                session_id=session.id,
                events=[
                    {
                        "type": "user.message",
                        "content": [{"type": "text", "text": user_message}],
                    }
                ],
            )

            for event in stream:
                if event.type == "agent.message":
                    for block in event.content:
                        if block.type == "text":
                            print(block.text, end="", flush=True)

                elif event.type == "session.status_idle":
                    stop_type = getattr(event.stop_reason, "type", None)
                    if stop_type != "requires_action":
                        print()
                        break

                elif event.type == "session.status_terminated":
                    print()
                    break

                elif event.type == "session.error":
                    print(f"\nSession error: {event}", file=sys.stderr)
                    sys.exit(1)

    except anthropic.APIError as e:
        print(f"\nAPI error: {e}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nInterrupted.", file=sys.stderr)
        sys.exit(130)


if __name__ == "__main__":
    main()
