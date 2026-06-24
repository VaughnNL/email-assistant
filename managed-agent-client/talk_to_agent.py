#!/usr/bin/env python3
"""Local email assistant — Ollama chat with live Outlook context injection."""

import json
import os
import subprocess
import sys
import urllib.request

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

OLLAMA_URL     = os.environ.get("OLLAMA_URL", "http://localhost:11434")
MODEL          = os.environ.get("OLLAMA_MODEL", "gemma4:latest")
EMAIL_FETCHER  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "outlook_downloader", "fetch_emails_text.ps1")
BRIEF_KEYWORDS = {"brief", "email", "inbox", "mail", "summary", "summarise", "summarize"}

SYSTEM_PROMPT = (
    "You are a personal assistant with access to the user's recent emails when provided. "
    "Give specific, concrete answers. Be concise and direct."
)


def fetch_emails(hours: int = 24) -> str | None:
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
    return bool(set(message.lower().split()) & BRIEF_KEYWORDS)


def stream_chat(messages: list) -> str:
    body = json.dumps({"model": MODEL, "messages": messages, "stream": True}).encode("utf-8")
    req  = urllib.request.Request(
        f"{OLLAMA_URL}/api/chat",
        data=body,
        headers={"Content-Type": "application/json"},
    )
    full = ""
    with urllib.request.urlopen(req, timeout=300) as resp:
        for line in resp:
            if line.strip():
                chunk = json.loads(line.decode("utf-8"))
                token = chunk.get("message", {}).get("content", "")
                if token:
                    print(token, end="", flush=True)
                    full += token
                if chunk.get("done"):
                    break
    print()
    return full


def inject_email_context(message: str) -> str:
    print("Fetching emails from Outlook...", file=sys.stderr)
    email_data = fetch_emails()
    if email_data:
        return email_data + "\n\n" + message
    print("Could not fetch emails — continuing without context.", file=sys.stderr)
    return message


def main() -> None:
    history = [{"role": "system", "content": SYSTEM_PROMPT}]

    # Single-shot mode: message passed as CLI argument
    if len(sys.argv) > 1:
        message = " ".join(sys.argv[1:])
        if wants_email_context(message):
            message = inject_email_context(message)
        history.append({"role": "user", "content": message})
        stream_chat(history)
        return

    # Interactive REPL
    print(f"Assistant ready ({MODEL}) — Ctrl+C to exit\n")
    while True:
        try:
            user_input = input("You: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nBye!")
            break

        if not user_input:
            continue

        message = user_input
        if wants_email_context(message):
            message = inject_email_context(message)

        history.append({"role": "user", "content": message})
        print("Assistant: ", end="", flush=True)
        response = stream_chat(history)
        history.append({"role": "assistant", "content": response})


if __name__ == "__main__":
    main()
