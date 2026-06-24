#!/usr/bin/env python3
"""
Outlook Email Downloader
Connects to local Outlook via COM and saves emails organized by day.

Usage:
    python outlook_downloader.py --hours 18
    python outlook_downloader.py --count 20
    python outlook_downloader.py --unread
    python outlook_downloader.py --hours 24 --unread
    python outlook_downloader.py --count 20 --unread --output C:\\EmailBackup
    python outlook_downloader.py --hours 12 --folder "Sent Items"
"""

import os
import re
import datetime
import argparse
from pathlib import Path

try:
    import win32com.client
except ImportError:
    print("Missing dependency. Run:  pip install pywin32")
    raise SystemExit(1)


def sanitize_filename(text: str, max_len: int = 80) -> str:
    text = re.sub(r'[<>:"/\\|?*\r\n\t]', "_", text or "")
    text = text.strip(". ")
    return text[:max_len] or "untitled"


def com_to_datetime(com_time) -> datetime.datetime:
    """Convert a pywintypes COM datetime to a naive local datetime."""
    # pywintypes.datetime is UTC-aware; convert to local naive
    utc = datetime.datetime(
        com_time.year, com_time.month, com_time.day,
        com_time.hour, com_time.minute, com_time.second,
        tzinfo=datetime.timezone.utc,
    )
    return utc.astimezone().replace(tzinfo=None)


WELL_KNOWN_FOLDERS = {
    "inbox":       6,
    "sent":        5,
    "sent items":  5,
    "drafts":      16,
    "deleted":     3,
    "junk":        23,
    "junk email":  23,
}


def get_outlook_folder(namespace, folder_name: str):
    key = folder_name.strip().lower()
    if key in WELL_KNOWN_FOLDERS:
        return namespace.GetDefaultFolder(WELL_KNOWN_FOLDERS[key])
    # Search all stores by name
    for store in namespace.Stores:
        try:
            root = store.GetRootFolder()
            for sub in root.Folders:
                if sub.Name.strip().lower() == key:
                    return sub
        except Exception:
            continue
    print(f"Folder '{folder_name}' not found — falling back to Inbox.")
    return namespace.GetDefaultFolder(6)


def download_emails(
    hours: int = None,
    count: int = None,
    unread_only: bool = False,
    output_dir: str = "emails",
    folder_name: str = "Inbox",
):
    print("Connecting to Outlook...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        namespace.Logon()
    except Exception as exc:
        print(f"Could not connect to Outlook: {exc}")
        print("Make sure Microsoft Outlook is installed and you are signed in.")
        return

    folder = get_outlook_folder(namespace, folder_name)
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # newest first

    # MAPI Restrict for unread (simple boolean — reliable across locales)
    if unread_only:
        messages = messages.Restrict("[UnRead] = True")
        messages.Sort("[ReceivedTime]", True)

    cutoff: datetime.datetime | None = None
    if hours is not None:
        cutoff = datetime.datetime.now() - datetime.timedelta(hours=hours)

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    print(f"\nFolder      : {folder.Name}")
    if cutoff:
        print(f"Time window : past {hours} hour(s)  (since {cutoff.strftime('%Y-%m-%d %H:%M')})")
    if count:
        print(f"Limit       : {count} email(s)")
    if unread_only:
        print(f"Filter      : unread only")
    print(f"Output      : {output_path.resolve()}\n")

    # Per-day name sets to avoid collisions
    seen: dict[Path, set] = {}

    downloaded = 0
    scanned = 0

    for message in messages:
        if count is not None and downloaded >= count:
            break
        if scanned > 10_000:
            print("Reached scan limit of 10 000 items — stopping.")
            break
        scanned += 1

        try:
            received = com_to_datetime(message.ReceivedTime)
        except Exception:
            continue

        # Messages are newest-first; stop as soon as we pass the cutoff
        if cutoff is not None and received < cutoff:
            break

        try:
            subject = sanitize_filename(message.Subject)
            sender  = sanitize_filename(message.SenderName)
            ts      = received.strftime("%H%M%S")

            day_folder = output_path / received.strftime("%Y-%m-%d")
            day_folder.mkdir(parents=True, exist_ok=True)
            if day_folder not in seen:
                seen[day_folder] = set()

            base = f"{ts}_{sender}_{subject}"
            filename = f"{base}.msg"
            n = 1
            while filename in seen[day_folder]:
                filename = f"{base}_{n}.msg"
                n += 1
            seen[day_folder].add(filename)

            message.SaveAs(str(day_folder / filename), 3)  # 3 = olMSG

            downloaded += 1
            read_flag = "" if message.UnRead else " [read]"
            print(f"  [{downloaded:>4}] {received.strftime('%Y-%m-%d %H:%M')} | "
                  f"{(message.Subject or '(no subject)')[:60]}{read_flag}")

        except Exception as exc:
            print(f"  Warning: could not save item — {exc}")
            continue

    print(f"\n{'─'*60}")
    print(f"Done. {downloaded} email(s) saved to '{output_path.resolve()}'")
    if downloaded == 0:
        print("No emails matched the given criteria.")


def main():
    parser = argparse.ArgumentParser(
        description="Download emails from local Outlook, organised by day.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples
--------
  python outlook_downloader.py --hours 18
  python outlook_downloader.py --count 20
  python outlook_downloader.py --unread
  python outlook_downloader.py --hours 24 --unread
  python outlook_downloader.py --count 20 --unread --output C:\\EmailBackup
  python outlook_downloader.py --hours 12 --folder "Sent Items"
        """,
    )
    parser.add_argument("--hours",  type=int, metavar="N",
                        help="Emails from the past N hours")
    parser.add_argument("--count",  type=int, metavar="N",
                        help="Last N emails (newest first)")
    parser.add_argument("--unread", action="store_true",
                        help="Unread emails only")
    parser.add_argument("--output", type=str, default="emails",
                        help="Output folder (default: ./emails)")
    parser.add_argument("--folder", type=str, default="Inbox",
                        help="Outlook folder to pull from (default: Inbox)")

    args = parser.parse_args()

    if not any([args.hours, args.count, args.unread]):
        parser.print_help()
        print("\nError: specify at least one of --hours, --count, or --unread")
        return

    download_emails(
        hours=args.hours,
        count=args.count,
        unread_only=args.unread,
        output_dir=args.output,
        folder_name=args.folder,
    )


if __name__ == "__main__":
    main()
