"""
gmail_extractor.py

Fetches emails from Gmail -- subject, sender, To, Cc, date, body, and any
PDF/Word attachments -- and saves everything locally:

    email_input/<original>_<sender>[_<cc>].<ext>      -- saved attachments
    output/email_metadata.xlsx                         -- one row per email

This is the FRONT END of the pipeline: it only captures raw email data.
It does NOT yet try to pull student name/reg no/department out of the
subject or body, because that depends on the email format the library
ends up requiring (see the open question in the architecture doc). Once
that's settled, a small email_parser.py can read these rows the same way
filename_parser.py reads filenames today.

A processed_message_ids.txt file tracks what's already been fetched, so
re-running this script only pulls NEW mail -- same idea as needed for the
eventual 30k-email run, just already built in now.

Setup:
    pip install -r requirements.txt
    Follow SETUP_GMAIL.md to get credentials.json, then run once to
    authorize in your browser.

Usage:
    python gmail_extractor.py                    # fetch 10 most recent emails
    python gmail_extractor.py --max 50            # fetch 50
    python gmail_extractor.py --subject "ai check" --max 5   # subject phrase
    python gmail_extractor.py --query "from:someone@x.com"   # override filter

By default only mail matching "has:attachment filename:pdf" is fetched.
"""

from __future__ import annotations

import argparse
import base64
import re
from datetime import datetime
from email.utils import getaddresses
from pathlib import Path

from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from email_parser import parse_email_body
from gmail_auth import get_gmail_service

BASE_DIR = Path(__file__).resolve().parent
ATTACHMENT_DIR = BASE_DIR / "email_input"
METADATA_LOG = BASE_DIR / "output" / "email_metadata.xlsx"
PROCESSED_IDS_FILE = BASE_DIR / "output" / "processed_message_ids.txt"

ALLOWED_ATTACHMENT_EXT = {".pdf", ".docx", ".doc"}

DEFAULT_QUERY = "has:attachment filename:pdf"

# Saved attachments are named <original>_<sender>[_<first cc>].<ext>
MAX_NAME_PART = 40  # cap for the sender / cc pieces
MAX_ORIGINAL_STEM = 80  # cap for the original filename piece

METADATA_HEADERS = [
    "Fetched At",
    "Message ID",
    "Date",
    "From",
    "To",
    "Cc",
    "Subject",
    "Body (first 500 chars)",
    "Attachments",
    "Attachment Count",
    "SRN",
    "Student Name",
    "Mentor",
    "Status",
    "Review Reason",
]


def _get_header(headers: list[dict], name: str) -> str:
    for h in headers:
        if h["name"].lower() == name.lower():
            return h["value"]
    return ""


def _decode_body_part(data: str) -> str:
    return base64.urlsafe_b64decode(data.encode("UTF-8")).decode(
        "UTF-8", errors="replace"
    )


def _extract_body(payload: dict) -> str:
    """
    Walk a Gmail message payload for the best available body text,
    preferring text/plain and falling back to a stripped text/html.
    """
    plain_text = None
    html_text = None

    def walk(part: dict):
        nonlocal plain_text, html_text
        mime = part.get("mimeType", "")
        body = part.get("body", {})
        if mime == "text/plain" and body.get("data"):
            plain_text = _decode_body_part(body["data"])
        elif mime == "text/html" and body.get("data"):
            html_text = _decode_body_part(body["data"])
        for sub in part.get("parts", []) or []:
            walk(sub)

    walk(payload)

    if plain_text:
        return plain_text.strip()
    if html_text:
        return BeautifulSoup(html_text, "html.parser").get_text(
            separator="\n"
        ).strip()
    return ""


def _safe_part(text: str, max_len: int = MAX_NAME_PART) -> str:
    """Make one piece of a filename safe: replace unsafe chars, trim, cap length."""
    cleaned = re.sub(r"_+", "_", re.sub(r"[^\w.\-]", "_", text))
    return cleaned[:max_len].strip("_.")


def _address_label(header_value: str) -> str:
    """
    Display name (else email address) of the FIRST address in a From/Cc
    header, made filename-safe. Empty string if the header has none.
    """
    for name, addr in getaddresses([header_value]):
        label = _safe_part(name.strip() or addr.strip())
        if label:
            return label
    return ""


def _attachment_filename(original: str, sender: str, cc: str) -> tuple[str, str]:
    """
    Returns (stem, extension) for <original>_<sender>[_<cc>]. The Cc part
    (the mentor) is left out entirely when there is no Cc.
    """
    original_path = Path(original)
    parts = [
        _safe_part(original_path.stem, MAX_ORIGINAL_STEM),
        _address_label(sender),
        _address_label(cc),
    ]
    stem = "_".join(p for p in parts if p)
    return stem, original_path.suffix.lower()


def _resolve_target(stem: str, ext: str, data: bytes, message_id: str):
    """
    Pick where to save `data` without ever overwriting a different file.

    Returns (path, needs_write). If a file with the target name already
    exists with identical bytes it is reused. If the bytes differ, the new
    file gets the first 6 chars of the message ID appended (then _2, _3...).
    """
    candidate = ATTACHMENT_DIR / f"{stem}{ext}"
    if not candidate.exists():
        return candidate, True
    if candidate.read_bytes() == data:
        return candidate, False

    suffixed = f"{stem}_{message_id[:6]}"
    candidate = ATTACHMENT_DIR / f"{suffixed}{ext}"
    n = 2
    while candidate.exists():
        if candidate.read_bytes() == data:
            return candidate, False
        candidate = ATTACHMENT_DIR / f"{suffixed}_{n}{ext}"
        n += 1
    return candidate, True


def _save_attachments(
    service, message_id: str, payload: dict, sender: str, cc: str
) -> list[str]:
    """
    Downloads any PDF/Word attachments on this message and returns their
    saved filenames, named <original>_<sender>[_<cc>].<ext>.
    """
    saved: list[str] = []

    def walk(part: dict):
        filename = part.get("filename", "")
        body = part.get("body", {})
        ext = Path(filename).suffix.lower()

        if filename and ext in ALLOWED_ATTACHMENT_EXT:
            attachment_id = body.get("attachmentId")
            data = None
            if attachment_id:
                att = (
                    service.users()
                    .messages()
                    .attachments()
                    .get(userId="me", messageId=message_id, id=attachment_id)
                    .execute()
                )
                data = base64.urlsafe_b64decode(att["data"])
            elif body.get("data"):
                data = base64.urlsafe_b64decode(body["data"])

            if data:
                stem, out_ext = _attachment_filename(filename, sender, cc)
                out_path, needs_write = _resolve_target(
                    stem, out_ext, data, message_id
                )
                if needs_write:
                    out_path.write_bytes(data)
                saved.append(out_path.name)

        for sub in part.get("parts", []) or []:
            walk(sub)

    walk(payload)
    return saved


def _ensure_metadata_log() -> None:
    if METADATA_LOG.exists():
        return
    METADATA_LOG.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Emails"
    ws.append(METADATA_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(METADATA_LOG)


def _append_metadata_row(row: list) -> None:
    _ensure_metadata_log()
    wb = load_workbook(METADATA_LOG)
    ws = wb["Emails"]
    ws.append(row)
    wb.save(METADATA_LOG)


def _load_processed_ids() -> set[str]:
    if not PROCESSED_IDS_FILE.exists():
        return set()
    return set(PROCESSED_IDS_FILE.read_text().splitlines())


def _mark_processed(message_id: str) -> None:
    PROCESSED_IDS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(PROCESSED_IDS_FILE, "a") as f:
        f.write(message_id + "\n")


def fetch_emails(max_results: int = 10, query: str = DEFAULT_QUERY) -> None:
    ATTACHMENT_DIR.mkdir(exist_ok=True)
    service = get_gmail_service()
    processed_ids = _load_processed_ids()

    response = (
        service.users()
        .messages()
        .list(userId="me", maxResults=max_results, q=query)
        .execute()
    )
    messages = response.get("messages", [])

    if not messages:
        print("No emails matched. (Try a broader --query, or none at all.)")
        return

    print(f"Found {len(messages)} email(s). Fetching details...")

    for msg_ref in messages:
        msg_id = msg_ref["id"]

        if msg_id in processed_ids:
            print(f"  [SKIP] {msg_id} already fetched.")
            continue

        full = (
            service.users()
            .messages()
            .get(userId="me", id=msg_id, format="full")
            .execute()
        )
        payload = full.get("payload", {})
        headers = payload.get("headers", [])

        subject = _get_header(headers, "Subject")
        sender = _get_header(headers, "From")
        to = _get_header(headers, "To")
        cc = _get_header(headers, "Cc")
        date = _get_header(headers, "Date")
        body = _extract_body(payload)
        attachments = _save_attachments(service, msg_id, payload, sender, cc)
        student = parse_email_body(body)

        _append_metadata_row(
            [
                datetime.now().isoformat(timespec="seconds"),
                msg_id,
                date,
                sender,
                to,
                cc,
                subject,
                body[:500],
                "; ".join(attachments),
                len(attachments),
                student.srn,
                student.student_name,
                student.mentor,
                "Needs Review" if student.needs_review else "OK",
                student.review_reason,
            ]
        )
        _mark_processed(msg_id)

        print(
            f'  [OK] {msg_id}: "{subject}" from {sender} '
            f"({len(attachments)} attachment(s))"
        )

    print(f"\nDone. Metadata: {METADATA_LOG}")
    print(f"Attachments saved to: {ATTACHMENT_DIR}/")


if __name__ == "__main__":
    arg_parser = argparse.ArgumentParser(description="Fetch emails from Gmail.")
    arg_parser.add_argument(
        "--max", type=int, default=10, help="Max emails to fetch (default 10)"
    )
    arg_parser.add_argument(
        "--query",
        type=str,
        default=DEFAULT_QUERY,
        help=f'Gmail search query (default: "{DEFAULT_QUERY}"), '
        'e.g. "from:someone@x.com". Pass "" for no filter.',
    )
    arg_parser.add_argument(
        "--subject",
        type=str,
        default="",
        help='Only emails whose subject contains this phrase, e.g. "ai check". '
        "Added to --query (or the default filter).",
    )
    args = arg_parser.parse_args()

    query = args.query
    if args.subject:
        phrase = args.subject.replace('"', "")
        query = f'{query} subject:"{phrase}"'.strip()

    fetch_emails(max_results=args.max, query=query)
