"""
email_address_extractor.py

Searches Gmail for messages matching a query and records who they were
between -- Subject, From (name + email), To, Cc, Date -- into:

    output/extracted_emails.xlsx
        "Emails" sheet            -- one row per matched message
        "Unique Addresses" sheet  -- every distinct address seen in
                                     From/To/Cc, with how many matched
                                     messages it appeared in
        "External Contacts" sheet -- Unique Addresses minus the library's
                                     own mailbox (original@pes.edu)

An output/extracted_processed_ids.txt file tracks which messages have
already been recorded, so re-running only adds NEW mail (the Unique
Addresses sheet is rebuilt from the Emails sheet each run).

Read-only on Gmail: never sends, deletes, labels or modifies mail. With
--sheet-id it also writes both tabs to that Google Sheet (clearing and
rewriting them), which is the only thing it ever modifies.

Setup:
    pip install -r requirements.txt
    Follow SETUP_GMAIL.md to get credentials.json, then run once to
    authorize in your browser.

Usage:
    python email_address_extractor.py                       # default query, 100 msgs
    python email_address_extractor.py --max 500             # up to 500 messages
    python email_address_extractor.py --query "from:someone@x.com"
    python email_address_extractor.py --token token_original.json
    python email_address_extractor.py --sheet-id <ID from the Sheet's URL>

By default the query is:
    ("SoC" OR "System-on-Chip" OR "Analog Memory" OR "Photonics"
     OR "Security" OR "Accelerators" OR "PQC" OR "Post-Quantum")
"""

from __future__ import annotations

import argparse
from collections import Counter
from email.utils import getaddresses, parseaddr
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from gmail_auth import get_gmail_service, get_sheets_service

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_FILE = BASE_DIR / "output" / "extracted_emails.xlsx"
PROCESSED_IDS_FILE = BASE_DIR / "output" / "extracted_processed_ids.txt"

EMAILS_SHEET = "Emails"
UNIQUE_SHEET = "Unique Addresses"
EXTERNAL_SHEET = "External Contacts"

# The library's own mailbox: kept in Emails/Unique Addresses, but left out of
# External Contacts since that tab is the invite list.
EXCLUDED_ADDRESSES = {"original@pes.edu"}

DEFAULT_QUERY = (
    '("SoC" OR "System-on-Chip" OR "Analog Memory" OR "Photonics" '
    'OR "Security" OR "Accelerators" OR "PQC" OR "Post-Quantum")'
)

EMAIL_HEADERS = [
    "Message ID",
    "Date",
    "Subject",
    "From Name",
    "From Email",
    "To",
    "Cc",
]
UNIQUE_HEADERS = ["Email Address", "Name", "Message Count"]

# Gmail's messages.list returns at most 500 results per page.
PAGE_SIZE_CAP = 500


def _get_header(headers: list[dict], name: str) -> str:
    for h in headers:
        if h["name"].lower() == name.lower():
            return h["value"]
    return ""


def _bold_header_row(ws) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True)


def _ensure_workbook() -> None:
    if OUTPUT_FILE.exists():
        return
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = EMAILS_SHEET
    ws.append(EMAIL_HEADERS)
    _bold_header_row(ws)
    wb.save(OUTPUT_FILE)


def _append_rows(rows: list[list]) -> None:
    _ensure_workbook()
    wb = load_workbook(OUTPUT_FILE)
    ws = wb[EMAILS_SHEET]
    for row in rows:
        ws.append(row)
    wb.save(OUTPUT_FILE)


def _rebuild_unique_sheet() -> int:
    """
    Recompute the Unique Addresses sheet from the Emails sheet: each address
    is counted once per message it appears in (From/To/Cc combined).
    Returns the number of unique addresses.
    """
    wb = load_workbook(OUTPUT_FILE)
    ws = wb[EMAILS_SHEET]
    col = {name: i for i, name in enumerate(EMAIL_HEADERS)}

    def _clean_name(name: str, addr: str) -> str:
        name = (name or "").strip()
        # A "name" that is just the address itself isn't a real name.
        return "" if name.lower() == addr else name

    counts: Counter[str] = Counter()
    sender_names: dict[str, Counter[str]] = {}
    other_names: dict[str, Counter[str]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        from_name = row[col["From Name"]] or ""
        from_email = row[col["From Email"]] or ""
        to = row[col["To"]] or ""
        cc = row[col["Cc"]] or ""
        addrs: set[str] = set()

        sender = from_email.strip().lower()
        if sender:
            addrs.add(sender)
            name = _clean_name(from_name, sender)
            if name:
                sender_names.setdefault(sender, Counter())[name] += 1

        for name, addr in getaddresses([to, cc]):
            addr = addr.strip().lower()
            if not addr:
                continue
            addrs.add(addr)
            name = _clean_name(name, addr)
            if name:
                other_names.setdefault(addr, Counter())[name] += 1
        counts.update(addrs)

    def _best_name(addr: str) -> str:
        # Sender name wins; otherwise the most common To/Cc name; else blank.
        for source in (sender_names, other_names):
            if addr in source:
                return source[addr].most_common(1)[0][0]
        return ""

    if UNIQUE_SHEET in wb.sheetnames:
        del wb[UNIQUE_SHEET]
    uws = wb.create_sheet(UNIQUE_SHEET)
    uws.append(UNIQUE_HEADERS)
    _bold_header_row(uws)
    for addr, n in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])):
        uws.append([addr, _best_name(addr), n])
    wb.save(OUTPUT_FILE)
    return len(counts)


def _rebuild_external_sheet() -> int:
    """
    Derive the External Contacts sheet from Unique Addresses: same columns,
    minus EXCLUDED_ADDRESSES (case-insensitive). Emails and Unique Addresses
    are left untouched. Returns the number of external addresses.
    """
    wb = load_workbook(OUTPUT_FILE)
    uws = wb[UNIQUE_SHEET]

    if EXTERNAL_SHEET in wb.sheetnames:
        del wb[EXTERNAL_SHEET]
    ews = wb.create_sheet(EXTERNAL_SHEET)
    ews.append(UNIQUE_HEADERS)
    _bold_header_row(ews)

    count = 0
    for addr, name, n in uws.iter_rows(min_row=2, values_only=True):
        if not addr or str(addr).strip().lower() in EXCLUDED_ADDRESSES:
            continue
        ews.append([addr, name, n])
        count += 1
    wb.save(OUTPUT_FILE)
    return count


def _read_sheet_rows(sheet_name: str) -> list[list]:
    """All rows (header included) of a tab in the local workbook, as strings/numbers."""
    wb = load_workbook(OUTPUT_FILE, read_only=True)
    ws = wb[sheet_name]
    return [
        ["" if v is None else v for v in row]
        for row in ws.iter_rows(values_only=True)
    ]


def _export_to_google_sheet(sheet_id: str, token_path: str | None) -> str:
    """
    Writes the Emails, Unique Addresses and External Contacts tabs of the
    local workbook into
    the Google Sheet `sheet_id` (clearing each tab first, creating missing
    tabs). Returns the Sheet's URL.
    """
    sheets = get_sheets_service(token_path)
    spreadsheets = sheets.spreadsheets()

    meta = spreadsheets.get(
        spreadsheetId=sheet_id, fields="spreadsheetUrl,sheets.properties.title"
    ).execute()
    existing = {s["properties"]["title"] for s in meta.get("sheets", [])}

    tabs = (EMAILS_SHEET, UNIQUE_SHEET, EXTERNAL_SHEET)
    missing = [t for t in tabs if t not in existing]
    if missing:
        spreadsheets.batchUpdate(
            spreadsheetId=sheet_id,
            body={
                "requests": [
                    {"addSheet": {"properties": {"title": t}}} for t in missing
                ]
            },
        ).execute()

    for title in tabs:
        rng = f"'{title}'"
        spreadsheets.values().clear(
            spreadsheetId=sheet_id, range=rng, body={}
        ).execute()
        spreadsheets.values().update(
            spreadsheetId=sheet_id,
            range=f"{rng}!A1",
            valueInputOption="RAW",
            body={"values": _read_sheet_rows(title)},
        ).execute()

    return meta.get(
        "spreadsheetUrl",
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit",
    )


def _load_processed_ids() -> set[str]:
    if not PROCESSED_IDS_FILE.exists():
        return set()
    return set(PROCESSED_IDS_FILE.read_text().splitlines())


def _mark_processed(message_ids: list[str]) -> None:
    PROCESSED_IDS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(PROCESSED_IDS_FILE, "a") as f:
        for message_id in message_ids:
            f.write(message_id + "\n")


def _list_message_ids(service, query: str, max_results: int) -> list[str]:
    ids: list[str] = []
    page_token = None
    while len(ids) < max_results:
        response = (
            service.users()
            .messages()
            .list(
                userId="me",
                q=query,
                maxResults=min(PAGE_SIZE_CAP, max_results - len(ids)),
                pageToken=page_token,
            )
            .execute()
        )
        ids.extend(m["id"] for m in response.get("messages", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            break
    return ids


def extract_addresses(
    max_results: int = 100,
    query: str = DEFAULT_QUERY,
    token_path: str | None = None,
    sheet_id: str | None = None,
) -> None:
    service = get_gmail_service(token_path)
    processed_ids = _load_processed_ids()

    message_ids = _list_message_ids(service, query, max_results)
    if not message_ids:
        print("No emails matched. (Try a broader --query.)")
        return

    print(f"Found {len(message_ids)} email(s). Fetching headers...")

    new_rows: list[list] = []
    new_ids: list[str] = []

    for msg_id in message_ids:
        if msg_id in processed_ids:
            print(f"  [SKIP] {msg_id} already extracted.")
            continue

        full = (
            service.users()
            .messages()
            .get(
                userId="me",
                id=msg_id,
                format="metadata",
                metadataHeaders=["Subject", "From", "To", "Cc", "Date"],
            )
            .execute()
        )
        headers = full.get("payload", {}).get("headers", [])

        subject = _get_header(headers, "Subject")
        sender = _get_header(headers, "From")
        from_name, from_email = parseaddr(sender)

        new_rows.append(
            [
                msg_id,
                _get_header(headers, "Date"),
                subject,
                from_name,
                from_email,
                _get_header(headers, "To"),
                _get_header(headers, "Cc"),
            ]
        )
        new_ids.append(msg_id)
        print(f'  [OK] {msg_id}: "{subject}" from {from_email or sender}')

    if new_rows:
        _append_rows(new_rows)
        _mark_processed(new_ids)
    elif not OUTPUT_FILE.exists():
        return

    unique_count = _rebuild_unique_sheet()
    external_count = _rebuild_external_sheet()

    print(f"\nDone. {len(new_rows)} new message(s) added.")
    print(f"{unique_count} unique address(es), {external_count} external. "
          f"Workbook: {OUTPUT_FILE}")

    if sheet_id:
        url = _export_to_google_sheet(sheet_id, token_path)
        print(f"Google Sheet updated: {url}")


if __name__ == "__main__":
    arg_parser = argparse.ArgumentParser(
        description="Extract Subject/From/To/Cc/Date from matching Gmail "
        "messages into an Excel workbook (read-only)."
    )
    arg_parser.add_argument(
        "--max", type=int, default=100, help="Max emails to scan (default 100)"
    )
    arg_parser.add_argument(
        "--query",
        type=str,
        default=DEFAULT_QUERY,
        help=f"Gmail search query (default: {DEFAULT_QUERY})",
    )
    arg_parser.add_argument(
        "--token",
        type=str,
        default="token.json",
        help='OAuth token file to use (default "token.json"), '
        'e.g. "token_original.json". Relative to the project folder.',
    )
    arg_parser.add_argument(
        "--sheet-id",
        type=str,
        default=None,
        help="Also write the Emails, Unique Addresses and External Contacts "
        "tabs to the Google Sheet with this ID (the long ID in its URL). "
        "Each tab is cleared and rewritten.",
    )
    args = arg_parser.parse_args()

    extract_addresses(
        max_results=args.max,
        query=args.query,
        token_path=args.token,
        sheet_id=args.sheet_id,
    )
