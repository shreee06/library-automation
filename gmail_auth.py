"""
gmail_auth.py

Handles Google OAuth for Gmail API access. Run once interactively to
authorize; after that, a cached token (token.json) is reused automatically.

Setup required before this works (see SETUP_GMAIL.md):
    1. A Google Cloud project with the Gmail API enabled.
    2. An OAuth 2.0 Client ID (type: Desktop app), downloaded as
       "credentials.json" and placed in this project folder.

SCOPES below only request READ access to Gmail -- this script never
sends, deletes or modifies mail. The Sheets scope is used only by the
optional Google Sheets export (get_sheets_service) to write results to a
spreadsheet you point it at.
"""

from __future__ import annotations

import json
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"

# Gmail is read-only: this project can only READ mail, never send/delete/modify.
# The Sheets scope is only used by the optional --sheet-id export.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]


def _get_credentials(token_path: str | Path | None = None) -> Credentials:
    """
    Loads cached credentials from the token file (token.json by default),
    refreshing or re-running the browser sign-in as needed, and saves the
    result back to the same file.

    A token created before the Sheets scope was added is treated as missing,
    so one extra browser sign-in is requested to grant the new scope.
    """
    creds = None

    if token_path is None:
        token_file = TOKEN_FILE
    else:
        token_file = Path(token_path)
        if not token_file.is_absolute():
            token_file = BASE_DIR / token_file

    if token_file.exists():
        try:
            granted = set(json.loads(token_file.read_text()).get("scopes") or [])
        except ValueError:
            granted = set()
        if set(SCOPES) <= granted:
            creds = Credentials.from_authorized_user_file(
                str(token_file), SCOPES
            )

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                raise FileNotFoundError(
                    f"'{CREDENTIALS_FILE.name}' not found. Follow "
                    f"SETUP_GMAIL.md to create a Google Cloud OAuth "
                    f"client and download it into this folder."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES
            )
            creds = flow.run_local_server(port=0)

        token_file.write_text(creds.to_json())

    return creds


def get_gmail_service(token_path: str | Path | None = None):
    """
    Returns an authorized Gmail API service object. The first run opens a
    browser window to sign in and approve access; after that, the cached
    token in token.json is reused (and refreshed automatically) until it's
    revoked.

    token_path: optional token file to use instead of token.json (e.g.
    "token_original.json"). Relative paths are resolved against this
    project folder.
    """
    return build("gmail", "v1", credentials=_get_credentials(token_path))


def get_sheets_service(token_path: str | Path | None = None):
    """
    Returns an authorized Google Sheets API service object, using the same
    token file and sign-in flow as get_gmail_service().
    """
    return build("sheets", "v4", credentials=_get_credentials(token_path))



if __name__ == "__main__":
    # Quick manual test: python gmail_auth.py
    service = get_gmail_service()
    profile = service.users().getProfile(userId="me").execute()
    print(f"Authorized as: {profile['emailAddress']}")
