"""
gmail_client.py - fetches emails from Gmail and converts them into RawEmail objects.

Concepts to be able to explain:

* OAuth 2.0 "installed app" flow: the first run opens a browser, you click Allow,
  and Google gives the script a token. We NEVER see or store your password.
* READ-ONLY scope: the script can read mail but cannot send, delete or modify.
* Gmail search syntax: we reuse Gmail's own query language (subject:, newer_than:)
  so Google's servers do the filtering and we download only candidate emails.
* MIME: an email is a tree of "parts" (plain text, HTML, attachments), with the
  content base64url-encoded. We walk the tree and decode the best text part.

Setup steps are in README.md.
"""
from __future__ import annotations

import base64
import logging
from datetime import datetime
from email.utils import parseaddr
from html.parser import HTMLParser

import config
from models import RawEmail

log = logging.getLogger(__name__)
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def get_service():
    # Imported here so --demo mode works even if the Google libraries aren't installed.
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    creds = None
    if config.TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(config.TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())  # silently get a new access token
        else:
            if not config.CREDENTIALS_PATH.exists():
                raise FileNotFoundError(
                    f"{config.CREDENTIALS_PATH.name} not found. Follow the Gmail setup steps in README.md."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(config.CREDENTIALS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)  # opens your browser once
        config.TOKEN_PATH.write_text(creds.to_json())
    return build("gmail", "v1", credentials=creds)


def build_query(days: int) -> str:
    """e.g.  newer_than:90d subject:("your application" OR "interview" ...)"""
    phrases = " OR ".join(f'"{k}"' for k in config.SUBJECT_KEYWORDS)
    return f"newer_than:{days}d subject:({phrases})"


class _TextExtractor(HTMLParser):
    """Minimal HTML -> text converter (standard library only)."""
    BLOCK = {"br", "p", "div", "tr", "li", "h1", "h2", "h3", "table"}

    def __init__(self):
        super().__init__()
        self.parts, self._skip = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self._skip += 1
        elif tag in self.BLOCK:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self._skip = max(0, self._skip - 1)
        elif tag in self.BLOCK:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self._skip:
            self.parts.append(data)


def html_to_text(html: str) -> str:
    p = _TextExtractor()
    p.feed(html)
    text = "".join(p.parts)
    lines = (" ".join(line.split()) for line in text.splitlines())
    return "\n".join(line for line in lines if line)


def _decode(data: str) -> str:
    return base64.urlsafe_b64decode(data + "=" * (-len(data) % 4)).decode("utf-8", errors="replace")


def _extract_body(payload: dict) -> str:
    """Walk the MIME tree; prefer text/plain, fall back to text/html."""
    plain, html = [], []

    def walk(part):
        mime, data = part.get("mimeType", ""), part.get("body", {}).get("data")
        if data and mime == "text/plain":
            plain.append(_decode(data))
        elif data and mime == "text/html":
            html.append(_decode(data))
        for sub in part.get("parts", []) or []:
            walk(sub)

    walk(payload)
    if plain:
        return "\n".join(plain)
    return html_to_text("\n".join(html))


def _to_raw_email(msg: dict) -> RawEmail:
    headers = {h["name"].lower(): h["value"] for h in msg["payload"].get("headers", [])}
    name, addr = parseaddr(headers.get("from", ""))
    return RawEmail(
        id=msg["id"],
        date=datetime.fromtimestamp(int(msg["internalDate"]) / 1000),  # ms since epoch
        sender_name=name or addr,
        sender_email=addr,
        subject=headers.get("subject", ""),
        body=_extract_body(msg["payload"]),
    )


def fetch_emails(days: int, skip_ids: set[str] | None = None) -> list[RawEmail]:
    """Search Gmail and return matching emails (skipping ones we've already processed)."""
    skip_ids = skip_ids or set()
    service = get_service()
    query = build_query(days)
    log.info("Gmail query: %s", query)

    ids, token = [], None
    while True:  # results come in pages; keep going until there is no next page
        resp = service.users().messages().list(
            userId="me", q=query, maxResults=100, pageToken=token).execute()
        ids += [m["id"] for m in resp.get("messages", [])]
        token = resp.get("nextPageToken")
        if not token:
            break

    new_ids = [i for i in ids if i not in skip_ids]
    log.info("%d emails match, %d are new.", len(ids), len(new_ids))
    return [_to_raw_email(service.users().messages().get(userId="me", id=i, format="full").execute())
            for i in new_ids]
