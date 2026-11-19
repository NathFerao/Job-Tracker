"""
models.py - the two data shapes that flow through the pipeline.

    RawEmail  --(parser.py)-->  JobEvent  --(excel_store.py)-->  a row in Excel

Using small dataclasses instead of loose dictionaries means each stage has a
clear contract: the Gmail client only has to produce RawEmail objects, and the
Excel writer only has to understand JobEvent objects. You could swap Gmail for
Outlook without touching anything downstream.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime


@dataclass
class RawEmail:
    """An email exactly as fetched, before any interpretation."""
    id: str
    date: datetime
    sender_name: str
    sender_email: str
    subject: str
    body: str


@dataclass
class JobEvent:
    """One meaningful thing that happened to a job application."""
    email_id: str
    event_date: date
    status: str                      # one of config.STATUSES
    organisation: str | None = None
    role: str | None = None
    location: str | None = None
    region: str = "Unknown"          # India / Outside India / Unknown
    site: str | None = None
    skills: str = ""                 # comma separated
    url: str = ""
    subject: str = ""                # kept for the Notes column
