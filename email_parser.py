"""
parser.py - turns a RawEmail into a JobEvent (or None if it isn't job related).

This is a RULE-BASED parser: regular expressions plus keyword lists. It is
deliberately simple and transparent - you can read every rule and explain why
an email was classified the way it was. (See README for how you could swap this
for an LLM later.)

The pipeline for one email:
    1. Is it noise (job alert, newsletter)?          -> ignore
    2. What STATUS is it (applied / interview / ...)?  -> classify_status()
    3. Which company and role is it about?             -> extract_org_role()
    4. Location, skills, site, link                    -> small helpers
"""
from __future__ import annotations

import re
from typing import Optional

import config
from models import JobEvent, RawEmail

# --------------------------------------------------------------------------- #
# 1. Noise filter
# --------------------------------------------------------------------------- #
IGNORE_SUBJECT = re.compile(
    r"job alert|jobs? for you|new jobs?|recommended|jobs you may|similar jobs|"
    r"newsletter|digest|webinar|top picks|people also viewed",
    re.I,
)

# --------------------------------------------------------------------------- #
# 2. Status classification
# --------------------------------------------------------------------------- #
# ORDER MATTERS. We test the most specific/decisive statuses first. A rejection
# email often contains the words "interview" or "offer" ("we will not be
# extending an offer"), so "Rejected" must be checked before them.
STATUS_RULES: list[tuple[str, list[str]]] = [
    ("Rejected", [
        r"unfortunately", r"regret to inform", r"not (?:be )?moving forward",
        r"not (?:been )?selected", r"decided to (?:move|proceed) (?:forward )?with (?:other|another)",
        r"will not be (?:proceeding|moving)", r"position has been filled",
        r"unable to (?:offer|move forward)",
    ]),
    ("Offer", [
        r"pleased to offer", r"delighted to offer", r"offer of employment",
        r"offer letter", r"extend (?:you )?an offer", r"job offer",
    ]),
    ("Interview", [
        r"interview invitation", r"invit\w* you to (?:an? )?(?:\w+ )?interview",
        r"schedule (?:an?|your|the) (?:\w+ )?interview", r"interview (?:is )?scheduled",
        r"interview with", r"next round", r"speak with you", r"schedule a (?:call|chat|meeting)",
    ]),
    ("Assessment", [
        r"assessment", r"coding (?:test|challenge)", r"take-home", r"online test",
        r"hackerrank", r"codility", r"complete (?:the|this) (?:test|task)",
    ]),
    ("Under Review", [
        r"viewed your application", r"application (?:was|has been) viewed",
        r"reviewing your application", r"application is (?:under|being) review",
    ]),
    ("Applied", [
        r"thank(?:s| you) for (?:applying|your application)",
        r"application (?:received|submitted|sent)", r"we(?:'ve| have) received your application",
        r"successfully applied", r"you applied", r"your application (?:was|has been) sent",
        r"indeed application", r"thank you for your interest",
        r"^(?:\w+, )?your application (?:to|for|was sent)",   # anchored: subject only
    ]),
]
_COMPILED_RULES = [(s, [re.compile(p) for p in pats]) for s, pats in STATUS_RULES]

# Confirmation emails say things like "IF selected, we will schedule an
# interview". That sentence is conditional, not an event, so we drop any
# sentence containing these words before matching the body.
CONDITIONAL = re.compile(r"\b(?:if|should|in case|once|when|whether)\b")
SENTENCE_SPLIT = re.compile(r"(?<=[.!?])\s+|\n+")


def classify_status(subject: str, body: str) -> Optional[str]:
    """Return one of config.STATUSES, or None if the email isn't about an application.

    The SUBJECT is checked before the body because it is the most reliable
    signal: a confirmation email whose body happens to say "unfortunately we
    cannot reply to every applicant" should still be 'Applied'.
    """
    subject_l = subject.lower().strip()
    sentences = [s for s in SENTENCE_SPLIT.split(body.lower()) if not CONDITIONAL.search(s)]
    body_l = " ".join(sentences)

    for text in (subject_l, body_l):
        for status, patterns in _COMPILED_RULES:
            if any(p.search(text) for p in patterns):
                return status
    return None


# --------------------------------------------------------------------------- #
# 3. Company and role extraction
# --------------------------------------------------------------------------- #
# A company name = text up to the next punctuation mark.
COMP = r"[^.,!|\n–—:()]{2,60}"
# A role name = text that can't cross a sentence boundary (so "interest in X. We will
# contact you at Y" can never be mis-read as role="X. We will contact you", company="Y").
ROLE = r"[^\n.,!|:]{3,70}?"

# (pattern, subject_only). Patterns that capture BOTH role and company come first.
ROLE_AND_COMPANY = [
    # "Your application to Data Analyst at Zeta Analytics"
    # "Thank you for applying for the Software Engineer position at Orbit Robotics"
    # "Thank you for your interest in the Data Analyst position at Zeta Analytics"
    (re.compile(
        rf"(?:application(?:\s+sent)?\s+(?:to|for)|appl(?:ying|ied)\s+(?:to|for)|interest\s+in)\s+"
        rf"(?:the\s+|a\s+|an\s+)?(?P<role>{ROLE})\s+(?:(?:position|role|job|opening)\s+)?"
        rf"(?:at|with)\s+(?P<company>{COMP})", re.I), False),
    # "...the position of Data Engineer at Finlytics"
    (re.compile(
        rf"(?:position|role)\s+of\s+(?:the\s+)?(?P<role>{ROLE})\s+(?:at|with)\s+(?P<company>{COMP})",
        re.I), False),
    # "Interview Invitation – Python Developer at Nimbus Software"  (subject only: too loose for bodies)
    (re.compile(
        rf"(?:\s[–—-]\s|:\s)(?P<role>[^–—:\n.,]{{3,60}}?)\s+at\s+(?P<company>{COMP})", re.I), True),
]

# Patterns that only give us the company.
COMPANY_ONLY = [
    re.compile(rf"(?:sent|submitted)\s+to\s+(?P<company>{COMP})", re.I),
    re.compile(rf"(?:applying|applied)\s+(?:to|at|with)\s+(?:the\s+)?(?P<company>{COMP})", re.I),
    re.compile(rf"interest\s+in\s+(?:the\s+)?(?P<company>{COMP})", re.I),
    re.compile(rf"update\s+from\s+(?P<company>{COMP})", re.I),
    re.compile(rf"viewed\s+by\s+(?P<company>{COMP})", re.I),
    re.compile(rf"interview\s+with\s+(?P<company>{COMP})", re.I),
]

# "Indeed Application: Junior Backend Engineer" -> role only
ROLE_ONLY = [re.compile(r"indeed application:\s*(?P<role>[^\n]{3,80})", re.I)]

LEGAL_SUFFIX = re.compile(r"(?:\s+(?:pvt|private|ltd|limited|inc|llc|llp|corp|co))+\.?$", re.I)
TRAILING_WORDS = re.compile(r"\s+(?:has|have|is|was|will|because|where)\b.*$", re.I)
NOT_A_COMPANY = {"you", "your email", "us", "this role", "the role", "this position"}


def clean_company(raw: str | None) -> str | None:
    if not raw:
        return None
    name = TRAILING_WORDS.sub("", raw.strip(" .,:;!-–—|\"'"))
    name = re.sub(r"^the\s+", "", name, flags=re.I)
    name = LEGAL_SUFFIX.sub("", name).strip(" .,:;!-–—|\"'")
    if len(name) < 2 or name.lower() in NOT_A_COMPANY or name.lower().startswith("your "):
        return None
    return name


def clean_role(raw: str | None) -> str | None:
    if not raw:
        return None
    role = raw.strip(" .,:;!-–—|\"'")
    return role if len(role) >= 3 else None


def extract_org_role(subject: str, body: str) -> tuple[str | None, str | None]:
    """Try patterns on the subject first, then the body, most specific first."""
    org = role = None

    # Pass 1: patterns that give both at once (they're the most trustworthy).
    for is_subject, text in ((True, subject), (False, body)):
        for pattern, subject_only in ROLE_AND_COMPANY:
            if subject_only and not is_subject:
                continue
            m = pattern.search(text)
            if m:
                org = org or clean_company(m.group("company"))
                role = role or clean_role(m.group("role"))
        if org and role:
            return org, role

    # Pass 2: fill in whatever is still missing.
    if not org:
        for text in (subject, body):
            for pattern in COMPANY_ONLY:
                m = pattern.search(text)
                if m and clean_company(m.group("company")):
                    org = clean_company(m.group("company"))
                    break
            if org:
                break
    if not role:
        for text in (subject, body):
            for pattern in ROLE_ONLY:
                m = pattern.search(text)
                if m:
                    role = clean_role(m.group("role"))
                    break
            if role:
                break
    return org, role


def company_from_sender(sender_name: str, sender_email: str, site: str) -> str | None:
    """Last-resort guess: the 'From' display name, e.g. 'Orbit Robotics Recruiting'."""
    if site in config.JOB_BOARDS:
        return None  # LinkedIn/Naukri write on behalf of employers; the sender isn't the employer
    name = re.sub(
        r"\b(careers?|recruit(?:ing|ment|ers?)|talent|hr|human resources|team|jobs?|hiring|"
        r"notifications?|no-?reply|acquisition|via\s+\w+)\b", "", sender_name, flags=re.I)
    name = re.sub(r"\s+", " ", name).strip(" -–—|,")
    if len(name) >= 2:
        return clean_company(name)
    if site.startswith("Company Website ("):
        return None  # ATS vendor domain (e.g. greenhouse.io) says nothing about the employer
    parts = sender_email.split("@")[-1].lower().split(".")
    if len(parts) >= 3 and parts[-2] in {"co", "com", "org", "net"}:
        label = parts[-3]
    elif len(parts) >= 2:
        label = parts[-2]
    else:
        return None
    return label.capitalize()


# --------------------------------------------------------------------------- #
# 4. Small helpers: site, location, skills, link
# --------------------------------------------------------------------------- #
def detect_site(sender_email: str) -> str:
    """Map the sender's domain to a friendly job-site name."""
    domain = sender_email.split("@")[-1].lower()
    for key, name in config.JOB_SITES.items():
        if domain == key or domain.endswith("." + key):
            return name
    return config.DEFAULT_SITE


LOCATION_PATTERNS = [
    re.compile(r"(?:job\s+)?location\s*[:\-]\s*(?P<loc>[^\n|]{2,60})", re.I),
    # LinkedIn style line: "Zeta Analytics · Bengaluru, Karnataka, India"
    re.compile(r"^.+?[·•]\s*(?P<loc>[A-Z][\w .'-]+(?:,\s*[\w .'-]+){0,2})\s*$", re.M),
]


def extract_location(text: str) -> str:
    for pattern in LOCATION_PATTERNS:
        m = pattern.search(text)
        if m:
            return re.sub(r"\s+", " ", m.group("loc")).strip(" .,;")
    if re.search(r"\bremote\b", text, re.I):
        return "Remote"
    return ""


def _skill_regex(skill: str) -> re.Pattern:
    # (?<![\w+#]) / (?![\w+#]) = "not glued to another letter, + or #".
    # Plain \b would let "Java" match inside "JavaScript" and break on "C++".
    flags = 0 if skill in config.CASE_SENSITIVE_SKILLS else re.I
    return re.compile(rf"(?<![\w+#]){re.escape(skill)}(?![\w+#])", flags)


_SKILL_PATTERNS = [(s, _skill_regex(s)) for s in config.SKILLS]


def extract_skills(text: str) -> str:
    return ", ".join(skill for skill, rx in _SKILL_PATTERNS if rx.search(text))


URL_RE = re.compile(r"https?://[^\s<>\")\]]+")


def extract_link(text: str) -> str:
    for url in URL_RE.findall(text):
        url = url.rstrip(".,;")
        if re.search(r"job|career|apply|view", url, re.I):
            return url.split("?")[0]  # drop tracking parameters
    return ""


# --------------------------------------------------------------------------- #
# 5. Region: India vs outside India
# --------------------------------------------------------------------------- #
_INDIA_RE = re.compile(r"\b(?:" + "|".join(re.escape(x) for x in config.INDIA_PLACES) + r")\b", re.I)
_WORK_MODE = re.compile(r"\b(?:remote|hybrid|on-?site|work from home|wfh|anywhere)\b|[()\-,/]", re.I)
_INDIA_MONEY = re.compile(r"₹|\bINR\b|\bLPA\b|\blakhs?\b", re.I)


def detect_region(location: str, site: str, sender_email: str, text: str = "") -> str:
    """Decide India / Outside India / Unknown. Strongest evidence first:
    1. a real place name in the location   2. weaker hints when the job is just 'Remote'.
    'Unknown' is a legitimate answer - you can set the Region cell by hand."""
    if location:
        if _INDIA_RE.search(location):
            return config.REGION_INDIA
        if _WORK_MODE.sub("", location).strip():   # something left besides "Remote/Hybrid"
            return config.REGION_OUTSIDE
    if site in config.INDIA_SITES or sender_email.lower().endswith(".in") or _INDIA_MONEY.search(text):
        return config.REGION_INDIA
    return config.REGION_UNKNOWN


# --------------------------------------------------------------------------- #
# Public entry point
# --------------------------------------------------------------------------- #
def parse_email(mail: RawEmail) -> Optional[JobEvent]:
    """RawEmail -> JobEvent, or None if the email isn't about a job application."""
    if IGNORE_SUBJECT.search(mail.subject):
        return None

    body = mail.body[: config.MAX_BODY_CHARS]
    status = classify_status(mail.subject, body)
    if status is None:
        return None

    site = detect_site(mail.sender_email)
    org, role = extract_org_role(mail.subject, body)
    org = org or company_from_sender(mail.sender_name, mail.sender_email, site)

    text = f"{mail.subject}\n{body}"
    location = extract_location(body)
    return JobEvent(
        email_id=mail.id,
        event_date=mail.date.date(),
        status=status,
        organisation=org,
        role=role,
        location=location,
        region=detect_region(location, site, mail.sender_email, text),
        site=site,
        skills=extract_skills(text),
        url=extract_link(body),
        subject=mail.subject,
    )
