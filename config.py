"""
config.py - every setting you might want to tweak lives here.

Keeping configuration in one place means you can change behaviour (colours,
which job sites you recognise, which skills to look for) without touching the
logic in the other modules.
"""
from pathlib import Path

# --------------------------------------------------------------------------- #
# Paths
# --------------------------------------------------------------------------- #
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "job_applications.xlsx"
DEMO_EXCEL_PATH = BASE_DIR / "demo_job_applications.xlsx"
STATE_PATH = BASE_DIR / "state.json"            # remembers which emails we've processed
CREDENTIALS_PATH = BASE_DIR / "credentials.json"  # downloaded from Google Cloud Console
TOKEN_PATH = BASE_DIR / "token.json"            # created automatically after first login

# --------------------------------------------------------------------------- #
# Gmail search
# --------------------------------------------------------------------------- #
LOOKBACK_DAYS = 90  # how far back to search on a run

# Gmail only returns emails whose SUBJECT contains one of these phrases.
# This keeps us from downloading (and reading) your whole inbox.
SUBJECT_KEYWORDS = [
    "your application", "thank you for applying", "thanks for applying",
    "application received", "application sent", "application update",
    "application status", "applied", "update from", "interview",
    "assessment", "offer", "next steps", "viewed",
]

# Only the first N characters of a body are analysed. Job-board emails append
# long "similar jobs" footers that would pollute skills/location extraction.
MAX_BODY_CHARS = 3000

# --------------------------------------------------------------------------- #
# Statuses and colours
# --------------------------------------------------------------------------- #
# Order matters: this is the order shown in the dropdown and on the Stats sheet.
STATUSES = [
    "Applied", "Under Review", "Assessment", "Interview",
    "Offer", "Rejected", "No Response",
]

# status -> (strong colour for the Status cell, light colour for the whole row)
STATUS_COLORS = {
    "Applied":      ("9DC3E6", "DDEBF7"),  # blue
    "Under Review": ("C9B8E8", "EDE7F6"),  # purple
    "Assessment":   ("FFE699", "FFF2CC"),  # yellow
    "Interview":    ("F8CBAD", "FCE4D6"),  # orange
    "Offer":        ("A9D08E", "E2EFDA"),  # green
    "Rejected":     ("FF9B9B", "FDE2E2"),  # red
    "No Response":  ("BFBFBF", "EDEDED"),  # grey
}

# How "advanced" each status is. A status email can only move a job FORWARD
# (e.g. Applied -> Interview), never backwards. "No Response" is -1 because a
# late reply should be allowed to revive a job we had given up on.
STATUS_RANK = {
    "No Response": -1, "Applied": 0, "Under Review": 1,
    "Assessment": 2, "Interview": 3, "Rejected": 4, "Offer": 5,
}

# Jobs with no activity for this many days are flagged "No Response".
GHOST_AFTER_DAYS = 30

# If a rejection/interview email arrives for a job that isn't in the sheet yet
# (e.g. you applied before the lookback window), add a row for it anyway?
CREATE_ROWS_FROM_STATUS_EMAILS = True

# --------------------------------------------------------------------------- #
# Recognising where an email came from
# --------------------------------------------------------------------------- #
# sender domain -> friendly name. Matching is "domain or any subdomain of it".
JOB_SITES = {
    "linkedin.com": "LinkedIn",
    "naukri.com": "Naukri",
    "indeed.com": "Indeed",
    "foundit.in": "Foundit",
    "monsterindia.com": "Foundit",
    "instahyre.com": "Instahyre",
    "internshala.com": "Internshala",
    "wellfound.com": "Wellfound",
    "glassdoor.com": "Glassdoor",
    # Applicant-tracking systems: the *company's* own careers page, run by a vendor.
    "greenhouse.io": "Company Website (Greenhouse)",
    "greenhouse-mail.io": "Company Website (Greenhouse)",
    "lever.co": "Company Website (Lever)",
    "myworkday.com": "Company Website (Workday)",
    "smartrecruiters.com": "Company Website (SmartRecruiters)",
    "ashbyhq.com": "Company Website (Ashby)",
    "zohorecruit.com": "Company Website (Zoho Recruit)",
}
DEFAULT_SITE = "Company Website"

# Job boards send mail ON BEHALF of companies, so their sender name is the
# board's name, not the employer's. We must not use it as the company name.
JOB_BOARDS = {"LinkedIn", "Naukri", "Indeed", "Foundit", "Instahyre",
              "Internshala", "Wellfound", "Glassdoor"}

# --------------------------------------------------------------------------- #
# Skills to look for (matched against the email text)
# --------------------------------------------------------------------------- #
SKILLS = [
    "Python", "Java", "JavaScript", "TypeScript", "C++", "C#", "Golang", "Rust",
    "SQL", "NoSQL", "PostgreSQL", "MySQL", "MongoDB", "Redis",
    "Django", "Flask", "FastAPI", "React", "Node.js", "Spring Boot",
    "AWS", "Azure", "GCP", "Docker", "Kubernetes", "Git", "Linux", "CI/CD",
    "REST APIs", "GraphQL", "Kafka", "Spark", "Airflow", "ETL", "Hadoop",
    "Machine Learning", "Deep Learning", "NLP", "TensorFlow", "PyTorch",
    "Pandas", "NumPy", "Scikit-learn",
    "Tableau", "Power BI", "Excel", "Data Analysis", "Data Engineering",
    "HTML", "CSS", "Selenium", "Agile",
]
# These words are also ordinary English, so only match them with exact casing.
CASE_SENSITIVE_SKILLS = {"Excel", "Spark"}

# --------------------------------------------------------------------------- #
# Workbook layout (shared by excel_store.py and stats.py)
# --------------------------------------------------------------------------- #
APPLICATIONS_SHEET = "Applications"
# Set to True to add a Region column (India / Outside India) and separate stats sheets for each.
# Off by default: most job-application emails don't contain the job's location, so most rows
# end up "Unknown" and the split isn't useful unless you fill locations in by hand.
REGION_SPLIT = False
REGION_INDIA, REGION_OUTSIDE, REGION_UNKNOWN = "India", "Outside India", "Unknown"
REGIONS = [REGION_INDIA, REGION_OUTSIDE, REGION_UNKNOWN]
# (sheet name, region filter or None for everything)
if REGION_SPLIT:
    STATS_SHEETS = [("Stats - All", None), ("Stats - India", REGION_INDIA),
                    ("Stats - Outside India", REGION_OUTSIDE)]
else:
    STATS_SHEETS = [("Stats", None)]
# Every dashboard name we might have created in the past; all are removed before a rebuild.
ALL_STATS_NAMES = ["Stats", "Stats - All", "Stats - India", "Stats - Outside India"]
MAX_ROWS = 2000  # formulas / dropdowns / colour rules cover rows 2..MAX_ROWS

# (header, column width). Your original 7 columns come first, unchanged.
COLUMNS = [
    ("Sr. No.", 8),
    ("Date Applied", 14),
    ("Organisation", 26),
    ("Location", 24),
    ("Role", 32),
    ("Site Used", 26),
    ("Skills", 38),
    # --- extra columns I added ---
    ("Status", 15),
    ("Last Updated", 14),
    ("Days to Response", 12),
    ("Job Link", 30),
    ("Notes", 48),
    ("Email ID", 18),  # hidden; lets us trace a row back to the Gmail message
]
if REGION_SPLIT:
    COLUMNS.append(("Region", 14))  # India / Outside India / Unknown (appended last so older files upgrade cleanly)
COL = {name: i for i, (name, _) in enumerate(COLUMNS, start=1)}
FONT_NAME = "Arial"

# --------------------------------------------------------------------------- #
# Region detection (India vs outside India)
# --------------------------------------------------------------------------- #
# If the job's location mentions any of these, it is in India. Add your own.
INDIA_PLACES = [
    "india", "bengaluru", "bangalore", "mumbai", "navi mumbai", "delhi", "new delhi", "noida",
    "gurgaon", "gurugram", "hyderabad", "secunderabad", "chennai", "pune", "kolkata", "ahmedabad",
    "goa", "panaji", "margao", "madgaon", "vasco", "porvorim", "jaipur", "kochi", "cochin",
    "thiruvananthapuram", "trivandrum", "coimbatore", "chandigarh", "indore", "nagpur", "lucknow",
    "bhubaneswar", "vadodara", "surat", "mysuru", "mysore", "visakhapatnam", "mangalore", "nashik",
    "faridabad", "ghaziabad", "karnataka", "maharashtra", "kerala", "tamil nadu", "telangana",
    "gujarat", "rajasthan", "uttar pradesh", "haryana", "west bengal", "andhra pradesh",
    "madhya pradesh", "odisha",
]
# Used only when the location is missing or just "Remote": these sites are India-focused.
INDIA_SITES = {"Naukri", "Foundit", "Instahyre", "Internshala"}
