# Job Application Tracker

A Python tool that reads your Gmail, spots job-application emails (confirmations,
interview invites, rejections, offers), and keeps an Excel workbook up to date:
one colour-coded row per application, plus a statistics dashboard.

```
 Gmail ──► gmail_client.py ──► email_parser.py ──► excel_store.py ──► job_applications.xlsx
 (API)      RawEmail            JobEvent             row add/update     ├─ Applications (colour-coded)
                                                                         └─ Stats (live formulas + charts)
```

## 1. Quick start (no Gmail needed)

```bash
python -m venv .venv && source .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python main.py --demo --as-of 2026-11-10              # builds demo_job_applications.xlsx from fake emails
python -m unittest -v                                  # run the tests
```

Open `demo_job_applications.xlsx`. Formulas calculate when Excel opens the file
(some phone/web previewers show empty Stats cells; use desktop Excel or LibreOffice).

## 2. Connect your real Gmail (one-time, ~10 minutes)

1. Go to <https://console.cloud.google.com> and create a project (e.g. "job-tracker").
2. **APIs & Services → Library →** search "Gmail API" → **Enable**.
3. **APIs & Services → OAuth consent screen →** User type **External** → fill in the app name and
   your email → under **Test users** add your own Gmail address.
4. **APIs & Services → Credentials → Create credentials → OAuth client ID →** type **Desktop app**.
   Click **Download JSON** and save it in this folder as `credentials.json`.
5. Run `python main.py`. A browser opens, you sign in and click Allow. A `token.json` is saved.

Notes
* The scope requested is `gmail.readonly`: the script **cannot** send, delete or change email.
* While the Google app is in "Testing" mode the token expires after about 7 days. If you get an
  auth error, delete `token.json` and run again.
* Everything runs on your own computer. No email content is sent to any third party.
* Never commit `credentials.json`, `token.json` or your real `.xlsx` (the `.gitignore` already blocks them).
* Close the Excel file before running, otherwise Windows blocks the save.

Run it whenever you like (or schedule it daily with Task Scheduler on Windows / `cron` on Mac/Linux).
It remembers processed emails in `state.json`, so re-running is safe.

## 3. What's in the workbook

**Applications sheet.** Your 7 columns first, then the extras:

| Column | Purpose |
|---|---|
| Sr. No., Date Applied, Organisation, Location, Role, Site Used, Skills | as you specified |
| Status | Applied / Under Review / Assessment / Interview / Offer / Rejected / No Response (dropdown) |
| Last Updated | date of the latest email about this job |
| Days to Response | formula: Last Updated − Date Applied (blank while still waiting) |
| Job Link | link to the posting, if the email contains one |
| Notes | automatic timeline, e.g. `[20 Sep] Interview: Interview Invitation – ...` |
| Email ID (hidden) | Gmail message id, so a row can be traced back to its email |

Colour code (a bold shade on the Status cell, a light shade across the row):
blue = Applied, purple = Under Review, yellow = Assessment, orange = Interview,
green = Offer, red = Rejected, grey = No Response (no activity for 30 days).
If you change a status by hand using the dropdown, the colour changes automatically.

**Optional: India / outside India split (off by default).** Most job-application emails do not
contain the job's location, so in practice most rows came out "Unknown" and the split wasn't useful.
It is therefore switched off (`REGION_SPLIT = False` in `config.py`). If you turn it on, each job gets a Region. The location is checked against a list of Indian cities,
states and the word "India" (`INDIA_PLACES` in `config.py`); any other real place counts as Outside India.
If the job is just "Remote" the tool falls back to weaker hints (India-focused sites such as Naukri,
a `.in` sender domain, salary in ₹/INR/LPA), and otherwise says **Unknown** rather than guessing.
You can fix any Region with the dropdown and the tool will never overwrite your choice. Use the
filter arrow on the Region header to see only one group.
Every run re-checks rows still marked Unknown (so a location you type in by hand is picked up). To
fill gaps on rows from older runs, run `python main.py --reprocess` once: it re-reads your emails but
only fills *blank* cells of existing rows. It never adds rows, changes statuses or overwrites edits.

**Stats sheet.** One combined dashboard (with the split on you get three: All, India, Outside India). It contains: total applications, active applications, response / interview / offer / rejection
rates, average days to a response, applications in the last 7 / 30 days, counts by status, a
"which job site works best" table (response rate per site), applications per month, and the
top requested skills, with four charts.

## 4. How it works, module by module

### `config.py` — settings in one place
Statuses, colours, job-site domains, skills list, search keywords. Change behaviour here without
touching logic. This is the "configuration vs code" separation.

### `models.py` — the data contract
`RawEmail` (what Gmail gave us) and `JobEvent` (what we understood). Each stage only knows these
two shapes, so Gmail could be replaced by Outlook by writing one new function that returns
`RawEmail` objects.

### `gmail_client.py` — getting the data
* **OAuth 2.0:** the browser login gives the script a revocable token instead of your password.
* **Server-side filtering:** the query `newer_than:90d subject:("your application" OR ...)` runs on
  Google's servers, so we download only candidate emails.
* **Pagination:** results arrive 100 at a time; we follow `nextPageToken` until it runs out.
* **MIME decoding:** an email is a tree of parts, with bodies base64url-encoded. We walk the tree,
  prefer `text/plain`, and fall back to converting HTML to text.

### `email_parser.py` — understanding the data (the interesting part)
It is **rule-based** (regular expressions and keyword lists) so every decision is explainable.

1. **Noise filter**: job alerts and newsletters are dropped.
2. **Status classification** with deliberate design decisions:
   * *Rule order matters.* Rejections are tested before offers/interviews, because rejection
     emails often contain the word "offer" or "interview".
   * *Subject is checked before body.* Confirmation emails often include boilerplate like
     "unfortunately we can't reply to everyone".
   * *Conditional sentences are removed.* "**If** selected, we will schedule an interview" is not
     an interview invitation, so sentences containing if/should/once/when are dropped before the
     body is matched.
3. **Company/role extraction**: patterns such as `application to {role} at {company}` are tried on
   the subject, then the body. If those fail, fall back to the sender's display name (but never
   for LinkedIn/Naukri, because those write *on behalf of* the employer). Legal suffixes
   ("Pvt Ltd") are stripped so the same company always looks the same.
4. **Site** = lookup of the sender's domain in `JOB_SITES`.
5. **Skills** = keyword matching with `(?<![\w+#])skill(?![\w+#])` instead of `\b`, so "Java"
   doesn't match inside "JavaScript" and "C++" / "C#" work.
6. Only the first 3000 characters of a body are read, because job-board emails end with
   "similar jobs" lists that would add wrong skills and locations.

### `excel_store.py` — updating the workbook
* **Matching:** a rejection must update the existing row, not create a new one. We match on a
  normalised company name (lower-case, punctuation and legal suffixes removed) and a fuzzy role
  comparison (`difflib.SequenceMatcher`). A *new application* needs company **and** role to match
  (so a second role at the same company gets its own row); a *status email* falls back to the
  newest row for that company, since status emails often omit the role.
* **De-duplication:** the same job confirmed by LinkedIn and by the company's own email is merged.
* **Status only moves forward** (`STATUS_RANK`): a late "Applied" email can't overwrite "Offer".
* **Conditional formatting** instead of painted cells, so colours follow the Status value.
* **Data validation** gives the Status dropdown.
* **`mark_stale`** flags jobs with 30+ days of silence as "No Response".

### `stats.py` — the dashboards
One function builds every dashboard: the regional versions (if enabled) just append an extra criteria pair
(`, Region, "India"`) to each formula. Every number is an Excel formula (`COUNTIFS`, `AVERAGEIFS`, `EDATE`)
pointing at the Applications sheet. Edit a status by hand and the dashboard and charts update instantly. Python is
only used to decide the layout (which sites/months exist). One subtle formula: to count "Java"
without counting "JavaScript", the skills formula matches the four positions a whole item can
occupy in the comma-separated list (only, first, middle, last).

### `main.py` — orchestration
Fetch → sort **oldest first** (so "Applied" is processed before the "Interview" that follows) →
parse → update → flag stale → rebuild stats → save. The processed-email IDs are written to
`state.json` **only after a successful save**, so a crash or an open Excel file never loses emails.
Processing is *idempotent*: running it twice gives the same result as once.

### `test_tracker.py`
23 unit tests covering the tricky cases above (conditional sentences, rejection vs offer wording,
Java vs JavaScript, duplicates, status never going backwards, stale flagging).

## 5. Known limitations (say these in an interview, it shows maturity)

* **Rules are brittle.** A new email template can be misclassified. Mitigation: tests, and the
  "Unknown" company flag with a note asking you to check it.
* **Confirmation emails rarely contain the full job description**, so Skills is often partial. A
  future version could follow the job link and extract skills from the posting.
* **Rates use the latest status only.** A job that reached "Interview" and was then rejected shows
  as Rejected, so "interview rate" slightly under-counts. A fix: a second sheet logging every
  status change (an event history).
* **Applications made outside email** (e.g. via a phone call) won't appear. Add them by hand; on existing rows the
  tool only fills in *blank* cells and updates Status / Last Updated / Notes, so your other edits are safe.

## 6. Ideas to extend it

* **Job location** is the missing piece for any regional analysis. Emails rarely include it. Options:
  type it into the Location cell, or fetch the posting from the Job Link (job sites often block bots,
  so this is fragile). Then set `REGION_SPLIT = True`.
* Replace or back up the regex parser with an **LLM call** that returns structured JSON
  (company, role, status), keeping the rules as a fast first pass.
* Add **Outlook** support via Microsoft Graph (just a new function that returns `RawEmail`).
* Add an **event-history sheet** and a funnel chart (Applied → Responded → Interview → Offer).
* Build a **Streamlit** dashboard on top of the same data.
* **Schedule** a daily run and send yourself a summary ("2 new rejections, 1 interview invite").

## 7. How to present this project

**One-line description (CV):**
> Built a Python automation that reads job-application emails through the Gmail API (OAuth 2.0),
> classifies them and extracts company, role, location and skills with regex-based NLP, and
> maintains a colour-coded Excel tracker with a live analytics dashboard.

**Resume bullets (pick 3):**
* Designed a multi-stage pipeline (fetch → parse → store → analyse) with clean module boundaries and
  typed data models, so the email source can be swapped without touching downstream code.
* Implemented rule-based email classification with ordered rules, conditional-sentence filtering and
  company-name normalisation, handling duplicates and out-of-order emails.
* Made processing idempotent using persisted message IDs; wrote 23 unit tests for edge cases.
* Generated an Excel dashboard using formula-driven KPIs, conditional formatting and charts
  (openpyxl), so metrics update live when statuses are edited.
* Used least-privilege OAuth (`gmail.readonly`), local-only processing and secret-safe repo hygiene.

**60-second pitch:**
"I was tracking applications by hand and it didn't scale, so I automated it. A script logs into my
Gmail through the official API with read-only access, searches for application-related subjects, and
decodes the email. A parser decides if each email is a confirmation, interview invite, rejection or
offer, and pulls out the company, role, location and skills. Another module finds the matching row in
my Excel sheet, or adds one, and updates its status, which drives a colour code. A second sheet has
formulas and charts for response rate, interviews per site and so on. It's idempotent and tested, and
I know where it breaks: it's rule-based, so I documented the limits and how an LLM could replace it."

**Questions you should be ready for**

| Question | Short answer |
|---|---|
| Why regex instead of ML/LLM? | Explainable, free, offline, private, fast to test. Emails from ATS systems are templated, so rules work well. An LLM would handle unusual wording but adds cost, latency and sends personal data to a third party. |
| How do you avoid duplicates? | Two layers: processed Gmail message IDs in `state.json`, and company+role matching so the same job from two sites merges. |
| How is a rejection linked to the right job? | Normalised company name plus fuzzy role similarity; status emails fall back to the newest row for that company. |
| What if the script crashes mid-run? | State is saved only after the Excel file saves, so unprocessed emails are retried next run. |
| Is it secure? | Read-only scope, OAuth instead of passwords, local processing, secrets in `.gitignore`. |
| Why formulas rather than Python-computed stats? | The dashboard stays live when a status is edited by hand. |
| How would you scale it? | Move storage to SQLite/Postgres, use Gmail push notifications (Pub/Sub) instead of polling, and add a web UI. |
