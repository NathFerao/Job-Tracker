"""
excel_store.py - everything that touches the Excel file.

Key ideas worth being able to explain:

1. MATCHING: a rejection email must update the *existing* row for that job, not
   create a new one. We match on a normalised company name + a fuzzy role match.
2. COLOUR CODING USES CONDITIONAL FORMATTING, not hard-coded cell colours. The
   colour is a rule ("if Status = Interview, fill orange"), so if you change a
   status by hand in Excel (the dropdown), the colour updates by itself.
3. STATUS ONLY MOVES FORWARD: an old "Applied" email processed late must never
   overwrite "Interview". config.STATUS_RANK defines the ordering.
"""
from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import config
from config import COL, COLUMNS
from email_parser import detect_region
from models import JobEvent

DATE_FMT = "DD-MMM-YYYY"
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _letter(column_name: str) -> str:
    return get_column_letter(COL[column_name])


# --------------------------------------------------------------------------- #
# Name normalisation for matching
# --------------------------------------------------------------------------- #
_LEGAL = {"pvt", "private", "ltd", "limited", "inc", "llc", "llp", "corp", "co"}


def normalise_org(name: str) -> str:
    """'Nimbus Software Pvt. Ltd.' and 'nimbus software' -> 'nimbus software'."""
    words = re.sub(r"[^a-z0-9 ]", " ", (name or "").lower()).split()
    return " ".join(w for w in words if w not in _LEGAL)


def normalise_role(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (name or "").lower())


def roles_similar(a: str, b: str) -> bool:
    a, b = normalise_role(a), normalise_role(b)
    if not a or not b:
        return True  # nothing to compare -> don't block the match
    return a == b or a in b or b in a or SequenceMatcher(None, a, b).ratio() >= 0.7


def orgs_match(a: str, b: str) -> bool:
    a, b = normalise_org(a), normalise_org(b)
    if not a or not b:
        return False
    return a == b or (min(len(a), len(b)) >= 4 and (a in b or b in a))


def _as_date(value):
    return value.date() if isinstance(value, datetime) else value


class JobTracker:
    def __init__(self, path: Path | str):
        self.path = Path(path)
        if self.path.exists():
            self.wb = load_workbook(self.path)
            if config.APPLICATIONS_SHEET not in self.wb.sheetnames:
                raise ValueError(f"{self.path} has no '{config.APPLICATIONS_SHEET}' sheet.")
        else:
            self.wb = Workbook()
            self.wb.active.title = config.APPLICATIONS_SHEET
            self._build_sheet(self.wb[config.APPLICATIONS_SHEET])
        self.ws = self.wb[config.APPLICATIONS_SHEET]
        self._upgrade_if_needed()

    # ------------------------------------------------------------------ #
    # One-time sheet construction
    # ------------------------------------------------------------------ #
    def _style_header(self, ws, col: int, name: str, width: int) -> None:
        cell = ws.cell(row=1, column=col, value=name)
        cell.font = Font(name=config.FONT_NAME, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="1F3864", end_color="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = width

    def _build_sheet(self, ws) -> None:
        for i, (name, width) in enumerate(COLUMNS, start=1):
            self._style_header(ws, i, name, width)
        ws.row_dimensions[1].height = 32
        ws.freeze_panes = "A2"
        self._apply_sheet_rules(ws)

    def _apply_sheet_rules(self, ws) -> None:
        """Filter, hidden column, dropdowns and colour rules. Safe to call again on an
        existing sheet: it resets the rules first, so nothing is duplicated."""
        last_col = get_column_letter(len(COLUMNS))
        ws.auto_filter.ref = f"A1:{last_col}{config.MAX_ROWS}"
        ws.column_dimensions[_letter("Email ID")].hidden = True

        ws.data_validations.dataValidation = []
        dropdowns = [("Status", config.STATUSES, "status")]
        if config.REGION_SPLIT:
            dropdowns.append(("Region", config.REGIONS, "region"))
        for column, options, title in dropdowns:
            letter = _letter(column)
            dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
            dv.error, dv.errorTitle = f"Pick a {title} from the list.", f"Invalid {title}"
            ws.add_data_validation(dv)
            dv.add(f"{letter}2:{letter}{config.MAX_ROWS}")

        # Colour rules. The strong colour for the Status cell is added FIRST so it
        # wins over the lighter whole-row colour (earlier rule = higher priority).
        ws.conditional_formatting = ConditionalFormattingList()
        status_col = _letter("Status")
        for status, (strong, _light) in config.STATUS_COLORS.items():
            ws.conditional_formatting.add(
                f"{status_col}2:{status_col}{config.MAX_ROWS}",
                FormulaRule(formula=[f'${status_col}2="{status}"'],
                            fill=PatternFill("solid", start_color=strong, end_color=strong),
                            font=Font(bold=True)))
        for status, (_strong, light) in config.STATUS_COLORS.items():
            ws.conditional_formatting.add(
                f"A2:{last_col}{config.MAX_ROWS}",
                FormulaRule(formula=[f'${status_col}2="{status}"'],
                            fill=PatternFill("solid", start_color=light, end_color=light)))

    def _upgrade_if_needed(self) -> None:
        """Keep older workbooks in step with the REGION_SPLIT setting."""
        ws = self.ws
        region_col = len(COLUMNS) + 1 if not config.REGION_SPLIT else COL["Region"]
        has_region = ws.cell(1, region_col).value == "Region"

        if not config.REGION_SPLIT:
            if has_region:                      # feature switched off: remove the column
                ws.delete_cols(region_col)
                self._apply_sheet_rules(ws)
            return

        if has_region:
            return
        self._style_header(ws, region_col, "Region", dict(COLUMNS)["Region"])
        for row in range(2, self.last_row() + 1):
            ws.cell(row, region_col, config.REGION_UNKNOWN)
            self._style_row(row)
        self.refresh_regions()
        self._apply_sheet_rules(ws)

    def refresh_regions(self) -> int:
        """Re-check every row whose Region is blank/Unknown using its Location and Site.
        Never touches a row that already has a region, so manual choices are safe."""
        changed = 0
        if not config.REGION_SPLIT:
            return 0
        for row in range(2, self.last_row() + 1):
            if self._get(row, "Region") in (None, "", config.REGION_UNKNOWN):
                new = detect_region(self._get(row, "Location") or "", self._get(row, "Site Used") or "", "")
                if new != config.REGION_UNKNOWN:
                    self.ws.cell(row, COL["Region"], new)
                    changed += 1
        return changed

    # ------------------------------------------------------------------ #
    # Reading
    # ------------------------------------------------------------------ #
    def last_row(self) -> int:
        """Last row that has an organisation filled in (1 = header only)."""
        row = self.ws.max_row
        while row > 1 and not self.ws.cell(row, COL["Organisation"]).value:
            row -= 1
        return row

    def _get(self, row: int, column: str):
        return self.ws.cell(row, COL[column]).value

    def find_match(self, org: str, role: str | None, strict: bool) -> int | None:
        """Find the row for this job. Newest rows are checked first.

        strict=True  (a new 'Applied' email): require org AND similar role, so a
                     second, different role at the same company is a NEW row.
        strict=False (a status email): if the role doesn't match, fall back to the
                     newest row for that company - status emails often omit the role.
        """
        candidates = [r for r in range(self.last_row(), 1, -1)
                      if orgs_match(org, self._get(r, "Organisation") or "")]
        for r in candidates:
            if roles_similar(role or "", self._get(r, "Role") or ""):
                return r
        return None if strict or not candidates else candidates[0]

    # ------------------------------------------------------------------ #
    # Writing
    # ------------------------------------------------------------------ #
    def apply_event(self, ev: JobEvent, fill_only: bool = False) -> tuple[str, int | None]:
        """Apply one JobEvent. Returns (what happened, row number).

        fill_only=True is used by --reprocess: it only fills BLANK cells of rows that already
        exist. It never adds rows, never changes a status and never touches your edits."""
        org = ev.organisation or ""

        if fill_only:
            row = self.find_match(org, ev.role, strict=False) if org else None
            if row is None:
                return "skipped (no existing row)", None
            self._fill_blanks(row, ev)
            return "enriched", row

        if ev.status == "Applied":
            row = self.find_match(org, ev.role, strict=True) if org else None
            if row:
                self._fill_blanks(row, ev)
                return "duplicate (merged)", row
            return "added", self._append(ev, org or "Unknown")

        # A status update (interview, rejection, ...)
        if not org:
            return "skipped (company unknown)", None
        row = self.find_match(org, ev.role, strict=False)
        if row is None:
            if config.CREATE_ROWS_FROM_STATUS_EMAILS:
                return "added from status email", self._append(ev, org, from_status=True)
            return "skipped (no matching application)", None
        return self._update_status(row, ev), row

    def _append(self, ev: JobEvent, org: str, from_status: bool = False) -> int:
        row = self.last_row() + 1
        values = {
            "Date Applied": None if from_status else ev.event_date,
            "Organisation": org,
            "Location": ev.location or "",
            "Role": ev.role or "",
            "Site Used": ev.site or "",
            "Skills": ev.skills,
            "Status": ev.status,
            "Last Updated": ev.event_date,
            "Job Link": ev.url,
            "Notes": ("Added from a status email - application date unknown."
                      if from_status else ("Company name not found - please check."
                                           if org == "Unknown" else "")),
            "Email ID": ev.email_id,
        }
        if config.REGION_SPLIT:
            values["Region"] = ev.region
        for name, value in values.items():
            self.ws.cell(row, COL[name], value)
        b, h, i = _letter("Date Applied"), _letter("Status"), _letter("Last Updated")
        self.ws.cell(row, COL["Days to Response"],
                     f'=IF(OR({b}{row}="",{h}{row}="Applied",{h}{row}="No Response"),"",{i}{row}-{b}{row})')
        self._style_row(row)
        if from_status:
            self._add_note(row, f"[{ev.event_date:%d %b}] {ev.status}: {ev.subject}")
        return row

    def _update_status(self, row: int, ev: JobEvent) -> str:
        current = self._get(row, "Status") or "Applied"
        if config.STATUS_RANK[ev.status] < config.STATUS_RANK.get(current, 0):
            return f"ignored (already '{current}')"
        self.ws.cell(row, COL["Status"], ev.status)
        self.ws.cell(row, COL["Last Updated"], ev.event_date)
        self._add_note(row, f"[{ev.event_date:%d %b}] {ev.status}: {ev.subject}")
        self._fill_blanks(row, ev)
        return f"updated -> {ev.status}"

    def _fill_blanks(self, row: int, ev: JobEvent) -> None:
        """Enrich a row with anything a later email tells us that we didn't know."""
        for column, value in (("Location", ev.location), ("Role", ev.role),
                              ("Skills", ev.skills), ("Job Link", ev.url)):
            if value and not self._get(row, column):
                self.ws.cell(row, COL[column], value)
        if config.REGION_SPLIT:
            current = self._get(row, "Region")
            if ev.region != config.REGION_UNKNOWN and current in (None, "", config.REGION_UNKNOWN):
                self.ws.cell(row, COL["Region"], ev.region)

    def _add_note(self, row: int, text: str) -> None:
        cell = self.ws.cell(row, COL["Notes"])
        if cell.value and text in str(cell.value):
            return  # already logged (keeps re-runs from duplicating the timeline)
        cell.value = f"{cell.value}\n{text}" if cell.value else text

    def _style_row(self, row: int) -> None:
        for name, _ in COLUMNS:
            cell = self.ws.cell(row, COL[name])
            cell.font = Font(name=config.FONT_NAME, size=10)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=name in {"Skills", "Notes", "Role", "Organisation", "Location", "Site Used"},
                                       horizontal="center" if name in {
                                           "Sr. No.", "Date Applied", "Last Updated",
                                           "Days to Response", "Status", "Region"} else None)
        for name in ("Date Applied", "Last Updated"):
            self.ws.cell(row, COL[name]).number_format = DATE_FMT

    def mark_stale(self, today: date, days: int = config.GHOST_AFTER_DAYS) -> int:
        """Flag jobs with no activity for `days` days as 'No Response'."""
        cutoff, count = today - timedelta(days=days), 0
        for row in range(2, self.last_row() + 1):
            if self._get(row, "Status") in ("Applied", "Under Review"):
                last = _as_date(self._get(row, "Last Updated"))
                if last and last < cutoff:
                    self.ws.cell(row, COL["Status"], "No Response")
                    self._add_note(row, f"[{today:%d %b}] Auto-flagged: no activity for {days}+ days.")
                    count += 1
        return count

    def save(self) -> None:
        for row in range(2, self.last_row() + 1):  # keep Sr. No. sequential
            self.ws.cell(row, COL["Sr. No."], row - 1)
        self.wb.save(self.path)
