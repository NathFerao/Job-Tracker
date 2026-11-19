"""
stats.py - builds the dashboards: "Stats - All", "Stats - India", "Stats - Outside India".

Design choices worth explaining:

* Every number is an Excel FORMULA (COUNTIFS, AVERAGEIFS...) pointing at the
  Applications sheet, so editing a status or region by hand updates everything.
* ONE function builds all three dashboards. The only difference is an extra
  criteria pair appended to each formula: ..., Region, "India". For the "All"
  sheet that pair is simply empty. Same code, three views - no copy-paste.
* Python only decides the layout (which sites/months/skills exist).
"""
from __future__ import annotations

from collections import Counter
from datetime import date, datetime

from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import config
from config import COL

F = config.FONT_NAME
HEADER_FILL = PatternFill("solid", start_color="1F3864", end_color="1F3864")
SECTION_FILL = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
RESPONDED = ("Assessment", "Interview", "Offer", "Rejected")  # statuses meaning "a human replied"


def _rng(column_name: str) -> str:
    """Absolute range of one Applications column, e.g. 'Applications'!$H$2:$H$2000."""
    letter = get_column_letter(COL[column_name])
    return f"'{config.APPLICATIONS_SHEET}'!${letter}$2:${letter}${config.MAX_ROWS}"


def _section(ws, row: int, text: str) -> None:
    for col in range(1, 5):
        ws.cell(row, col).fill = SECTION_FILL
    ws.cell(row, 1, text).font = Font(name=F, bold=True, size=11)


def _table_header(ws, row: int, labels: list[str]) -> None:
    for i, label in enumerate(labels, start=1):
        c = ws.cell(row, i, label)
        c.font = Font(name=F, bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")


def _style_chart(ch, title, anchor, width=16, height=7.5):
    ch.title, ch.legend = title, None
    ch.width, ch.height = width, height
    ch.x_axis.delete = False  # openpyxl 3.1 hides axes unless told otherwise
    ch.y_axis.delete = False
    ch.anchor = anchor


def _read_applications(ws_apps, region):
    """Collect the distinct sites / months / skills present (only for this region)."""
    sites, months, skills = Counter(), Counter(), Counter()
    for row in range(2, ws_apps.max_row + 1):
        if not ws_apps.cell(row, COL["Organisation"]).value:
            continue
        if region and ws_apps.cell(row, COL["Region"]).value != region:
            continue
        site = ws_apps.cell(row, COL["Site Used"]).value
        if site:
            sites[site] += 1
        applied = ws_apps.cell(row, COL["Date Applied"]).value
        if isinstance(applied, (datetime, date)):
            months[date(applied.year, applied.month, 1)] += 1
        for s in (ws_apps.cell(row, COL["Skills"]).value or "").split(","):
            if s.strip():
                skills[s.strip()] += 1
    return sites, months, skills


def _build_one(wb, sheet_name: str, region: str | None, today: date) -> None:
    ws = wb.create_sheet(sheet_name)
    sites, months, skills = _read_applications(wb[config.APPLICATIONS_SHEET], region)

    # The extra criteria pair that turns an "All" formula into a region-only one.
    rx = f',{_rng("Region")},"{region}"' if region else ""
    label = region if region else "all regions"
    title = f"Job Search Dashboard - {label}" if config.REGION_SPLIT else "Job Search Dashboard"

    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDE", (52, 14, 14, 16, 3)):
        ws.column_dimensions[col].width = w
    ws["A1"] = title
    ws["A1"].font = Font(name=F, bold=True, size=16, color="1F3864")
    ws["A2"] = f"Refreshed {today:%d %b %Y}. All figures are live formulas on the Applications sheet."
    ws["A2"].font = Font(name=F, italic=True, size=9, color="595959")

    # ---- row plan (computed first so KPI formulas can point at the status table) ----
    kpi_start = 5
    status_hdr = kpi_start + 9 + 2
    status_first = status_hdr + 1
    status_row = {s: status_first + i for i, s in enumerate(config.STATUSES)}
    cnt = lambda s: f"B{status_row[s]}"

    # ---- KPIs ----
    _section(ws, kpi_start - 1, f"Key metrics ({label})" if config.REGION_SPLIT else "Key metrics")
    total = f"B{kpi_start}"
    if region:   # AVERAGEIFS = average of a column, only for rows matching a condition
        avg_days = f'=IFERROR(AVERAGEIFS({_rng("Days to Response")},{_rng("Region")},"{region}"),"n/a")'
    else:
        avg_days = f'=IFERROR(AVERAGE({_rng("Days to Response")}),"n/a")'
    kpis = [
        ("Total applications", f'=COUNTIFS({_rng("Organisation")},"<>"{rx})', "0"),
        ("Active (Applied / Review / Assessment / Interview)",
         "=" + "+".join(cnt(s) for s in ("Applied", "Under Review", "Assessment", "Interview")), "0"),
        ("Response rate (assessment, interview, offer or rejection)",
         f"=IFERROR(({'+'.join(cnt(s) for s in RESPONDED)})/{total},0)", "0.0%"),
        ("Interview rate (interview + offer)",
         f"=IFERROR(({cnt('Interview')}+{cnt('Offer')})/{total},0)", "0.0%"),
        ("Offer rate", f"=IFERROR({cnt('Offer')}/{total},0)", "0.0%"),
        ("Rejection rate", f"=IFERROR({cnt('Rejected')}/{total},0)", "0.0%"),
        ("Avg. days from applying to a response", avg_days, "0.0"),
        ("Applied in last 7 days",
         f'=COUNTIFS({_rng("Date Applied")},">="&(TODAY()-7){rx})', "0"),
        ("Applied in last 30 days",
         f'=COUNTIFS({_rng("Date Applied")},">="&(TODAY()-30){rx})', "0"),
    ]
    for i, (text, formula, fmt) in enumerate(kpis):
        r = kpi_start + i
        ws.cell(r, 1, text).font = Font(name=F, size=10)
        c = ws.cell(r, 2, formula)
        c.font, c.number_format = Font(name=F, bold=True, size=10), fmt
        c.alignment = Alignment(horizontal="right")

    # ---- Status table (doubles as the colour legend) ----
    _section(ws, status_hdr - 1, "Applications by status (colours match the Applications sheet)")
    _table_header(ws, status_hdr, ["Status", "Count", "% of total"])
    for s, r in status_row.items():
        strong = config.STATUS_COLORS[s][0]
        a = ws.cell(r, 1, s)
        a.fill = PatternFill("solid", start_color=strong, end_color=strong)
        a.font = Font(name=F, bold=True, size=10)
        ws.cell(r, 2, f'=COUNTIFS({_rng("Status")},A{r}{rx})').font = Font(name=F, size=10)
        p = ws.cell(r, 3, f"=IFERROR(B{r}/{total},0)")
        p.number_format, p.font = "0.0%", Font(name=F, size=10)
    status_last = status_first + len(config.STATUSES) - 1

    chart = BarChart()
    chart.type = "col"
    chart.add_data(Reference(ws, min_col=2, min_row=status_hdr, max_row=status_last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=status_first, max_row=status_last))
    for i, s in enumerate(config.STATUSES):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = config.STATUS_COLORS[s][0]
        chart.series[0].dPt.append(pt)
    _style_chart(chart, "Applications by status", "G3")
    ws.add_chart(chart)
    row = status_last + 3

    # ---- India vs Outside India comparison (only on the "All" sheet) ----
    if region is None and config.REGION_SPLIT:
        _section(ws, row, "India vs outside India")
        _table_header(ws, row + 1, ["Region", "Applications", "Responses", "Response rate"])
        for i, reg in enumerate(config.REGIONS):
            r = row + 2 + i
            ws.cell(r, 1, reg)
            ws.cell(r, 2, f'=COUNTIFS({_rng("Organisation")},"<>",{_rng("Region")},A{r})')
            ws.cell(r, 3, "=" + "+".join(
                f'COUNTIFS({_rng("Region")},A{r},{_rng("Status")},"{s}")' for s in RESPONDED))
            ws.cell(r, 4, f"=IFERROR(C{r}/B{r},0)").number_format = "0.0%"
            for col in range(1, 5):
                ws.cell(r, col).font = Font(name=F, size=10)
        ws.cell(row + 5, 1, '"Unknown" jobs are in neither regional dashboard - set their Region '
                            'cell on the Applications sheet.').font = Font(name=F, italic=True, size=9, color="595959")
        row += 8

    # ---- By site ----
    if sites:
        _section(ws, row, "Which job site works best?")
        _table_header(ws, row + 1, ["Site", "Applications", "Responses", "Response rate"])
        first = row + 2
        for i, (site, _n) in enumerate(sites.most_common()):
            r = first + i
            ws.cell(r, 1, site)
            ws.cell(r, 2, f'=COUNTIFS({_rng("Site Used")},A{r}{rx})')
            ws.cell(r, 3, "=" + "+".join(
                f'COUNTIFS({_rng("Site Used")},A{r},{_rng("Status")},"{s}"{rx})' for s in RESPONDED))
            ws.cell(r, 4, f"=IFERROR(C{r}/B{r},0)").number_format = "0.0%"
            for col in range(1, 5):
                ws.cell(r, col).font = Font(name=F, size=10)
        last = first + len(sites) - 1
        ch = BarChart()
        ch.type = "bar"
        ch.add_data(Reference(ws, min_col=2, min_row=row + 1, max_row=last), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
        _style_chart(ch, "Applications by job site", "G19")
        ws.add_chart(ch)
        row = last + 3

    # ---- By month ----
    if months:
        _section(ws, row, "Applications per month")
        _table_header(ws, row + 1, ["Month", "Applications"])
        first = row + 2
        y, m = min(months).year, min(months).month
        end = max(max(months), date(today.year, today.month, 1))
        ordered = []
        while date(y, m, 1) <= end:
            ordered.append(date(y, m, 1))
            y, m = (y + 1, 1) if m == 12 else (y, m + 1)
        for i, start in enumerate(ordered):
            r = first + i
            ws.cell(r, 1, start).number_format = "mmm yyyy"
            ws.cell(r, 1).alignment = Alignment(horizontal="left")
            ws.cell(r, 2, f'=COUNTIFS({_rng("Date Applied")},">="&A{r},'
                          f'{_rng("Date Applied")},"<"&EDATE(A{r},1){rx})')
            for col in (1, 2):
                ws.cell(r, col).font = Font(name=F, size=10)
        last = first + len(ordered) - 1
        ch = BarChart()
        ch.type = "col"
        ch.add_data(Reference(ws, min_col=2, min_row=row + 1, max_row=last), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
        _style_chart(ch, "Applications per month", "G35")
        ws.add_chart(ch)
        row = last + 3

    # ---- Skills ----
    if skills:
        top = [s for s, _ in skills.most_common(10)]
        _section(ws, row, "Most requested skills (top 10)")
        _table_header(ws, row + 1, ["Skill", "Jobs mentioning it"])
        first = row + 2
        sk = _rng("Skills")
        for i, skill in enumerate(top):
            r = first + i
            ws.cell(r, 1, skill)
            # Skills are stored as "Python, SQL, Java". Matching "Java" with *Java* would also
            # hit "JavaScript", so match the four positions a whole item can occupy.
            ws.cell(r, 2, f'=COUNTIFS({sk},A{r}{rx})+COUNTIFS({sk},A{r}&", *"{rx})'
                          f'+COUNTIFS({sk},"*, "&A{r}&", *"{rx})+COUNTIFS({sk},"*, "&A{r}{rx})')
            for col in (1, 2):
                ws.cell(r, col).font = Font(name=F, size=10)
        last = first + len(top) - 1
        ch = BarChart()
        ch.type = "bar"
        ch.add_data(Reference(ws, min_col=2, min_row=row + 1, max_row=last), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
        _style_chart(ch, "Most requested skills", "G51")
        ws.add_chart(ch)


def build_stats_sheet(wb, today: date | None = None) -> None:
    """Rebuild every dashboard from scratch (they are cheap and fully formula-driven)."""
    today = today or date.today()
    for name in config.ALL_STATS_NAMES:
        if name in wb.sheetnames:
            del wb[name]
    for sheet_name, region in config.STATS_SHEETS:
        _build_one(wb, sheet_name, region, today)
