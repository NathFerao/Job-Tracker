"""
main.py - the entry point. Wires the stages together:

    fetch emails -> parse -> update Excel -> flag stale jobs -> rebuild stats -> save

Usage:
    python main.py --demo                 # try it on fake emails, no Gmail needed
    python main.py                        # real run against your Gmail
    python main.py --reprocess            # fill blank Region/Location on old rows (safe)
    python main.py --days 180             # look further back
    python main.py --demo --as-of 2026-11-10   # pretend today is this date (for testing)
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import date

import config
from excel_store import JobTracker
from email_parser import parse_email
from stats import build_stats_sheet

log = logging.getLogger("job_tracker")


def load_state() -> set[str]:
    """IDs of emails already processed, so re-running never creates duplicates."""
    if config.STATE_PATH.exists():
        return set(json.loads(config.STATE_PATH.read_text()).get("processed", []))
    return set()


def save_state(ids: set[str]) -> None:
    config.STATE_PATH.write_text(json.dumps({"processed": sorted(ids)}, indent=2))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Track job applications from your email in Excel.")
    ap.add_argument("--demo", action="store_true", help="use built-in sample emails instead of Gmail")
    ap.add_argument("--days", type=int, default=config.LOOKBACK_DAYS, help="how many days back to search")
    ap.add_argument("--excel", help="path of the Excel file to create/update")
    ap.add_argument("--reprocess", action="store_true",
                    help="re-read already-processed emails ONLY to fill blank cells (e.g. region, location)")
    ap.add_argument("--as-of", help="treat this YYYY-MM-DD as 'today' (handy for demos/tests)")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(message)s")
    today = date.fromisoformat(args.as_of) if args.as_of else date.today()
    excel_path = args.excel or (config.DEMO_EXCEL_PATH if args.demo else config.EXCEL_PATH)

    # 1. Fetch -------------------------------------------------------------
    previously_processed = set() if args.demo else load_state()
    processed = set() if (args.demo or args.reprocess) else set(previously_processed)
    if args.demo:
        from sample_data import DEMO_EMAILS
        emails = [e for e in DEMO_EMAILS]
    else:
        from gmail_client import fetch_emails
        emails = fetch_emails(args.days, skip_ids=processed)

    # Oldest first: "Applied" must be processed before the "Interview" that follows it.
    emails.sort(key=lambda e: e.date)

    # 2. Parse + 3. Update Excel -------------------------------------------
    tracker = JobTracker(excel_path)
    counts = {"relevant": 0, "ignored": 0}
    for mail in emails:
        processed.add(mail.id)
        event = parse_email(mail)
        if event is None:
            counts["ignored"] += 1
            log.info("  - ignored   : %s", mail.subject[:70])
            continue
        counts["relevant"] += 1
        outcome, row = tracker.apply_event(event, fill_only=args.reprocess)
        log.info("  + %-12s %-22s %-30s -> %s%s", event.status, (event.organisation or "?")[:22],
                 (event.role or "-")[:30], outcome, f" (row {row})" if row else "")

    # 4. Housekeeping ------------------------------------------------------
    regions_fixed = tracker.refresh_regions()
    stale = tracker.mark_stale(today) if not args.reprocess else 0
    build_stats_sheet(tracker.wb, today)

    try:
        tracker.save()
    except PermissionError:
        log.error("\nCould not save %s - is it open in Excel? Close it and run again.", excel_path)
        return 1
    if not args.demo:
        save_state(processed | previously_processed)  # only AFTER a successful save, so a crash never loses emails

    log.info("\nDone. %d job emails processed, %d ignored, %d marked 'No Response'.",
             counts["relevant"], counts["ignored"], stale)
    if config.REGION_SPLIT:
        log.info("%d previously-unknown regions resolved.", regions_fixed)
    log.info("Excel file: %s", excel_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
