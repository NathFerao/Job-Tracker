"""
test_tracker.py - run with:   python -m unittest -v

Tests document the behaviour you care about AND protect it when you change
the rules later (e.g. add a new phrase and accidentally break another case).
"""
import tempfile
import unittest

import config
from datetime import date, datetime
from pathlib import Path

from email_parser import classify_status, clean_company, detect_region, extract_skills, parse_email
from excel_store import JobTracker, orgs_match, roles_similar
from models import RawEmail
from sample_data import DEMO_EMAILS


def mail(subject, body="", name="Acme Careers", addr="careers@acme.com", when="2026-10-01T10:00:00"):
    return RawEmail("t1", datetime.fromisoformat(when), name, addr, subject, body)


class StatusClassification(unittest.TestCase):
    def test_basic_statuses(self):
        self.assertEqual(classify_status("Thank you for applying to Acme", ""), "Applied")
        self.assertEqual(classify_status("Update", "Unfortunately, we will not be moving forward."), "Rejected")
        self.assertEqual(classify_status("Interview Invitation", ""), "Interview")
        self.assertEqual(classify_status("Update", "We are pleased to offer you the role."), "Offer")

    def test_rejection_beats_offer_wording(self):
        body = "Unfortunately we will not be extending an offer to you."
        self.assertEqual(classify_status("Your update", body), "Rejected")

    def test_conditional_sentence_is_not_an_interview(self):
        body = "We received your application. If selected, we will schedule an interview."
        self.assertEqual(classify_status("Thank you for applying", body), "Applied")

    def test_subject_wins_over_boilerplate_body(self):
        body = "Unfortunately we cannot reply to every applicant."
        self.assertEqual(classify_status("Application received", body), "Applied")

    def test_unrelated_email(self):
        self.assertIsNone(classify_status("Your Amazon order has shipped", "Track your parcel"))

    def test_job_alert_is_ignored(self):
        self.assertIsNone(parse_email(mail("Your job alert: Python Developer in Goa", "Apply now")))


class Extraction(unittest.TestCase):
    def test_company_and_role_from_subject(self):
        ev = parse_email(mail("Your application to Data Analyst at Zeta Analytics",
                              addr="jobs-noreply@linkedin.com", name="LinkedIn"))
        self.assertEqual((ev.organisation, ev.role, ev.site), ("Zeta Analytics", "Data Analyst", "LinkedIn"))

    def test_legal_suffix_removed(self):
        self.assertEqual(clean_company("Nimbus Software Pvt Ltd"), "Nimbus Software")

    def test_java_does_not_match_javascript(self):
        self.assertEqual(extract_skills("We use JavaScript and React"), "JavaScript, React")
        self.assertEqual(extract_skills("Java, C++ and C#"), "Java, C++, C#")

    def test_job_board_sender_is_not_used_as_company(self):
        ev = parse_email(mail("Application sent", "Your application has been sent.",
                              name="Naukri", addr="donotreply@naukri.com"))
        self.assertIsNone(ev.organisation)


class Matching(unittest.TestCase):
    def test_org_matching(self):
        self.assertTrue(orgs_match("Nimbus Software Pvt. Ltd.", "nimbus software"))
        self.assertFalse(orgs_match("Nimbus Software", "Orbit Robotics"))

    def test_role_matching(self):
        self.assertTrue(roles_similar("Software Engineer", "Software Engineer – Platform"))
        self.assertFalse(roles_similar("Data Analyst", "Backend Engineer"))


class EndToEnd(unittest.TestCase):
    def run_demo(self, path):
        tracker = JobTracker(path)
        for m in sorted(DEMO_EMAILS, key=lambda e: e.date):
            ev = parse_email(m)
            if ev:
                tracker.apply_event(ev)
        tracker.save()
        return JobTracker(path)

    def test_demo_outcome(self):
        with tempfile.TemporaryDirectory() as d:
            t = self.run_demo(Path(d) / "x.xlsx")
            rows = {t.ws.cell(r, 3).value: t.ws.cell(r, 8).value for r in range(2, t.last_row() + 1)}
            self.assertEqual(rows, {
                "Zeta Analytics": "Rejected", "Nimbus Software": "Interview",
                "Coastal Labs": "Applied", "Orbit Robotics": "Interview",
                "Finlytics": "Offer", "Quantum Retail": "Under Review",
                "Northwind Cloud": "Interview", "Helvetia Systems": "Applied",
            })  # the duplicate Nimbus email and the job alert created nothing

    def test_status_never_moves_backwards(self):
        with tempfile.TemporaryDirectory() as d:
            t = self.run_demo(Path(d) / "x.xlsx")
            late_applied = parse_email(mail("Application received - Finlytics Data Engineer",
                                            "Thank you for applying for the Data Engineer role at Finlytics."))
            t.apply_event(late_applied)
            row = t.find_match("Finlytics", "Data Engineer", strict=True)
            self.assertEqual(t.ws.cell(row, 8).value, "Offer")

    def test_stale_jobs_are_flagged(self):
        with tempfile.TemporaryDirectory() as d:
            t = self.run_demo(Path(d) / "x.xlsx")
            self.assertEqual(t.mark_stale(date(2026, 11, 10)), 1)  # only Coastal Labs


@unittest.skipUnless(config.REGION_SPLIT, "REGION_SPLIT is off")
class Regions(unittest.TestCase):
    def test_detect_region(self):
        self.assertEqual(detect_region("Bengaluru, Karnataka, India", "LinkedIn", ""), "India")
        self.assertEqual(detect_region("Pune, India (Hybrid)", "LinkedIn", ""), "India")
        self.assertEqual(detect_region("Dublin, Ireland", "LinkedIn", ""), "Outside India")
        self.assertEqual(detect_region("Remote, US", "LinkedIn", ""), "Outside India")

    def test_remote_needs_a_hint(self):
        self.assertEqual(detect_region("Remote", "LinkedIn", "jobs@linkedin.com"), "Unknown")
        self.assertEqual(detect_region("Remote", "Naukri", "x@naukri.com"), "India")   # India-focused site
        self.assertEqual(detect_region("", "Company Website", "hr@acme.co.in"), "India")  # .in domain
        self.assertEqual(detect_region("Remote", "Indeed", "x@indeed.com", "Salary 12 LPA"), "India")

    def test_regions_in_demo(self):
        with tempfile.TemporaryDirectory() as d:
            t = EndToEnd().run_demo(Path(d) / "x.xlsx")
            reg = {t.ws.cell(r, 3).value: t.ws.cell(r, 14).value for r in range(2, t.last_row() + 1)}
            self.assertEqual(reg["Zeta Analytics"], "India")
            self.assertEqual(reg["Northwind Cloud"], "Outside India")
            self.assertEqual(reg["Helvetia Systems"], "Outside India")
            self.assertEqual(reg["Coastal Labs"], "Unknown")   # remote on Indeed: can't tell

    def test_manual_region_is_never_overwritten(self):
        with tempfile.TemporaryDirectory() as d:
            t = EndToEnd().run_demo(Path(d) / "x.xlsx")
            row = t.find_match("Coastal Labs", "Junior Backend Engineer", strict=True)
            t.ws.cell(row, 14, "Outside India")           # the user decides by hand
            ev = parse_email(mail("Interview Invitation – Junior Backend Engineer at Coastal Labs",
                                  "Please schedule an interview. Location: Mumbai, India"))
            t.apply_event(ev)
            self.assertEqual(t.ws.cell(row, 14).value, "Outside India")

    def test_reprocess_only_fills_blanks(self):
        with tempfile.TemporaryDirectory() as d:
            t = EndToEnd().run_demo(Path(d) / "x.xlsx")
            row = t.find_match("Northwind Cloud", "Backend Engineer", strict=True)
            status, notes, n_rows = t.ws.cell(row, 8).value, t.ws.cell(row, 12).value, t.last_row()
            t.ws.cell(row, 4, None)                      # blank out location and region
            t.ws.cell(row, 14, "Unknown")
            for m in sorted(DEMO_EMAILS, key=lambda e: e.date):
                ev = parse_email(m)
                if ev:
                    t.apply_event(ev, fill_only=True)
            self.assertEqual(t.ws.cell(row, 4).value, "Dublin, Ireland")
            self.assertEqual(t.ws.cell(row, 14).value, "Outside India")
            self.assertEqual((t.ws.cell(row, 8).value, t.ws.cell(row, 12).value, t.last_row()),
                             (status, notes, n_rows))   # status, notes, row count unchanged

    def test_refresh_regions_uses_hand_typed_location(self):
        with tempfile.TemporaryDirectory() as d:
            t = EndToEnd().run_demo(Path(d) / "x.xlsx")
            row = t.find_match("Coastal Labs", "Junior Backend Engineer", strict=True)
            t.ws.cell(row, 4, "Panaji, Goa")             # user types the location
            self.assertEqual(t.refresh_regions(), 1)
            self.assertEqual(t.ws.cell(row, 14).value, "India")

    def test_old_workbook_is_upgraded(self):
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "x.xlsx"
            EndToEnd().run_demo(path)
            wb = load_workbook(path)                       # simulate a pre-Region file
            wb["Applications"].delete_cols(14)
            wb.save(path)
            t = JobTracker(path)
            self.assertEqual(t.ws.cell(1, 14).value, "Region")
            self.assertEqual(t.ws.cell(2, 14).value, "India")


@unittest.skipIf(config.REGION_SPLIT, "REGION_SPLIT is on")
class RegionSplitOff(unittest.TestCase):
    def test_region_column_and_extra_sheets_are_removed(self):
        from openpyxl import load_workbook
        from stats import build_stats_sheet
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "x.xlsx"
            EndToEnd().run_demo(path)
            wb = load_workbook(path)                         # simulate a file made with the split on
            wb["Applications"].cell(1, 14, "Region")
            wb["Applications"].cell(2, 14, "India")
            for n in ("Stats - All", "Stats - India"):
                wb.create_sheet(n)
            wb.save(path)
            t = JobTracker(path)
            self.assertIsNone(t.ws.cell(1, 14).value)        # column gone
            build_stats_sheet(t.wb)
            self.assertEqual(t.wb.sheetnames, ["Applications", "Stats"])


if __name__ == "__main__":
    unittest.main()
