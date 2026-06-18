import contextlib
import io
import json
import re
import tempfile
import unittest
from pathlib import Path

from scripts import scan_sensitive_fixtures


class SensitiveFixtureScanTest(unittest.TestCase):
    def test_current_committed_fixture_scan_has_no_blocking_findings(self):
        findings = scan_sensitive_fixtures.scan_tracked_files(Path.cwd())
        blocking = [finding for finding in findings if finding.severity == "block"]

        self.assertEqual(blocking, [])

    def test_cli_reports_path_and_category_without_sensitive_values(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            export_path = root / "private.quote-company-profile.json"
            export_path.write_text(
                json.dumps({
                    "company": "Do Not Echo Pte Ltd",
                    "bank": "Do Not Echo Bank",
                    "account_number": "1234567890",
                }),
                encoding="utf-8",
            )

            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                exit_code = scan_sensitive_fixtures.main([
                    "--root",
                    str(root),
                    "--path",
                    str(export_path),
                ])

        text = output.getvalue()
        self.assertEqual(exit_code, 1)
        self.assertIn("private-profile-export", text)
        self.assertIn("bank-payment-marker", text)
        self.assertIn("private.quote-company-profile.json", text)
        self.assertNotIn("Do Not Echo", text)
        self.assertNotIn("1234567890", text)

    def test_unreviewed_company_or_customer_markers_fail_without_echoing_values(self):
        original_company_re = scan_sensitive_fixtures.REAL_COMPANY_RE
        original_customer_re = scan_sensitive_fixtures.CUSTOMER_SAMPLE_RE
        scan_sensitive_fixtures.REAL_COMPANY_RE = re.compile(r"synthetic\s+sensitive\s+company", re.IGNORECASE)
        scan_sensitive_fixtures.CUSTOMER_SAMPLE_RE = re.compile(r"synthetic\s+sensitive\s+customer", re.IGNORECASE)

        try:
            with tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                fixture_path = root / "profiles" / "new" / "profile.json"
                fixture_path.parent.mkdir(parents=True)
                fixture_path.write_text(
                    json.dumps({
                        "company": "Synthetic Sensitive Company",
                        "customer": "Synthetic Sensitive Customer",
                    }),
                    encoding="utf-8",
                )

                output = io.StringIO()
                with contextlib.redirect_stdout(output):
                    exit_code = scan_sensitive_fixtures.main([
                        "--root",
                        str(root),
                        "--path",
                        str(fixture_path),
                    ])
        finally:
            scan_sensitive_fixtures.REAL_COMPANY_RE = original_company_re
            scan_sensitive_fixtures.CUSTOMER_SAMPLE_RE = original_customer_re

        text = output.getvalue()
        self.assertEqual(exit_code, 1)
        self.assertIn("real-company-identity-marker", text)
        self.assertIn("customer-sample-marker", text)
        self.assertIn("profiles/new/profile.json", text)
        self.assertNotIn("Synthetic Sensitive Company", text)
        self.assertNotIn("Synthetic Sensitive Customer", text)


if __name__ == "__main__":
    unittest.main()
