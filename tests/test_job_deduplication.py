import unittest
import inspect
from unittest.mock import patch

import pandas as pd

import java_contract_job_agent as agent
from java_contract_job_agent import _deduplicate_jobs, is_c2c, rejects_c2c
from linkedin_posts_scraper import scrape_posts


class DeduplicateJobsTests(unittest.TestCase):
    def test_posts_scrape_window_is_two_hours(self):
        self.assertEqual(agent.POSTS_MAX_AGE_MINUTES, 120)
        self.assertEqual(inspect.signature(scrape_posts).parameters["max_age_minutes"].default, 120)

    def test_explicit_c2c_denials_override_positive_mentions(self):
        denials = [
            "Will not accept on C2C",
            "Not accepting C2C candidates",
            "C2C not accepted",
            "C2C and Sponsorship: Not Available",
        ]

        for wording in denials:
            with self.subTest(wording=wording):
                self.assertTrue(is_c2c(wording, "Java Developer"))
                self.assertTrue(rejects_c2c(wording, "Java Developer"))

    def test_sponsorship_denial_alone_does_not_reject_c2c(self):
        description = "C2C is available; visa sponsorship is not available."

        self.assertTrue(is_c2c(description, "Java Developer"))
        self.assertFalse(rejects_c2c(description, "Java Developer"))

    def test_keeps_distinct_posts_from_same_author(self):
        jobs = pd.DataFrame(
            {
                "site": ["linkedin_posts", "linkedin_posts"],
                "job_url": ["https://linkedin.com/in/recruiter"] * 2,
                "description": ["Java contract role, Austin", "QA contract role, Boston"],
            }
        )

        self.assertEqual(len(_deduplicate_jobs(jobs)), 2)

    def test_deduplicates_same_post_with_normalized_whitespace(self):
        jobs = pd.DataFrame(
            {
                "site": ["linkedin_posts", "linkedin_posts"],
                "job_url": ["https://linkedin.com/in/recruiter"] * 2,
                "description": ["Java contract role\nAustin", " java   contract role Austin "],
            }
        )

        self.assertEqual(len(_deduplicate_jobs(jobs)), 1)

    def test_keeps_job_board_url_deduplication(self):
        jobs = pd.DataFrame(
            {
                "site": ["linkedin", "linkedin"],
                "job_url": ["https://jobs.example/123"] * 2,
                "description": ["First listing text", "Duplicate listing text"],
            }
        )

        self.assertEqual(len(_deduplicate_jobs(jobs)), 1)

    @patch("java_contract_job_agent.smtplib.SMTP_SSL")
    def test_empty_category_sends_no_match_summary(self, smtp_ssl):
        smtp = smtp_ssl.return_value.__enter__.return_value
        smtp.send_message.return_value = {}
        with patch.object(agent, "GMAIL_USER", "sender@example.test"), patch.object(
            agent, "GMAIL_APP_PASSWORD", "test-password"
        ):
            agent.send_email_notification(
                pd.DataFrame(columns=["vendor_emails"]), "", "one@example.test", ".NET"
            )

        message = smtp.send_message.call_args.args[0]
        plain_body = message.get_body(preferencelist=("plain",)).get_content()
        self.assertIn("No contract/C2C .NET postings matched this run.", plain_body)

    def test_standalone_1099_is_not_explicit_c2c(self):
        self.assertFalse(agent.is_c2c("Rate: $60/hour, 1099", "Java Developer"))

    def test_explicit_c2c_with_1099_is_accepted(self):
        self.assertTrue(agent.is_c2c("Work Type: C2C / 1099", "Java Developer"))


if __name__ == "__main__":
    unittest.main()