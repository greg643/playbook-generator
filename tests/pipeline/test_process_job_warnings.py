import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "pipeline"))

from process_job import (  # noqa: E402
    build_complete_status,
    visible_pipeline_warnings,
)


class ProcessJobWarningTests(unittest.TestCase):
    def test_known_warning_is_exposed_when_offense_output_is_requested(self):
        result = {
            "warnings": [
                {
                    "code": "assumed_offense_before_defense",
                    "playCount": 2,
                    "ignored": "not part of the public contract",
                },
                {
                    "code": "skipped_no_field_rectangle",
                    "slideCount": 3,
                },
            ]
        }

        warnings = visible_pipeline_warnings(result, include_offense=True)

        self.assertEqual(warnings, [
            {
                "code": "assumed_offense_before_defense",
                "playCount": 2,
            },
            {
                "code": "skipped_no_field_rectangle",
                "slideCount": 3,
            },
        ])
        self.assertEqual(
            build_complete_status(["offense_coach_card.pdf"], warnings),
            {
                "status": "complete",
                "files": ["offense_coach_card.pdf"],
                "warnings": warnings,
            },
        )

    def test_only_offense_assumption_is_suppressed_for_defense_only_output(self):
        result = {
            "warnings": [
                {
                    "code": "assumed_offense_before_defense",
                    "playCount": 2,
                },
                {
                    "code": "skipped_before_first_divider",
                    "slideCount": 1,
                },
            ]
        }

        warnings = visible_pipeline_warnings(result, include_offense=False)
        self.assertEqual(
            warnings,
            [{
                "code": "skipped_before_first_divider",
                "slideCount": 1,
            }],
        )
        self.assertEqual(
            build_complete_status(["defense_coach_card.pdf"], warnings),
            {
                "status": "complete",
                "files": ["defense_coach_card.pdf"],
                "warnings": warnings,
            },
        )

    def test_unknown_and_malformed_warnings_are_dropped(self):
        malformed = [
            {"code": "future_warning", "playCount": 2},
            {"code": "assumed_offense_before_defense", "playCount": True},
            {"code": "assumed_offense_before_defense", "playCount": 0},
            {"code": "assumed_offense_before_defense", "playCount": 65},
            {"code": "skipped_before_first_divider", "slideCount": True},
            {"code": "skipped_no_field_rectangle", "slideCount": 0},
            {"code": "skipped_no_field_rectangle", "slideCount": 101},
            "not-an-object",
        ]

        self.assertEqual(
            visible_pipeline_warnings(
                {"warnings": malformed},
                include_offense=True,
            ),
            [],
        )
        self.assertEqual(
            visible_pipeline_warnings(None, include_offense=True),
            [],
        )


if __name__ == "__main__":
    unittest.main()
