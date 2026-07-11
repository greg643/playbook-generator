import sys
import unittest
from pathlib import Path

from botocore.exceptions import ClientError


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "pipeline"))

from process_job import JobCancelled, ensure_job_active, scrub_job_payload  # noqa: E402


class FakeS3:
    def __init__(self, keys=()):
        self.keys = set(keys)

    def head_object(self, *, Bucket, Key):
        if Key in self.keys:
            return {}
        raise ClientError(
            {"Error": {"Code": "404", "Message": "Not Found"}},
            "HeadObject",
        )

    def list_objects_v2(self, *, Bucket, Prefix, MaxKeys):
        return {
            "Contents": [
                {"Key": key}
                for key in sorted(self.keys)
                if key.startswith(Prefix)
            ][:MaxKeys]
        }

    def delete_objects(self, *, Bucket, Delete):
        for item in Delete["Objects"]:
            self.keys.discard(item["Key"])
        return {}


class JobCancellationTests(unittest.TestCase):
    def test_cancellation_marker_stops_processing(self):
        client = FakeS3({"jobs/job/cancelled.json"})
        with self.assertRaises(JobCancelled):
            ensure_job_active(client, "bucket", "job")

    def test_scrub_retains_only_non_content_metadata(self):
        prefix = "jobs/job/"
        client = FakeS3(
            {
                prefix + "owner.json",
                prefix + "status.json",
                prefix + "cancelled.json",
                prefix + "input.pptx",
                prefix + "plays/01.png",
                prefix + "offense_coach_card.pdf",
            }
        )
        scrub_job_payload(client, "bucket", "job")
        self.assertEqual(
            client.keys,
            {
                prefix + "owner.json",
                prefix + "status.json",
                prefix + "cancelled.json",
            },
        )


if __name__ == "__main__":
    unittest.main()
