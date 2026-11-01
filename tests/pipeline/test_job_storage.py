import io
import unittest
from unittest.mock import Mock
from botocore.exceptions import ClientError
from pipeline.verify_job_storage import JOB_BUCKET, verify


class JobStorageTests(unittest.TestCase):
    def client(self, denial="AccessDenied"):
        s3 = Mock()
        saved = {}
        def listing(**kw):
            if kw["Bucket"] != JOB_BUCKET:
                if denial is None:
                    return {}
                raise ClientError({"Error": {"Code": denial}}, "ListObjectsV2")
            return {"Contents": [{"Key": saved["Key"]}]}
        def put(**kw):
            saved.update(kw)
        s3.list_objects_v2.side_effect = listing
        s3.put_object.side_effect = put
        s3.get_object.side_effect = lambda **kw: {"Body": io.BytesIO(saved["Body"])}
        return s3

    def test_isolation_and_round_trip(self):
        s3 = self.client()
        self.assertEqual(verify(s3, JOB_BUCKET)["protected_bucket_denials"], 2)
        s3.delete_object.assert_called_once()
        self.assertEqual(s3.delete_object.call_args.kwargs["Bucket"], JOB_BUCKET)

    def test_broad_credential_fails_before_any_write(self):
        s3 = self.client(None)
        with self.assertRaises(RuntimeError): verify(s3, JOB_BUCKET)
        s3.put_object.assert_not_called()

    def test_wrong_bucket_fails_before_io(self):
        s3 = self.client()
        with self.assertRaises(ValueError): verify(s3, "playbook-files")
        self.assertEqual(s3.mock_calls, [])

    def test_authentication_or_transport_failure_is_not_scope_proof(self):
        for code in ("InvalidAccessKeyId", "NoSuchBucket", "SlowDown"):
            s3 = self.client(code)
            with self.assertRaises(RuntimeError): verify(s3, JOB_BUCKET)
            s3.put_object.assert_not_called()

    def test_failed_readback_still_cleans_only_its_probe(self):
        s3 = self.client()
        s3.get_object.side_effect = RuntimeError("read failed")
        with self.assertRaises(RuntimeError): verify(s3, JOB_BUCKET)
        s3.delete_object.assert_called_once()
