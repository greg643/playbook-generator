"""Fail-closed deployment check for the renderer's isolated R2 credential.

Only a uniquely named synthetic object is written/deleted, in the job bucket.
Protected buckets are probed read-only; their objects are never printed.
"""
import json
import os
import uuid

import boto3
from botocore.exceptions import ClientError

JOB_BUCKET = "gss-playbook-jobs"
PROTECTED_BUCKETS = ("playbook-files", "gss-playbook-auth-state")


def verify(s3, bucket):
    if bucket != JOB_BUCKET:
        raise ValueError("Renderer must use the isolated job bucket")
    for protected in PROTECTED_BUCKETS:
        try:
            s3.list_objects_v2(Bucket=protected, MaxKeys=1)
        except ClientError as error:
            if error.response.get("Error", {}).get("Code") != "AccessDenied":
                raise RuntimeError("Protected bucket denial was not proven") from None
        else:
            raise RuntimeError("Renderer can access a protected bucket")
    key = f"qualification/{uuid.uuid4()}/storage-probe.txt"
    payload = b"GSS isolated renderer storage qualification\n"
    created = False
    try:
        s3.put_object(Bucket=bucket, Key=key, Body=payload, ContentType="text/plain")
        created = True
        obj = s3.get_object(Bucket=bucket, Key=key)
        with obj["Body"] as body:
            if body.read(len(payload) + 1) != payload:
                raise RuntimeError("Job object readback mismatch")
        page = s3.list_objects_v2(Bucket=bucket, Prefix=key, MaxKeys=2)
        if not any(item["Key"] == key for item in page.get("Contents", [])):
            raise RuntimeError("Job object listing mismatch")
    finally:
        if created:
            s3.delete_object(Bucket=bucket, Key=key)
    return {"job_write_read_list_delete": "PASS", "protected_bucket_denials": 2}


if __name__ == "__main__":
    client = boto3.client("s3", endpoint_url=os.environ["R2_ENDPOINT"],
                          aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
                          aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"],
                          region_name="auto")
    print(json.dumps(verify(client, os.environ["R2_BUCKET"])))
