#!/usr/bin/env python3
"""Best-effort GitHub Actions finalizer for failures before process_job runs."""

import json
import os
import sys
import uuid

import boto3
from botocore.exceptions import ClientError

from process_job import scrub_job_payload


TERMINAL_STATUSES = {"complete", "error"}


def main():
    if len(sys.argv) != 2:
        raise SystemExit("Usage: mark_job_failed.py <job_id>")
    job_id = sys.argv[1]
    try:
        parsed = uuid.UUID(job_id)
    except ValueError as exc:
        raise SystemExit("Invalid job ID") from exc
    if str(parsed) != job_id.lower():
        raise SystemExit("Invalid job ID")

    bucket = os.environ["R2_BUCKET"]
    key = f"jobs/{job_id}/status.json"
    client = boto3.client(
        "s3",
        endpoint_url=os.environ["R2_ENDPOINT"],
        aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
        aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"],
        region_name="auto",
    )

    # process_job already attempts this; retry in the always() finalizer so a
    # transient delete failure does not leave cancelled uploads/PDFs behind.
    try:
        client.head_object(Bucket=bucket, Key=f"jobs/{job_id}/cancelled.json")
    except ClientError as exc:
        code = str(exc.response.get("Error", {}).get("Code", ""))
        if code not in {"404", "NoSuchKey", "NotFound"}:
            raise
    else:
        scrub_job_payload(client, bucket, job_id)

    try:
        current = json.loads(client.get_object(Bucket=bucket, Key=key)["Body"].read())
    except Exception:
        current = {}
    if current.get("status") in TERMINAL_STATUSES:
        return

    preserved = {
        name: current[name]
        for name in ("mode", "options", "createdAt", "ownerId")
        if name in current
    }
    body = {
        **preserved,
        "status": "error",
        "message": "The processing worker stopped before the job completed. Please try again.",
    }
    client.put_object(
        Bucket=bucket,
        Key=key,
        Body=json.dumps(body),
        ContentType="application/json",
    )


if __name__ == "__main__":
    main()
