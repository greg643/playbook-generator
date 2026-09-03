#!/usr/bin/env python3
"""
GitHub Actions job orchestrator.

Downloads job input from R2 (an uploaded PPTX, or play PNGs exported by the
web editor in images mode), runs the playbook pipeline, uploads PDFs back to R2.

Usage:
    python pipeline/process_job.py <job_id>

Environment variables:
    R2_ENDPOINT, R2_ACCESS_KEY_ID, R2_SECRET_ACCESS_KEY, R2_BUCKET
"""

import os
import sys
import json
import time
import tempfile
import traceback
import uuid
import re
from pathlib import Path

import boto3
from botocore.exceptions import ClientError


class JobCancelled(RuntimeError):
    """Raised when account deletion has cancelled an in-flight job."""


PLAY_IMAGE_NAME_RE = re.compile(r"(?:0[1-9]|1[0-6]|D[1-6])\.png")
MAX_PLAY_IMAGES = 22
MAX_PLAY_IMAGE_BYTES = 4 * 1024 * 1024
ASSUMED_OFFENSE_WARNING_CODE = "assumed_offense_before_defense"
SKIPPED_BEFORE_DIVIDER_WARNING_CODE = "skipped_before_first_divider"
SKIPPED_NO_FIELD_WARNING_CODE = "skipped_no_field_rectangle"
MAX_ASSUMED_OFFENSE_WARNING_PLAYS = 64
MAX_WARNING_SLIDES = 100


def job_is_cancelled(s3, bucket, job_id):
    try:
        s3.head_object(Bucket=bucket, Key=f"jobs/{job_id}/cancelled.json")
        return True
    except ClientError as exc:
        code = str(exc.response.get("Error", {}).get("Code", ""))
        if code in {"404", "NoSuchKey", "NotFound"}:
            return False
        raise


def ensure_job_active(s3, bucket, job_id):
    if job_is_cancelled(s3, bucket, job_id):
        raise JobCancelled("Job was cancelled because its account was deleted")


def scrub_job_payload(s3, bucket, job_id):
    """Delete user content while retaining owner/status/cancellation metadata."""
    prefix = f"jobs/{job_id}/"
    preserved = {
        f"{prefix}owner.json",
        f"{prefix}status.json",
        f"{prefix}cancelled.json",
    }
    while True:
        listing = s3.list_objects_v2(Bucket=bucket, Prefix=prefix, MaxKeys=1000)
        keys = [
            {"Key": item["Key"]}
            for item in listing.get("Contents", [])
            if item["Key"] not in preserved
        ]
        if not keys:
            return
        s3.delete_objects(Bucket=bucket, Delete={"Objects": keys, "Quiet": True})


def get_r2_client():
    """Create an S3-compatible client for Cloudflare R2."""
    return boto3.client(
        "s3",
        endpoint_url=os.environ["R2_ENDPOINT"],
        aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
        aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"],
        region_name="auto",
    )


def update_status(s3, bucket, job_id, status_dict, request_fields=None):
    """Write status.json to R2, preserving the original job request fields
    (mode/options/createdAt) so a re-run of the job still knows what it is."""
    body = {**(request_fields or {}), **status_dict}
    s3.put_object(
        Bucket=bucket,
        Key=f"jobs/{job_id}/status.json",
        Body=json.dumps(body),
        ContentType="application/json",
    )


def visible_pipeline_warnings(pipeline_result, *, include_offense):
    """Return the small warning subset that is useful for this job's outputs."""
    if not isinstance(pipeline_result, dict):
        return []
    raw_warnings = pipeline_result.get("warnings")
    if not isinstance(raw_warnings, list):
        return []

    # Only fixed-code, bounded diagnostics are public. Ignore unknown or
    # malformed values rather than turning an otherwise successful job into an
    # error or reflecting arbitrary text through the status API.
    visible = []
    seen_codes = set()
    for warning in raw_warnings:
        if not isinstance(warning, dict):
            continue
        code = warning.get("code")
        if code in seen_codes:
            continue
        if code == ASSUMED_OFFENSE_WARNING_CODE:
            play_count = warning.get("playCount")
            if (
                include_offense
                and type(play_count) is int
                and 1 <= play_count <= MAX_ASSUMED_OFFENSE_WARNING_PLAYS
            ):
                visible.append({
                    "code": ASSUMED_OFFENSE_WARNING_CODE,
                    "playCount": play_count,
                })
                seen_codes.add(code)
        elif code in {
            SKIPPED_BEFORE_DIVIDER_WARNING_CODE,
            SKIPPED_NO_FIELD_WARNING_CODE,
        }:
            slide_count = warning.get("slideCount")
            if type(slide_count) is int and 1 <= slide_count <= MAX_WARNING_SLIDES:
                visible.append({"code": code, "slideCount": slide_count})
                seen_codes.add(code)
    return visible


def build_complete_status(files, warnings=None):
    """Build the backward-compatible terminal status returned to the UI."""
    status = {"status": "complete", "files": files}
    if warnings:
        status["warnings"] = warnings
    return status


def download_play_images(s3, bucket, job_id, plays_dir):
    """Download every object under jobs/<job_id>/plays/ from R2 (paginated)."""
    count = 0
    prefix = f"jobs/{job_id}/plays/"
    list_kwargs = {"Bucket": bucket, "Prefix": prefix}
    names = set()
    while True:
        listing = s3.list_objects_v2(**list_kwargs)
        for obj in listing.get("Contents", []):
            name = obj["Key"].rsplit("/", 1)[-1]
            if not name:
                continue
            if (
                obj["Key"] != prefix + name
                or PLAY_IMAGE_NAME_RE.fullmatch(name) is None
                or name in names
            ):
                raise ValueError("Job contains an unsupported play image name")
            if count >= MAX_PLAY_IMAGES:
                raise ValueError(f"Job contains too many play images (max {MAX_PLAY_IMAGES})")
            declared_size = obj.get("Size")
            if declared_size is not None and (
                declared_size <= 0 or declared_size > MAX_PLAY_IMAGE_BYTES
            ):
                raise ValueError("Job contains an invalid play image size")
            dest = plays_dir / name
            s3.download_file(bucket, obj["Key"], str(dest))
            if dest.stat().st_size <= 0 or dest.stat().st_size > MAX_PLAY_IMAGE_BYTES:
                raise ValueError("Job contains an invalid play image size")
            print(f"  Downloaded {name} ({dest.stat().st_size} bytes)")
            names.add(name)
            count += 1
        if listing.get("IsTruncated"):
            list_kwargs["ContinuationToken"] = listing["NextContinuationToken"]
        else:
            break
    return count


def main():
    if len(sys.argv) < 2:
        print("Usage: python pipeline/process_job.py <job_id>")
        sys.exit(1)

    job_id = sys.argv[1]
    try:
        parsed_job_id = uuid.UUID(job_id)
    except (ValueError, AttributeError) as exc:
        raise SystemExit("Invalid job ID") from exc
    if str(parsed_job_id) != job_id.lower():
        raise SystemExit("Invalid job ID")
    bucket = os.environ["R2_BUCKET"]
    s3 = get_r2_client()

    print(f"Processing job: {job_id}")

    request_fields = {}
    pipeline_warnings = []
    try:
        # Read the job request first: the status update below overwrites
        # status.json, and the editor flow stores mode + options there.
        # Retry transient failures rather than misrouting an images job
        # into the PPTX path on a blip.
        request_data = None
        for attempt in range(3):
            try:
                request_obj = s3.get_object(Bucket=bucket, Key=f"jobs/{job_id}/status.json")
                request_data = json.loads(request_obj["Body"].read())
                break
            except s3.exceptions.NoSuchKey:
                request_data = {}  # genuinely absent (legacy job): PPTX defaults
                break
            except Exception as e:
                print(f"  status.json pre-read attempt {attempt + 1} failed: {e}")
                time.sleep(2 * (attempt + 1))
        if request_data is None:
            raise RuntimeError("Could not read the job request (status.json) from R2")
        request_fields = {
            k: request_data[k]
            for k in ("mode", "options", "createdAt", "ownerId")
            if k in request_data
        }
        mode = request_data.get("mode", "pptx")
        if mode not in {"pptx", "images"}:
            raise ValueError("Unsupported job mode")
        ensure_job_active(s3, bucket, job_id)
        print(f"Job mode: {mode}")

        # Update status to processing
        update_status(s3, bucket, job_id, {"status": "processing", "step": "downloading"}, request_fields)

        # Create temporary working directory
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            output_dir = tmpdir / "output"
            output_dir.mkdir()

            if mode == "images":
                # Editor flow: play PNGs were uploaded to jobs/<job_id>/plays/ by /api/generate
                options = request_data.get("options", {})
                offense_coach_card = options.get("offense_coach_card", options.get("offense", True))
                offense_wristband = options.get("offense_wristband", options.get("offense", True))
                defense_coach_card = options.get("defense_coach_card", options.get("defense", True))
                defense_wristband = options.get("defense_wristband", options.get("defense", True))
                gen_offense = offense_coach_card or offense_wristband
                gen_defense = defense_coach_card or defense_wristband
                print(f"  Outputs: offense_coach_card={offense_coach_card}, offense_wristband={offense_wristband}, "
                      f"defense_coach_card={defense_coach_card}, defense_wristband={defense_wristband}")

                # Download play images from R2
                print("Downloading play images from R2...")
                plays_dir = tmpdir / "plays"
                plays_dir.mkdir()
                count = download_play_images(s3, bucket, job_id, plays_dir)
                if count == 0:
                    raise RuntimeError("No play images found for this job")
                ensure_job_active(s3, bucket, job_id)

                # Generate PDFs straight from the play images (no LibreOffice step)
                update_status(s3, bucket, job_id, {"status": "processing", "step": "generating"}, request_fields)

                pipeline_dir = Path(__file__).parent
                sys.path.insert(0, str(pipeline_dir))
                from playbook_pipeline import PlaybookGenerator

                generator = PlaybookGenerator(str(plays_dir), str(output_dir))
                generator.generate_all(
                    gen_offense=gen_offense,
                    gen_defense=gen_defense,
                    offense_coach_card=offense_coach_card,
                    offense_wristband=offense_wristband,
                    defense_coach_card=defense_coach_card,
                    defense_wristband=defense_wristband,
                    show_offense_title=options.get("show_offense_title") is True,
                    show_defense_title=options.get("show_defense_title") is not False,
                )
            else:
                pptx_path = tmpdir / "input.pptx"

                # Use the pre-read job request: the "processing" status write above
                # replaced status.json, so re-reading it here would lose the options
                # (that re-read bug silently ignored output selections until now).
                options = request_data.get("options", {})
                # Support both old-style (offense/defense booleans) and new-style (4 granular outputs)
                offense_coach_card = options.get("offense_coach_card", options.get("offense", True))
                offense_wristband = options.get("offense_wristband", options.get("offense", True))
                defense_coach_card = options.get("defense_coach_card", options.get("defense", True))
                defense_wristband = options.get("defense_wristband", options.get("defense", True))
                gen_offense = offense_coach_card or offense_wristband
                gen_defense = defense_coach_card or defense_wristband
                sections = "both" if (gen_offense and gen_defense) else ("offense" if gen_offense else "defense")
                print(f"  Outputs: offense_coach_card={offense_coach_card}, offense_wristband={offense_wristband}, "
                      f"defense_coach_card={defense_coach_card}, defense_wristband={defense_wristband}")

                # Download PPTX from R2
                print("Downloading PPTX from R2...")
                s3.download_file(bucket, f"jobs/{job_id}/input.pptx", str(pptx_path))
                print(f"  Downloaded {pptx_path.stat().st_size} bytes")
                ensure_job_active(s3, bucket, job_id)

                # Run the pipeline
                update_status(s3, bucket, job_id, {"status": "processing", "step": "generating"}, request_fields)

                # Import and run pipeline from the same package
                pipeline_dir = Path(__file__).parent
                sys.path.insert(0, str(pipeline_dir))
                from playbook_pipeline import main as pipeline_main

                # Build the list of selected outputs
                selected = []
                if offense_coach_card: selected.append("offense_coach_card")
                if offense_wristband: selected.append("offense_wristband")
                if defense_coach_card: selected.append("defense_coach_card")
                if defense_wristband: selected.append("defense_wristband")

                titles = []
                if options.get("show_offense_title") is True:
                    titles.append("offense")
                if options.get("show_defense_title") is not False:
                    titles.append("defense")

                # Override sys.argv for the pipeline
                original_argv = sys.argv
                sys.argv = ["playbook_pipeline.py", str(pptx_path), str(output_dir),
                             "--sections", sections, "--outputs", ",".join(selected),
                             "--titles", ",".join(titles) if titles else "none"]

                # Change to tmpdir so _playbook_work is created there
                original_cwd = os.getcwd()
                os.chdir(str(tmpdir))

                try:
                    pipeline_result = pipeline_main()
                    pipeline_warnings = visible_pipeline_warnings(
                        pipeline_result,
                        include_offense=gen_offense,
                    )
                finally:
                    os.chdir(original_cwd)
                    sys.argv = original_argv

            # Account deletion can arrive while LibreOffice/report generation
            # is running. Never upload its results after a cancellation marker.
            ensure_job_active(s3, bucket, job_id)

            # Upload PDFs to R2
            update_status(s3, bucket, job_id, {"status": "processing", "step": "uploading"}, request_fields)

            pdf_files = sorted(output_dir.glob("*.pdf"))
            uploaded = []

            for pdf in pdf_files:
                ensure_job_active(s3, bucket, job_id)
                key = f"jobs/{job_id}/{pdf.name}"
                print(f"  Uploading {pdf.name} ({pdf.stat().st_size} bytes)...")
                s3.upload_file(
                    str(pdf),
                    bucket,
                    key,
                    ExtraArgs={"ContentType": "application/pdf"},
                )
                uploaded.append(pdf.name)
                ensure_job_active(s3, bucket, job_id)

            if not uploaded:
                raise RuntimeError("Pipeline produced no PDF files")

            # Write final success status
            ensure_job_active(s3, bucket, job_id)
            complete_status = build_complete_status(uploaded, pipeline_warnings)
            update_status(s3, bucket, job_id, complete_status, request_fields)
            ensure_job_active(s3, bucket, job_id)
            print(f"Job {job_id} complete. Uploaded: {uploaded}")

    except Exception as e:
        error_msg = f"{type(e).__name__}: {e}"
        tb = traceback.format_exc()
        print(f"Job {job_id} failed: {error_msg}")
        print(tb)

        # Build a user-friendly message with detail
        if isinstance(e, JobCancelled):
            friendly = "This job was cancelled because its account was deleted."
            try:
                scrub_job_payload(s3, bucket, job_id)
            except Exception as scrub_error:
                print(f"Failed to scrub cancelled job payload: {scrub_error}")
        elif "LibreOffice" in str(e) or "soffice" in str(e):
            friendly = "LibreOffice conversion failed. The PPTX file may be corrupted or in an unsupported format."
        elif "pdftoppm" in str(e):
            friendly = "PDF to image conversion failed. This is a server-side dependency issue."
        elif "No PDF" in str(e) or "didn't produce" in str(e):
            friendly = "Could not convert the PowerPoint file to PDF. Please check the file is a valid .pptx."
        elif "no field rectangle" in str(e).lower():
            friendly = "Could not detect play diagrams in the playbook. Make sure slides have rectangle shapes marking the field area."
        elif "no play images" in str(e).lower():
            friendly = "No play images were found for this job. Please try generating again from the editor."
        elif "No slide images" in str(e) or "no PDF files" in str(e).lower():
            friendly = "Pipeline produced no output. The playbook may not have recognizable offense/defense sections."
        elif isinstance(e, ValueError):
            friendly = str(e)
        else:
            friendly = "Processing failed unexpectedly. Please try again or check the input file."

        try:
            update_status(s3, bucket, job_id, {
                "status": "error",
                "message": friendly,
            }, request_fields)
        except Exception:
            print("Failed to update error status in R2")
        sys.exit(1)


if __name__ == "__main__":
    main()
