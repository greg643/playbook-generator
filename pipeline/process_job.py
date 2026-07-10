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
from pathlib import Path

import boto3


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


def download_play_images(s3, bucket, job_id, plays_dir):
    """Download every object under jobs/<job_id>/plays/ from R2 (paginated)."""
    count = 0
    list_kwargs = {"Bucket": bucket, "Prefix": f"jobs/{job_id}/plays/"}
    while True:
        listing = s3.list_objects_v2(**list_kwargs)
        for obj in listing.get("Contents", []):
            name = obj["Key"].rsplit("/", 1)[-1]
            if not name:
                continue
            dest = plays_dir / name
            s3.download_file(bucket, obj["Key"], str(dest))
            print(f"  Downloaded {name} ({dest.stat().st_size} bytes)")
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
    bucket = os.environ["R2_BUCKET"]
    s3 = get_r2_client()

    print(f"Processing job: {job_id}")

    request_fields = {}
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
        request_fields = {k: request_data[k] for k in ("mode", "options", "createdAt") if k in request_data}
        mode = request_data.get("mode", "pptx")
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

                # Override sys.argv for the pipeline
                original_argv = sys.argv
                sys.argv = ["playbook_pipeline.py", str(pptx_path), str(output_dir),
                             "--sections", sections, "--outputs", ",".join(selected)]

                # Change to tmpdir so _playbook_work is created there
                original_cwd = os.getcwd()
                os.chdir(str(tmpdir))

                try:
                    pipeline_main()
                finally:
                    os.chdir(original_cwd)
                    sys.argv = original_argv

            # Upload PDFs to R2
            update_status(s3, bucket, job_id, {"status": "processing", "step": "uploading"}, request_fields)

            pdf_files = sorted(output_dir.glob("*.pdf"))
            uploaded = []

            for pdf in pdf_files:
                key = f"jobs/{job_id}/{pdf.name}"
                print(f"  Uploading {pdf.name} ({pdf.stat().st_size} bytes)...")
                s3.upload_file(
                    str(pdf),
                    bucket,
                    key,
                    ExtraArgs={"ContentType": "application/pdf"},
                )
                uploaded.append(pdf.name)

            if not uploaded:
                raise RuntimeError("Pipeline produced no PDF files")

            # Write final success status
            update_status(s3, bucket, job_id, {
                "status": "complete",
                "files": uploaded,
            }, request_fields)
            print(f"Job {job_id} complete. Uploaded: {uploaded}")

    except Exception as e:
        error_msg = f"{type(e).__name__}: {e}"
        tb = traceback.format_exc()
        print(f"Job {job_id} failed: {error_msg}")
        print(tb)

        # Build a user-friendly message with detail
        if "LibreOffice" in str(e) or "soffice" in str(e):
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
        else:
            friendly = str(e)

        try:
            update_status(s3, bucket, job_id, {
                "status": "error",
                "message": friendly,
                "detail": error_msg,
            }, request_fields)
        except Exception:
            print("Failed to update error status in R2")
        sys.exit(1)


if __name__ == "__main__":
    main()
