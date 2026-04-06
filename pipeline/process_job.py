#!/usr/bin/env python3
"""
GitHub Actions job orchestrator.

Downloads a PPTX from R2, runs the playbook pipeline, uploads PDFs back to R2.

Usage:
    python pipeline/process_job.py <job_id>

Environment variables:
    R2_ENDPOINT, R2_ACCESS_KEY_ID, R2_SECRET_ACCESS_KEY, R2_BUCKET
"""

import os
import sys
import json
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


def update_status(s3, bucket, job_id, status_dict):
    """Write status.json to R2."""
    s3.put_object(
        Bucket=bucket,
        Key=f"{job_id}/status.json",
        Body=json.dumps(status_dict),
        ContentType="application/json",
    )


def main():
    if len(sys.argv) < 2:
        print("Usage: python pipeline/process_job.py <job_id>")
        sys.exit(1)

    job_id = sys.argv[1]
    bucket = os.environ["R2_BUCKET"]
    s3 = get_r2_client()

    print(f"Processing job: {job_id}")

    try:
        # Update status to processing
        update_status(s3, bucket, job_id, {"status": "processing", "step": "downloading"})

        # Create temporary working directory
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            pptx_path = tmpdir / "input.pptx"
            output_dir = tmpdir / "output"
            output_dir.mkdir()

            # Read options from status.json
            print("Reading job options from R2...")
            status_obj = s3.get_object(Bucket=bucket, Key=f"{job_id}/status.json")
            status_data = json.loads(status_obj["Body"].read())
            options = status_data.get("options", {"offense": True, "defense": True})
            gen_offense = options.get("offense", True)
            gen_defense = options.get("defense", True)
            sections = "both" if (gen_offense and gen_defense) else ("offense" if gen_offense else "defense")
            print(f"  Sections: {sections}")

            # Download PPTX from R2
            print("Downloading PPTX from R2...")
            s3.download_file(bucket, f"{job_id}/input.pptx", str(pptx_path))
            print(f"  Downloaded {pptx_path.stat().st_size} bytes")

            # Run the pipeline
            update_status(s3, bucket, job_id, {"status": "processing", "step": "generating"})

            # Import and run pipeline from the same package
            pipeline_dir = Path(__file__).parent
            sys.path.insert(0, str(pipeline_dir))
            from playbook_pipeline import main as pipeline_main

            # Override sys.argv for the pipeline
            original_argv = sys.argv
            sys.argv = ["playbook_pipeline.py", str(pptx_path), str(output_dir), "--sections", sections]

            # Change to tmpdir so _playbook_work is created there
            original_cwd = os.getcwd()
            os.chdir(str(tmpdir))

            try:
                pipeline_main()
            finally:
                os.chdir(original_cwd)
                sys.argv = original_argv

            # Upload PDFs to R2
            update_status(s3, bucket, job_id, {"status": "processing", "step": "uploading"})

            pdf_files = sorted(output_dir.glob("*.pdf"))
            uploaded = []

            for pdf in pdf_files:
                key = f"{job_id}/{pdf.name}"
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
            })
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
        elif "No slide images" in str(e) or "no PDF files" in str(e).lower():
            friendly = "Pipeline produced no output. The playbook may not have recognizable offense/defense sections."
        else:
            friendly = str(e)

        try:
            update_status(s3, bucket, job_id, {
                "status": "error",
                "message": friendly,
                "detail": error_msg,
            })
        except Exception:
            print("Failed to update error status in R2")
        sys.exit(1)


if __name__ == "__main__":
    main()
