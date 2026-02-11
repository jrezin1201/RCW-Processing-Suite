"""API routes for the Lennar Excel processor and Gas & Rig processor services."""
import os
import uuid
import logging
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, File, UploadFile, HTTPException
from fastapi.responses import FileResponse

from app.models.schemas import UploadResponse, JobStatusResponse
from app.services.jobs import enqueue_job, get_job, set_job
from app.services.gas_rig import compute_job_costs_from_xlsx, build_output_workbook

logger = logging.getLogger(__name__)


router = APIRouter(prefix="/api/v1")


@router.post("/uploads", response_model=UploadResponse)
async def upload_file(file: UploadFile = File(...)) -> UploadResponse:
    """
    Upload an Excel file for processing.

    Args:
        file: The uploaded Excel file

    Returns:
        UploadResponse with job_id
    """
    # Validate file type
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="File must be an Excel file (.xlsx or .xls)")

    # Generate unique job ID
    job_id = str(uuid.uuid4())

    # Create uploads directory if it doesn't exist
    upload_dir = Path("data/uploads")
    upload_dir.mkdir(parents=True, exist_ok=True)

    # Save uploaded file
    file_path = upload_dir / f"{job_id}.xlsx"
    try:
        content = await file.read()
        with open(file_path, "wb") as f:
            f.write(content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save file: {str(e)}")

    # Initialize job status
    set_job(job_id, {
        "status": "queued",
        "progress": 0.0,
        "message": "File uploaded, job queued",
        "result": None
    })

    # Run processing job
    from app.services.worker_tasks import process_lennar_file
    original_filename = Path(file.filename).stem  # Get filename without extension
    enqueue_job(process_lennar_file, job_id, str(file_path), original_filename)

    return UploadResponse(job_id=job_id)


@router.get("/jobs/{job_id}", response_model=JobStatusResponse)
async def get_job_status(job_id: str) -> JobStatusResponse:
    """
    Get the status of a processing job.

    Args:
        job_id: The job identifier

    Returns:
        JobStatusResponse with current job status
    """
    job_data = get_job(job_id)

    if not job_data:
        raise HTTPException(status_code=404, detail=f"Job {job_id} not found")

    return JobStatusResponse(**job_data)


@router.get("/jobs/{job_id}/download")
async def download_result(job_id: str):
    """
    Download the processed Excel file.

    Args:
        job_id: The job identifier

    Returns:
        The processed Excel file
    """
    job_data = get_job(job_id)

    if not job_data:
        raise HTTPException(status_code=404, detail=f"Job {job_id} not found")

    if job_data["status"] != "succeeded":
        raise HTTPException(
            status_code=400,
            detail=f"Job {job_id} has not completed successfully. Current status: {job_data['status']}"
        )

    if not job_data.get("result") or not job_data["result"].get("output_path"):
        raise HTTPException(status_code=404, detail=f"Output file not found for job {job_id}")

    output_path = job_data["result"]["output_path"]

    if not os.path.exists(output_path):
        raise HTTPException(status_code=404, detail=f"Output file not found: {output_path}")

    # Use the actual filename from the output path
    download_filename = Path(output_path).name
    return FileResponse(
        path=output_path,
        filename=download_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@router.post("/gas-rig/process")
async def process_gas_rig(file: UploadFile = File(...)):
    """
    Process a Gas & Rig hours worked Excel file.

    Args:
        file: The uploaded Excel file (.xlsx)

    Returns:
        The processed summary Excel file with job costs
    """
    # Validate file type
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="File must be an Excel file (.xlsx)")

    try:
        # Read file content
        content = await file.read()

        # Process the file
        rows = compute_job_costs_from_xlsx(content)

        if not rows:
            raise HTTPException(
                status_code=400,
                detail="No job data found in file. Please ensure the file contains location data with 4-digit job numbers."
            )

        # Build output workbook
        output_bytes = build_output_workbook(rows)

        # Save output temporarily
        output_dir = Path("data/outputs")
        output_dir.mkdir(parents=True, exist_ok=True)

        output_filename = f"gas_rig_{uuid.uuid4()}.xlsx"
        output_path = output_dir / output_filename

        with open(output_path, "wb") as f:
            f.write(output_bytes)

        # Return the file as a response
        return FileResponse(
            path=output_path,
            filename=f"gas_rig_summary_{file.filename}",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        logger.error(f"Error processing gas & rig file: {e}")
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")