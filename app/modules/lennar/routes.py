"""API routes for the Lennar Excel processor."""
import logging
import uuid
from pathlib import Path

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.core.config import settings
from app.core.security import safe_path_under, validate_upload, verify_api_key
from app.modules.lennar.jobs import enqueue_job, get_job, set_job
from app.modules.lennar.schemas import JobStatusResponse, UploadResponse
from app.modules.lennar.worker_tasks import process_lennar_file

logger = logging.getLogger(__name__)

router = APIRouter(
    prefix="/api/v1",
    tags=["lennar"],
    dependencies=[Depends(verify_api_key)],
)


@router.post("/uploads", response_model=UploadResponse)
async def upload_file(file: UploadFile = File(...)) -> UploadResponse:
    """Upload a Lennar Excel file for processing."""
    content = await validate_upload(file)

    job_id = str(uuid.uuid4())
    file_path = settings.UPLOAD_DIR / f"{job_id}.xlsx"
    try:
        file_path.write_bytes(content)
    except OSError:
        logger.exception("Failed to persist upload for job %s", job_id)
        raise HTTPException(status_code=500, detail="Failed to save file")

    set_job(job_id, {
        "job_id": job_id,
        "status": "queued",
        "progress": 0.0,
        "message": "File uploaded, job queued",
        "result": None,
    })

    original_filename = Path(file.filename or "upload").stem
    enqueue_job(process_lennar_file, job_id, str(file_path), original_filename)

    return UploadResponse(job_id=job_id)


@router.get("/jobs/{job_id}", response_model=JobStatusResponse)
async def get_job_status(job_id: str) -> JobStatusResponse:
    """Get the status of a processing job."""
    try:
        uuid.UUID(job_id)
    except ValueError:
        raise HTTPException(status_code=404, detail="Job not found")

    job_data = get_job(job_id)
    if not job_data:
        raise HTTPException(status_code=404, detail="Job not found")
    return JobStatusResponse(**job_data)


@router.get("/jobs/{job_id}/download")
async def download_result(job_id: str):
    """Download the processed Excel file."""
    try:
        uuid.UUID(job_id)
    except ValueError:
        raise HTTPException(status_code=404, detail="Job not found")

    job_data = get_job(job_id)
    if not job_data:
        raise HTTPException(status_code=404, detail="Job not found")

    if job_data.get("status") != "succeeded":
        raise HTTPException(
            status_code=400,
            detail=f"Job not ready. Status: {job_data.get('status')}",
        )

    result = job_data.get("result") or {}
    output_path = result.get("output_path")
    if not output_path:
        raise HTTPException(status_code=404, detail="Output file not found")

    # Defense in depth: ensure the recorded output path lives under OUTPUT_DIR.
    safe = safe_path_under(settings.OUTPUT_DIR, output_path)
    return FileResponse(
        path=safe,
        filename=safe.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
