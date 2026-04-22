"""API routes for the Gas & Rig job-cost processor."""
import logging
import uuid
from pathlib import Path

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.core.config import settings
from app.core.security import validate_upload, verify_api_key
from app.modules.gas_rig.services import (
    build_output_workbook,
    compute_job_costs_from_xlsx,
)

logger = logging.getLogger(__name__)

router = APIRouter(
    prefix="/api/v1/gas-rig",
    tags=["gas-rig"],
    dependencies=[Depends(verify_api_key)],
)


@router.post("/process")
async def process_gas_rig(file: UploadFile = File(...)):
    """Process a Gas & Rig hours-worked Excel file."""
    content = await validate_upload(file, allowed_extensions={".xlsx"})

    try:
        rows = compute_job_costs_from_xlsx(content)
    except Exception:
        logger.exception("Gas & Rig processing failed")
        raise HTTPException(status_code=500, detail="Error processing file")

    if not rows:
        raise HTTPException(
            status_code=400,
            detail="No job data found. Please ensure the file contains location data with 4-digit job numbers.",
        )

    try:
        output_bytes = build_output_workbook(rows)
    except Exception:
        logger.exception("Gas & Rig output generation failed")
        raise HTTPException(status_code=500, detail="Error generating output file")

    output_path = settings.OUTPUT_DIR / f"gas_rig_{uuid.uuid4()}.xlsx"
    output_path.write_bytes(output_bytes)

    download_name = f"gas_rig_summary_{Path(file.filename or 'input').name}"
    return FileResponse(
        path=output_path,
        filename=download_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
