"""API routes for the Capital One Card Report processor."""
import logging
import uuid
from datetime import UTC, datetime

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.core.config import settings
from app.core.security import validate_upload, verify_api_key
from app.modules.capital_one_card.services import (
    CapitalOneCardError,
    process_capital_one_card,
)

logger = logging.getLogger(__name__)

router = APIRouter(
    prefix="/api/v1/capital-one-card",
    tags=["capital-one-card"],
    dependencies=[Depends(verify_api_key)],
)


@router.post("/process")
async def process(file: UploadFile = File(...)):
    """Process a Capital One card transactions export (CSV or XLSX)."""
    content = await validate_upload(file, allowed_extensions={".csv", ".xlsx"})

    try:
        output_bytes = process_capital_one_card(content)
    except CapitalOneCardError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception:
        logger.exception("Capital One Card processing failed")
        raise HTTPException(status_code=500, detail="Error processing file")

    output_path = settings.OUTPUT_DIR / f"capital_one_card_{uuid.uuid4()}.xlsx"
    output_path.write_bytes(output_bytes)

    download_name = f"Capital-One-Card_Report_{datetime.now(UTC).strftime('%Y-%m-%d')}.xlsx"
    return FileResponse(
        path=output_path,
        filename=download_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
