"""API routes for the Merchant Charges processor."""
import logging
import uuid
from datetime import UTC, datetime

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.core.config import settings
from app.core.security import validate_upload, verify_api_key
from app.modules.merchant_charges.services import (
    MerchantChargesError,
    process_merchant_charges,
)

logger = logging.getLogger(__name__)

router = APIRouter(
    prefix="/api/v1/merchant-charges",
    tags=["merchant-charges"],
    dependencies=[Depends(verify_api_key)],
)


@router.post("/process")
async def process(file: UploadFile = File(...)):
    """Process a Merchant Charges credit-card transactions Excel file."""
    content = await validate_upload(file, allowed_extensions={".xlsx"})

    try:
        output_bytes = process_merchant_charges(content)
    except MerchantChargesError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception:
        logger.exception("Merchant Charges processing failed")
        raise HTTPException(status_code=500, detail="Error processing file")

    output_path = settings.OUTPUT_DIR / f"merchant_charges_{uuid.uuid4()}.xlsx"
    output_path.write_bytes(output_bytes)

    download_name = f"Merchant-Charges_Report_{datetime.now(UTC).strftime('%Y-%m-%d')}.xlsx"
    return FileResponse(
        path=output_path,
        filename=download_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
