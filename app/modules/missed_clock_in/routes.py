"""API routes for the Missed Clock-In warning-notice generator."""
import io
import logging
import math
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from app.core.security import validate_upload, verify_api_key
from app.modules.missed_clock_in.generate_warnings import (
    TARGET_ERRORS,
    build_workbook,
    format_date,
    parse_exception_list,
)

logger = logging.getLogger(__name__)

router = APIRouter(
    prefix="/api/v1/missed-clock-in",
    tags=["missed-clock-in"],
    dependencies=[Depends(verify_api_key)],
)


def _clean(val: Any) -> str:
    """Return a JSON-safe string for a cell value from pandas (handles NaN/None)."""
    if val is None:
        return ""
    if isinstance(val, float) and math.isnan(val):
        return ""
    return str(val).strip()


def _parse_bytes(raw: bytes) -> list[dict]:
    """Write bytes to a temp file so pandas can read them, then parse."""
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        tmp.write(raw)
        tmp.close()
        return parse_exception_list(Path(tmp.name))
    finally:
        try:
            Path(tmp.name).unlink()
        except OSError:
            pass


@router.post("/preview")
async def preview_violations(file: UploadFile = File(...)):
    """Parse the Exception List and return a JSON summary without generating the workbook."""
    raw = await validate_upload(file, allowed_extensions={".xlsx"})

    try:
        records = _parse_bytes(raw)
    except Exception:
        logger.exception("Exception List parse failed")
        raise HTTPException(status_code=400, detail="Unable to parse Exception List")

    notices = [r for r in records if r["error"] in TARGET_ERRORS]
    clocked_twice = [r for r in records if r["error"] == "Clocked In Twice"]
    unique_employees = sorted({r["employee"] for r in records})

    preview = [
        {
            "employee": r["employee"],
            "date": format_date(r["date"]),
            "error": r["error"],
            "job": _clean(r["start_location"]) or _clean(r["stop_location"]),
        }
        for r in records
    ]

    return {
        "total_records": len(records),
        "notices": len(notices),
        "clocked_twice": len(clocked_twice),
        "unique_employees": len(unique_employees),
        "records": preview,
    }


@router.post("/process")
async def process_exception_list(file: UploadFile = File(...)):
    """Parse the Exception List and return the generated warning-notices workbook."""
    raw = await validate_upload(file, allowed_extensions={".xlsx"})

    try:
        records = _parse_bytes(raw)
    except Exception:
        logger.exception("Exception List parse failed")
        raise HTTPException(status_code=400, detail="Unable to parse Exception List")

    if not records:
        raise HTTPException(
            status_code=400,
            detail="No tracked violations found in file.",
        )

    try:
        wb = build_workbook(records)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
    except Exception:
        logger.exception("Warning-notice workbook generation failed")
        raise HTTPException(status_code=500, detail="Error generating output file")

    filename = f"Warning_Notices_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
