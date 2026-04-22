"""Shared security + input-validation helpers.

Used by every module so validation logic doesn't drift across modules.
"""
from __future__ import annotations

import secrets
from pathlib import Path

from fastapi import Header, HTTPException, UploadFile, status

from app.core.config import settings

# Magic-byte signatures. openpyxl/pandas can be fed arbitrary bytes — checking
# the extension alone is trivially spoofable, so we also sniff the first few
# bytes to ensure the upload is actually an Office file (xlsx is a ZIP, xls is
# the OLE compound document).
_XLSX_MAGIC = b"PK\x03\x04"
_XLS_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


async def verify_api_key(x_api_key: str | None = Header(default=None, alias="X-API-Key")) -> None:
    """FastAPI dependency: if API_KEY is configured, require it on the request."""
    if settings.API_KEY is None:
        return
    if x_api_key is None or not secrets.compare_digest(x_api_key, settings.API_KEY):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid or missing API key",
        )


async def validate_upload(
    file: UploadFile,
    allowed_extensions: set[str] | None = None,
) -> bytes:
    """Read + validate an UploadFile. Returns the raw bytes on success.

    Checks:
      1. Filename has an allowed extension.
      2. Byte length is within MAX_UPLOAD_SIZE_BYTES.
      3. First bytes match a known xlsx/xls magic signature.

    Raises HTTPException(400) on any failure.
    """
    allowed = allowed_extensions if allowed_extensions is not None else settings.ALLOWED_EXTENSIONS

    filename = file.filename or ""
    ext = Path(filename).suffix.lower()
    if ext not in allowed:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file type. Allowed: {', '.join(sorted(allowed))}",
        )

    content = await file.read()
    if len(content) == 0:
        raise HTTPException(status_code=400, detail="Uploaded file is empty")
    if len(content) > settings.MAX_UPLOAD_SIZE_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"File exceeds {settings.MAX_UPLOAD_SIZE_MB} MB limit",
        )

    # Magic-byte sniff. xlsx files are ZIP archives; xls is OLE2.
    head = content[:8]
    is_xlsx = head.startswith(_XLSX_MAGIC)
    is_xls = head.startswith(_XLS_MAGIC)
    if ext == ".xlsx" and not is_xlsx:
        raise HTTPException(status_code=400, detail="File is not a valid .xlsx workbook")
    if ext == ".xls" and not is_xls:
        raise HTTPException(status_code=400, detail="File is not a valid .xls workbook")

    return content


def safe_path_under(base: Path, candidate: Path | str) -> Path:
    """Return `candidate` resolved, but only if it lives under `base`.

    Defense against path traversal on endpoints that serve files by path
    (e.g. download-by-job-id). Raises HTTPException(404) if the resolved
    path escapes the base directory.
    """
    base_resolved = base.resolve()
    resolved = Path(candidate).resolve()
    try:
        resolved.relative_to(base_resolved)
    except ValueError:
        raise HTTPException(status_code=404, detail="File not found") from None
    if not resolved.exists():
        raise HTTPException(status_code=404, detail="File not found")
    return resolved
