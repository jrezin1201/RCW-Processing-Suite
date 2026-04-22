"""Environment-driven settings for the RCW Processing Suite."""
from __future__ import annotations

import os
from pathlib import Path


def _bool(name: str, default: bool = False) -> bool:
    return os.getenv(name, "true" if default else "false").strip().lower() in {"1", "true", "yes", "on"}


def _int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, str(default)))
    except ValueError:
        return default


def _list(name: str, default: list[str]) -> list[str]:
    raw = os.getenv(name, "").strip()
    if not raw:
        return default
    return [item.strip() for item in raw.split(",") if item.strip()]


class Settings:
    """Application settings. Production defaults are safe; dev overrides via env."""

    PROJECT_NAME: str = "RCW Processing Suite"
    VERSION: str = "1.1.0"

    # development | staging | production
    ENVIRONMENT: str = os.getenv("ENVIRONMENT", "production").strip().lower()
    DEBUG: bool = _bool("DEBUG", default=False)

    # CORS — production default is deny-all; supply CORS_ORIGINS to allowlist.
    # Example: CORS_ORIGINS="https://rcw.example.com,https://app.rcw.example.com"
    CORS_ORIGINS: list[str] = _list("CORS_ORIGINS", default=[])

    # OpenAPI docs are off by default; enable explicitly via ENABLE_DOCS=true
    # (or implicitly in development).
    ENABLE_DOCS: bool = _bool("ENABLE_DOCS", default=False) or ENVIRONMENT == "development"

    # Optional API key. If unset, endpoints are unauthenticated. If set, every
    # mutating endpoint requires `X-API-Key: <value>` to match.
    API_KEY: str | None = os.getenv("API_KEY") or None

    # File storage (resolved to absolute paths so path-traversal checks work).
    BASE_DIR: Path = Path(__file__).resolve().parents[2]
    UPLOAD_DIR: Path = (BASE_DIR / os.getenv("UPLOAD_DIR", "data/uploads")).resolve()
    OUTPUT_DIR: Path = (BASE_DIR / os.getenv("OUTPUT_DIR", "data/outputs")).resolve()

    # Upload limits
    MAX_UPLOAD_SIZE_MB: int = _int("MAX_UPLOAD_SIZE_MB", 50)
    MAX_UPLOAD_SIZE_BYTES: int = MAX_UPLOAD_SIZE_MB * 1024 * 1024

    ALLOWED_EXTENSIONS: set[str] = {".xlsx", ".xls"}

    # Lennar job processing
    JOB_TIMEOUT_SECONDS: int = _int("JOB_TIMEOUT_SECONDS", 300)


settings = Settings()

# Ensure data dirs exist at import time so upload/download code can trust them.
settings.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
settings.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
