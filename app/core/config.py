"""
Configuration for Lennar Excel Processor Service.
"""

import os
from typing import Optional


class Settings:
    """Application settings."""

    # Application
    PROJECT_NAME: str = "Lennar Excel Processor"
    VERSION: str = "1.0.0"
    API_V1_PREFIX: str = "/api/v1"
    DEBUG: bool = os.getenv("DEBUG", "True").lower() == "true"

    # Redis
    REDIS_URL: str = os.getenv("REDIS_URL", "redis://localhost:6379/0")

    # File Storage
    UPLOAD_DIR: str = os.getenv("UPLOAD_DIR", "data/uploads")
    OUTPUT_DIR: str = os.getenv("OUTPUT_DIR", "data/outputs")

    # Processing Configuration
    MAX_UPLOAD_SIZE_MB: int = int(os.getenv("MAX_UPLOAD_SIZE_MB", "50"))
    ALLOWED_EXTENSIONS: list = [".xlsx", ".xls"]

    # Job Processing
    JOB_TIMEOUT_SECONDS: int = int(os.getenv("JOB_TIMEOUT_SECONDS", "300"))
    MAX_CONCURRENT_JOBS: int = int(os.getenv("MAX_CONCURRENT_JOBS", "5"))

    # Environment
    ENVIRONMENT: str = os.getenv("ENVIRONMENT", "development")


settings = Settings()
