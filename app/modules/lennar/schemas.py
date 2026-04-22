"""Pydantic schemas for the Lennar Excel processor service."""
from datetime import datetime
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class JobStatus(str, Enum):
    """Job status enum."""
    QUEUED = "queued"
    RUNNING = "running"
    SUCCEEDED = "succeeded"
    FAILED = "failed"


class UploadResponse(BaseModel):
    """Response model for file upload."""
    job_id: str = Field(..., description="Unique job identifier")


class JobResult(BaseModel):
    """Job result model."""
    output_path: str = Field(..., description="Path to the generated Excel file")
    qa_report: dict[str, Any] = Field(..., description="QA report with statistics")


class JobStatusResponse(BaseModel):
    """Response model for job status."""
    job_id: str = Field(..., description="Unique job identifier")
    status: JobStatus = Field(..., description="Current job status")
    progress: float = Field(0.0, ge=0.0, le=1.0, description="Job progress (0-1)")
    message: str | None = Field(None, description="Status message")
    result: JobResult | None = Field(None, description="Job results when succeeded")


class ParsedRow(BaseModel):
    """Model for a parsed row from Lennar Excel."""
    lot_block: str | None = Field(None, description="Lot/Block identifier")
    plan: str | None = Field(None, description="Plan identifier")
    elevation: str | None = Field(None, description="Elevation")
    swing: str | None = Field(None, description="Swing")
    task_start_date: datetime | None = Field(None, description="Task start date")
    task_text: str | None = Field(None, description="Normalized task name (extracted from raw)")
    task_text_raw: str | None = Field(None, description="Original task text for audit/debugging")
    subtotal: float | None = Field(None, description="Job subtotal amount")
    tax: float | None = Field(None, description="Tax amount")
    total: float | None = Field(None, description="Total amount")


class QAMeta(BaseModel):
    """QA metadata from parsing."""
    total_rows_seen: int = Field(0, description="Total rows processed")
    rows_parsed: int = Field(0, description="Successfully parsed rows")
    rows_skipped_missing_fields: int = Field(0, description="Rows skipped due to missing fields")


class QAReport(BaseModel):
    """QA report model."""
    counts_per_bucket: dict[str, int] = Field(..., description="Count of rows per bucket")
    unmapped_examples: list[dict[str, Any]] = Field(..., description="Top unmapped task examples")
    suspicious_totals: list[dict[str, Any]] = Field(..., description="Lots/plans with suspicious totals")
    parse_meta: QAMeta = Field(..., description="Parsing metadata")
