"""Pydantic schemas for the Lennar Excel processor service."""
from datetime import datetime
from typing import Optional, Dict, Any, List
from pydantic import BaseModel, Field
from enum import Enum


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
    qa_report: Dict[str, Any] = Field(..., description="QA report with statistics")


class JobStatusResponse(BaseModel):
    """Response model for job status."""
    job_id: str = Field(..., description="Unique job identifier")
    status: JobStatus = Field(..., description="Current job status")
    progress: float = Field(0.0, ge=0.0, le=1.0, description="Job progress (0-1)")
    message: Optional[str] = Field(None, description="Status message")
    result: Optional[JobResult] = Field(None, description="Job results when succeeded")


class ParsedRow(BaseModel):
    """Model for a parsed row from Lennar Excel."""
    lot_block: Optional[str] = Field(None, description="Lot/Block identifier")
    plan: Optional[str] = Field(None, description="Plan identifier")
    elevation: Optional[str] = Field(None, description="Elevation")
    swing: Optional[str] = Field(None, description="Swing")
    task_start_date: Optional[datetime] = Field(None, description="Task start date")
    task_text: Optional[str] = Field(None, description="Task description")
    subtotal: Optional[float] = Field(None, description="Job subtotal amount")
    tax: Optional[float] = Field(None, description="Tax amount")
    total: Optional[float] = Field(None, description="Total amount")


class QAMeta(BaseModel):
    """QA metadata from parsing."""
    total_rows_seen: int = Field(0, description="Total rows processed")
    rows_parsed: int = Field(0, description="Successfully parsed rows")
    rows_skipped_missing_fields: int = Field(0, description="Rows skipped due to missing fields")


class SummaryRow(BaseModel):
    """Model for aggregated summary row."""
    lot_block: str = Field(..., description="Lot/Block identifier")
    plan: str = Field(..., description="Plan identifier")
    ext_prime: float = Field(0.0, description="EXT PRIME total")
    extere: float = Field(0.0, description="EXTERE total")
    exterior_ua: float = Field(0.0, description="EXTERIOR UA total")
    interior: float = Field(0.0, description="INTERIOR total")
    total: float = Field(0.0, description="Total amount")


class QAReport(BaseModel):
    """QA report model."""
    counts_per_bucket: Dict[str, int] = Field(..., description="Count of rows per bucket")
    unmapped_examples: List[Dict[str, Any]] = Field(..., description="Top unmapped task examples")
    suspicious_totals: List[Dict[str, Any]] = Field(..., description="Lots/plans with suspicious totals")
    parse_meta: QAMeta = Field(..., description="Parsing metadata")