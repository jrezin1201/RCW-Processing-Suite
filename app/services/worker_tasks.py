"""Background worker tasks for processing Lennar Excel files."""
import traceback
import json
import logging
from typing import Dict, Any

from app.services.jobs import set_job, update_job_progress
from app.services.parser_lennar import parse_lennar_export
from app.services.aggregator import aggregate_data
from app.services.excel_writer import write_summary_excel

logger = logging.getLogger(__name__)


def process_lennar_file(job_id: str, filepath: str, original_filename: str = None) -> None:
    """
    Process a Lennar scheduled tasks Excel file.

    This is the main orchestrator function that runs in the background.

    Args:
        job_id: The job identifier
        filepath: Path to the uploaded Excel file
        original_filename: Original uploaded filename (without extension)
    """
    try:
        # Update job status to running
        set_job(job_id, {
            "job_id": job_id,
            "status": "running",
            "progress": 0.0,
            "message": "Starting file processing",
            "result": None
        })

        # Step 1: Parse the Excel file
        update_job_progress(job_id, 0.1, "Parsing Excel file")
        parsed_rows, qa_meta, phase, project_name, house_string = parse_lennar_export(filepath)

        if not parsed_rows:
            set_job(job_id, {
                "job_id": job_id,
                "status": "failed",
                "progress": 0.1,
                "message": "No valid data rows found in Excel file",
                "result": None
            })
            return

        # Step 2: Classify and aggregate data (with auto-category creation)
        update_job_progress(job_id, 0.5, f"Processing {len(parsed_rows)} rows")
        summary_rows, qa_report, category_headers = aggregate_data(parsed_rows, qa_meta)

        # Log category summary
        logger.info(f"Categories used: {category_headers}")
        auto_created = [h for h in category_headers if h not in [
            "EXT PRIME", "EXTERIOR", "EXTERIOR UA", "INTERIOR",
            "BASE SHOE", "ROLL WALLS FINAL", "TOUCH UP", "Q4 REVERSAL"
        ]]
        if auto_created:
            logger.info(f"Auto-created categories: {auto_created}")

        # Step 3: Write output Excel file
        update_job_progress(job_id, 0.8, "Generating output Excel file")
        output_path = write_summary_excel(
            summary_rows, qa_report, job_id, category_headers,
            phase=phase, project_name=project_name, house_string=house_string,
            original_filename=original_filename
        )

        # Step 4: Prepare result data
        result_data = {
            "output_path": output_path,
            "qa_report": {
                "parse_meta": {
                    "total_rows_seen": qa_report.parse_meta.total_rows_seen,
                    "rows_parsed": qa_report.parse_meta.rows_parsed,
                    "rows_skipped_missing_fields": qa_report.parse_meta.rows_skipped_missing_fields
                },
                "counts_per_bucket": qa_report.counts_per_bucket,
                "unmapped_count": len(qa_report.unmapped_examples),
                "unmapped_examples": qa_report.unmapped_examples[:10],  # Limit for display
                "suspicious_totals_count": len(qa_report.suspicious_totals),
                "suspicious_totals": qa_report.suspicious_totals[:10],  # Limit for display
                "summary_rows_generated": len(summary_rows)
            }
        }

        # Mark job as succeeded
        set_job(job_id, {
            "job_id": job_id,
            "status": "succeeded",
            "progress": 1.0,
            "message": f"Successfully processed {len(parsed_rows)} rows into {len(summary_rows)} summary rows",
            "result": result_data
        })

    except Exception as e:
        # Handle any errors
        error_message = str(e)
        error_traceback = traceback.format_exc()

        # Truncate traceback if too long
        if len(error_traceback) > 1000:
            error_traceback = error_traceback[:1000] + "... (truncated)"

        set_job(job_id, {
            "job_id": job_id,
            "status": "failed",
            "progress": 0.0,
            "message": f"Processing failed: {error_message}\n\nTraceback:\n{error_traceback}",
            "result": None
        })