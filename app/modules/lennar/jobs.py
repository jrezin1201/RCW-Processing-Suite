"""Job management service using in-memory storage."""
import uuid
from typing import Dict, Any, Optional
import logging

logger = logging.getLogger(__name__)

# In-memory job storage
job_storage: Dict[str, Dict[str, Any]] = {}


def enqueue_job(func, *args, **kwargs):
    """
    Run a job synchronously.

    Args:
        func: The function to execute
        *args: Function arguments
        **kwargs: Function keyword arguments

    Returns:
        A job ID string
    """
    fake_job_id = str(uuid.uuid4())

    try:
        func(*args, **kwargs)
    except Exception as e:
        logger.error(f"Error running job {fake_job_id}: {e}")

    return fake_job_id


def set_job(job_id: str, data: Dict[str, Any]) -> None:
    """
    Store job status and data.

    Args:
        job_id: The job identifier
        data: Job data including status, progress, message, result
    """
    job_storage[job_id] = data


def get_job(job_id: str) -> Optional[Dict[str, Any]]:
    """
    Retrieve job status and data.

    Args:
        job_id: The job identifier

    Returns:
        Job data or None if not found
    """
    return job_storage.get(job_id)


def update_job_progress(job_id: str, progress: float, message: str = "") -> None:
    """
    Update job progress.

    Args:
        job_id: The job identifier
        progress: Progress value (0.0 to 1.0)
        message: Optional status message
    """
    job_data = get_job(job_id)
    if job_data:
        job_data["progress"] = progress
        if message:
            job_data["message"] = message
        set_job(job_id, job_data)
