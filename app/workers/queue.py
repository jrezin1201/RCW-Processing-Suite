"""
RQ queue configuration and utilities.
Provides Redis connection and queue instances for job management.
"""

from typing import Any, Callable

from redis import Redis
from rq import Queue

from app.core.config import settings
from app.core.logging import get_logger

logger = get_logger(__name__)

# Redis connection for RQ
redis_conn = Redis(
    host=settings.REDIS_HOST,
    port=settings.REDIS_PORT,
    db=settings.REDIS_DB,
    decode_responses=True,
)

# Default queue for background tasks
default_queue = Queue("default", connection=redis_conn)


def enqueue_task(func: Callable[..., Any], *args: Any, **kwargs: Any) -> str:
    """
    Enqueue a background task.

    Args:
        func: Function to execute
        *args: Positional arguments for the function
        **kwargs: Keyword arguments for the function

    Returns:
        Job ID
    """
    job = default_queue.enqueue(func, *args, **kwargs)
    logger.info(f"Enqueued task {func.__name__} with job ID: {job.id}")
    return job.id


def get_job_status(job_id: str) -> dict[str, Any]:
    """
    Get the status of a background job.

    Args:
        job_id: Job ID to check

    Returns:
        Job status information
    """
    from rq.job import Job

    try:
        job = Job.fetch(job_id, connection=redis_conn)
        return {
            "job_id": job.id,
            "status": job.get_status(),
            "result": job.result if job.is_finished else None,
            "error": str(job.exc_info) if job.is_failed else None,
        }
    except Exception as e:
        logger.error(f"Error fetching job {job_id}: {e}")
        return {
            "job_id": job_id,
            "status": "not_found",
            "error": str(e),
        }
