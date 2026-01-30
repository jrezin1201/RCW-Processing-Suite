"""Job management service using Redis and RQ (with fallback to in-memory)."""
import os
from typing import Dict, Any, Optional
import logging

logger = logging.getLogger(__name__)

# Try to import Redis and RQ
try:
    import redis
    from rq import Queue
    from rq.job import Job
    REDIS_AVAILABLE = True
except ImportError:
    REDIS_AVAILABLE = False
    logger.warning("Redis/RQ not available, using in-memory storage only")

# Try to connect to Redis
redis_conn = None
queue = None

if REDIS_AVAILABLE:
    try:
        REDIS_URL = os.getenv("REDIS_URL", "redis://localhost:6379/0")
        redis_conn = redis.from_url(REDIS_URL)
        # Test connection
        redis_conn.ping()
        queue = Queue("default", connection=redis_conn)
        logger.info("Connected to Redis successfully")
    except Exception as e:
        logger.warning(f"Could not connect to Redis: {e}. Using in-memory storage only")
        redis_conn = None
        queue = None

# In-memory job storage (fallback when Redis is not available)
job_storage: Dict[str, Dict[str, Any]] = {}


def enqueue_job(func, *args, **kwargs):
    """
    Enqueue a job to the RQ queue (or simulate if Redis not available).

    Args:
        func: The function to execute
        *args: Function arguments
        **kwargs: Function keyword arguments

    Returns:
        The enqueued RQ job or a fake job ID
    """
    if queue:
        return queue.enqueue(func, *args, **kwargs)
    else:
        # Simulate job enqueueing when Redis is not available
        # In a real scenario, you might want to run this in a thread
        import uuid
        fake_job_id = str(uuid.uuid4())
        logger.warning(f"Redis not available, job {fake_job_id} will run synchronously")

        # Run the function synchronously (for testing only)
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
    # Always store in memory
    job_storage[job_id] = data

    # Also store in Redis if available
    if redis_conn:
        try:
            redis_conn.hset(f"job:{job_id}", mapping={
                "status": data.get("status", "queued"),
                "progress": str(data.get("progress", 0.0)),
                "message": data.get("message", ""),
                "result": str(data.get("result", ""))
            })
        except Exception as e:
            logger.warning(f"Could not store job in Redis: {e}")


def get_job(job_id: str) -> Optional[Dict[str, Any]]:
    """
    Retrieve job status and data.

    Args:
        job_id: The job identifier

    Returns:
        Job data or None if not found
    """
    # Try memory first
    if job_id in job_storage:
        return job_storage[job_id]

    # Try Redis if available
    if redis_conn:
        try:
            job_data = redis_conn.hgetall(f"job:{job_id}")
            if job_data:
                # Convert bytes to strings and parse
                return {
                    "job_id": job_id,
                    "status": job_data.get(b"status", b"").decode("utf-8"),
                    "progress": float(job_data.get(b"progress", b"0.0").decode("utf-8")),
                    "message": job_data.get(b"message", b"").decode("utf-8"),
                    "result": eval(job_data.get(b"result", b"None").decode("utf-8")) if job_data.get(b"result", b"None") != b"None" else None
                }
        except Exception as e:
            logger.warning(f"Could not retrieve job from Redis: {e}")

    return None


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