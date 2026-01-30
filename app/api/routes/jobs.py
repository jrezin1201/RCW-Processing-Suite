"""
Background job routes for enqueueing and monitoring tasks.
Demonstrates async task processing with RQ.
"""

from typing import Annotated

from fastapi import APIRouter, Depends

from app.api.deps import get_current_active_user
from app.models.user import User
from app.workers.queue import enqueue_task, get_job_status
from app.workers.tasks import example_task, process_data_task, send_email_task

router = APIRouter(prefix="/jobs", tags=["jobs"])


@router.post("/example")
def create_example_job(
    name: str,
    duration: int = 5,
    current_user: User = Depends(get_current_active_user),
) -> dict:
    """
    Create an example background job.

    Args:
        name: Name for the task
        duration: How long the task should run
        current_user: Current authenticated user

    Returns:
        Job information including job_id
    """
    job_id = enqueue_task(example_task, name=name, duration=duration)

    return {
        "message": "Job enqueued successfully",
        "job_id": job_id,
        "task": "example_task",
        "parameters": {"name": name, "duration": duration},
    }


@router.post("/send-email")
def create_email_job(
    to: str,
    subject: str,
    body: str,
    current_user: User = Depends(get_current_active_user),
) -> dict:
    """
    Enqueue an email sending job.

    Args:
        to: Recipient email
        subject: Email subject
        body: Email body
        current_user: Current authenticated user

    Returns:
        Job information
    """
    job_id = enqueue_task(send_email_task, to=to, subject=subject, body=body)

    return {
        "message": "Email job enqueued successfully",
        "job_id": job_id,
        "task": "send_email_task",
    }


@router.get("/{job_id}/status")
def get_job_status_endpoint(
    job_id: str,
    current_user: User = Depends(get_current_active_user),
) -> dict:
    """
    Get the status of a background job.

    Args:
        job_id: Job ID to check
        current_user: Current authenticated user

    Returns:
        Job status information
    """
    return get_job_status(job_id)
