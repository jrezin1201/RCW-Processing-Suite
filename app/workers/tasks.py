"""
Background tasks using RQ (Redis Queue).
Define long-running or asynchronous tasks here.
"""

import time
from typing import Any

from app.core.logging import get_logger

logger = get_logger(__name__)


def example_task(name: str, duration: int = 5) -> dict[str, Any]:
    """
    Example background task that simulates a long-running operation.

    Args:
        name: Name identifier for the task
        duration: How long the task should run (seconds)

    Returns:
        Task result dictionary
    """
    logger.info(f"Starting example task: {name} (duration: {duration}s)")

    # Simulate long-running work
    for i in range(duration):
        time.sleep(1)
        logger.debug(f"Task {name} progress: {i + 1}/{duration}")

    result = {
        "task_name": name,
        "duration": duration,
        "status": "completed",
        "message": f"Task {name} completed successfully",
    }

    logger.info(f"Completed example task: {name}")
    return result


def send_email_task(to: str, subject: str, body: str) -> dict[str, Any]:
    """
    Example email sending task.
    In production, this would integrate with an email service.

    Args:
        to: Recipient email address
        subject: Email subject
        body: Email body

    Returns:
        Task result dictionary
    """
    logger.info(f"Sending email to {to}: {subject}")

    # Simulate email sending
    time.sleep(2)

    # In production, integrate with SendGrid, AWS SES, etc.
    logger.info(f"Email sent to {to}")

    return {
        "to": to,
        "subject": subject,
        "status": "sent",
        "message": "Email sent successfully",
    }


def process_data_task(data: dict[str, Any]) -> dict[str, Any]:
    """
    Example data processing task.
    Useful for ETL operations, data transformations, etc.

    Args:
        data: Input data to process

    Returns:
        Processed data
    """
    logger.info("Starting data processing task")

    # Simulate data processing
    time.sleep(3)

    processed_data = {
        "input": data,
        "processed_at": time.time(),
        "result": "Data processed successfully",
    }

    logger.info("Data processing task completed")
    return processed_data
