"""
Tests for background job endpoints.
"""

from unittest.mock import MagicMock, patch

from fastapi.testclient import TestClient

from app.core.config import settings


def test_create_example_job(client: TestClient, user_token: str) -> None:
    """Test creating an example background job."""
    with patch("app.api.routes.jobs.enqueue_task") as mock_enqueue:
        mock_enqueue.return_value = "test-job-id-123"

        response = client.post(
            f"{settings.API_V1_PREFIX}/jobs/example",
            params={"name": "test-task", "duration": 3},
            headers={"Authorization": f"Bearer {user_token}"},
        )

        assert response.status_code == 200
        data = response.json()
        assert data["job_id"] == "test-job-id-123"
        assert data["task"] == "example_task"
        mock_enqueue.assert_called_once()


def test_create_email_job(client: TestClient, user_token: str) -> None:
    """Test creating an email sending job."""
    with patch("app.api.routes.jobs.enqueue_task") as mock_enqueue:
        mock_enqueue.return_value = "email-job-id-456"

        response = client.post(
            f"{settings.API_V1_PREFIX}/jobs/send-email",
            params={
                "to": "recipient@example.com",
                "subject": "Test Email",
                "body": "Test body",
            },
            headers={"Authorization": f"Bearer {user_token}"},
        )

        assert response.status_code == 200
        data = response.json()
        assert data["job_id"] == "email-job-id-456"
        assert data["task"] == "send_email_task"


def test_get_job_status(client: TestClient, user_token: str) -> None:
    """Test getting job status."""
    with patch("app.api.routes.jobs.get_job_status") as mock_get_status:
        mock_get_status.return_value = {
            "job_id": "test-job-123",
            "status": "finished",
            "result": {"message": "Task completed"},
        }

        response = client.get(
            f"{settings.API_V1_PREFIX}/jobs/test-job-123/status",
            headers={"Authorization": f"Bearer {user_token}"},
        )

        assert response.status_code == 200
        data = response.json()
        assert data["job_id"] == "test-job-123"
        assert data["status"] == "finished"


def test_job_unauthorized(client: TestClient) -> None:
    """Test that job endpoints require authentication."""
    response = client.post(
        f"{settings.API_V1_PREFIX}/jobs/example",
        params={"name": "test", "duration": 1},
    )
    assert response.status_code == 401
