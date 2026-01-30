"""
Tests for health check endpoints.
"""

from fastapi.testclient import TestClient

from app.core.config import settings


def test_health_check(client: TestClient) -> None:
    """Test basic health check."""
    response = client.get(f"{settings.API_V1_PREFIX}/health")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "healthy"
    assert data["service"] == settings.PROJECT_NAME
    assert data["version"] == settings.VERSION


def test_database_health_check(client: TestClient) -> None:
    """Test database health check."""
    response = client.get(f"{settings.API_V1_PREFIX}/health/db")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "healthy"
    assert data["database"] == "connected"
