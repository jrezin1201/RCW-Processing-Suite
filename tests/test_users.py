"""
Tests for user endpoints.
"""

from fastapi.testclient import TestClient

from app.core.config import settings
from app.models.user import User


def test_get_current_user(client: TestClient, user_token: str, test_user: User) -> None:
    """Test getting current user profile."""
    response = client.get(
        f"{settings.API_V1_PREFIX}/users/me",
        headers={"Authorization": f"Bearer {user_token}"},
    )
    assert response.status_code == 200
    data = response.json()
    assert data["email"] == test_user.email
    assert data["id"] == test_user.id


def test_get_current_user_unauthorized(client: TestClient) -> None:
    """Test that accessing profile without token fails."""
    response = client.get(f"{settings.API_V1_PREFIX}/users/me")
    assert response.status_code == 401


def test_admin_only_route_as_admin(client: TestClient, admin_token: str) -> None:
    """Test admin-only route with admin user."""
    response = client.get(
        f"{settings.API_V1_PREFIX}/users/admin-only",
        headers={"Authorization": f"Bearer {admin_token}"},
    )
    assert response.status_code == 200
    data = response.json()
    assert "admin-only" in data["message"].lower()


def test_admin_only_route_as_regular_user(client: TestClient, user_token: str) -> None:
    """Test that regular users cannot access admin-only routes."""
    response = client.get(
        f"{settings.API_V1_PREFIX}/users/admin-only",
        headers={"Authorization": f"Bearer {user_token}"},
    )
    assert response.status_code == 403
