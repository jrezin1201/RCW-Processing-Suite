"""
Tests for authentication endpoints.
"""

from fastapi.testclient import TestClient
from sqlmodel import Session

from app.core.config import settings
from app.models.user import User


def test_register_user(client: TestClient) -> None:
    """Test user registration."""
    response = client.post(
        f"{settings.API_V1_PREFIX}/auth/register",
        json={
            "email": "newuser@example.com",
            "password": "newpassword123",
            "full_name": "New User",
        },
    )
    assert response.status_code == 201
    data = response.json()
    assert data["email"] == "newuser@example.com"
    assert data["full_name"] == "New User"
    assert "id" in data
    assert "hashed_password" not in data


def test_register_duplicate_email(client: TestClient, test_user: User) -> None:
    """Test that duplicate email registration fails."""
    response = client.post(
        f"{settings.API_V1_PREFIX}/auth/register",
        json={
            "email": test_user.email,
            "password": "password123",
            "full_name": "Duplicate User",
        },
    )
    assert response.status_code == 400
    assert "already registered" in response.json()["detail"].lower()


def test_login_success(client: TestClient, test_user: User) -> None:
    """Test successful login."""
    response = client.post(
        f"{settings.API_V1_PREFIX}/auth/login",
        data={
            "username": "test@example.com",
            "password": "testpassword123",
        },
    )
    assert response.status_code == 200
    data = response.json()
    assert "access_token" in data
    assert data["token_type"] == "bearer"


def test_login_wrong_password(client: TestClient, test_user: User) -> None:
    """Test login with wrong password."""
    response = client.post(
        f"{settings.API_V1_PREFIX}/auth/login",
        data={
            "username": "test@example.com",
            "password": "wrongpassword",
        },
    )
    assert response.status_code == 401


def test_login_nonexistent_user(client: TestClient) -> None:
    """Test login with non-existent user."""
    response = client.post(
        f"{settings.API_V1_PREFIX}/auth/login",
        data={
            "username": "nonexistent@example.com",
            "password": "password123",
        },
    )
    assert response.status_code == 401
