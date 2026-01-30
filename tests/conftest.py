"""
Pytest configuration and fixtures.
Provides test database, client, and common test utilities.
"""

from typing import Generator

import pytest
from fastapi.testclient import TestClient
from sqlmodel import Session, SQLModel, create_engine
from sqlmodel.pool import StaticPool

from app.core.config import settings
from app.db.session import get_session
from app.main import app
from app.models.user import User, UserRole
from app.schemas.user import UserCreate
from app.services.user_service import UserService


@pytest.fixture(name="session")
def session_fixture() -> Generator[Session, None, None]:
    """
    Create a test database session.
    Uses an in-memory SQLite database for fast tests.
    """
    engine = create_engine(
        "sqlite:///:memory:",
        connect_args={"check_same_thread": False},
        poolclass=StaticPool,
    )
    SQLModel.metadata.create_all(engine)

    with Session(engine) as session:
        yield session


@pytest.fixture(name="client")
def client_fixture(session: Session) -> Generator[TestClient, None, None]:
    """
    Create a test client with dependency overrides.
    """

    def get_session_override() -> Generator[Session, None, None]:
        yield session

    app.dependency_overrides[get_session] = get_session_override

    with TestClient(app) as test_client:
        yield test_client

    app.dependency_overrides.clear()


@pytest.fixture(name="test_user")
def test_user_fixture(session: Session) -> User:
    """
    Create a test user.
    """
    user_create = UserCreate(
        email="test@example.com",
        password="testpassword123",
        full_name="Test User",
    )
    return UserService.create(session, user_create)


@pytest.fixture(name="test_admin")
def test_admin_fixture(session: Session) -> User:
    """
    Create a test admin user.
    """
    user_create = UserCreate(
        email="admin@example.com",
        password="adminpassword123",
        full_name="Admin User",
    )
    return UserService.create(session, user_create, role=UserRole.ADMIN)


@pytest.fixture(name="user_token")
def user_token_fixture(client: TestClient, test_user: User) -> str:
    """
    Get an access token for a regular user.
    """
    response = client.post(
        f"{settings.API_V1_PREFIX}/auth/login",
        data={
            "username": "test@example.com",
            "password": "testpassword123",
        },
    )
    assert response.status_code == 200
    return response.json()["access_token"]


@pytest.fixture(name="admin_token")
def admin_token_fixture(client: TestClient, test_admin: User) -> str:
    """
    Get an access token for an admin user.
    """
    response = client.post(
        f"{settings.API_V1_PREFIX}/auth/login",
        data={
            "username": "admin@example.com",
            "password": "adminpassword123",
        },
    )
    assert response.status_code == 200
    return response.json()["access_token"]
