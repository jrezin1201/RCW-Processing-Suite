"""
Tests for service layer.
"""

from sqlmodel import Session

from app.models.user import User, UserRole
from app.schemas.user import UserCreate
from app.services.user_service import UserService


def test_create_user(session: Session) -> None:
    """Test user creation."""
    user_create = UserCreate(
        email="service@example.com",
        password="password123",
        full_name="Service Test User",
    )
    user = UserService.create(session, user_create)

    assert user.id is not None
    assert user.email == "service@example.com"
    assert user.full_name == "Service Test User"
    assert user.role == UserRole.USER
    assert user.is_active is True


def test_get_user_by_email(session: Session, test_user: User) -> None:
    """Test retrieving user by email."""
    user = UserService.get_by_email(session, test_user.email)
    assert user is not None
    assert user.id == test_user.id
    assert user.email == test_user.email


def test_get_user_by_id(session: Session, test_user: User) -> None:
    """Test retrieving user by ID."""
    user = UserService.get_by_id(session, test_user.id)  # type: ignore
    assert user is not None
    assert user.id == test_user.id
    assert user.email == test_user.email


def test_authenticate_user_success(session: Session, test_user: User) -> None:
    """Test successful user authentication."""
    user = UserService.authenticate(session, "test@example.com", "testpassword123")
    assert user is not None
    assert user.id == test_user.id


def test_authenticate_user_wrong_password(session: Session, test_user: User) -> None:
    """Test authentication with wrong password."""
    user = UserService.authenticate(session, "test@example.com", "wrongpassword")
    assert user is None


def test_authenticate_nonexistent_user(session: Session) -> None:
    """Test authentication with non-existent user."""
    user = UserService.authenticate(session, "nobody@example.com", "password")
    assert user is None


def test_is_admin(session: Session, test_admin: User, test_user: User) -> None:
    """Test admin check."""
    assert UserService.is_admin(test_admin) is True
    assert UserService.is_admin(test_user) is False
