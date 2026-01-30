"""
User service layer implementing business logic for user operations.
Separates business logic from API routes and database operations.
"""

from typing import Optional

from sqlmodel import Session, select

from app.core.security import get_password_hash, verify_password
from app.models.user import User, UserRole
from app.schemas.user import UserCreate


class UserService:
    """Service class for user-related operations."""

    @staticmethod
    def get_by_email(session: Session, email: str) -> Optional[User]:
        """
        Retrieve a user by email address.

        Args:
            session: Database session
            email: Email address to search for

        Returns:
            User if found, None otherwise
        """
        statement = select(User).where(User.email == email)
        return session.exec(statement).first()

    @staticmethod
    def get_by_id(session: Session, user_id: int) -> Optional[User]:
        """
        Retrieve a user by ID.

        Args:
            session: Database session
            user_id: User ID to search for

        Returns:
            User if found, None otherwise
        """
        return session.get(User, user_id)

    @staticmethod
    def create(session: Session, user_create: UserCreate, role: UserRole = UserRole.USER) -> User:
        """
        Create a new user with hashed password.

        Args:
            session: Database session
            user_create: User creation data
            role: User role (defaults to USER)

        Returns:
            Created user instance
        """
        db_user = User(
            email=user_create.email,
            hashed_password=get_password_hash(user_create.password),
            full_name=user_create.full_name,
            role=role,
        )
        session.add(db_user)
        session.commit()
        session.refresh(db_user)
        return db_user

    @staticmethod
    def authenticate(session: Session, email: str, password: str) -> Optional[User]:
        """
        Authenticate a user by email and password.

        Args:
            session: Database session
            email: User's email
            password: Plain text password

        Returns:
            User if authentication successful, None otherwise
        """
        user = UserService.get_by_email(session, email)
        if not user:
            return None
        if not verify_password(password, user.hashed_password):
            return None
        if not user.is_active:
            return None
        return user

    @staticmethod
    def is_admin(user: User) -> bool:
        """
        Check if a user has admin privileges.

        Args:
            user: User to check

        Returns:
            True if user is admin, False otherwise
        """
        return user.role == UserRole.ADMIN
