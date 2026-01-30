"""
User model with role-based access control.
Implements a simple admin/user role system.
"""

from datetime import datetime
from enum import Enum
from typing import Optional

from sqlmodel import Field, SQLModel


class UserRole(str, Enum):
    """User role enumeration for RBAC."""

    ADMIN = "admin"
    USER = "user"


class User(SQLModel, table=True):
    """
    User model with authentication and role support.

    Attributes:
        id: Primary key
        email: Unique email address (used for login)
        hashed_password: Bcrypt hashed password
        full_name: Optional user's full name
        role: User role (admin or user)
        is_active: Whether the account is active
        created_at: Timestamp of account creation
        updated_at: Timestamp of last update
    """

    __tablename__ = "users"  # type: ignore

    id: Optional[int] = Field(default=None, primary_key=True)
    email: str = Field(unique=True, index=True, max_length=255)
    hashed_password: str
    full_name: Optional[str] = Field(default=None, max_length=255)
    role: UserRole = Field(default=UserRole.USER)
    is_active: bool = Field(default=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)
