"""
User schemas for API request/response validation.
Separates internal models from API contracts using Pydantic.
"""

from datetime import datetime
from typing import Optional

from pydantic import BaseModel, EmailStr

from app.models.user import UserRole


class UserBase(BaseModel):
    """Base user schema with common fields."""

    email: EmailStr
    full_name: Optional[str] = None


class UserCreate(UserBase):
    """Schema for user registration."""

    password: str


class UserLogin(BaseModel):
    """Schema for user login."""

    email: EmailStr
    password: str


class UserResponse(UserBase):
    """
    Schema for user data in API responses.
    Excludes sensitive information like hashed_password.
    """

    id: int
    role: UserRole
    is_active: bool
    created_at: datetime
    updated_at: datetime

    model_config = {"from_attributes": True}


class User(UserResponse):
    """Full user schema for internal use."""

    pass
