"""
User routes for user profile and management operations.
Demonstrates protected routes and role-based access control.
"""

from typing import Annotated

from fastapi import APIRouter, Depends

from app.api.deps import get_current_active_user, get_current_admin_user
from app.models.user import User
from app.schemas.user import UserResponse

router = APIRouter(prefix="/users", tags=["users"])


@router.get("/me", response_model=UserResponse)
def get_current_user_profile(
    current_user: Annotated[User, Depends(get_current_active_user)],
) -> UserResponse:
    """
    Get current user's profile.
    This is a protected route that requires authentication.

    Args:
        current_user: Current authenticated user

    Returns:
        User profile data
    """
    return UserResponse.model_validate(current_user)


@router.get("/admin-only", response_model=dict)
def admin_only_route(
    current_user: Annotated[User, Depends(get_current_admin_user)],
) -> dict:
    """
    Example admin-only route demonstrating role-based access control.
    Only users with admin role can access this endpoint.

    Args:
        current_user: Current authenticated admin user

    Returns:
        Success message
    """
    return {
        "message": "This is an admin-only endpoint",
        "admin_user": current_user.email,
    }
