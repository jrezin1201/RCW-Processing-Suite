"""
Authentication routes for user registration and login.
Provides JWT token-based authentication.
"""

from typing import Annotated

from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.security import OAuth2PasswordRequestForm
from sqlmodel import Session

from app.core.logging import get_logger
from app.core.security import create_access_token
from app.db.session import get_session
from app.schemas.token import Token
from app.schemas.user import UserCreate, UserResponse
from app.services.user_service import UserService

logger = get_logger(__name__)

router = APIRouter(prefix="/auth", tags=["auth"])


@router.post("/register", response_model=UserResponse, status_code=status.HTTP_201_CREATED)
def register(
    user_in: UserCreate,
    session: Annotated[Session, Depends(get_session)],
) -> UserResponse:
    """
    Register a new user.

    Args:
        user_in: User registration data
        session: Database session

    Returns:
        Created user data

    Raises:
        HTTPException: If email already registered
    """
    # Check if user already exists
    existing_user = UserService.get_by_email(session, email=user_in.email)
    if existing_user:
        logger.warning(f"Registration attempt with existing email: {user_in.email}")
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Email already registered",
        )

    # Create new user
    user = UserService.create(session, user_create=user_in)
    logger.info(f"New user registered: {user.email} (ID: {user.id})")

    return UserResponse.model_validate(user)


@router.post("/login", response_model=Token)
def login(
    session: Annotated[Session, Depends(get_session)],
    form_data: Annotated[OAuth2PasswordRequestForm, Depends()],
) -> Token:
    """
    OAuth2 compatible token login.

    Args:
        session: Database session
        form_data: OAuth2 form with username (email) and password

    Returns:
        Access token

    Raises:
        HTTPException: If credentials are invalid
    """
    user = UserService.authenticate(
        session, email=form_data.username, password=form_data.password
    )
    if not user:
        logger.warning(f"Failed login attempt for email: {form_data.username}")
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect email or password",
        )

    # Create access token
    access_token = create_access_token(subject=user.id)
    logger.info(f"User logged in: {user.email} (ID: {user.id})")

    return Token(access_token=access_token)
