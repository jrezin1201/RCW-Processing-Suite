"""
Database session management using SQLModel.
Provides session factory and dependency injection for FastAPI routes.
"""

from typing import Generator

from sqlmodel import Session, create_engine

from app.core.config import settings

# Create database engine with connection pooling
# pool_pre_ping ensures connections are alive before using them
engine = create_engine(
    settings.SQLALCHEMY_DATABASE_URI,
    echo=settings.DEBUG,
    pool_pre_ping=True,
    pool_size=5,
    max_overflow=10,
)


def get_session() -> Generator[Session, None, None]:
    """
    Dependency that provides a database session for FastAPI routes.
    Automatically commits on success and rolls back on exceptions.

    Yields:
        Database session instance
    """
    with Session(engine) as session:
        yield session
