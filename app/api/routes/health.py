"""
Health check routes for monitoring and service discovery.
Provides endpoints to verify service health and database connectivity.
"""

from fastapi import APIRouter, Depends
from sqlmodel import Session, text

from app.core.config import settings
from app.db.session import get_session

router = APIRouter(tags=["health"])


@router.get("/health")
def health_check() -> dict:
    """
    Basic health check endpoint.
    Returns service status and version information.

    Returns:
        Health status
    """
    return {
        "status": "healthy",
        "service": settings.PROJECT_NAME,
        "version": settings.VERSION,
    }


@router.get("/health/db")
def database_health_check(session: Session = Depends(get_session)) -> dict:
    """
    Database health check endpoint.
    Verifies database connectivity by executing a simple query.

    Args:
        session: Database session

    Returns:
        Database health status
    """
    try:
        # Execute a simple query to verify database connectivity
        result = session.exec(text("SELECT 1")).first()
        return {
            "status": "healthy",
            "database": "connected",
            "query_result": result,
        }
    except Exception as e:
        return {
            "status": "unhealthy",
            "database": "disconnected",
            "error": str(e),
        }
