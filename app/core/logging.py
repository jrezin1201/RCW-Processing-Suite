"""
Structured logging configuration using python-json-logger.
Provides consistent, machine-readable logs for production environments.
"""

import logging
import sys
from typing import Any

from pythonjsonlogger import jsonlogger

from app.core.config import settings


class CustomJsonFormatter(jsonlogger.JsonFormatter):
    """Custom JSON formatter that adds application-specific fields."""

    def add_fields(
        self,
        log_record: dict[str, Any],
        record: logging.LogRecord,
        message_dict: dict[str, Any],
    ) -> None:
        """Add custom fields to each log record."""
        super().add_fields(log_record, record, message_dict)
        log_record["service"] = settings.PROJECT_NAME
        log_record["version"] = settings.VERSION
        log_record["level"] = record.levelname


def setup_logging() -> None:
    """
    Configure application-wide logging.
    Uses JSON format in production, simpler format in development.
    """
    log_level = logging.DEBUG if settings.DEBUG else logging.INFO

    # Create handler
    handler = logging.StreamHandler(sys.stdout)

    if settings.DEBUG:
        # Simple format for development
        formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    else:
        # JSON format for production
        formatter = CustomJsonFormatter(
            "%(timestamp)s %(level)s %(name)s %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )

    handler.setFormatter(formatter)

    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)
    root_logger.addHandler(handler)

    # Reduce noise from third-party libraries
    logging.getLogger("uvicorn.access").setLevel(logging.WARNING)


def get_logger(name: str) -> logging.Logger:
    """
    Get a logger instance for a specific module.

    Args:
        name: Name of the module (typically __name__)

    Returns:
        Configured logger instance
    """
    return logging.getLogger(name)
