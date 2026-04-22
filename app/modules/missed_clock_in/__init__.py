"""Missed Clock-In warning-notice generator module."""
from fastapi import FastAPI

from app.modules.missed_clock_in.routes import router

MODULE_META = {
    "id": "missed_clock_in",
    "name": "Missed Clock-In",
    "description": "Generate warning notices for timekeeping violations",
    "route_prefix": "/api/v1/missed-clock-in",
}


def register(app: FastAPI) -> None:
    app.include_router(router)
