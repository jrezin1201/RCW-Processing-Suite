"""Lennar scheduled-tasks processor module."""
from fastapi import FastAPI

from app.modules.lennar.routes import router

MODULE_META = {
    "id": "lennar",
    "name": "Lennar Tasks",
    "description": "Parse and summarize Lennar scheduled-task exports",
    "route_prefix": "/api/v1",
}


def register(app: FastAPI) -> None:
    app.include_router(router)
