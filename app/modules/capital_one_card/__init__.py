"""Capital One Card Report processor module."""
from fastapi import FastAPI

from app.modules.capital_one_card.routes import router

MODULE_META = {
    "id": "capital_one_card",
    "name": "Capital One Card Report",
    "description": (
        "Groups Capital One card transactions by Category, then Description, "
        "while keeping each transaction separate for job costing"
    ),
    "route_prefix": "/api/v1/capital-one-card",
}


def register(app: FastAPI) -> None:
    app.include_router(router)
