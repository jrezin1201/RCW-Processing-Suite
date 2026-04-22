"""Gas & Rig job-cost processor module."""
from fastapi import FastAPI

from app.modules.gas_rig.routes import router

MODULE_META = {
    "id": "gasrig",
    "name": "Gas & Rig",
    "description": "Compute job costs from hours-worked exports",
    "route_prefix": "/api/v1/gas-rig",
}


def register(app: FastAPI) -> None:
    app.include_router(router)
