"""Merchant Charges credit-card transaction processor module."""
from fastapi import FastAPI

from app.modules.merchant_charges.routes import router

MODULE_META = {
    "id": "merchant_charges",
    "name": "Merchant Charges",
    "description": "Group credit card transactions by merchant with subtotals and grand total",
    "route_prefix": "/api/v1/merchant-charges",
}


def register(app: FastAPI) -> None:
    app.include_router(router)
