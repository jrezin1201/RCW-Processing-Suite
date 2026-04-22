"""Main FastAPI application for the RCW Processing Suite.

Modules are auto-discovered from ``app.modules``. To add a new module, drop a
new subpackage under ``app/modules/`` that exposes a ``register(app)`` function —
no changes here are required.
"""
import logging

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.core.config import settings
from app.core.registry import load_modules

logging.basicConfig(
    level=logging.DEBUG if settings.DEBUG else logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

# Disable the OpenAPI / Swagger / ReDoc endpoints unless explicitly enabled —
# they expose the full API surface and are unnecessary attack surface in prod.
_docs_url = "/docs" if settings.ENABLE_DOCS else None
_redoc_url = "/redoc" if settings.ENABLE_DOCS else None
_openapi_url = "/openapi.json" if settings.ENABLE_DOCS else None

app = FastAPI(
    title=settings.PROJECT_NAME,
    description="Excel processing suite with pluggable modules",
    version=settings.VERSION,
    docs_url=_docs_url,
    redoc_url=_redoc_url,
    openapi_url=_openapi_url,
)

# CORS: production default is deny-all. Set CORS_ORIGINS in the environment to
# allowlist specific origins (comma-separated).
if settings.CORS_ORIGINS:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.CORS_ORIGINS,
        allow_credentials=True,
        allow_methods=["GET", "POST"],
        allow_headers=["*"],
    )
    logger.info("CORS enabled for origins: %s", settings.CORS_ORIGINS)
else:
    logger.info("CORS disabled — same-origin only")

app.mount("/static", StaticFiles(directory="app/static"), name="static")
templates = Jinja2Templates(directory="app/templates")

# Auto-discover and register every module under app.modules.
registered_modules = load_modules(app)
logger.info(
    "Started %s %s [env=%s] with modules: %s",
    settings.PROJECT_NAME,
    settings.VERSION,
    settings.ENVIRONMENT,
    [m.get("id") for m in registered_modules],
)


@app.get("/", response_class=HTMLResponse, tags=["Root"])
async def root(request: Request):
    """Serve the professional interface."""
    return templates.TemplateResponse(
        "professional_interface.html",
        {"request": request, "modules": registered_modules},
    )


@app.get("/api/info", tags=["API"])
async def api_info():
    """API information endpoint (public)."""
    return {
        "service": settings.PROJECT_NAME,
        "version": settings.VERSION,
        "environment": settings.ENVIRONMENT,
        "modules": [m.get("id") for m in registered_modules],
    }


@app.get("/api/modules", tags=["API"])
async def list_modules():
    """List registered modules (ids + route prefixes)."""
    return {"modules": registered_modules}


@app.get("/health", tags=["Health"])
async def health_check():
    """Health check endpoint (public, cheap)."""
    return {"status": "healthy"}
