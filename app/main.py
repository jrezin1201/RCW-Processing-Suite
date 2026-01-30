"""Main FastAPI application for Lennar Excel processor service."""
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse

from app.api.lennar_routes import router


# Create FastAPI app with documentation
app = FastAPI(
    title="RCW Processing Tools",
    description="Service for processing Excel exports - Hours Worked & Lennar Tasks",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json"
)

# Configure CORS for development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allow all origins in development
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static files
app.mount("/static", StaticFiles(directory="app/static"), name="static")

# Setup templates
templates = Jinja2Templates(directory="app/templates")

# Include API routes
app.include_router(router)


@app.get("/", response_class=HTMLResponse, tags=["Root"])
async def root(request: Request):
    """Root endpoint - serves the professional interface."""
    return templates.TemplateResponse("professional_interface.html", {"request": request})


@app.get("/api/info", tags=["API"])
async def api_info():
    """API information endpoint."""
    return {
        "service": "RCW Processing Tools",
        "version": "1.0.0",
        "docs": "/docs",
        "health": "/health",
        "interface": "/"
    }


@app.get("/health", tags=["Health"])
async def health_check():
    """Health check endpoint."""
    return {"status": "healthy", "service": "rcw-processing-tools"}