#!/bin/bash
# Development server startup script

echo "Starting FastAPI development server..."
echo ""
echo "Services will be available at:"
echo "  - API: http://localhost:8000"
echo "  - API Docs: http://localhost:8000/api/v1/docs"
echo "  - ReDoc: http://localhost:8000/api/v1/redoc"
echo ""

uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
