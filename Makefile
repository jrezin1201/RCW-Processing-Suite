.PHONY: help install dev test lint format typecheck clean docker-build docker-up docker-down docker-logs

help:
	@echo "RCW Processing Suite - Available Commands"
	@echo ""
	@echo "Setup:"
	@echo "  make install       Install dependencies with Poetry"
	@echo ""
	@echo "Development:"
	@echo "  make dev           Start development server (reload on change)"
	@echo "  make docker-build  Build the production Docker image"
	@echo "  make docker-up     Start via docker-compose"
	@echo "  make docker-down   Stop docker-compose stack"
	@echo "  make docker-logs   Tail API container logs"
	@echo ""
	@echo "Quality:"
	@echo "  make test          Run tests"
	@echo "  make test-cov      Run tests with coverage"
	@echo "  make lint          Run ruff linter"
	@echo "  make format        Format with black + ruff fix"
	@echo "  make typecheck     Run mypy"
	@echo "  make clean         Remove caches and build artifacts"

install:
	poetry install

dev:
	uvicorn app.main:app --reload --host 0.0.0.0 --port 8000

docker-build:
	docker compose build

docker-up:
	docker compose up -d
	@echo ""
	@echo "Services started:"
	@echo "  API: http://localhost:8000"

docker-down:
	docker compose down

docker-logs:
	docker compose logs -f api

test:
	pytest

test-cov:
	pytest --cov=app --cov-report=html --cov-report=term

lint:
	ruff check .

format:
	black .
	ruff check --fix .

typecheck:
	mypy app

clean:
	find . -type d -name "__pycache__" -not -path "./venv/*" -exec rm -rf {} +
	find . -type f -name "*.pyc" -not -path "./venv/*" -delete
	find . -type f -name "*.pyo" -not -path "./venv/*" -delete
	find . -type d -name ".pytest_cache" -not -path "./venv/*" -exec rm -rf {} +
	find . -type d -name ".ruff_cache" -not -path "./venv/*" -exec rm -rf {} +
	rm -rf htmlcov/ .coverage
