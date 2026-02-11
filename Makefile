.PHONY: help install dev test lint format clean docker-up docker-down migrate

help:
	@echo "FastAPI SaaS Starter - Available Commands"
	@echo ""
	@echo "Setup:"
	@echo "  make install      Install dependencies with Poetry"
	@echo "  make setup        Copy .env.example to .env"
	@echo ""
	@echo "Development:"
	@echo "  make dev          Start development server"
	@echo "  make docker-up    Start all services with Docker"
	@echo "  make docker-down  Stop all Docker services"
	@echo ""
	@echo "Database:"
	@echo "  make migrate      Run database migrations"
	@echo "  make migration    Create new migration (use MSG='message')"
	@echo ""
	@echo "Testing:"
	@echo "  make test         Run tests"
	@echo "  make test-cov     Run tests with coverage"
	@echo ""
	@echo "Code Quality:"
	@echo "  make lint         Run linter"
	@echo "  make format       Format code"
	@echo "  make typecheck    Run type checker"
	@echo ""
	@echo "Cleanup:"
	@echo "  make clean        Remove cache and build files"

install:
	poetry install

setup:
	cp .env.example .env
	@echo "Created .env file. Please update with your configuration."

dev:
	uvicorn app.main:app --reload --host 0.0.0.0 --port 8000

docker-up:
	docker-compose up -d
	@echo ""
	@echo "Services started:"
	@echo "  - API: http://localhost:8000"
	@echo "  - Docs: http://localhost:8000/api/v1/docs"

docker-down:
	docker-compose down

docker-logs:
	docker-compose logs -f

migrate:
	alembic upgrade head

migration:
	@if [ -z "$(MSG)" ]; then \
		echo "Error: MSG is required. Usage: make migration MSG='your message'"; \
		exit 1; \
	fi
	alembic revision --autogenerate -m "$(MSG)"

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
	find . -type d -name "__pycache__" -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete
	find . -type f -name "*.pyo" -delete
	find . -type d -name "*.egg-info" -exec rm -rf {} +
	find . -type d -name ".pytest_cache" -exec rm -rf {} +
	find . -type d -name ".ruff_cache" -exec rm -rf {} +
	rm -rf htmlcov/
	rm -rf .coverage
