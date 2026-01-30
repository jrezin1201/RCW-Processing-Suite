# FastAPI SaaS Starter - Quick Start Guide

## 🎉 What Was Created

A complete, production-ready FastAPI template with:

### Core Features
✅ FastAPI with Pydantic v2
✅ SQLModel ORM + PostgreSQL
✅ Alembic migrations
✅ Redis + RQ background jobs
✅ JWT authentication
✅ Role-based access control (admin/user)
✅ Structured JSON logging
✅ Comprehensive test suite (pytest)
✅ Docker + docker-compose setup

### Project Structure (33 Python files)
```
app/
├── api/routes/     # Auth, Users, Jobs, Health endpoints
├── core/           # Config, Security, Logging
├── db/             # Database session
├── models/         # User model with RBAC
├── schemas/        # Pydantic request/response schemas
├── services/       # Business logic layer
└── workers/        # Background tasks with RQ

tests/              # 5 test modules (22 tests)
alembic/            # Database migrations
scripts/            # Helper scripts
```

## 🚀 How to Run

### Option 1: Docker (Fastest - Recommended)

```bash
# Start all services
docker-compose up -d

# View logs
docker-compose logs -f api

# Stop
docker-compose down
```

**Services will be available at:**
- API: http://localhost:8000
- Docs: http://localhost:8000/api/v1/docs
- ReDoc: http://localhost:8000/api/v1/redoc

### Option 2: Using Makefile

```bash
make help           # See all commands
make docker-up      # Start with Docker
make docker-logs    # View logs
make test           # Run tests (needs databases)
make format         # Format code
```

### Option 3: Local Development

```bash
# 1. Install Poetry
pip install --user poetry

# 2. Install dependencies
poetry install

# 3. Copy environment file
cp .env.example .env

# 4. Start databases (Docker)
docker-compose up -d postgres redis

# 5. Run migrations
poetry run alembic upgrade head

# 6. Start API server
poetry run uvicorn app.main:app --reload

# 7. Start worker (separate terminal)
poetry run rq worker default --with-scheduler
```

## 🔑 Default Credentials

```
Email: admin@example.com
Password: changethis123
```

⚠️ **Change these in production!**

## 📝 Example API Calls

### Register a User
```bash
curl -X POST "http://localhost:8000/api/v1/auth/register" \
  -H "Content-Type: application/json" \
  -d '{
    "email": "user@example.com",
    "password": "password123",
    "full_name": "John Doe"
  }'
```

### Login
```bash
curl -X POST "http://localhost:8000/api/v1/auth/login" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  -d "username=user@example.com&password=password123"
```

### Get Current User (Protected Route)
```bash
TOKEN="your-token-from-login"

curl -X GET "http://localhost:8000/api/v1/users/me" \
  -H "Authorization: Bearer $TOKEN"
```

### Create Background Job
```bash
curl -X POST "http://localhost:8000/api/v1/jobs/example?name=test&duration=5" \
  -H "Authorization: Bearer $TOKEN"
```

### Check Job Status
```bash
JOB_ID="job-id-from-previous-response"

curl -X GET "http://localhost:8000/api/v1/jobs/$JOB_ID/status" \
  -H "Authorization: Bearer $TOKEN"
```

## 🧪 Testing

The test suite includes 22 tests covering:
- Authentication (registration, login)
- Protected routes
- Role-based access control
- Background jobs
- Health checks
- Service layer

**Note:** Tests require PostgreSQL and Redis running. Use Docker for easiest setup:

```bash
# Start databases
docker-compose up -d postgres redis

# Run tests
poetry run pytest -v

# With coverage
poetry run pytest --cov=app --cov-report=html
```

## 🏗️ Architecture Highlights

### Clean Layer Separation
1. **Routes** - Handle HTTP requests/responses
2. **Services** - Business logic (reusable)
3. **Models** - Database schema (SQLModel)
4. **Schemas** - API contracts (Pydantic)

### Security
- JWT tokens for authentication
- Bcrypt password hashing
- Role-based access control
- Protected route dependencies

### Background Jobs
- RQ (Redis Queue) for async tasks
- Independent worker process
- Job status tracking
- Examples: email, data processing

### Production Ready
- Type hints throughout
- Structured JSON logging
- Health check endpoints
- Database connection pooling
- Docker multi-stage builds

## 📚 Next Steps

1. **Customize Environment**
   - Edit `.env` with your settings
   - Update `SECRET_KEY` for production

2. **Add Your Models**
   - Create models in `app/models/`
   - Generate migrations: `make migration MSG="add your_table"`
   - Apply: `make migrate`

3. **Add Your Routes**
   - Create routes in `app/api/routes/`
   - Create services in `app/services/`
   - Register in `app/main.py`

4. **Add Background Tasks**
   - Define in `app/workers/tasks.py`
   - Enqueue via `enqueue_task()`

5. **Deploy**
   - Build: `docker build -t your-app .`
   - Push to registry
   - Deploy to your platform

## 📖 Full Documentation

See `README.md` for complete documentation including:
- Detailed architecture explanations
- How to extend the template
- Production deployment guide
- Troubleshooting tips

## 🤝 Contributing

See `CONTRIBUTING.md` for:
- Development setup
- Code standards
- Testing guidelines
- Pull request process

## ✨ What Makes This Template Great

- **Type-Safe** - Type hints everywhere
- **Tested** - Comprehensive test coverage
- **Documented** - Clear code comments
- **Modular** - Clean separation of concerns
- **Scalable** - Background jobs, async-ready
- **Production-Ready** - Logging, health checks, Docker
- **Extensible** - Easy to add features

Ready to build your SaaS! 🚀
