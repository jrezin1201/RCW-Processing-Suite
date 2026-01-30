# Contributing Guide

Thank you for considering contributing to the FastAPI SaaS Starter template!

## Development Setup

1. **Fork and clone the repository**
   ```bash
   git clone https://github.com/your-username/fastapi-saas-starter.git
   cd fastapi-saas-starter
   ```

2. **Install dependencies**
   ```bash
   make install
   # OR
   poetry install
   ```

3. **Setup environment**
   ```bash
   make setup
   # OR
   cp .env.example .env
   ```

4. **Start services**
   ```bash
   make docker-up
   ```

## Code Standards

### Python Style Guide

- Follow **PEP 8** style guide
- Use **type hints** for all function signatures
- Maximum line length: **100 characters**
- Use **black** for code formatting
- Use **ruff** for linting

### Code Formatting

Before submitting, format your code:

```bash
make format
```

### Type Checking

Ensure type safety:

```bash
make typecheck
```

### Linting

Check for code issues:

```bash
make lint
```

## Testing

### Writing Tests

- Place tests in the `tests/` directory
- Name test files as `test_*.py`
- Use descriptive test function names: `test_<what>_<condition>_<expected_result>`
- Use fixtures from `tests/conftest.py`

Example:

```python
def test_user_registration_with_valid_data_creates_user(client: TestClient) -> None:
    """Test that user registration with valid data creates a new user."""
    response = client.post(
        f"{settings.API_V1_PREFIX}/auth/register",
        json={
            "email": "new@example.com",
            "password": "password123",
            "full_name": "New User",
        },
    )
    assert response.status_code == 201
    assert response.json()["email"] == "new@example.com"
```

### Running Tests

```bash
# Run all tests
make test

# Run with coverage
make test-cov

# Run specific test file
pytest tests/test_auth.py -v

# Run specific test
pytest tests/test_auth.py::test_register_user -v
```

### Test Coverage

- Aim for **>80% coverage**
- Critical paths should have **100% coverage**
- All new features must include tests

## Architecture Guidelines

### Layer Separation

1. **Routes** (`app/api/routes/`)
   - Handle HTTP concerns only
   - Validate input via Pydantic schemas
   - Call service layer for business logic
   - Return appropriate HTTP responses

2. **Services** (`app/services/`)
   - Contain all business logic
   - Orchestrate database operations
   - Can be called from multiple routes
   - Should be framework-agnostic

3. **Models** (`app/models/`)
   - Define database schema
   - Use SQLModel for ORM
   - Include only database-related logic

4. **Schemas** (`app/schemas/`)
   - Define API contracts
   - Validate request/response data
   - Separate from database models

### Naming Conventions

- **Files**: snake_case (e.g., `user_service.py`)
- **Classes**: PascalCase (e.g., `UserService`)
- **Functions**: snake_case (e.g., `get_user_by_email`)
- **Constants**: UPPER_SNAKE_CASE (e.g., `MAX_LOGIN_ATTEMPTS`)
- **Private**: prefix with underscore (e.g., `_internal_helper`)

### Documentation

- Add **docstrings** to all public functions and classes
- Use **Google-style** docstrings
- Include type information in docstrings
- Document complex logic with inline comments

Example:

```python
def create_user(session: Session, user_create: UserCreate) -> User:
    """
    Create a new user in the database.

    Args:
        session: Database session
        user_create: User creation data

    Returns:
        Created user instance

    Raises:
        ValueError: If email already exists
    """
    # Implementation
```

## Adding New Features

### 1. Database Models

```bash
# Create model in app/models/
# Create migration
make migration MSG="add your_model table"

# Apply migration
make migrate
```

### 2. API Endpoints

1. Create Pydantic schemas in `app/schemas/`
2. Create service in `app/services/`
3. Create route in `app/api/routes/`
4. Register router in `app/main.py`
5. Write tests in `tests/`

### 3. Background Tasks

1. Define task function in `app/workers/tasks.py`
2. Create endpoint to enqueue task in `app/api/routes/jobs.py`
3. Test task execution

## Commit Messages

Follow **Conventional Commits** specification:

```
<type>(<scope>): <subject>

<body>

<footer>
```

Types:
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `style`: Code style changes (formatting)
- `refactor`: Code refactoring
- `test`: Adding or updating tests
- `chore`: Maintenance tasks

Examples:

```
feat(auth): add password reset functionality

Implements password reset via email with time-limited tokens.
Includes new endpoints for request and confirmation.

Closes #123
```

```
fix(jobs): prevent duplicate job enqueuing

Adds check to prevent the same job from being enqueued
multiple times within a 5-minute window.
```

## Pull Request Process

1. **Create a feature branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make your changes**
   - Write code
   - Add tests
   - Update documentation

3. **Run quality checks**
   ```bash
   make format
   make lint
   make typecheck
   make test
   ```

4. **Commit your changes**
   ```bash
   git add .
   git commit -m "feat: add your feature"
   ```

5. **Push to your fork**
   ```bash
   git push origin feature/your-feature-name
   ```

6. **Create pull request**
   - Provide clear description
   - Reference any related issues
   - Add screenshots if relevant

### PR Checklist

- [ ] Code follows style guidelines
- [ ] All tests pass
- [ ] New code has tests
- [ ] Documentation is updated
- [ ] Commit messages follow convention
- [ ] No merge conflicts

## Questions?

If you have questions, please:

1. Check existing documentation
2. Search existing issues
3. Open a new issue with the `question` label

Thank you for contributing!
