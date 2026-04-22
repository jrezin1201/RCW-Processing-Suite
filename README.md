# RCW Processing Suite

FastAPI service that processes Excel exports for RC Wendt Painting. Ships
three pluggable modules:

| Module | What it does |
|--------|--------------|
| **Lennar Tasks** (`/api/v1/uploads`) | Parses Lennar scheduled-task Excel exports, classifies tasks via signal-based category mapping, aggregates by lot/plan, and emits a formatted summary + QA report |
| **Gas & Rig** (`/api/v1/gas-rig/process`) | Parses hours-worked files, extracts 4-digit job numbers, sums hours per job, calculates billing at $0.75/hour |
| **Missed Clock-In** (`/api/v1/missed-clock-in/process`) | Parses timekeeping Exception List exports and generates formatted employee warning notices (Overview sheet + one 34-row notice per violation) |

Modules auto-register at startup — see
[ARCHITECTURE.md](ARCHITECTURE.md#module-system) for the plug-in contract
and step-by-step instructions for adding a 4th or 5th module.

## Quick Start (local)

```bash
# 1. Create venv and install
python3.11 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 2. Configure
cp .env.example .env
# Edit .env — at minimum, set ENVIRONMENT=development for local use

# 3. Run
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Open [http://localhost:8000](http://localhost:8000).

## Docker

```bash
cp .env.example .env   # edit first!
docker compose up -d
```

The image runs as a non-root user, mounts `./data` for uploads/outputs, and
ships with a health check on `/health`.

## Configuration

All configuration is environment-driven. See [`.env.example`](.env.example)
for the full list. Key variables:

| Variable | Default | Notes |
|----------|---------|-------|
| `ENVIRONMENT` | `production` | `development` auto-enables `/docs` |
| `DEBUG` | `false` | Verbose logging. Never enable in prod. |
| `CORS_ORIGINS` | *(empty → deny)* | Comma-separated allowlist |
| `ENABLE_DOCS` | `false` | Gate `/docs`, `/redoc`, `/openapi.json` |
| `API_KEY` | *(empty → open)* | If set, `X-API-Key` header required on `/api/v1/*` |
| `MAX_UPLOAD_SIZE_MB` | `50` | Enforced at upload time |
| `UPLOAD_DIR` / `OUTPUT_DIR` | `data/uploads`, `data/outputs` | Relative to repo root |

## API

### Lennar (async job pattern)

```bash
# 1. Upload
curl -X POST http://localhost:8000/api/v1/uploads \
  -H "X-API-Key: $API_KEY" \
  -F "file=@lennar_export.xlsx"
# -> { "job_id": "550e8400-..." }

# 2. Poll
curl -H "X-API-Key: $API_KEY" \
  http://localhost:8000/api/v1/jobs/550e8400-.../

# 3. Download when status = succeeded
curl -OJ -H "X-API-Key: $API_KEY" \
  http://localhost:8000/api/v1/jobs/550e8400-.../download
```

### Gas & Rig (synchronous)

```bash
curl -X POST http://localhost:8000/api/v1/gas-rig/process \
  -H "X-API-Key: $API_KEY" \
  -F "file=@hours_worked.xlsx" \
  --output gas_rig_summary.xlsx
```

### Missed Clock-In (synchronous)

```bash
curl -X POST http://localhost:8000/api/v1/missed-clock-in/process \
  -H "X-API-Key: $API_KEY" \
  -F "file=@ExceptionList.xlsx" \
  -F "exclude_test=true" \
  --output Warning_Notices.xlsx
```

(`X-API-Key` is only required if `API_KEY` is set.)

## Architecture

The app is organized as a collection of self-contained modules under
`app/modules/`. Each module owns its routes, business logic, and data
models. The core app auto-discovers them at startup.

```
app/
├── core/
│   ├── config.py        # Env-driven settings
│   ├── registry.py      # Module auto-discovery
│   └── security.py      # Shared: API key + upload validation + path safety
├── modules/
│   ├── lennar/          # Excel parser + category mapper + aggregator + writer
│   ├── gas_rig/         # Hours → job costs
│   └── missed_clock_in/ # Exception List → warning notices
├── static/
├── templates/
│   └── professional_interface.html
└── main.py
```

See [ARCHITECTURE.md](ARCHITECTURE.md) for detailed module internals,
request flows, trade-offs, and the **Adding a New Module** guide.

## Security

- **CORS**: deny-all by default, allowlist via `CORS_ORIGINS`
- **OpenAPI docs**: disabled by default, enable via `ENABLE_DOCS` or
  `ENVIRONMENT=development`
- **Optional API key**: `X-API-Key` header enforced on all `/api/v1/*`
  endpoints when `API_KEY` is set (constant-time compare)
- **Upload validation**: extension + size cap + magic-byte sniff (xlsx/xls)
- **Path-traversal defense**: download endpoints validate the resolved
  path is inside `OUTPUT_DIR`
- **Error sanitization**: server-side tracebacks logged, client gets
  generic messages
- **Non-root Docker user** (UID 1000)

## Testing

```bash
pytest
```

Tests currently cover Lennar category mapping (signal extraction,
template mapping, auto-creation, deduplication, real-world task examples).

## Development

```bash
make help        # list commands
make dev         # uvicorn with reload
make test        # pytest
make format      # black + ruff fix
make lint        # ruff check
make typecheck   # mypy
make clean       # remove caches
```
