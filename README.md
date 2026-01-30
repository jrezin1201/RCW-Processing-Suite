# Lennar Excel Processor Service

A FastAPI service that processes Lennar "scheduled tasks" Excel exports and generates formatted summary reports with task categorization and QA reporting.

## Features

- **Excel Processing**: Parse Lennar scheduled tasks Excel exports with intelligent header detection
- **Task Classification**: YAML-driven rules engine for categorizing tasks into buckets (EXT PRIME, EXTERE, EXTERIOR UA, INTERIOR)
- **Data Aggregation**: Group and sum costs by Lot/Block and Plan
- **Background Jobs**: Asynchronous processing using Redis Queue (RQ)
- **Formatted Output**: Generate professional Excel summaries matching template requirements
- **QA Reporting**: Comprehensive reporting on parsing, classification, and data quality issues

## Tech Stack

- **Python 3.11**
- **FastAPI** for REST API
- **openpyxl** for Excel processing
- **Redis + RQ** for background job processing
- **Pydantic** for data validation
- **YAML** for classification rules

## Project Structure

```
.
├── app/
│   ├── api/
│   │   └── routes.py            # API endpoints
│   ├── core/
│   │   └── config.py            # Configuration
│   ├── services/
│   │   ├── jobs.py              # Redis/RQ job management
│   │   ├── parser_lennar.py     # Excel parsing logic
│   │   ├── classifier.py        # Task classification
│   │   ├── aggregator.py        # Data aggregation
│   │   ├── excel_writer.py      # Output Excel generation
│   │   └── worker_tasks.py      # Background job orchestration
│   ├── models/
│   │   └── schemas.py           # Pydantic models
│   ├── data/
│   │   └── mapping_rules.yaml   # Task classification rules
│   └── main.py                  # FastAPI application
├── tests/
│   └── test_classifier.py       # Classification tests
├── data/
│   ├── uploads/                 # Uploaded files directory
│   └── outputs/                 # Generated files directory
├── requirements.txt
├── .env.example
└── README.md
```

## Local Development Setup

### Prerequisites

- Python 3.11+
- Redis (for background jobs)
- pip or poetry

### Installation

1. **Clone the repository**
```bash
git clone <repository-url>
cd rcw-hoursworked-an-contracts
```

2. **Create virtual environment**
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

4. **Set up environment variables**
```bash
cp .env.example .env
# Edit .env file with your configuration
```

5. **Start Redis** (using Docker)
```bash
docker run -d -p 6379:6379 --name redis redis:latest
```

### Running the Service

1. **Start the RQ worker** (in one terminal)
```bash
source venv/bin/activate
rq worker default
```

2. **Start the API server** (in another terminal)
```bash
source venv/bin/activate
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

The API will be available at:
- API: http://localhost:8000
- Interactive docs: http://localhost:8000/docs
- Alternative docs: http://localhost:8000/redoc

## API Usage

### Upload Excel File

```bash
# Upload a Lennar scheduled tasks Excel file
curl -X POST "http://localhost:8000/api/v1/uploads" \
  -F "file=@path/to/lennar_export.xlsx"
```

Response:
```json
{
  "job_id": "550e8400-e29b-41d4-a716-446655440000"
}
```

### Check Job Status

```bash
# Get job status and progress
curl "http://localhost:8000/api/v1/jobs/550e8400-e29b-41d4-a716-446655440000"
```

Response:
```json
{
  "job_id": "550e8400-e29b-41d4-a716-446655440000",
  "status": "succeeded",
  "progress": 1.0,
  "message": "Successfully processed 150 rows into 25 summary rows",
  "result": {
    "output_path": "data/outputs/550e8400-e29b-41d4-a716-446655440000_summary.xlsx",
    "qa_report": {
      "parse_meta": {
        "total_rows_seen": 180,
        "rows_parsed": 150,
        "rows_skipped_missing_fields": 30
      },
      "counts_per_bucket": {
        "EXT PRIME": 40,
        "EXTERE": 35,
        "EXTERIOR UA": 30,
        "INTERIOR": 35,
        "UNMAPPED": 10
      },
      "unmapped_count": 10,
      "summary_rows_generated": 25
    }
  }
}
```

### Download Processed File

```bash
# Download the generated Excel summary
curl -O "http://localhost:8000/api/v1/jobs/550e8400-e29b-41d4-a716-446655440000/download"
```

## Processing Workflow

1. **Upload**: Client uploads Lennar Excel export via POST /api/v1/uploads
2. **Queue**: Job is queued in Redis with unique job_id
3. **Parse**: Worker parses Excel, finding headers and extracting data
4. **Classify**: Each task is classified using YAML rules (EXT PRIME, EXTERE, etc.)
5. **Aggregate**: Data is grouped by Lot/Block + Plan and summed by category
6. **Generate**: Formatted Excel summary is created with:
   - Main sheet: Summary table with totals
   - QA sheet: Statistics and unmapped tasks
7. **Download**: Client retrieves the processed file

## Task Classification Rules

The service uses a YAML-based rules engine to classify tasks. Rules are defined in `app/data/mapping_rules.yaml`:

```yaml
rules:
  - bucket: "EXT PRIME"
    all_contains: ["exterior", "prime"]

  - bucket: "EXTERIOR UA"
    all_contains: ["exterior", "[ua]"]

  - bucket: "INTERIOR"
    any_contains: ["interior"]

  - bucket: "EXTERE"
    all_contains: ["exterior"]
    none_contains: ["prime", "[ua]"]
```

Each rule supports:
- `all_contains`: All terms must be present
- `any_contains`: At least one term must be present
- `none_contains`: None of these terms should be present

## Output Format

The generated Excel file contains:

### Summary Sheet
- Headers: LOT, PLAN, EXT PRIME, EXTERE, EXTERIOR UA, INTERIOR, Total
- One row per unique Lot/Block + Plan combination
- Currency formatting for all money columns
- Yellow highlight on Total column
- Bottom TOTAL row summing all columns

### QA Report Sheet
- Parsing statistics (rows seen, parsed, skipped)
- Classification counts per bucket
- Top 30 unmapped task examples with counts
- Suspicious totals (negative or > $100k)

## Testing

Run the classifier tests:
```bash
pytest tests/test_classifier.py -v
```

## Docker Deployment (Optional)

Build and run with Docker:

```bash
# Build image
docker build -t lennar-processor .

# Run with docker-compose
docker-compose up -d
```

## Troubleshooting

### Redis Connection Errors
```bash
# Check if Redis is running
redis-cli ping
# Should return: PONG

# If using Docker, check container
docker ps | grep redis
```

### Worker Not Processing Jobs
```bash
# Check RQ worker logs
# Ensure worker is running in separate terminal
# Check Redis queue
redis-cli LLEN rq:queue:default
```

### Excel Parsing Issues
- Ensure the Excel file has required headers: "Lot/Block", "Plan", "Task", "Task Start Date"
- Headers must be within the first 50 rows
- File must be .xlsx format (not .xls or .csv)

## License

MIT