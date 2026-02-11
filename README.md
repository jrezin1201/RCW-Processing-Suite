# RCW Processing Suite

A FastAPI service that processes Lennar "scheduled tasks" Excel exports and Gas & Rig hours-worked files, generating formatted summary reports.

## Features

- **Lennar Excel Processing**: Parse Lennar scheduled tasks Excel exports with intelligent header detection
- **Task Classification**: Signal-based category mapping with auto-category creation for new task types
- **Duplicate Draw Handling**: When the same draw name appears multiple times for one house, each gets its own column
- **Data Aggregation**: Group costs by Lot/Block and Plan with dynamic category columns
- **Gas & Rig Processing**: Parse hours-worked files and compute job costs
- **Formatted Output**: Generate professional Excel summaries with accounting formatting
- **QA Reporting**: Comprehensive reporting on parsing, classification, and data quality

## Tech Stack

- **Python 3.11+**
- **FastAPI** for REST API
- **openpyxl** / **pandas** for Excel processing
- **Pydantic** for data validation

## Project Structure

```
.
├── app/
│   ├── api/
│   │   └── lennar_routes.py     # API endpoints (upload, status, download, gas-rig)
│   ├── core/
│   │   └── config.py            # Configuration
│   ├── services/
│   │   ├── jobs.py              # In-memory job management
│   │   ├── parser_lennar.py     # Excel parsing logic
│   │   ├── category_mapper.py   # Signal-based task classification
│   │   ├── aggregator.py        # Data aggregation (with duplicate column support)
│   │   ├── excel_writer.py      # Output Excel generation
│   │   ├── gas_rig.py           # Gas & Rig processor
│   │   └── worker_tasks.py      # Job orchestration
│   ├── models/
│   │   └── schemas.py           # Pydantic models
│   └── main.py                  # FastAPI application
├── tests/
│   └── test_category_mapper.py  # Category mapper tests
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
- pip or poetry

### Installation

1. **Clone the repository**
```bash
git clone <repository-url>
cd rcw-processing-suite
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
```

### Running the Service

```bash
source venv/bin/activate
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

The API will be available at:
- API: http://localhost:8000
- Interactive docs: http://localhost:8000/docs
- Alternative docs: http://localhost:8000/redoc

## API Usage

### Upload Lennar Excel File

```bash
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
curl "http://localhost:8000/api/v1/jobs/550e8400-e29b-41d4-a716-446655440000"
```

### Download Processed File

```bash
curl -O "http://localhost:8000/api/v1/jobs/550e8400-e29b-41d4-a716-446655440000/download"
```

### Process Gas & Rig File

```bash
curl -X POST "http://localhost:8000/api/v1/gas-rig/process" \
  -F "file=@path/to/hours_worked.xlsx" \
  --output gas_rig_summary.xlsx
```

## Processing Workflow

1. **Upload**: Client uploads Lennar Excel export via POST /api/v1/uploads
2. **Parse**: Excel file is parsed, finding headers and extracting data rows
3. **Classify**: Each task is classified using signal-based category mapping (EXT PRIME, EXTERIOR, INTERIOR, etc.)
4. **Aggregate**: Data is grouped by Lot/Block + Plan with dynamic category columns
   - If a house has duplicate draws (e.g. two "Touch Up" entries), each gets its own column ("TOUCH UP", "TOUCH UP (2)")
5. **Generate**: Formatted Excel summary is created with:
   - Summary sheet: Table with LOT, PLAN, category columns, and Total
   - QA sheet: Statistics, category counts, and unmapped tasks
6. **Download**: Client retrieves the processed file

## Task Classification

The service uses signal-based classification via `category_mapper.py`. Signals extracted from task text include:

- **Location**: (EXT), (INT), EXTERIOR, INTERIOR
- **Designation**: [UA], [OP], [LS]
- **Keywords**: PRIME, TOUCH UP, ROLL WALLS, BASE SHOE, etc.

Categories are mapped template-first. If a task doesn't match any template category, a new column is auto-created so no dollars are lost.

## Output Format

### Summary Sheet
- Dynamic columns: LOT, PLAN, [category columns], Total
- One row per unique Lot/Block + Plan combination
- Duplicate draws get separate numbered columns
- Currency (accounting) formatting for all money columns
- Yellow highlight on Total column
- Bottom TOTAL row, plus LABOR (43%) and MATERIAL (28%) calculations

### QA Report Sheet
- Parsing statistics (rows seen, parsed, skipped)
- Classification counts per category
- Auto-created categories with example tasks
- Top unmapped task examples
- Suspicious totals (negative or > $100k)

## Testing

```bash
pytest tests/ -v
```

## Docker Deployment

```bash
docker-compose up -d
```

## Troubleshooting

### Excel Parsing Issues
- Ensure the Excel file has required headers: "Lot/Block", "Plan", "Task", "Task Start Date"
- Headers must be within the first 50 rows
- File must be .xlsx format (not .xls or .csv)

## License

MIT
