"""Parser for Lennar scheduled tasks Excel exports."""
from typing import List, Dict, Tuple, Optional, Any
from datetime import datetime
import os
import logging

from app.models.schemas import ParsedRow, QAMeta

logger = logging.getLogger(__name__)


def parse_lennar_export(filepath: str) -> Tuple[List[ParsedRow], QAMeta]:
    """
    Parse a Lennar scheduled tasks Excel export (supports both .xls and .xlsx).

    Args:
        filepath: Path to the Excel file

    Returns:
        Tuple of (parsed rows, QA metadata)
    """
    # Try to detect format by actually attempting to parse
    # This handles cases where the extension doesn't match the actual format
    try:
        # Try modern format first (xlsx)
        logger.info(f"Attempting to parse {filepath} as .xlsx format")
        return parse_with_openpyxl(filepath)
    except Exception as e:
        logger.info(f"Failed with openpyxl (likely .xls format): {str(e)[:100]}")
        # Fallback to pandas for old format
        try:
            logger.info(f"Attempting to parse {filepath} as .xls format with pandas")
            return parse_with_pandas(filepath)
        except Exception as e2:
            logger.error(f"Failed with both parsers: {e2}")
            # Try one more time with pandas using None engine (auto-detect)
            import pandas as pd
            try:
                df = pd.read_excel(filepath, engine=None)
                logger.info(f"Successfully read with pandas auto-detect, processing...")
                return parse_with_pandas_df(df)
            except Exception as e3:
                logger.error(f"All parsing attempts failed: {e3}")
                raise Exception(f"Could not parse file as either .xls or .xlsx format: {str(e)[:200]}")


def parse_with_openpyxl(filepath: str) -> Tuple[List[ParsedRow], QAMeta]:
    """Parse using openpyxl for .xlsx files."""
    import openpyxl
    from openpyxl.worksheet.worksheet import Worksheet

    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    # Find header row
    header_row_idx, column_map = find_header_row_openpyxl(ws)

    if header_row_idx is None:
        return [], QAMeta(total_rows_seen=0, rows_parsed=0, rows_skipped_missing_fields=0)

    # Parse data rows
    parsed_rows = []
    qa_meta = QAMeta()
    consecutive_blank_rows = 0
    max_consecutive_blanks = 30

    for row_idx, row in enumerate(ws.iter_rows(min_row=header_row_idx + 1, values_only=True), start=header_row_idx + 1):
        qa_meta.total_rows_seen += 1

        # Check if row is effectively blank
        if is_row_blank(row):
            consecutive_blank_rows += 1
            if consecutive_blank_rows >= max_consecutive_blanks:
                break
            continue

        consecutive_blank_rows = 0

        # Parse the row
        parsed_row = parse_row(row, column_map)

        # Skip if missing required fields
        if not parsed_row.lot_block or not parsed_row.task_text:
            qa_meta.rows_skipped_missing_fields += 1
            continue

        parsed_rows.append(parsed_row)
        qa_meta.rows_parsed += 1

    wb.close()
    return parsed_rows, qa_meta


def parse_with_pandas(filepath: str) -> Tuple[List[ParsedRow], QAMeta]:
    """Parse using pandas for .xls files (old format)."""
    import pandas as pd

    # Read the Excel file with pandas (try xlrd first for .xls)
    try:
        df = pd.read_excel(filepath, engine='xlrd')
    except:
        # Fallback to auto-detect
        df = pd.read_excel(filepath, engine=None)

    return parse_with_pandas_df(df)


def parse_with_pandas_df(df) -> Tuple[List[ParsedRow], QAMeta]:
    """Parse a pandas DataFrame."""
    import pandas as pd

    # Convert DataFrame to list of lists for compatibility
    data = df.values.tolist()
    headers = df.columns.tolist()

    # Find header row in data
    header_row_idx, column_map = find_header_row_pandas(headers, data)

    if header_row_idx is None:
        return [], QAMeta(total_rows_seen=0, rows_parsed=0, rows_skipped_missing_fields=0)

    # Parse data rows
    parsed_rows = []
    qa_meta = QAMeta()
    consecutive_blank_rows = 0
    max_consecutive_blanks = 30

    # Start from the row after headers
    start_idx = 0 if header_row_idx == -1 else header_row_idx

    for row_idx, row in enumerate(data[start_idx:], start=start_idx):
        qa_meta.total_rows_seen += 1

        # Check if row is effectively blank
        if is_row_blank_pandas(row):
            consecutive_blank_rows += 1
            if consecutive_blank_rows >= max_consecutive_blanks:
                break
            continue

        consecutive_blank_rows = 0

        # Parse the row
        parsed_row = parse_row(row, column_map)

        # Skip if missing required fields
        if not parsed_row.lot_block or not parsed_row.task_text:
            qa_meta.rows_skipped_missing_fields += 1
            continue

        parsed_rows.append(parsed_row)
        qa_meta.rows_parsed += 1

    return parsed_rows, qa_meta


def find_header_row_openpyxl(ws) -> Tuple[Optional[int], Dict[str, int]]:
    """Find header row for openpyxl worksheet."""
    from openpyxl.worksheet.worksheet import Worksheet

    required_headers = ["lot/block", "plan", "task", "task start date"]
    max_rows_to_check = 50

    for row_idx, row in enumerate(ws.iter_rows(max_row=max_rows_to_check, values_only=True), start=1):
        if row is None:
            continue

        # Convert row to lowercase string list for comparison
        row_lower = [str(cell).lower().strip() if cell else "" for cell in row]

        # Check if all required headers are present
        found_headers = {}
        for header in required_headers:
            for col_idx, cell_value in enumerate(row_lower):
                if header in cell_value:
                    found_headers[header] = col_idx
                    break

        # If we found all required headers, build the column map
        if len(found_headers) == len(required_headers):
            column_map = build_column_map(row)
            return row_idx, column_map

    return None, {}


def find_header_row_pandas(headers: List, data: List) -> Tuple[Optional[int], Dict[str, int]]:
    """Find header row for pandas data."""
    required_headers = ["lot/block", "plan", "task", "task start date"]

    # First check if headers are in the column names
    headers_lower = [str(h).lower().strip() if h else "" for h in headers]
    found_headers = {}

    for header in required_headers:
        for col_idx, cell_value in enumerate(headers_lower):
            if header in cell_value:
                found_headers[header] = col_idx
                break

    if len(found_headers) == len(required_headers):
        column_map = build_column_map(headers)
        return -1, column_map  # -1 indicates headers are in column names

    # Otherwise search in data rows
    max_rows_to_check = min(50, len(data))

    for row_idx in range(max_rows_to_check):
        row = data[row_idx]
        if row is None:
            continue

        row_lower = [str(cell).lower().strip() if cell else "" for cell in row]
        found_headers = {}

        for header in required_headers:
            for col_idx, cell_value in enumerate(row_lower):
                if header in cell_value:
                    found_headers[header] = col_idx
                    break

        if len(found_headers) == len(required_headers):
            column_map = build_column_map(row)
            return row_idx, column_map

    return None, {}


def find_header_row(ws) -> Tuple[Optional[int], Dict[str, int]]:
    """Legacy function - redirects to openpyxl version."""
    return find_header_row_openpyxl(ws)


def is_row_blank_pandas(row: list) -> bool:
    """Check if a pandas row is effectively blank."""
    import pandas as pd
    if row is None:
        return True
    return all(pd.isna(cell) or str(cell).strip() == "" for cell in row)


def build_column_map(header_row: tuple) -> Dict[str, int]:
    """
    Build a mapping of column names to indices.

    Args:
        header_row: The header row values

    Returns:
        Dictionary mapping column names to indices
    """
    column_map = {}

    for idx, header in enumerate(header_row):
        if header is None:
            continue

        header_lower = str(header).lower().strip()

        # Map headers to standardized names
        if "lot" in header_lower and "block" in header_lower:
            column_map["lot_block"] = idx
        elif "plan" in header_lower:
            column_map["plan"] = idx
        elif "elevation" in header_lower:
            column_map["elevation"] = idx
        elif "swing" in header_lower:
            column_map["swing"] = idx
        elif "task start date" in header_lower:
            column_map["task_start_date"] = idx
        elif "task" in header_lower and "date" not in header_lower:
            column_map["task_text"] = idx
        elif "subtotal" in header_lower:
            column_map["subtotal"] = idx
        elif "tax" in header_lower:
            column_map["tax"] = idx
        elif "total" in header_lower and "subtotal" not in header_lower:
            column_map["total"] = idx

    return column_map


def is_row_blank(row: tuple) -> bool:
    """
    Check if a row is effectively blank.

    Args:
        row: Row values

    Returns:
        True if row is blank
    """
    if row is None:
        return True
    return all(cell is None or str(cell).strip() == "" for cell in row)


def parse_row(row: tuple, column_map: Dict[str, int]) -> ParsedRow:
    """
    Parse a single data row.

    Args:
        row: Row values
        column_map: Column name to index mapping

    Returns:
        ParsedRow object
    """
    def get_value(key: str) -> Any:
        """Get value from row using column map."""
        idx = column_map.get(key)
        if idx is not None and idx < len(row):
            return row[idx]
        return None

    def parse_money(value: Any) -> Optional[float]:
        """Parse money value to float."""
        if value is None:
            return None

        # If it's already a number
        if isinstance(value, (int, float)):
            return float(value)

        # Convert to string and clean
        value_str = str(value).strip()
        # Remove currency symbols and commas
        value_str = value_str.replace("$", "").replace(",", "").strip()

        try:
            return float(value_str)
        except (ValueError, TypeError):
            return None

    def parse_date(value: Any) -> Optional[datetime]:
        """Parse date value."""
        if value is None:
            return None

        # If it's already a datetime
        if isinstance(value, datetime):
            return value

        # Try to parse string date
        try:
            # Common date formats
            date_str = str(value).strip()
            for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y"]:
                try:
                    return datetime.strptime(date_str, fmt)
                except ValueError:
                    continue
        except:
            pass

        return None

    return ParsedRow(
        lot_block=str(get_value("lot_block")).strip() if get_value("lot_block") else None,
        plan=str(get_value("plan")).strip() if get_value("plan") else None,
        elevation=str(get_value("elevation")).strip() if get_value("elevation") else None,
        swing=str(get_value("swing")).strip() if get_value("swing") else None,
        task_start_date=parse_date(get_value("task_start_date")),
        task_text=str(get_value("task_text")).strip() if get_value("task_text") else None,
        subtotal=parse_money(get_value("subtotal")),
        tax=parse_money(get_value("tax")),
        total=parse_money(get_value("total"))
    )