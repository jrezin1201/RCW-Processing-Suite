"""Excel output writer service for generating formatted summary files."""
import os
from pathlib import Path
from typing import List, Dict, Any
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from app.models.schemas import SummaryRow, QAReport


def write_summary_excel(
    summary_rows: List[SummaryRow],
    qa_report: QAReport,
    job_id: str
) -> str:
    """
    Write summary data to a formatted Excel file.

    Args:
        summary_rows: List of aggregated summary rows
        qa_report: QA report data
        job_id: Job identifier for filename

    Returns:
        Path to the generated Excel file
    """
    # Create output directory if it doesn't exist
    output_dir = Path("data/outputs")
    output_dir.mkdir(parents=True, exist_ok=True)

    # Create workbook and worksheets
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"

    # Define styles
    header_font = Font(bold=True)
    currency_format = '"$"#,##0.00'
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # Write title row (placeholders)
    ws["A1"] = "Project Name:"
    ws["C1"] = "Phase:"
    ws["E1"] = "Job #:"
    ws["A1"].font = header_font
    ws["C1"].font = header_font
    ws["E1"].font = header_font

    # Write headers
    headers = ["LOT", "PLAN", "EXT PRIME", "EXTERE", "EXTERIOR UA", "INTERIOR", "Total"]
    header_row = 3
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        if header == "Total":
            cell.fill = yellow_fill

    # Write data rows
    data_start_row = header_row + 1
    for row_idx, summary_row in enumerate(summary_rows, data_start_row):
        ws.cell(row=row_idx, column=1, value=summary_row.lot_block).border = thin_border
        ws.cell(row=row_idx, column=2, value=summary_row.plan).border = thin_border

        # Money columns with formatting
        money_values = [
            summary_row.ext_prime,
            summary_row.extere,
            summary_row.exterior_ua,
            summary_row.interior,
            summary_row.total
        ]

        for col_idx, value in enumerate(money_values, 3):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.number_format = currency_format
            cell.border = thin_border
            if col_idx == 7:  # Total column
                cell.fill = yellow_fill

    # Add total row
    total_row = data_start_row + len(summary_rows)
    ws.cell(row=total_row, column=1, value="TOTAL").font = header_font
    ws.cell(row=total_row, column=1).border = thin_border

    # Calculate and write column totals
    for col_idx in range(3, 8):
        col_letter = get_column_letter(col_idx)
        formula = f"=SUM({col_letter}{data_start_row}:{col_letter}{total_row-1})"
        cell = ws.cell(row=total_row, column=col_idx, value=formula)
        cell.number_format = currency_format
        cell.font = header_font
        cell.border = thin_border
        if col_idx == 7:  # Total column
            cell.fill = yellow_fill

    # Adjust column widths
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 12

    # Create QA sheet
    qa_ws = wb.create_sheet(title="QA Report")
    write_qa_sheet(qa_ws, qa_report)

    # Save the file
    output_path = output_dir / f"{job_id}_summary.xlsx"
    wb.save(output_path)
    wb.close()

    return str(output_path)


def write_qa_sheet(ws, qa_report: QAReport):
    """
    Write QA report data to a worksheet.

    Args:
        ws: The worksheet to write to
        qa_report: QA report data
    """
    header_font = Font(bold=True)
    current_row = 1

    # Parsing statistics
    ws.cell(row=current_row, column=1, value="Parsing Statistics").font = header_font
    current_row += 1
    ws.cell(row=current_row, column=1, value="Total Rows Seen:")
    ws.cell(row=current_row, column=2, value=qa_report.parse_meta.total_rows_seen)
    current_row += 1
    ws.cell(row=current_row, column=1, value="Rows Parsed:")
    ws.cell(row=current_row, column=2, value=qa_report.parse_meta.rows_parsed)
    current_row += 1
    ws.cell(row=current_row, column=1, value="Rows Skipped (Missing Fields):")
    ws.cell(row=current_row, column=2, value=qa_report.parse_meta.rows_skipped_missing_fields)
    current_row += 2

    # Classification counts
    ws.cell(row=current_row, column=1, value="Classification Counts").font = header_font
    current_row += 1
    for bucket, count in qa_report.counts_per_bucket.items():
        ws.cell(row=current_row, column=1, value=bucket)
        ws.cell(row=current_row, column=2, value=count)
        current_row += 1
    current_row += 1

    # Unmapped tasks
    if qa_report.unmapped_examples:
        ws.cell(row=current_row, column=1, value="Top Unmapped Tasks").font = header_font
        current_row += 1
        ws.cell(row=current_row, column=1, value="Task Text").font = header_font
        ws.cell(row=current_row, column=2, value="Count").font = header_font
        current_row += 1

        for example in qa_report.unmapped_examples[:30]:
            ws.cell(row=current_row, column=1, value=example["task_text"])
            ws.cell(row=current_row, column=2, value=example["count"])
            current_row += 1
        current_row += 1

    # Suspicious totals
    if qa_report.suspicious_totals:
        ws.cell(row=current_row, column=1, value="Suspicious Totals").font = header_font
        current_row += 1
        ws.cell(row=current_row, column=1, value="Lot/Block").font = header_font
        ws.cell(row=current_row, column=2, value="Plan").font = header_font
        ws.cell(row=current_row, column=3, value="Total").font = header_font
        ws.cell(row=current_row, column=4, value="Reason").font = header_font
        current_row += 1

        for suspicious in qa_report.suspicious_totals:
            ws.cell(row=current_row, column=1, value=suspicious["lot_block"])
            ws.cell(row=current_row, column=2, value=suspicious["plan"])
            ws.cell(row=current_row, column=3, value=suspicious["total"])
            ws.cell(row=current_row, column=3).number_format = '"$"#,##0.00'
            ws.cell(row=current_row, column=4, value=suspicious["reason"])
            current_row += 1

    # Adjust column widths
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 20