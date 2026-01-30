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
    job_id: str,
    phase: str = None,
    project_name: str = None,
    house_string: str = None
) -> str:
    """
    Write summary data to a formatted Excel file.

    Args:
        summary_rows: List of aggregated summary rows
        qa_report: QA report data
        job_id: Job identifier for filename
        phase: Phase number extracted from input file
        project_name: Project name extracted from B3
        house_string: House string extracted from B5

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
    header_font = Font(bold=True, size=16)
    normal_font = Font(size=16)
    # Accounting format: aligns currency symbols, shows negatives in parentheses, zeros as dash
    accounting_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # Write title row (placeholders) - shifted one column to the right
    ws["B1"] = f"Project Name: {project_name if project_name else ''}"
    ws["G1"] = f"Phase: {phase if phase else ''}"  # Moved from D1 to G1
    ws["I1"] = "Job #:"  # Moved from F1 to I1
    ws["H1"] = house_string if house_string else ""  # House string in column H, row 1

    # Merge cells B1 and C1 for Project Name
    ws.merge_cells('B1:C1')

    ws["B1"].font = header_font
    ws["G1"].font = header_font  # Updated from D1
    ws["H1"].font = header_font
    ws["I1"].font = header_font  # Updated from F1

    # Write headers - starting from column B (column 2)
    headers = ["LOT", "PLAN", "EXT PRIME", "EXTERE", "EXTERIOR UA", "INTERIOR",
               "Rolls Walls Final", "Touch Up", "Q4 Reversal", "Total"]
    header_row = 3
    for col_idx, header in enumerate(headers, 2):  # Start from column 2 (B)
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        if header == "Total":
            cell.fill = yellow_fill

    # Write data rows - starting from column B (column 2)
    data_start_row = header_row + 1
    for row_idx, summary_row in enumerate(summary_rows, data_start_row):
        # Convert lot_block to number if possible
        try:
            lot_value = int(summary_row.lot_block)
        except (ValueError, TypeError):
            try:
                lot_value = float(summary_row.lot_block)
            except (ValueError, TypeError):
                lot_value = summary_row.lot_block  # Keep as string if not numeric

        lot_cell = ws.cell(row=row_idx, column=2, value=lot_value)
        lot_cell.border = thin_border
        lot_cell.font = normal_font  # Column B

        plan_cell = ws.cell(row=row_idx, column=3, value=summary_row.plan)
        plan_cell.border = thin_border
        plan_cell.font = normal_font  # Column C

        # Money columns with formatting
        money_values = [
            summary_row.ext_prime,      # Column D
            summary_row.extere,          # Column E
            summary_row.exterior_ua,     # Column F
            summary_row.interior,        # Column G
            0,                          # Column H - Rolls Walls Final
            0,                          # Column I - Touch Up
            0,                          # Column J - Q4 Reversal
            summary_row.total           # Column K - Total
        ]

        for col_idx, value in enumerate(money_values, 4):  # Start from column 4 (D)
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.number_format = accounting_format
            cell.font = normal_font
            cell.border = thin_border
            if col_idx == 11:  # Total column (now column K)
                cell.fill = yellow_fill

    # Add total row
    total_row = data_start_row + len(summary_rows)

    # Create larger font for total row
    total_font = Font(bold=True, size=16)

    # Add empty cells for columns B through I (no borders, no values)
    for col_idx in range(2, 10):  # Columns 2-9 (B-I)
        ws.cell(row=total_row, column=col_idx, value="")
        # No border applied

    # Add "TOTAL" label in column J (column 10)
    total_label_cell = ws.cell(row=total_row, column=10, value="TOTAL")
    total_label_cell.font = total_font
    # No border applied

    # Calculate and write total ONLY for the Total column (column K)
    col_letter = get_column_letter(11)  # Column K
    formula = f"=SUM({col_letter}{data_start_row}:{col_letter}{total_row-1})"
    cell = ws.cell(row=total_row, column=11, value=formula)
    cell.number_format = accounting_format
    cell.font = total_font
    # No border applied
    cell.fill = yellow_fill

    # Add Labor calculations (absolute rows 16 and 18)
    labor_row_1 = 16
    labor_row_2 = 18

    # Add "LABOR:" in column E, row 16
    labor_cell = ws.cell(row=labor_row_1, column=5, value="LABOR:")  # Column E is column 5
    labor_cell.font = normal_font

    # Add 43% of total in column F, row 16
    # Reference the actual total row (row 13 as user specified, but we'll use dynamic total_row)
    # User said "column K, row 13" but the total row is dynamic, so we use total_row
    total_ref = f"K{total_row}"  # Reference to total in column K
    formula_43 = f"={total_ref}*0.43"
    cell_43 = ws.cell(row=labor_row_1, column=6, value=formula_43)  # Column F is column 6
    cell_43.number_format = accounting_format
    cell_43.font = normal_font

    # Add description in column G, row 16
    desc_43 = ws.cell(row=labor_row_1, column=7, value="43% of total amount")  # Column G is column 7
    desc_43.font = normal_font

    # Add "MATERIAL:" in column E, row 18
    material_cell = ws.cell(row=labor_row_2, column=5, value="MATERIAL:")  # Column E is column 5
    material_cell.font = normal_font

    # Add 28% of total in column F, row 18
    formula_28 = f"={total_ref}*0.28"
    cell_28 = ws.cell(row=labor_row_2, column=6, value=formula_28)  # Column F is column 6
    cell_28.number_format = accounting_format
    cell_28.font = normal_font

    # Add description in column G, row 18
    desc_28 = ws.cell(row=labor_row_2, column=7, value="28% of total amount")  # Column G is column 7
    desc_28.font = normal_font

    # Auto-size column widths based on content
    ws.column_dimensions["A"].width = 2.5  # Keep narrow column A fixed

    # Calculate optimal widths for each column based on content
    for column in ws.columns:
        column_letter = get_column_letter(column[0].column)

        # Skip column A (keep it narrow)
        if column_letter == "A":
            continue

        max_length = 0

        # Check all cells in the column for the longest content
        for cell in column:
            try:
                if cell.value:
                    # Convert to string to measure length
                    cell_value = str(cell.value)

                    # For formulas, estimate based on typical result length
                    if cell_value.startswith("="):
                        cell_length = 12  # Estimate for formula results
                    else:
                        # Account for currency formatting adding extra characters
                        if isinstance(cell.value, (int, float)):
                            # Add extra space for accounting format ($ symbol, commas, parentheses)
                            cell_length = len(f"${cell.value:,.2f}") + 2
                        else:
                            cell_length = len(str(cell.value))

                    # Account for bold font taking slightly more space
                    if cell.font and cell.font.bold:
                        cell_length = cell_length * 1.1

                    if cell_length > max_length:
                        max_length = cell_length
            except:
                pass

        # Set minimum and maximum widths
        adjusted_width = max(8, min(max_length + 2, 40))  # Min 8, Max 40

        # Apply specific adjustments for known columns
        if column_letter == "G":  # Description column needs more space
            adjusted_width = max(adjusted_width, 20)

        ws.column_dimensions[column_letter].width = adjusted_width

    # Create QA sheet
    qa_ws = wb.create_sheet(title="QA Report")
    write_qa_sheet(qa_ws, qa_report)

    # Save the file with contracts-forms prefix
    output_path = output_dir / f"contracts-forms_{job_id}.xlsx"
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
    header_font = Font(bold=True, size=16)
    normal_font = Font(size=16)
    accounting_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    current_row = 1

    # Parsing statistics
    ws.cell(row=current_row, column=1, value="Parsing Statistics").font = header_font
    current_row += 1
    ws.cell(row=current_row, column=1, value="Total Rows Seen:").font = normal_font
    ws.cell(row=current_row, column=2, value=qa_report.parse_meta.total_rows_seen).font = normal_font
    current_row += 1
    ws.cell(row=current_row, column=1, value="Rows Parsed:").font = normal_font
    ws.cell(row=current_row, column=2, value=qa_report.parse_meta.rows_parsed).font = normal_font
    current_row += 1
    ws.cell(row=current_row, column=1, value="Rows Skipped (Missing Fields):").font = normal_font
    ws.cell(row=current_row, column=2, value=qa_report.parse_meta.rows_skipped_missing_fields).font = normal_font
    current_row += 2

    # Classification counts
    ws.cell(row=current_row, column=1, value="Classification Counts").font = header_font
    current_row += 1
    for bucket, count in qa_report.counts_per_bucket.items():
        ws.cell(row=current_row, column=1, value=bucket).font = normal_font
        ws.cell(row=current_row, column=2, value=count).font = normal_font
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
            ws.cell(row=current_row, column=1, value=example["task_text"]).font = normal_font
            ws.cell(row=current_row, column=2, value=example["count"]).font = normal_font
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
            ws.cell(row=current_row, column=1, value=suspicious["lot_block"]).font = normal_font
            ws.cell(row=current_row, column=2, value=suspicious["plan"]).font = normal_font
            total_cell = ws.cell(row=current_row, column=3, value=suspicious["total"])
            total_cell.number_format = accounting_format
            total_cell.font = normal_font
            ws.cell(row=current_row, column=4, value=suspicious["reason"]).font = normal_font
            current_row += 1

    # Adjust column widths
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 20