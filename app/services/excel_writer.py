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
    light_blue_fill = PatternFill(start_color="E0F2F7", end_color="E0F2F7", fill_type="solid")  # Light blue/teal
    light_gray_fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")  # Light gray for headers
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # Row 1: Project header
    ws["B1"] = f"Project Name: {project_name if project_name else ''}"
    ws["G1"] = f"Phase:{phase if phase else ''}"  # No space after "Phase:"
    ws["H1"] = house_string if house_string else ""
    ws["I1"] = "Job #:"

    # Merge cells B1, C1, and D1 for Project Name
    ws.merge_cells('B1:D1')

    ws["B1"].font = header_font
    ws["G1"].font = header_font
    ws["H1"].font = header_font
    ws["H1"].fill = light_blue_fill  # Apply light blue/teal background to H1
    ws["I1"].font = header_font

    # Row 2: Empty (leave blank)

    # Row 3: Headers
    headers = ["LOT", "PLAN", "EXT PRIME", "EXTERIOR", "EXTERIOR UA", "INTERIOR",
               "ROLL WALLS FINAL", "TOUCH UP", "Q4 REVERSAL", "Total"]
    header_row = 3

    # Column A3 is blank for headers
    ws.cell(row=header_row, column=1, value="")

    for col_idx, header in enumerate(headers, 2):  # Start from column 2 (B)
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if header == "Total":
            cell.fill = yellow_fill
        else:
            cell.fill = light_gray_fill  # Light gray background for headers

    # Write data rows starting from row 4
    data_start_row = 4
    for row_num, summary_row in enumerate(summary_rows, 1):
        row_idx = data_start_row + row_num - 1

        # Column A: Row number
        ws.cell(row=row_idx, column=1, value=row_num).font = normal_font

        # Column B: LOT
        try:
            lot_value = int(summary_row.lot_block)
        except (ValueError, TypeError):
            try:
                lot_value = float(summary_row.lot_block)
            except (ValueError, TypeError):
                lot_value = summary_row.lot_block

        ws.cell(row=row_idx, column=2, value=lot_value).font = normal_font

        # Column C: PLAN
        ws.cell(row=row_idx, column=3, value=summary_row.plan).font = normal_font

        # Money columns D-J (not including K which will be a formula)
        money_values = [
            summary_row.ext_prime,      # Column D - EXT PRIME
            summary_row.extere,          # Column E - EXTERIOR
            summary_row.exterior_ua,     # Column F - EXTERIOR UA
            summary_row.interior,        # Column G - INTERIOR
            0,                          # Column H - ROLL WALLS FINAL
            0,                          # Column I - TOUCH UP
            0,                          # Column J - Q4 REVERSAL
        ]

        for col_idx, value in enumerate(money_values, 4):  # Start from column 4 (D)
            cell = ws.cell(row=row_idx, column=col_idx, value=value if value else 0)
            cell.number_format = accounting_format
            cell.font = normal_font

        # Column K: Total (as SUM formula)
        total_formula = f"=SUM(D{row_idx}:J{row_idx})"
        total_cell = ws.cell(row=row_idx, column=11, value=total_formula)
        total_cell.number_format = accounting_format
        total_cell.font = normal_font
        total_cell.fill = yellow_fill

    # After last data row: Skip one row (empty row)
    total_row = data_start_row + len(summary_rows) + 1  # +1 for the empty row

    # Apply yellow fill to all cells in Total column (K) from header through totals
    for row in range(header_row, max(total_row + 1, 14)):
        cell = ws.cell(row=row, column=11)  # Column K is column 11
        cell.fill = yellow_fill

    # Create larger font for total row
    total_font = Font(bold=True, size=16)

    # Add "TOTAL" label in column I (column 9)
    total_label_cell = ws.cell(row=total_row, column=9, value="TOTAL")
    total_label_cell.font = total_font

    # Calculate and write sum in column K
    col_letter = get_column_letter(11)  # Column K
    formula = f"=SUM({col_letter}{data_start_row}:{col_letter}{total_row-2})"  # -2 because we added empty row
    cell = ws.cell(row=total_row, column=11, value=formula)
    cell.number_format = accounting_format
    cell.font = total_font
    cell.fill = yellow_fill

    # Skip 2 rows after total, then add Labor row
    labor_row = total_row + 3
    material_row = labor_row + 2  # Skip 1 row after labor

    # Add "LABOR:" in column D
    labor_cell = ws.cell(row=labor_row, column=4, value="LABOR:")  # Column D is column 4
    labor_cell.font = normal_font

    # Add labor calculation in column E (43% of total)
    labor_calc_cell = ws.cell(row=labor_row, column=5, value=f"=K{total_row}*0.43")  # Column E is column 5
    labor_calc_cell.number_format = accounting_format
    labor_calc_cell.font = normal_font

    # Add description in column G for labor
    desc_labor = ws.cell(row=labor_row, column=7, value="Will be 43% of total amount")  # Column G
    desc_labor.font = normal_font

    # Add "MATERIAL:" in column D
    material_cell = ws.cell(row=material_row, column=4, value="MATERIAL:")  # Column D
    material_cell.font = normal_font

    # Add material calculation in column E (28% of total)
    material_calc_cell = ws.cell(row=material_row, column=5, value=f"=K{total_row}*0.28")  # Column E is column 5
    material_calc_cell.number_format = accounting_format
    material_calc_cell.font = normal_font

    # Add description in column G for material
    desc_material = ws.cell(row=material_row, column=7, value="will be 28% of total amount")  # Column G, lowercase 'will'
    desc_material.font = normal_font

    # Set column widths according to specifications
    # Converting pixel widths to Excel units (approximately 7 pixels = 1 unit)
    ws.column_dimensions["A"].width = 4     # 30px (narrow index column)
    ws.column_dimensions["B"].width = 7     # 50px (LOT)
    ws.column_dimensions["C"].width = 7     # 50px (PLAN)
    ws.column_dimensions["D"].width = 12    # 90px (EXT PRIME)
    ws.column_dimensions["E"].width = 12    # 90px (EXTERIOR)
    ws.column_dimensions["F"].width = 12    # 90px (EXTERIOR UA)
    ws.column_dimensions["G"].width = 12    # 90px (INTERIOR)
    ws.column_dimensions["H"].width = 16    # 120px (ROLL WALLS FINAL)
    ws.column_dimensions["I"].width = 12    # 90px (TOUCH UP)
    ws.column_dimensions["J"].width = 12    # 90px (Q4 REVERSAL)
    ws.column_dimensions["K"].width = 12    # 90px (Total)

    # Create QA sheet
    qa_ws = wb.create_sheet(title="QA Report")
    write_qa_sheet(qa_ws, qa_report)

    # Save the file with contracts_forms prefix
    output_path = output_dir / f"contracts_forms_{job_id}.xlsx"
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