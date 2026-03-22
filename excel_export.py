"""
Excel Export — styled like the example screenshot
RTL layout, colored cells per employee, Hebrew headers
"""
from datetime import date, timedelta
from typing import Dict, List, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from scheduler import SHIFTS, SHIFT_HOURS, MORNING, EVENING, NIGHT, get_day_name

# Color palette for employees (matching screenshot style)
EMPLOYEE_COLORS = [
    "00B0F0",  # light blue
    "FF00FF",  # pink/magenta
    "FFC000",  # orange/gold
    "00B050",  # green
    "FFFF00",  # yellow
    "FF6600",  # dark orange
    "9966FF",  # purple
    "FF9999",  # light pink
]


def assign_colors(employees: List[str]) -> Dict[str, str]:
    """Assign a color to each employee."""
    colors = {}
    for i, emp in enumerate(employees):
        colors[emp] = EMPLOYEE_COLORS[i % len(EMPLOYEE_COLORS)]
    return colors


def export_to_excel(
    schedule: Dict[Tuple[date, str], str],
    employees: List[str],
    start_date: date,
    num_days: int,
    output_path: str = "משמרות.xlsx",
):
    """Export schedule to a styled Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "משמרות"
    ws.sheet_view.rightToLeft = True

    dates = [start_date + timedelta(days=i) for i in range(num_days)]
    colors = assign_colors(employees)

    # Styles
    header_font = Font(name="Arial", bold=True, size=12)
    cell_font = Font(name="Arial", size=11)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(name="Arial", bold=True, size=11, color="FFFFFF")

    # --- Row 1: Date numbers ---
    # Column A = shift info, columns B onwards = dates (RTL so rightmost = first date)
    # We'll put dates left-to-right in columns, RTL view will flip them

    # Column 1: "יום" header
    ws.cell(row=1, column=1, value="תאריך\nיום")
    ws.cell(row=1, column=1).font = header_font_white
    ws.cell(row=1, column=1).fill = header_fill
    ws.cell(row=1, column=1).alignment = center_align
    ws.cell(row=1, column=1).border = thin_border

    # Date columns
    for col_idx, d in enumerate(dates, start=2):
        # Row 1: date + day name
        cell = ws.cell(
            row=1,
            column=col_idx,
            value=f"{d.strftime('%d.%m')}\n{get_day_name(d)}",
        )
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        cell.fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

    # --- Shift rows ---
    shift_labels = {
        MORNING: f"משמרת\nבוקר\n{SHIFT_HOURS[MORNING]}",
        EVENING: f"משמרת\nצהריים\nערב\n{SHIFT_HOURS[EVENING]}",
        NIGHT: f"לילה\n{SHIFT_HOURS[NIGHT]}",
    }

    for shift_idx, shift in enumerate(SHIFTS):
        row = 2 + shift_idx

        # Shift label in column 1
        cell = ws.cell(row=row, column=1, value=shift_labels[shift])
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

        # Employee assignments
        for col_idx, d in enumerate(dates, start=2):
            emp = schedule.get((d, shift), "")
            cell = ws.cell(row=row, column=col_idx, value=emp)
            cell.font = cell_font
            cell.alignment = center_align
            cell.border = thin_border

            # Color the cell based on employee
            if emp and emp in colors:
                fill = PatternFill(
                    start_color=colors[emp],
                    end_color=colors[emp],
                    fill_type="solid",
                )
                cell.fill = fill

                # Use white font on dark backgrounds
                dark_colors = {"FF00FF", "00B050", "4472C4", "9966FF", "FF6600"}
                if colors[emp] in dark_colors:
                    cell.font = Font(name="Arial", size=11, color="FFFFFF")

    # --- Column widths ---
    ws.column_dimensions["A"].width = 16
    for col_idx in range(2, len(dates) + 2):
        ws.column_dimensions[get_column_letter(col_idx)].width = 12

    # --- Row heights ---
    ws.row_dimensions[1].height = 40
    for row in range(2, 5):
        ws.row_dimensions[row].height = 50

    # --- Legend: employee color reference ---
    legend_row = 6
    ws.cell(row=legend_row, column=1, value="מקרא צבעים:").font = Font(
        name="Arial", bold=True, size=11
    )
    for i, emp in enumerate(employees):
        col = 2 + i
        cell = ws.cell(row=legend_row, column=col, value=emp)
        cell.fill = PatternFill(
            start_color=colors[emp], end_color=colors[emp], fill_type="solid"
        )
        cell.font = cell_font
        cell.alignment = center_align
        cell.border = thin_border

    # --- Summary: total points per employee ---
    summary_row = 8
    ws.cell(row=summary_row, column=1, value="סה״כ נקודות:").font = Font(
        name="Arial", bold=True, size=11
    )
    from scheduler import SHIFT_COST

    for i, emp in enumerate(employees):
        total = 0
        for d_idx in range(num_days):
            d = start_date + timedelta(days=d_idx)
            for shift in SHIFTS:
                if schedule.get((d, shift)) == emp:
                    total += SHIFT_COST[shift]
        col = 2 + i
        cell = ws.cell(row=summary_row, column=col, value=f"{emp}: {total}")
        cell.font = cell_font
        cell.alignment = center_align
        cell.border = thin_border

    wb.save(output_path)
    return output_path
