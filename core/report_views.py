from collections import defaultdict
from datetime import date
from io import BytesIO
from pathlib import Path

from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.shortcuts import render
from django.views.decorators.http import require_http_methods
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


MAX_REPORT_SIZE = 50 * 1024 * 1024
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}


class ReportUploadError(ValueError):
    pass


def normalize_log_number(value):
    return str(value or "").strip().upper()


def clean_text(value):
    if value in (None, ""):
        return ""

    # Remove extra spaces and line breaks.
    return " ".join(str(value).split())


def validate_upload(upload):
    extension = Path(upload.name).suffix.lower()

    if extension not in ALLOWED_EXTENSIONS:
        raise ReportUploadError(
            f"{upload.name}: please upload an .xlsx or .xlsm file."
        )

    if upload.size > MAX_REPORT_SIZE:
        raise ReportUploadError(
            f"{upload.name}: the file must be 50 MB or smaller."
        )


def open_report_worksheet(upload):
    validate_upload(upload)

    try:
        upload.seek(0)

        workbook = load_workbook(
            upload,
            read_only=True,
            data_only=True,
        )

    except Exception as exc:
        raise ReportUploadError(
            f"{upload.name} could not be read as an Excel workbook."
        ) from exc

    if "Report" in workbook.sheetnames:
        worksheet = workbook["Report"]
    else:
        worksheet = workbook.active

    return workbook, worksheet


def read_bin_location_report(upload, expected_warehouse):
    """
        Inventory report layout:

        A = Warehouse
        B = Description
        C = Log Number
        D = Location Type (usually F)
        E = Bin Location
        H = Available Weight
        I = Available Pieces
        L = On Hand Weight
        M = On Hand Pieces
        """
    workbook, worksheet = open_report_worksheet(upload)

    rows = []
    header_found = False

    try:
        for values in worksheet.iter_rows(values_only=True):
            values = list(values)

            # Ensure the row has at least 13 columns.
            values.extend([None] * max(0, 13 - len(values)))

            if (
                normalize_log_number(values[2]) == "LOG NUMBER"
                and clean_text(values[3]).upper() == "BIN LOCATION"
            ):
                header_found = True
                continue

            if not header_found:
                continue

            warehouse = clean_text(values[0])
            log_number = normalize_log_number(values[2])

            if warehouse != str(expected_warehouse) or not log_number:
                continue

            rows.append({
                "warehouse": warehouse,
                "log_number": log_number,
                "description": clean_text(values[1]),
                "bin_location": clean_text(values[4]),
                "available_weight": values[7],
                "available_pieces": values[8],
                "on_hand_weight": values[11],
                "on_hand_pieces": values[12],
            })

    finally:
        workbook.close()

    if not header_found:
        raise ReportUploadError(
            f"{upload.name} does not appear to be a bin-location report."
        )

    if not rows:
        raise ReportUploadError(
            f"{upload.name} does not contain warehouse "
            f"{expected_warehouse} inventory."
        )

    return rows


def read_remarks_report(upload):
    """
    Remarks report layout:

    A = Warehouse
    E = Log
    X = Remarks 1
    Y = Remarks 2
    Z = Remarks 3
    """
    workbook, worksheet = open_report_worksheet(upload)

    remarks_by_log = defaultdict(list)
    header_found = False

    try:
        for values in worksheet.iter_rows(values_only=True):
            values = list(values)

            # Ensure the row has at least 26 columns (through column Z).
            values.extend([None] * max(0, 26 - len(values)))

            if (
                normalize_log_number(values[4]) == "LOG"
                and clean_text(values[23]).upper() == "REMARKS 1"
            ):
                header_found = True
                continue

            if not header_found:
                continue

            warehouse = clean_text(values[0])
            log_number = normalize_log_number(values[4])

            if warehouse not in {"20", "21"} or not log_number:
                continue

            existing_lowercase = {
                remark.casefold()
                for remark in remarks_by_log[log_number]
            }

            for value in values[23:26]:
                remark = clean_text(value)

                if (
                    remark
                    and remark.casefold() not in existing_lowercase
                ):
                    remarks_by_log[log_number].append(remark)
                    existing_lowercase.add(remark.casefold())

    finally:
        workbook.close()

    if not header_found:
        raise ReportUploadError(
            f"{upload.name} does not appear to be a log-remarks report."
        )

    return remarks_by_log


def combine_reports(bin_report_20, bin_report_21, remarks_report):
    warehouse_20_rows = read_bin_location_report(
        bin_report_20,
        expected_warehouse=20,
    )

    warehouse_21_rows = read_bin_location_report(
        bin_report_21,
        expected_warehouse=21,
    )

    remarks_by_log = read_remarks_report(remarks_report)

    combined_rows = []

    for inventory_row in warehouse_20_rows + warehouse_21_rows:
        log_number = inventory_row["log_number"]
        remarks = remarks_by_log.get(log_number, [])

        combined_rows.append({
            "log_number": log_number,
            "description": inventory_row["description"],
            "bin_location": inventory_row["bin_location"],
            "available_pieces": inventory_row["available_pieces"],
            "available_weight": inventory_row["available_weight"],
            "on_hand_pieces": inventory_row["on_hand_pieces"],
            "on_hand_weight": inventory_row["on_hand_weight"],
            "remarks": " | ".join(remarks),
        })

    combined_rows.sort(
        key=lambda row: row["log_number"]
    )

    return combined_rows


def build_excel_report(rows):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Combined Report"

    headers = [
        "Log Number",
        "Description",
        "Bin Location",
        "Available Pieces",
        "Available Weight",
        "On Hand Pieces",
        "On Hand Weight",
        "Remarks",
    ]

    worksheet.append(headers)

    for row in rows:
        worksheet.append([
            row["log_number"],
            row["description"],
            row["bin_location"],
            row["available_pieces"],
            row["available_weight"],
            row["on_hand_pieces"],
            row["on_hand_weight"],
            row["remarks"],
        ])

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="0D2C54",
    )

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = Font(
            color="FFFFFF",
            bold=True,
        )
        cell.alignment = Alignment(
            vertical="center",
        )

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.row_dimensions[1].height = 25

    column_widths = {
        "A": 18,
        "B": 55,
        "C": 18,
        "D": 18,
        "E": 18,
        "F": 18,
        "G": 18,
        "H": 70,
    }

    for column, width in column_widths.items():
        worksheet.column_dimensions[column].width = width

    # Pieces
    for row_number in range(2, worksheet.max_row + 1):
        # Pieces
        worksheet[f"D{row_number}"].number_format = "#,##0"
        worksheet[f"F{row_number}"].number_format = "#,##0"

        # Weight
        worksheet[f"E{row_number}"].number_format = "#,##0.##"
        worksheet[f"G{row_number}"].number_format = "#,##0.##"

        # Description and remarks
        worksheet[f"B{row_number}"].alignment = Alignment(
            wrap_text=True,
            vertical="top",
        )

        worksheet[f"H{row_number}"].alignment = Alignment(
            wrap_text=True,
            vertical="top",
        )

    output = BytesIO()
    workbook.save(output)
    workbook.close()
    output.seek(0)

    return output


@login_required
@require_http_methods(["GET", "POST"])
def reports(request):
    context = {
        "combined_rows": [],
        "error": None,
    }

    if request.method == "POST":
        bin_report_20 = request.FILES.get("bin_report_20")
        bin_report_21 = request.FILES.get("bin_report_21")
        remarks_report = request.FILES.get("remarks_report")

        if not all([
            bin_report_20,
            bin_report_21,
            remarks_report,
        ]):
            context["error"] = (
                "Choose all three reports before continuing."
            )

        else:
            try:
                combined_rows = combine_reports(
                    bin_report_20,
                    bin_report_21,
                    remarks_report,
                )

            except ReportUploadError as exc:
                context["error"] = str(exc)

            else:
                action = request.POST.get("action", "display")

                if action == "download":
                    output = build_excel_report(combined_rows)

                    filename = (
                        f"combined_log_report_"
                        f"{date.today().isoformat()}.xlsx"
                    )

                    response = HttpResponse(
                        output.getvalue(),
                        content_type=(
                            "application/vnd.openxmlformats-"
                            "officedocument.spreadsheetml.sheet"
                        ),
                    )

                    response["Content-Disposition"] = (
                        f'attachment; filename="{filename}"'
                    )

                    return response

                context["combined_rows"] = combined_rows
                context["total_rows"] = len(combined_rows)
                context["remarks_count"] = sum(
                    1
                    for row in combined_rows
                    if row["remarks"]
                )

    return render(
        request,
        "core/reports.html",
        context,
    )