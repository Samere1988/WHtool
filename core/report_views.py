from collections import defaultdict
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
import re

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.http import HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import (
    require_GET,
    require_http_methods,
    require_POST,
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from .models import (
    HymusTransferItem,
    InventoryItem,
    InventoryReport,
)


MAX_REPORT_SIZE = 50 * 1024 * 1024
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}

SHEET_DESCRIPTION_PREFIXES = (
    "SSH",
    "SPL",
    "APL",
    "ASH",
    "ACP",
    "SCL",
    "ACL",
)

class ReportUploadError(ValueError):
    pass


def normalize_log_number(value):
    return str(value or "").strip().upper()


def clean_text(value):
    if value in (None, ""):
        return ""

    return " ".join(str(value).split())
def bin_location_sort_key(item):
    location = clean_text(
        item.bin_location
    ).upper()

    return [
        int(part)
        if part.isdigit()
        else part
        for part in re.split(
            r"(\d+)",
            location,
        )
    ]


def is_sheet_material(item):
    description = clean_text(
        item.description
    ).upper()

    return description.startswith(
        SHEET_DESCRIPTION_PREFIXES
    )

def clean_decimal(value, field_name, log_number):
    if value in (None, ""):
        return None

    try:
        number = Decimal(
            str(value).replace(",", "").strip()
        )
    except (InvalidOperation, ValueError) as exc:
        raise ReportUploadError(
            f"Log {log_number}: "
            f"{field_name} is not a valid number."
        ) from exc

    if not number.is_finite():
        raise ReportUploadError(
            f"Log {log_number}: "
            f"{field_name} is not a valid number."
        )

    return number.quantize(Decimal("0.01"))


def clean_integer(value, field_name, log_number):
    number = clean = clean_decimal(
        value,
        field_name,
        log_number,
    )

    if number is None:
        return None

    if number != number.to_integral_value():
        raise ReportUploadError(
            f"Log {log_number}: "
            f"{field_name} must be a whole number."
        )

    return int(number)


def validate_upload(upload):
    extension = Path(upload.name).suffix.lower()

    if extension not in ALLOWED_EXTENSIONS:
        raise ReportUploadError(
            f"{upload.name}: please upload "
            f"an .xlsx or .xlsm file."
        )

    if upload.size > MAX_REPORT_SIZE:
        raise ReportUploadError(
            f"{upload.name}: the file must be "
            f"50 MB or smaller."
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
            f"{upload.name} could not be read "
            f"as an Excel workbook."
        ) from exc

    if "Report" in workbook.sheetnames:
        worksheet = workbook["Report"]
    else:
        worksheet = workbook.active

    return workbook, worksheet


def read_bin_location_report(
    upload,
    expected_warehouse,
):
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
    workbook, worksheet = open_report_worksheet(
        upload
    )

    rows = []
    header_found = False

    try:
        for values in worksheet.iter_rows(
            values_only=True
        ):
            values = list(values)

            values.extend(
                [None] * max(
                    0,
                    13 - len(values),
                )
            )

            if (
                normalize_log_number(values[2])
                == "LOG NUMBER"
                and clean_text(values[3]).upper()
                == "BIN LOCATION"
            ):
                header_found = True
                continue

            if not header_found:
                continue

            warehouse = clean_text(values[0])
            log_number = normalize_log_number(
                values[2]
            )

            if (
                warehouse
                != str(expected_warehouse)
                or not log_number
            ):
                continue

            rows.append({
                "warehouse": warehouse,
                "log_number": log_number,
                "description": clean_text(
                    values[1]
                ),
                "bin_location": clean_text(
                    values[4]
                ),
                "available_weight": clean_decimal(
                    values[7],
                    "available weight",
                    log_number,
                ),
                "available_pieces": clean_integer(
                    values[8],
                    "available pieces",
                    log_number,
                ),
                "on_hand_weight": clean_decimal(
                    values[11],
                    "on-hand weight",
                    log_number,
                ),
                "on_hand_pieces": clean_integer(
                    values[12],
                    "on-hand pieces",
                    log_number,
                ),
            })
    finally:
        workbook.close()

    if not header_found:
        raise ReportUploadError(
            f"{upload.name} does not appear to be "
            f"a bin-location report."
        )

    if not rows:
        raise ReportUploadError(
            f"{upload.name} does not contain "
            f"warehouse {expected_warehouse} inventory."
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
    workbook, worksheet = open_report_worksheet(
        upload
    )

    remarks_by_log = defaultdict(list)
    header_found = False

    try:
        for values in worksheet.iter_rows(
            values_only=True
        ):
            values = list(values)

            values.extend(
                [None] * max(
                    0,
                    26 - len(values),
                )
            )

            if (
                normalize_log_number(values[4])
                == "LOG"
                and clean_text(values[23]).upper()
                == "REMARKS 1"
            ):
                header_found = True
                continue

            if not header_found:
                continue

            warehouse = clean_text(values[0])
            log_number = normalize_log_number(
                values[4]
            )

            if (
                warehouse not in {"20", "21"}
                or not log_number
            ):
                continue

            existing_lowercase = {
                remark.casefold()
                for remark
                in remarks_by_log[log_number]
            }

            for value in values[23:26]:
                remark = clean_text(value)

                if (
                    remark
                    and remark.casefold()
                    not in existing_lowercase
                ):
                    remarks_by_log[
                        log_number
                    ].append(remark)

                    existing_lowercase.add(
                        remark.casefold()
                    )
    finally:
        workbook.close()

    if not header_found:
        raise ReportUploadError(
            f"{upload.name} does not appear to be "
            f"a log-remarks report."
        )

    return remarks_by_log


def combine_reports(
    bin_report_20,
    bin_report_21,
    remarks_report,
):
    warehouse_20_rows = (
        read_bin_location_report(
            bin_report_20,
            expected_warehouse=20,
        )
    )

    warehouse_21_rows = (
        read_bin_location_report(
            bin_report_21,
            expected_warehouse=21,
        )
    )

    remarks_by_log = read_remarks_report(
        remarks_report
    )

    combined_rows = []

    for inventory_row in (
        warehouse_20_rows
        + warehouse_21_rows
    ):
        log_number = inventory_row["log_number"]

        combined_rows.append({
            **inventory_row,
            "remarks": " | ".join(
                remarks_by_log.get(
                    log_number,
                    [],
                )
            ),
        })

    combined_rows.sort(
        key=lambda row: (
            row["log_number"],
            row["warehouse"],
            row["bin_location"],
        )
    )

    return combined_rows


def row_value(row, field_name):
    if isinstance(row, dict):
        return row.get(field_name)

    return getattr(row, field_name)


def build_excel_report(rows):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Inventory Report"

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
            row_value(row, "log_number"),
            row_value(row, "description"),
            row_value(row, "bin_location"),
            row_value(row, "available_pieces"),
            row_value(row, "available_weight"),
            row_value(row, "on_hand_pieces"),
            row_value(row, "on_hand_weight"),
            row_value(row, "remarks"),
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
            vertical="center"
        )

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )
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

    for column, width in (
        column_widths.items()
    ):
        worksheet.column_dimensions[
            column
        ].width = width

    for row_number in range(
        2,
        worksheet.max_row + 1,
    ):
        worksheet[
            f"D{row_number}"
        ].number_format = "#,##0"

        worksheet[
            f"F{row_number}"
        ].number_format = "#,##0"

        worksheet[
            f"E{row_number}"
        ].number_format = "#,##0.00"

        worksheet[
            f"G{row_number}"
        ].number_format = "#,##0.00"

        worksheet[
            f"B{row_number}"
        ].alignment = Alignment(
            wrap_text=True,
            vertical="top",
        )

        worksheet[
            f"H{row_number}"
        ].alignment = Alignment(
            wrap_text=True,
            vertical="top",
        )

    output = BytesIO()

    workbook.save(output)
    workbook.close()

    output.seek(0)

    return output


def replace_saved_inventory(
    rows,
    request,
    uploads,
):
    """
    Replace the current inventory atomically.

    If creating the new inventory fails, the
    previous inventory remains in the database.
    """
    with transaction.atomic():
        InventoryReport.objects.all().delete()

        report = InventoryReport.objects.create(
            uploaded_by=(
                request.user.get_username()
            ),
            warehouse_20_filename=(
                uploads[0].name
            ),
            warehouse_21_filename=(
                uploads[1].name
            ),
            remarks_filename=uploads[2].name,
        )

        InventoryItem.objects.bulk_create(
            [
                InventoryItem(
                    report=report,
                    warehouse=row[
                        "warehouse"
                    ],
                    log_number=row[
                        "log_number"
                    ],
                    description=row[
                        "description"
                    ],
                    bin_location=row[
                        "bin_location"
                    ],
                    available_pieces=row[
                        "available_pieces"
                    ],
                    available_weight=row[
                        "available_weight"
                    ],
                    on_hand_pieces=row[
                        "on_hand_pieces"
                    ],
                    on_hand_weight=row[
                        "on_hand_weight"
                    ],
                    remarks=row["remarks"],
                )
                for row in rows
            ],
            batch_size=1000,
        )

    return report


def current_inventory_report():
    return (
        InventoryReport.objects
        .order_by("-uploaded_at")
        .first()
    )


@login_required
@require_http_methods(["GET", "POST"])
def reports(request):
    error = None

    if request.method == "POST":
        uploads = (
            request.FILES.get(
                "bin_report_20"
            ),
            request.FILES.get(
                "bin_report_21"
            ),
            request.FILES.get(
                "remarks_report"
            ),
        )

        if not all(uploads):
            error = (
                "Choose all three reports before "
                "replacing the inventory."
            )
        else:
            try:
                combined_rows = combine_reports(
                    *uploads
                )

                replace_saved_inventory(
                    combined_rows,
                    request,
                    uploads,
                )

            except ReportUploadError as exc:
                error = str(exc)

            else:
                messages.success(
                    request,
                    (
                        "Inventory replaced "
                        "successfully with "
                        f"{len(combined_rows):,} rows."
                    ),
                )

                return redirect("reports")

    current_report = (
        current_inventory_report()
    )

    if current_report:
        inventory_rows = (
            current_report.items.all()
        )
    else:
        inventory_rows = (
            InventoryItem.objects.none()
        )

    total_rows = inventory_rows.count()

    remarks_count = (
        inventory_rows
        .exclude(remarks="")
        .count()
    )

    return render(
        request,
        "core/reports.html",
        {
            "current_report": current_report,
            "combined_rows": inventory_rows,
            "total_rows": total_rows,
            "remarks_count": remarks_count,
            "error": error,
        },
    )


@login_required
@require_GET
def download_inventory_report(request):
    current_report = (
        current_inventory_report()
    )

    if current_report is None:
        messages.error(
            request,
            (
                "Upload an inventory report "
                "before downloading."
            ),
        )

        return redirect("reports")

    output = build_excel_report(
        current_report.items.all()
    )

    filename = (
        "inventory_report_"
        f"{timezone.localdate().isoformat()}"
        ".xlsx"
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

@login_required
@require_http_methods(["GET", "POST"])
def hymus_transfer(request):
    if request.method == "POST":
        action = request.POST.get("action", "")

        if action == "add_inventory":
            current_report = (
                current_inventory_report()
            )

            if current_report is None:
                messages.error(
                    request,
                    (
                        "Upload the Inventory Report "
                        "before selecting a log."
                    ),
                )

                return redirect(
                    "hymus_transfer"
                )

            inventory_item = get_object_or_404(
                InventoryItem,
                pk=request.POST.get(
                    "inventory_item_id"
                ),
                report=current_report,
            )

            _, created = (
                HymusTransferItem.objects
                .get_or_create(
                    warehouse=(
                        inventory_item.warehouse
                    ),
                    log_number=(
                        inventory_item.log_number
                    ),
                    bin_location=(
                        inventory_item.bin_location
                    ),
                    defaults={
                        "description": (
                            inventory_item.description
                        ),
                        "added_by": (
                            request.user
                            .get_username()
                        ),
                    },
                )
            )

            if created:
                messages.success(
                    request,
                    (
                        "Added log "
                        f"{inventory_item.log_number} "
                        "to the transfer."
                    ),
                )
            else:
                messages.info(
                    request,
                    (
                        "That inventory row is "
                        "already on the transfer."
                    ),
                )

        elif action == "add_manual":
            log_number = normalize_log_number(
                request.POST.get(
                    "manual_log_number",
                    "",
                )
            )

            description = clean_text(
                request.POST.get(
                    "manual_description",
                    "",
                )
            )

            bin_location = clean_text(
                request.POST.get(
                    "manual_bin_location",
                    "",
                )
            )

            if not all([
                log_number,
                description,
                bin_location,
            ]):
                messages.error(
                    request,
                    (
                        "Enter the log number, "
                        "description, and location."
                    ),
                )

                return redirect(
                    "hymus_transfer"
                )

            current_report = (
                current_inventory_report()
            )

            if (
                current_report
                and current_report.items.filter(
                    log_number__iexact=(
                        log_number
                    )
                ).exists()
            ):
                messages.warning(
                    request,
                    (
                        f"Log {log_number} exists "
                        "in inventory. Select it "
                        "from the search results "
                        "instead."
                    ),
                )

                return redirect(
                    "hymus_transfer"
                )

            _, created = (
                HymusTransferItem.objects
                .get_or_create(
                    warehouse="",
                    log_number=log_number,
                    bin_location=bin_location,
                    defaults={
                        "description": description,
                        "added_by": (
                            request.user
                            .get_username()
                        ),
                    },
                )
            )

            if created:
                messages.success(
                    request,
                    (
                        f"Added manual log "
                        f"{log_number} to the "
                        "transfer."
                    ),
                )
            else:
                messages.info(
                    request,
                    (
                        "That manual log and "
                        "location are already "
                        "on the transfer."
                    ),
                )

        else:
            messages.error(
                request,
                (
                    "Choose a log before "
                    "adding it."
                ),
            )

        return redirect("hymus_transfer")

    transfer_items = sorted(
        HymusTransferItem.objects.all(),
        key=bin_location_sort_key,
    )

    sheet_items = [
        item
        for item in transfer_items
        if is_sheet_material(item)
    ]

    bundle_items = [
        item
        for item in transfer_items
        if not is_sheet_material(item)
    ]

    transfer_documents = [
        {
            "key": "sheets",
            "title": "26 Hymus Sheets",
            "items": sheet_items,
            "count": len(sheet_items),
        },
        {
            "key": "bundles",
            "title": "26 Hymus Bundles",
            "items": bundle_items,
            "count": len(bundle_items),
        },
    ]

    return render(
        request,
        "core/hymus_transfer.html",
        {
            "transfer_items": transfer_items,
            "transfer_count": len(
                transfer_items
            ),
            "sheet_items": sheet_items,
            "sheet_count": len(
                sheet_items
            ),
            "bundle_items": bundle_items,
            "bundle_count": len(
                bundle_items
            ),
            "transfer_documents": (
                transfer_documents
            ),
            "current_report": (
                current_inventory_report()
            ),
            "printed_on": (
                timezone.localtime()
            ),
        },
    )


@login_required
@require_GET
def inventory_log_search(request):
    query = normalize_log_number(
        request.GET.get("q", "")
    )

    current_report = (
        current_inventory_report()
    )

    if (
        not query
        or current_report is None
    ):
        return JsonResponse({
            "results": [],
        })

    inventory_items = list(
        current_report.items
        .filter(
            log_number__icontains=query
        )
        .order_by(
            "log_number",
            "warehouse",
            "bin_location",
        )[:25]
    )

    existing_keys = set(
        HymusTransferItem.objects
        .values_list(
            "warehouse",
            "log_number",
            "bin_location",
        )
    )

    results = [
        {
            "id": item.id,
            "log_number": (
                item.log_number
            ),
            "description": (
                item.description
            ),
            "bin_location": (
                item.bin_location
            ),
            "warehouse": (
                item.warehouse
            ),
            "already_added": (
                item.warehouse,
                item.log_number,
                item.bin_location,
            ) in existing_keys,
        }
        for item in inventory_items
    ]

    return JsonResponse({
        "results": results,
    })

@login_required
@require_POST
def update_hymus_transfer_note(
    request,
    item_id,
):
    item = get_object_or_404(
        HymusTransferItem,
        pk=item_id,
    )

    item.notes = request.POST.get(
        "notes",
        "",
    ).strip()

    item.save(
        update_fields=[
            "notes",
            "updated_at",
        ]
    )

    messages.success(
        request,
        (
            f"Notes saved for log "
            f"{item.log_number}."
        ),
    )

    return redirect("hymus_transfer")


@login_required
@require_POST
def remove_hymus_transfer_item(
    request,
    item_id,
):
    item = get_object_or_404(
        HymusTransferItem,
        pk=item_id,
    )

    log_number = item.log_number

    item.delete()

    messages.success(
        request,
        (
            f"Removed log {log_number} "
            f"from the transfer."
        ),
    )

    return redirect("hymus_transfer")


@login_required
@require_POST
def clear_hymus_transfer(request):
    if (
        request.POST.get("confirm_clear")
        != "yes"
    ):
        messages.error(
            request,
            "The transfer was not cleared.",
        )

        return redirect("hymus_transfer")

    deleted_count, _ = (
        HymusTransferItem.objects
        .all()
        .delete()
    )

    messages.success(
        request,
        (
            "26 Hymus Transfer cleared "
            f"({deleted_count:,} rows removed)."
        ),
    )

    return redirect("hymus_transfer")