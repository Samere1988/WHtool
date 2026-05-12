from copy import copy
from io import BytesIO

from django.contrib.auth.decorators import login_required

from openpyxl import load_workbook

from .models import RunSheet
from . import views as base_views


def _metadata_col_for_block(block):
    # Existing layout uses AA/AB hidden columns under hidden_ids_col.
    # Keep this only to clear old hidden metadata so stale IDs cannot remain behind.
    return block.get("customer_code_col") or block.get("hidden_ids_col")


def _customer_code_for_stop(stop):
    ids = stop.get("ids") or []
    if not ids:
        return ""

    try:
        first_id = int(ids[0])
    except (TypeError, ValueError):
        return ""

    order = RunSheet.objects.filter(id=first_id).only("customer_id").first()
    return order.customer_id if order else ""


def _copy_cell(source, target):
    target.value = source.value
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.alignment = copy(source.alignment)
    target.number_format = source.number_format
    target.protection = copy(source.protection)


def _shift_block_right_for_code_column(ws, block):
    """
    Converts a generated 9-column transport block:
        Customer Name | City | Weight | Skids | Bundles | Coils | Closes at | Pickup | blank

    into a 10-column transport block:
        Code | Customer Name | City | Weight | Skids | Bundles | Coils | Closes at | Pickup | blank

    This keeps the code visible in the Excel file so it moves naturally when rows are copied/pasted.
    """
    start_col = block["start_col"]

    for row_num in range(block["header_row"], block["total_row"] + 1):
        for col in range(start_col + 8, start_col - 1, -1):
            _copy_cell(
                ws.cell(row=row_num, column=col),
                ws.cell(row=row_num, column=col + 1),
            )

    # The code column uses the old customer-name column style.
    code_header = ws.cell(row=block["header_row"], column=start_col)
    code_header.value = "Code"

    for row_num in range(block["start_row"], block["end_row"] + 1):
        ws.cell(row=row_num, column=start_col).value = None


def _write_customer_codes(workbook, shipping_date):
    """
    Calls the existing export first, then adds a visible Code column to each region block.
    Old hidden AA/AB metadata is cleared so stale row IDs cannot affect future imports.
    """
    for ws in workbook.worksheets:
        for region, block in base_views.TRANSPORT_REGION_BLOCKS.items():
            _shift_block_right_for_code_column(ws, block)

            metadata_col = _metadata_col_for_block(block)
            if metadata_col:
                for row_num in range(block["start_row"], block["end_row"] + 1):
                    ws.cell(row=row_num, column=metadata_col).value = None

            stops = base_views.build_transport_stops_for_region(region, shipping_date)
            for index, stop in enumerate(stops):
                row_num = block["start_row"] + index
                if row_num > block["end_row"]:
                    break
                ws.cell(row=row_num, column=block["start_col"]).value = _customer_code_for_stop(stop)

            # Rebuild totals because formulas copied by openpyxl do not automatically shift references.
            total_row = block["total_row"]
            start_row = block["start_row"]
            end_row = block["end_row"]
            start_col = block["start_col"]

            weight_col = start_col + 3
            skids_col = start_col + 4
            bundles_col = start_col + 5
            coils_col = start_col + 6

            ws.cell(row=total_row, column=weight_col).value = (
                f"=SUM({ws.cell(start_row, weight_col).coordinate}:{ws.cell(end_row, weight_col).coordinate})"
            )
            ws.cell(row=total_row, column=skids_col).value = (
                f"=SUM({ws.cell(start_row, skids_col).coordinate}:{ws.cell(end_row, skids_col).coordinate})"
            )
            ws.cell(row=total_row, column=bundles_col).value = (
                f"=SUM({ws.cell(start_row, bundles_col).coordinate}:{ws.cell(end_row, bundles_col).coordinate})"
            )
            ws.cell(row=total_row, column=coils_col).value = (
                f"=SUM({ws.cell(start_row, coils_col).coordinate}:{ws.cell(end_row, coils_col).coordinate})"
            )

    for ws in workbook.worksheets:
        # Keep old metadata columns hidden/empty for backwards safety, but new matching uses visible Code.
        ws.column_dimensions["AA"].hidden = True
        ws.column_dimensions["AB"].hidden = True
        ws.print_area = "A1:T75"

        # Make the visible Code columns usable without making them huge.
        ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 12)
        ws.column_dimensions["K"].width = max(ws.column_dimensions["K"].width or 0, 12)


@login_required
def export_run_sheet_excel(request):
    """
    Wrapper around the original Excel export.

    It preserves the existing generated workbook/layout, then adds a visible Code column
    at the start of each transport region block. This avoids hidden-cell mismatch issues
    when the transport sheet is copy/pasted or rearranged.
    """
    shipping_date = base_views.get_selected_shipping_date(request)
    response = base_views.export_run_sheet_excel(request)

    workbook = load_workbook(BytesIO(response.content))
    _write_customer_codes(workbook, shipping_date)

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    response.content = output.getvalue()
    return response
