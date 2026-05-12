import re
import unicodedata
from difflib import SequenceMatcher
from collections import defaultdict

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.utils.dateparse import parse_date

from openpyxl import load_workbook

from .models import RunSheet, TransportImportBatch, TransportImportRow, TransportImportPreviousState
from .views import TRANSPORT_REGION_BLOCKS, TRANSPORT_REGIONS, get_selected_shipping_date, redirect_run_sheet_for_date


LEGAL_WORDS = {
    "INC", "INCORPORATED", "LTD", "LTEE", "LTÉE", "LIMITED", "CORP", "CORPORATION",
    "CO", "COMPANY", "THE", "LES", "LE", "LA", "DES", "DE", "DU", "DISTRIBUTION",
}


def clean_match_text(value):
    if value is None:
        return ""
    text = str(value).strip().upper()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^A-Z0-9 ]+", " ", text)
    parts = [p for p in text.split() if p not in LEGAL_WORDS]
    return " ".join(parts)


def score_match(imported_name, imported_city, order):
    name_score = SequenceMatcher(None, clean_match_text(imported_name), clean_match_text(order.customer_name)).ratio()
    city_score = SequenceMatcher(None, clean_match_text(imported_city), clean_match_text(order.city)).ratio()

    if imported_city:
        return round((name_score * 0.8) + (city_score * 0.2), 3)
    return round(name_score, 3)


def stop_group_key(order):
    return (
        order.customer_id or "",
        order.customer_name or "",
        order.city or "",
        order.address or "",
        order.postal_code or "",
        order.closing_time or "",
        bool(order.is_pickup),
        bool(order.is_return),
    )


def build_current_stop_options(shipping_date):
    grouped = {}
    orders = RunSheet.objects.filter(shipping_date=shipping_date).order_by("region", "load_index", "customer_name", "id")

    for order in orders:
        key = stop_group_key(order)
        if key not in grouped:
            grouped[key] = {
                "ids": [],
                "label": f"{order.customer_name or ''} — {order.city or ''} — {order.region or ''}",
                "customer_name": order.customer_name or "",
                "city": order.city or "",
                "region": order.region or "",
                "order_numbers": [],
            }
        grouped[key]["ids"].append(str(order.id))
        if order.order_number:
            grouped[key]["order_numbers"].append(order.order_number)

    options = []
    for item in grouped.values():
        order_refs = " / ".join(item["order_numbers"])
        label = item["label"]
        if order_refs:
            label = f"{label} ({order_refs})"
        options.append({"value": ",".join(item["ids"]), "label": label})

    return sorted(options, key=lambda x: x["label"])


def ids_from_cell(value):
    if value in (None, ""):
        return []
    text = str(value)
    return [int(x) for x in re.findall(r"\d+", text)]


def parse_website_export(file_obj):
    wb = load_workbook(file_obj, data_only=True)
    parsed_rows = []

    for ws in wb.worksheets:
        for region, block in TRANSPORT_REGION_BLOCKS.items():
            driver = ws[block["driver_cell"]].value or ""
            stop_no = 1

            for row_num in range(block["start_row"], block["end_row"] + 1):
                start_col = block["start_col"]
                customer_name = ws.cell(row=row_num, column=start_col).value
                city = ws.cell(row=row_num, column=start_col + 1).value
                hidden_ids = ids_from_cell(ws.cell(row=row_num, column=block["hidden_ids_col"]).value)

                # Ignore empty rows.
                if not customer_name and not city and not hidden_ids:
                    continue

                parsed_rows.append({
                    "sheet_name": ws.title,
                    "source_row_number": row_num,
                    "imported_run_name": region,
                    "imported_driver": str(driver).strip(),
                    "imported_truck": "",
                    "imported_stop_number": stop_no,
                    "imported_customer_name": str(customer_name or "").strip(),
                    "imported_city": str(city or "").strip(),
                    "hidden_ids": hidden_ids,
                })
                stop_no += 1

    return parsed_rows


def auto_match_row(row_data, shipping_date):
    hidden_ids = row_data.get("hidden_ids") or []
    if hidden_ids:
        valid_ids = list(
            RunSheet.objects.filter(id__in=hidden_ids, shipping_date=shipping_date)
            .values_list("id", flat=True)
        )
        if valid_ids:
            return valid_ids, 1.0, "matched"

    imported_name = row_data.get("imported_customer_name") or ""
    imported_city = row_data.get("imported_city") or ""

    candidates = RunSheet.objects.filter(shipping_date=shipping_date)
    if not candidates.exists() or not imported_name:
        return [], 0, "unmatched"

    best_order = None
    best_score = 0
    for order in candidates:
        score = score_match(imported_name, imported_city, order)
        if score > best_score:
            best_score = score
            best_order = order

    if not best_order:
        return [], 0, "unmatched"

    key = stop_group_key(best_order)
    matched_ids = [o.id for o in candidates if stop_group_key(o) == key]
    status = "matched" if best_score >= 0.82 else "review" if best_score >= 0.58 else "unmatched"
    return matched_ids, best_score, status


@login_required
def upload_transport_import(request):
    if request.method != "POST":
        return redirect_run_sheet_for_date(get_selected_shipping_date(request))

    shipping_date = get_selected_shipping_date(request)
    excel_file = request.FILES.get("excel_file")

    if not excel_file:
        messages.error(request, "Please choose an Excel file first.")
        return redirect_run_sheet_for_date(shipping_date)

    batch = TransportImportBatch.objects.create(
        shipping_date=shipping_date,
        original_filename=excel_file.name,
        uploaded_by=request.user.username if request.user.is_authenticated else "",
        status="review",
    )

    try:
        parsed_rows = parse_website_export(excel_file)
    except Exception as exc:
        batch.status = "failed"
        batch.notes = str(exc)
        batch.save(update_fields=["status", "notes"])
        messages.error(request, f"Could not read that Excel file: {exc}")
        return redirect_run_sheet_for_date(shipping_date)

    if not parsed_rows:
        batch.status = "failed"
        batch.notes = "No stops were detected in the uploaded Excel file."
        batch.save(update_fields=["status", "notes"])
        messages.error(request, "No stops were detected in the uploaded Excel file.")
        return redirect_run_sheet_for_date(shipping_date)

    for sort_order, row_data in enumerate(parsed_rows, start=1):
        matched_ids, confidence, status = auto_match_row(row_data, shipping_date)
        TransportImportRow.objects.create(
            batch=batch,
            sort_order=sort_order,
            source_sheet_name=row_data["sheet_name"],
            source_row_number=row_data["source_row_number"],
            imported_run_name=row_data["imported_run_name"],
            imported_driver=row_data["imported_driver"],
            imported_truck=row_data["imported_truck"],
            imported_stop_number=row_data["imported_stop_number"],
            imported_customer_name=row_data["imported_customer_name"],
            imported_city=row_data["imported_city"],
            matched_run_sheet_ids=",".join(str(i) for i in matched_ids),
            confidence=confidence,
            status=status,
        )

    messages.info(request, "Transport sheet uploaded. Review the matches before applying the order.")
    return redirect("review_transport_import", batch_id=batch.id)


@login_required
def review_transport_import(request, batch_id):
    batch = get_object_or_404(TransportImportBatch, pk=batch_id)
    rows = batch.rows.all().order_by("sort_order", "id")
    options = build_current_stop_options(batch.shipping_date)

    if request.method == "POST":
        for row in rows:
            selected_ids = request.POST.get(f"match_{row.id}", "").strip()
            row.matched_run_sheet_ids = selected_ids
            row.status = "matched" if selected_ids else "unmatched"
            row.save(update_fields=["matched_run_sheet_ids", "status"])

        if request.POST.get("apply_now") == "1":
            return redirect("apply_transport_import", batch_id=batch.id)

        messages.success(request, "Matches saved.")
        return redirect("review_transport_import", batch_id=batch.id)

    matched_count = rows.filter(status="matched").count()
    review_count = rows.filter(status="review").count()
    unmatched_count = rows.exclude(status__in=["matched", "review"]).count()

    return render(request, "core/transport_import_review.html", {
        "batch": batch,
        "rows": rows,
        "options": options,
        "matched_count": matched_count,
        "review_count": review_count,
        "unmatched_count": unmatched_count,
    })


@login_required
def apply_transport_import(request, batch_id):
    batch = get_object_or_404(TransportImportBatch, pk=batch_id)
    rows = batch.rows.exclude(matched_run_sheet_ids="").order_by("sort_order", "id")

    if not rows.exists():
        messages.error(request, "There are no matched rows to apply.")
        return redirect("review_transport_import", batch_id=batch.id)

    with transaction.atomic():
        TransportImportPreviousState.objects.filter(batch=batch).delete()
        touched_ids = set()

        for row in rows:
            ids = [int(x) for x in row.matched_run_sheet_ids.split(",") if x.strip().isdigit()]
            for run_item in RunSheet.objects.select_for_update().filter(id__in=ids, shipping_date=batch.shipping_date):
                if run_item.id not in touched_ids:
                    TransportImportPreviousState.objects.create(
                        batch=batch,
                        run_sheet_id=run_item.id,
                        previous_transport_run_name=run_item.transport_run_name,
                        previous_transport_driver=run_item.transport_driver,
                        previous_transport_truck=run_item.transport_truck,
                        previous_transport_stop_number=run_item.transport_stop_number,
                        previous_transport_import_batch_id=run_item.transport_import_batch_id,
                        previous_driver_name=run_item.driver_name,
                        previous_load_index=run_item.load_index,
                    )
                    touched_ids.add(run_item.id)

                run_item.transport_run_name = row.imported_run_name or run_item.region or "Transport Run"
                run_item.transport_driver = row.imported_driver or ""
                run_item.transport_truck = row.imported_truck or ""
                run_item.transport_stop_number = row.imported_stop_number or row.sort_order
                run_item.transport_import_batch = batch
                if row.imported_driver:
                    run_item.driver_name = row.imported_driver
                run_item.load_index = row.imported_stop_number or row.sort_order
                run_item.save(update_fields=[
                    "transport_run_name", "transport_driver", "transport_truck",
                    "transport_stop_number", "transport_import_batch", "driver_name", "load_index",
                ])

        batch.status = "applied"
        batch.applied_at = timezone.now()
        batch.save(update_fields=["status", "applied_at"])

    messages.success(request, f"Transport order applied. {len(touched_ids)} run sheet rows were updated. You can undo this import if needed.")
    return redirect("transport_run_sheet_view", batch_id=batch.id)


@login_required
def undo_transport_import(request, batch_id):
    batch = get_object_or_404(TransportImportBatch, pk=batch_id)

    if request.method != "POST":
        return redirect("transport_import_history")

    states = TransportImportPreviousState.objects.filter(batch=batch)
    if not states.exists():
        messages.error(request, "No saved previous state was found for this import.")
        return redirect("transport_import_history")

    with transaction.atomic():
        for state in states:
            RunSheet.objects.filter(id=state.run_sheet_id).update(
                transport_run_name=state.previous_transport_run_name,
                transport_driver=state.previous_transport_driver,
                transport_truck=state.previous_transport_truck,
                transport_stop_number=state.previous_transport_stop_number,
                transport_import_batch_id=state.previous_transport_import_batch_id,
                driver_name=state.previous_driver_name,
                load_index=state.previous_load_index,
            )

        batch.status = "undone"
        batch.undone_at = timezone.now()
        batch.save(update_fields=["status", "undone_at"])

    messages.success(request, "Transport import was undone and the previous order was restored.")
    return redirect_run_sheet_for_date(batch.shipping_date)


@login_required
def transport_import_history(request):
    selected_shipping_date = get_selected_shipping_date(request)
    batches = TransportImportBatch.objects.filter(shipping_date=selected_shipping_date).order_by("-created_at")
    return render(request, "core/transport_import_history.html", {
        "batches": batches,
        "selected_shipping_date": selected_shipping_date,
    })


@login_required
def transport_run_sheet_view(request, batch_id=None):
    selected_shipping_date = get_selected_shipping_date(request)

    if batch_id:
        batch = get_object_or_404(TransportImportBatch, pk=batch_id)
    else:
        batch = TransportImportBatch.objects.filter(
            shipping_date=selected_shipping_date,
            status="applied",
        ).order_by("-applied_at", "-created_at").first()

    if not batch:
        messages.warning(request, "No applied transport import exists for this shipping date yet.")
        return redirect_run_sheet_for_date(selected_shipping_date)

    orders = RunSheet.objects.filter(
        shipping_date=batch.shipping_date,
        transport_import_batch=batch,
    ).order_by("transport_run_name", "transport_driver", "transport_stop_number", "customer_name", "order_number", "id")

    grouped = defaultdict(lambda: {"orders": [], "totals": {"weight": 0, "skids": 0, "bundles": 0, "coils": 0}})
    for order in orders:
        run_label = order.transport_run_name or order.region or "Transport Run"
        if order.transport_driver:
            run_label = f"{run_label} — {order.transport_driver}"
        grouped[run_label]["orders"].append(order)
        grouped[run_label]["totals"]["weight"] += order.weight or 0
        grouped[run_label]["totals"]["skids"] += order.skids or 0
        grouped[run_label]["totals"]["bundles"] += order.bundles or 0
        grouped[run_label]["totals"]["coils"] += order.coils or 0

    return render(request, "core/transport_run_sheet.html", {
        "batch": batch,
        "grouped_runs": dict(grouped),
        "selected_shipping_date": batch.shipping_date,
    })
