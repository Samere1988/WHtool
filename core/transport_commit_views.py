from collections import defaultdict

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.shortcuts import get_object_or_404, redirect, render

from .models import RunSheet, TransportImportBatch
from .transport_import_views import redirect_run_sheet_for_date


def _transport_stop_key(order):
    """
    Groups individual RunSheet rows into one visible stop for the transport preview.
    This mirrors the main run sheet grouping behaviour, while preserving the imported
    transport run order.
    """
    return (
        order.customer_id or "",
        order.customer_name or "",
        order.city or "",
        order.address or "",
        order.postal_code or "",
        order.closing_time or "",
        bool(order.is_pickup),
        bool(order.is_return),
        order.transport_stop_number or order.load_index or 0,
    )


def _group_orders_into_stops(orders):
    grouped = {}

    for order in orders:
        key = _transport_stop_key(order)

        if key not in grouped:
            grouped[key] = {
                "main_order": order,
                "orders": [],
                "customer_name": order.customer_name or "",
                "city": order.city or "",
                "region": order.region or "",
                "transport_stop_number": order.transport_stop_number or order.load_index or 0,
                "closing_time": order.closing_time or "",
                "is_pickup": bool(order.is_pickup),
                "is_return": bool(order.is_return),
                "weight": 0,
                "skids": 0,
                "bundles": 0,
                "coils": 0,
            }

        grouped[key]["orders"].append(order)
        grouped[key]["weight"] += order.weight or 0
        grouped[key]["skids"] += order.skids or 0
        grouped[key]["bundles"] += order.bundles or 0
        grouped[key]["coils"] += order.coils or 0

        current_stop = grouped[key]["transport_stop_number"] or 0
        order_stop = order.transport_stop_number or order.load_index or 0
        if order_stop and (not current_stop or order_stop < current_stop):
            grouped[key]["transport_stop_number"] = order_stop
            grouped[key]["main_order"] = order

    stops = list(grouped.values())
    stops.sort(key=lambda stop: (stop["transport_stop_number"], stop["customer_name"], stop["main_order"].id))

    for stop in stops:
        stop["orders"].sort(key=lambda order: (order.order_number or "", order.id))

    return stops


@login_required
def transport_preview_view(request, batch_id=None):
    if batch_id:
        batch = get_object_or_404(TransportImportBatch, pk=batch_id)
    else:
        # Use latest applied/review batch as the preview for the selected date.
        from .views import get_selected_shipping_date

        selected_shipping_date = get_selected_shipping_date(request)
        batch = TransportImportBatch.objects.filter(
            shipping_date=selected_shipping_date,
            status__in=["review", "applied"],
        ).order_by("-applied_at", "-created_at").first()

    if not batch:
        messages.warning(request, "No transport import preview exists for this shipping date yet.")
        from .views import get_selected_shipping_date
        return redirect_run_sheet_for_date(get_selected_shipping_date(request))

    if request.method == "POST":
        original_run_name = request.POST.get("original_run_name", "").strip()
        original_driver = request.POST.get("original_driver", "").strip()
        original_start_time = request.POST.get("original_start_time", "").strip()
        new_driver = request.POST.get("transport_driver", "").strip()
        new_start_time = request.POST.get("transport_start_time", "").strip()

        updated = RunSheet.objects.filter(
            shipping_date=batch.shipping_date,
            transport_import_batch=batch,
            transport_run_name=original_run_name,
            transport_driver=original_driver,
            transport_start_time=original_start_time,
        ).update(
            transport_driver=new_driver,
            transport_start_time=new_start_time,
            driver_name=new_driver,
        )

        messages.success(request, f"Updated driver/start time for {updated} rows.")
        return redirect("transport_run_sheet_view", batch_id=batch.id)

    orders = RunSheet.objects.filter(
        shipping_date=batch.shipping_date,
        transport_import_batch=batch,
    ).order_by(
        "transport_run_name",
        "transport_driver",
        "transport_start_time",
        "transport_stop_number",
        "customer_name",
        "order_number",
        "id",
    )

    grouped = defaultdict(lambda: {
        "orders": [],
        "stops": [],
        "transport_run_name": "",
        "transport_driver": "",
        "transport_start_time": "",
        "totals": {"weight": 0, "skids": 0, "bundles": 0, "coils": 0},
    })

    run_order_buckets = defaultdict(list)

    for order in orders:
        run_name = order.transport_run_name or order.region or "Transport Run"
        driver = order.transport_driver or ""
        start_time = order.transport_start_time or ""
        group_key = f"{run_name}|{driver}|{start_time}"

        grouped[group_key]["transport_run_name"] = run_name
        grouped[group_key]["transport_driver"] = driver
        grouped[group_key]["transport_start_time"] = start_time
        grouped[group_key]["orders"].append(order)
        grouped[group_key]["totals"]["weight"] += order.weight or 0
        grouped[group_key]["totals"]["skids"] += order.skids or 0
        grouped[group_key]["totals"]["bundles"] += order.bundles or 0
        grouped[group_key]["totals"]["coils"] += order.coils or 0
        run_order_buckets[group_key].append(order)

    for group_key, run_orders in run_order_buckets.items():
        grouped[group_key]["stops"] = _group_orders_into_stops(run_orders)

    return render(request, "core/transport_run_sheet.html", {
        "batch": batch,
        "grouped_runs": dict(grouped),
        "selected_shipping_date": batch.shipping_date,
    })


@login_required
def save_transport_import_to_run_sheet(request, batch_id):
    """
    Commits the transport preview into the main RunSheet layout.

    This is the point where the imported transport runs overwrite the normal run sheet
    region/order. Until this button is clicked, the main run sheet remains unchanged.
    """
    if request.method != "POST":
        return redirect("transport_run_sheet_view", batch_id=batch_id)

    batch = get_object_or_404(TransportImportBatch, pk=batch_id)

    rows = batch.rows.all().order_by("sort_order", "id")

    with transaction.atomic():
        updated_count = 0

        for row in rows:
            ids = [int(x) for x in row.matched_run_sheet_ids.split(",") if x.strip().isdigit()]
            if not ids:
                continue

            run_items = RunSheet.objects.select_for_update().filter(
                id__in=ids,
                shipping_date=batch.shipping_date,
            )

            for run_item in run_items:
                run_item.region = row.imported_run_name or run_item.region
                run_item.driver_name = row.imported_driver or run_item.driver_name or ""
                run_item.load_index = row.imported_stop_number or row.sort_order
                run_item.transport_run_name = row.imported_run_name or run_item.transport_run_name
                run_item.transport_driver = row.imported_driver or run_item.transport_driver or ""
                run_item.transport_start_time = row.imported_start_time or run_item.transport_start_time or ""
                run_item.transport_stop_number = row.imported_stop_number or row.sort_order
                run_item.transport_import_batch = batch
                run_item.save(update_fields=[
                    "region",
                    "driver_name",
                    "load_index",
                    "transport_run_name",
                    "transport_driver",
                    "transport_start_time",
                    "transport_stop_number",
                    "transport_import_batch",
                ])
                updated_count += 1

        batch.status = "applied"
        batch.applied_at = batch.applied_at or __import__("django.utils.timezone").utils.timezone.now()
        batch.save(update_fields=["status", "applied_at"])

    messages.success(request, f"Run sheet saved. {updated_count} rows were updated.")
    return redirect_run_sheet_for_date(batch.shipping_date)
