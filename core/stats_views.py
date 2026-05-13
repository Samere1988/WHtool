import datetime
import json
from collections import defaultdict

from django.contrib.auth.decorators import login_required
from django.db.models import Sum, Count
from django.shortcuts import render
from django.utils import timezone
from django.utils.dateparse import parse_date

from .models import OrderArchive, PickupLog, RunSheet


STATIONS = [
    ("bar", "Bar", "bar_prep", "bar_lines"),
    ("sheet", "Sheet", "sheet_prep", "sheet_lines"),
    ("covering", "Covering", "covering_prep", "covering_lines"),
]


def _safe_int(value):
    return value or 0


def _fmt(value):
    try:
        return f"{int(value or 0):,}"
    except (TypeError, ValueError):
        return value


def _add_fmt_fields(row, keys):
    for key in keys:
        row[f"{key}_fmt"] = _fmt(row.get(key, 0))
    return row


def _parse_period(request):
    now = timezone.localtime(timezone.now())
    active_tab = request.GET.get("tab", "day")
    if active_tab not in {"day", "week", "month"}:
        active_tab = "day"

    day_str = request.GET.get("day")
    week_str = request.GET.get("week")
    month_str = request.GET.get("month")

    target_day = parse_date(day_str) if day_str else now.date()
    if not target_day:
        target_day = now.date()

    day_start_date = target_day
    day_end_date = target_day
    day_start_dt = timezone.make_aware(datetime.datetime.combine(day_start_date, datetime.time.min))
    day_end_dt = timezone.make_aware(datetime.datetime.combine(day_end_date, datetime.time.max))

    if week_str and "-W" in week_str:
        year, week = int(week_str.split("-W")[0]), int(week_str.split("-W")[1])
        week_start_date = datetime.date.fromisocalendar(year, week, 1)
    else:
        week_start_date = now.date() - datetime.timedelta(days=now.weekday())
        year, week, _ = week_start_date.isocalendar()
        week_str = f"{year}-W{week:02d}"

    week_end_date = week_start_date + datetime.timedelta(days=6)
    week_start_dt = timezone.make_aware(datetime.datetime.combine(week_start_date, datetime.time.min))
    week_end_dt = timezone.make_aware(datetime.datetime.combine(week_end_date, datetime.time.max))

    if month_str and "-" in month_str:
        year, month = int(month_str.split("-")[0]), int(month_str.split("-")[1])
        month_start_date = datetime.date(year, month, 1)
    else:
        month_start_date = now.date().replace(day=1)
        month_str = month_start_date.strftime("%Y-%m")

    if month_start_date.month == 12:
        next_month_date = datetime.date(month_start_date.year + 1, 1, 1)
    else:
        next_month_date = datetime.date(month_start_date.year, month_start_date.month + 1, 1)
    month_end_date = next_month_date - datetime.timedelta(days=1)
    month_start_dt = timezone.make_aware(datetime.datetime.combine(month_start_date, datetime.time.min))
    month_end_dt = timezone.make_aware(datetime.datetime.combine(month_end_date, datetime.time.max))

    ranges = {
        "day": {
            "start_date": day_start_date,
            "end_date": day_end_date,
            "start_dt": day_start_dt,
            "end_dt": day_end_dt,
            "title": target_day.strftime("%B %d, %Y"),
        },
        "week": {
            "start_date": week_start_date,
            "end_date": week_end_date,
            "start_dt": week_start_dt,
            "end_dt": week_end_dt,
            "title": f"Week of {week_start_date.strftime('%B %d, %Y')}",
        },
        "month": {
            "start_date": month_start_date,
            "end_date": month_end_date,
            "start_dt": month_start_dt,
            "end_dt": month_end_dt,
            "title": month_start_date.strftime("%B %Y"),
        },
    }

    return {
        "active_tab": active_tab,
        "day_val": target_day.strftime("%Y-%m-%d"),
        "week_val": week_str,
        "month_val": month_str,
        "ranges": ranges,
        "active_range": ranges[active_tab],
    }


def _split_names(value):
    if not value:
        return []
    return [name.strip() for name in str(value).split(",") if name.strip()]


def _worker_dataset(start_dt, end_dt):
    orders = list(OrderArchive.objects.filter(created_at__range=(start_dt, end_dt)))
    pickups = list(PickupLog.objects.filter(date_completed__range=(start_dt.date(), end_dt.date())))

    stats = defaultdict(lambda: {
        "name": "",
        "total_orders": 0,
        "total_lines": 0,
        "bar_orders": 0,
        "bar_lines": 0,
        "sheet_orders": 0,
        "sheet_lines": 0,
        "covering_orders": 0,
        "covering_lines": 0,
    })

    for item in orders + pickups:
        for station_key, _label, prep_field, lines_field in STATIONS:
            names = _split_names(getattr(item, prep_field, ""))
            lines = _safe_int(getattr(item, lines_field, 0))
            for name in names:
                row = stats[name]
                row["name"] = name
                row["total_orders"] += 1
                row["total_lines"] += lines
                row[f"{station_key}_orders"] += 1
                row[f"{station_key}_lines"] += lines

    table = []
    for row in stats.values():
        row["avg_lines"] = round(row["total_lines"] / row["total_orders"], 1) if row["total_orders"] else 0
        _add_fmt_fields(row, [
            "total_orders", "total_lines", "bar_orders", "bar_lines",
            "sheet_orders", "sheet_lines", "covering_orders", "covering_lines",
        ])
        table.append(row)

    table.sort(key=lambda r: (r["total_lines"], r["total_orders"], r["name"]), reverse=True)

    total_orders = sum(row["total_orders"] for row in table)
    total_lines = sum(row["total_lines"] for row in table)
    summary = {
        "employees": len(table),
        "orders": total_orders,
        "lines": total_lines,
        "avg_lines": round(total_lines / total_orders, 1) if total_orders else 0,
        "archived_orders": len(orders),
        "pickup_orders": len(pickups),
    }
    _add_fmt_fields(summary, ["employees", "orders", "lines", "archived_orders", "pickup_orders"])

    return {
        "table": table,
        "labels": json.dumps([row["name"] for row in table]),
        "orders": json.dumps([row["total_orders"] for row in table]),
        "lines": json.dumps([row["total_lines"] for row in table]),
        "summary": summary,
    }


def _run_sheet_queryset(start_date, end_date):
    return RunSheet.objects.filter(shipping_date__range=(start_date, end_date))


def _run_summary(start_date, end_date):
    qs = _run_sheet_queryset(start_date, end_date)
    summary = qs.aggregate(
        total_weight=Sum("weight"),
        total_skids=Sum("skids"),
        total_bundles=Sum("bundles"),
        total_coils=Sum("coils"),
        total_orders=Count("id"),
    )
    summary = {key: _safe_int(value) for key, value in summary.items()}
    _add_fmt_fields(summary, ["total_weight", "total_skids", "total_bundles", "total_coils", "total_orders"])
    return summary


def _daily_run_rows(start_date, end_date):
    rows = []
    current = start_date
    while current <= end_date:
        totals = _run_summary(current, current)
        totals["date"] = current
        rows.append(totals)
        current += datetime.timedelta(days=1)
    return rows


def _region_rows(start_date, end_date):
    rows = list(
        _run_sheet_queryset(start_date, end_date)
        .values("region")
        .annotate(
            total_weight=Sum("weight"),
            total_skids=Sum("skids"),
            total_bundles=Sum("bundles"),
            total_coils=Sum("coils"),
            total_orders=Count("id"),
        )
        .order_by("-total_weight", "region")
    )
    for row in rows:
        row["region"] = row["region"] or "Unassigned"
        for key in ["total_weight", "total_skids", "total_bundles", "total_coils", "total_orders"]:
            row[key] = _safe_int(row[key])
        _add_fmt_fields(row, ["total_weight", "total_skids", "total_bundles", "total_coils", "total_orders"])
    return rows


def _customer_rows(start_date, end_date, sort_by="weight"):
    allowed_sort = {
        "orders": "-total_orders",
        "weight": "-total_weight",
        "skids": "-total_skids",
        "bundles": "-total_bundles",
        "coils": "-total_coils",
    }
    order_field = allowed_sort.get(sort_by, "-total_weight")

    rows = list(
        _run_sheet_queryset(start_date, end_date)
        .values("customer_id", "customer_name", "city")
        .annotate(
            total_weight=Sum("weight"),
            total_skids=Sum("skids"),
            total_bundles=Sum("bundles"),
            total_coils=Sum("coils"),
            total_orders=Count("id"),
        )
        .order_by(order_field, "customer_name")[:25]
    )
    for row in rows:
        row["customer_name"] = row["customer_name"] or "Unknown Customer"
        row["city"] = row["city"] or ""
        for key in ["total_weight", "total_skids", "total_bundles", "total_coils", "total_orders"]:
            row[key] = _safe_int(row[key])
        _add_fmt_fields(row, ["total_weight", "total_skids", "total_bundles", "total_coils", "total_orders"])
    return rows


def _run_dataset(period_range, customer_sort="weight"):
    start_date = period_range["start_date"]
    end_date = period_range["end_date"]
    daily_rows = _daily_run_rows(start_date, end_date)
    region_rows = _region_rows(start_date, end_date)
    customer_rows = _customer_rows(start_date, end_date, customer_sort)
    summary = _run_summary(start_date, end_date)

    return {
        "summary": summary,
        "daily_rows": daily_rows,
        "region_rows": region_rows,
        "customer_rows": customer_rows,
        "daily_labels": json.dumps([row["date"].strftime("%b %d") for row in daily_rows]),
        "daily_weights": json.dumps([row["total_weight"] for row in daily_rows]),
        "daily_orders": json.dumps([row["total_orders"] for row in daily_rows]),
        "region_labels": json.dumps([row["region"] for row in region_rows]),
        "region_weights": json.dumps([row["total_weight"] for row in region_rows]),
    }


@login_required
def stats_landing(request):
    today = timezone.localdate()
    week_start = today - datetime.timedelta(days=today.weekday())
    month_start = today.replace(day=1)

    worker_today = _worker_dataset(
        timezone.make_aware(datetime.datetime.combine(today, datetime.time.min)),
        timezone.make_aware(datetime.datetime.combine(today, datetime.time.max)),
    )
    run_today = _run_summary(today, today)
    run_week = _run_summary(week_start, today)
    run_month = _run_summary(month_start, today)

    recent_days = _daily_run_rows(today - datetime.timedelta(days=6), today)

    return render(request, "core/stats/landing.html", {
        "worker_today": worker_today,
        "run_today": run_today,
        "run_week": run_week,
        "run_month": run_month,
        "recent_days": recent_days,
        "recent_labels": json.dumps([row["date"].strftime("%b %d") for row in recent_days]),
        "recent_weights": json.dumps([row["total_weight"] for row in recent_days]),
    })


@login_required
def worker_stats(request):
    period = _parse_period(request)
    datasets = {
        key: _worker_dataset(value["start_dt"], value["end_dt"])
        for key, value in period["ranges"].items()
    }

    return render(request, "core/stats/worker_stats.html", {
        **period,
        "daily": datasets["day"],
        "weekly": datasets["week"],
        "monthly": datasets["month"],
        "active_dataset": datasets[period["active_tab"]],
        "active_title": period["active_range"]["title"],
    })


@login_required
def run_sheet_stats(request):
    period = _parse_period(request)
    customer_sort = request.GET.get("customer_sort", "weight")
    datasets = {
        key: _run_dataset(value, customer_sort)
        for key, value in period["ranges"].items()
    }

    return render(request, "core/stats/run_sheet_stats.html", {
        **period,
        "customer_sort": customer_sort,
        "daily": datasets["day"],
        "weekly": datasets["week"],
        "monthly": datasets["month"],
        "active_dataset": datasets[period["active_tab"]],
        "active_title": period["active_range"]["title"],
    })
