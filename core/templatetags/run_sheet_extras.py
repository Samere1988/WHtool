from django import template

register = template.Library()


@register.filter
def comma(value):
    try:
        return format(int(value or 0), ",")
    except (TypeError, ValueError):
        return value


@register.filter
def run_totals(grouped_orders):
    totals = {
        "weight": 0,
        "skids": 0,
        "bundles": 0,
        "coils": 0,
        "orders": 0,
    }

    if not grouped_orders:
        return totals

    for region_data in grouped_orders.values():
        region_totals = region_data.get("totals", {})
        totals["weight"] += int(region_totals.get("weight") or 0)
        totals["skids"] += int(region_totals.get("skids") or 0)
        totals["bundles"] += int(region_totals.get("bundles") or 0)
        totals["coils"] += int(region_totals.get("coils") or 0)
        totals["orders"] += len(region_data.get("orders", []))

    return totals


@register.filter
def get_item(dictionary, key):
    if not dictionary:
        return ""
    return dictionary.get(key, "")
