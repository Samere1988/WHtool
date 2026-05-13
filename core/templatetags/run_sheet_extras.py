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
    """
    Totals for the small top summary on the run sheet.
    Only regular delivery orders are counted. Pickups and returns are excluded.
    """
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
        for stop in region_data.get("orders", []):
            if stop.get("is_pickup") or stop.get("is_return"):
                continue

            totals["weight"] += int(stop.get("weight") or 0)
            totals["skids"] += int(stop.get("skids") or 0)
            totals["bundles"] += int(stop.get("bundles") or 0)
            totals["coils"] += int(stop.get("coils") or 0)
            totals["orders"] += len(stop.get("orders", [])) or 1

    return totals


@register.filter
def get_item(dictionary, key):
    if not dictionary:
        return ""
    return dictionary.get(key, "")
