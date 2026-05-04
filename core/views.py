import pandas as pd
import json
import datetime
from datetime import date, timedelta
from collections import Counter

from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.contrib import messages
from django.db.models import Sum, Count, Max
from django.utils import timezone, dateparse
from django.utils.dateparse import parse_date

from openpyxl.styles import Font, PatternFill

from .models import (
    CustomerList, RunSheet, OrderArchive, FinalizedRunSheet,
    OutboundLoad, OutboundPhoto, Vendor, PickupLog,
    PickupPhotoLog, PickupPhoto, Container, ContainerPhoto
)


# ==========================================
# --- 1. HELPER FUNCTIONS ---
# ==========================================

def get_next_business_day():
    """Calculates the next shipping day, skipping weekends."""
    today = date.today()
    if today.weekday() == 4:  # Friday -> Monday
        days_to_add = 3
    elif today.weekday() == 5:  # Saturday -> Monday
        days_to_add = 2
    else:  # Sun-Thu -> Next Day
        days_to_add = 1
    return today + timedelta(days=days_to_add)


def calculate_performance(orders_queryset):
    """Calculates picker KPIs for the Stats page (Chart.js version)"""
    picker_counts = Counter()
    line_counts = Counter()

    for o in orders_queryset:
        if o.bar_prep:
            names = [n.strip() for n in o.bar_prep.split(',') if n.strip()]
            for name in names:
                picker_counts[name] += 1
                line_counts[name] += (o.bar_lines or 0)
        if o.sheet_prep:
            names = [n.strip() for n in o.sheet_prep.split(',') if n.strip()]
            for name in names:
                picker_counts[name] += 1
                line_counts[name] += (o.sheet_lines or 0)
        if o.covering_prep:
            names = [n.strip() for n in o.covering_prep.split(',') if n.strip()]
            for name in names:
                picker_counts[name] += 1
                line_counts[name] += (o.covering_lines or 0)

    performance = []
    labels = []
    orders_data = []
    lines_data = []

    sorted_names = sorted(picker_counts.keys(), key=lambda n: line_counts[n], reverse=True)

    for name in sorted_names:
        tot_orders = picker_counts[name]
        tot_lines = line_counts[name]
        avg = round(tot_lines / tot_orders, 1) if tot_orders > 0 else 0

        performance.append({
            'prepared_by': name, 'total_orders': tot_orders,
            'total_lines': tot_lines, 'avg_lines': avg
        })
        labels.append(name)
        orders_data.append(tot_orders)
        lines_data.append(tot_lines)

    return {
        'table': performance, 'labels': json.dumps(labels),
        'orders': json.dumps(orders_data), 'lines': json.dumps(lines_data)
    }


# ==========================================
# --- 2. MAIN DASHBOARDS & STATS ---
# ==========================================

@login_required
def home(request):
    return render(request, 'core/home.html')


@login_required
def run_sheet(request):
    """The main dashboard grouped by Truck/Region, sorted by load sequence."""
    regions = ["North Shore", "Drummond", "Quebec", "Beauce", "Montreal", "South Shore", "Ontario", "Sherbrooke"]
    grouped_orders = {}
    for region in regions:
        orders = RunSheet.objects.filter(region=region).order_by('load_index')
        if orders.exists():
            grouped_orders[region] = {
                'orders': orders,
                'totals': {
                    'weight': sum(o.weight or 0 for o in orders),
                    'skids': sum(o.skids or 0 for o in orders),
                    'bundles': sum(o.bundles or 0 for o in orders),
                    'coils': sum(o.coils or 0 for o in orders),
                }
            }
    return render(request, 'core/run_sheet.html', {'grouped_orders': grouped_orders, 'regions': regions})


@login_required
def stats(request):
    """KPIs and Staff Performance Tracking"""
    now = timezone.localtime(timezone.now())
    active_tab = request.GET.get('tab', 'day')

    day_str = request.GET.get('day')
    week_str = request.GET.get('week')
    month_str = request.GET.get('month')

    # DAILY
    target_day = parse_date(day_str) if day_str else now.date()
    day_start = timezone.make_aware(datetime.datetime.combine(target_day, datetime.time.min))
    day_end = timezone.make_aware(datetime.datetime.combine(target_day, datetime.time.max))

    # WEEKLY
    if week_str and '-W' in week_str:
        year, week = int(week_str.split('-W')[0]), int(week_str.split('-W')[1])
        week_start_date = datetime.date.fromisocalendar(year, week, 1)
    else:
        week_start_date = now.date() - timedelta(days=now.weekday())
        year, week, _ = week_start_date.isocalendar()
        week_str = f"{year}-W{week:02d}"

    week_start = timezone.make_aware(datetime.datetime.combine(week_start_date, datetime.time.min))
    week_end = timezone.make_aware(datetime.datetime.combine(week_start_date + timedelta(days=6), datetime.time.max))

    # MONTHLY
    if month_str and '-' in month_str:
        year, month = int(month_str.split('-')[0]), int(month_str.split('-')[1])
        month_start_date = datetime.date(year, month, 1)
    else:
        month_start_date = now.date().replace(day=1)
        month_str = month_start_date.strftime('%Y-%m')

    month_start = timezone.make_aware(datetime.datetime.combine(month_start_date, datetime.time.min))
    if month_start.month == 12:
        next_month = month_start.replace(year=month_start.year + 1, month=1)
    else:
        next_month = month_start.replace(month=month_start.month + 1)
    month_end = next_month - timedelta(seconds=1)

    # Queries
    orders_day = list(OrderArchive.objects.filter(created_at__range=(day_start, day_end)))
    pickups_day = list(PickupLog.objects.filter(date_completed__range=(day_start.date(), day_end.date())))

    orders_week = list(OrderArchive.objects.filter(created_at__range=(week_start, week_end)))
    pickups_week = list(PickupLog.objects.filter(date_completed__range=(week_start.date(), week_end.date())))

    orders_month = list(OrderArchive.objects.filter(created_at__range=(month_start, month_end)))
    pickups_month = list(PickupLog.objects.filter(date_completed__range=(month_start.date(), month_end.date())))

    context = {
        'active_tab': active_tab,
        'day_val': target_day.strftime('%Y-%m-%d'), 'week_val': week_str, 'month_val': month_str,
        'title_day': target_day.strftime('%B %d, %Y'), 'title_week': f"Week of {week_start_date.strftime('%B %d, %Y')}",
        'title_month': month_start_date.strftime('%B %Y'),
        'daily': calculate_performance(orders_day + pickups_day),
        'weekly': calculate_performance(orders_week + pickups_week),
        'monthly': calculate_performance(orders_month + pickups_month),
    }
    return render(request, 'core/stats.html', context)


# ==========================================
# --- 3. DISPATCH & ORDER ENTRY ---
# ==========================================

@login_required
def add_to_run_sheet(request):
    customers = CustomerList.objects.all()
    return render(request, 'core/truck_dispatch/add_to_run_sheet.html', {'customers': customers})


@login_required
def entry_form(request, customer_id):
    """Main manual entry form for truck orders"""
    customer = get_object_or_404(CustomerList, customer_id=str(customer_id))
    all_employees = ["Ahmad", "Mikey", "Michael", "Danilio", "Kasim", "Carl", "David", "Douglas", "Jean Duval",
                     "Jean Thomas", "Cooper", "Aperam", "Elvis", "Leon", "Ismail"]
    regions = ['Beauce', 'Drummond', 'Montreal', 'North Shore', 'Ontario', 'Quebec', 'Sherbrooke', 'South Shore']

    def sort_for_station(priority_names):
        top = [name for name in priority_names if name in all_employees]
        rest = sorted([name for name in all_employees if name not in priority_names])
        return top + rest

    bar_employees = sort_for_station(["Ismail", "Jean Duval", "Jean Thomas", "Aperam", "Cooper", "Elvis"])
    sheet_employees = sort_for_station(["Mikey", "Danilio", "Kasim", "Carl"])
    cov_employees = sort_for_station(["David", "Douglas", "Ahmad"])

    if request.method == "POST":
        order_num = request.POST.get('order_number', '').strip()
        if order_num and not order_num.upper().startswith('W'):
            order_num = f"W{order_num}"

        closing_time = request.POST.get('closing_time', '').strip()
        custom_address = request.POST.get('address', '').strip() or customer.address
        custom_city = request.POST.get('city', '').strip() or customer.city
        custom_postal = request.POST.get('postal_code', '').strip() or getattr(customer, 'postal_code', '')
        custom_region = request.POST.get('region', '').strip() or customer.region

        w = int(request.POST.get('weight') or 0)
        s = int(request.POST.get('skids') or 0)
        b = int(request.POST.get('bundles') or 0)
        c = int(request.POST.get('coils') or 0)

        bar_lines = int(request.POST.get('bar_lines') or 0)
        bar_pickers = request.POST.getlist('bar_prep')
        bar_prep_str = ", ".join(sorted(bar_pickers)) if bar_pickers else ""

        sheet_lines = int(request.POST.get('sheet_lines') or 0)
        sheet_pickers = request.POST.getlist('sheet_prep')
        sheet_prep_str = ", ".join(sorted(sheet_pickers)) if sheet_pickers else ""

        cov_lines = int(request.POST.get('covering_lines') or 0)
        cov_pickers = request.POST.getlist('covering_prep')
        cov_prep_str = ", ".join(sorted(cov_pickers)) if cov_pickers else ""

        total_lines = bar_lines + sheet_lines + cov_lines
        total_prep_list = list(set(bar_pickers + sheet_pickers + cov_pickers))
        total_prep_str = ", ".join(sorted(total_prep_list)) if total_prep_list else "Unknown"

        max_index = RunSheet.objects.filter(region=custom_region).aggregate(Max('load_index'))['load_index__max'] or 0

        OrderArchive.objects.create(
            order_number=order_num, customer_id=customer.customer_id, customer_name=customer.customer_name,
            bar_prep=bar_prep_str, bar_lines=bar_lines, sheet_prep=sheet_prep_str, sheet_lines=sheet_lines,
            covering_prep=cov_prep_str, covering_lines=cov_lines, prepared_by=total_prep_str,
            line_items=total_lines, skids=s, bundles=b, coils=c, weight=w
        )

        RunSheet.objects.create(
            customer_id=customer.customer_id, customer_name=customer.customer_name,
            address=custom_address, city=custom_city, postal_code=custom_postal, region=custom_region,
            order_number=order_num, prepared_by=total_prep_str, closing_time=closing_time,
            weight=w, skids=s, bundles=b, coils=c, load_index=max_index + 1
        )
        return redirect('run_sheet')

    return render(request, 'core/truck_dispatch/entry_form.html', {
        'customer': customer, 'bar_employees': bar_employees, 'sheet_employees': sheet_employees,
        'cov_employees': cov_employees, 'regions': regions
    })


@login_required
def edit_order(request, pk):
    """Edits granular line item details for an existing Run Sheet order"""
    run_item = get_object_or_404(RunSheet, pk=pk)
    customer = get_object_or_404(CustomerList, customer_id=run_item.customer_id)
    archive = OrderArchive.objects.filter(order_number=run_item.order_number).order_by('-id').first()

    saved_bar = archive.bar_prep.split(', ') if archive and archive.bar_prep else []
    saved_sheet = archive.sheet_prep.split(', ') if archive and archive.sheet_prep else []
    saved_cov = archive.covering_prep.split(', ') if archive and archive.covering_prep else []

    all_employees = ["Ahmad", "Mikey", "Michael", "Danilio", "Kasim", "Carl", "David", "Douglas", "Jean Duval",
                     "Jean Thomas", "Cooper", "Aperam", "Elvis", "Leon", "Ismail"]

    def sort_for_station(priority_names):
        top = [name for name in priority_names if name in all_employees]
        rest = sorted([name for name in all_employees if name not in priority_names])
        return top + rest

    bar_employees = sort_for_station(["Ismail", "Jean Duval", "Jean Thomas", "Aperam", "Cooper", "Elvis"])
    sheet_employees = sort_for_station(["Mikey", "Danilio", "Kasim", "Carl"])
    cov_employees = sort_for_station(["David", "Douglas", "Ahmad"])

    if request.method == "POST":
        order_num = request.POST.get('order_number', '').strip()
        if order_num and not order_num.upper().startswith('W'): order_num = f"W{order_num}"

        closing_time = request.POST.get('closing_time', '').strip()
        w = int(request.POST.get('weight') or 0)
        s = int(request.POST.get('skids') or 0)
        b = int(request.POST.get('bundles') or 0)
        c = int(request.POST.get('coils') or 0)

        bar_lines = int(request.POST.get('bar_lines') or 0)
        bar_pickers = request.POST.getlist('bar_prep')
        bar_prep_str = ", ".join(sorted(bar_pickers)) if bar_pickers else ""

        sheet_lines = int(request.POST.get('sheet_lines') or 0)
        sheet_pickers = request.POST.getlist('sheet_prep')
        sheet_prep_str = ", ".join(sorted(sheet_pickers)) if sheet_pickers else ""

        cov_lines = int(request.POST.get('covering_lines') or 0)
        cov_pickers = request.POST.getlist('covering_prep')
        cov_prep_str = ", ".join(sorted(cov_pickers)) if cov_pickers else ""

        total_lines = bar_lines + sheet_lines + cov_lines
        total_prep_list = list(set(bar_pickers + sheet_pickers + cov_pickers))
        total_prep_str = ", ".join(sorted(total_prep_list)) if total_prep_list else "Unknown"

        run_item.order_number = order_num
        run_item.address = request.POST.get('address', run_item.address)
        run_item.city = request.POST.get('city', run_item.city)
        run_item.closing_time = closing_time
        run_item.weight, run_item.skids, run_item.bundles, run_item.coils = w, s, b, c
        run_item.prepared_by = total_prep_str
        run_item.save()

        if archive:
            archive.order_number = order_num
            archive.bar_prep, archive.bar_lines = bar_prep_str, bar_lines
            archive.sheet_prep, archive.sheet_lines = sheet_prep_str, sheet_lines
            archive.covering_prep, archive.covering_lines = cov_prep_str, cov_lines
            archive.prepared_by, archive.line_items = total_prep_str, total_lines
            archive.skids, archive.bundles, archive.coils, archive.weight = s, b, c, w
            archive.save()

        return redirect('run_sheet')

    return render(request, 'core/truck_dispatch/entry_form.html', {
        'customer': customer, 'order': run_item, 'archive': archive,
        'saved_bar': saved_bar, 'saved_sheet': saved_sheet, 'saved_cov': saved_cov,
        'bar_employees': bar_employees, 'sheet_employees': sheet_employees, 'cov_employees': cov_employees
    })


@login_required
def edit_specific_order(request, pk):
    """Simple edit form for basic order details"""
    order = get_object_or_404(RunSheet, pk=pk)
    employees = ["Ahmad", "Mikey", "Michael", "..."]
    if request.method == "POST":
        order.order_number = request.POST.get('order_number')
        order.weight = int(request.POST.get('weight') or 0)
        order.skids = int(request.POST.get('skids') or 0)
        order.save()
        return redirect('run_sheet')
    return render(request, 'core/edit_order_form.html', {'order': order, 'employees': sorted(employees)})


@login_required
def delete_stop(request, pk):
    stop = get_object_or_404(RunSheet, pk=pk)
    order_num = stop.order_number
    customer_name = stop.customer_name
    stop.delete()
    OrderArchive.objects.filter(order_number=order_num).delete()
    messages.warning(request, f"Order {order_num} for {customer_name} removed from truck and stats.")
    return redirect('run_sheet')


@login_required
def clear_run_sheet(request):
    if request.method == "POST":
        RunSheet.objects.all().delete()
        OrderArchive.objects.all().delete()
        messages.warning(request, "All run sheet data and order details have been cleared.")
    return redirect('run_sheet')


@login_required
def commit_and_clear_day(request):
    if request.method == "POST":
        current_stops = RunSheet.objects.all()
        if not current_stops.exists():
            messages.info(request, "Run sheet is already empty. Nothing to commit.")
            return redirect('run_sheet')
        for stop in current_stops:
            FinalizedRunSheet.objects.create(
                customer_name=stop.customer_name, region=stop.region,
                order_numbers=stop.order_number, weight=stop.weight,
                skids=stop.skids, bundles=stop.bundles, coils=stop.coils
            )
        RunSheet.objects.all().delete()
        messages.success(request, "Day officially committed to memory! Live list cleared.")
    return redirect('run_sheet')


@login_required
def finalize_run_sheet(request):
    shipping_date = get_next_business_day()
    now = timezone.now()
    today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    regions = ["North Shore", "Drummond", "Quebec", "Beauce", "Montreal", "South Shore", "Ontario", "Sherbrooke"]
    grouped_orders = {}

    for region in regions:
        orders = RunSheet.objects.filter(region=region).order_by('load_index')
        if orders.exists():
            cust_ids = orders.values_list('customer_id', flat=True)
            detailed_archive = OrderArchive.objects.filter(customer_id__in=cust_ids,
                                                           created_at__gte=today_start).order_by('created_at')
            grouped_orders[region] = {
                'orders': orders, 'detailed_archive': detailed_archive,
                'totals': {
                    'weight': sum(o.weight or 0 for o in orders), 'skids': sum(o.skids or 0 for o in orders),
                    'bundles': sum(o.bundles or 0 for o in orders), 'coils': sum(o.coils or 0 for o in orders),
                }
            }
    return render(request, 'core/finalize_view.html',
                  {'grouped_orders': grouped_orders, 'shipping_date': shipping_date})


@login_required
def run_sheet_history(request):
    history = FinalizedRunSheet.objects.all().order_by('-commit_date')
    return render(request, 'core/history.html', {'history': history})


# ==========================================
# --- 4. EXCEL IMPORT / EXPORT ---
# ==========================================

@login_required
def upload_run_sheet(request):
    if request.method == "POST" and request.FILES.get('excel_file'):
        file = request.FILES['excel_file']
        try:
            df = pd.read_csv(file, header=None).fillna('') if file.name.endswith('.csv') else pd.read_excel(file,
                                                                                                            header=None).fillna(
                '')
            RunSheet.objects.all().delete()
            import_count = 0
            for r in range(len(df)):
                for c in range(len(df.columns)):
                    cell_val = str(df.iloc[r, c]).strip()
                    is_old = "region" in cell_val.lower()
                    is_new = cell_val.startswith("---") and "---" in cell_val[3:]
                    if is_old or is_new:
                        current_region = cell_val.replace("region", "").replace("---", "").replace("TRUCK",
                                                                                                   "").strip().title()
                        driver_cell = str(df.iloc[r, c + 1])
                        d_name = driver_cell.lower().split("driver:")[1].split("-")[
                            0].strip().title() if "driver:" in driver_cell.lower() else ""
                        anchor_col = c if is_new else c - 1

                        for idx, i in enumerate(range(r + 1, len(df))):
                            row_id, row_name = str(df.iloc[i, anchor_col]).strip(), str(
                                df.iloc[i, anchor_col + 1]).strip()
                            if any(x in row_id.lower() for x in ["customer", "code", ""]): continue
                            if any(x in row_id.upper() for x in ["TOTAL", "REGION", "---"]) or row_name == "": break

                            def get_num(val):
                                clean = "".join(filter(str.isdigit, str(val).split('.')[0]))
                                return int(clean) if clean else 0

                            RunSheet.objects.create(
                                customer_id=row_id, customer_name=row_name,
                                city=str(df.iloc[i, anchor_col + 2]).strip(),
                                weight=get_num(df.iloc[i, anchor_col + 3]), skids=get_num(df.iloc[i, anchor_col + 4]),
                                bundles=get_num(df.iloc[i, anchor_col + 5]), coils=get_num(df.iloc[i, anchor_col + 6]),
                                order_number="W", region=current_region, driver_name=d_name,
                                prepared_by="Trucking Import", load_index=idx
                            )
                            import_count += 1
            if import_count > 0:
                messages.success(request, f"Successfully imported {import_count} orders!")
            else:
                messages.warning(request, "Found headers, but no data in expected columns.")
        except Exception as e:
            messages.error(request, f"Upload error: {e}")
    return redirect('run_sheet')


@login_required
def export_run_sheet_excel(request):
    regions = ["North Shore", "Drummond", "Quebec", "Beauce", "Montreal", "South Shore", "Ontario", "Sherbrooke"]
    final_data = []
    columns = ['Customer', 'City', 'Order #', 'Weight', 'Skids', 'Bundles', 'Coils', 'Lines', 'Prepared By']

    for region in regions:
        orders = RunSheet.objects.filter(region=region)
        if orders.exists():
            final_data.append(
                {'Customer': f'--- {region.upper()} TRUCK ---', 'City': '', 'Order #': '', 'Weight': '', 'Skids': '',
                 'Bundles': '', 'Coils': '', 'Lines': '', 'Prepared By': ''})
            t_w = t_s = t_b = t_c = 0
            for o in orders:
                final_data.append(
                    {'Customer': o.customer_name, 'City': o.city, 'Order #': o.order_number, 'Weight': o.weight or 0,
                     'Skids': o.skids or 0, 'Bundles': o.bundles or 0, 'Coils': o.coils or 0,
                     'Lines': o.line_items or 0, 'Prepared By': o.prepared_by})
                t_w += (o.weight or 0);
                t_s += (o.skids or 0);
                t_b += (o.bundles or 0);
                t_c += (o.coils or 0)
            final_data.append(
                {'Customer': f'TOTAL {region.upper()}', 'City': '', 'Order #': '', 'Weight': t_w, 'Skids': t_s,
                 'Bundles': t_b, 'Coils': t_c, 'Lines': '', 'Prepared By': ''})
            final_data.append({k: '' for k in columns})

    df = pd.DataFrame(final_data, columns=columns)
    filename = f"{get_next_business_day().strftime('%A %B %d')} Run Sheet.xlsx"
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'

    with pd.ExcelWriter(response, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Run Sheet')
        ws = writer.sheets['Run Sheet']
        header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        for row in ws.iter_rows(min_row=2):
            val = str(row[0].value)
            if "---" in val or "TOTAL" in val:
                for cell in row:
                    cell.font = Font(bold=True)
                    if "---" in val: cell.fill = header_fill
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = max(len(str(cell.value or "")) for cell in col) + 2
    return response


# ==========================================
# --- 5. CUSTOMER & VENDOR DATABASES ---
# ==========================================

@login_required
def customer_list(request):
    """Alternate view for customer list"""
    customers = CustomerList.objects.all()
    return render(request, 'core/database/customers.html', {'customers': customers})


@login_required
def manage_customers(request):
    customers = CustomerList.objects.all().order_by('customer_name')
    return render(request, 'core/database/manage_customers.html', {'customers': customers})


@login_required
def edit_customer(request, pk=None):
    regions = ["North Shore", "Drummond", "Quebec", "Beauce", "Montreal", "South Shore", "Ontario", "Sherbrooke"]
    customer = get_object_or_404(CustomerList, pk=pk) if pk else None
    title = f"Edit {customer.customer_name}" if pk else "Add New Customer"

    if request.method == "POST":
        cid, name = request.POST.get('customer_id'), request.POST.get('customer_name')
        addr, city = request.POST.get('address'), request.POST.get('city')
        reg, post = request.POST.get('region'), request.POST.get('postal_code')

        if pk:
            customer.customer_id, customer.customer_name = cid, name
            customer.address, customer.city = addr, city
            customer.region, customer.postal_code = reg, post
            customer.save()
        else:
            CustomerList.objects.create(customer_id=cid, customer_name=name, address=addr, city=city, region=reg,
                                        postal_code=post)
        return redirect('manage_customers')
    return render(request, 'core/database/edit_customer.html',
                  {'customer': customer, 'title': title, 'regions': sorted(regions)})


@login_required
def manage_vendors(request):
    vendors = Vendor.objects.all().order_by('name')
    return render(request, 'core/database/manage_vendors.html', {'vendors': vendors})


@login_required
def edit_vendor(request, pk=None):
    regions = ["North Shore", "Drummond", "Quebec", "Beauce", "Montreal", "South Shore", "Ontario", "Sherbrooke"]
    vendor = get_object_or_404(Vendor, pk=pk) if pk else None
    title = f"Edit {vendor.name}" if pk else "Add New Vendor"

    if request.method == "POST":
        name, addr = request.POST.get('name'), request.POST.get('address')
        city, reg, post = request.POST.get('city'), request.POST.get('region'), request.POST.get('postal_code')

        if pk:
            vendor.name, vendor.address, vendor.city = name, addr, city
            vendor.region, vendor.postal_code = reg, post
            vendor.save()
            messages.success(request, f"Vendor {name} updated.")
        else:
            Vendor.objects.create(name=name, address=addr, city=city, region=reg, postal_code=post)
            messages.success(request, f"New vendor {name} added.")
        return redirect('manage_vendors')
    return render(request, 'core/database/edit_vendor.html', {'vendor': vendor, 'title': title, 'regions': sorted(regions)})


# ==========================================
# --- 6. TRUCK LOGIC: PICKUPS & RETURNS ---
# ==========================================

@login_required
def select_vendor_pickup(request):
    vendors = Vendor.objects.all().order_by('name')
    return render(request, 'core/truck_dispatch/select_vendor_pickup.html', {'vendors': vendors})


@login_required
def add_pickup(request):
    vendors = Vendor.objects.all().order_by('name')
    regions = ["North Shore", "Drummond", "Quebec", "Beauce", "Montreal", "South Shore", "Ontario", "Sherbrooke"]
    if request.method == "POST":
        v_name, pos = request.POST.get('vendor_name').strip(), request.POST.get('po_numbers').strip().upper()
        addr, city = request.POST.get('address').strip(), request.POST.get('city').strip()
        region = request.POST.get('region')

        if request.POST.get('save_vendor'):
            Vendor.objects.get_or_create(name=v_name, address=addr, city=city, region=region)

        max_idx = RunSheet.objects.filter(region=region).aggregate(Max('load_index'))['load_index__max'] or 0
        RunSheet.objects.create(
            customer_name=f"PICKUP: {v_name}", address=addr, city=city,
            region=region, order_number=pos, is_pickup=True, load_index=max_idx + 1
        )
        return redirect('run_sheet')
    return render(request, 'core/truck_dispatch/add_pickup.html', {'vendors': vendors, 'regions': regions})


@login_required
def add_pickup_form(request, vendor_id):
    vendor = get_object_or_404(Vendor, pk=vendor_id)
    regions = ["North Shore", "Drummond", "Quebec", "Beauce", "Montreal", "South Shore", "Ontario", "Sherbrooke"]

    if request.method == "POST":
        pos = request.POST.get('po_numbers', '').strip().upper()
        w = max(0, int(request.POST.get('weight') or 0))
        s = max(0, int(request.POST.get('skids') or 0))
        b = max(0, int(request.POST.get('bundles') or 0))
        c = max(0, int(request.POST.get('coils') or 0))

        p_city = request.POST.get('city', '').strip() or vendor.city
        p_addr = request.POST.get('address', '').strip() or vendor.address
        selected_truck = request.POST.get('region')

        is_redirect = request.POST.get('is_redirect') == 'on'
        d_name, d_city = request.POST.get('dest_name', '').strip(), request.POST.get('dest_city', '').strip()
        d_postal = request.POST.get('dest_postal', '').strip().upper()

        if is_redirect and d_name:
            display_name = f"PICKUP: {vendor.name} ({p_city}) ➔ DELIVER: {d_name} ({d_city} {d_postal})"
        else:
            display_name = f"PICKUP: {vendor.name}"

        max_idx = RunSheet.objects.filter(region=selected_truck).aggregate(Max('load_index'))['load_index__max'] or 0
        RunSheet.objects.create(
            customer_name=display_name, address=p_addr, city=p_city, order_number=pos,
            weight=w, skids=s, bundles=b, coils=c, is_pickup=True, region=selected_truck,
            load_index=max_idx + 1, prepared_by="Office"
        )
        messages.success(request, f"Pickup for {vendor.name} successfully added.")
        return redirect('run_sheet')
    return render(request, 'core/truck_dispatch/add_pickup_form.html', {'vendor': vendor, 'regions': regions})


@login_required
def select_customer_return(request):
    customers = CustomerList.objects.all().order_by('customer_name')
    return render(request, 'core/truck_dispatch/select_customer_return.html', {'customers': customers})


@login_required
def add_return(request):
    customers = CustomerList.objects.all().order_by('customer_name')
    if request.method == "POST":
        customer = get_object_or_404(CustomerList, customer_id=request.POST.get('customer_id'))
        ret_num = request.POST.get('return_number', '').strip().upper()
        if ret_num and not ret_num.startswith('K'): ret_num = f"K{ret_num}"
        max_idx = RunSheet.objects.filter(region=customer.region).aggregate(Max('load_index'))['load_index__max'] or 0

        RunSheet.objects.create(
            customer_id=customer.customer_id, customer_name=f"RETURN: {customer.customer_name}",
            address=customer.address, city=customer.city, region=customer.region, order_number=ret_num,
            is_return=True, load_index=max_idx + 1, prepared_by="Office"
        )
        messages.success(request, f"Return {ret_num} added to board.")
        return redirect('run_sheet')
    return render(request, 'core/truck_dispatch/add_return.html', {'customers': customers})


@login_required
def add_return_form(request, customer_id):
    customer = get_object_or_404(CustomerList, customer_id=str(customer_id))
    regions = ['Beauce', 'Drummond', 'Montreal', 'North Shore', 'Ontario', 'Quebec', 'Sherbrooke', 'South Shore']

    if request.method == "POST":
        ret_num = request.POST.get('return_number', '').strip().upper()
        if ret_num and not ret_num.startswith('K'): ret_num = f"K{ret_num}"

        w = int(request.POST.get('weight') or 0)
        s = int(request.POST.get('skids') or 0)
        b = int(request.POST.get('bundles') or 0)
        c = int(request.POST.get('coils') or 0)

        custom_address = request.POST.get('address', '').strip() or customer.address
        custom_city = request.POST.get('city', '').strip() or customer.city
        custom_postal = request.POST.get('postal_code', '').strip() or getattr(customer, 'postal_code', '')
        custom_region = request.POST.get('region', '').strip() or customer.region

        max_idx = RunSheet.objects.filter(region=custom_region).aggregate(Max('load_index'))['load_index__max'] or 0

        RunSheet.objects.create(
            customer_id=customer.customer_id, customer_name=f"RETURN: {customer.customer_name}",
            address=custom_address, city=custom_city, postal_code=custom_postal, region=custom_region,
            order_number=ret_num, weight=w, skids=s, bundles=b, coils=c, is_return=True, load_index=max_idx + 1
        )
        return redirect('run_sheet')
    return render(request, 'core/truck_dispatch/add_return_form.html', {'customer': customer, 'regions': regions})


# ==========================================
# --- 7. COUNTER PICKUPS (PERMANENT LOG) ---
# ==========================================

@login_required
def select_customer_pickup(request):
    customers = CustomerList.objects.all().order_by('customer_name')
    return render(request, 'core/customer_pickup/select_customer_pickup.html', {'customers': customers})


@login_required
def pickup_log_list(request):
    dates = PickupLog.objects.dates('date_completed', 'day', order='DESC')
    return render(request, 'core/customer_pickup/cust_pickup.html', {'dates': dates})


@login_required
def daily_pickup_detail(request, date):
    orders = PickupLog.objects.filter(date_completed=date).order_by('customer_name', 'order_number')
    return render(request, 'core/customer_pickup/daily_pickup_detail.html', {'orders': orders, 'date': date})


@login_required
def add_pickup_order(request, customer_id):
    customer = get_object_or_404(CustomerList, customer_id=str(customer_id))
    all_employees = ["Ahmad", "Mikey", "Michael", "Danilio", "Kasim", "Carl", "David", "Douglas", "Jean Duval",
                     "Jean Thomas", "Cooper", "Aperam", "Elvis", "Leon", "Ismail"]

    def sort_for_station(priority_names):
        top = [name for name in priority_names if name in all_employees]
        rest = sorted([name for name in all_employees if name not in priority_names])
        return top + rest

    bar_employees = sort_for_station(["Ismail", "Jean Duval", "Jean Thomas", "Aperam", "Cooper", "Elvis"])
    sheet_employees = sort_for_station(["Mikey", "Danilio", "Kasim", "Carl"])
    cov_employees = sort_for_station(["David", "Douglas", "Ahmad"])

    if request.method == "POST":
        order_num = request.POST.get('order_number', '').strip()
        if order_num and not order_num.upper().startswith('W'): order_num = f"W{order_num}"

        w = max(0, int(request.POST.get('weight') or 0))
        s = max(0, int(request.POST.get('skids') or 0))
        b = max(0, int(request.POST.get('bundles') or 0))
        c = max(0, int(request.POST.get('coils') or 0))

        bar_lines, sheet_lines, cov_lines = max(0, int(request.POST.get('bar_lines') or 0)), max(0,
                                                                                                 int(request.POST.get(
                                                                                                     'sheet_lines') or 0)), max(
            0, int(request.POST.get('covering_lines') or 0))

        bar_pickers = request.POST.getlist('bar_prep')
        sheet_pickers = request.POST.getlist('sheet_prep')
        cov_pickers = request.POST.getlist('covering_prep')

        PickupLog.objects.create(
            customer_name=customer.customer_name, order_number=order_num,
            weight=w, skids=s, bundles=b, coils=c,
            bar_lines=bar_lines, sheet_lines=sheet_lines, covering_lines=cov_lines,
            bar_prep=", ".join(sorted(bar_pickers)) if bar_pickers else "",
            sheet_prep=", ".join(sorted(sheet_pickers)) if sheet_pickers else "",
            covering_prep=", ".join(sorted(cov_pickers)) if cov_pickers else ""
        )
        messages.success(request, f"Counter Pickup {order_num} logged.")
        return redirect('pickup_log_list')

    return render(request, 'core/truck_dispatch/entry_form.html', {
        'customer': customer, 'bar_employees': bar_employees, 'sheet_employees': sheet_employees,
        'cov_employees': cov_employees, 'is_counter_pickup': True
    })


@login_required
def edit_pickup_order(request, pk):
    order = get_object_or_404(PickupLog, pk=pk)
    all_employees = ["Ahmad", "Mikey", "Michael", "Danilio", "Kasim", "Carl", "David", "Douglas", "Jean Duval",
                     "Jean Thomas", "Cooper", "Aperam", "Elvis", "Leon", "Ismail"]
    saved_bar = [n.strip() for n in order.bar_prep.split(',')]
    saved_sheet = [n.strip() for n in order.sheet_prep.split(',')]
    saved_cov = [n.strip() for n in order.covering_prep.split(',')]

    if request.method == "POST":
        order.order_number = request.POST.get('order_number')
        order.weight = max(0, int(request.POST.get('weight') or 0))
        order.skids = max(0, int(request.POST.get('skids') or 0))
        order.bundles = max(0, int(request.POST.get('bundles') or 0))
        order.coils = max(0, int(request.POST.get('coils') or 0))
        order.bar_lines = max(0, int(request.POST.get('bar_lines') or 0))
        order.sheet_lines = max(0, int(request.POST.get('sheet_lines') or 0))
        order.covering_lines = max(0, int(request.POST.get('covering_lines') or 0))
        order.bar_prep = ", ".join(request.POST.getlist('bar_prep'))
        order.sheet_prep = ", ".join(request.POST.getlist('sheet_prep'))
        order.covering_prep = ", ".join(request.POST.getlist('covering_prep'))
        order.save()
        messages.warning(request, f"Pickup {order.order_number} updated.")
        return redirect('daily_pickup_detail', date=order.date_completed.strftime('%Y-%m-%d'))

    return render(request, 'core/truck_dispatch/entry_form.html', {
        'order': order, 'is_counter_pickup': True, 'bar_employees': all_employees,
        'sheet_employees': all_employees, 'cov_employees': all_employees,
        'saved_bar': saved_bar, 'saved_sheet': saved_sheet, 'saved_cov': saved_cov
    })


@login_required
def delete_pickup_order(request, pk):
    order = get_object_or_404(PickupLog, pk=pk)
    target_date = order.date_completed.strftime('%Y-%m-%d')
    order.delete()
    messages.error(request, "Pickup record deleted.")
    return redirect('daily_pickup_detail', date=target_date)


# ==========================================
# --- 8. PHOTO LOGS: CONTAINERS ---
# ==========================================

@login_required
def container_list(request):
    containers = Container.objects.all().order_by('-date_received')
    return render(request, 'core/containers/list.html', {'containers': containers})


@login_required
def container_detail(request, pk):
    container = get_object_or_404(Container, pk=pk)
    return render(request, 'core/containers/detail.html', {'container': container})


@login_required
def add_container(request):
    prefill_num = request.GET.get('container_num', '')
    if request.method == 'POST':
        container_num = request.POST.get('container_number', '').strip().upper()
        images = request.FILES.getlist('photos')
        if container_num and images:
            container, created = Container.objects.get_or_create(container_number=container_num)
            for image in images:
                ContainerPhoto.objects.create(container=container, image=image)
            messages.success(request, f"Successfully uploaded {len(images)} photos.")
            return redirect('container_detail', pk=container.pk)
        messages.error(request, "Provide a container number and photo.")
    return render(request, 'core/containers/add.html', {'prefill_num': prefill_num})


@login_required
def upload_more_container_photos(request, pk):
    container = get_object_or_404(Container, pk=pk)
    if request.method == 'POST':
        images = request.FILES.getlist('photos')
        if images:
            for img in images:
                ContainerPhoto.objects.create(container=container, image=img)
            messages.success(request, f"Added {len(images)} photos.")
    return redirect('container_detail', pk=container.pk)


@login_required
def delete_container(request, pk):
    container = get_object_or_404(Container, pk=pk)
    container.delete()
    messages.warning(request, "Container and photos deleted.")
    return redirect('container_list')


@login_required
def delete_container_photo(request, photo_id):
    photo = get_object_or_404(ContainerPhoto, pk=photo_id)
    container_id = photo.container.pk
    photo.delete()
    return redirect('container_detail', pk=container_id)


# ==========================================
# --- 9. PHOTO LOGS: OUTBOUND TRUCKS ---
# ==========================================

@login_required
def outbound_list(request):
    loads = OutboundLoad.objects.all().order_by('-date_loaded', 'truck_name')
    return render(request, 'core/outbound/list.html', {'loads': loads})


@login_required
def outbound_detail(request, pk):
    load = get_object_or_404(OutboundLoad, pk=pk)
    return render(request, 'core/outbound/detail.html', {'load': load})


@login_required
def add_outbound_photos(request):
    if request.method == 'POST':
        truck_name = request.POST.get('truck_name', '').strip().upper()
        images = request.FILES.getlist('photos')
        if truck_name and images:
            load, created = OutboundLoad.objects.get_or_create(truck_name=truck_name, date_loaded=timezone.now().date())
            for img in images:
                OutboundPhoto.objects.create(load=load, image=img)
            messages.success(request, f"Photos uploaded for {truck_name}")
            return redirect('outbound_detail', pk=load.pk)
        messages.error(request, "Provide a truck name and photo.")
    return render(request, 'core/outbound/add.html')


@login_required
def upload_more_outbound_photos(request, pk):
    load = get_object_or_404(OutboundLoad, pk=pk)
    if request.method == 'POST':
        images = request.FILES.getlist('photos')
        if images:
            for img in images:
                OutboundPhoto.objects.create(load=load, image=img)
            messages.success(request, f"Added {len(images)} photos.")
    return redirect('outbound_detail', pk=load.pk)


@login_required
def delete_outbound_load(request, pk):
    load = get_object_or_404(OutboundLoad, pk=pk)
    load.delete()
    messages.warning(request, "Load and photos deleted.")
    return redirect('outbound_list')


@login_required
def delete_outbound_photo(request, photo_id):
    photo = get_object_or_404(OutboundPhoto, pk=photo_id)
    load_id = photo.load.pk
    photo.delete()
    return redirect('outbound_detail', pk=load_id)


# ==========================================
# --- 10. PHOTO LOGS: CUSTOMER PICKUPS ---
# ==========================================

@login_required
def pickup_photo_list(request):
    logs = PickupPhotoLog.objects.all().order_by('-date_picked_up', 'customer_name')
    return render(request, 'core/pickups/list.html', {'logs': logs})


@login_required
def pickup_photo_detail(request, pk):
    log = get_object_or_404(PickupPhotoLog, pk=pk)
    return render(request, 'core/pickups/detail.html', {'log': log})


@login_required
def add_pickup_photos(request):
    if request.method == 'POST':
        cust_name = request.POST.get('customer_name', '').strip().upper()
        order_num = request.POST.get('order_number', '').strip().upper()
        images = request.FILES.getlist('photos')
        if cust_name and images:
            log, created = PickupPhotoLog.objects.get_or_create(customer_name=cust_name, order_number=order_num,
                                                                date_picked_up=timezone.now().date())
            for img in images:
                PickupPhoto.objects.create(log=log, image=img)
            messages.success(request, f"Photos saved for {cust_name}")
            return redirect('pickup_photo_detail', pk=log.pk)
    return render(request, 'core/pickups/add.html')


@login_required
def upload_more_pickup_photos(request, pk):
    log = get_object_or_404(PickupPhotoLog, pk=pk)
    if request.method == 'POST':
        images = request.FILES.getlist('photos')
        if images:
            for img in images:
                PickupPhoto.objects.create(log=log, image=img)
            messages.success(request, f"Added {len(images)} photos.")
    return redirect('pickup_photo_detail', pk=log.pk)


@login_required
def delete_pickup_photo_log(request, pk):
    log = get_object_or_404(PickupPhotoLog, pk=pk)
    log.delete()
    messages.warning(request, "Pickup photo folder deleted.")
    return redirect('pickup_photo_list')


@login_required
def delete_pickup_individual_photo(request, photo_id):
    photo = get_object_or_404(PickupPhoto, pk=photo_id)
    log_id = photo.log.pk
    photo.delete()
    return redirect('pickup_photo_detail', pk=log_id)