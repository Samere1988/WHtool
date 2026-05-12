from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.shortcuts import get_object_or_404, redirect

from .models import Container, OutboundLoad, PickupPhotoLog


@login_required
def edit_outbound_load(request, pk):
    load = get_object_or_404(OutboundLoad, pk=pk)

    if request.method != "POST":
        return redirect("outbound_detail", pk=load.pk)

    truck_name = request.POST.get("truck_name", "").strip()
    if not truck_name:
        messages.error(request, "Truck name is required.")
        return redirect("outbound_detail", pk=load.pk)

    load.truck_name = truck_name
    load.save(update_fields=["truck_name"])
    messages.success(request, "Outbound load info updated.")
    return redirect("outbound_detail", pk=load.pk)


@login_required
def edit_container(request, pk):
    container = get_object_or_404(Container, pk=pk)

    if request.method != "POST":
        return redirect("container_detail", pk=container.pk)

    container_number = request.POST.get("container_number", "").strip()
    unloaded_by = request.POST.get("unloaded_by", "").strip()

    if not container_number:
        messages.error(request, "Container number is required.")
        return redirect("container_detail", pk=container.pk)

    container.container_number = container_number
    container.unloaded_by = unloaded_by
    container.save(update_fields=["container_number", "unloaded_by"])
    messages.success(request, "Container info updated.")
    return redirect("container_detail", pk=container.pk)


@login_required
def edit_pickup_photo_log(request, pk):
    log = get_object_or_404(PickupPhotoLog, pk=pk)

    if request.method != "POST":
        return redirect("pickup_photo_detail", pk=log.pk)

    customer_name = request.POST.get("customer_name", "").strip()
    customer_id = request.POST.get("customer_id", "").strip()
    order_number = request.POST.get("order_number", "").strip()
    date_picked_up = request.POST.get("date_picked_up", "").strip()

    if not customer_name or not order_number:
        messages.error(request, "Customer name and order number are required.")
        return redirect("pickup_photo_detail", pk=log.pk)

    log.customer_name = customer_name
    log.customer_id = customer_id
    log.order_number = order_number

    if date_picked_up:
        log.date_picked_up = date_picked_up

    log.save(update_fields=["customer_name", "customer_id", "order_number", "date_picked_up"])
    messages.success(request, "Pickup photo folder info updated.")
    return redirect("pickup_photo_detail", pk=log.pk)
