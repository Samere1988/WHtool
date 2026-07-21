from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.db.models import Count
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.utils.dateparse import parse_date
from django.views.decorators.http import require_POST

from .models import (
    Container,
    ContainerPhoto,
    OutboundLoad,
    OutboundPhoto,
    PickupPhoto,
    PickupPhotoLog,
)
from .photo_utils import (
    PhotoUploadError,
    create_photo_records,
    validate_photo_uploads,
)


def save_more_photos(
    request,
    parent,
    photo_model,
    parent_field,
    detail_url_name,
):
    try:
        uploads = validate_photo_uploads(
            request.FILES.getlist("photos")
        )

        with transaction.atomic():
            photo_count = create_photo_records(
                photo_model,
                parent_field,
                parent,
                uploads,
            )

    except PhotoUploadError as exc:
        messages.error(request, str(exc))
    else:
        messages.success(
            request,
            f"Added {photo_count} photos.",
        )

    return redirect(
        detail_url_name,
        pk=parent.pk,
    )


@login_required
def container_list(request):
    containers = (
        Container.objects
        .annotate(photo_count=Count("photos"))
        .order_by("-date_received", "-id")
    )

    return render(
        request,
        "core/containers/list.html",
        {"containers": containers},
    )


@login_required
def container_detail(request, pk):
    container = get_object_or_404(
        Container,
        pk=pk,
    )

    return render(
        request,
        "core/containers/detail.html",
        {
            "container": container,
            "photos": container.photos.order_by(
                "uploaded_at",
                "id",
            ),
        },
    )


@login_required
def add_container(request):
    prefill_num = request.GET.get(
        "container_num",
        "",
    )

    if request.method == "POST":
        container_number = request.POST.get(
            "container_number",
            "",
        ).strip().upper()

        unloaded_by = request.POST.get(
            "unloaded_by",
            "",
        ).strip()

        unloaded_at = request.POST.get(
            "unloaded_at",
            "",
        ).strip()

        date_received = parse_date(
            request.POST.get(
                "date_received",
                "",
            ).strip()
        )

        if not container_number:
            messages.error(
                request,
                "Enter a container number.",
            )
        elif not unloaded_by:
            messages.error(
                request,
                "Enter who unloaded the container.",
            )
        elif unloaded_at not in {
            "20 Hymus",
            "26 Hymus",
        }:
            messages.error(
                request,
                "Choose where the container was unloaded.",
            )
        elif date_received is None:
            messages.error(
                request,
                "Choose a valid unloaded date.",
            )
        else:
            try:
                uploads = validate_photo_uploads(
                    request.FILES.getlist("photos")
                )

                with transaction.atomic():
                    container = (
                        Container.objects
                        .select_for_update()
                        .filter(
                            container_number=container_number,
                            date_received=date_received,
                        )
                        .order_by("id")
                        .first()
                    )

                    if container is None:
                        container = Container.objects.create(
                            container_number=container_number,
                            date_received=date_received,
                            unloaded_by=unloaded_by,
                            unloaded_at=unloaded_at,
                        )
                    else:
                        container.unloaded_by = unloaded_by
                        container.unloaded_at = unloaded_at
                        container.save(
                            update_fields=[
                                "unloaded_by",
                                "unloaded_at",
                            ]
                        )

                    photo_count = create_photo_records(
                        ContainerPhoto,
                        "container",
                        container,
                        uploads,
                    )

            except PhotoUploadError as exc:
                messages.error(request, str(exc))
            else:
                messages.success(
                    request,
                    f"Successfully uploaded {photo_count} photos.",
                )

                return redirect(
                    "container_detail",
                    pk=container.pk,
                )

    return render(
        request,
        "core/containers/add.html",
        {
            "prefill_num": prefill_num,
            "today": timezone.localdate(),
        },
    )


@login_required
@require_POST
def upload_more_container_photos(request, pk):
    container = get_object_or_404(
        Container,
        pk=pk,
    )

    return save_more_photos(
        request,
        container,
        ContainerPhoto,
        "container",
        "container_detail",
    )


@login_required
@require_POST
def edit_container(request, pk):
    container = get_object_or_404(
        Container,
        pk=pk,
    )

    container_number = request.POST.get(
        "container_number",
        "",
    ).strip().upper()

    unloaded_by = request.POST.get(
        "unloaded_by",
        "",
    ).strip()

    unloaded_at = request.POST.get(
        "unloaded_at",
        "",
    ).strip()

    date_received = parse_date(
        request.POST.get(
            "date_received",
            "",
        ).strip()
    )

    if not container_number:
        messages.error(
            request,
            "Container number is required.",
        )
    elif not unloaded_by:
        messages.error(
            request,
            "Unloaded by is required.",
        )
    elif unloaded_at not in {
        "20 Hymus",
        "26 Hymus",
    }:
        messages.error(
            request,
            "Choose where the container was unloaded.",
        )
    elif date_received is None:
        messages.error(
            request,
            "Choose a valid unloaded date.",
        )
    else:
        container.container_number = container_number
        container.unloaded_by = unloaded_by
        container.unloaded_at = unloaded_at
        container.date_received = date_received

        container.save(
            update_fields=[
                "container_number",
                "unloaded_by",
                "unloaded_at",
                "date_received",
            ]
        )

        messages.success(
            request,
            "Container info updated.",
        )

    return redirect(
        "container_detail",
        pk=container.pk,
    )


@login_required
@require_POST
def delete_container(request, pk):
    container = get_object_or_404(
        Container,
        pk=pk,
    )

    container.delete()

    messages.warning(
        request,
        "Container and photos deleted.",
    )

    return redirect("container_list")


@login_required
@require_POST
def delete_container_photo(request, photo_id):
    photo = get_object_or_404(
        ContainerPhoto,
        pk=photo_id,
    )

    container_id = photo.container_id
    photo.delete()

    return redirect(
        "container_detail",
        pk=container_id,
    )


@login_required
def outbound_list(request):
    loads = (
        OutboundLoad.objects
        .annotate(photo_count=Count("photos"))
        .order_by("-date_loaded", "truck_name", "-id")
    )

    return render(
        request,
        "core/outbound/list.html",
        {"loads": loads},
    )


@login_required
def outbound_detail(request, pk):
    load = get_object_or_404(
        OutboundLoad,
        pk=pk,
    )

    return render(
        request,
        "core/outbound/detail.html",
        {
            "load": load,
            "photos": load.photos.order_by(
                "uploaded_at",
                "id",
            ),
        },
    )


@login_required
def add_outbound_photos(request):
    if request.method == "POST":
        driver_name = request.POST.get(
            "truck_name",
            "",
        ).strip().upper()

        loaded_by = request.POST.get(
            "loaded_by",
            "",
        ).strip()

        date_loaded = parse_date(
            request.POST.get(
                "date_loaded",
                "",
            ).strip()
        )

        if not driver_name:
            messages.error(
                request,
                "Enter the driver name.",
            )
        elif not loaded_by:
            messages.error(
                request,
                "Enter who loaded the truck.",
            )
        elif date_loaded is None:
            messages.error(
                request,
                "Choose a valid loaded date.",
            )
        else:
            try:
                uploads = validate_photo_uploads(
                    request.FILES.getlist("photos")
                )

                with transaction.atomic():
                    load = (
                        OutboundLoad.objects
                        .select_for_update()
                        .filter(
                            truck_name=driver_name,
                            date_loaded=date_loaded,
                        )
                        .order_by("id")
                        .first()
                    )

                    if load is None:
                        load = OutboundLoad.objects.create(
                            truck_name=driver_name,
                            date_loaded=date_loaded,
                            loaded_by=loaded_by,
                        )
                    else:
                        load.loaded_by = loaded_by
                        load.save(
                            update_fields=["loaded_by"]
                        )

                    photo_count = create_photo_records(
                        OutboundPhoto,
                        "load",
                        load,
                        uploads,
                    )

            except PhotoUploadError as exc:
                messages.error(request, str(exc))
            else:
                messages.success(
                    request,
                    f"Uploaded {photo_count} photos for {driver_name}.",
                )

                return redirect(
                    "outbound_detail",
                    pk=load.pk,
                )

    return render(
        request,
        "core/outbound/add.html",
        {"today": timezone.localdate()},
    )


@login_required
@require_POST
def upload_more_outbound_photos(request, pk):
    load = get_object_or_404(
        OutboundLoad,
        pk=pk,
    )

    return save_more_photos(
        request,
        load,
        OutboundPhoto,
        "load",
        "outbound_detail",
    )


@login_required
@require_POST
def edit_outbound_load(request, pk):
    load = get_object_or_404(
        OutboundLoad,
        pk=pk,
    )

    date_loaded = parse_date(
        request.POST.get(
            "date_loaded",
            "",
        ).strip()
    )

    loaded_by = request.POST.get(
        "loaded_by",
        "",
    ).strip()

    driver_name = request.POST.get(
        "truck_name",
        "",
    ).strip().upper()

    if date_loaded is None:
        messages.error(
            request,
            "Choose a valid loaded date.",
        )
    elif not loaded_by:
        messages.error(
            request,
            "Loaded by is required.",
        )
    elif not driver_name:
        messages.error(
            request,
            "Driver is required.",
        )
    else:
        load.date_loaded = date_loaded
        load.loaded_by = loaded_by
        load.truck_name = driver_name

        load.save(
            update_fields=[
                "date_loaded",
                "loaded_by",
                "truck_name",
            ]
        )

        messages.success(
            request,
            "Outbound load info updated.",
        )

    return redirect(
        "outbound_detail",
        pk=load.pk,
    )


@login_required
@require_POST
def delete_outbound_load(request, pk):
    load = get_object_or_404(
        OutboundLoad,
        pk=pk,
    )

    load.delete()

    messages.warning(
        request,
        "Load and photos deleted.",
    )

    return redirect("outbound_list")


@login_required
@require_POST
def delete_outbound_photo(request, photo_id):
    photo = get_object_or_404(
        OutboundPhoto,
        pk=photo_id,
    )

    load_id = photo.load_id
    photo.delete()

    return redirect(
        "outbound_detail",
        pk=load_id,
    )


@login_required
def pickup_photo_list(request):
    logs = (
        PickupPhotoLog.objects
        .annotate(photo_count=Count("photos"))
        .order_by("-date_picked_up", "customer_name", "-id")
    )

    return render(
        request,
        "core/pickups/list.html",
        {"logs": logs},
    )


@login_required
def pickup_photo_detail(request, pk):
    log = get_object_or_404(
        PickupPhotoLog,
        pk=pk,
    )

    return render(
        request,
        "core/pickups/detail.html",
        {
            "log": log,
            "photos": log.photos.order_by(
                "uploaded_at",
                "id",
            ),
        },
    )


@login_required
def add_pickup_photos(request):
    if request.method == "POST":
        customer_name = request.POST.get(
            "customer_name",
            "",
        ).strip().upper()

        order_number = request.POST.get(
            "order_number",
            "",
        ).strip().upper()

        loaded_by = request.POST.get(
            "loaded_by",
            "",
        ).strip()

        date_picked_up = parse_date(
            request.POST.get(
                "date_picked_up",
                "",
            ).strip()
        )

        if not customer_name:
            messages.error(
                request,
                "Enter the customer name.",
            )
        elif not loaded_by:
            messages.error(
                request,
                "Enter who loaded the pickup.",
            )
        elif date_picked_up is None:
            messages.error(
                request,
                "Choose a valid pickup date.",
            )
        else:
            try:
                uploads = validate_photo_uploads(
                    request.FILES.getlist("photos")
                )

                with transaction.atomic():
                    log = (
                        PickupPhotoLog.objects
                        .select_for_update()
                        .filter(
                            customer_name=customer_name,
                            order_number=order_number,
                            date_picked_up=date_picked_up,
                        )
                        .order_by("id")
                        .first()
                    )

                    if log is None:
                        log = PickupPhotoLog.objects.create(
                            customer_name=customer_name,
                            order_number=order_number,
                            date_picked_up=date_picked_up,
                            loaded_by=loaded_by,
                        )
                    else:
                        log.loaded_by = loaded_by
                        log.save(
                            update_fields=["loaded_by"]
                        )

                    photo_count = create_photo_records(
                        PickupPhoto,
                        "log",
                        log,
                        uploads,
                    )

            except PhotoUploadError as exc:
                messages.error(request, str(exc))
            else:
                messages.success(
                    request,
                    f"Saved {photo_count} photos for {customer_name}.",
                )

                return redirect(
                    "pickup_photo_detail",
                    pk=log.pk,
                )

    return render(
        request,
        "core/pickups/add.html",
        {"today": timezone.localdate()},
    )


@login_required
@require_POST
def upload_more_pickup_photos(request, pk):
    log = get_object_or_404(
        PickupPhotoLog,
        pk=pk,
    )

    return save_more_photos(
        request,
        log,
        PickupPhoto,
        "log",
        "pickup_photo_detail",
    )


@login_required
@require_POST
def edit_pickup_photo_log(request, pk):
    log = get_object_or_404(
        PickupPhotoLog,
        pk=pk,
    )

    customer_name = request.POST.get(
        "customer_name",
        "",
    ).strip().upper()

    order_number = request.POST.get(
        "order_number",
        "",
    ).strip().upper()

    loaded_by = request.POST.get(
        "loaded_by",
        "",
    ).strip()

    date_picked_up = parse_date(
        request.POST.get(
            "date_picked_up",
            "",
        ).strip()
    )

    if not customer_name:
        messages.error(
            request,
            "Customer name is required.",
        )
    elif not loaded_by:
        messages.error(
            request,
            "Loaded by is required.",
        )
    elif date_picked_up is None:
        messages.error(
            request,
            "Choose a valid pickup date.",
        )
    else:
        log.customer_name = customer_name
        log.order_number = order_number
        log.loaded_by = loaded_by
        log.date_picked_up = date_picked_up

        log.save(
            update_fields=[
                "customer_name",
                "order_number",
                "loaded_by",
                "date_picked_up",
            ]
        )

        messages.success(
            request,
            "Pickup photo folder info updated.",
        )

    return redirect(
        "pickup_photo_detail",
        pk=log.pk,
    )


@login_required
@require_POST
def delete_pickup_photo_log(request, pk):
    log = get_object_or_404(
        PickupPhotoLog,
        pk=pk,
    )

    log.delete()

    messages.warning(
        request,
        "Pickup photo folder deleted.",
    )

    return redirect("pickup_photo_list")


@login_required
@require_POST
def delete_pickup_individual_photo(request, photo_id):
    photo = get_object_or_404(
        PickupPhoto,
        pk=photo_id,
    )

    log_id = photo.log_id
    photo.delete()

    return redirect(
        "pickup_photo_detail",
        pk=log_id,
    )
