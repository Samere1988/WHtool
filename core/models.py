from django.db import models
from django.utils import timezone
import uuid


class CustomerList(models.Model):
    # We changed managed to True so you can edit these in the app
    customer_id = models.TextField(db_column='Customer ID', primary_key=True)
    customer_name = models.TextField(db_column='Customer Name')
    address = models.TextField(db_column='Address')
    city = models.TextField(db_column='City')
    province = models.TextField(db_column='Province')
    postal_code = models.CharField(max_length=10, blank=True, null=True)
    region = models.TextField(db_column='Region')

    class Meta:
        managed = True
        db_table = 'Customer List'

    def __str__(self):
        return f"{self.customer_name} - {self.postal_code}"

class RunSheet(models.Model):
    # Added an ID field so Django can track individual rows easily
    id = models.AutoField(primary_key=True)
    customer_id = models.TextField(db_column='Customer ID', blank=True, null=True)
    order_number = models.CharField(max_length=20, default="W", blank=True, null=True)
    prepared_by = models.CharField(max_length=100, blank=True, null=True)
    line_items = models.IntegerField(default=0, blank=True, null=True)
    customer_name = models.TextField(db_column='Customer Name', blank=True, null=True)
    address = models.TextField(db_column='Address', blank=True, null=True)
    city = models.TextField(db_column='City', blank=True, null=True)
    driver_name = models.CharField(max_length=100, null=True, blank=True)
    region = models.TextField(db_column='Region', blank=True, null=True)
    weight = models.IntegerField(db_column='Weight', blank=True, null=True)
    skids = models.IntegerField(db_column='Skids', blank=True, null=True)
    bundles = models.IntegerField(db_column='Bundles', blank=True, null=True)
    coils = models.IntegerField(db_column='Coils', blank=True, null=True)
    closing_time = models.TextField(db_column='Closing Time', blank=True, null=True)
    is_pickup = models.BooleanField(default=False)
    is_return = models.BooleanField(default=False)
    created_at = models.DateTimeField(default=timezone.now)
    postal_code = models.CharField(max_length=20, blank=True, null=True)
    load_index = models.IntegerField(default=0)
    shipping_date = models.DateField(default=timezone.now)

    # Transport-company import fields. These do not replace the original region data.
    transport_run_name = models.CharField(max_length=100, blank=True, null=True)
    transport_driver = models.CharField(max_length=100, blank=True, null=True)
    transport_truck = models.CharField(max_length=100, blank=True, null=True)
    transport_start_time = models.CharField(max_length=50, blank=True, null=True)
    transport_stop_number = models.IntegerField(blank=True, null=True)
    transport_import_batch = models.ForeignKey(
        "TransportImportBatch",
        blank=True,
        null=True,
        related_name="run_sheet_items",
        on_delete=models.SET_NULL,
    )

    class Meta:
        managed = True
        db_table = 'Run Sheet'

class OrderArchive(models.Model):
    order_number = models.CharField(max_length=20)
    customer_id = models.CharField(max_length=50)
    customer_name = models.TextField()
    prepared_by = models.CharField(max_length=100)
    line_items = models.IntegerField(default=0)
    bar_prep = models.CharField(max_length=255, blank=True, null=True)
    bar_lines = models.IntegerField(default=0)
    sheet_prep = models.CharField(max_length=255, blank=True, null=True)
    sheet_lines = models.IntegerField(default=0)
    covering_prep = models.CharField(max_length=255, blank=True, null=True)
    covering_lines = models.IntegerField(default=0)
    skids = models.IntegerField(default=0)
    bundles = models.IntegerField(default=0)
    coils = models.IntegerField(default=0)
    weight = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    region = models.CharField(max_length=100, blank=True, null=True)
    is_tallied = models.BooleanField(default=False)

    class Meta:
        db_table = 'Order Archive'


class FinalizedRunSheet(models.Model):
    customer_name = models.TextField()
    region = models.CharField(max_length=100)
    order_numbers = models.TextField() # Combined string (e.g., W111 / W222)
    weight = models.IntegerField(default=0)
    skids = models.IntegerField(default=0)
    bundles = models.IntegerField(default=0)
    coils = models.IntegerField(default=0)
    finalized_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'finalized_run_sheets'

class DailyRunSheetCommit(models.Model):
    shipping_date = models.DateField()
    committed_at = models.DateTimeField(auto_now_add=True)

    total_weight = models.IntegerField(default=0)
    total_skids = models.IntegerField(default=0)
    total_bundles = models.IntegerField(default=0)
    total_coils = models.IntegerField(default=0)

    def __str__(self):
        return f"Run Sheet Commit - {self.shipping_date}"


class DailyRunSheetEntry(models.Model):
    commit = models.ForeignKey(
        DailyRunSheetCommit,
        related_name="entries",
        on_delete=models.CASCADE
    )

    original_run_sheet_id = models.IntegerField(blank=True, null=True)

    customer_id = models.CharField(max_length=50, blank=True, null=True)
    customer_name = models.TextField(blank=True, null=True)
    order_number = models.CharField(max_length=50, blank=True, null=True)

    address = models.TextField(blank=True, null=True)
    city = models.CharField(max_length=100, blank=True, null=True)
    postal_code = models.CharField(max_length=20, blank=True, null=True)

    region = models.CharField(max_length=100, blank=True, null=True)
    driver_name = models.CharField(max_length=100, blank=True, null=True)
    load_index = models.IntegerField(default=0)

    closing_time = models.CharField(max_length=50, blank=True, null=True)

    weight = models.IntegerField(default=0)
    skids = models.IntegerField(default=0)
    bundles = models.IntegerField(default=0)
    coils = models.IntegerField(default=0)

    is_pickup = models.BooleanField(default=False)
    is_return = models.BooleanField(default=False)

    prepared_by = models.CharField(max_length=255, blank=True, null=True)
    line_items = models.IntegerField(default=0)

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["region", "load_index", "customer_name", "order_number"]


class EmployeeDailyStat(models.Model):
    commit = models.ForeignKey(
        DailyRunSheetCommit,
        related_name="employee_stats",
        on_delete=models.CASCADE
    )

    employee_name = models.CharField(max_length=100)

    orders_picked = models.IntegerField(default=0)
    total_lines = models.IntegerField(default=0)

    bar_orders = models.IntegerField(default=0)
    bar_lines = models.IntegerField(default=0)

    sheet_orders = models.IntegerField(default=0)
    sheet_lines = models.IntegerField(default=0)

    covering_orders = models.IntegerField(default=0)
    covering_lines = models.IntegerField(default=0)

    class Meta:
        unique_together = ("commit", "employee_name")
        ordering = ["employee_name"]

    def __str__(self):
        return f"{self.employee_name} - {self.commit.shipping_date}"


class TransportImportBatch(models.Model):
    STATUS_CHOICES = [
        ("review", "Review"),
        ("applied", "Applied"),
        ("undone", "Undone"),
        ("failed", "Failed"),
    ]

    shipping_date = models.DateField()
    original_filename = models.CharField(max_length=255, blank=True)
    uploaded_by = models.CharField(max_length=150, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    applied_at = models.DateTimeField(blank=True, null=True)
    undone_at = models.DateTimeField(blank=True, null=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="review")
    notes = models.TextField(blank=True)

    class Meta:
        ordering = ["-created_at"]

    def __str__(self):
        return f"Transport import {self.id} - {self.shipping_date} - {self.status}"


class TransportImportRow(models.Model):
    STATUS_CHOICES = [
        ("matched", "Matched"),
        ("review", "Needs Review"),
        ("unmatched", "Unmatched"),
    ]

    batch = models.ForeignKey(TransportImportBatch, related_name="rows", on_delete=models.CASCADE)
    sort_order = models.IntegerField(default=0)
    source_sheet_name = models.CharField(max_length=100, blank=True)
    source_row_number = models.IntegerField(default=0)

    imported_run_name = models.CharField(max_length=100, blank=True)
    imported_driver = models.CharField(max_length=100, blank=True)
    imported_truck = models.CharField(max_length=100, blank=True)
    imported_start_time = models.CharField(max_length=50, blank=True)
    imported_stop_number = models.IntegerField(default=0)
    imported_customer_name = models.CharField(max_length=255, blank=True)
    imported_city = models.CharField(max_length=150, blank=True)

    matched_run_sheet_ids = models.TextField(blank=True)
    confidence = models.FloatField(default=0)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="unmatched")

    class Meta:
        ordering = ["sort_order", "id"]

    def matched_id_list(self):
        return [x.strip() for x in (self.matched_run_sheet_ids or "").split(",") if x.strip()]


class TransportImportPreviousState(models.Model):
    batch = models.ForeignKey(TransportImportBatch, related_name="previous_states", on_delete=models.CASCADE)
    run_sheet_id = models.IntegerField()

    previous_transport_run_name = models.CharField(max_length=100, blank=True, null=True)
    previous_transport_driver = models.CharField(max_length=100, blank=True, null=True)
    previous_transport_truck = models.CharField(max_length=100, blank=True, null=True)
    previous_transport_start_time = models.CharField(max_length=50, blank=True, null=True)
    previous_transport_stop_number = models.IntegerField(blank=True, null=True)
    previous_transport_import_batch_id = models.IntegerField(blank=True, null=True)
    previous_driver_name = models.CharField(max_length=100, blank=True, null=True)
    previous_load_index = models.IntegerField(default=0)

    class Meta:
        unique_together = ("batch", "run_sheet_id")


class Container(models.Model):
    container_number = models.CharField(max_length=100)
    date_received = models.DateField(
        default=timezone.localdate
    )
    unloaded_by = models.CharField(max_length=100, blank=True, null=True)
    unloaded_at = models.CharField(max_length=20, blank=True, null=True)


    def __str__(self):
        return self.container_number

class ContainerPhoto(models.Model):
    container = models.ForeignKey(Container, related_name='photos', on_delete=models.CASCADE)
    image = models.ImageField(upload_to='container_photos/')
    uploaded_at = models.DateTimeField(auto_now_add=True)



class OutboundLoad(models.Model):
    # Just a clean text field now, no choices!
    truck_name = models.CharField(max_length=100)
    date_loaded = models.DateField(
        default=timezone.localdate
    )
    loaded_by = models.CharField(max_length=100, blank=True, null=True)


    def __str__(self):
        return f"{self.truck_name} - {self.date_loaded}"

class OutboundPhoto(models.Model):
    load = models.ForeignKey(OutboundLoad, related_name='photos', on_delete=models.CASCADE)
    image = models.ImageField(upload_to='outbound_photos/')
    uploaded_at = models.DateTimeField(auto_now_add=True)

class Vendor(models.Model):
    name = models.CharField(max_length=255)
    address = models.CharField(max_length=255)
    city = models.CharField(max_length=100)
    postal_code = models.CharField(max_length=20)
    region = models.CharField(max_length=100)

    def __str__(self):
        return self.name


class PickupLog(models.Model):
    customer_name = models.CharField(max_length=255)
    customer_id = models.CharField(max_length=50, blank=True, null=True)
    order_number = models.CharField(max_length=50)
    date_completed = models.DateField(auto_now_add=True)

    # Load Details
    weight = models.IntegerField(default=0)
    skids = models.IntegerField(default=0)
    bundles = models.IntegerField(default=0)
    coils = models.IntegerField(default=0)

    # Stats Tracking (Same as your Archive)
    bar_lines = models.IntegerField(default=0)
    sheet_lines = models.IntegerField(default=0)
    covering_lines = models.IntegerField(default=0)

    # This stores the names of the guys who did the work
    bar_prep = models.CharField(max_length=255, blank=True)
    sheet_prep = models.CharField(max_length=255, blank=True)
    covering_prep = models.CharField(max_length=255, blank=True)

    def __str__(self):
        return f"{self.order_number} - {self.customer_name}"


class PickupPhotoLog(models.Model):
    customer_name = models.CharField(max_length=255)
    order_number = models.CharField(max_length=100)
    date_picked_up = models.DateField(default=timezone.now)
    loaded_by = models.CharField(max_length=100, blank=True, null=True)


    def __str__(self):
        return f"{self.customer_name} - {self.order_number}"

class PickupPhoto(models.Model):
    log = models.ForeignKey(PickupPhotoLog, related_name='photos', on_delete=models.CASCADE)
    image = models.ImageField(upload_to='pickup_photos/%Y/%m/%d/')
    uploaded_at = models.DateTimeField(auto_now_add=True)

class RegionRunInfo(models.Model):
    shipping_date = models.DateField()
    region = models.CharField(max_length=100)

    driver_name = models.CharField(max_length=100, blank=True, null=True)
    start_time = models.CharField(max_length=50, blank=True, null=True)

    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        unique_together = ("shipping_date", "region")
        ordering = ["shipping_date", "region"]

    def __str__(self):
        return f"{self.shipping_date} - {self.region} - {self.driver_name or 'Unassigned'}"

class ExtraRun(models.Model):
    shipping_date = models.DateField()
    name = models.CharField(max_length=100)

    created_at = models.DateTimeField(
        auto_now_add=True,
    )

    class Meta:
        ordering = [
            "shipping_date",
            "created_at",
            "id",
        ]

        constraints = [
            models.UniqueConstraint(
                fields=[
                    "shipping_date",
                    "name",
                ],
                name="unique_extra_run_name_per_date",
            ),
        ]

    def __str__(self):
        return (
            f"{self.shipping_date} - "
            f"{self.name}"
        )

class BillOfLadingCustomer(models.Model):
    name = models.CharField(max_length=255)
    address = models.TextField(blank=True)
    city = models.CharField(max_length=100, blank=True)
    province = models.CharField(max_length=50, blank=True)
    postal_code = models.CharField(max_length=20, blank=True)
    account_number = models.CharField(max_length=100, blank=True)


class BillOfLading(models.Model):
    bol_number = models.CharField(max_length=100, blank=True)
    bol_date = models.DateField()
    consignor_name = models.CharField(max_length=255, blank=True)
    consignor_address = models.TextField(blank=True)
    consignor_account_number = models.CharField(max_length=100, blank=True)

    consignee_name = models.CharField(max_length=255, blank=True)
    consignee_address = models.TextField(blank=True)

    declared_value = models.CharField(max_length=100, blank=True)
    freight_collect = models.BooleanField(default=False)
    freight_prepaid = models.BooleanField(default=False)
    cod = models.BooleanField(default=False)
    cod_amount = models.CharField(max_length=100, blank=True)
    other_charges = models.CharField(max_length=100, blank=True)
    total_charges = models.CharField(max_length=100, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

class BillOfLadingLine(models.Model):
    class PackageType(models.TextChoices):
        SKID = "skid", "Skid(s)"
        COIL = "coil", "Coil(s)"
        BUNDLE = "bundle", "Bundle(s)"

    bill = models.ForeignKey(
        BillOfLading,
        related_name="lines",
        on_delete=models.CASCADE,
    )

    order_po_number = models.CharField(
        max_length=100,
        blank=True,
    )

    description = models.TextField(
        blank=True,
    )

    total_packages = models.CharField(
        max_length=50,
        blank=True,
    )

    package_type = models.CharField(
        max_length=10,
        choices=PackageType.choices,
        blank=True,
    )

    weight = models.CharField(
        max_length=50,
        blank=True,
    )

class InventoryReport(models.Model):
    uploaded_at = models.DateTimeField(auto_now_add=True)
    uploaded_by = models.CharField(max_length=150, blank=True)

    warehouse_20_filename = models.CharField(
        max_length=255,
        blank=True,
    )

    warehouse_21_filename = models.CharField(
        max_length=255,
        blank=True,
    )

    remarks_filename = models.CharField(
        max_length=255,
        blank=True,
    )

    class Meta:
        ordering = ["-uploaded_at"]

    def __str__(self):
        return (
            f"Inventory report - "
            f"{self.uploaded_at:%Y-%m-%d %H:%M}"
        )


class InventoryItem(models.Model):
    report = models.ForeignKey(
        InventoryReport,
        related_name="items",
        on_delete=models.CASCADE,
    )

    warehouse = models.CharField(max_length=10)

    log_number = models.CharField(
        max_length=100,
        db_index=True,
    )

    description = models.TextField(blank=True)

    bin_location = models.CharField(
        max_length=255,
        blank=True,
    )

    available_pieces = models.IntegerField(
        blank=True,
        null=True,
    )

    available_weight = models.DecimalField(
        max_digits=18,
        decimal_places=2,
        blank=True,
        null=True,
    )

    on_hand_pieces = models.IntegerField(
        blank=True,
        null=True,
    )

    on_hand_weight = models.DecimalField(
        max_digits=18,
        decimal_places=2,
        blank=True,
        null=True,
    )

    remarks = models.TextField(blank=True)

    class Meta:
        ordering = [
            "log_number",
            "warehouse",
            "bin_location",
            "id",
        ]

    def __str__(self):
        return self.log_number


class HymusTransferItem(models.Model):
    warehouse = models.CharField(
        max_length=10,
        blank=True,
    )

    log_number = models.CharField(max_length=100)

    description = models.TextField(blank=True)

    bin_location = models.CharField(
        max_length=255,
        blank=True,
    )

    notes = models.TextField(blank=True)

    added_by = models.CharField(
        max_length=150,
        blank=True,
    )

    added_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["added_at", "id"]

        constraints = [
            models.UniqueConstraint(
                fields=[
                    "warehouse",
                    "log_number",
                    "bin_location",
                ],
                name=(
                    "unique_hymus_transfer_"
                    "inventory_location"
                ),
            ),
        ]

    def __str__(self):
        return (
            f"{self.log_number} - "
            f"{self.bin_location}"
        )


class CycleCount(models.Model):
    class Category(models.TextChoices):
        SHEETS = "sheets", "Sheets"
        LONG_PRODUCTS = (
            "long_products",
            "Long Products",
        )

    batch_id = models.UUIDField(
        default=uuid.uuid4,
        editable=False,
        db_index=True,
    )

    category = models.CharField(
        max_length=30,
        choices=Category.choices,
    )

    rack = models.CharField(
        max_length=100,
        db_index=True,
    )

    source_inventory_uploaded_at = models.DateTimeField(
        blank=True,
        null=True,
    )

    created_by = models.CharField(
        max_length=150,
        blank=True,
    )

    created_at = models.DateTimeField(
        auto_now_add=True,
    )

    completed_by = models.CharField(
        max_length=150,
        blank=True,
    )

    completed_at = models.DateTimeField(
        blank=True,
        null=True,
    )

    class Meta:
        ordering = ["-created_at", "rack", "id"]

    def __str__(self):
        return (
            f"{self.get_category_display()} - "
            f"{self.rack} - "
            f"{self.created_at:%Y-%m-%d}"
        )


class CycleCountItem(models.Model):
    cycle_count = models.ForeignKey(
        CycleCount,
        related_name="items",
        on_delete=models.CASCADE,
    )

    position = models.PositiveIntegerField(
        default=0,
    )

    log_number = models.CharField(
        max_length=100,
    )

    description = models.TextField(blank=True)

    bin_location = models.CharField(
        max_length=255,
        blank=True,
    )

    on_hand_pieces = models.IntegerField(
        blank=True,
        null=True,
    )

    on_hand_weight = models.DecimalField(
        max_digits=18,
        decimal_places=2,
        blank=True,
        null=True,
    )

    notes = models.TextField(blank=True)

    class Meta:
        ordering = ["position", "id"]

    def __str__(self):
        return (
            f"{self.log_number} - "
            f"{self.bin_location}"
        )

class CycleCountCounter(models.Model):
    cycle_count = models.ForeignKey(
        CycleCount,
        related_name="counters",
        on_delete=models.CASCADE,
    )

    employee_name = models.CharField(
        max_length=100,
        db_index=True,
    )

    class Meta:
        ordering = ["employee_name", "id"]

        constraints = [
            models.UniqueConstraint(
                fields=[
                    "cycle_count",
                    "employee_name",
                ],
                name="unique_cycle_count_counter",
            ),
        ]

    def __str__(self):
        return (
            f"{self.employee_name} - "
            f"{self.cycle_count.rack}"
        )