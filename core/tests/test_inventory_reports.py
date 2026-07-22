from decimal import Decimal
from io import BytesIO

from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import SimpleTestCase
from openpyxl import Workbook

from core.report_views import (
    ReportUploadError,
    combine_reports,
    read_make_and_hold_report,
)


def workbook_upload(name, report_rows, warehouse):
    workbook = Workbook()
    report = workbook.active
    report.title = "Report"

    for row in report_rows:
        report.append(row)

    parameters = workbook.create_sheet("Parameters")
    parameters.append(["Parameters"])
    parameters.append([])
    parameters.append(["Warehouse", str(warehouse)])

    output = BytesIO()
    workbook.save(output)
    workbook.close()

    return SimpleUploadedFile(
        name,
        output.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
    )


def bin_report(name, warehouse, log_number):
    return workbook_upload(
        name,
        [
            [
                "Warehouse",
                "Description",
                "Log Number",
                "Bin Location",
                None,
                "Case",
                "Heat",
                "Weight",
                "Pieces",
                "Inventory Qty",
                None,
                "Weight",
                "Pieces",
            ],
            [
                str(warehouse),
                f"Warehouse {warehouse} stock",
                log_number,
                "F",
                f"{warehouse}A1",
                None,
                None,
                125.5,
                3,
                None,
                None,
                150.5,
                4,
            ],
        ],
        warehouse,
    )


def mh_report(
    name,
    warehouse,
    log_number=None,
    on_hand_weight=None,
    on_hand_pieces=None,
):
    header = [None] * 14
    header[0] = "Warehouse"
    header[7] = "Log"
    header[8] = "Description"
    header[11] = "Bin"
    header[12] = "Qty"
    header[13] = "Pcs"
    rows = [header]

    if log_number:
        data = [None] * 14
        data[0] = str(warehouse)
        data[7] = log_number
        data[8] = f"Warehouse {warehouse} MH stock"
        data[11] = f"{warehouse}MH"
        data[12] = on_hand_weight
        data[13] = on_hand_pieces
        rows.append(data)

    return workbook_upload(name, rows, warehouse)


def remarks_report(
    name,
    warehouse,
    log_number,
    remark,
):
    header = [None] * 26
    header[4] = "Log"
    header[23] = "Remarks 1"
    header[24] = "Remarks 2"
    header[25] = "Remarks 3"

    data = [None] * 26
    data[0] = str(warehouse)
    data[4] = log_number
    data[23] = remark

    return workbook_upload(
        name,
        [header, data],
        warehouse,
    )


class InventoryReportMergeTests(SimpleTestCase):
    def test_six_reports_merge_in_order_and_keep_remarks_separate(self):
        rows = combine_reports(
            bin_report_20=bin_report(
                "bin-20.xlsx",
                20,
                "A200",
            ),
            bin_report_21=bin_report(
                "bin-21.xlsx",
                21,
                "A210",
            ),
            mh_report_20=mh_report(
                "mh-20.xlsx",
                20,
                "A205",
                on_hand_weight=725.5,
                on_hand_pieces=12,
            ),
            mh_report_21=mh_report(
                "mh-21.xlsx",
                21,
            ),
            remarks_report_20=remarks_report(
                "remarks-20.xlsx",
                20,
                "A205",
                "Warehouse 20 note",
            ),
            remarks_report_21=remarks_report(
                "remarks-21.xlsx",
                21,
                "A210",
                "Warehouse 21 note",
            ),
        )

        self.assertEqual(
            [row["log_number"] for row in rows],
            ["A200", "A205", "A210"],
        )
        self.assertEqual(rows[1]["remarks"], "Warehouse 20 note")
        self.assertEqual(rows[2]["remarks"], "Warehouse 21 note")
        self.assertIsNone(rows[1]["available_weight"])
        self.assertIsNone(rows[1]["available_pieces"])
        self.assertEqual(
            rows[1]["on_hand_weight"],
            Decimal("725.50"),
        )
        self.assertEqual(rows[1]["on_hand_pieces"], 12)

    def test_empty_mh_report_is_valid_for_expected_warehouse(self):
        rows = read_make_and_hold_report(
            mh_report("mh-21.xlsx", 21),
            expected_warehouse=21,
        )

        self.assertEqual(rows, [])

    def test_empty_mh_report_rejects_wrong_warehouse_slot(self):
        with self.assertRaises(ReportUploadError):
            read_make_and_hold_report(
                mh_report("mh-20.xlsx", 20),
                expected_warehouse=21,
            )