import io
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook

from zalos_sheetops.delivery_classification import classify_delivery


def test_workbook_requires_receipt_regardless_of_name() -> None:
    output = io.BytesIO()
    Workbook().save(output)
    for name in ("output.xlsx", "output.bin", "output.zip", "output.mp4"):
        assert (
            classify_delivery(name, io.BytesIO(output.getvalue()))["action"]
            == "require_receipt"
        )
    assert (
        classify_delivery("output.bin", io.BytesIO(b"prefix" + output.getvalue()))[
            "action"
        ]
        == "require_receipt"
    )


def test_other_files_keep_delivery_and_oversized_metadata_blocks() -> None:
    assert classify_delivery("notes.txt", io.BytesIO(b"hello"))["action"] == "allow"
    output = io.BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        archive.writestr("notes.txt", "hello")
    assert (
        classify_delivery("archive.zip", io.BytesIO(output.getvalue()))["action"]
        == "allow"
    )
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", b" " * (256 * 1024 + 1))
    assert (
        classify_delivery("archive.bin", io.BytesIO(output.getvalue()))["action"]
        == "block"
    )


def test_damaged_archive_is_unknown_and_cannot_skip_validation() -> None:
    output = io.BytesIO()
    Workbook().save(output)
    assert (
        classify_delivery("output.bin", io.BytesIO(output.getvalue()[:-22]))["action"]
        == "block"
    )
    assert classify_delivery("notes.txt", io.BytesIO(b"PK text"))["action"] == "allow"
