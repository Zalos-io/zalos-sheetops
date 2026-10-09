"""Pinned public inputs; workbook contents are data, never executable instructions."""

import hashlib
import json
import tarfile
import urllib.request
from datetime import date, timedelta
from pathlib import Path
from random import Random

EXCELBENCH_REV = "5f71713413c43f3733e303ba8a4eed6a756760cd"
READBENCH_REV = "d8ddc2b5cfab51d53e196015d7037b1d5f463a8a"
VERIFIED_REV = "ab0b742b0fc95b946f212d80ac7771b5531272e4"
HEADERS = [
    "id",
    "sku",
    "name",
    "category",
    "qty",
    "price",
    "weight",
    "ratio",
    "active",
    "created",
    "region",
    "note",
]


def file_sha256(path: Path) -> str:
    with path.open("rb") as source:
        return hashlib.file_digest(source, "sha256").hexdigest()


def download(url: str, destination: Path) -> dict:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if not destination.exists():
        with urllib.request.urlopen(url, timeout=60) as source:
            with destination.open("wb") as target:
                while block := source.read(1024 * 1024):
                    target.write(block)
    return {
        "path": str(destination),
        "url": url,
        "sha256": file_sha256(destination),
    }


def prepare_public_inputs(root: Path, archive: Path | None = None) -> dict:
    base = f"https://raw.githubusercontent.com/SynthGL/ExcelBench/{EXCELBENCH_REV}"
    records = []
    manifest_path = root / "excelbench/manifest.json"
    records.append(download(f"{base}/fixtures/excel/manifest.json", manifest_path))
    records.append(download(f"{base}/LICENSE", root / "excelbench/LICENSE"))
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    supported = {
        "cell_values",
        "formulas",
        "multiple_sheets",
        "merged_cells",
        "hyperlinks",
        "number_formats",
        "tables",
    }
    for entry in manifest["files"]:
        if entry["feature"] in supported:
            relative = entry["path"]
            records.append(
                download(
                    f"{base}/fixtures/excel/{relative}", root / "excelbench" / relative
                )
            )
    readbase = f"https://raw.githubusercontent.com/nazmul284/excel-read-benchmark/{READBENCH_REV}"
    records.append(
        download(f"{readbase}/bench/make_files.py", root / "readbench/make_files.py")
    )
    records.append(download(f"{readbase}/LICENSE", root / "readbench/LICENSE"))
    if archive is None:
        archive = root / "verified400.tar.gz"
        url = (
            f"https://huggingface.co/datasets/KAKA22/SpreadsheetBench/resolve/{VERIFIED_REV}"
            "/spreadsheetbench_verified_400.tar.gz"
        )
        records.append(download(url, archive))
    else:
        with archive.open("rb") as source:
            records.append(
                {
                    "path": str(archive),
                    "sha256": hashlib.file_digest(source, "sha256").hexdigest(),
                    "source": "KAKA22/SpreadsheetBench",
                    "revision": VERIFIED_REV,
                }
            )
    with tarfile.open(archive) as source:
        source.extractall(root / "verified", filter="data")
    result = {
        "schemaVersion": 1,
        "excelbenchRevision": EXCELBENCH_REV,
        "readbenchRevision": READBENCH_REV,
        "verifiedRevision": VERIFIED_REV,
        "inputs": records,
    }
    (root / "sources.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    return result


def generate_readbench(path: Path, rows: int) -> None:
    """MIT readbench adaptation at READBENCH_REV; same seed, types and layout."""
    import xlsxwriter

    path.parent.mkdir(parents=True, exist_ok=True)
    rng = Random(11)
    with xlsxwriter.Workbook(
        str(path), {"constant_memory": True, "default_date_format": "yyyy-mm-dd"}
    ) as book:
        sheet = book.add_worksheet("data")
        sheet.write_row(0, 0, HEADERS)
        for row in range(1, rows + 1):
            values = [
                row,
                f"SKU-{row:08d}",
                f"item name number {row}",
                ["alpha", "beta", "gamma", "delta", "epsilon"][row % 5],
                rng.randint(1, 500),
                round(rng.uniform(1, 9999), 2),
                round(rng.uniform(0.1, 50), 3),
                rng.random(),
                row % 3 == 0,
                date(2020, 1, 1) + timedelta(days=row % 2000),
                ["north", "south", "east", "west"][row % 4],
                "lorem ipsum dolor sit amet " * 2,
            ]
            sheet.write_row(row, 0, values)
