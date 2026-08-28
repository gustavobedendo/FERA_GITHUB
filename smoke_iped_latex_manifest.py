"""Creates and validates a minimal portable IPED LaTeX bundle for FERA.

The fixture exercises the same resolver invoked by PDF links in ``fera.py``:
one asset is already materialized under ``IPED/Exportados/arquivos`` and the
other exists only as a gzip-compressed blob in an IPED ``storage-*.db``.
"""

import argparse
import gzip
import hashlib
import json
import shutil
import sqlite3
import sys
from pathlib import Path

import fitz

FERA_ROOT = Path(__file__).resolve().parent
if str(FERA_ROOT) not in sys.path:
    sys.path.insert(0, str(FERA_ROOT))

import global_settings
from indexador_fera import build_db_with_reports_commandline
from utilities_general import materialize_iped_archive_link, resolve_iped_latex_manifest_link


def md5_bytes(data):
    return hashlib.md5(data).hexdigest().upper()


def write_storage_db(path, storage_id, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(path) as connection:
        connection.execute("CREATE TABLE t1 (id TEXT PRIMARY KEY, data BLOB)")
        connection.execute("INSERT INTO t1 (id, data) VALUES (?, ?)",
                           (storage_id, gzip.compress(data)))


def write_manifest(path, physical, stored):
    records = [
        {
            "logicalStoredName": "files/already-exported.txt",
            "storedName": "files/already-exported.txt",
            "openTarget": "../Exportados/arquivos/already-exported.txt",
            "backingKind": "exportados-file",
            "backingPath": "Exportados/arquivos/already-exported.txt",
            "md5": md5_bytes(physical),
            "size": len(physical),
            "name": "already-exported.txt",
            "path": "fixture/already-exported.txt",
        },
        {
            "logicalStoredName": "files/storage-only.txt",
            "storedName": "files/storage-only.txt",
            "backingKind": "sqlite-storage-v1",
            "md5": md5_bytes(stored),
            "size": len(stored),
            "name": "storage-only.txt",
            "path": "fixture/storage-only.txt",
            "locator": {
                "type": "sqlite-storage-v1",
                "storageDb": "storage/storage-0.db",
                "id": "SMOKE-STORAGE-001",
            },
        },
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as manifest:
        for record in records:
            manifest.write(json.dumps(record, ensure_ascii=False) + "\n")


def write_smoke_pdf(path):
    document = fitz.open()
    page = document.new_page()
    page.insert_text((72, 72), "FERA / IPED LaTeX - fixture de abertura de arquivos")
    page.insert_text((72, 112), "1. Abrir arquivo ja exportado", fontsize=12)
    page.insert_link({
        "kind": fitz.LINK_LAUNCH,
        "from": fitz.Rect(68, 94, 330, 122),
        "file": "files/already-exported.txt",
    })
    page.insert_text((72, 160), "2. Abrir arquivo somente em storage-0.db", fontsize=12)
    page.insert_link({
        "kind": fitz.LINK_LAUNCH,
        "from": fitz.Rect(68, 142, 390, 170),
        "file": "files/storage-only.txt",
    })
    document.new_page().insert_text((72, 72), "Pagina tecnica da fixture FERA.")
    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(path))
    document.close()


def assert_content(path, expected, label):
    if path is None:
        raise AssertionError(f"{label}: FERA did not resolve the link")
    resolved = Path(path)
    if not resolved.is_file():
        raise AssertionError(f"{label}: resolved path is not a file: {resolved}")
    if resolved.read_bytes() != expected:
        raise AssertionError(f"{label}: recovered bytes do not match the manifest")
    return resolved


def build_fixture(root):
    root = root.resolve()
    if root.exists():
        shutil.rmtree(root)

    physical = b"FERA smoke fixture: exported physical evidence\n"
    stored = b"FERA smoke fixture: sqlite-storage-v1 evidence\n"
    equipment_root = root / "IPED" / "Eq02"
    export_dir = equipment_root / "Exportados" / "arquivos"
    physical_path = export_dir / "already-exported.txt"
    physical_path.parent.mkdir(parents=True, exist_ok=True)
    physical_path.write_bytes(physical)

    storage_id = "SMOKE-STORAGE-001"
    # Every equipment uses the same storage-*.db names.  Eq01 deliberately
    # contains the same storage id with different bytes: resolving the Eq02
    # report must never select Eq01 merely because it is scanned first.
    write_storage_db(root / "IPED" / "Eq01" / "iped" / "storage" / "storage-0.db", storage_id,
                     b"FERA smoke fixture: wrong Eq01 evidence\n")
    # Portable IPED keeps operational storage below the equipment's iped
    # module; the locator is relative to that module.
    write_storage_db(equipment_root / "iped" / "storage" / "storage-0.db", storage_id, stored)
    manifest_path = equipment_root / "RelatoriosPDF" / "sources" / "iped-latex-assets.jsonl"
    write_manifest(manifest_path, physical, stored)

    latex_dir = equipment_root / "RelatoriosPDF" / "sources"
    latex_dir.mkdir(parents=True, exist_ok=True)
    fera_db = root / "FERA" / "fera-smoke.db"
    fera_db.parent.mkdir(parents=True, exist_ok=True)
    report_path = equipment_root / "RelatoriosPDF" / "fera-ipeds-latex-smoke.pdf"
    write_smoke_pdf(report_path)
    global_settings.initiate_variables(commandline=True)
    status = build_db_with_reports_commandline(
        str(fera_db), [str(report_path)], str(manifest_path)
    )
    if status != 0:
        raise RuntimeError(f"FERA could not index the smoke report (status={status})")

    return {
        "root": root,
        "fera_db": fera_db,
        "latex_dir": latex_dir,
        "physical": physical,
        "stored": stored,
        "export_dir": export_dir,
        "report_path": report_path,
    }


def validate_fixture(fixture):
    root = fixture["root"]
    latex_dir = fixture["latex_dir"]
    fera_db = fixture["fera_db"]
    physical = fixture["physical"]
    stored = fixture["stored"]
    export_dir = fixture["export_dir"]

    physical_resolved = assert_content(
        resolve_iped_latex_manifest_link(str(latex_dir / "already-exported.txt"), str(fera_db)),
        physical,
        "exportados-file",
    )
    on_demand_target = export_dir / "storage-only.txt"
    if on_demand_target.exists():
        raise AssertionError("sqlite-storage-v1 fixture was materialized before FERA opened it")
    precheck_resolved = resolve_iped_latex_manifest_link(
        str(latex_dir / "storage-only.txt"),
        str(fera_db),
        materialize_from_storage=False,
    )
    if precheck_resolved is not None or on_demand_target.exists():
        raise AssertionError("manifest precheck materialized a sqlite-storage-v1 fixture")
    storage_resolved = assert_content(
        materialize_iped_archive_link(str(latex_dir / "storage-only.txt"), str(fera_db)),
        stored,
        "sqlite-storage-v1",
    )
    cached_storage = assert_content(
        resolve_iped_latex_manifest_link(str(latex_dir / "storage-only.txt"), str(fera_db)),
        stored,
        "sqlite-storage-v1 cached",
    )
    if storage_resolved != cached_storage:
        raise AssertionError("sqlite-storage-v1: repeated open did not reuse the materialized file")

    print("FERA IPED LaTeX manifest smoke test: PASS")
    print(f"fixture: {root}")
    print(f"physical asset: {physical_resolved}")
    print(f"storage asset: {storage_resolved}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--root",
        type=Path,
        default=Path(r"B:\IPED\target\fera-ipeds-latex-smoke"),
        help="directory where the portable fixture is created",
    )
    parser.add_argument(
        "--prepare-only",
        action="store_true",
        help="create the bundle without opening the storage-only asset",
    )
    arguments = parser.parse_args()
    fixture = build_fixture(arguments.root)
    if arguments.prepare_only:
        print("FERA IPED LaTeX manifest fixture: READY")
        print(f"fixture: {arguments.root.resolve()}")
        print(f"FERA database: {fixture['fera_db']}")
        print(f"PDF para abertura no FERA: {fixture['report_path']}")
        print(f"physical asset: {fixture['export_dir'] / 'already-exported.txt'}")
        print(f"storage-only asset: {fixture['export_dir'] / 'storage-only.txt'} (not materialized)")
        return
    validate_fixture(fixture)


if __name__ == "__main__":
    main()
