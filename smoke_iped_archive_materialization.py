"""Smoke test for materializing a PDF attachment from an embedded-7-Zip archive."""

import sqlite3
import subprocess
import tempfile
import os
from pathlib import Path

from utilities_general import _seven_zip_executable, materialize_iped_archive_link


def main():
    with tempfile.TemporaryDirectory(prefix="fera-archive-smoke-") as temporary:
        bundle_root = Path(temporary) / "88886-26" / "Validador" / "FERA"
        fera_dir = bundle_root / "FERA"
        fera_dir.mkdir(parents=True)
        fera_db = fera_dir / "fera-smoke.db"
        connection = sqlite3.connect(fera_db)
        connection.close()

        member = "files/attachments/chat/mensagem.txt"
        expected = os.urandom(4096)
        source_root = bundle_root / "source"
        source_file = source_root / Path(member)
        source_file.parent.mkdir(parents=True)
        source_file.write_bytes(expected)
        seven_zip = _seven_zip_executable()
        if(seven_zip is None):
            raise AssertionError("The embedded 7-Zip runtime was not found")
        archive = bundle_root / "anexo.zip"
        created = subprocess.run(
            [seven_zip, "a", "-tzip", "-v1k", str(archive), "files"],
            cwd=source_root,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        selected_archive = Path(str(archive) + ".001")
        if(created.returncode != 0 or not selected_archive.is_file() or not Path(str(archive) + ".002").is_file()):
            raise AssertionError("The smoke fixture could not create a split ZIP archive")

        target = bundle_root / Path(member)
        first = materialize_iped_archive_link(str(target), str(fera_db), str(selected_archive), member)
        if(first != str(target) or target.read_bytes() != expected):
            raise AssertionError("The selected archive did not materialize the PDF-linked path")

        target.unlink()
        second = materialize_iped_archive_link(str(target), str(fera_db), link_reference=member)
        if(second != str(target) or target.read_bytes() != expected):
            raise AssertionError("The remembered archive was not reused")

        enveloped_member = "88886-26-Anexo/Anexo/Eq01/files/attachments/chat/PTT-20260118-WA0005.opus"
        enveloped_link = "files/attachments/chat/PTT-20260118-WA0005.opus"
        enveloped_expected = os.urandom(4096)
        enveloped_source_root = bundle_root / "source-enveloped"
        enveloped_source_file = enveloped_source_root / Path(enveloped_member)
        enveloped_source_file.parent.mkdir(parents=True)
        enveloped_source_file.write_bytes(enveloped_expected)
        enveloped_archive = bundle_root / "REP88886-26-Anexo.zip"
        created = subprocess.run(
            [seven_zip, "a", "-tzip", "-v1k", str(enveloped_archive), "88886-26-Anexo"],
            cwd=enveloped_source_root,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        selected_enveloped_archive = Path(str(enveloped_archive) + ".001")
        if(created.returncode != 0 or not selected_enveloped_archive.is_file()):
            raise AssertionError("The enveloped smoke fixture could not create a split ZIP archive")

        enveloped_target = bundle_root / "Anexo" / "Eq01" / Path(enveloped_link)
        third = materialize_iped_archive_link(
            str(enveloped_target),
            str(fera_db),
            str(selected_enveloped_archive),
            enveloped_link,
        )
        if(third != str(enveloped_target) or enveloped_target.read_bytes() != enveloped_expected):
            raise AssertionError("The enveloped ZIP member did not materialize the PDF-linked path")

        legacy_member = "Exportados/arquivos/legacy-audio.opus"
        legacy_expected = os.urandom(4096)
        legacy_source_root = bundle_root / "source-legacy"
        legacy_source_file = legacy_source_root / Path(legacy_member)
        legacy_source_file.parent.mkdir(parents=True)
        legacy_source_file.write_bytes(legacy_expected)
        legacy_archive = bundle_root / "legacy-anexo.zip"
        created = subprocess.run(
            [seven_zip, "a", "-tzip", "-v1k", str(legacy_archive), "Exportados"],
            cwd=legacy_source_root,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        selected_legacy_archive = Path(str(legacy_archive) + ".001")
        if(created.returncode != 0 or not selected_legacy_archive.is_file()):
            raise AssertionError("The legacy smoke fixture could not create a split ZIP archive")

        legacy_missing_path = bundle_root.parent / "Exportados" / "arquivos" / "legacy-audio.opus"
        legacy_target = bundle_root / Path(legacy_member)
        fourth = materialize_iped_archive_link(
            str(legacy_missing_path),
            str(fera_db),
            str(selected_legacy_archive),
            "../../../Exportados/arquivos/legacy-audio.opus",
        )
        if(fourth != str(legacy_target) or legacy_target.read_bytes() != legacy_expected):
            raise AssertionError("The legacy ZIP member was not saved inside the bundle")

        legacy_target.unlink()
        fifth = materialize_iped_archive_link(
            str(legacy_missing_path),
            str(fera_db),
            link_reference="../../../Exportados/arquivos/legacy-audio.opus",
        )
        if(fifth != str(legacy_target) or legacy_target.read_bytes() != legacy_expected):
            raise AssertionError("The remembered legacy archive was not reused")

        connection = sqlite3.connect(fera_db)
        try:
            persisted_archive_sources = connection.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name='Anexo_Eletronico_Iped_Archive_Sources'"
            ).fetchone()
            if(persisted_archive_sources is not None):
                raise AssertionError("Archive source selection must not be persisted in the FERA database")
        finally:
            connection.close()

    print("FERA archive materialization smoke test: PASS")


if __name__ == "__main__":
    main()
