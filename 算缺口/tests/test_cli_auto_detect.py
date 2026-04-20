from __future__ import annotations

import os

from shipment_planner.cli import _list_xlsx_candidates


def test_list_xlsx_candidates_prefers_filename_export_timestamp_over_copy_time(tmp_path):
    older_export_copied_later = tmp_path / "订单_20260213112213_142723699_1.xlsx"
    latest_export_copied_earlier = tmp_path / "订单_20260420084730_156092592_1.xlsx"
    older_export_copied_later.write_bytes(b"")
    latest_export_copied_earlier.write_bytes(b"")
    os.utime(older_export_copied_later, (2000, 2000))
    os.utime(latest_export_copied_earlier, (1000, 1000))

    candidates = _list_xlsx_candidates(tmp_path)

    assert candidates[:2] == [latest_export_copied_earlier, older_export_copied_later]


def test_list_xlsx_candidates_uses_mtime_when_filename_has_no_export_timestamp(tmp_path):
    older_file = tmp_path / "older.xlsx"
    latest_file = tmp_path / "latest.xlsx"
    older_file.write_bytes(b"")
    latest_file.write_bytes(b"")
    os.utime(older_file, (1000, 1000))
    os.utime(latest_file, (2000, 2000))

    candidates = _list_xlsx_candidates(tmp_path)

    assert candidates[:2] == [latest_file, older_file]
