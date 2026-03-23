from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

from fileops.models import OperationStatus
from fileops.operations import CommonOptions, copy_items, delete_items, move_items, rename_items


def test_copy_dry_run_does_not_modify_fs() -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        src = root / "source.txt"
        src.write_text("hello", encoding="utf-8")

        options = CommonOptions(workspace=root, dry_run=True, overwrite="never")
        results = copy_items([src], root / "target.txt", options)

        assert len(results) == 1
        assert results[0].status == OperationStatus.DRY_RUN
        assert not (root / "target.txt").exists()


def test_move_with_auto_rename_policy() -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        src = root / "to-move.txt"
        src.write_text("hello", encoding="utf-8")
        existing = root / "dest.txt"
        existing.write_text("existing", encoding="utf-8")

        options = CommonOptions(workspace=root, dry_run=False, overwrite="rename")
        results = move_items([src], root / "dest.txt", options)

        assert len(results) == 1
        assert results[0].status == OperationStatus.SUCCESS
        assert (root / "dest_1.txt").exists()
        assert existing.exists()


def test_rename_pattern_applies_index() -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        src = root / "sample.txt"
        src.write_text("hello", encoding="utf-8")

        options = CommonOptions(workspace=root, dry_run=False, overwrite="never")
        results = rename_items([src], pattern="{stem}_{index}{ext}", start_index=7, options=options)

        assert len(results) == 1
        assert results[0].status == OperationStatus.SUCCESS
        assert (root / "sample_7.txt").exists()


def test_delete_hard_removes_file() -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        src = root / "to-delete.txt"
        src.write_text("bye", encoding="utf-8")

        results = delete_items([src], workspace=root, dry_run=False, use_trash=False)

        assert len(results) == 1
        assert results[0].status == OperationStatus.SUCCESS
        assert not src.exists()
