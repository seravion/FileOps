from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

from fileops.document_split import split_documents_by_structure
from fileops.models import OperationStatus
from fileops.operations import CommonOptions, copy_items, delete_items, move_items, rename_items, split_items


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


def test_split_items_by_size() -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        src = root / "large.bin"
        src.write_bytes(b"abcdefghijklmnopqrstuvwxyz")

        options = CommonOptions(workspace=root, dry_run=False, overwrite="never")
        out_dir = root / "chunks"
        results = split_items([src], destination=out_dir, chunk_size_mb=0.00001, options=options)

        assert len(results) == 1
        assert results[0].status == OperationStatus.SUCCESS
        assert (out_dir / "large.part001.bin").exists()
        assert (out_dir / "large.part002.bin").exists()
        assert (out_dir / "large.part003.bin").exists()


def test_doc_split_markdown_by_heading() -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        src = root / "sample.md"
        src.write_text(
            "# Title One\n"
            "Intro line\n"
            "![diagram](img.png)\n"
            "## Sub Two\n"
            "Second section text\n",
            encoding="utf-8",
        )

        out_dir = root / "doc_parts"
        results = split_documents_by_structure(
            sources=[src],
            destination=out_dir,
            workspace=root,
            dry_run=False,
            heading_mode="h1_h2",
            include_image_text=True,
        )

        assert len(results) == 1
        assert results[0].status == OperationStatus.SUCCESS
        assert (out_dir / "sample_split_index.json").exists()

        produced = sorted(out_dir.glob("sample_*.txt"))
        assert len(produced) >= 2
        assert "[图片说明] diagram" in produced[0].read_text(encoding="utf-8")

class _BrokenStyleParagraph:
    @property
    def style(self) -> object:
        raise KeyError("missing style")


class _NormalStyleParagraph:
    class _Style:
        name = "Heading 1"

    @property
    def style(self) -> object:
        return self._Style()


def test_safe_docx_style_name_handles_style_lookup_error() -> None:
    from fileops.document_split import _safe_docx_style_name

    assert _safe_docx_style_name(_BrokenStyleParagraph()) == ""
    assert _safe_docx_style_name(_NormalStyleParagraph()) == "Heading 1"

def test_doc_split_invalid_docx_reports_friendly_error() -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        src = root / "not_a_real_docx.docx"
        src.write_text("this is plain text, not docx package", encoding="utf-8")

        out_dir = root / "doc_parts"
        results = split_documents_by_structure(
            sources=[src],
            destination=out_dir,
            workspace=root,
            dry_run=False,
            heading_mode="h1_h2",
            include_image_text=False,
        )

        assert len(results) == 1
        assert results[0].status == OperationStatus.FAILED
        assert ".docx" in results[0].message
        assert "Word/WPS" in results[0].message

