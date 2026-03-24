from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document

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

        produced = sorted(out_dir.glob("sample_*.md"))
        assert len(produced) >= 2
        assert "[Image Alt] diagram" in produced[0].read_text(encoding="utf-8")

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


def test_doc_split_docx_keeps_docx_and_tables() -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        src = root / "sample.docx"

        doc = Document()
        doc.add_paragraph("封面段落")
        h1 = doc.add_paragraph("第一章")
        h1.style = "Heading 1"
        doc.add_paragraph("第一章正文")
        table = doc.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "A1"
        table.cell(0, 1).text = "B1"
        h2 = doc.add_paragraph("第一章-小节")
        h2.style = "Heading 2"
        doc.add_paragraph("小节正文")
        doc.save(src)

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
        assert results[0].status == OperationStatus.SUCCESS
        assert (out_dir / "sample_split_index.json").exists()

        produced_docx = sorted(out_dir.glob("sample_*.docx"))
        assert len(produced_docx) >= 2
        assert not list(out_dir.glob("sample_*.txt"))

        split_docs = [Document(str(path)) for path in produced_docx]
        assert any("第一章正文" in "\n".join(p.text for p in split_doc.paragraphs) for split_doc in split_docs)
        assert any(len(split_doc.tables) > 0 for split_doc in split_docs)


def test_doc_split_markdown_export_docx() -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        src = root / "sample.md"
        src.write_text(
            "# Title One\n"
            "Body text\n"
            "| Col1 | Col2 |\n"
            "| --- | --- |\n"
            "| A | B |\n",
            encoding="utf-8",
        )

        out_dir = root / "doc_parts"
        results = split_documents_by_structure(
            sources=[src],
            destination=out_dir,
            workspace=root,
            dry_run=False,
            heading_mode="h1_h2",
            include_image_text=False,
            input_format="markdown",
            output_format="docx",
        )

        assert len(results) == 1
        assert results[0].status == OperationStatus.SUCCESS
        produced_docx = sorted(out_dir.glob("sample_*.docx"))
        assert produced_docx

        first_doc = Document(str(produced_docx[0]))
        text_content = "\n".join(paragraph.text for paragraph in first_doc.paragraphs)
        assert "Title One" in text_content


def test_doc_split_docx_export_markdown() -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        src = root / "sample.docx"

        doc = Document()
        heading = doc.add_paragraph("Chapter One")
        heading.style = "Heading 1"
        doc.add_paragraph("Body paragraph")
        table = doc.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "A1"
        table.cell(0, 1).text = "B1"
        doc.save(src)

        out_dir = root / "doc_parts"
        results = split_documents_by_structure(
            sources=[src],
            destination=out_dir,
            workspace=root,
            dry_run=False,
            heading_mode="h1_h2",
            include_image_text=False,
            output_format="md",
        )

        assert len(results) == 1
        assert results[0].status == OperationStatus.SUCCESS
        produced_md = sorted(out_dir.glob("sample_*.md"))
        assert produced_md

        merged_text = "\n".join(path.read_text(encoding="utf-8") for path in produced_md)
        assert "# Chapter One" in merged_text
        assert "| A1 | B1 |" in merged_text


def test_doc_split_import_format_mismatch() -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        src = root / "sample.txt"
        src.write_text("plain text", encoding="utf-8")

        out_dir = root / "doc_parts"
        results = split_documents_by_structure(
            sources=[src],
            destination=out_dir,
            workspace=root,
            dry_run=False,
            heading_mode="h1",
            include_image_text=False,
            input_format="docx",
        )

        assert len(results) == 1
        assert results[0].status == OperationStatus.FAILED
        assert "import format" in results[0].message.lower()
