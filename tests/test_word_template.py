from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document
from PIL import Image

from fileops.models import OperationStatus
from fileops.word_template import format_word_documents, import_word_template, list_word_templates


def _create_sample_png(path: Path) -> None:
    image = Image.new("RGB", (24, 24), color=(255, 0, 0))
    image.save(path)


def test_template_library_import_and_list(monkeypatch) -> None:
    with TemporaryDirectory() as temp_dir:
        home = Path(temp_dir)
        monkeypatch.setenv("HOME", str(home))
        monkeypatch.setenv("USERPROFILE", str(home))
        monkeypatch.setenv("LOCALAPPDATA", str(home))
        monkeypatch.setenv("APPDATA", str(home))

        template_file = home / "sample_template.docx"
        doc = Document()
        doc.add_paragraph("template")
        doc.save(template_file)

        imported = import_word_template(template_file)
        templates = list_word_templates()

        assert imported.exists()
        assert imported in templates


def test_word_format_applies_template_and_preserves_content(monkeypatch) -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        monkeypatch.setenv("HOME", str(root))
        monkeypatch.setenv("USERPROFILE", str(root))
        monkeypatch.setenv("LOCALAPPDATA", str(root))
        monkeypatch.setenv("APPDATA", str(root))

        template = root / "template.docx"
        template_doc = Document()
        template_doc.add_heading("Template Title", level=1)
        template_doc.save(template)

        source = root / "source.docx"
        image_path = root / "img.png"
        _create_sample_png(image_path)

        source_doc = Document()
        source_doc.add_heading("Source Heading", level=1)
        source_doc.add_paragraph("Source paragraph")
        table = source_doc.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "C1"
        table.cell(0, 1).text = "C2"
        source_doc.add_picture(str(image_path))
        source_doc.save(source)

        out_dir = root / "out"
        results = format_word_documents(
            sources=[source],
            destination=out_dir,
            workspace=root,
            dry_run=False,
            template_path=template,
        )

        assert len(results) == 1
        assert results[0].status == OperationStatus.SUCCESS

        output = out_dir / "source_formatted.docx"
        assert output.exists()

        out_doc = Document(str(output))
        text_blob = "\n".join(par.text for par in out_doc.paragraphs)
        assert "Source Heading" in text_blob
        assert "Source paragraph" in text_blob
        assert any(tbl.cell(0, 0).text == "C1" for tbl in out_doc.tables)
        assert len(out_doc.inline_shapes) >= 1


def test_word_format_dry_run(monkeypatch) -> None:
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        monkeypatch.setenv("HOME", str(root))
        monkeypatch.setenv("USERPROFILE", str(root))
        monkeypatch.setenv("LOCALAPPDATA", str(root))
        monkeypatch.setenv("APPDATA", str(root))

        template = root / "template.docx"
        doc = Document()
        doc.add_paragraph("template")
        doc.save(template)

        source = root / "source.docx"
        source_doc = Document()
        source_doc.add_paragraph("source")
        source_doc.save(source)

        out_dir = root / "out"
        results = format_word_documents(
            sources=[source],
            destination=out_dir,
            workspace=root,
            dry_run=True,
            template_path=template,
        )

        assert len(results) == 1
        assert results[0].status == OperationStatus.DRY_RUN
        assert not out_dir.exists()
