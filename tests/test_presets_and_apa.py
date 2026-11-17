"""Tests for context-adaptive document profiles, page sizes, TOC, and APA formatting."""

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Inches, Pt


def test_resolve_profile_all_valid_and_invalid():
    from word_mcp.core.presets import PRESETS, resolve_profile

    for name in ("modern", "formal", "report", "academic", "apa", "compact"):
        profile = resolve_profile(name)
        assert profile["name"] == name
        assert name in PRESETS

    assert resolve_profile(None)["name"] == "modern"

    with pytest.raises(ValueError, match="Unknown preset 'nonexistent'"):
        resolve_profile("nonexistent")


def test_resolve_page_size_valid_and_invalid():
    from word_mcp.core.presets import resolve_page_size

    w_a4, h_a4 = resolve_page_size("A4")
    assert round(w_a4.cm, 2) == 21.0
    assert round(h_a4.cm, 2) == 29.7

    w_let, h_let = resolve_page_size("letter")
    assert w_let == Inches(8.5)
    assert h_let == Inches(11.0)

    with pytest.raises(ValueError, match="Unknown page_size"):
        resolve_page_size("tabloid")


def test_apply_profile_typography_apa_and_report():
    from word_mcp.core.presets import apply_profile_typography, resolve_profile

    doc_apa = Document()
    apa = resolve_profile("apa")
    apply_profile_typography(doc_apa, apa)
    assert doc_apa.styles["Normal"].font.name == "Times New Roman"
    assert doc_apa.styles["Normal"].font.size == Pt(12)
    assert doc_apa.styles["Heading 1"].font.name == "Times New Roman"
    assert doc_apa.styles["Heading 1"].paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert doc_apa.styles["Heading 3"].font.italic is True

    doc_rep = Document()
    rep = resolve_profile("report")
    apply_profile_typography(doc_rep, rep)
    assert doc_rep.styles["Normal"].font.name == "Calibri"
    assert doc_rep.styles["Heading 1"].font.name == "Calibri"
    assert doc_rep.styles["Heading 1"].font.bold is True
    assert round(Cm(21.0).cm, 1) == 21.0


def test_core_add_table_of_contents_enables_update_fields_and_dirty():
    from docx.oxml.ns import qn

    from word_mcp.core.content import core_add_table_of_contents

    doc = Document()
    res = core_add_table_of_contents(
        doc, title="Índice", max_level=3, page_break=False
    )
    assert res["title"] == "Índice"

    update_fields = doc.settings.element.find(qn("w:updateFields"))
    assert update_fields is not None
    assert update_fields.get(qn("w:val")) == "true"

    xml_str = doc.element.xml
    assert 'w:dirty="true"' in xml_str


def test_core_set_page_layout_supports_page_size():
    from word_mcp.core.properties import core_set_page_layout

    doc = Document()
    core_set_page_layout(doc, section_idx=0, orientation="portrait", page_size="a4")
    sec = doc.sections[0]
    assert round(sec.page_width.cm, 1) == 21.0
    assert round(sec.page_height.cm, 1) == 29.7

    core_set_page_layout(doc, section_idx=0, orientation="landscape", page_size="a4")
    assert round(sec.page_width.cm, 1) == 29.7
    assert round(sec.page_height.cm, 1) == 21.0


def test_core_build_document_report_and_apa_dynamic_elements():
    from docx.oxml.ns import qn

    from word_mcp.core.batch import core_build_document

    doc_report = Document()
    res_rep = core_build_document(
        doc_report,
        title="Annual Report",
        preset="report",
        elements=[
            {"type": "title_block", "title": "Annual Report", "subtitle": "Q4", "author": "Team"},
            {"type": "toc", "title": "Índice", "max_level": 3, "page_break": False},
            {"type": "heading", "level": 1, "text": "Summary"},
            {
                "type": "bullet_list",
                "items": [
                    "Flat bullet",
                    {"text": "Parent bullet", "children": ["Nested child"]},
                ],
            },
            {
                "type": "references",
                "title": "References",
                "items": [
                    "Doe, J. (2025). *Sample Book*. Publisher.",
                    {
                        "authors": "Smith, A.",
                        "year": "2024",
                        "title": "Dynamic Docs",
                        "source": "Tech Press",
                        "doi": "https://doi.org/10.1000/xyz",
                    },
                ],
            },
        ],
    )
    assert res_rep["status"] == "success"
    assert round(doc_report.sections[0].page_width.cm, 1) == 21.0
    assert doc_report.settings.element.find(qn("w:updateFields")) is not None

    # Check nested bullet indentation
    nested_p = [p for p in doc_report.paragraphs if p.text == "Nested child"][0]
    assert nested_p.paragraph_format.left_indent is not None
    assert nested_p.paragraph_format.left_indent.cm > 0.5

    # Check APA hanging indent on references
    ref_p = [p for p in doc_report.paragraphs if "Doe, J." in p.text][0]
    assert ref_p.paragraph_format.left_indent.cm > 1.0
    assert ref_p.paragraph_format.first_line_indent.cm < -1.0


def test_core_build_document_rejects_invalid_preset_and_unknown_element():
    from word_mcp.core.batch import core_build_document

    with pytest.raises(ValueError, match="Unknown preset"):
        core_build_document(Document(), preset="invalid_preset")

    with pytest.raises(ValueError, match="Unknown element type"):
        core_build_document(Document(), elements=[{"type": "magic_widget"}])


def test_parse_markdown_preserves_nested_lists():
    from word_mcp.core.batch import parse_markdown_to_elements

    md = "- Top item\n  - Child item\n- Second top"
    elems = parse_markdown_to_elements(md)
    assert len(elems) == 1
    assert elems[0]["type"] == "bullet_list"
    items = elems[0]["items"]
    assert items[0] == "Top item"
    assert isinstance(items[1], dict) and items[1]["level"] == 1
    assert items[1]["text"] == "Child item"
