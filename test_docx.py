"""
Tests for the DOCX tools in writing_mcp.py.

All tests are self-contained: minimal DOCX files are created in-memory using
zipfile so no Word/LibreOffice installation is required. Zotero API calls in
docx_zotero_insert_citations are patched with a lightweight mock.
"""
import io
import json
import zipfile
from pathlib import Path
from unittest.mock import patch

import pytest

# ---------------------------------------------------------------------------
# Minimal DOCX builder
# ---------------------------------------------------------------------------
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = f"{{{_W_NS}}}"

_CONTENT_TYPES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
   Target="word/document.xml"/>
</Relationships>
"""

_DOC_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
   Target="styles.xml"/>
</Relationships>
"""


def _make_document_xml(*paragraphs: str) -> str:
    """Build a minimal word/document.xml with one run per paragraph."""
    paras = ""
    for text in paragraphs:
        escaped = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        paras += f'<w:p><w:r><w:t>{escaped}</w:t></w:r></w:p>\n'
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{_W_NS}">'
        f"<w:body>{paras}<w:sectPr/></w:body>"
        f"</w:document>"
    )


def _make_heading_document_xml() -> str:
    """Document with heading paragraphs that have pStyle elements."""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{_W_NS}"><w:body>'
        f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>'
        f'<w:p><w:r><w:t>Some body text.</w:t></w:r></w:p>'
        f'<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Background</w:t></w:r></w:p>'
        f'<w:p><w:r><w:t>More body text.</w:t></w:r></w:p>'
        f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Methods</w:t></w:r></w:p>'
        f"<w:sectPr/></w:body></w:document>"
    )


def _build_docx(tmp_path: Path, name: str = "test.docx", *paragraphs: str) -> Path:
    """Write a minimal DOCX to tmp_path and return its path."""
    if not paragraphs:
        paragraphs = ("Hello world.", "This paper cites Smith [12] and Jones [34].")
    doc_xml = _make_document_xml(*paragraphs)
    out = tmp_path / name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _RELS)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)
    out.write_bytes(buf.getvalue())
    return out


def _build_heading_docx(tmp_path: Path, name: str = "headings.docx") -> Path:
    doc_xml = _make_heading_document_xml()
    out = tmp_path / name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _RELS)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)
    out.write_bytes(buf.getvalue())
    return out


def _read_output_text(out_path: str) -> str:
    """Extract visible text from a DOCX output path."""
    from xml.etree import ElementTree as ET
    with zipfile.ZipFile(out_path, "r") as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
    root = ET.fromstring(xml.encode("utf-8"))
    texts = []
    for t in root.iter(f"{{{_W_NS}}}t"):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


# Import module under test (after helpers defined so namespace is registered)
import writing_mcp as mcp_mod


# ---------------------------------------------------------------------------
# Helper / internal unit tests
# ---------------------------------------------------------------------------

class TestInternalHelpers:
    def test_tc_date_format(self):
        date = mcp_mod._tc_date()
        import re
        assert re.match(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z", date)

    def test_max_wid_empty(self):
        assert mcp_mod._max_wid("<w:document/>") == 0

    def test_max_wid_finds_max(self):
        xml = '<doc w:id="3" w:id="7" w:id="2"/>'
        assert mcp_mod._max_wid(xml) == 7

    def test_docx_visible_paragraphs_basic(self):
        xml = _make_document_xml("Alpha", "Beta", "")
        paras = mcp_mod._docx_visible_paragraphs(xml)
        texts = [p["text"] for p in paras]
        assert "Alpha" in texts
        assert "Beta" in texts
        assert "" not in texts  # empty paragraph filtered out

    def test_docx_live_zotero_summary_empty(self):
        xml = _make_document_xml("No citations here.")
        summary = mcp_mod._docx_live_zotero_summary(xml)
        assert summary["csl_citation_marker_count"] == 0
        assert summary["unique_zotero_item_keys_in_live_fields"] == 0

    def test_short_author_year_empty(self):
        result = mcp_mod._short_author_year([])
        assert result == "[Zotero citation]"

    def test_short_author_year_single_author(self):
        csl = [{"author": [{"family": "Smith", "given": "J"}],
                "issued": {"date-parts": [[2021]]}}]
        result = mcp_mod._short_author_year(csl)
        assert "Smith" in result
        assert "2021" in result

    def test_short_author_year_many_authors(self):
        csl = [{"author": [
            {"family": "Smith"}, {"family": "Jones"}, {"family": "Brown"}
        ], "issued": {"date-parts": [[2020]]}}]
        result = mcp_mod._short_author_year(csl)
        assert "et al." in result


# ---------------------------------------------------------------------------
# docx_zotero_audit
# ---------------------------------------------------------------------------

class TestDocxZoteroAudit:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_zotero_audit(str(tmp_path / "missing.docx")))
        assert result["status"] == "error"
        assert "not found" in result["error"]

    def test_basic_audit(self, tmp_path):
        docx = _build_docx(tmp_path)
        result = json.loads(mcp_mod.docx_zotero_audit(str(docx)))
        assert result["status"] == "ok"
        assert result["csl_citation_marker_count"] == 0
        assert isinstance(result["visible_numeric_reference_labels"], list)
        assert 12 in result["visible_numeric_reference_labels"]
        assert 34 in result["visible_numeric_reference_labels"]

    def test_suspect_terms_found(self, tmp_path):
        docx = _build_docx(tmp_path, "test.docx",
                            "Smith et al. found something important.")
        result = json.loads(mcp_mod.docx_zotero_audit(str(docx), suspect_terms=["Smith"]))
        assert result["status"] == "ok"
        assert len(result["suspect_hits"]) >= 1
        assert result["suspect_hits"][0]["term"] == "Smith"

    def test_suspect_terms_not_found(self, tmp_path):
        docx = _build_docx(tmp_path)
        result = json.loads(mcp_mod.docx_zotero_audit(str(docx), suspect_terms=["Zhan"]))
        assert result["status"] == "ok"
        assert result["suspect_hits"] == []

    def test_no_suspect_terms(self, tmp_path):
        docx = _build_docx(tmp_path)
        result = json.loads(mcp_mod.docx_zotero_audit(str(docx)))
        assert result["status"] == "ok"
        assert result["suspect_hits"] == []


# ---------------------------------------------------------------------------
# docx_text_replace
# ---------------------------------------------------------------------------

class TestDocxTextReplace:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_text_replace(
            str(tmp_path / "missing.docx"),
            [{"find": "x", "replace": "y"}],
        ))
        assert result["status"] == "error"

    def test_simple_replace(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "The quick brown fox.")
        result = json.loads(mcp_mod.docx_text_replace(
            str(docx),
            [{"find": "quick brown", "replace": "slow red"}],
            output_path=str(tmp_path / "out.docx"),
            overwrite=True,
        ))
        assert result["status"] == "ok"
        assert result["results"][0]["occurrences_edited"] == 1
        text = _read_output_text(result["output_path"])
        assert "slow red fox" in text
        assert "quick brown" not in text

    def test_replace_not_found(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Hello world.")
        result = json.loads(mcp_mod.docx_text_replace(
            str(docx),
            [{"find": "NOTHERE", "replace": "x"}],
            output_path=str(tmp_path / "out.docx"),
            overwrite=True,
        ))
        assert result["results"][0]["status"] == "not_found"

    def test_overwrite_protection(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Hello world.")
        out = tmp_path / "out.docx"
        out.write_bytes(b"existing")
        result = json.loads(mcp_mod.docx_text_replace(
            str(docx),
            [{"find": "Hello", "replace": "Hi"}],
            output_path=str(out),
        ))
        assert result["status"] == "error"
        assert "set_overwrite" in result

    def test_output_is_valid_zip(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "The quick brown fox.")
        out = tmp_path / "out.docx"
        mcp_mod.docx_text_replace(
            str(docx),
            [{"find": "quick", "replace": "slow"}],
            output_path=str(out),
            overwrite=True,
        )
        assert zipfile.is_zipfile(str(out))

    def test_track_changes(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Old text here.")
        out = tmp_path / "tracked.docx"
        result = json.loads(mcp_mod.docx_text_replace(
            str(docx),
            [{"find": "Old", "replace": "New"}],
            output_path=str(out),
            overwrite=True,
            track_changes=True,
            author="TestAuthor",
        ))
        assert result["status"] == "ok"
        assert result["results"][0]["tracked"] is True
        # Verify the DOCX contains track-change markup
        with zipfile.ZipFile(str(out)) as zf:
            doc_xml = zf.read("word/document.xml").decode()
        assert "w:del" in doc_xml
        assert "w:ins" in doc_xml
        assert "TestAuthor" in doc_xml

    def test_empty_find_skipped(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Hello world.")
        result = json.loads(mcp_mod.docx_text_replace(
            str(docx),
            [{"find": "", "replace": "x"}],
            output_path=str(tmp_path / "out.docx"),
            overwrite=True,
        ))
        assert result["results"][0]["status"] == "skipped"

    def test_default_output_name(self, tmp_path):
        docx = _build_docx(tmp_path, "paper.docx", "Hello world.")
        result = json.loads(mcp_mod.docx_text_replace(
            str(docx),
            [{"find": "Hello", "replace": "Hi"}],
        ))
        assert result["status"] == "ok"
        assert "paper_edited.docx" in result["output_path"]


# ---------------------------------------------------------------------------
# docx_add_comment
# ---------------------------------------------------------------------------

class TestDocxAddComment:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_add_comment(
            str(tmp_path / "missing.docx"),
            [{"find": "x", "comment": "note"}],
        ))
        assert result["status"] == "error"

    def test_add_single_comment(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "This is the anchor text.")
        out = tmp_path / "commented.docx"
        result = json.loads(mcp_mod.docx_add_comment(
            str(docx),
            [{"find": "anchor text", "comment": "Review this."}],
            output_path=str(out),
            overwrite=True,
        ))
        assert result["status"] == "ok"
        assert result["comments_added"] == 1
        assert result["results"][0]["status"] == "ok"

    def test_comment_text_not_found(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Hello world.")
        out = tmp_path / "commented.docx"
        result = json.loads(mcp_mod.docx_add_comment(
            str(docx),
            [{"find": "NOTHERE", "comment": "note"}],
            output_path=str(out),
            overwrite=True,
        ))
        assert result["comments_added"] == 0
        assert result["results"][0]["status"] == "not_found"

    def test_output_has_comments_xml(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Anchor here please.")
        out = tmp_path / "commented.docx"
        mcp_mod.docx_add_comment(
            str(docx),
            [{"find": "Anchor", "comment": "Check this."}],
            output_path=str(out),
            overwrite=True,
        )
        with zipfile.ZipFile(str(out)) as zf:
            assert "word/comments.xml" in zf.namelist()
            cxml = zf.read("word/comments.xml").decode()
        assert "Check this." in cxml

    def test_output_is_valid_zip(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Text here.")
        out = tmp_path / "out.docx"
        mcp_mod.docx_add_comment(
            str(docx),
            [{"find": "Text", "comment": "A note."}],
            output_path=str(out),
            overwrite=True,
        )
        assert zipfile.is_zipfile(str(out))

    def test_overwrite_protection(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Hello.")
        out = tmp_path / "out.docx"
        out.write_bytes(b"existing")
        result = json.loads(mcp_mod.docx_add_comment(
            str(docx),
            [{"find": "Hello", "comment": "note"}],
            output_path=str(out),
        ))
        assert result["status"] == "error"

    def test_custom_author_and_initials(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Annotate this word.")
        out = tmp_path / "commented.docx"
        mcp_mod.docx_add_comment(
            str(docx),
            [{"find": "Annotate", "comment": "Good point."}],
            output_path=str(out),
            overwrite=True,
            author="Dr. Tester",
            initials="DT",
        )
        with zipfile.ZipFile(str(out)) as zf:
            cxml = zf.read("word/comments.xml").decode()
        assert "Dr. Tester" in cxml
        assert "DT" in cxml

    def test_skips_empty_find_or_comment(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Hello world.")
        out = tmp_path / "out.docx"
        result = json.loads(mcp_mod.docx_add_comment(
            str(docx),
            [{"find": "", "comment": "note"}, {"find": "Hello", "comment": ""}],
            output_path=str(out),
            overwrite=True,
        ))
        for r in result["results"]:
            assert r["status"] == "skipped"


# ---------------------------------------------------------------------------
# docx_zotero_insert_citations  (Zotero API mocked)
# ---------------------------------------------------------------------------

_MOCK_ZOTERO_ITEM = {
    "key": "ABC12345",
    "data": {
        "key": "ABC12345",
        "itemType": "journalArticle",
        "title": "A Great Paper",
        "creators": [{"creatorType": "author", "lastName": "Smith", "firstName": "J."}],
        "date": "2020",
    },
    "library": {"id": 999},
}

_MOCK_CSL = {
    "id": "ABC12345",
    "type": "article-journal",
    "title": "A Great Paper",
    "author": [{"family": "Smith", "given": "J."}],
    "issued": {"date-parts": [[2020]]},
}


def _mock_get_csl(key: str):
    return _MOCK_ZOTERO_ITEM, _MOCK_CSL


class TestDocxZoteroInsertCitations:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_zotero_insert_citations(
            str(tmp_path / "missing.docx"),
            [{"old_text": "Smith [12]", "zotero_keys": ["ABC12345"]}],
        ))
        assert result["status"] == "error"

    def test_simple_citation_insertion(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx",
                            "Introduction.", "Smith [12] found something important.")
        out = tmp_path / "out.docx"
        with patch.object(mcp_mod, "_get_csl_item_data", side_effect=_mock_get_csl):
            result = json.loads(mcp_mod.docx_zotero_insert_citations(
                str(docx),
                [{"old_text": "Smith [12]", "zotero_keys": ["ABC12345"]}],
                output_path=str(out),
                overwrite=True,
            ))
        assert result["status"] == "ok"
        assert result["results"][0]["replacements_made"] >= 1

    def test_output_contains_zotero_field(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Jones [34] studied this.")
        out = tmp_path / "out.docx"
        with patch.object(mcp_mod, "_get_csl_item_data", side_effect=_mock_get_csl):
            mcp_mod.docx_zotero_insert_citations(
                str(docx),
                [{"old_text": "Jones [34]", "zotero_keys": ["ABC12345"]}],
                output_path=str(out),
                overwrite=True,
            )
        with zipfile.ZipFile(str(out)) as zf:
            doc_xml = zf.read("word/document.xml").decode()
        assert "CSL_CITATION" in doc_xml
        assert "fldCharType" in doc_xml

    def test_old_text_not_found(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Hello world.")
        out = tmp_path / "out.docx"
        with patch.object(mcp_mod, "_get_csl_item_data", side_effect=_mock_get_csl):
            result = json.loads(mcp_mod.docx_zotero_insert_citations(
                str(docx),
                [{"old_text": "NOTHERE [99]", "zotero_keys": ["ABC12345"]}],
                output_path=str(out),
                overwrite=True,
            ))
        assert result["results"][0]["status"] == "not_found_or_spans_runs"

    def test_output_is_valid_zip(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Jones [34] studied this.")
        out = tmp_path / "out.docx"
        with patch.object(mcp_mod, "_get_csl_item_data", side_effect=_mock_get_csl):
            mcp_mod.docx_zotero_insert_citations(
                str(docx),
                [{"old_text": "Jones [34]", "zotero_keys": ["ABC12345"]}],
                output_path=str(out),
                overwrite=True,
            )
        assert zipfile.is_zipfile(str(out))

    def test_missing_old_text_or_keys_skipped(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Hello world.")
        out = tmp_path / "out.docx"
        with patch.object(mcp_mod, "_get_csl_item_data", side_effect=_mock_get_csl):
            result = json.loads(mcp_mod.docx_zotero_insert_citations(
                str(docx),
                [{"old_text": "", "zotero_keys": ["ABC12345"]}],
                output_path=str(out),
                overwrite=True,
            ))
        assert result["results"][0]["status"] == "skipped"

    def test_overwrite_protection(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Hello world.")
        out = tmp_path / "out.docx"
        out.write_bytes(b"existing")
        with patch.object(mcp_mod, "_get_csl_item_data", side_effect=_mock_get_csl):
            result = json.loads(mcp_mod.docx_zotero_insert_citations(
                str(docx),
                [{"old_text": "Hello", "zotero_keys": ["ABC12345"]}],
                output_path=str(out),
            ))
        assert result["status"] == "error"


# ---------------------------------------------------------------------------
# docx_extract_text (new tool)
# ---------------------------------------------------------------------------

class TestDocxExtractText:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_extract_text(str(tmp_path / "missing.docx")))
        assert result["status"] == "error"

    def test_extracts_all_paragraphs(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx",
                            "First paragraph.", "Second paragraph.", "Third.")
        result = json.loads(mcp_mod.docx_extract_text(str(docx)))
        assert result["status"] == "ok"
        assert result["paragraph_count"] == 3
        texts = [p["text"] for p in result["paragraphs"]]
        assert "First paragraph." in texts
        assert "Second paragraph." in texts
        assert "Third." in texts

    def test_full_text_field(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Alpha.", "Beta.")
        result = json.loads(mcp_mod.docx_extract_text(str(docx)))
        assert "Alpha." in result["full_text"]
        assert "Beta." in result["full_text"]


# ---------------------------------------------------------------------------
# docx_get_headings (new tool)
# ---------------------------------------------------------------------------

class TestDocxGetHeadings:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_get_headings(str(tmp_path / "missing.docx")))
        assert result["status"] == "error"

    def test_extracts_headings(self, tmp_path):
        docx = _build_heading_docx(tmp_path)
        result = json.loads(mcp_mod.docx_get_headings(str(docx)))
        assert result["status"] == "ok"
        headings = result["headings"]
        assert len(headings) == 3
        texts = [h["text"] for h in headings]
        assert "Introduction" in texts
        assert "Background" in texts
        assert "Methods" in texts

    def test_heading_levels(self, tmp_path):
        docx = _build_heading_docx(tmp_path)
        result = json.loads(mcp_mod.docx_get_headings(str(docx)))
        levels = {h["text"]: h["level"] for h in result["headings"]}
        assert levels["Introduction"] == 1
        assert levels["Background"] == 2
        assert levels["Methods"] == 1

    def test_no_headings(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Just body text here.")
        result = json.loads(mcp_mod.docx_get_headings(str(docx)))
        assert result["status"] == "ok"
        assert result["headings"] == []


# ---------------------------------------------------------------------------
# DOCX builder helpers for new tools
# ---------------------------------------------------------------------------

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

_CONTENT_TYPES_WITH_IMAGE = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml"
   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml"
   ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"
   ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""

_DOC_RELS_WITH_IMAGE = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{_PKG_NS}">
  <Relationship Id="rId1"
   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
   Target="styles.xml"/>
  <Relationship Id="rId5"
   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
   Target="media/image1.png"/>
</Relationships>
"""

_CORE_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Test Paper</dc:title>
  <dc:creator>Jane Smith</dc:creator>
  <cp:lastModifiedBy>Jane Smith</cp:lastModifiedBy>
  <cp:revision>3</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-06-01T12:00:00Z</dcterms:modified>
</cp:coreProperties>
"""

_APP_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Word</Application>
  <Pages>4</Pages>
  <Words>1234</Words>
  <Characters>7000</Characters>
  <Paragraphs>50</Paragraphs>
</Properties>
"""

_TINY_PNG = (
    b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01'
    b'\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00'
    b'\x00\x01\x01\x00\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82'
)


def _make_doc_xml_with_image() -> str:
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{_W_NS}"'
        f' xmlns:r="{_R_NS}"'
        f' xmlns:wp="{_WP_NS}"'
        f' xmlns:a="{_A_NS}">'
        f'<w:body>'
        f'<w:p><w:r><w:t>See Figure 1.</w:t></w:r></w:p>'
        f'<w:p><w:r><w:drawing>'
        f'<wp:inline>'
        f'<wp:docPr id="1" name="Figure1" descr="My figure alt text"/>'
        f'<a:graphic><a:graphicData>'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill>'
        f'</pic:pic>'
        f'</a:graphicData></a:graphic>'
        f'</wp:inline>'
        f'</w:drawing></w:r></w:p>'
        f'<w:sectPr/></w:body></w:document>'
    )


def _make_doc_xml_with_table() -> str:
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{_W_NS}"><w:body>'
        f'<w:p><w:r><w:t>Before table.</w:t></w:r></w:p>'
        f'<w:tbl>'
        f'<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>'
        f'<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>'
        f'<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>'
        f'<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>'
        f'</w:tbl>'
        f'<w:tbl>'
        f'<w:tr><w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr>'
        f'</w:tbl>'
        f'<w:sectPr/></w:body></w:document>'
    )


def _make_doc_xml_with_tracked_changes() -> str:
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{_W_NS}"><w:body>'
        f'<w:p>'
        f'<w:del w:id="1" w:author="Alice" w:date="2024-01-01T00:00:00Z">'
        f'<w:r><w:delText>old word</w:delText></w:r>'
        f'</w:del>'
        f'<w:ins w:id="2" w:author="Alice" w:date="2024-01-01T00:00:00Z">'
        f'<w:r><w:t>new word</w:t></w:r>'
        f'</w:ins>'
        f'<w:r><w:t> remains.</w:t></w:r>'
        f'</w:p>'
        f'<w:sectPr/></w:body></w:document>'
    )


def _build_docx_with_image(tmp_path: Path, name: str = "img.docx") -> Path:
    out = tmp_path / name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES_WITH_IMAGE)
        zf.writestr("_rels/.rels", _RELS)
        zf.writestr("word/document.xml", _make_doc_xml_with_image())
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS_WITH_IMAGE)
        zf.writestr("word/media/image1.png", _TINY_PNG)
        zf.writestr("docProps/core.xml", _CORE_XML)
        zf.writestr("docProps/app.xml", _APP_XML)
    out.write_bytes(buf.getvalue())
    return out


def _build_docx_with_table(tmp_path: Path, name: str = "table.docx") -> Path:
    out = tmp_path / name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _RELS)
        zf.writestr("word/document.xml", _make_doc_xml_with_table())
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)
    out.write_bytes(buf.getvalue())
    return out


def _build_docx_with_tracked_changes(tmp_path: Path, name: str = "tracked.docx") -> Path:
    out = tmp_path / name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _RELS)
        zf.writestr("word/document.xml", _make_doc_xml_with_tracked_changes())
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)
    out.write_bytes(buf.getvalue())
    return out


def _build_docx_with_properties(tmp_path: Path, name: str = "props.docx") -> Path:
    out = tmp_path / name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES_WITH_IMAGE)
        zf.writestr("_rels/.rels", _RELS)
        zf.writestr("word/document.xml", _make_document_xml("Body text."))
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)
        zf.writestr("docProps/core.xml", _CORE_XML)
        zf.writestr("docProps/app.xml", _APP_XML)
    out.write_bytes(buf.getvalue())
    return out


# ---------------------------------------------------------------------------
# docx_extract_image
# ---------------------------------------------------------------------------

class TestDocxExtractImage:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_extract_image(str(tmp_path / "missing.docx"), "rId5"))
        assert result["status"] == "error"

    def test_invalid_rid(self, tmp_path):
        docx = _build_docx_with_image(tmp_path)
        result = json.loads(mcp_mod.docx_extract_image(str(docx), "rId99"))
        assert result["status"] == "error"
        assert "not found" in result["error"]

    def test_returns_image_object(self, tmp_path):
        from fastmcp.utilities.types import Image as MCPImage
        docx = _build_docx_with_image(tmp_path)
        result = mcp_mod.docx_extract_image(str(docx), "rId5")
        assert isinstance(result, MCPImage)

    def test_image_data_matches_original(self, tmp_path):
        docx = _build_docx_with_image(tmp_path)
        result = mcp_mod.docx_extract_image(str(docx), "rId5")
        # MCPImage._get_data() returns base64; decode and compare bytes
        import base64
        raw = base64.b64decode(result._get_data())
        assert raw == _TINY_PNG

    def test_image_format_inferred(self, tmp_path):
        docx = _build_docx_with_image(tmp_path)
        result = mcp_mod.docx_extract_image(str(docx), "rId5")
        assert "png" in result._mime_type

    def test_emf_returns_not_renderable(self, tmp_path):
        # Build a DOCX with an EMF relationship
        emf_rels = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{_PKG_NS}">
  <Relationship Id="rId1"
   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
   Target="styles.xml"/>
  <Relationship Id="rId6"
   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
   Target="media/figure.emf"/>
</Relationships>
"""
        out = tmp_path / "emf.docx"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
            zf.writestr("_rels/.rels", _RELS)
            zf.writestr("word/document.xml", _make_document_xml("EMF figure doc."))
            zf.writestr("word/_rels/document.xml.rels", emf_rels)
            zf.writestr("word/media/figure.emf", b"\x01\x00\x00\x00")
        out.write_bytes(buf.getvalue())
        result = json.loads(mcp_mod.docx_extract_image(str(out), "rId6"))
        assert result["status"] == "not_renderable"


# ---------------------------------------------------------------------------
# docx_list_images
# ---------------------------------------------------------------------------

class TestDocxListImages:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_list_images(str(tmp_path / "missing.docx")))
        assert result["status"] == "error"

    def test_lists_one_image(self, tmp_path):
        docx = _build_docx_with_image(tmp_path)
        result = json.loads(mcp_mod.docx_list_images(str(docx)))
        assert result["status"] == "ok"
        assert result["image_count"] == 1
        img = result["images"][0]
        assert img["rId"] == "rId5"
        assert img["filename"] == "image1.png"
        assert img["exists_in_archive"] is True

    def test_alt_text_extracted(self, tmp_path):
        docx = _build_docx_with_image(tmp_path)
        result = json.loads(mcp_mod.docx_list_images(str(docx)))
        assert result["images"][0]["alt_text"] == "My figure alt text"

    def test_no_images(self, tmp_path):
        docx = _build_docx(tmp_path)
        result = json.loads(mcp_mod.docx_list_images(str(docx)))
        assert result["status"] == "ok"
        assert result["image_count"] == 0


# ---------------------------------------------------------------------------
# docx_replace_image
# ---------------------------------------------------------------------------

class TestDocxReplaceImage:
    def test_file_not_found(self, tmp_path):
        img = tmp_path / "new.png"
        img.write_bytes(_TINY_PNG)
        result = json.loads(mcp_mod.docx_replace_image(
            str(tmp_path / "missing.docx"), "rId5", str(img),
        ))
        assert result["status"] == "error"

    def test_new_image_not_found(self, tmp_path):
        docx = _build_docx_with_image(tmp_path)
        result = json.loads(mcp_mod.docx_replace_image(
            str(docx), "rId5", str(tmp_path / "ghost.png"),
        ))
        assert result["status"] == "error"
        assert "new image not found" in result["error"]

    def test_invalid_rid(self, tmp_path):
        docx = _build_docx_with_image(tmp_path)
        img = tmp_path / "new.png"
        img.write_bytes(_TINY_PNG)
        result = json.loads(mcp_mod.docx_replace_image(
            str(docx), "rId99", str(img),
            output_path=str(tmp_path / "out.docx"), overwrite=True,
        ))
        assert result["status"] == "error"
        assert "not found" in result["error"]

    def test_successful_replace(self, tmp_path):
        docx = _build_docx_with_image(tmp_path)
        new_img = tmp_path / "new.png"
        new_img.write_bytes(_TINY_PNG + b"\x00\x00")  # slightly different bytes
        out = tmp_path / "out.docx"
        result = json.loads(mcp_mod.docx_replace_image(
            str(docx), "rId5", str(new_img),
            output_path=str(out), overwrite=True,
        ))
        assert result["status"] == "ok"
        assert result["rId"] == "rId5"
        with zipfile.ZipFile(str(out)) as zf:
            stored = zf.read("word/media/image1.png")
        assert stored == _TINY_PNG + b"\x00\x00"

    def test_overwrite_protection(self, tmp_path):
        docx = _build_docx_with_image(tmp_path)
        img = tmp_path / "new.png"
        img.write_bytes(_TINY_PNG)
        out = tmp_path / "out.docx"
        out.write_bytes(b"existing")
        result = json.loads(mcp_mod.docx_replace_image(
            str(docx), "rId5", str(img), output_path=str(out),
        ))
        assert result["status"] == "error"
        assert "set_overwrite" in result

    def test_output_is_valid_zip(self, tmp_path):
        docx = _build_docx_with_image(tmp_path)
        img = tmp_path / "new.png"
        img.write_bytes(_TINY_PNG)
        out = tmp_path / "out.docx"
        mcp_mod.docx_replace_image(
            str(docx), "rId5", str(img), output_path=str(out), overwrite=True,
        )
        assert zipfile.is_zipfile(str(out))


# ---------------------------------------------------------------------------
# docx_get_tables
# ---------------------------------------------------------------------------

class TestDocxGetTables:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_get_tables(str(tmp_path / "missing.docx")))
        assert result["status"] == "error"

    def test_finds_two_tables(self, tmp_path):
        docx = _build_docx_with_table(tmp_path)
        result = json.loads(mcp_mod.docx_get_tables(str(docx)))
        assert result["status"] == "ok"
        assert result["table_count"] == 2

    def test_table_content(self, tmp_path):
        docx = _build_docx_with_table(tmp_path)
        result = json.loads(mcp_mod.docx_get_tables(str(docx)))
        t1 = result["tables"][0]
        assert t1["row_count"] == 2
        assert t1["col_count"] == 2
        assert t1["rows"][0] == ["A1", "B1"]
        assert t1["rows"][1] == ["A2", "B2"]

    def test_second_table(self, tmp_path):
        docx = _build_docx_with_table(tmp_path)
        result = json.loads(mcp_mod.docx_get_tables(str(docx)))
        t2 = result["tables"][1]
        assert t2["rows"][0] == ["X"]

    def test_no_tables(self, tmp_path):
        docx = _build_docx(tmp_path)
        result = json.loads(mcp_mod.docx_get_tables(str(docx)))
        assert result["status"] == "ok"
        assert result["table_count"] == 0


# ---------------------------------------------------------------------------
# docx_get_comments
# ---------------------------------------------------------------------------

class TestDocxGetComments:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_get_comments(str(tmp_path / "missing.docx")))
        assert result["status"] == "error"

    def test_no_comments_file(self, tmp_path):
        docx = _build_docx(tmp_path)
        result = json.loads(mcp_mod.docx_get_comments(str(docx)))
        assert result["status"] == "ok"
        assert result["comment_count"] == 0
        assert result["comments"] == []

    def test_reads_existing_comments(self, tmp_path):
        # Build a DOCX that already has a comments.xml via docx_add_comment
        docx = _build_docx(tmp_path, "input.docx", "Annotate this sentence here.")
        commented = tmp_path / "commented.docx"
        mcp_mod.docx_add_comment(
            str(docx),
            [{"find": "Annotate", "comment": "Check this reference."}],
            output_path=str(commented),
            overwrite=True,
            author="Reviewer1",
        )
        result = json.loads(mcp_mod.docx_get_comments(str(commented)))
        assert result["status"] == "ok"
        assert result["comment_count"] == 1
        c = result["comments"][0]
        assert c["author"] == "Reviewer1"
        assert "Check this reference." in c["text"]

    def test_multiple_comments(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx",
                            "First anchor text here.", "Second anchor text here.")
        commented = tmp_path / "commented.docx"
        mcp_mod.docx_add_comment(
            str(docx),
            [
                {"find": "First anchor", "comment": "Note one."},
                {"find": "Second anchor", "comment": "Note two."},
            ],
            output_path=str(commented),
            overwrite=True,
        )
        result = json.loads(mcp_mod.docx_get_comments(str(commented)))
        assert result["comment_count"] == 2
        texts = [c["text"] for c in result["comments"]]
        assert "Note one." in texts
        assert "Note two." in texts


# ---------------------------------------------------------------------------
# docx_get_tracked_changes
# ---------------------------------------------------------------------------

class TestDocxGetTrackedChanges:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_get_tracked_changes(str(tmp_path / "missing.docx")))
        assert result["status"] == "error"

    def test_no_tracked_changes(self, tmp_path):
        docx = _build_docx(tmp_path)
        result = json.loads(mcp_mod.docx_get_tracked_changes(str(docx)))
        assert result["status"] == "ok"
        assert result["change_count"] == 0

    def test_finds_deletion_and_insertion(self, tmp_path):
        docx = _build_docx_with_tracked_changes(tmp_path)
        result = json.loads(mcp_mod.docx_get_tracked_changes(str(docx)))
        assert result["status"] == "ok"
        assert result["change_count"] == 2
        types = {c["type"] for c in result["changes"]}
        assert "deletion" in types
        assert "insertion" in types

    def test_deletion_text(self, tmp_path):
        docx = _build_docx_with_tracked_changes(tmp_path)
        result = json.loads(mcp_mod.docx_get_tracked_changes(str(docx)))
        deletions = [c for c in result["changes"] if c["type"] == "deletion"]
        assert deletions[0]["text"] == "old word"
        assert deletions[0]["author"] == "Alice"

    def test_insertion_text(self, tmp_path):
        docx = _build_docx_with_tracked_changes(tmp_path)
        result = json.loads(mcp_mod.docx_get_tracked_changes(str(docx)))
        insertions = [c for c in result["changes"] if c["type"] == "insertion"]
        assert insertions[0]["text"] == "new word"

    def test_track_changes_tool_produces_detectable_changes(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "This is original text.")
        edited = tmp_path / "edited.docx"
        mcp_mod.docx_text_replace(
            str(docx),
            [{"find": "original", "replace": "revised"}],
            output_path=str(edited),
            overwrite=True,
            track_changes=True,
        )
        result = json.loads(mcp_mod.docx_get_tracked_changes(str(edited)))
        assert result["change_count"] == 2
        types = {c["type"] for c in result["changes"]}
        assert "deletion" in types
        assert "insertion" in types


# ---------------------------------------------------------------------------
# docx_accept_tracked_changes
# ---------------------------------------------------------------------------

class TestDocxAcceptTrackedChanges:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_accept_tracked_changes(str(tmp_path / "missing.docx")))
        assert result["status"] == "error"

    def test_accept_on_clean_doc(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Clean document.")
        out = tmp_path / "accepted.docx"
        result = json.loads(mcp_mod.docx_accept_tracked_changes(
            str(docx), output_path=str(out), overwrite=True,
        ))
        assert result["status"] == "ok"
        assert result["deletions_removed"] == 0
        assert result["insertions_accepted"] == 0

    def test_accepts_changes(self, tmp_path):
        docx = _build_docx_with_tracked_changes(tmp_path)
        out = tmp_path / "accepted.docx"
        result = json.loads(mcp_mod.docx_accept_tracked_changes(
            str(docx), output_path=str(out), overwrite=True,
        ))
        assert result["status"] == "ok"
        assert result["deletions_removed"] == 1
        assert result["insertions_accepted"] == 1

    def test_output_text_contains_insertion_not_deletion(self, tmp_path):
        docx = _build_docx_with_tracked_changes(tmp_path)
        out = tmp_path / "accepted.docx"
        mcp_mod.docx_accept_tracked_changes(
            str(docx), output_path=str(out), overwrite=True,
        )
        text = _read_output_text(str(out))
        assert "new word" in text
        assert "old word" not in text

    def test_no_del_ins_in_output_xml(self, tmp_path):
        docx = _build_docx_with_tracked_changes(tmp_path)
        out = tmp_path / "accepted.docx"
        mcp_mod.docx_accept_tracked_changes(
            str(docx), output_path=str(out), overwrite=True,
        )
        with zipfile.ZipFile(str(out)) as zf:
            doc_xml = zf.read("word/document.xml").decode()
        assert "w:del" not in doc_xml
        assert "w:ins" not in doc_xml

    def test_overwrite_protection(self, tmp_path):
        docx = _build_docx_with_tracked_changes(tmp_path)
        out = tmp_path / "out.docx"
        out.write_bytes(b"existing")
        result = json.loads(mcp_mod.docx_accept_tracked_changes(str(docx), output_path=str(out)))
        assert result["status"] == "error"
        assert "set_overwrite" in result

    def test_output_is_valid_zip(self, tmp_path):
        docx = _build_docx_with_tracked_changes(tmp_path)
        out = tmp_path / "accepted.docx"
        mcp_mod.docx_accept_tracked_changes(str(docx), output_path=str(out), overwrite=True)
        assert zipfile.is_zipfile(str(out))

    def test_roundtrip_via_text_replace_track_then_accept(self, tmp_path):
        docx = _build_docx(tmp_path, "input.docx", "Manuscript text goes here.")
        tracked = tmp_path / "tracked.docx"
        mcp_mod.docx_text_replace(
            str(docx),
            [{"find": "Manuscript", "replace": "Paper"}],
            output_path=str(tracked),
            overwrite=True,
            track_changes=True,
        )
        accepted = tmp_path / "accepted.docx"
        mcp_mod.docx_accept_tracked_changes(
            str(tracked), output_path=str(accepted), overwrite=True,
        )
        text = _read_output_text(str(accepted))
        assert "Paper" in text
        assert "Manuscript" not in text


# ---------------------------------------------------------------------------
# docx_get_properties
# ---------------------------------------------------------------------------

class TestDocxGetProperties:
    def test_file_not_found(self, tmp_path):
        result = json.loads(mcp_mod.docx_get_properties(str(tmp_path / "missing.docx")))
        assert result["status"] == "error"

    def test_reads_core_properties(self, tmp_path):
        docx = _build_docx_with_properties(tmp_path)
        result = json.loads(mcp_mod.docx_get_properties(str(docx)))
        assert result["status"] == "ok"
        assert result["title"] == "Test Paper"
        assert result["author"] == "Jane Smith"
        assert result["last_modified_by"] == "Jane Smith"
        assert result["revision"] == "3"

    def test_reads_app_properties(self, tmp_path):
        docx = _build_docx_with_properties(tmp_path)
        result = json.loads(mcp_mod.docx_get_properties(str(docx)))
        assert result["word_count"] == 1234
        assert result["page_count"] == 4
        assert result["paragraph_count"] == 50

    def test_no_props_files(self, tmp_path):
        docx = _build_docx(tmp_path)
        result = json.loads(mcp_mod.docx_get_properties(str(docx)))
        assert result["status"] == "ok"
        assert "title" not in result
