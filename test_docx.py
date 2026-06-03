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
