#!/usr/bin/env python3
"""
Academic Writing MCP Server

Unified MCP for dissertation/academic writing:
  - Zotero: search library, get scannable cites for ODF Scan workflow
  - LibreOffice: edit open ODT/DOCX documents via the LibreOffice MCP extension
  - DocConvert: convert between document formats via pandoc

Architecture (this machine):
  Claude Code (WSL2)
    ├── Zotero tools  → Zotero Desktop (Windows, port 23119)
    └── LibreOffice tools → LibreOffice Writer (VirtualBox Ubuntu VM, port 8765)

Required environment variables:
  ZOTERO_HOST      IP/host of the machine running Zotero (default: auto-detect Windows host)
  ZOTERO_API_KEY   Zotero Web API key (alternative to local API)
  ZOTERO_LIBRARY_ID  Zotero library/user ID (required for web API; leave unset for local)
  LIBREOFFICE_URL  Full URL to LibreOffice extension (default: http://VBOX_VM_IP:8765)
  VBOX_VM_IP       IP of the VirtualBox VM running LibreOffice (required for LibreOffice tools)

Enabling Zotero local API:
  In Zotero: Edit → Preferences → Advanced → "Allow other applications on this computer
  to communicate with Zotero" must be checked. Then port 23119 must be reachable from WSL.
  Easiest fix: add a Windows Firewall inbound rule for TCP port 23119.
"""

import os
import subprocess
import json
import re
import shutil
import uuid
import zipfile
from typing import Optional
from pathlib import Path
from xml.etree import ElementTree as ET
import httpx
from fastmcp import FastMCP

# Load .env from repo directory if present (machine-specific overrides)
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env")
except ImportError:
    pass

mcp = FastMCP("AcademicWriting")


# ---------------------------------------------------------------------------
# Network helpers
# ---------------------------------------------------------------------------

def _win_host_ip() -> str:
    """Auto-detect the Windows host IP from WSL2's default route."""
    try:
        r = subprocess.run(["ip", "route", "show", "default"], capture_output=True, text=True)
        for line in r.stdout.split("\n"):
            if "default" in line and "via" in line:
                parts = line.split()
                return parts[parts.index("via") + 1]
    except Exception:
        pass
    return "172.28.32.1"  # fallback common WSL2 gateway


def _zotero_base() -> str:
    """Base URL for the Zotero local API."""
    host = os.environ.get("ZOTERO_HOST") or _win_host_ip()
    return f"http://{host}:23119/api"


def _libre_url() -> str:
    """Base URL for the LibreOffice extension HTTP API."""
    if url := os.environ.get("LIBREOFFICE_URL"):
        return url.rstrip("/")
    vm_ip = os.environ.get("VBOX_VM_IP", "")
    if vm_ip:
        return f"http://{vm_ip}:8765"
    # VirtualBox host-only network: Windows side is 192.168.56.1, VM is typically .101+
    return "http://192.168.56.101:8765"


# ---------------------------------------------------------------------------
# Zotero helpers — direct HTTP to local API (supports custom host)
# ---------------------------------------------------------------------------

def _zotero_get(path: str, params: dict = None) -> dict | list:
    """
    Call the Zotero local API directly via HTTP so we can target the Windows host.
    Falls back to the Zotero Web API if ZOTERO_API_KEY is set.

    Zotero's local API (httpd.js) requires Host header to include the port:
      Host: localhost:23119  ← must include :23119, plain 'localhost' is rejected.
    """
    api_key = os.environ.get("ZOTERO_API_KEY", "")
    library_id = os.environ.get("ZOTERO_LIBRARY_ID", "")

    if api_key and library_id:
        # Web API mode
        base = f"https://api.zotero.org/users/{library_id}"
        headers = {"Zotero-API-Key": api_key, "Zotero-API-Version": "3"}
    else:
        # Local API mode — hits the Windows host's Zotero via WSL→Windows portproxy.
        # Zotero rejects requests unless Host: localhost:23119 is sent exactly.
        host = os.environ.get("ZOTERO_HOST") or _win_host_ip()
        base = f"http://{host}:23119/api/users/0"
        headers = {
            "Host": "localhost:23119",
            "Zotero-API-Version": "3",
        }

    with httpx.Client(timeout=15, headers=headers) as client:
        resp = client.get(f"{base}{path}", params=params or {})
        resp.raise_for_status()
        return resp.json()


def _zotero_post_raw(path: str, payload: dict) -> httpx.Response:
    """POST to Zotero/BBT local HTTP API and return the raw response."""
    host = os.environ.get("ZOTERO_HOST") or _win_host_ip()
    url = f"http://{host}:23119{path}"
    headers = {
        "Host": "localhost:23119",
        "Zotero-API-Version": "3",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    with httpx.Client(timeout=30, headers=headers) as client:
        resp = client.post(url, json=payload)
        resp.raise_for_status()
        return resp


def _bbt_rpc(method: str, params: list | None = None) -> dict:
    """
    Call Better BibTeX's JSON-RPC endpoint.

    BBT exposes this at /better-bibtex/json-rpc on the Zotero local HTTP server.
    The Zotero local server still requires Host: localhost:23119.
    """
    payload = {
        "jsonrpc": "2.0",
        "method": method,
        "params": params or [],
        "id": str(uuid.uuid4()),
    }
    resp = _zotero_post_raw("/better-bibtex/json-rpc", payload)
    data = resp.json()
    if "error" in data:
        raise RuntimeError(data["error"])
    return data.get("result")


def _format_scannable_cite(item: dict) -> str:
    """Format a Zotero item dict as an ODF Scannable Cite marker."""
    key = item.get("key", "")
    data = item.get("data", item)
    creators = data.get("creators", [])
    year = str(data.get("date", ""))[:4]

    if creators:
        first = creators[0]
        last = first.get("lastName") or first.get("name", "Unknown")
        if len(creators) == 1:
            author = last
        elif len(creators) == 2:
            second = creators[1].get("lastName") or creators[1].get("name", "")
            author = f"{last} & {second}"
        else:
            author = f"{last} et al."
    else:
        author = "Unknown"

    return f"{{  | {author}, ({year}) |  |  |zu:0:{key}}}"


_ZOTERO_SETUP_HINT = """
Zotero not reachable. To fix:

Option A — Enable local API (recommended):
  1. In Zotero on Windows: Edit → Preferences → Advanced
     → check "Allow other applications on this computer to communicate with Zotero"
  2. Add Windows Firewall inbound rule for TCP port 23119
     (or run in PowerShell as admin):
     netsh advfirewall firewall add rule name="Zotero Local API" ^
       dir=in action=allow protocol=TCP localport=23119

Option B — Use Zotero Web API:
  1. Go to https://www.zotero.org/settings/keys → create a new API key
  2. Find your library ID at https://www.zotero.org/settings/keys
  3. Add to ~/.claude/settings.json under "writing" → "env":
       "ZOTERO_API_KEY": "your-key",
       "ZOTERO_LIBRARY_ID": "your-numeric-library-id"
"""


# ---------------------------------------------------------------------------
# Zotero tools
# ---------------------------------------------------------------------------

@mcp.tool
def zotero_search(query: str, limit: int = 10) -> str:
    """
    Search the Zotero library and return items with scannable cite markers.

    Args:
        query: Search terms (title, author, keyword)
        limit: Max results (default 10)
    """
    try:
        items = _zotero_get("/items", {"q": query, "limit": limit, "itemType": "-attachment"})
        results = []
        for item in items:
            data = item.get("data", {})
            creators = data.get("creators", [])
            author_str = ", ".join(
                c.get("lastName") or c.get("name", "") for c in creators[:3]
            )
            if len(creators) > 3:
                author_str += " et al."
            results.append({
                "key": item.get("key", ""),
                "title": data.get("title", ""),
                "authors": author_str,
                "year": str(data.get("date", ""))[:4],
                "item_type": data.get("itemType", ""),
                "scannable_cite": _format_scannable_cite(item),
            })
        return json.dumps(results, indent=2)
    except Exception as e:
        return f"Error: {e}\n{_ZOTERO_SETUP_HINT}"


@mcp.tool
def zotero_get_cite(key: str) -> str:
    """
    Get the scannable cite marker for a Zotero item by its key.

    Args:
        key: Zotero item key (e.g. 'ABCD1234')

    Returns:
        Scannable cite: {  | Author, (Year) |  |  |zu:0:KEY}
    """
    try:
        item = _zotero_get(f"/items/{key}")
        return _format_scannable_cite(item)
    except Exception as e:
        return f"Error fetching {key}: {e}\n{_ZOTERO_SETUP_HINT}"


@mcp.tool
def zotero_get_cites_batch(keys: list[str]) -> str:
    """
    Get scannable cite markers for multiple Zotero items.

    Args:
        keys: List of Zotero item keys
    """
    results = {}
    for key in keys:
        try:
            item = _zotero_get(f"/items/{key}")
            results[key] = _format_scannable_cite(item)
        except Exception as e:
            results[key] = f"ERROR: {e}"
    return json.dumps(results, indent=2)


@mcp.tool
def zotero_collections() -> str:
    """List all Zotero collections with their keys and item counts."""
    try:
        colls = _zotero_get("/collections")
        result = [
            {
                "key": c.get("key"),
                "name": c.get("data", {}).get("name"),
                "parent": c.get("data", {}).get("parentCollection"),
                "items": c.get("data", {}).get("numItems", 0),
            }
            for c in colls
        ]
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error: {e}\n{_ZOTERO_SETUP_HINT}"


@mcp.tool
def zotero_check_connection() -> str:
    """
    Test the Zotero connection and show configuration status.
    Run this first to verify the setup is working.
    """
    win_ip = _win_host_ip()
    api_key = os.environ.get("ZOTERO_API_KEY", "")
    library_id = os.environ.get("ZOTERO_LIBRARY_ID", "")

    if api_key and library_id:
        mode = f"Web API (library_id={library_id})"
        url = f"https://api.zotero.org/users/{library_id}/items?limit=1"
        headers = {"Zotero-API-Key": api_key}
    else:
        mode = f"Local API (host={win_ip}, portproxy → Windows:23119)"
        url = f"http://{win_ip}:23119/api/users/0/items?limit=1"
        # Must include port in Host header — Zotero rejects bare 'localhost'
        headers = {"Host": "localhost:23119", "Zotero-API-Version": "3"}

    try:
        with httpx.Client(timeout=5) as client:
            resp = client.get(url, headers=headers)
        status = "CONNECTED" if resp.status_code < 400 else f"HTTP {resp.status_code}"
        count = resp.headers.get("Total-Results", "?")
        return json.dumps({"status": status, "mode": mode, "total_items": count}, indent=2)
    except Exception as e:
        return json.dumps({
            "status": "FAILED",
            "mode": mode,
            "error": str(e),
            "fix": _ZOTERO_SETUP_HINT,
        }, indent=2)


# ---------------------------------------------------------------------------
# Direct DOCX + Zotero/Better BibTeX tools
#
# These tools enhance, but do not replace, the Zotero Word plugin. They can
# insert Zotero-compatible Word fields into a DOCX so Word/Zotero can refresh
# numbering and bibliography later. Final citation formatting is still done by
# Zotero in Word.
# ---------------------------------------------------------------------------

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_XML_NS = "http://www.w3.org/XML/1998/namespace"
ET.register_namespace("w", _W_NS)
ET.register_namespace("xml", _XML_NS)


def _tc_date() -> str:
    """Current UTC datetime in Word track-change format."""
    from datetime import datetime, timezone
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _max_wid(xml: str) -> int:
    """Return the maximum w:id integer present in document XML."""
    ids = [int(x) for x in re.findall(r'w:id="(\d+)"', xml)]
    return max(ids) if ids else 0


def _set_xml_space(el: ET.Element) -> None:
    el.set(f"{{{_XML_NS}}}space", "preserve")


def _make_tracked_replacement(
    old_text: str,
    new_text: str,
    del_id: int,
    ins_id: int,
    author: str,
    date: str,
    template_run: ET.Element | None = None,
) -> list[ET.Element]:
    """Return [<w:del>, <w:ins>] elements for a tracked find-replace."""
    rpr_src = _clone_rpr(template_run) if template_run is not None else None

    def _rpr_copy() -> ET.Element | None:
        if rpr_src is None:
            return None
        return ET.fromstring(ET.tostring(rpr_src))

    # <w:del>
    del_el = ET.Element(_w("del"))
    del_el.set(_w("id"), str(del_id))
    del_el.set(_w("author"), author)
    del_el.set(_w("date"), date)
    del_run = ET.SubElement(del_el, _w("r"))
    rpr = _rpr_copy()
    if rpr is not None:
        del_run.insert(0, rpr)
    dt = ET.SubElement(del_run, _w("delText"))
    dt.text = old_text
    if old_text != old_text.strip():
        _set_xml_space(dt)

    # <w:ins>
    ins_el = ET.Element(_w("ins"))
    ins_el.set(_w("id"), str(ins_id))
    ins_el.set(_w("author"), author)
    ins_el.set(_w("date"), date)
    ins_run = ET.SubElement(ins_el, _w("r"))
    rpr = _rpr_copy()
    if rpr is not None:
        ins_run.insert(0, rpr)
    it = ET.SubElement(ins_run, _w("t"))
    it.text = new_text
    if new_text != new_text.strip():
        _set_xml_space(it)

    return [del_el, ins_el]


def _register_doc_namespaces(xml: str) -> None:
    """Register every xmlns:prefix found in document XML so ET preserves original prefixes."""
    for prefix, uri in re.findall(r'xmlns:([A-Za-z0-9_\-]+)="([^"]+)"', xml):
        try:
            ET.register_namespace(prefix, uri)
        except Exception:
            pass


def _restore_root_namespaces(orig_xml: str, new_xml_bytes: bytes) -> bytes:
    """Re-inject namespace declarations ET drops for namespaces unused in the element tree.

    ET only emits xmlns: for namespaces that appear in the serialized elements.
    Word requires all original root declarations to be present (e.g. w15, w16*, wpc, cx*).
    """
    # Scan only the root element region (first 4000 chars is enough)
    orig_ns = dict(re.findall(r'xmlns:([A-Za-z0-9_\-]+)="([^"]+)"', orig_xml[:4000]))
    new_str = new_xml_bytes.decode("utf-8")
    new_ns = dict(re.findall(r'xmlns:([A-Za-z0-9_\-]+)="([^"]+)"', new_str[:4000]))
    missing = {k: v for k, v in orig_ns.items() if k not in new_ns}
    # Also restore standalone="yes" in the XML declaration if original had it
    if "standalone='yes'" in orig_xml[:200] or 'standalone="yes"' in orig_xml[:200]:
        new_str = re.sub(
            r"(<\?xml[^?]*)\?>",
            r"\1 standalone='yes'?>",
            new_str,
            count=1,
        )
    if not missing:
        return new_str.encode("utf-8")
    inject = " ".join(f'xmlns:{k}="{v}"' for k, v in sorted(missing.items()))
    new_str = re.sub(r'(<w:document\s)', r'\1' + inject + ' ', new_str, count=1)
    return new_str.encode("utf-8")


def _w(tag: str) -> str:
    return f"{{{_W_NS}}}{tag}"


def _as_path(path: str) -> Path:
    return Path(os.path.expanduser(path)).resolve()


def _docx_xml(docx_path: Path, member: str = "word/document.xml") -> str:
    with zipfile.ZipFile(docx_path, "r") as zf:
        return zf.read(member).decode("utf-8")


def _docx_visible_paragraphs(document_xml: str) -> list[dict]:
    root = ET.fromstring(document_xml.encode("utf-8"))
    paragraphs = []
    for i, p in enumerate(root.iter(_w("p")), start=1):
        texts = []
        for t in p.iter(_w("t")):
            texts.append(t.text or "")
        text = "".join(texts)
        if text.strip():
            paragraphs.append({"n": i, "text": text})
    return paragraphs


def _docx_live_zotero_summary(document_xml: str) -> dict:
    citation_fields = document_xml.count("CSL_CITATION")
    bibliography_fields = document_xml.count("CSL_BIBLIOGRAPHY")
    field_begins = document_xml.count('w:fldCharType="begin"')
    instr_text_runs = document_xml.count("<w:instrText")
    item_keys = sorted(set(re.findall(r"/items/([A-Z0-9]{8})", document_xml)))
    return {
        "csl_citation_marker_count": citation_fields,
        "csl_bibliography_marker_count": bibliography_fields,
        "field_begin_count": field_begins,
        "instr_text_run_count": instr_text_runs,
        "unique_zotero_item_keys_in_live_fields": len(item_keys),
        "zotero_item_keys": item_keys,
    }


def _docx_copy_with_replaced_xml(input_path: Path, output_path: Path, replacements: dict[str, bytes]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(input_path, "r") as zin, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        written = set()
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in replacements:
                data = replacements[item.filename]
            zout.writestr(item, data)
            written.add(item.filename)
        # Write new files that didn't exist in the input archive (e.g. comments.xml)
        for name, data in replacements.items():
            if name not in written:
                zout.writestr(name, data)


def _minimal_csl_from_zotero_item(item: dict) -> dict:
    """Fallback Zotero item → minimal CSL JSON converter."""
    data = item.get("data", item)
    creators = data.get("creators", [])
    csl = {
        "id": data.get("key") or item.get("key"),
        "type": {
            "journalArticle": "article-journal",
            "conferencePaper": "paper-conference",
            "bookSection": "chapter",
            "report": "report",
            "webpage": "webpage",
            "book": "book",
            "thesis": "thesis",
        }.get(data.get("itemType"), data.get("itemType", "article-journal")),
        "title": data.get("title", ""),
    }
    if data.get("date"):
        year = str(data.get("date"))[:4]
        if year.isdigit():
            csl["issued"] = {"date-parts": [[int(year)]]}
    if data.get("DOI"):
        csl["DOI"] = data.get("DOI")
    if data.get("url"):
        csl["URL"] = data.get("url")
    if data.get("publicationTitle"):
        csl["container-title"] = data.get("publicationTitle")
    if data.get("journalAbbreviation"):
        csl["journalAbbreviation"] = data.get("journalAbbreviation")
    if data.get("volume"):
        csl["volume"] = data.get("volume")
    if data.get("issue"):
        csl["issue"] = data.get("issue")
    if data.get("pages"):
        csl["page"] = data.get("pages")
    authors = []
    for creator in creators:
        if creator.get("creatorType") not in ("author", None):
            continue
        if creator.get("lastName") or creator.get("firstName"):
            authors.append({"family": creator.get("lastName", ""), "given": creator.get("firstName", "")})
        elif creator.get("name"):
            authors.append({"literal": creator.get("name")})
    if authors:
        csl["author"] = authors
    return csl


def _get_csl_item_data(zotero_key: str) -> tuple[dict, dict]:
    """
    Return (normal Zotero item, CSL itemData) for a Zotero item key.
    Uses Zotero's csljson format when available, falls back to a local converter.
    """
    item = _zotero_get(f"/items/{zotero_key}")
    try:
        csl = _zotero_get(f"/items/{zotero_key}", {"format": "csljson"})
        if isinstance(csl, list) and csl:
            csl = csl[0]
        if isinstance(csl, dict) and csl:
            return item, csl
    except Exception:
        pass
    return item, _minimal_csl_from_zotero_item(item)


def _short_author_year(csl_items: list[dict]) -> str:
    if not csl_items:
        return "[Zotero citation]"
    item = csl_items[0]
    authors = item.get("author") or []
    if authors:
        first = authors[0]
        family = first.get("family") or first.get("literal") or "Unknown"
        label = f"{family} et al." if len(authors) > 2 else family
    else:
        label = "Unknown"
    issued = item.get("issued", {}).get("date-parts", [[""]])
    year = issued[0][0] if issued and issued[0] else ""
    return f"[{label}, {year}]"


def _extract_local_zotero_hash(xml: str) -> str | None:
    """Extract the local Zotero user hash from existing CSL_CITATION fields in document XML."""
    m = re.search(r'http://zotero\.org/users/local/([A-Za-z0-9]+)/items/', xml)
    return m.group(1) if m else None


def _citation_field_json(zotero_keys: list[str], display_text: str, local_hash: str | None = None) -> dict:
    citation_items = []
    csl_items = []
    for key in zotero_keys:
        item, csl = _get_csl_item_data(key)
        library = item.get("library", {}) if isinstance(item, dict) else {}
        library_id = library.get("id", 0)
        # Use the Zotero item key as the canonical CSL id (not the BibTeX key).
        csl["id"] = key
        uris = []
        if local_hash:
            uris.append(f"http://zotero.org/users/local/{local_hash}/items/{key}")
        if library_id:
            uris.append(f"http://zotero.org/users/{library_id}/items/{key}")
        if not uris:
            uris = [f"http://zotero.org/users/local/{key}/items/{key}"]
        csl_items.append(csl)
        citation_items.append({
            "id": key,
            "uris": uris,
            "itemData": csl,
            "locator": "",
            "prefix": "",
            "suffix": "",
        })
    if not display_text:
        display_text = _short_author_year(csl_items)
    return {
        "citationID": uuid.uuid4().hex[:8],
        "properties": {
            "formattedCitation": display_text,
            "plainCitation": display_text,
            "noteIndex": 0,
        },
        "citationItems": citation_items,
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",
    }


def _clone_rpr(template_run: ET.Element | None) -> ET.Element | None:
    if template_run is None:
        return None
    rpr = template_run.find(_w("rPr"))
    if rpr is None:
        return None
    return ET.fromstring(ET.tostring(rpr, encoding="utf-8"))


def _make_run(child: ET.Element, template_run: ET.Element | None = None) -> ET.Element:
    run = ET.Element(_w("r"))
    rpr = _clone_rpr(template_run)
    if rpr is not None:
        run.append(rpr)
    run.append(child)
    return run


def _make_text_run(text: str, template_run: ET.Element | None = None) -> ET.Element:
    t = ET.Element(_w("t"))
    if text.startswith(" ") or text.endswith(" "):
        t.set(f"{{{_XML_NS}}}space", "preserve")
    t.text = text
    return _make_run(t, template_run)


def _make_zotero_field_runs(field_json: dict, display_text: str, template_run: ET.Element | None = None) -> list[ET.Element]:
    instr = " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(field_json, ensure_ascii=False, separators=(",", ":")) + " "

    begin = ET.Element(_w("fldChar"))
    begin.set(_w("fldCharType"), "begin")

    instr_el = ET.Element(_w("instrText"))
    instr_el.set(f"{{{_XML_NS}}}space", "preserve")
    instr_el.text = instr

    separate = ET.Element(_w("fldChar"))
    separate.set(_w("fldCharType"), "separate")

    visible = ET.Element(_w("t"))
    if display_text.startswith(" ") or display_text.endswith(" "):
        visible.set(f"{{{_XML_NS}}}space", "preserve")
    visible.text = display_text

    end = ET.Element(_w("fldChar"))
    end.set(_w("fldCharType"), "end")

    return [
        _make_run(begin, template_run),
        _make_run(instr_el, template_run),
        _make_run(separate, template_run),
        _make_run(visible, template_run),
        _make_run(end, template_run),
    ]


def _replacement_parts(old_text: str, zotero_keys: list[str], display_text: str | None, keep_prefix_text: bool, local_hash: str | None = None) -> tuple[str, list[ET.Element]]:
    prefix = ""
    field_display = display_text
    if keep_prefix_text and display_text is None:
        match = re.match(r"^(.*?)(\[[0-9,\-\s;]+\])([,.;:]?)$", old_text)
        if match:
            prefix = match.group(1)
            field_display = "[Zotero citation]"
            suffix = match.group(3)
        else:
            suffix = ""
            field_display = "[Zotero citation]"
    else:
        suffix = ""
        field_display = display_text or "[Zotero citation]"

    field_json = _citation_field_json(zotero_keys, field_display, local_hash=local_hash)
    return prefix, _make_zotero_field_runs(field_json, field_display) + ([_make_text_run(suffix)] if suffix else [])


def _replace_text_in_single_run(root: ET.Element, old_text: str, zotero_keys: list[str], display_text: str | None, keep_prefix_text: bool, local_hash: str | None = None) -> int:
    count = 0
    for p in root.iter(_w("p")):
        runs = list(p.findall(_w("r")))
        for run in runs:
            text_nodes = [t for t in run.findall(_w("t")) if t.text and old_text in t.text]
            if not text_nodes:
                continue
            t = text_nodes[0]
            before, after = t.text.split(old_text, 1)
            idx = list(p).index(run)
            new_runs = []
            if before:
                new_runs.append(_make_text_run(before, run))
            prefix, field_runs = _replacement_parts(old_text, zotero_keys, display_text, keep_prefix_text, local_hash=local_hash)
            if prefix:
                new_runs.append(_make_text_run(prefix, run))
            for fr in field_runs:
                # Add run properties matching the source run if the helper had none.
                if fr.find(_w("rPr")) is None:
                    rpr = _clone_rpr(run)
                    if rpr is not None:
                        fr.insert(0, rpr)
                new_runs.append(fr)
            if after:
                new_runs.append(_make_text_run(after, run))
            p.remove(run)
            for offset, nr in enumerate(new_runs):
                p.insert(idx + offset, nr)
            count += 1
            break
    return count


def _replace_text_in_adjacent_runs(
    root: ET.Element,
    old_text: str,
    zotero_keys: list[str],
    display_text: str | None,
    keep_prefix_text: bool,
    local_hash: str | None = None,
) -> int:
    """Replace old_text that spans adjacent runs within a paragraph.

    Handles cases like run_i="Lu et al. " / run_{i+1}="[46]," where
    old_text="Lu et al. [46]," is never contained in any single run.
    """
    count = 0
    for p in root.iter(_w("p")):
        while True:
            runs = [c for c in p if c.tag == _w("r")]
            if len(runs) < 2:
                break
            run_texts = ["".join(t.text or "" for t in r.findall(_w("t"))) for r in runs]
            concat = "".join(run_texts)
            if old_text not in concat:
                break
            # Skip if a single run already contains it
            if any(old_text in rt for rt in run_texts):
                break

            start_pos = concat.index(old_text)
            end_pos = start_pos + len(old_text)

            # Map each char position to (run_idx, offset_within_run)
            char_map: list[tuple[int, int]] = []
            for ri, rt in enumerate(run_texts):
                for ci in range(len(rt)):
                    char_map.append((ri, ci))

            if start_pos >= len(char_map) or end_pos - 1 >= len(char_map):
                break

            start_run_idx, start_char = char_map[start_pos]
            end_run_idx, end_char = char_map[end_pos - 1]

            if start_run_idx == end_run_idx:
                break  # shouldn't happen given the single-run check above

            prefix, field_runs = _replacement_parts(old_text, zotero_keys, display_text, keep_prefix_text, local_hash=local_hash)
            template = runs[start_run_idx]

            before_text = run_texts[start_run_idx][:start_char]
            after_text = run_texts[end_run_idx][end_char + 1:]

            new_runs: list[ET.Element] = []
            if before_text:
                new_runs.append(_make_text_run(before_text, template))
            if prefix:
                new_runs.append(_make_text_run(prefix, template))
            for fr in field_runs:
                if fr.find(_w("rPr")) is None:
                    rpr = _clone_rpr(template)
                    if rpr is not None:
                        fr.insert(0, rpr)
                new_runs.append(fr)
            if after_text:
                new_runs.append(_make_text_run(after_text, template))

            # Remove middle runs that are fully consumed by old_text
            # (runs between start_run_idx+1 and end_run_idx-1 are fully inside old_text)
            insert_idx = list(p).index(runs[start_run_idx])
            for ri in range(start_run_idx, end_run_idx + 1):
                p.remove(runs[ri])
            for offset, nr in enumerate(new_runs):
                p.insert(insert_idx + offset, nr)
            count += 1

    return count


def _rewrite_plain_paragraph(root: ET.Element, old_text: str, zotero_keys: list[str], display_text: str | None, keep_prefix_text: bool, local_hash: str | None = None) -> int:
    """
    Fallback replacement for paragraphs where old_text spans multiple runs.
    This preserves paragraph properties but rewrites the paragraph's runs.
    """
    count = 0
    for p in root.iter(_w("p")):
        runs = list(p.findall(_w("r")))
        para_text = "".join((t.text or "") for r in runs for t in r.findall(_w("t")))
        if old_text not in para_text:
            continue
        before, after = para_text.split(old_text, 1)
        template = runs[0] if runs else None
        ppr = p.find(_w("pPr"))
        for child in list(p):
            if child is not ppr:
                p.remove(child)
        insert_at = 1 if ppr is not None else 0
        new_runs = []
        if before:
            new_runs.append(_make_text_run(before, template))
        prefix, field_runs = _replacement_parts(old_text, zotero_keys, display_text, keep_prefix_text, local_hash=local_hash)
        if prefix:
            new_runs.append(_make_text_run(prefix, template))
        new_runs.extend(field_runs)
        if after:
            new_runs.append(_make_text_run(after, template))
        for offset, run in enumerate(new_runs):
            p.insert(insert_at + offset, run)
        count += 1
    return count


@mcp.tool
def betterbibtex_check_connection() -> str:
    """
    Test Better BibTeX JSON-RPC connectivity through Zotero's local API.
    """
    try:
        result = _bbt_rpc("api.ready")
        return json.dumps({"status": "CONNECTED", "result": result}, indent=2)
    except Exception as e:
        return json.dumps({
            "status": "FAILED",
            "error": str(e),
            "hint": (
                "Zotero must be running with Better BibTeX installed. "
                "The endpoint is /better-bibtex/json-rpc on Zotero's local server."
            ),
        }, indent=2)


@mcp.tool
def betterbibtex_citation_keys(item_keys: list[str]) -> str:
    """
    Fetch Better BibTeX citation keys for Zotero item keys.

    Args:
        item_keys: Zotero item keys, e.g. ["WIVFG3LY"].
    """
    try:
        qualified = [k if ":" in k else f"0:{k}" for k in item_keys]
        result = _bbt_rpc("item.citationkey", [qualified])
        return json.dumps({"status": "ok", "citation_keys": result}, indent=2)
    except Exception as e:
        return json.dumps({"status": "error", "error": str(e)}, indent=2)


@mcp.tool
def betterbibtex_export_items(citekeys: list[str], translator: str = "Better BibTeX") -> str:
    """
    Export Zotero items through Better BibTeX by citation key.

    Args:
        citekeys: Better BibTeX citation keys.
        translator: BBT translator name, e.g. "Better BibTeX", "Better BibLaTeX",
                    or "Better CSL JSON".
    """
    try:
        result = _bbt_rpc("item.export", [citekeys, translator])
        return json.dumps({"status": "ok", "translator": translator, "export": result}, indent=2)
    except Exception as e:
        return json.dumps({"status": "error", "error": str(e)}, indent=2)


@mcp.tool
def docx_extract_text(docx_path: str) -> str:
    """
    Extract all visible paragraph text from a DOCX without LibreOffice.

    Returns each non-empty paragraph with its 1-based index, plus the full
    concatenated text. Useful for reading document content before editing.

    Args:
        docx_path: Path to a .docx file.
    """
    path = _as_path(docx_path)
    if not path.exists():
        return json.dumps({"status": "error", "error": f"file not found: {path}"}, indent=2)
    try:
        xml = _docx_xml(path)
        paragraphs = _docx_visible_paragraphs(xml)
        full_text = "\n".join(p["text"] for p in paragraphs)
        return json.dumps({
            "status": "ok",
            "docx_path": str(path),
            "paragraph_count": len(paragraphs),
            "paragraphs": paragraphs,
            "full_text": full_text,
        }, indent=2)
    except Exception as e:
        return json.dumps({"status": "error", "error": str(e)}, indent=2)


@mcp.tool
def docx_get_headings(docx_path: str) -> str:
    """
    Extract the heading structure (outline) of a DOCX without LibreOffice.

    Returns headings in document order with their level (1 = Heading 1,
    2 = Heading 2, etc.) and paragraph index. Useful for navigating large
    manuscripts before targeted edits.

    Args:
        docx_path: Path to a .docx file.
    """
    path = _as_path(docx_path)
    if not path.exists():
        return json.dumps({"status": "error", "error": f"file not found: {path}"}, indent=2)
    try:
        xml = _docx_xml(path)
        _register_doc_namespaces(xml)
        root = ET.fromstring(xml.encode("utf-8"))
        headings = []
        for i, p in enumerate(root.iter(_w("p")), start=1):
            ppr = p.find(_w("pPr"))
            if ppr is None:
                continue
            pstyle = ppr.find(_w("pStyle"))
            if pstyle is None:
                continue
            val = pstyle.get(_w("val"), "")
            m = re.match(r"[Hh]eading(\d+)", val)
            if not m:
                continue
            level = int(m.group(1))
            text = "".join(t.text or "" for t in p.iter(_w("t"))).strip()
            if text:
                headings.append({"paragraph_index": i, "level": level, "text": text})
        return json.dumps({
            "status": "ok",
            "docx_path": str(path),
            "heading_count": len(headings),
            "headings": headings,
        }, indent=2)
    except Exception as e:
        return json.dumps({"status": "error", "error": str(e)}, indent=2)


@mcp.tool
def docx_zotero_audit(docx_path: str, suspect_terms: Optional[list[str]] = None) -> str:
    """
    Audit a DOCX for live Zotero fields and plain-text citation-like labels.

    Args:
        docx_path: Path to a .docx file.
        suspect_terms: Optional author/name fragments to find in visible text.
    """
    path = _as_path(docx_path)
    if not path.exists():
        return json.dumps({"status": "error", "error": f"file not found: {path}"}, indent=2)
    try:
        xml = _docx_xml(path)
        paragraphs = _docx_visible_paragraphs(xml)
        visible_text = "\n".join(p["text"] for p in paragraphs)
        numeric_refs = sorted(set(int(n) for n in re.findall(r"\[([0-9]{1,3})\]", visible_text)))
        suspect_hits = []
        if suspect_terms:
            lowered = [(term, term.lower()) for term in suspect_terms]
            for p in paragraphs:
                low = p["text"].lower()
                for term, term_low in lowered:
                    if term_low in low:
                        suspect_hits.append({"paragraph": p["n"], "term": term, "text": p["text"][:600]})

        summary = _docx_live_zotero_summary(xml)
        summary.update({
            "status": "ok",
            "docx_path": str(path),
            "visible_numeric_reference_labels": numeric_refs,
            "visible_numeric_reference_label_count": len(numeric_refs),
            "suspect_hits": suspect_hits[:200],
            "note": (
                "Visible numeric labels can include live Zotero field results and plain text. "
                "Use suspect_terms or exact replacement auditing to distinguish them."
            ),
        })
        return json.dumps(summary, indent=2)
    except Exception as e:
        return json.dumps({"status": "error", "error": str(e)}, indent=2)


@mcp.tool
def docx_text_replace(
    docx_path: str,
    replacements: list[dict],
    output_path: Optional[str] = None,
    overwrite: bool = False,
    track_changes: bool = False,
    author: str = "writing-mcp",
) -> str:
    """
    Find-and-replace plain text in a DOCX without LibreOffice.

    Edits word/document.xml directly, preserving all images, styles,
    formatting, namespace declarations, and Zotero citation fields.

    Args:
        docx_path: Input .docx path.
        replacements: List of {\"find\": \"old text\", \"replace\": \"new text\"}.
            find must appear within a single Word run (no cross-run spans).
        output_path: Destination .docx. Defaults to *_edited.docx.
        overwrite: Allow writing over output_path if it already exists.
        track_changes: If true, wrap edits as Word tracked changes
            (<w:del>/<w:ins>) so the author can Accept or Reject them.
        author: Author name shown in tracked changes. Default \"writing-mcp\".
    """
    input_path = _as_path(docx_path)
    if not input_path.exists():
        return json.dumps({"status": "error", "error": f"file not found: {input_path}"}, indent=2)
    out_path = _as_path(output_path) if output_path else \
        input_path.with_name(f"{input_path.stem}_edited{input_path.suffix}")
    if out_path.exists() and out_path != input_path and not overwrite:
        return json.dumps({"status": "error", "error": f"output exists: {out_path}", "set_overwrite": True}, indent=2)

    try:
        xml = _docx_xml(input_path)
        _register_doc_namespaces(xml)
        root = ET.fromstring(xml.encode("utf-8"))
        next_id = _max_wid(xml) + 1
        date = _tc_date()
        results = []

        for repl in replacements:
            find = repl.get("find", "")
            replace = repl.get("replace", "")
            if not find:
                results.append({"find": find, "status": "skipped", "reason": "find string required"})
                continue
            count = 0
            for p in root.iter(_w("p")):
                for run in list(p.findall(_w("r"))):
                    t_nodes = [t for t in run.findall(_w("t")) if t.text and find in t.text]
                    if not t_nodes:
                        continue
                    t = t_nodes[0]
                    before, after = t.text.split(find, 1)
                    idx = list(p).index(run)
                    new_nodes: list[ET.Element] = []
                    if before:
                        new_nodes.append(_make_text_run(before, run))
                    if track_changes:
                        tc_elems = _make_tracked_replacement(
                            find, replace, next_id, next_id + 1, author, date, run
                        )
                        next_id += 2
                        new_nodes.extend(tc_elems)
                    else:
                        new_nodes.append(_make_text_run(replace, run))
                    if after:
                        new_nodes.append(_make_text_run(after, run))
                    p.remove(run)
                    for offset, node in enumerate(new_nodes):
                        p.insert(idx + offset, node)
                    count += 1
            results.append({"find": find, "replace": replace, "occurrences_edited": count,
                            "tracked": track_changes,
                            "status": "ok" if count else "not_found"})

        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        new_xml = _restore_root_namespaces(xml, new_xml)
        if out_path == input_path:
            backup = input_path.with_name(f"{input_path.stem}.bak{input_path.suffix}")
            shutil.copy2(input_path, backup)
        _docx_copy_with_replaced_xml(input_path, out_path, {"word/document.xml": new_xml})
        with zipfile.ZipFile(out_path, "r") as zf:
            bad = zf.testzip()
        return json.dumps({
            "status": "ok" if bad is None else "zip_error",
            "output_path": str(out_path),
            "results": results,
        }, indent=2)
    except Exception as e:
        return json.dumps({"status": "error", "error": str(e)}, indent=2)


@mcp.tool
def docx_add_comment(
    docx_path: str,
    comments: list[dict],
    output_path: Optional[str] = None,
    overwrite: bool = False,
    author: str = "writing-mcp",
    initials: str = "WM",
) -> str:
    """
    Add review comments anchored to specific text in a DOCX.

    Comments appear in Word's comment sidebar. Existing comments and
    track changes are preserved. Safe for manuscripts with Zotero fields.

    Args:
        docx_path: Input .docx path.
        comments: List of {\"find\": \"text to anchor to\", \"comment\": \"comment text\"}.
            find must appear in a single Word run. The comment is anchored
            to the first occurrence.
        output_path: Destination .docx. Defaults to *_commented.docx.
        overwrite: Allow writing over output_path.
        author: Author shown in comment bubble. Default \"writing-mcp\".
        initials: Initials shown in comment bubble. Default \"WM\".
    """
    input_path = _as_path(docx_path)
    if not input_path.exists():
        return json.dumps({"status": "error", "error": f"file not found: {input_path}"}, indent=2)
    out_path = _as_path(output_path) if output_path else \
        input_path.with_name(f"{input_path.stem}_commented{input_path.suffix}")
    if out_path.exists() and out_path != input_path and not overwrite:
        return json.dumps({"status": "error", "error": f"output exists: {out_path}", "set_overwrite": True}, indent=2)

    try:
        xml = _docx_xml(input_path)
        _register_doc_namespaces(xml)
        root = ET.fromstring(xml.encode("utf-8"))
        next_id = _max_wid(xml) + 1
        date = _tc_date()

        # Load or create comments.xml
        with zipfile.ZipFile(input_path, "r") as zf:
            all_files = zf.namelist()
            if "word/comments.xml" in all_files:
                cxml = zf.read("word/comments.xml").decode("utf-8")
                _register_doc_namespaces(cxml)
                croot = ET.fromstring(cxml.encode("utf-8"))
            else:
                # Build minimal comments.xml root
                croot = ET.fromstring(
                    f'<w:comments xmlns:w="{_W_NS}"></w:comments>'
                )
                cxml = ""
            rels_xml = zf.read("word/_rels/document.xml.rels").decode("utf-8")
            content_types_xml = zf.read("[Content_Types].xml").decode("utf-8")

        results = []
        for item in comments:
            find = item.get("find", "")
            comment_text = item.get("comment", "")
            if not find or not comment_text:
                results.append({"find": find, "status": "skipped", "reason": "find and comment required"})
                continue

            # Find the run containing the text
            found = False
            for p in root.iter(_w("p")):
                for run in list(p.findall(_w("r"))):
                    t_nodes = [t for t in run.findall(_w("t")) if t.text and find in t.text]
                    if not t_nodes:
                        continue
                    t = t_nodes[0]
                    before, after = t.text.split(find, 1)
                    idx = list(p).index(run)
                    cid = next_id
                    next_id += 1

                    # Build replacement sequence
                    new_nodes: list[ET.Element] = []
                    if before:
                        new_nodes.append(_make_text_run(before, run))

                    # commentRangeStart
                    crs = ET.Element(_w("commentRangeStart"))
                    crs.set(_w("id"), str(cid))
                    new_nodes.append(crs)

                    # The commented text run (copy rPr from original)
                    text_run = ET.Element(_w("r"))
                    rpr = _clone_rpr(run)
                    if rpr is not None:
                        text_run.insert(0, rpr)
                    ct = ET.SubElement(text_run, _w("t"))
                    ct.text = find
                    if find != find.strip():
                        _set_xml_space(ct)
                    new_nodes.append(text_run)

                    # commentRangeEnd
                    cre = ET.Element(_w("commentRangeEnd"))
                    cre.set(_w("id"), str(cid))
                    new_nodes.append(cre)

                    # commentReference run
                    ref_run = ET.Element(_w("r"))
                    ref_rpr = ET.SubElement(ref_run, _w("rPr"))
                    ref_style = ET.SubElement(ref_rpr, _w("rStyle"))
                    ref_style.set(_w("val"), "CommentReference")
                    ref = ET.SubElement(ref_run, _w("commentReference"))
                    ref.set(_w("id"), str(cid))
                    new_nodes.append(ref_run)

                    if after:
                        new_nodes.append(_make_text_run(after, run))

                    p.remove(run)
                    for offset, node in enumerate(new_nodes):
                        p.insert(idx + offset, node)

                    # Add comment to comments.xml
                    comment_el = ET.SubElement(croot, _w("comment"))
                    comment_el.set(_w("id"), str(cid))
                    comment_el.set(_w("author"), author)
                    comment_el.set(_w("date"), date)
                    comment_el.set(_w("initials"), initials)
                    cp = ET.SubElement(comment_el, _w("p"))
                    cp_pr = ET.SubElement(cp, _w("pPr"))
                    ps = ET.SubElement(cp_pr, _w("pStyle"))
                    ps.set(_w("val"), "CommentText")
                    ann_run = ET.SubElement(cp, _w("r"))
                    ann_rpr = ET.SubElement(ann_run, _w("rPr"))
                    ann_style = ET.SubElement(ann_rpr, _w("rStyle"))
                    ann_style.set(_w("val"), "CommentReference")
                    ET.SubElement(ann_run, _w("annotationRef"))
                    txt_run = ET.SubElement(cp, _w("r"))
                    txt_t = ET.SubElement(txt_run, _w("t"))
                    txt_t.text = comment_text

                    results.append({"find": find, "comment_id": cid, "status": "ok"})
                    found = True
                    break  # anchor to first occurrence only
                if found:
                    break
            if not found:
                results.append({"find": find, "status": "not_found"})

        # Serialize both XMLs
        new_doc_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        new_doc_xml = _restore_root_namespaces(xml, new_doc_xml)
        new_c_xml = ET.tostring(croot, encoding="utf-8", xml_declaration=True)
        if cxml:
            new_c_xml = _restore_root_namespaces(cxml, new_c_xml)

        file_replacements = {
            "word/document.xml": new_doc_xml,
            "word/comments.xml": new_c_xml,
        }

        # Ensure the comments part has a content type. Without this Word can
        # report "unreadable content" even when the ZIP/XML are well-formed.
        comments_content_type = (
            '<Override PartName="/word/comments.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
        )
        if 'PartName="/word/comments.xml"' not in content_types_xml:
            content_types_xml = content_types_xml.replace(
                "</Types>",
                comments_content_type + "</Types>",
            )
            file_replacements["[Content_Types].xml"] = content_types_xml.encode("utf-8")

        # Ensure comments relationship exists in document.xml.rels
        COMMENTS_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        if COMMENTS_TYPE not in rels_xml and "word/comments.xml" not in all_files:
            # Find max rId number
            rids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)]
            new_rid = f"rId{max(rids) + 1 if rids else 1}"
            new_rel = f'<Relationship Id="{new_rid}" Type="{COMMENTS_TYPE}" Target="comments.xml"/>'
            rels_xml = rels_xml.replace("</Relationships>", new_rel + "</Relationships>")
            file_replacements["word/_rels/document.xml.rels"] = rels_xml.encode("utf-8")

        _docx_copy_with_replaced_xml(input_path, out_path, file_replacements)
        with zipfile.ZipFile(out_path, "r") as zf:
            bad = zf.testzip()
        return json.dumps({
            "status": "ok" if bad is None else "zip_error",
            "output_path": str(out_path),
            "comments_added": sum(1 for r in results if r.get("status") == "ok"),
            "results": results,
        }, indent=2)
    except Exception as e:
        return json.dumps({"status": "error", "error": str(e)}, indent=2)


@mcp.tool
def docx_zotero_insert_citations(
    docx_path: str,
    replacements: list[dict],
    output_path: Optional[str] = None,
    overwrite: bool = False,
    fallback_paragraph_rewrite: bool = False,
) -> str:
    """
    Replace exact plain-text citation labels in a DOCX with live Zotero Word fields.

    This is for stale plain-text labels such as "Chen & Zhan [78]". It does not
    replace existing live Zotero fields. Word/Zotero must still refresh the
    document to compute final numbering and update the bibliography.

    Args:
        docx_path: Input .docx path.
        replacements: List of replacements. Each item:
            {
              "old_text": "Chen & Zhan [78]",
              "zotero_keys": ["WIVFG3LY"],
              "display_text": "[Zotero citation]",      # optional placeholder
              "keep_prefix_text": true                  # optional, default true
            }
            With keep_prefix_text=true, the tool keeps "Chen & Zhan " as normal
            text and replaces only the bracketed numeric citation with a field.
        output_path: Destination .docx. If omitted, writes *_with_zotero_fields.docx.
        overwrite: If true and output_path == input, modifies in place after creating
            a timestamp-free .bak next to the file.
        fallback_paragraph_rewrite: If true, rewrite whole paragraphs when old_text
            spans multiple runs. Use only for simple figure/table labels.
    """
    input_path = _as_path(docx_path)
    if not input_path.exists():
        return json.dumps({"status": "error", "error": f"file not found: {input_path}"}, indent=2)
    if output_path:
        out_path = _as_path(output_path)
    else:
        out_path = input_path.with_name(f"{input_path.stem}_with_zotero_fields{input_path.suffix}")
    if out_path.exists() and out_path != input_path and not overwrite:
        return json.dumps({"status": "error", "error": f"output exists: {out_path}", "set_overwrite": True}, indent=2)
    if out_path == input_path and not overwrite:
        return json.dumps({"status": "error", "error": "in-place edit requires overwrite=true"}, indent=2)

    try:
        xml = _docx_xml(input_path)
        _register_doc_namespaces(xml)
        root = ET.fromstring(xml.encode("utf-8"))
        local_hash = _extract_local_zotero_hash(xml)
        results = []
        for repl in replacements:
            old_text = repl.get("old_text", "")
            zotero_keys = repl.get("zotero_keys") or repl.get("item_keys") or []
            if isinstance(zotero_keys, str):
                zotero_keys = [zotero_keys]
            display_text = repl.get("display_text")
            keep_prefix_text = bool(repl.get("keep_prefix_text", True))
            if not old_text or not zotero_keys:
                results.append({"old_text": old_text, "status": "skipped", "reason": "old_text and zotero_keys required"})
                continue
            count = _replace_text_in_single_run(root, old_text, zotero_keys, display_text, keep_prefix_text, local_hash=local_hash)
            if count == 0:
                count = _replace_text_in_adjacent_runs(root, old_text, zotero_keys, display_text, keep_prefix_text, local_hash=local_hash)
            if count == 0 and fallback_paragraph_rewrite:
                count = _rewrite_plain_paragraph(root, old_text, zotero_keys, display_text, keep_prefix_text, local_hash=local_hash)
            results.append({
                "old_text": old_text,
                "zotero_keys": zotero_keys,
                "replacements_made": count,
                "status": "ok" if count else "not_found_or_spans_runs",
            })

        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        new_xml = _restore_root_namespaces(xml, new_xml)
        if out_path == input_path:
            backup = input_path.with_name(f"{input_path.stem}.bak{input_path.suffix}")
            shutil.copy2(input_path, backup)
        else:
            backup = None
        _docx_copy_with_replaced_xml(input_path, out_path, {"word/document.xml": new_xml})

        with zipfile.ZipFile(out_path, "r") as zf:
            bad_zip = zf.testzip()
        status = "ok" if bad_zip is None else "zip_error"
        return json.dumps({
            "status": status,
            "input_path": str(input_path),
            "output_path": str(out_path),
            "backup_path": str(backup) if backup else None,
            "zip_test": "ok" if bad_zip is None else bad_zip,
            "results": results,
            "next_step": "Open the DOCX in Word and use Zotero Refresh to renumber citations and update bibliography.",
        }, indent=2)
    except Exception as e:
        return json.dumps({"status": "error", "error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# LibreOffice tools — connects to VirtualBox VM running LibreOffice
#
# Extension API structure:
#   GET  /health                 → status check
#   GET  /tools                  → list available tools
#   POST /tools/{tool_name}      → execute tool with JSON body params
#
# Network setup required (run in Windows PowerShell as admin):
#   # 1. Forward VirtualBox NAT port 8765 from Windows localhost to VM
#   VBoxManage controlvm "Ubuntu24" natpf1 "libreoffice,tcp,,8765,,8765"
#   # 2. Forward from WSL gateway to Windows localhost
#   netsh interface portproxy add v4tov4 listenport=8765 listenaddress=172.28.32.1 connectport=8765 connectaddress=127.0.0.1
#   netsh advfirewall firewall add rule name="LibreOffice WSL Bridge" dir=in action=allow protocol=TCP localport=8765
# ---------------------------------------------------------------------------

_LIBRE_HINT = (
    "LibreOffice not reachable. Setup steps:\n"
    "1. VirtualBox VM must be running with LibreOffice Writer open\n"
    "2. Start MCP extension: Tools → MCP Server → Start MCP Server\n"
    "3. In Windows PowerShell (admin), run:\n"
    '   VBoxManage controlvm "Ubuntu24" natpf1 "libreoffice,tcp,,8765,,8765"\n'
    "   netsh interface portproxy add v4tov4 listenport=8765 listenaddress=172.28.32.1 connectport=8765 connectaddress=127.0.0.1\n"
    '   netsh advfirewall firewall add rule name="LibreOffice WSL Bridge" dir=in action=allow protocol=TCP localport=8765'
)


def _libre_tool(tool_name: str, params: dict = None) -> dict:
    """Call a LibreOffice extension tool via POST /tools/{tool_name}."""
    url = f"{_libre_url()}/tools/{tool_name}"
    try:
        with httpx.Client(timeout=30) as client:
            resp = client.post(url, json=params or {})
        return resp.json()
    except Exception as e:
        return {"error": str(e), "hint": _LIBRE_HINT}


def _libre_get(path: str) -> dict:
    """GET request to the LibreOffice extension."""
    url = f"{_libre_url()}{path}"
    try:
        with httpx.Client(timeout=10) as client:
            resp = client.get(url)
        return resp.json()
    except Exception as e:
        return {"error": str(e), "hint": _LIBRE_HINT}


@mcp.tool
def libre_check_connection() -> str:
    """
    Test LibreOffice connection. Run this first to verify the VM is reachable.
    Shows setup instructions if not connected.
    """
    url = _libre_url()
    try:
        with httpx.Client(timeout=5) as client:
            resp = client.get(f"{url}/health")
        return json.dumps({"status": "CONNECTED", "url": url, "response": resp.json()}, indent=2)
    except Exception as e:
        return json.dumps({
            "status": "FAILED",
            "url": url,
            "error": str(e),
            "setup": _LIBRE_HINT,
        }, indent=2)


@mcp.tool
def libre_content() -> str:
    """Get the full text content of the active LibreOffice document."""
    return json.dumps(_libre_tool("get_text_content_live"), indent=2)


@mcp.tool
def libre_document_info() -> str:
    """Get info about the active document (filename, word count, page count, etc.)."""
    return json.dumps(_libre_tool("get_document_info_live"), indent=2)


@mcp.tool
def libre_list_documents() -> str:
    """List all open LibreOffice documents."""
    return json.dumps(_libre_tool("list_open_documents"), indent=2)


@mcp.tool
def libre_search_replace(old_text: str, new_text: str) -> str:
    """
    Find and replace ALL occurrences in the active LibreOffice document.
    Primary use: replace placeholder text with scannable cite markers.

    Args:
        old_text: Text to find
        new_text: Replacement text (may include scannable cite markers like {  | Author, (Year) |  |  |zu:0:KEY})
    """
    return json.dumps(_libre_tool("find_and_replace_all_live", {"old": old_text, "new": new_text}), indent=2)


@mcp.tool
def libre_find(query: str) -> str:
    """Find text in the active LibreOffice document and return matches with positions."""
    return json.dumps(_libre_tool("find_text_live", {"query": query}), indent=2)


@mcp.tool
def libre_insert_text(content: str, paragraph: Optional[int] = None) -> str:
    """
    Insert text at the cursor position (or navigate to a paragraph first).

    Args:
        content: Text to insert (may include scannable cite markers)
        paragraph: Navigate to this paragraph number first (None = current cursor)
    """
    if paragraph is not None:
        _libre_tool("goto_paragraph_live", {"n": paragraph})
    return json.dumps(_libre_tool("insert_text_live", {"text": content}), indent=2)


@mcp.tool
def libre_paragraph(n: int) -> str:
    """Get the text of paragraph N from the active document (1-based)."""
    return json.dumps(_libre_tool("get_paragraph_live", {"n": n}), indent=2)


@mcp.tool
def libre_paragraph_count() -> str:
    """Get the total number of paragraphs in the active document."""
    return json.dumps(_libre_tool("get_paragraph_count_live"), indent=2)


@mcp.tool
def libre_outline() -> str:
    """Get the document outline (headings and structure) of the active document."""
    return json.dumps(_libre_tool("get_document_outline_live"), indent=2)


@mcp.tool
def libre_save(file_path: str) -> str:
    """
    Save the active LibreOffice document.

    Args:
        file_path: Path on the VM's Linux filesystem (e.g. /home/vboxuser/Documents/paper.odt)
                   Use shared folder path if you need access from Windows.
    """
    return json.dumps(_libre_tool("save_document_live", {"file_path": file_path}), indent=2)


# ---------------------------------------------------------------------------
# Document conversion (runs in WSL via pandoc)
# ---------------------------------------------------------------------------

@mcp.tool
def convert_document(input_path: str, output_format: str, output_path: Optional[str] = None) -> str:
    """
    Convert a document between formats using pandoc (runs in WSL).

    Formats: pdf, docx, odt, html, md/markdown, tex/latex, rst, epub, rtf, txt

    Args:
        input_path: Path to input file (WSL path, e.g. /mnt/c/Users/<username>/doc.odt)
        output_format: Target format (e.g. 'docx', 'odt', 'markdown')
        output_path: Output path (auto-generated from input if omitted)
    """
    input_path = os.path.expanduser(input_path)
    if not os.path.exists(input_path):
        return f"Error: file not found: {input_path}"

    if output_path is None:
        ext_map = {
            "markdown": "md", "latex": "tex",
            "docx": "docx", "odt": "odt", "html": "html",
            "pdf": "pdf", "epub": "epub", "rst": "rst", "txt": "txt",
        }
        ext = ext_map.get(output_format, output_format)
        output_path = f"{os.path.splitext(input_path)[0]}.{ext}"

    try:
        r = subprocess.run(
            ["pandoc", input_path, "-o", output_path],
            capture_output=True, text=True, timeout=120,
        )
        if r.returncode == 0:
            return json.dumps({"status": "ok", "output": output_path})
        return json.dumps({"status": "error", "stderr": r.stderr})
    except FileNotFoundError:
        return "Error: pandoc not found — install with: sudo apt install pandoc"
    except Exception as e:
        return f"Error: {e}"


@mcp.tool
def odf_scan_guide() -> str:
    """
    Show the steps to run the Zotero ODF Scan to convert scannable cite markers
    to live Zotero citations.

    NOTE: The ODF scan plugin runs INSIDE Zotero (on Windows), not in LibreOffice.
    It is a Zotero plugin that processes the ODT file directly using Zotero's internal
    citation formatting engine (CSL). It cannot be triggered via the Zotero HTTP API,
    so this step remains manual.
    """
    return json.dumps({
        "what_this_does": (
            "Converts {  | Author, (Year) |  |  |zu:0:KEY} plain-text markers "
            "into live LibreOffice/Zotero citation fields that Zotero can manage "
            "(update, reformat, generate bibliography from)."
        ),
        "where_it_runs": "Inside Zotero Desktop on Windows — NOT in LibreOffice",
        "steps": [
            "1. LibreOffice does NOT need to be open for this step",
            "2. In Zotero on Windows: Tools → ODF Scan",
            "3. 'Input file': select your ODT file with scannable cite markers",
            "4. 'Output file': choose where to save the converted file",
            "5. Click 'Scan' — Zotero reads the ODT, resolves each zu:0:KEY citation,",
            "   formats them via CSL, and writes a new ODT with live citation fields",
            "6. Open the output ODT in LibreOffice to verify citations",
            "7. To get a Word file: File → Save As → .docx in LibreOffice",
            "   OR: use convert_document('/path/to/output.odt', 'docx') from WSL via pandoc",
        ],
        "why_not_automated": (
            "The ODF scan runs Zotero's internal JavaScript CSL citation engine — "
            "it is not exposed via the Zotero local HTTP API (port 23119). "
            "Automation would require reimplementing CSL formatting in Python."
        ),
        "plugin": "https://github.com/Juris-M/zotero-odf-scan-plugin",
    }, indent=2)


if __name__ == "__main__":
    mcp.run()
