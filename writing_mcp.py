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
from typing import Optional
import httpx
from fastmcp import FastMCP

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
        input_path: Path to input file (WSL path, e.g. /mnt/c/Users/jayw/doc.odt)
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
    Show the steps to run the Zotero ODF Scan in LibreOffice to convert
    scannable cite markers to live Zotero citations.
    """
    return json.dumps({
        "what_this_does": "Converts {  | Author, (Year) |  |  |zu:0:KEY} markers to live Zotero citations",
        "steps": [
            "1. Make sure the ODT file is open in LibreOffice Writer (in the VirtualBox VM)",
            "2. In LibreOffice: Tools → Macros → Organize Basic Macros → find 'ODF Scan'",
            "   OR use the Zotero toolbar → 'Switch Word Processors' → ODF Scan",
            "3. Click 'Scan Document' — Zotero scans and replaces all markers",
            "4. Save the file as .odt (keeps live citations)",
            "5. To get a Word file: File → Save As → .docx",
            "   OR from WSL: use the convert_document tool (odt → docx via pandoc)",
        ],
        "note": "ODF scan plugin: https://github.com/Juris-M/zotero-odf-scan-plugin (already at ~/github/zotero-odf-scan-plugin)",
    }, indent=2)


if __name__ == "__main__":
    mcp.run()
