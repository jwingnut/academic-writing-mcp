# Academic Writing MCP

Unified MCP server for dissertation/academic writing workflows. Combines Zotero citation lookup, LibreOffice document editing, and document format conversion into a single Claude Code tool.

## Architecture

```
Claude Code (WSL2)
    │
    └── writing MCP server (this repo)
            │
            ├── zotero_* tools ──► Windows portproxy (172.28.32.1:23119)
            │                          └──► Zotero Desktop (Windows localhost:23119)
            │
            ├── libre_* tools  ──► Windows portproxy (172.28.32.1:8765)
            │                          └──► VirtualBox NAT (Windows localhost:8765)
            │                                   └──► LibreOffice MCP extension (VM:8765)
            │
            └── convert_document ──► pandoc (WSL)
```

## Workflow: Zotero Scannable Cite → Word Document

1. **Search Zotero** → get scannable cite markers  
   `{  | Author, (Year) |  |  |zu:0:ITEMKEY}`
2. **Edit document in LibreOffice** (in VirtualBox VM) → insert markers
3. **ODF Scan** in Zotero: Tools → ODF Scan → select the ODT file  
   → markers become live Zotero citations
4. **Export** as `.docx` for Word compatibility  
   (or use `convert_document` tool via pandoc)

## Setup

### Prerequisites

| Component | Where it runs | Purpose |
|-----------|--------------|---------|
| Zotero Desktop | Windows | Citation library (local API port 23119) |
| LibreOffice Writer | VirtualBox Ubuntu VM | Document editing |
| LibreOffice MCP extension | Same VM | HTTP bridge (port 8765) |
| Zotero ODF Scan plugin | LibreOffice in VM | Converts scannable cites to live citations |
| This MCP server | WSL2 | Claude Code integration |

### Step 1 — Enable Zotero local API

In Zotero on Windows:  
`Edit → Preferences → Advanced → check "Allow other applications on this computer to communicate with Zotero"`

### Step 2 — Install LibreOffice MCP extension (in VM, one-time)

```bash
# Inside VirtualBox VM
bash setup-vboxvm.sh
```

Then in LibreOffice: `Tools → MCP Server → Start MCP Server`

### Step 3 — Windows port forwarding (run after each Windows reboot)

```powershell
# In PowerShell as Administrator
.\setup-windows.ps1
```

This sets up:
- `netsh portproxy`: WSL gateway → Windows localhost for ports 23119 (Zotero) and 8765 (LibreOffice)
- `netsh advfirewall`: inbound rules for both ports
- `VBoxManage natpf1`: VirtualBox NAT port forwarding for LibreOffice

**Note:** portproxy rules are lost on reboot — re-run `setup-windows.ps1` each time.

### Step 4 — WSL setup (one-time)

```bash
# In WSL2
bash setup-wsl.sh
```

This creates the Python venv and writes `~/.claude/settings.json`.

### Step 5 — Restart Claude Code

The `writing` MCP server loads on Claude Code startup. Verify with:
```
libre_check_connection()
zotero_check_connection()
```

## Tools

### Zotero tools

| Tool | Description |
|------|-------------|
| `zotero_check_connection()` | Test Zotero API connectivity |
| `zotero_search(query, limit)` | Search library, returns items with `scannable_cite` field |
| `zotero_get_cite(key)` | Get `{  \| Author, (Year) \|  \|  \|zu:0:KEY}` for one item |
| `zotero_get_cites_batch(keys)` | Bulk cite lookup |
| `zotero_collections()` | List all collections |

### LibreOffice tools

| Tool | Description |
|------|-------------|
| `libre_check_connection()` | Test LibreOffice extension connectivity |
| `libre_document_info()` | Filename, word count, modification status |
| `libre_list_documents()` | List all open documents |
| `libre_content()` | Full text of active document |
| `libre_outline()` | Document headings and structure |
| `libre_paragraph(n)` | Get text of paragraph N |
| `libre_paragraph_count()` | Total paragraph count |
| `libre_find(query)` | Find text, return matches with positions |
| `libre_search_replace(old, new)` | Replace all occurrences (for inserting cite markers) |
| `libre_insert_text(content, paragraph)` | Insert text at cursor or paragraph N |
| `libre_save(file_path)` | Save active document |

### Conversion tools

| Tool | Description |
|------|-------------|
| `convert_document(input, format)` | pandoc-based format conversion (odt↔docx↔md↔html etc.) |
| `odf_scan_guide()` | Step-by-step guide for ODF scan workflow |

## Configuration

`~/.claude/settings.json` entry written by `setup-wsl.sh`:

```json
{
  "mcpServers": {
    "writing": {
      "command": "/home/<user>/github/academic-writing-mcp/.venv/bin/python",
      "args": ["/home/<user>/github/academic-writing-mcp/writing_mcp.py"],
      "env": {
        "LIBREOFFICE_URL": "http://172.28.32.1:8765",
        "FASTMCP_SHOW_SERVER_BANNER": "false"
      }
    },
    "reference-mcp": {
      "command": "uvx",
      "args": ["reference-mcp"]
    }
  }
}
```

### Environment variables (optional overrides)

| Variable | Default | Purpose |
|----------|---------|---------|
| `LIBREOFFICE_URL` | `http://<win-ip>:8765` | LibreOffice extension URL |
| `ZOTERO_HOST` | auto-detected WSL gateway | IP of machine running Zotero |
| `ZOTERO_API_KEY` | — | Use Zotero Web API instead of local API |
| `ZOTERO_LIBRARY_ID` | — | Required with `ZOTERO_API_KEY` |

## Key Technical Details

### Why `Host: localhost:23119` is required for Zotero

Zotero's local HTTP server (`httpd.js`) validates the `Host` header for CSRF protection. Requests must include the port — `Host: localhost:23119` — not just `Host: localhost`. This is handled automatically in the MCP server.

### Why WSL cannot reach Windows localhost directly

WSL2 uses a virtual network with a separate subnet (172.28.x.x). Windows services listening on `127.0.0.1` are not reachable from WSL at `localhost`. The `localhostforwarding=true` `.wslconfig` setting only forwards WSL→Windows for certain connection types. The workaround is `netsh portproxy` to bridge the WSL gateway IP to Windows localhost.

### Why LibreOffice runs in VirtualBox, not WSL

The LibreOffice MCP extension requires a desktop environment with a running LibreOffice instance. WSL2 does not have a native display server suitable for this. VirtualBox provides a full Ubuntu desktop where LibreOffice can run normally.

## Recreating on a New Machine

1. Clone this repo to `~/github/academic-writing-mcp/`
2. Clone `libreoffice-mcp-ubuntu` to `~/github/libreoffice-mcp-ubuntu/` (for the `.oxt` extension file)
3. Follow Steps 1–5 above
4. Update `LIBREOFFICE_URL` in `~/.claude/settings.json` if WSL gateway IP differs
