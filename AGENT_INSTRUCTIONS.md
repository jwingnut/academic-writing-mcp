# Writing Tools — Agent Instructions

Read this file at the start of any writing session. It describes every tool available,
the citation workflow, and conventions to follow.

---

## 1. Available MCP Servers

### `writing` — core tool for this project

Unified server combining Zotero, LibreOffice, and document conversion.
**Always verify connectivity at the start of a session:**

```
zotero_check_connection()    → should return CONNECTED, shows library item count
libre_check_connection()     → should return CONNECTED, confirms LibreOffice is open in VM
betterbibtex_check_connection() → optional; should return CONNECTED when BBT is needed
```

If either fails, see the Troubleshooting section below.

### `reference-mcp`
Lookup citations by DOI, CrossRef, Google Scholar. Use when a paper isn't in Zotero yet.

### `office-word-mcp`
Direct editing of `.docx` files on the Windows filesystem. Use for Word documents that
don't go through the Zotero ODF scan workflow (e.g. quick edits, formatting).

The `writing` MCP also has direct DOCX/Zotero repair tools. Use these when a
Word document already has Zotero fields but some citations are stale plain text
that Word/Zotero will not renumber.

### `arXivPaper`
Search and fetch arXiv papers. Useful for finding recent preprints.

---

## 2. The Citation Workflow (full end-to-end)

This is the standard workflow for writing academic text with live Zotero citations.

```
Step 1 [AUTOMATED]  Search Zotero → get scannable cite markers
Step 2 [AUTOMATED]  Write/edit the ODT document in LibreOffice (via writing MCP)
Step 3 [MANUAL]     ODF Scan in Zotero on Windows → converts markers to live citations
Step 4 [AUTOMATED]  Export ODT → DOCX via pandoc (convert_document tool)
                    OR open in LibreOffice on Windows → Save As .docx
Step 5 [MANUAL]     Open in Word → use Zotero plugin to refresh/format bibliography
```

### Step 1 — Search Zotero and get cite markers

```python
zotero_search("wildfire landscape connectivity", limit=5)
# Returns a list. Each item has:
# {
#   "key": "ABCD1234",
#   "title": "...",
#   "authors": "Smith, Jones",
#   "year": "2023",
#   "scannable_cite": "{  | Smith & Jones, (2023) |  |  |zu:0:ABCD1234}"
# }

# For a specific key you already know:
zotero_get_cite("ABCD1234")

# Bulk lookup:
zotero_get_cites_batch(["KEY1", "KEY2", "KEY3"])
```

### Scannable cite format

```
{  | Author et al., (Year) |  |  |zu:0:ITEMKEY}
```

Rules:
- Two spaces after `{` and before `|`
- Year in parentheses: `(2023)` not `2023`
- `zu:0:` prefix before the Zotero item key
- For 1 author: `Smith,`  |  2 authors: `Smith & Jones,`  |  3+: `Smith et al.,`

### Step 2 — Edit in LibreOffice

```python
# See what's open
libre_list_documents()
libre_document_info()

# Read content
libre_content()            # full text
libre_outline()            # headings only
libre_paragraph(n)         # paragraph N (1-based)
libre_paragraph_count()    # total paragraphs

# Write
libre_insert_text("Some text with {  | Smith, (2023) |  |  |zu:0:KEY} inline.")
libre_insert_text("New paragraph text", paragraph=15)  # navigate first

# Find and replace (most useful for inserting cites into existing text)
libre_search_replace("(Smith 2023)", "{  | Smith, (2023) |  |  |zu:0:KEY}")

# Save (path is on the VirtualBox VM filesystem)
libre_save("/home/vboxuser/Documents/my_paper.odt")
```

### Step 3 — ODF Scan (MANUAL, in Zotero on Windows)

1. In Zotero: **Tools → ODF Scan**
2. Input: the `.odt` file saved in the VM  
   (access via VirtualBox shared folder, or copy to a shared path)
3. Output: choose a save location
4. Click **Scan** — Zotero replaces `{  | ... |  |  |zu:0:KEY}` markers with live citation fields
5. The output ODT now has proper Zotero citations that Word/LibreOffice can manage

### Step 4 — Convert to DOCX

```python
# Via pandoc in WSL (fast, loses live citation fields — use for draft sharing)
convert_document("/mnt/c/Users/<username>/Documents/paper.odt", "docx")

# For a Word file WITH live Zotero citations:
# Open the scanned ODT in LibreOffice → File → Save As → .docx
# (preserves the citation XML that the Zotero Word plugin can read)
```

---

## 3. Direct DOCX Zotero Field Repair

Use this path for existing `.docx` manuscripts where citations are already mostly
live Zotero fields, but some labels are plain text, often in figure/table text.

This does **not** replace Zotero. It inserts Word fields that Zotero can refresh.
Final numbering and bibliography updates still happen in Word through the Zotero
plugin.

### Audit

```python
docx_zotero_audit(
    "/mnt/c/Users/<username>/Documents/paper.docx",
    suspect_terms=["Chen & Zhan", "Sbayti", "Gan"]
)
```

### Verify Zotero and Better BibTeX metadata

```python
zotero_get_cites_batch(["WIVFG3LY", "ZTSKQQU4", "BVF58E8R"])
betterbibtex_citation_keys(["WIVFG3LY", "ZTSKQQU4", "BVF58E8R"])
```

### Insert live Word/Zotero fields

```python
docx_zotero_insert_citations(
    docx_path="/mnt/c/Users/<username>/Documents/paper.docx",
    output_path="/mnt/c/Users/<username>/Documents/paper_zotero_fixed.docx",
    replacements=[
        {
            "old_text": "Chen & Zhan [78]",
            "zotero_keys": ["WIVFG3LY"],
            "display_text": "[Zotero citation]",
            "keep_prefix_text": True
        }
    ]
)
```

With `keep_prefix_text=True`, the author text (`Chen & Zhan`) remains normal text
and only the numeric bracket is converted to a live Zotero field. After opening
the output DOCX in Word, click **Zotero → Refresh** so Zotero renumbers citations
and updates the bibliography.

Use `fallback_paragraph_rewrite=True` only for simple figure/table labels where
the target text spans multiple Word runs. It rewrites the affected paragraph's
runs and is therefore more invasive.

---

## 4. File Path Conventions

| Location | Path | Notes |
|----------|------|-------|
| Writing folder (Windows) | `/mnt/c/Users/<username>/OneDrive/Documents/` | Your primary writing location |
| LibreOffice VM documents | `/home/vboxuser/Documents/` | VM internal path |
| Shared folder (VM ↔ Windows) | Set up in VirtualBox settings | If configured, use for file transfer |
| Writing MCP repo | `~/github/academic-writing-mcp/` | Tool source |

---

## 5. Creating a New Document

```python
# Create new Writer document in LibreOffice (VM)
writing_mcp._libre_tool('create_document_live', {'doc_type': 'writer'})

# Check it appeared
libre_list_documents()   # new doc will be at the end of the list

# Write content (inserts at cursor — beginning of new doc)
libre_insert_text("Title\n\n")
libre_insert_text("Body paragraph with citation {  | Smith, (2023) |  |  |zu:0:KEY}.\n\n")

# Save
libre_save("/home/vboxuser/Documents/new_paper.odt")
```

---

## 6. Troubleshooting

### Zotero not connected
```
zotero_check_connection()  → FAILED
```
**Fix:** Run `setup-windows.ps1` as admin on Windows (resets portproxy rules lost on reboot):
```powershell
# In PowerShell (admin) on Windows:
netsh interface portproxy add v4tov4 listenport=23119 listenaddress=172.28.32.1 connectport=23119 connectaddress=127.0.0.1
netsh advfirewall firewall add rule name="Zotero WSL Bridge" dir=in action=allow protocol=TCP localport=23119
```
Zotero Desktop must be running on Windows with local API enabled:
`Edit → Preferences → Advanced → "Allow other applications on this computer to communicate with Zotero"`

### LibreOffice not connected
```
libre_check_connection()  → FAILED
```
**Fix steps:**
1. Start the VirtualBox VM ("Ubuntu 24.04")
2. Open LibreOffice Writer in the VM
3. `Tools → MCP Server → Start MCP Server`
4. Then on Windows (PowerShell admin):
```powershell
P:\VirtualBox\VBoxManage.exe controlvm "Ubuntu 24.04" natpf1 "libreoffice,tcp,,8765,,8765"
netsh interface portproxy add v4tov4 listenport=8765 listenaddress=172.28.32.1 connectport=8765 connectaddress=127.0.0.1
netsh advfirewall firewall add rule name="LibreOffice WSL Bridge" dir=in action=allow protocol=TCP localport=8765
```

### Zotero items return "Unknown" author/year
The item's metadata is incomplete in the library. Look it up manually:
```python
zotero_search("author name title keywords", limit=10)
# Then use the key to check metadata
writing_mcp._zotero_get("/items/KEY")
```

---

## 7. Configuring This MCP in a New Claude Code Session

The `writing` server is configured globally at `~/.claude/settings.json`.
Any Claude Code session started in WSL2 on this machine already has the tools.

If starting on a **new machine**, run in WSL:
```bash
cd ~/github/academic-writing-mcp
bash setup-wsl.sh
```
Then run `setup-windows.ps1` on Windows. Full docs in `README.md`.

## 8. Configuring for Codex (Windows)

Add to `C:\Users\<username>\.codex\config.toml`:

```toml
[mcp_servers.writing]
command = "/home/<user>/github/academic-writing-mcp/.venv/bin/python"
args = ["/home/<user>/github/academic-writing-mcp/writing_mcp.py"]
startup_timeout_sec = 30.0

[mcp_servers.writing.env]
LIBREOFFICE_URL = "http://172.28.32.1:8765"
FASTMCP_SHOW_SERVER_BANNER = "false"
```

This assumes Codex runs inside WSL (the standard setup for codex-cli on WSL2).
The WSL gateway IP and portproxy setup must still be in place.

---

## 9. Quick-Start Prompt for a New Agent Session

Paste this at the start of a new Claude Code or Codex session to orient the agent:

```
Read /home/jay/github/academic-writing-mcp/AGENT_INSTRUCTIONS.md for the full
writing toolkit guide. Then:
1. Run zotero_check_connection() and libre_check_connection() to verify setup
2. Run libre_list_documents() to see what's open in LibreOffice
3. Ask me what writing task we're working on today
```
