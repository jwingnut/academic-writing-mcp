#!/bin/bash
# Academic Writing MCP — WSL Setup Script
# Run this in WSL2 (Ubuntu) as your normal user
#
# What this does:
#   1. Creates the Python venv with required packages
#   2. Writes the MCP server config to ~/.claude/settings.json
#   3. Verifies connectivity to Zotero and LibreOffice

set -e

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$REPO_DIR/.venv"

echo "========================================"
echo "Academic Writing MCP — WSL Setup"
echo "========================================"
echo "Repo: $REPO_DIR"
echo ""

# ---------------------------------------------------------------------------
# Step 1: Python venv
# ---------------------------------------------------------------------------
echo "[1/4] Setting up Python venv..."
if ! command -v uv &>/dev/null; then
    echo "  Installing uv..."
    curl -LsSf https://astral.sh/uv/install.sh | sh
    source "$HOME/.local/bin/env"
fi

cd "$REPO_DIR"
uv venv 2>&1 | tail -2
uv pip install fastmcp httpx pyzotero 2>&1 | tail -3
echo "  OK: venv at $VENV"

# ---------------------------------------------------------------------------
# Step 2: Detect Windows host IP
# ---------------------------------------------------------------------------
echo ""
echo "[2/4] Detecting network..."
WIN_IP=$(ip route show default 2>/dev/null | awk '{print $3}' | head -1)
echo "  Windows host IP: $WIN_IP"

# Test Zotero
ZOTERO_STATUS=$(curl -s --connect-timeout 3 \
    -H "Host: localhost:23119" \
    "http://$WIN_IP:23119/api/users/0/items?limit=1" \
    -o /dev/null -w "%{http_code}" 2>/dev/null)
if [ "$ZOTERO_STATUS" = "200" ]; then
    ZOTERO_COUNT=$(curl -s -H "Host: localhost:23119" \
        "http://$WIN_IP:23119/api/users/0/items?limit=1" 2>/dev/null | \
        python3 -c "import sys,json; _ = json.load(sys.stdin)" 2>/dev/null && \
        curl -s -H "Host: localhost:23119" \
        "http://$WIN_IP:23119/api/users/0/items" 2>/dev/null | \
        python3 -c "import sys; import urllib.request; \
            req = urllib.request.Request('http://$WIN_IP:23119/api/users/0/items?limit=0', \
            headers={'Host':'localhost:23119'}); print('?')" 2>/dev/null || echo "?")
    echo "  Zotero: CONNECTED (HTTP $ZOTERO_STATUS)"
else
    echo "  Zotero: NOT REACHABLE (HTTP $ZOTERO_STATUS)"
    echo "    -> Run setup-windows.ps1 as admin on Windows first"
fi

# Test LibreOffice
LIBRE_STATUS=$(curl -s --connect-timeout 3 \
    "http://$WIN_IP:8765/health" -o /dev/null -w "%{http_code}" 2>/dev/null)
if [ "$LIBRE_STATUS" = "200" ]; then
    echo "  LibreOffice: CONNECTED (HTTP $LIBRE_STATUS)"
else
    echo "  LibreOffice: NOT REACHABLE (HTTP $LIBRE_STATUS)"
    echo "    -> Run setup-windows.ps1 and start MCP in LibreOffice first"
fi

# ---------------------------------------------------------------------------
# Step 3: Write Claude Code settings
# ---------------------------------------------------------------------------
echo ""
echo "[3/4] Configuring Claude Code MCP settings..."

SETTINGS="$HOME/.claude/settings.json"
mkdir -p "$HOME/.claude"

# Check if settings.json already has mcpServers
if [ -f "$SETTINGS" ] && python3 -c "import json; d=json.load(open('$SETTINGS')); d['mcpServers']" 2>/dev/null; then
    echo "  SKIP: mcpServers already configured in $SETTINGS"
    echo "        To reset, remove the mcpServers key and re-run"
else
    # Build MCP config — preserve existing settings if file exists
    python3 - <<PYEOF
import json, os

settings_path = "$SETTINGS"
repo_dir = "$REPO_DIR"
win_ip = "$WIN_IP"
venv = "$VENV"

# Load existing settings
if os.path.exists(settings_path):
    with open(settings_path) as f:
        settings = json.load(f)
else:
    settings = {}

settings["mcpServers"] = {
    "writing": {
        "command": f"{venv}/bin/python",
        "args": [f"{repo_dir}/writing_mcp.py"],
        "env": {
            "LIBREOFFICE_URL": f"http://{win_ip}:8765"
        }
    },
    "reference-mcp": {
        "command": "uvx",
        "args": ["reference-mcp"]
    }
}

with open(settings_path, "w") as f:
    json.dump(settings, f, indent=2)
    f.write("\n")

print(f"  OK: Written to {settings_path}")
PYEOF
fi

# ---------------------------------------------------------------------------
# Step 4: Smoke test
# ---------------------------------------------------------------------------
echo ""
echo "[4/4] Running smoke test..."
RESULT=$("$VENV/bin/python" -c "
import os
os.environ['LIBREOFFICE_URL'] = 'http://$WIN_IP:8765'
import writing_mcp, json

z = writing_mcp.zotero_check_connection()
l = writing_mcp.libre_check_connection()
print('Zotero:', json.loads(z)['status'])
print('LibreOffice:', json.loads(l)['status'])
" 2>&1)
echo "  $RESULT"

echo ""
echo "========================================"
echo "WSL setup complete!"
echo ""
echo "Restart Claude Code to load the new MCP servers."
echo "Then verify with: zotero_check_connection() and libre_check_connection()"
echo ""
echo "NOTE: Windows portproxy rules reset on reboot."
echo "Re-run setup-windows.ps1 after each Windows restart."
echo "========================================"
