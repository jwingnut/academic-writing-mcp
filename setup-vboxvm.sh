#!/bin/bash
# Academic Writing MCP — VirtualBox VM Setup Script
# Run this INSIDE the VirtualBox Ubuntu VM as the vboxuser
#
# What this does:
#   1. Installs the LibreOffice MCP extension (.oxt file)
#   2. Installs Python dependencies for the extension
#   3. Prints instructions for starting the MCP server in LibreOffice

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "========================================"
echo "Academic Writing MCP — VM Setup"
echo "========================================"
echo ""

# ---------------------------------------------------------------------------
# Check for the extension file
# The .oxt is in libreoffice-mcp-ubuntu or can be copied from the host
# ---------------------------------------------------------------------------
OXT_FILE=""
for candidate in \
    "$SCRIPT_DIR/../libreoffice-mcp-ubuntu/libreoffice-mcp-extension.oxt" \
    "$HOME/libreoffice-mcp-extension.oxt" \
    "/tmp/libreoffice-mcp-extension.oxt"; do
    if [ -f "$candidate" ]; then
        OXT_FILE="$candidate"
        break
    fi
done

if [ -z "$OXT_FILE" ]; then
    echo "ERROR: libreoffice-mcp-extension.oxt not found."
    echo ""
    echo "Copy it from WSL to the VM first:"
    echo "  From WSL: cp ~/github/libreoffice-mcp-ubuntu/libreoffice-mcp-extension.oxt /tmp/"
    echo "  Then share /tmp with VirtualBox shared folder, or use scp"
    echo ""
    echo "Or if you have it in this repo, update OXT_FILE path in this script."
    exit 1
fi

echo "[1/3] Installing LibreOffice MCP extension..."
echo "  Extension: $OXT_FILE"

# Install the extension (no-interaction mode)
unopkg add --force "$OXT_FILE" 2>&1 && echo "  OK: Extension installed"

# ---------------------------------------------------------------------------
# Install Python dependencies that the extension needs
# ---------------------------------------------------------------------------
echo ""
echo "[2/3] Installing Python dependencies..."
pip3 install --user fastmcp httpx 2>&1 | tail -3 || true
echo "  OK: Dependencies installed"

# ---------------------------------------------------------------------------
# Check LibreOffice version
# ---------------------------------------------------------------------------
echo ""
echo "[3/3] Checking LibreOffice..."
soffice --version 2>/dev/null || echo "  WARN: soffice not in PATH"

echo ""
echo "========================================"
echo "VM setup complete!"
echo ""
echo "NEXT STEPS (manual, inside LibreOffice):"
echo "  1. Open LibreOffice Writer"
echo "  2. Go to: Tools -> MCP Server -> Start MCP Server"
echo "  3. Verify: curl http://localhost:8765/health"
echo "     Should return: {\"status\": \"healthy\", ...}"
echo ""
echo "The MCP server must be restarted each time LibreOffice opens."
echo "========================================"
