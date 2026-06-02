# Academic Writing MCP — Windows Setup Script
# Run this in PowerShell as Administrator
# Sets up port forwarding so WSL2 can reach Zotero (Windows) and LibreOffice (VirtualBox VM)
#
# Prerequisites before running:
#   1. Zotero Desktop is installed and running
#      In Zotero: Edit -> Preferences -> Advanced
#      -> check "Allow other applications on this computer to communicate with Zotero"
#   2. VirtualBox is installed with a Linux VM that has LibreOffice + the MCP extension
#   3. WSL2 is installed and running

param(
    [string]$VBoxManagePath = "P:\VirtualBox\VBoxManage.exe",
    [string]$VMName = "Ubuntu 24.04",
    [string]$WslGatewayIP = ""  # auto-detected if empty
)

# Detect WSL2 gateway IP automatically
if (-not $WslGatewayIP) {
    $WslGatewayIP = (Get-NetIPAddress -InterfaceAlias "vEthernet (WSL)" -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress
    if (-not $WslGatewayIP) {
        # Fallback: check wsl route
        $WslGatewayIP = (wsl ip route show default 2>$null | Select-String "via (\S+)" | ForEach-Object { $_.Matches.Groups[1].Value } | Select-Object -First 1)
    }
    if (-not $WslGatewayIP) {
        Write-Error "Could not auto-detect WSL gateway IP. Pass -WslGatewayIP explicitly."
        exit 1
    }
}

Write-Host "WSL Gateway IP: $WslGatewayIP"
Write-Host "VM Name: $VMName"
Write-Host ""

# ---------------------------------------------------------------------------
# Step 1: VirtualBox NAT port forwarding for LibreOffice (port 8765)
# ---------------------------------------------------------------------------
Write-Host "[1/4] Setting up VirtualBox NAT port forward for LibreOffice (8765)..."
if (Test-Path $VBoxManagePath) {
    & $VBoxManagePath controlvm $VMName natpf1 "libreoffice,tcp,,8765,,8765" 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  OK: VirtualBox NAT port forward added"
    } else {
        Write-Host "  WARN: VBoxManage failed (VM may not be running, or rule already exists)"
    }
} else {
    Write-Host "  SKIP: VBoxManage not found at $VBoxManagePath"
    Write-Host "        Run manually: VBoxManage controlvm '$VMName' natpf1 'libreoffice,tcp,,8765,,8765'"
}

# ---------------------------------------------------------------------------
# Step 2: netsh portproxy — WSL gateway -> Windows localhost
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[2/4] Setting up netsh portproxy rules..."

# Remove any existing rules first (idempotent)
netsh interface portproxy delete v4tov4 listenport=23119 listenaddress=$WslGatewayIP 2>$null | Out-Null
netsh interface portproxy delete v4tov4 listenport=8765  listenaddress=$WslGatewayIP 2>$null | Out-Null

# Zotero local API
netsh interface portproxy add v4tov4 `
    listenport=23119 listenaddress=$WslGatewayIP `
    connectport=23119 connectaddress=127.0.0.1
Write-Host "  OK: Zotero portproxy ($WslGatewayIP`:23119 -> 127.0.0.1:23119)"

# LibreOffice MCP extension
netsh interface portproxy add v4tov4 `
    listenport=8765 listenaddress=$WslGatewayIP `
    connectport=8765 connectaddress=127.0.0.1
Write-Host "  OK: LibreOffice portproxy ($WslGatewayIP`:8765 -> 127.0.0.1:8765)"

# ---------------------------------------------------------------------------
# Step 3: Windows Firewall inbound rules
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[3/4] Adding Windows Firewall inbound rules..."

# Remove old rules if they exist
Remove-NetFirewallRule -DisplayName "Zotero WSL Bridge"    -ErrorAction SilentlyContinue
Remove-NetFirewallRule -DisplayName "LibreOffice WSL Bridge" -ErrorAction SilentlyContinue

New-NetFirewallRule -DisplayName "Zotero WSL Bridge" `
    -Direction Inbound -Action Allow -Protocol TCP -LocalPort 23119 | Out-Null
Write-Host "  OK: Firewall rule for Zotero (TCP 23119)"

New-NetFirewallRule -DisplayName "LibreOffice WSL Bridge" `
    -Direction Inbound -Action Allow -Protocol TCP -LocalPort 8765 | Out-Null
Write-Host "  OK: Firewall rule for LibreOffice (TCP 8765)"

# ---------------------------------------------------------------------------
# Step 4: Verify
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[4/4] Verifying portproxy rules..."
netsh interface portproxy show all

Write-Host ""
Write-Host "========================================"
Write-Host "Setup complete. Verify from WSL with:"
Write-Host "  curl -H 'Host: localhost:23119' http://$WslGatewayIP`:23119/api/users/0/items?limit=1"
Write-Host "  curl http://$WslGatewayIP`:8765/health"
Write-Host ""
Write-Host "NOTE: These portproxy rules are reset on Windows reboot."
Write-Host "To persist, schedule this script via Task Scheduler on startup,"
Write-Host "or add to your WSL startup routine."
Write-Host "========================================"
