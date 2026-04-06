# Redline Risk setup script for Windows

Write-Host "Installing Redline Risk dependencies..." -ForegroundColor Cyan

# Check if pip is available
if (-not (Get-Command pip -ErrorAction SilentlyContinue)) {
    Write-Host "Error: pip not found. Please install Python 3 first." -ForegroundColor Red
    exit 1
}

# Install dependencies
pip install -r requirements.txt

Write-Host ""
Write-Host "✓ All dependencies installed successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "To verify the installation, run:"
Write-Host "  python skills/redline-risk/tools/redline_risk.py setup-check"
