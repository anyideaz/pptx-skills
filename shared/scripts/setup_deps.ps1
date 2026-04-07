# Dependency installer for slide-deck AI skills (Windows PowerShell)
$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=== Slide Deck Skills: Dependency Setup ===" -ForegroundColor Cyan

# Check Python
$Python = $null
foreach ($cmd in @("python", "python3")) {
    if (Get-Command $cmd -ErrorAction SilentlyContinue) {
        $Python = $cmd
        break
    }
}
if (-not $Python) {
    Write-Host "ERROR: Python not found. Install Python 3.10+ from https://python.org" -ForegroundColor Red
    exit 1
}
Write-Host "✓ Python: $(& $Python --version)"

# Install Python packages
Write-Host "Installing Python packages (python-pptx, Pillow, markitdown)..."
& $Python -m pip install --user "python-pptx>=1.0.0" "Pillow>=10.0.0" "markitdown[all]>=0.1.0" --quiet
try {
    & $Python -c "import pptx" 2>$null
} catch {
    Write-Host "ERROR: python-pptx installation failed." -ForegroundColor Red
    exit 3
}
Write-Host "✓ python-pptx installed"
try {
    & $Python -c "import markitdown" 2>$null
} catch {
    Write-Host "ERROR: markitdown installation failed." -ForegroundColor Red
    exit 3
}
Write-Host "✓ markitdown installed"

# Check Node.js
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR: Node.js not found. Install Node.js 18+ from https://nodejs.org" -ForegroundColor Red
    exit 2
}
Write-Host "✓ Node.js: $(node --version)"

# Install npm packages
Write-Host "Installing npm packages (pptxgenjs, uuid)..."
Push-Location $ScriptDir
npm install --quiet
try {
    node -e "require('pptxgenjs')" 2>$null
} catch {
    Pop-Location
    Write-Host "ERROR: pptxgenjs installation failed." -ForegroundColor Red
    exit 4
}
Pop-Location
Write-Host "✓ pptxgenjs installed"

Write-Host ""
Write-Host "=== Setup complete. All dependencies ready. ===" -ForegroundColor Green
