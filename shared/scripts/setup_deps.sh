#!/usr/bin/env bash
# Dependency installer for slide-deck AI skills (Unix/macOS)
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "=== Slide Deck Skills: Dependency Setup ==="

# Check Python
if command -v python3 &>/dev/null; then
    PYTHON=python3
elif command -v python &>/dev/null; then
    PYTHON=python
else
    echo "ERROR: Python not found. Install Python 3.10+ from https://python.org" >&2
    exit 1
fi
echo "✓ Python: $($PYTHON --version)"

# Install Python packages
echo "Installing Python packages (python-pptx, Pillow, markitdown)..."
$PYTHON -m pip install --user "python-pptx>=1.0.0" "Pillow>=10.0.0" "markitdown[all]>=0.1.0" --quiet
if ! $PYTHON -c "import pptx" 2>/dev/null; then
    echo "ERROR: python-pptx installation failed." >&2
    exit 3
fi
echo "✓ python-pptx installed"
if ! $PYTHON -c "import markitdown" 2>/dev/null; then
    echo "ERROR: markitdown installation failed." >&2
    exit 3
fi
echo "✓ markitdown installed"

# Check Node.js
if ! command -v node &>/dev/null; then
    echo "ERROR: Node.js not found. Install Node.js 18+ from https://nodejs.org" >&2
    exit 2
fi
echo "✓ Node.js: $(node --version)"

# Install npm packages
echo "Installing npm packages (pptxgenjs, uuid)..."
cd "$SCRIPT_DIR"
npm install --quiet
if ! node -e "require('pptxgenjs')" 2>/dev/null; then
    echo "ERROR: pptxgenjs installation failed." >&2
    exit 4
fi
echo "✓ pptxgenjs installed"

echo ""
echo "=== Setup complete. All dependencies ready. ==="
