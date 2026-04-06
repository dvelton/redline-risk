#!/bin/bash

# Redline Risk setup script for macOS/Linux

set -e

echo "Installing Redline Risk dependencies..."

# Check if pip3 is available
if ! command -v pip3 &> /dev/null; then
    echo "Error: pip3 not found. Please install Python 3 first."
    exit 1
fi

# Install dependencies
pip3 install -r requirements.txt

echo ""
echo "✓ All dependencies installed successfully!"
echo ""
echo "To verify the installation, run:"
echo "  python3 skills/redline-risk/tools/redline_risk.py setup-check"
