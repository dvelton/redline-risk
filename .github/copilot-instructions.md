# Redline Risk

This repository contains instructions and configuration for the Redline Risk Copilot CLI plugin.

## What this does

Redline Risk analyzes contract redlines with directional risk assessment. It extracts tracked changes from Word documents and produces a color-coded report showing which edits favor or hurt your position.

## Plugin structure

- `.github/copilot-instructions.md` - Repository-level Copilot instructions
- `skills/redline-risk/` - The Redline Risk skill
  - `SKILL.md` - Skill instructions for Copilot
  - `tools/redline_risk.py` - Core Python tool
- `setup.sh` / `setup.ps1` - Setup scripts for macOS/Linux and Windows
- `requirements.txt` - Python dependencies
