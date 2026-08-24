@echo off
echo ========================================================
echo Updating PowerPoint MCP Installation for Antigravity CLI
echo ========================================================

echo 1. Installing in editable mode...
call uv pip install -e .

echo 2. Syncing MCP schemas and skill...
call uv run python scripts/sync_mcp.py

echo 3. Running test suite...
call uv run pytest

echo.
echo ========================================================
echo Update Complete!
echo ========================================================
