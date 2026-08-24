@echo off
start "" /min powershell.exe -NoProfile -WindowStyle Minimized -Command "cd 'C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP'; uv run python -m powerpoint_mcp.server"