"""Sync PowerPoint MCP server schemas and skill to Antigravity CLI."""

import asyncio
import json
import os
from pathlib import Path
import shutil
import sys

from powerpoint_mcp.server import app


def sync_mcp_schemas():
    """Extract tool schemas from MCP server and write them to Antigravity MCP directory."""
    user_home = Path.home()
    mcp_target_dir = user_home / ".gemini" / "antigravity-cli" / "mcp" / "powerpoint-mcp"
    mcp_target_dir.mkdir(parents=True, exist_ok=True)

    tools = asyncio.run(app.list_tools())
    print(f"[+] Found {len(tools)} MCP tools in powerpoint_mcp server.")

    # Remove obsolete json files
    tool_names = {t.name for t in tools}
    for existing_file in mcp_target_dir.glob("*.json"):
        if existing_file.stem not in tool_names:
            existing_file.unlink()
            print(f"[-] Removed obsolete schema: {existing_file.name}")

    # Write updated schema for each tool
    for t in tools:
        schema = {
            "name": t.name,
            "description": t.description or "",
            "parameters": t.input_schema if hasattr(t, "input_schema") and t.input_schema else getattr(t, "parameters", {}),
        }
        dest_path = mcp_target_dir / f"{t.name}.json"
        dest_path.write_text(json.dumps(schema, indent=2), encoding="utf-8")
        print(f"[OK] Updated {dest_path.name}")

    print(f"\n[SUCCESS] Synced {len(tools)} MCP schemas to {mcp_target_dir}")


def sync_skill():
    """Sync powerpoint-editor skill to user skills directory if present."""
    repo_root = Path(__file__).resolve().parent.parent
    source_skill = repo_root / ".agents" / "skills" / "powerpoint-editor"

    user_home = Path.home()
    target_skill_builtin = user_home / ".gemini" / "antigravity-cli" / "builtin" / "skills" / "powerpoint-editor"

    if source_skill.exists():
        if target_skill_builtin.parent.exists():
            target_skill_builtin.mkdir(parents=True, exist_ok=True)
            for item in source_skill.glob("*"):
                if item.is_file():
                    shutil.copy2(item, target_skill_builtin / item.name)
            print(f"[OK] Synced skill to {target_skill_builtin}")



if __name__ == "__main__":
    sync_mcp_schemas()
    sync_skill()
