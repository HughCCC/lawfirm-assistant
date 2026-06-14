"""Wrapper around the existing generate_docx.py script."""

import os
import re
import sys
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

# Add scripts directory to path for import
SCRIPTS_DIR = (Path(__file__).resolve().parent.parent / ".law-research" / "scripts")
sys.path.insert(0, str(SCRIPTS_DIR))

from generate_docx import generate_docx as gen_docx

from config import WEBAPP_OUTPUT_DIR, MAX_OUTPUT_FILES, COVER_TEMPLATE_PATH


def run_generate_docx(
    json_path: str,
    title: Optional[str] = None,
) -> tuple[bool, str]:
    """Generate a .docx from a report JSON file.

    Args:
        json_path: Path to the report JSON file.
        title: Optional report title for the output filename.

    Returns:
        (success, output_path_or_error)
    """
    if not os.path.exists(json_path):
        return False, f"JSON 文件不存在：{json_path}"

    # Generate output filename
    report_id = os.path.basename(json_path).replace("report_", "").replace(".json", "")
    date_fname = datetime.now().strftime('%Y%m%d')
    safe_title = "法律评估意见书"
    if title:
        # Sanitize title for filename (remove path-illegal chars)
        safe_title = re.sub(r'[\\/:*?"<>|]', '_', title)[:80]

    output_path = WEBAPP_OUTPUT_DIR / f"{safe_title}_{date_fname}.docx"

    # Update JSON title if provided (cover page uses this as the report title)
    if title:
        _update_json_title(json_path, title)

    try:
        # The generate_docx function runs its own scan internally and returns False on violations.
        # Since we already scanned in generator.py, this should always pass — but it's our safety net.
        cover_path = str(COVER_TEMPLATE_PATH) if COVER_TEMPLATE_PATH.exists() else None
        success = gen_docx(str(json_path), str(output_path), cover_path)

        if success:
            # Cleanup old files
            _cleanup_old_files()
            return True, str(output_path)
        else:
            return False, "违禁字符扫描未通过（generate_docx 内部扫描）"

    except Exception as e:
        return False, f"文档生成异常：{str(e)}"


def _update_json_title(json_path: str, new_title: str):
    """Update the title field in the report JSON file."""
    import json
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        data['title'] = new_title
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # Non-critical: document will still generate with original title


def _cleanup_old_files():
    """Remove oldest output files if exceeding MAX_OUTPUT_FILES."""
    files = sorted(
        WEBAPP_OUTPUT_DIR.glob("*"),
        key=lambda f: f.stat().st_mtime,
    )
    if len(files) > MAX_OUTPUT_FILES:
        for f in files[:len(files) - MAX_OUTPUT_FILES]:
            try:
                f.unlink()
            except OSError:
                pass
