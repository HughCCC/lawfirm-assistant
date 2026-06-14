"""FastAPI application: legal report generation web service."""

import logging
import sys
import uuid
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

# Add scripts directory for format_docx import
SCRIPTS_DIR = Path(__file__).parent.parent / ".law-research" / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from models import GenerateRequest, GenerateResponse
from generator import generate_report
from docx_runner import run_generate_docx

# ─── Logging Setup ────────────────────────────────────────────────────

# Filter out API keys from logs
class ApiKeyFilter(logging.Filter):
    def filter(self, record):
        msg = record.getMessage()
        if "api_key" in msg.lower():
            # Redact api_key values
            return False
        return True


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("legal-report")
logger.addFilter(ApiKeyFilter())

# ─── FastAPI App ──────────────────────────────────────────────────────

app = FastAPI(
    title="法律评估意见书生成系统",
    description="Law Firm Legal Assessment Report Generator — BIRACS framework",
    version="1.0.0",
)

# Mount static files
STATIC_DIR = Path(__file__).parent / "static"
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/")
async def root():
    """Serve the main page."""
    return FileResponse(str(STATIC_DIR / "index.html"))


@app.post("/api/generate", response_model=GenerateResponse)
async def generate(request: GenerateRequest):
    """Generate a legal assessment report.

    Accepts a legal question + LLM provider credentials, returns a
    download URL for the generated .docx file.
    """
    logger.info(
        "Generating report — provider=%s, model=%s, question_len=%d",
        request.provider,
        request.model or "default",
        len(request.question),
    )

    # Step 1: Generate the report JSON via LLM
    report_data, json_path, retries, error = generate_report(request)

    if error:
        logger.warning("Report generation failed after %d attempts: %s", retries + 1, error[:200])
        return GenerateResponse(
            success=False,
            error=error,
            retries_used=retries,
        )

    # Step 2: Generate .docx
    title = report_data.get("title", "法律评估意见书") if report_data else "法律评估意见书"
    docx_success, docx_result = run_generate_docx(json_path, title)

    if not docx_success:
        logger.error("Docx generation failed: %s", docx_result)
        return GenerateResponse(
            success=False,
            title=title,
            error=f"文档生成失败：{docx_result}",
            retries_used=retries,
        )

    # Success — append date to title for display
    date_str = datetime.now().strftime('%Y年%m月%d日')
    display_title = f"{title} — {date_str}"
    docx_filename = Path(docx_result).name
    logger.info("Report generated successfully: %s", docx_filename)

    return GenerateResponse(
        success=True,
        title=display_title,
        docx_filename=docx_filename,
        docx_url=f"/api/download/{docx_filename}",
        retries_used=retries,
    )


@app.post("/api/format")
async def format_document(file: UploadFile = File(...)):
    """格式调整校对：上传 .docx，自动统一为标准法律文书格式。

    Accepts a .docx file upload, reformats it to the standard BIRACS
    format, and returns a download URL.
    """
    # Validate file type
    if not file.filename or not file.filename.endswith('.docx'):
        raise HTTPException(
            status_code=400,
            detail="仅支持 .docx 格式的 Word 文档",
        )

    logger.info("Formatting document — filename=%s", file.filename)

    # Save uploaded file to temp location
    upload_id = uuid.uuid4().hex[:12]
    temp_input = Path(__file__).parent / "output" / f"upload_{upload_id}.docx"
    temp_input.parent.mkdir(exist_ok=True)

    try:
        content = await file.read()
        temp_input.write_bytes(content)
    except Exception as e:
        logger.error("Failed to save uploaded file: %s", e)
        raise HTTPException(status_code=500, detail=f"文件保存失败：{e}")

    # Run format_docx
    # Use original filename stem for output
    original_stem = file.filename
    if original_stem and original_stem.lower().endswith('.docx'):
        original_stem = original_stem[:-5]
    output_filename = f"{original_stem}_格式调整.docx"
    output_path = Path(__file__).parent / "output" / output_filename

    try:
        import importlib
        import format_docx
        importlib.reload(format_docx)  # Ensure latest code after hotfixes
        from format_docx import format_docx as run_format
        success, result_path = run_format(str(temp_input), str(output_path))
    except Exception as e:
        err_msg = str(e) if str(e) else f"未知错误（{type(e).__name__}）"
        logger.error("Format failed: %s", err_msg)
        # Clean up temp input
        temp_input.unlink(missing_ok=True)
        raise HTTPException(status_code=500, detail=f"格式调整失败：{err_msg}")

    # Clean up temp input
    temp_input.unlink(missing_ok=True)

    if not success:
        raise HTTPException(status_code=500, detail="格式调整失败，请检查文件是否损坏。")

    logger.info("Document formatted successfully: %s", output_filename)

    return {
        "success": True,
        "filename": output_filename,
        "url": f"/api/download/{output_filename}",
    }


@app.get("/api/download/{filename}")
async def download(filename: str):
    """Download a generated .docx file."""
    file_path = Path(__file__).parent / "output" / filename

    # Security: prevent path traversal
    file_path = file_path.resolve()
    output_dir = (Path(__file__).parent / "output").resolve()
    if not str(file_path).startswith(str(output_dir)):
        raise HTTPException(status_code=403, detail="禁止访问")

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在或已过期")

    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ─── Startup ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    import os
    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=port,
        reload=False,      # 生产环境不开启
        log_level="info",
    )
