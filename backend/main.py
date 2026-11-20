import io
import json
import zipfile
from pathlib import Path
from typing import List

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, FileResponse
from PIL import Image, ImageDraw

from converter_engine import (
    ALL_SUPPORTED_TARGETS,
    FORMAT_CATEGORIES,
    convert_single_file,
    merge_files_to_single_pdf,
)
from size_engine import transform_file_size_and_quality

BASE_DIR = Path(__file__).resolve().parent.parent
FRONTEND_DIR = BASE_DIR / "frontend"
FRONTEND_INDEX = FRONTEND_DIR / "index.html"
CUSTOM_LOGO_PATH = FRONTEND_DIR / "logo.png"


def _build_default_janus_logo_png(size: int = 256) -> bytes:
    """Generates crisp monochrome dual-profile Janus logo matching tavus.jpg if logo.png is not replaced."""
    img = Image.new("RGB", (size, size), (0, 0, 0))
    d = ImageDraw.Draw(img)
    s = size / 100.0
    # Outer white frame
    d.rectangle([28 * s, 28 * s, 64 * s, 72 * s], outline=(255, 255, 255), width=max(2, int(3.2 * s)))
    # Right white profile silhouette
    right_poly = [
        (46 * s, 28 * s),
        (64 * s, 28 * s),
        (72 * s, 52 * s),
        (67 * s, 55 * s),
        (70 * s, 60 * s),
        (67 * s, 64 * s),
        (70 * s, 72 * s),
        (48 * s, 72 * s),
        (48 * s, 64 * s),
        (51 * s, 60 * s),
        (48 * s, 55 * s),
        (53 * s, 52 * s),
    ]
    d.polygon(right_poly, fill=(255, 255, 255))
    # Left square white eye
    d.rectangle([36 * s, 40 * s, 42 * s, 46 * s], fill=(255, 255, 255))
    # Right circular black eye
    d.ellipse([55 * s, 39.5 * s, 61.5 * s, 46 * s], fill=(0, 0, 0))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


app = FastAPI(title="all2all", version="4.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def serve_frontend():
    if FRONTEND_INDEX.exists():
        return FileResponse(FRONTEND_INDEX)
    return {"message": "all2all API is running."}


@app.get("/logo.png", include_in_schema=False)
def serve_logo():
    if CUSTOM_LOGO_PATH.exists():
        return FileResponse(CUSTOM_LOGO_PATH, media_type="image/png")
    return Response(content=_build_default_janus_logo_png(256), media_type="image/png")


@app.get("/favicon.ico", include_in_schema=False)
def serve_favicon():
    png_bytes = CUSTOM_LOGO_PATH.read_bytes() if CUSTOM_LOGO_PATH.exists() else _build_default_janus_logo_png(64)
    img = Image.open(io.BytesIO(png_bytes)).resize((64, 64))
    ico_buf = io.BytesIO()
    img.save(ico_buf, format="ICO", sizes=[(64, 64), (32, 32), (16, 16)])
    return Response(content=ico_buf.getvalue(), media_type="image/x-icon")


@app.get("/api/health")
def health_check():
    n = len(ALL_SUPPORTED_TARGETS)
    return {
        "status": "online",
        "service": "all2all",
        "categories": FORMAT_CATEGORIES,
        "total_formats": n,
        "total_conversion_pairs": n * (n - 1),
    }


@app.post("/api/convert")
async def convert_endpoint(
    files: List[UploadFile] = File(...),
    target_format: str = Form(...),
    merge_into_pdf: bool = Form(False),
    options_json: str = Form("{}"),
):
    tgt = target_format.lower().lstrip(".")
    if tgt not in ALL_SUPPORTED_TARGETS:
        raise HTTPException(status_code=400, detail=f"Unsupported target format: {target_format}")
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded.")

    try:
        options = json.loads(options_json)
    except Exception:
        options = {}

    raw_items = [(f.filename or "file", await f.read()) for f in files]

    try:
        if merge_into_pdf and tgt == "pdf" and len(raw_items) > 1:
            merged_pdf = merge_files_to_single_pdf(raw_items)
            return Response(
                content=merged_pdf,
                media_type="application/pdf",
                headers={"Content-Disposition": 'attachment; filename="all2all_Merged_Bundle.pdf"'},
            )

        if len(raw_items) == 1:
            out_name, out_bytes, mime = convert_single_file(raw_items[0][0], raw_items[0][1], tgt, options)
            return Response(
                content=out_bytes,
                media_type=mime,
                headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
            )

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            seen = set()
            for fname, fbytes in raw_items:
                out_name, out_bytes, _ = convert_single_file(fname, fbytes, tgt, options)
                if out_name in seen:
                    out_name = f"{len(seen)}_{out_name}"
                seen.add(out_name)
                zf.writestr(out_name, out_bytes)

        return Response(
            content=zip_buf.getvalue(),
            media_type="application/zip",
            headers={"Content-Disposition": f'attachment; filename="all2all_Batch_{tgt.upper()}.zip"'},
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/convert-text")
async def convert_raw_text_endpoint(
    text_content: str = Form(...),
    source_ext: str = Form("txt"),
    target_format: str = Form(...),
    filename_stem: str = Form("all2all_snippet"),
    options_json: str = Form("{}"),
):
    try:
        options = json.loads(options_json)
        out_name, out_bytes, mime = convert_single_file(
            f"{filename_stem}.{source_ext.lstrip('.')}", text_content.encode("utf-8"), target_format, options
        )
        return Response(
            content=out_bytes,
            media_type=mime,
            headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/size-transform")
async def size_transform_endpoint(
    files: List[UploadFile] = File(...),
    mode: str = Form("reduce"),
    target_size_value: float = Form(200.0),
    target_size_unit: str = Form("KB"),
    enhance_json: str = Form("{}"),
):
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded.")

    try:
        enhance_opts = json.loads(enhance_json)
    except Exception:
        enhance_opts = {}

    raw_items = [(f.filename or "file", await f.read()) for f in files]

    try:
        if len(raw_items) == 1:
            out_name, out_bytes, mime = transform_file_size_and_quality(
                raw_items[0][0], raw_items[0][1], mode, target_size_value, target_size_unit, enhance_opts
            )
            return Response(
                content=out_bytes,
                media_type=mime,
                headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
            )

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            seen = set()
            for fname, fbytes in raw_items:
                out_name, out_bytes, _ = transform_file_size_and_quality(
                    fname, fbytes, mode, target_size_value, target_size_unit, enhance_opts
                )
                if out_name in seen:
                    out_name = f"{len(seen)}_{out_name}"
                seen.add(out_name)
                zf.writestr(out_name, out_bytes)

        return Response(
            content=zip_buf.getvalue(),
            media_type="application/zip",
            headers={"Content-Disposition": f'attachment; filename="all2all_{mode.upper()}_Batch.zip"'},
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))
