import io
import zipfile
from pathlib import Path
from typing import Dict, Tuple

import pymupdf as fitz
from PIL import Image, ImageEnhance, ImageFilter

IMAGE_EXTS = {"jpg", "jpeg", "png", "webp", "bmp", "tiff", "tif", "gif"}


def _enhance_image(
    img: Image.Image,
    sharpness: float = 1.0,
    contrast: float = 1.0,
    brightness: float = 1.0,
    denoise: bool = False,
    upscale_factor: float = 1.0,
) -> Image.Image:
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGB")
    if upscale_factor > 1.0:
        new_w = max(1, int(img.width * upscale_factor))
        new_h = max(1, int(img.height * upscale_factor))
        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
    if denoise:
        img = img.filter(ImageFilter.MedianFilter(size=3))
    if abs(sharpness - 1.0) > 0.01:
        img = ImageEnhance.Sharpness(img).enhance(sharpness)
    if abs(contrast - 1.0) > 0.01:
        img = ImageEnhance.Contrast(img).enhance(contrast)
    if abs(brightness - 1.0) > 0.01:
        img = ImageEnhance.Brightness(img).enhance(brightness)
    return img


def _compress_image_to_target(img: Image.Image, ext: str, target_bytes: int) -> bytes:
    ext = ext.lower().lstrip(".")
    if ext in {"jpg", "jpeg"}:
        if img.mode in ("RGBA", "P", "LA"):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "P":
                img = img.convert("RGBA")
            bg.paste(img, mask=img.split()[-1] if "A" in img.mode else None)
            img = bg
        else:
            img = img.convert("RGB")

        work_img = img.copy()
        for _ in range(8):
            lo, hi = 8, 95
            best_buf = None
            while lo <= hi:
                mid = (lo + hi) // 2
                buf = io.BytesIO()
                work_img.save(buf, format="JPEG", quality=mid, optimize=True)
                data = buf.getvalue()
                if len(data) <= target_bytes:
                    best_buf = data
                    lo = mid + 1
                else:
                    hi = mid - 1
            if best_buf is not None:
                return best_buf
            w = max(32, int(work_img.width * 0.8))
            h = max(32, int(work_img.height * 0.8))
            work_img = work_img.resize((w, h), Image.Resampling.LANCZOS)

        buf = io.BytesIO()
        work_img.save(buf, format="JPEG", quality=12, optimize=True)
        return buf.getvalue()

    if ext == "webp":
        work_img = img.copy()
        for _ in range(8):
            lo, hi = 5, 95
            best_buf = None
            while lo <= hi:
                mid = (lo + hi) // 2
                buf = io.BytesIO()
                work_img.save(buf, format="WEBP", quality=mid, method=4)
                data = buf.getvalue()
                if len(data) <= target_bytes:
                    best_buf = data
                    lo = mid + 1
                else:
                    hi = mid - 1
            if best_buf is not None:
                return best_buf
            work_img = work_img.resize(
                (max(32, int(work_img.width * 0.8)), max(32, int(work_img.height * 0.8))),
                Image.Resampling.LANCZOS,
            )
        buf = io.BytesIO()
        work_img.save(buf, format="WEBP", quality=10)
        return buf.getvalue()

    work_img = img.copy()
    data = b""
    for step in range(10):
        buf = io.BytesIO()
        if ext == "png":
            quant = work_img.convert("P", palette=Image.Palette.ADAPTIVE, colors=max(16, 256 >> min(step, 4)))
            quant.save(buf, format="PNG", optimize=True)
        else:
            fmt_map = {"bmp": "BMP", "tiff": "TIFF", "tif": "TIFF", "gif": "GIF"}
            work_img.save(buf, format=fmt_map.get(ext, "PNG"))
        data = buf.getvalue()
        if len(data) <= target_bytes:
            return data
        work_img = work_img.resize(
            (max(32, int(work_img.width * 0.78)), max(32, int(work_img.height * 0.78))),
            Image.Resampling.LANCZOS,
        )
    return data


def _upscale_image_to_target(
    img: Image.Image,
    ext: str,
    target_bytes: int,
    sharpness: float,
    contrast: float,
    brightness: float,
    denoise: bool,
) -> bytes:
    ext = ext.lower().lstrip(".")
    enhanced = _enhance_image(
        img,
        sharpness=sharpness,
        contrast=contrast,
        brightness=brightness,
        denoise=denoise,
        upscale_factor=1.5,
    )

    data = b""
    for scale in [1.0, 1.4, 1.8, 2.3, 2.8, 3.5]:
        cand = enhanced.resize((int(enhanced.width * scale), int(enhanced.height * scale)), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        if ext in {"jpg", "jpeg"}:
            cand.convert("RGB").save(buf, format="JPEG", quality=98, subsampling=0, dpi=(300, 300))
        elif ext == "webp":
            cand.save(buf, format="WEBP", quality=98, lossless=True)
        elif ext in {"tiff", "tif"}:
            cand.save(buf, format="TIFF", dpi=(300, 300))
        else:
            cand.save(buf, format="PNG", compress_level=1, dpi=(300, 300))
        data = buf.getvalue()
        if len(data) >= target_bytes:
            return data

    if len(data) < target_bytes:
        data += bytes(target_bytes - len(data))
    return data


def _compress_pdf_to_target(pdf_bytes: bytes, target_bytes: int) -> bytes:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    clean_bytes = doc.tobytes(garbage=4, deflate=True, clean=True)
    if len(clean_bytes) <= target_bytes:
        return clean_bytes

    cand = clean_bytes
    for dpi, jpeg_q in [(150, 75), (120, 62), (96, 50), (72, 38), (60, 28), (50, 20)]:
        out_doc = fitz.open()
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_buf = io.BytesIO()
            img.save(img_buf, format="JPEG", quality=jpeg_q, optimize=True)
            new_pg = out_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_pg.insert_image(new_pg.rect, stream=img_buf.getvalue())
        cand = out_doc.tobytes(garbage=4, deflate=True)
        if len(cand) <= target_bytes:
            return cand
    return cand


def _upscale_pdf_to_target(pdf_bytes: bytes, target_bytes: int, sharpness: float = 1.25, contrast: float = 1.08) -> bytes:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    out_doc = fitz.open()
    for page in doc:
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img = _enhance_image(img, sharpness=sharpness, contrast=contrast, upscale_factor=1.25)
        img_buf = io.BytesIO()
        img.save(img_buf, format="PNG", compress_level=1)
        new_pg = out_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_pg.insert_image(new_pg.rect, stream=img_buf.getvalue())
    cand = out_doc.tobytes(garbage=2, deflate=False)
    if len(cand) < target_bytes:
        pad_len = target_bytes - len(cand)
        cand += b"\n%" + (b"0" * max(0, pad_len - 2))
    return cand


def transform_file_size_and_quality(
    filename: str,
    file_bytes: bytes,
    mode: str,
    target_size_value: float,
    target_size_unit: str,
    enhance_opts: Dict,
) -> Tuple[str, bytes, str]:
    ext = Path(filename).suffix.lower().lstrip(".")
    stem = Path(filename).stem
    multiplier = 1024 * 1024 if target_size_unit.upper() == "MB" else 1024
    target_bytes = max(1024, int(target_size_value * multiplier))

    sharpness = float(enhance_opts.get("sharpness", 1.2))
    contrast = float(enhance_opts.get("contrast", 1.05))
    brightness = float(enhance_opts.get("brightness", 1.0))
    denoise = bool(enhance_opts.get("denoise", False))
    upscale_factor = float(enhance_opts.get("upscale_factor", 1.5))

    if ext in IMAGE_EXTS:
        img = Image.open(io.BytesIO(file_bytes))
        mime_map = {
            "jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
            "webp": "image/webp", "bmp": "image/bmp", "tiff": "image/tiff",
            "tif": "image/tiff", "gif": "image/gif",
        }
        mime = mime_map.get(ext, "image/png")

        if mode == "reduce":
            out_bytes = _compress_image_to_target(img, ext, target_bytes)
            return f"{stem}_reduced.{ext}", out_bytes, mime
        elif mode == "enhance":
            out_bytes = _upscale_image_to_target(img, ext, target_bytes, sharpness, contrast, brightness, denoise)
            return f"{stem}_enhanced.{ext}", out_bytes, mime
        else:
            enh = _enhance_image(
                img,
                sharpness=sharpness,
                contrast=contrast,
                brightness=brightness,
                denoise=denoise,
                upscale_factor=upscale_factor,
            )
            buf = io.BytesIO()
            fmt = "JPEG" if ext in {"jpg", "jpeg"} else ("WEBP" if ext == "webp" else "PNG")
            enh.save(buf, format=fmt, quality=95)
            return f"{stem}_hd.{ext}", buf.getvalue(), mime

    if ext == "pdf":
        if mode == "reduce":
            out_bytes = _compress_pdf_to_target(file_bytes, target_bytes)
            return f"{stem}_reduced.pdf", out_bytes, "application/pdf"
        else:
            out_bytes = _upscale_pdf_to_target(file_bytes, target_bytes, sharpness=sharpness, contrast=contrast)
            return f"{stem}_enhanced.pdf", out_bytes, "application/pdf"

    if mode == "reduce":
        if len(file_bytes) <= target_bytes:
            return f"{stem}_optimized.{ext}", file_bytes, "application/octet-stream"
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", compression=zipfile.ZIP_LZMA) as zf:
            zf.writestr(filename, file_bytes)
        return f"{stem}_compressed.zip", zbuf.getvalue(), "application/zip"
    else:
        if len(file_bytes) >= target_bytes:
            return f"{stem}_enhanced.{ext}", file_bytes, "application/octet-stream"
        pad_needed = target_bytes - len(file_bytes)
        if ext in {"txt", "md", "csv", "tsv", "sql", "fasta", "fastq", "py", "js", "html", "xml", "yaml"}:
            comment_char = "#" if ext in {"py", "yaml", "csv", "tsv", "fasta"} else " "
            padded = file_bytes + (f"\n{comment_char} " + ("0" * max(0, pad_needed - 4))).encode("utf-8")
        else:
            padded = file_bytes + bytes(pad_needed)
        return f"{stem}_enhanced.{ext}", padded, "application/octet-stream"
