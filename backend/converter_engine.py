import base64
import csv
import io
import json
import os
import re
import sqlite3
import tarfile
import tempfile
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pymupdf as fitz
from PIL import Image, ImageOps
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
import openpyxl
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import markdown as md_lib
import yaml
import qrcode


FORMAT_CATEGORIES = {
    "Documents & E-Books": ["pdf", "docx", "rtf", "txt", "md", "html", "epub", "tex", "org", "rst", "adoc"],
    "Presentations": ["pptx"],
    "Spreadsheets & Databases": ["xlsx", "csv", "tsv", "json", "jsonl", "xml", "yaml", "toml", "ini", "env", "sql", "sqlite"],
    "Images & Vector": ["jpg", "png", "webp", "bmp", "tiff", "gif", "ico", "svg", "eps", "ppm", "tga"],
    "Web, Code & Encoding": ["b64", "hex", "urlenc", "qr_png", "qr_svg", "zip", "tar_gz"],
    "Bioinformatics & Sequence": ["fasta", "fastq", "gbk", "gff3", "bed", "vcf", "sam"],
}

ALL_SUPPORTED_TARGETS = {
    fmt for group in FORMAT_CATEGORIES.values() for fmt in group
} | {"jpeg", "tif", "yml", "htm", "fa", "fna", "fq", "db", "tgz"}

IMAGE_FORMATS = {"jpg", "jpeg", "png", "webp", "bmp", "tiff", "tif", "gif", "ico", "ppm", "tga"}
TABULAR_FORMATS = {"csv", "tsv", "xlsx", "json", "jsonl", "xml", "yaml", "yml", "toml", "ini", "env", "sql", "sqlite", "db"}
BIO_FORMATS = {"fasta", "fa", "fna", "fastq", "fq", "gbk", "gff3", "bed", "vcf", "sam"}


def _normalize_ext(ext: str) -> str:
    e = ext.lower().lstrip(".")
    aliases = {
        "jpeg": "jpg", "tif": "tiff", "yml": "yaml", "htm": "html",
        "fa": "fasta", "fna": "fasta", "fq": "fastq", "db": "sqlite", "tgz": "tar_gz"
    }
    return aliases.get(e, e)


def _text_to_pdf_bytes(text: str, title: str = "all2all Universal Export") -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4
    margin = 46

    c.setFillColorRGB(0.0, 0.129, 0.278)
    c.rect(0, height - 34, width, 34, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(margin, height - 22, f"all2all  |  {title[:75]}")

    c.setFillColorRGB(0.06, 0.09, 0.16)
    c.setFont("Helvetica", 10)
    y = height - 58
    max_width = width - 2 * margin

    for raw_line in text.splitlines() or [""]:
        wrapped = simpleSplit(raw_line, "Helvetica", 10, max_width) or [""]
        for line in wrapped:
            if y < margin:
                c.showPage()
                c.setFont("Helvetica", 10)
                y = height - margin
            c.drawString(margin, y, line)
            y -= 14

    c.save()
    return buf.getvalue()


def _table_to_pdf_bytes(rows: List[List[str]], title: str = "Structured Data Export") -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=28, bottomMargin=24)
    styles = getSampleStyleSheet()
    elements = [Paragraph(f"<b>{title}</b>", styles["Heading2"])]
    if not rows:
        rows = [["Empty"]]

    trimmed_rows = [[str(cell)[:55] if cell is not None else "" for cell in row[:14]] for row in rows[:300]]
    table = Table(trimmed_rows, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#002147")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    elements.append(table)
    doc.build(elements)
    return buf.getvalue()


def _text_to_docx_bytes(text: str, title: str = "Converted Document") -> bytes:
    doc = Document()
    doc.add_heading(title, level=1)
    for para in text.split("\n"):
        doc.add_paragraph(para)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _text_to_pptx_bytes(text: str, title: str = "Presentation") -> bytes:
    prs = Presentation()
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = title
    title_slide.placeholders[1].text = "Converted via all2all"

    lines = [l.strip() for l in text.splitlines() if l.strip()]
    chunk_size = 8
    for idx in range(0, max(1, len(lines)), chunk_size):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"{title} ({idx // chunk_size + 1})"
        tf = slide.placeholders[1].text_frame
        tf.clear()
        chunk = lines[idx:idx + chunk_size] or ["(No content)"]
        for i, line in enumerate(chunk):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line[:240]
            p.font.size = Pt(15)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _text_to_rtf_bytes(text: str) -> bytes:
    escaped = text.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}").replace("\n", "\\par\n")
    return f"{{\\rtf1\\ansi\\deff0{{\\fonttbl{{\\f0 Arial;}}}}\\fs22 {escaped}}}".encode("utf-8")


def _text_to_latex_bytes(text: str, title: str) -> bytes:
    safe_title = re.sub(r"([#$%&_{}~^\\])", r"\\\1", title)
    safe_body = re.sub(r"([#$%&_{}~^\\])", r"\\\1", text)
    tex = (
        "\\documentclass[11pt,a4paper]{article}\n"
        "\\usepackage[utf8]{inputenc}\n"
        "\\usepackage{geometry}\n"
        "\\geometry{margin=1in}\n"
        f"\\title{{{safe_title}}}\n"
        "\\date{\\today}\n"
        "\\begin{document}\n"
        "\\maketitle\n\n"
        f"{safe_body}\n\n"
        "\\end{document}\n"
    )
    return tex.encode("utf-8")


def _text_to_epub_bytes(text: str, title: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("mimetype", "application/epub+zip", compress_type=zipfile.ZIP_STORED)
        zf.writestr(
            "META-INF/container.xml",
            '<?xml version="1.0"?><container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">'
            '<rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>'
        )
        html_paras = "".join(f"<p>{line}</p>" for line in text.splitlines() if line.strip())
        zf.writestr(
            "OEBPS/chapter1.xhtml",
            f'<?xml version="1.0" encoding="utf-8"?><html xmlns="http://www.w3.org/1999/xhtml"><head><title>{title}</title></head>'
            f'<body><h1>{title}</h1>{html_paras}</body></html>'
        )
        zf.writestr(
            "OEBPS/content.opf",
            f'<?xml version="1.0" encoding="utf-8"?><package xmlns="http://www.idpf.org/2007/opf" unique-identifier="BookId" version="2.0">'
            f'<metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>{title}</dc:title><dc:language>en</dc:language></metadata>'
            f'<manifest><item id="ch1" href="chapter1.xhtml" media-type="application/xhtml+xml"/></manifest>'
            f'<spine><itemref idref="ch1"/></spine></package>'
        )
    return buf.getvalue()


def _extract_text(data: bytes, ext: str) -> str:
    ext = _normalize_ext(ext)
    if ext in {"txt", "md", "html", "csv", "tsv", "json", "jsonl", "xml", "yaml", "toml", "ini", "env", "sql", "tex", "org", "rst", "adoc", "fasta", "fastq", "gbk", "gff3", "bed", "vcf", "sam", "svg"}:
        raw = data.decode("utf-8", errors="replace")
        if ext == "html":
            return re.sub(r"<[^>]+>", " ", raw)
        return raw
    if ext == "rtf":
        raw = data.decode("utf-8", errors="replace")
        return re.sub(r"\\[a-z]+\d* ?|[{}]", "", raw)
    if ext == "pdf":
        doc = fitz.open(stream=data, filetype="pdf")
        return "\n\n".join(f"--- Page {i+1} ---\n{p.get_text('text')}" for i, p in enumerate(doc))
    if ext == "docx":
        doc = Document(io.BytesIO(data))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for tbl in doc.tables:
            for r in tbl.rows:
                parts.append(" | ".join(c.text.strip() for c in r.cells))
        return "\n".join(parts)
    if ext == "pptx":
        prs = Presentation(io.BytesIO(data))
        parts = []
        for idx, slide in enumerate(prs.slides, 1):
            lines = [f"--- Slide {idx} ---"]
            for sh in slide.shapes:
                if sh.has_text_frame:
                    lines.append(sh.text_frame.text)
            parts.append("\n".join(lines))
        return "\n\n".join(parts)
    if ext == "xlsx":
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        parts = []
        for name in wb.sheetnames:
            ws = wb[name]
            parts.append(f"=== Sheet: {name} ===")
            for row in ws.iter_rows(values_only=True):
                parts.append("\t".join("" if v is None else str(v) for v in row))
        return "\n".join(parts)
    if ext == "epub":
        parts = []
        with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
            for name in sorted(zf.namelist()):
                if name.endswith((".xhtml", ".html", ".htm")):
                    txt = zf.read(name).decode("utf-8", errors="replace")
                    parts.append(re.sub(r"<[^>]+>", " ", txt))
        return "\n\n".join(parts)
    if ext == "sqlite":
        with tempfile.NamedTemporaryFile(suffix=".sqlite", delete=False) as tf:
            tf.write(data)
            tmp_path = tf.name
        try:
            conn = sqlite3.connect(tmp_path)
            cur = conn.cursor()
            tables = [r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
            out = []
            for tbl in tables:
                out.append(f"=== Table: {tbl} ===")
                rows = cur.execute(f'SELECT * FROM "{tbl}"').fetchall()
                cols = [d[0] for d in cur.description]
                out.append("\t".join(cols))
                for r in rows:
                    out.append("\t".join("" if v is None else str(v) for v in r))
            conn.close()
            return "\n".join(out)
        finally:
            os.remove(tmp_path)
    if ext in IMAGE_FORMATS:
        img = Image.open(io.BytesIO(data))
        return f"Image Metadata Summary\nFormat: {img.format}\nDimensions: {img.width} x {img.height} px\nColor Mode: {img.mode}"
    return data.decode("utf-8", errors="replace")


def _extract_rows(data: bytes, ext: str) -> List[List[str]]:
    ext = _normalize_ext(ext)
    if ext == "csv":
        return list(csv.reader(io.StringIO(data.decode("utf-8", errors="replace")))) or [["Empty"]]
    if ext == "tsv":
        return list(csv.reader(io.StringIO(data.decode("utf-8", errors="replace")), delimiter="\t")) or [["Empty"]]
    if ext == "xlsx":
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        ws = wb.active
        return [["" if v is None else str(v) for v in row] for row in ws.iter_rows(values_only=True)] or [["Empty"]]
    if ext in {"json", "jsonl"}:
        raw = data.decode("utf-8", errors="replace").strip()
        records = [json.loads(line) for line in raw.splitlines() if line.strip()] if ext == "jsonl" else (
            json.loads(raw) if isinstance(json.loads(raw), list) else [json.loads(raw)]
        )
        if records and isinstance(records[0], dict):
            headers = list(dict.fromkeys(k for r in records if isinstance(r, dict) for k in r.keys()))
            return [headers] + [["" if r.get(h) is None else str(r.get(h)) for h in headers] for r in records]
        return [["Value"]] + [[str(x)] for x in records]
    if ext in {"yaml", "yml"}:
        obj = yaml.safe_load(data.decode("utf-8", errors="replace"))
        records = obj if isinstance(obj, list) else [obj]
        if records and isinstance(records[0], dict):
            headers = list(dict.fromkeys(k for r in records if isinstance(r, dict) for k in r.keys()))
            return [headers] + [["" if r.get(h) is None else str(r.get(h)) for h in headers] for r in records]
        return [["Value"]] + [[str(x)] for x in records]
    if ext in {"ini", "env", "toml"}:
        rows = [["Key", "Value"]]
        for line in data.decode("utf-8", errors="replace").splitlines():
            line = line.strip()
            if not line or line.startswith(("#", ";", "[")):
                continue
            if "=" in line:
                k, v = line.split("=", 1)
                rows.append([k.strip(), v.strip().strip('"').strip("'")])
        return rows
    if ext == "xml":
        root = ET.fromstring(data.decode("utf-8", errors="replace"))
        children = list(root)
        if not children:
            return [["Tag", "Text"], [root.tag, root.text or ""]]
        headers = list(dict.fromkeys(sub.tag for item in children for sub in item))
        if not headers:
            return [["Tag", "Value"]] + [[item.tag, item.text or ""] for item in children]
        return [headers] + [[item.findtext(h, default="") for h in headers] for item in children]
    if ext == "sqlite":
        with tempfile.NamedTemporaryFile(suffix=".sqlite", delete=False) as tf:
            tf.write(data)
            tmp_path = tf.name
        try:
            conn = sqlite3.connect(tmp_path)
            cur = conn.cursor()
            tables = [r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
            if not tables:
                return [["Empty Database"]]
            rows = cur.execute(f'SELECT * FROM "{tables[0]}"').fetchall()
            cols = [d[0] for d in cur.description]
            conn.close()
            return [cols] + [["" if v is None else str(v) for v in r] for r in rows]
        finally:
            os.remove(tmp_path)
    if ext in {"fasta", "fastq", "gbk"}:
        seqs = _parse_bio_sequences(data.decode("utf-8", errors="replace"), ext)
        return [["ID", "Description", "Length", "GC_Percent", "Sequence"]] + [
            [s["id"], s["desc"], str(len(s["seq"])), f"{_gc_percent(s['seq']):.2f}", s["seq"]] for s in seqs
        ]
    text = _extract_text(data, ext)
    return [["Line", "Content"]] + [[str(i + 1), l] for i, l in enumerate(text.splitlines())]


def _gc_percent(seq: str) -> float:
    s = re.sub(r"[^ACGTUacgtu]", "", seq).upper()
    return (sum(1 for ch in s if ch in ("G", "C")) / len(s) * 100.0) if s else 0.0


def _parse_bio_sequences(text: str, ext: str) -> List[Dict[str, str]]:
    ext = _normalize_ext(ext)
    records = []
    if ext == "fasta":
        cur_id, cur_desc, chunks = "seq_1", "", []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            if line.startswith(">"):
                if chunks:
                    seq = "".join(chunks)
                    records.append({"id": cur_id, "desc": cur_desc, "seq": seq, "qual": "I" * len(seq)})
                header = line[1:].strip().split(None, 1)
                cur_id = header[0] if header else "seq"
                cur_desc = header[1] if len(header) > 1 else ""
                chunks = []
            else:
                chunks.append(line)
        if chunks:
            seq = "".join(chunks)
            records.append({"id": cur_id, "desc": cur_desc, "seq": seq, "qual": "I" * len(seq)})
    elif ext == "fastq":
        lines = [l.rstrip("\r\n") for l in text.splitlines() if l.strip()]
        for i in range(0, len(lines) - 3, 4):
            hdr = lines[i].lstrip("@").split(None, 1)
            seq = lines[i + 1].strip()
            qual = lines[i + 3].strip()
            records.append({
                "id": hdr[0] if hdr else f"read_{i//4+1}",
                "desc": hdr[1] if len(hdr) > 1 else "",
                "seq": seq,
                "qual": qual if len(qual) == len(seq) else "I" * len(seq),
            })
    elif ext == "gbk":
        for entry in text.split("//"):
            if "ORIGIN" not in entry:
                continue
            locus = re.search(r"^LOCUS\s+(\S+)", entry, re.M)
            defn = re.search(r"^DEFINITION\s+(.+)", entry, re.M)
            seq_id = locus.group(1) if locus else "genbank_record"
            desc = defn.group(1).rstrip(".") if defn else "Converted from GenBank"
            origin_match = re.search(r"ORIGIN([\s\S]+)", entry)
            seq = re.sub(r"[^A-Za-z]", "", origin_match.group(1)).upper() if origin_match else ""
            if seq:
                records.append({"id": seq_id, "desc": desc, "seq": seq, "qual": "I" * len(seq)})
    return records


def _sequences_to_bio_bytes(records: List[Dict[str, str]], tgt: str, filename: str) -> bytes:
    if tgt == "fasta":
        return ("\n".join(
            f">{r['id']} {r['desc']}".strip() + "\n" + "\n".join(r["seq"][i:i+80] for i in range(0, len(r["seq"]), 80))
            for r in records
        ) + "\n").encode("utf-8")
    if tgt == "fastq":
        return ("\n".join(f"@{r['id']} {r['desc']}".strip() + f"\n{r['seq']}\n+\n{r['qual']}" for r in records) + "\n").encode("utf-8")
    if tgt == "gbk":
        entries = []
        for r in records:
            seq = r["seq"].lower()
            origin_lines = [
                f"{i+1:>9} " + " ".join(seq[i:i+60][j:j+10] for j in range(0, len(seq[i:i+60]), 10))
                for i in range(0, len(seq), 60)
            ]
            entries.append(
                f"LOCUS       {r['id']:<16} {len(seq)} bp    DNA     linear   UNA\n"
                f"DEFINITION  {r['desc'] or filename}.\nFEATURES             Location/Qualifiers\n"
                f"     source          1..{len(seq)}\nORIGIN\n" + "\n".join(origin_lines) + "\n//"
            )
        return ("\n".join(entries) + "\n").encode("utf-8")
    if tgt == "gff3":
        return ("##gff-version 3\n" + "\n".join(
            f"{r['id']}\tall2all\tregion\t1\t{max(1, len(r['seq']))}\t.\t+\t.\tID={r['id']};Note={r['desc']}" for r in records
        ) + "\n").encode("utf-8")
    if tgt == "bed":
        return ("\n".join(f"{r['id']}\t0\t{max(1, len(r['seq']))}\t{r['id']}\t1000\t+" for r in records) + "\n").encode("utf-8")
    if tgt == "vcf":
        hdr = "##fileformat=VCFv4.2\n#CHROM\tPOS\tID\tREF\tALT\tQUAL\tFILTER\tINFO\n"
        rows = [f"{r['id']}\t1\t.\t{r['seq'][:1] or 'N'}\t.\t100\tPASS\tLEN={len(r['seq'])}" for r in records]
        return (hdr + "\n".join(rows) + "\n").encode("utf-8")
    if tgt == "sam":
        hdr = "@HD\tVN:1.6\tSO:unsorted\n" + "\n".join(f"@SQ\tSN:{r['id']}\tLN:{max(1, len(r['seq']))}" for r in records)
        aligns = "\n".join(f"{r['id']}\t4\t*\t0\t0\t*\t*\t0\t0\t{r['seq']}\t{r['qual']}" for r in records)
        return (hdr + "\n" + aligns + "\n").encode("utf-8")
    return b""


def _rows_to_target(rows: List[List[str]], tgt: str, stem: str) -> Tuple[str, bytes, str]:
    headers = rows[0] if rows else ["Column1"]
    data_rows = rows[1:] if len(rows) > 1 else []

    if tgt == "csv":
        buf = io.StringIO()
        csv.writer(buf).writerows(rows)
        return f"{stem}.csv", buf.getvalue().encode("utf-8"), "text/csv"
    if tgt == "tsv":
        buf = io.StringIO()
        csv.writer(buf, delimiter="\t").writerows(rows)
        return f"{stem}.tsv", buf.getvalue().encode("utf-8"), "text/tab-separated-values"
    if tgt == "xlsx":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        for r in rows:
            ws.append(r)
        buf = io.BytesIO()
        wb.save(buf)
        return f"{stem}.xlsx", buf.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if tgt == "json":
        records = [dict(zip(headers, r)) for r in data_rows] if data_rows else rows
        return f"{stem}.json", json.dumps(records, indent=2, ensure_ascii=False).encode("utf-8"), "application/json"
    if tgt == "jsonl":
        records = [dict(zip(headers, r)) for r in data_rows]
        return f"{stem}.jsonl", "\n".join(json.dumps(rec, ensure_ascii=False) for rec in records).encode("utf-8"), "application/x-ndjson"
    if tgt == "yaml":
        records = [dict(zip(headers, r)) for r in data_rows] if data_rows else rows
        return f"{stem}.yaml", yaml.safe_dump(records, sort_keys=False, allow_unicode=True).encode("utf-8"), "text/yaml"
    if tgt in {"toml", "ini", "env"}:
        lines = []
        for idx, r in enumerate(data_rows or [headers], 1):
            if len(r) >= 2:
                k = re.sub(r"[^A-Za-z0-9_]", "_", r[0]) or f"KEY_{idx}"
                v = str(r[1]).replace('"', '\"')
                lines.append(f'{k}="{v}"' if tgt != "ini" else f"{k} = {v}")
            else:
                lines.append(f'ITEM_{idx}="{r[0]}"')
        return f"{stem}.{tgt}", ("\n".join(lines) + "\n").encode("utf-8"), "text/plain"
    if tgt == "xml":
        root = ET.Element("dataset")
        safe_tags = [re.sub(r"[^A-Za-z0-9_]", "_", h) or f"col_{i}" for i, h in enumerate(headers)]
        for r in data_rows:
            rec_el = ET.SubElement(root, "record")
            for tag, val in zip(safe_tags, r):
                child = ET.SubElement(rec_el, tag)
                child.text = str(val)
        return f"{stem}.xml", b'<?xml version="1.0" encoding="utf-8"?>\n' + ET.tostring(root, encoding="utf-8"), "application/xml"
    if tgt == "sql":
        tbl = re.sub(r"[^A-Za-z0-9_]", "_", stem) or "exported_table"
        cols = [re.sub(r"[^A-Za-z0-9_]", "_", h) or f"col_{i}" for i, h in enumerate(headers)]
        ddl = f'CREATE TABLE "{tbl}" ({", ".join(f'"{c}" TEXT' for c in cols)});\n'
        inserts = [f'INSERT INTO "{tbl}" VALUES ({", ".join("'" + str(v).replace("'", "''") + "'" for v in r)});' for r in data_rows]
        return f"{stem}.sql", (ddl + "\n".join(inserts) + "\n").encode("utf-8"), "application/sql"
    if tgt == "sqlite":
        tbl = re.sub(r"[^A-Za-z0-9_]", "_", stem) or "exported_table"
        cols = [re.sub(r"[^A-Za-z0-9_]", "_", h) or f"col_{i}" for i, h in enumerate(headers)]
        with tempfile.NamedTemporaryFile(suffix=".sqlite", delete=False) as tf:
            tmp_path = tf.name
        try:
            conn = sqlite3.connect(tmp_path)
            cur = conn.cursor()
            cur.execute(f'CREATE TABLE "{tbl}" ({", ".join(f'"{c}" TEXT' for c in cols)})')
            placeholders = ", ".join("?" for _ in cols)
            for r in data_rows:
                cur.execute(f'INSERT INTO "{tbl}" VALUES ({placeholders})', (r + [""] * len(cols))[:len(cols)])
            conn.commit()
            conn.close()
            with open(tmp_path, "rb") as f:
                return f"{stem}.sqlite", f.read(), "application/vnd.sqlite3"
        finally:
            os.remove(tmp_path)
    if tgt == "pdf":
        return f"{stem}.pdf", _table_to_pdf_bytes(rows, title=stem), "application/pdf"
    raise ValueError(f"Unsupported tabular target: {tgt}")


def _apply_image_options(img: Image.Image, opts: Dict) -> Image.Image:
    rotate_deg = int(opts.get("rotate", 0)) % 360
    if rotate_deg:
        img = img.rotate(-rotate_deg, expand=True)
    if opts.get("grayscale"):
        img = ImageOps.grayscale(img).convert("RGB")
    resize_w = int(opts.get("resize_width", 0) or 0)
    if resize_w > 0 and img.width > 0:
        ratio = resize_w / float(img.width)
        img = img.resize((resize_w, max(1, int(img.height * ratio))), Image.Resampling.LANCZOS)
    return img


def _image_to_svg_bytes(img: Image.Image) -> bytes:
    png_buf = io.BytesIO()
    img.convert("RGBA").save(png_buf, format="PNG")
    b64 = base64.b64encode(png_buf.getvalue()).decode("ascii")
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{img.width}" height="{img.height}" viewBox="0 0 {img.width} {img.height}">'
        f'<image width="{img.width}" height="{img.height}" href="data:image/png;base64,{b64}"/></svg>'
    ).encode("utf-8")


def convert_single_file(filename: str, file_bytes: bytes, target_ext: str, options: Optional[Dict] = None) -> Tuple[str, bytes, str]:
    opts = options or {}
    src = _normalize_ext(Path(filename).suffix)
    stem = Path(filename).stem
    tgt = _normalize_ext(target_ext)
    quality = max(10, min(100, int(opts.get("quality", 90) or 90)))

    if tgt == "b64":
        return f"{stem}_{src}.b64.txt", base64.b64encode(file_bytes), "text/plain; charset=utf-8"
    if tgt == "hex":
        return f"{stem}_{src}.hex.txt", file_bytes.hex().encode("utf-8"), "text/plain; charset=utf-8"
    if tgt == "urlenc":
        txt = _extract_text(file_bytes, src)
        return f"{stem}.urlenc.txt", urllib.parse.quote(txt).encode("utf-8"), "text/plain; charset=utf-8"
    if tgt == "zip":
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(filename, file_bytes)
        return f"{stem}.zip", zbuf.getvalue(), "application/zip"
    if tgt == "tar_gz":
        tbuf = io.BytesIO()
        with tarfile.open(fileobj=tbuf, mode="w:gz") as tf:
            info = tarfile.TarInfo(name=filename)
            info.size = len(file_bytes)
            tf.addfile(info, io.BytesIO(file_bytes))
        return f"{stem}.tar.gz", tbuf.getvalue(), "application/gzip"
    if tgt in {"qr_png", "qr_svg"}:
        payload = _extract_text(file_bytes, src)[:1800] or filename
        qr = qrcode.QRCode(box_size=8, border=2)
        qr.add_data(payload)
        qr.make(fit=True)
        img = qr.make_image(fill_color="#002147", back_color="white").convert("RGB")
        if tgt == "qr_png":
            out = io.BytesIO()
            img.save(out, format="PNG")
            return f"{stem}_qr.png", out.getvalue(), "image/png"
        return f"{stem}_qr.svg", _image_to_svg_bytes(img), "image/svg+xml"

    if src in IMAGE_FORMATS:
        img = _apply_image_options(Image.open(io.BytesIO(file_bytes)), opts)
        if tgt in IMAGE_FORMATS:
            if tgt in {"jpg", "bmp", "eps", "ppm"} and img.mode in ("RGBA", "P", "LA"):
                bg = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                bg.paste(img, mask=img.split()[-1] if "A" in img.mode else None)
                img = bg
            elif img.mode == "CMYK":
                img = img.convert("RGB")
            if tgt == "ico":
                icon_img = img.copy()
                icon_img.thumbnail((256, 256))
                out = io.BytesIO()
                icon_img.save(out, format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (32, 32)])
                return f"{stem}.ico", out.getvalue(), "image/x-icon"
            pil_map = {"jpg": "JPEG", "png": "PNG", "webp": "WEBP", "bmp": "BMP", "tiff": "TIFF", "gif": "GIF", "ppm": "PPM", "tga": "TGA"}
            mime_map = {"jpg": "image/jpeg", "png": "image/png", "webp": "image/webp", "bmp": "image/bmp", "tiff": "image/tiff", "gif": "image/gif", "ppm": "image/x-portable-pixmap", "tga": "image/x-tga"}
            out = io.BytesIO()
            img.save(out, format=pil_map[tgt], quality=quality)
            return f"{stem}.{tgt}", out.getvalue(), mime_map[tgt]
        if tgt == "eps":
            out = io.BytesIO()
            img.convert("RGB").save(out, format="EPS")
            return f"{stem}.eps", out.getvalue(), "application/postscript"
        if tgt == "svg":
            return f"{stem}.svg", _image_to_svg_bytes(img), "image/svg+xml"
        if tgt == "pdf":
            out = io.BytesIO()
            img.convert("RGB").save(out, format="PDF", resolution=150.0)
            return f"{stem}.pdf", out.getvalue(), "application/pdf"
        if tgt == "pptx":
            img_buf = io.BytesIO()
            img.convert("RGB").save(img_buf, format="PNG")
            img_buf.seek(0)
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(img_buf, Inches(0.5), Inches(0.5), width=Inches(9.0))
            out = io.BytesIO()
            prs.save(out)
            return f"{stem}.pptx", out.getvalue(), "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    if src == "pdf" and tgt in (IMAGE_FORMATS | {"svg", "eps"}):
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        dpi = int(opts.get("dpi", 170) or 170)
        if len(doc) == 1:
            pix = doc[0].get_pixmap(dpi=dpi)
            img = _apply_image_options(Image.frombytes("RGB", [pix.width, pix.height], pix.samples), opts)
            png_buf = io.BytesIO()
            img.save(png_buf, format="PNG")
            return convert_single_file(f"{stem}_page1.png", png_buf.getvalue(), tgt, opts)
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, page in enumerate(doc, 1):
                pix = page.get_pixmap(dpi=dpi)
                img = _apply_image_options(Image.frombytes("RGB", [pix.width, pix.height], pix.samples), opts)
                png_buf = io.BytesIO()
                img.save(png_buf, format="PNG")
                p_name, p_bytes, _ = convert_single_file(f"{stem}_page_{idx}.png", png_buf.getvalue(), tgt, opts)
                zf.writestr(p_name, p_bytes)
        return f"{stem}_{tgt}_pages.zip", zip_buf.getvalue(), "application/zip"

    if src == "pdf" and tgt == "pptx":
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        for page in doc:
            pix = page.get_pixmap(dpi=150)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(io.BytesIO(pix.tobytes("png")), Inches(0), Inches(0), height=Inches(7.5))
        out = io.BytesIO()
        prs.save(out)
        return f"{stem}.pptx", out.getvalue(), "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    if src in BIO_FORMATS and tgt in BIO_FORMATS:
        records = _parse_bio_sequences(file_bytes.decode("utf-8", errors="replace"), src)
        if not records:
            raise ValueError(f"No valid biological sequences found in .{src} file.")
        return f"{stem}.{tgt}", _sequences_to_bio_bytes(records, tgt, filename), "text/plain; charset=utf-8"

    if src in TABULAR_FORMATS and tgt in (TABULAR_FORMATS | {"pdf"}):
        return _rows_to_target(_extract_rows(file_bytes, src), tgt, stem)

    extracted_text = _extract_text(file_bytes, src)
    if tgt == "pdf":
        return f"{stem}.pdf", _text_to_pdf_bytes(extracted_text, title=stem), "application/pdf"
    if tgt == "docx":
        return f"{stem}.docx", _text_to_docx_bytes(extracted_text, title=stem), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if tgt == "pptx":
        return f"{stem}.pptx", _text_to_pptx_bytes(extracted_text, title=stem), "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    if tgt == "rtf":
        return f"{stem}.rtf", _text_to_rtf_bytes(extracted_text), "application/rtf"
    if tgt == "tex":
        return f"{stem}.tex", _text_to_latex_bytes(extracted_text, stem), "application/x-tex"
    if tgt == "epub":
        return f"{stem}.epub", _text_to_epub_bytes(extracted_text, stem), "application/epub+zip"
    if tgt in {"txt", "org", "rst", "adoc"}:
        return f"{stem}.{tgt}", extracted_text.encode("utf-8"), "text/plain; charset=utf-8"
    if tgt == "md":
        return f"{stem}.md", f"# {stem}\n\n{extracted_text}".encode("utf-8"), "text/markdown; charset=utf-8"
    if tgt == "html":
        body = md_lib.markdown(extracted_text) if src == "md" else f"<pre style='white-space:pre-wrap;font-family:Inter,sans-serif'>{extracted_text}</pre>"
        html_doc = f"<!DOCTYPE html><html><head><meta charset='utf-8'><title>{stem}</title></head><body style='max-width:860px;margin:2rem auto;padding:1rem;'>{body}</body></html>"
        return f"{stem}.html", html_doc.encode("utf-8"), "text/html; charset=utf-8"
    if tgt in TABULAR_FORMATS:
        return _rows_to_target(_extract_rows(file_bytes, src), tgt, stem)

    raise ValueError(f"Conversion .{src} -> .{tgt} is not supported.")


def merge_files_to_single_pdf(files: List[Tuple[str, bytes]]) -> bytes:
    out_doc = fitz.open()
    for fname, raw in files:
        ext = _normalize_ext(Path(fname).suffix)
        if ext == "pdf":
            out_doc.insert_pdf(fitz.open(stream=raw, filetype="pdf"))
        elif ext in IMAGE_FORMATS:
            img = Image.open(io.BytesIO(raw)).convert("RGB")
            pdf_buf = io.BytesIO()
            img.save(pdf_buf, format="PDF", resolution=150.0)
            out_doc.insert_pdf(fitz.open(stream=pdf_buf.getvalue(), filetype="pdf"))
        else:
            _, pdf_bytes, _ = convert_single_file(fname, raw, "pdf")
            out_doc.insert_pdf(fitz.open(stream=pdf_bytes, filetype="pdf"))
    return out_doc.tobytes(garbage=4, deflate=True)
