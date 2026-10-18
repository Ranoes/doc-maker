#!/usr/bin/env python3
"""
doc_maker.py - Zero-dependency CLI engine to inspect and render .docx and .odt
documents using any existing .doc, .docx, or .odt file as a formatting template.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# XML Namespaces
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "xml": "http://schemas.openxmlformats.org/XML/1998/namespace",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
    "loext": "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
    "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
}

for prefix, uri in NS.items():
    if prefix != "xml":
        ET.register_namespace(prefix, uri)

W_P = f"{{{NS['w']}}}p"
W_R = f"{{{NS['w']}}}r"
W_T = f"{{{NS['w']}}}t"
W_TAB = f"{{{NS['w']}}}tab"
W_BR = f"{{{NS['w']}}}br"
W_PPR = f"{{{NS['w']}}}pPr"
W_RPR = f"{{{NS['w']}}}rPr"
W_PSTYLE = f"{{{NS['w']}}}pStyle"
W_NUMPR = f"{{{NS['w']}}}numPr"
W_ILVL = f"{{{NS['w']}}}ilvl"
W_NUMID = f"{{{NS['w']}}}numId"
W_JC = f"{{{NS['w']}}}jc"
W_TBL = f"{{{NS['w']}}}tbl"
W_TBLGRID = f"{{{NS['w']}}}tblGrid"
W_GRIDCOL = f"{{{NS['w']}}}gridCol"
W_TR = f"{{{NS['w']}}}tr"
W_TC = f"{{{NS['w']}}}tc"
W_TCPR = f"{{{NS['w']}}}tcPr"
W_TCW = f"{{{NS['w']}}}tcW"
W_GRIDSPAN = f"{{{NS['w']}}}gridSpan"
W_BODY = f"{{{NS['w']}}}body"
W_SECTPR = f"{{{NS['w']}}}sectPr"
W_VAL = f"{{{NS['w']}}}val"
W_W = f"{{{NS['w']}}}w"
W_TYPE = f"{{{NS['w']}}}type"
W_B = f"{{{NS['w']}}}b"
W_BCS = f"{{{NS['w']}}}bCs"
W_I = f"{{{NS['w']}}}i"
W_ICS = f"{{{NS['w']}}}iCs"
W_U = f"{{{NS['w']}}}u"
W_SZ = f"{{{NS['w']}}}sz"
W_SZCS = f"{{{NS['w']}}}szCs"
W_RFONTS = f"{{{NS['w']}}}rFonts"
W_COLOR = f"{{{NS['w']}}}color"
W_DRAWING = f"{{{NS['w']}}}drawing"
W_PICT = f"{{{NS['w']}}}pict"
W_BOOKMARK_START = f"{{{NS['w']}}}bookmarkStart"
W_BOOKMARK_END = f"{{{NS['w']}}}bookmarkEnd"
W_UPDATE_FIELDS = f"{{{NS['w']}}}updateFields"
XML_SPACE = f"{{{NS['xml']}}}space"
XML_ID = f"{{{NS['xml']}}}id"

ODF_P = f"{{{NS['text']}}}p"
ODF_H = f"{{{NS['text']}}}h"
ODF_SPAN = f"{{{NS['text']}}}span"
ODF_A = f"{{{NS['text']}}}a"
ODF_S = f"{{{NS['text']}}}s"
ODF_TAB = f"{{{NS['text']}}}tab"
ODF_LINE_BREAK = f"{{{NS['text']}}}line-break"
ODF_SOFT_PAGE_BREAK = f"{{{NS['text']}}}soft-page-break"
ODF_BOOKMARK = f"{{{NS['text']}}}bookmark"
ODF_BOOKMARK_START = f"{{{NS['text']}}}bookmark-start"
ODF_BOOKMARK_END = f"{{{NS['text']}}}bookmark-end"
ODF_LIST = f"{{{NS['text']}}}list"
ODF_LIST_ITEM = f"{{{NS['text']}}}list-item"
ODF_STYLE_NAME = f"{{{NS['text']}}}style-name"
ODF_CONTINUE_NUMBERING = f"{{{NS['text']}}}continue-numbering"
ODF_TOC = f"{{{NS['text']}}}table-of-content"
ODF_INDEX_BODY = f"{{{NS['text']}}}index-body"
ODF_TABLE = f"{{{NS['table']}}}table"
ODF_TABLE_COL = f"{{{NS['table']}}}table-column"
ODF_ROW = f"{{{NS['table']}}}table-row"
ODF_CELL = f"{{{NS['table']}}}table-cell"
ODF_COVERED_CELL = f"{{{NS['table']}}}covered-table-cell"
ODF_TABLE_STYLE_NAME = f"{{{NS['table']}}}style-name"
ODF_TABLE_NAME = f"{{{NS['table']}}}name"
ODF_NUM_COLS_SPANNED = f"{{{NS['table']}}}number-columns-spanned"
ODF_OFFICE_TEXT = f"{{{NS['office']}}}text"
ODF_OFFICE_AUTO_STYLES = f"{{{NS['office']}}}automatic-styles"
ODF_OFFICE_VALUE_TYPE = f"{{{NS['office']}}}value-type"
ODF_STYLE = f"{{{NS['style']}}}style"
ODF_STYLE_NAME_ATTR = f"{{{NS['style']}}}name"
ODF_STYLE_FAMILY_ATTR = f"{{{NS['style']}}}family"
ODF_STYLE_PARENT_ATTR = f"{{{NS['style']}}}parent-style-name"
ODF_TEXT_PROPS = f"{{{NS['style']}}}text-properties"
ODF_PARA_PROPS = f"{{{NS['style']}}}paragraph-properties"
ODF_TABLE_PROPS = f"{{{NS['style']}}}table-properties"
ODF_TABLE_COL_PROPS = f"{{{NS['style']}}}table-column-properties"
ODF_COL_WIDTH = f"{{{NS['style']}}}column-width"
ODF_DRAW_FRAME = f"{{{NS['draw']}}}frame"
ODF_DRAW_IMAGE = f"{{{NS['draw']}}}image"
ODF_SVG_DESC = f"{{{NS['svg']}}}desc"
ODF_SVG_TITLE = f"{{{NS['svg']}}}title"

PLACEHOLDER_RE = re.compile(r"(\{\{[^{}\n]+\}\}|\[\[[A-Za-z0-9_.\-\s]+\]\]|<[A-Za-z0-9_.\-]+>)")
MD_IMAGE_RE = re.compile(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$")

IDENTITY_KEYWORDS = {
    "name", "full name", "student name", "author", "prepared by", "submitted by",
    "nama", "nama lengkap", "nama mahasiswa", "nama siswa", "penulis", "penyusun",
    "nim", "npm", "nrp", "nis", "nisn", "nik", "nip", "id", "student id", "employee id", "no induk", "nomor induk",
    "age", "umur", "usia",
    "ttl", "tempat tanggal lahir", "date of birth", "dob", "birth date",
    "gender", "sex", "jenis kelamin",
    "class", "kelas", "semester", "angkatan", "batch", "group", "kelompok",
    "department", "major", "program studi", "prodi", "jurusan", "faculty", "fakultas",
    "university", "universitas", "institution", "instansi", "nama instansi", "school", "sekolah", "company", "perusahaan",
    "course", "subject", "mata kuliah", "matkul", "pelajaran",
    "lecturer", "instructor", "supervisor", "advisor", "dosen", "dosen pengampu", "pembimbing", "guru",
    "date", "tanggal", "hari/tanggal", "hari / tanggal",
    "email", "e-mail", "phone", "no hp", "no. hp", "no telp", "telepon", "whatsapp",
    "address", "alamat", "role", "jabatan", "position",
}

LABEL_VALUE_RE = re.compile(
    r"^\s*([A-Za-z][A-Za-z0-9\s./()\-_]{1,35}?)\s*(?:\t+|\s*:\s*|\t+:\s*)(.+?)\s*$"
)


def find_soffice() -> str | None:
    """Locate LibreOffice / soffice executable across Windows, macOS, and Linux."""
    for cmd in ("soffice.com", "soffice", "libreoffice"):
        found = shutil.which(cmd)
        if found:
            return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def convert_with_soffice(input_path: Path, target_ext: str, out_dir: Path) -> Path:
    """Convert a document using LibreOffice headless mode."""
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            f"LibreOffice (soffice) is required to convert '{input_path.suffix}' to '.{target_ext}', "
            "but was not found on PATH or standard install locations."
        )
    target_ext = target_ext.lstrip(".").lower()
    with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile_dir:
        profile_uri = Path(profile_dir).as_uri()
        cmd = [
            soffice,
            f"-env:UserInstallation={profile_uri}",
            "--headless",
            "--convert-to",
            target_ext,
            "--outdir",
            str(out_dir),
            str(input_path),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, check=False)
        expected = out_dir / f"{input_path.stem}.{target_ext}"
        if res.returncode != 0 or not expected.exists():
            raise RuntimeError(
                f"LibreOffice conversion failed (exit {res.returncode}): {res.stderr or res.stdout}"
            )
        return expected


def get_image_meta(img_path: Path) -> tuple[bytes, str, str, int, int]:
    """Return (bytes, mime_type, ext, width_px, height_px) by inspecting binary headers."""
    data = img_path.read_bytes()
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        w, h = struct.unpack(">II", data[16:24])
        return data, "image/png", "png", max(w, 1), max(h, 1)
    if data[:2] == b"\xff\xd8":
        idx = 2
        while idx < len(data):
            while idx < len(data) and data[idx] != 0xFF:
                idx += 1
            while idx < len(data) and data[idx] == 0xFF:
                idx += 1
            if idx + 1 >= len(data):
                break
            marker = data[idx]
            idx += 1
            if marker in (0xD8, 0xD9) or (0xD0 <= marker <= 0xD7):
                continue
            if idx + 2 > len(data):
                break
            length = struct.unpack(">H", data[idx : idx + 2])[0]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3) and idx + 7 <= len(data):
                h, w = struct.unpack(">HH", data[idx + 3 : idx + 7])
                return data, "image/jpeg", "jpg", max(w, 1), max(h, 1)
            idx += length
    ext = img_path.suffix.lstrip(".").lower() or "png"
    mime = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
    return data, mime, ext, 800, 500


def compute_display_dims_in(
    w_px: int, h_px: int, max_width_in: float = 5.4, max_height_in: float = 4.2
) -> tuple[float, float]:
    ratio = w_px / max(h_px, 1)
    w_in = max_width_in
    h_in = w_in / ratio
    if h_in > max_height_in:
        h_in = max_height_in
        w_in = h_in * ratio
    return round(w_in, 4), round(h_in, 4)


def normalize_label(text: str) -> str:
    cleaned = re.sub(r"^[\d.\-*\s()]+", "", text.strip().lower())
    cleaned = re.sub(r"\s*:\s*$", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned


def is_identity_label(label: str) -> bool:
    norm = normalize_label(label)
    if not norm or len(norm) > 40:
        return False
    if norm in IDENTITY_KEYWORDS:
        return True
    for kw in IDENTITY_KEYWORDS:
        if len(kw) >= 4 and (norm.startswith(kw + " ") or norm.endswith(" " + kw)):
            return True
    return False


def detect_identity_in_block(block_id: str, text: str) -> dict | None:
    m = LABEL_VALUE_RE.match(text)
    if not m:
        m_empty = re.match(r"^\s*([A-Za-z][A-Za-z0-9\s./()\-_]{1,35}?)\s*:\s*$", text)
        if m_empty and is_identity_label(m_empty.group(1)):
            return {
                "id": block_id,
                "label": m_empty.group(1).strip(),
                "current_value": "",
                "raw_text": text,
            }
        return None
    label, val = m.group(1).strip(), m.group(2).strip()
    if is_identity_label(label):
        return {
            "id": block_id,
            "label": label,
            "current_value": val,
            "raw_text": text,
        }
    return None


# ============================================================================
# DOCX (OOXML) Inspection & Manipulation
# ============================================================================

def _docx_prop_is_on(rpr: ET.Element, tag: str) -> bool:
    el = rpr.find(tag)
    if el is None:
        return False
    val = el.get(W_VAL)
    if val is None:
        return True
    return val.lower() not in ("false", "0", "none")


def _docx_set_prop_on(rpr: ET.Element, tag: str) -> None:
    el = rpr.find(tag)
    if el is None:
        ET.SubElement(rpr, tag)
    else:
        if W_VAL in el.attrib:
            del el.attrib[W_VAL]


def docx_get_paragraph_text(p_elem: ET.Element) -> str:
    parts: list[str] = []
    for elem in p_elem.iter():
        if elem.tag == W_T and elem.text:
            parts.append(elem.text)
        elif elem.tag == W_TAB:
            parts.append("\t")
        elif elem.tag == W_BR:
            parts.append("\n")
    return "".join(parts)


def docx_get_paragraph_style_info(p_elem: ET.Element) -> dict:
    style_name = "Normal"
    is_list = False
    ppr = p_elem.find(W_PPR)
    if ppr is not None:
        pstyle = ppr.find(W_PSTYLE)
        if pstyle is not None and pstyle.get(W_VAL):
            style_name = pstyle.get(W_VAL)
        if ppr.find(W_NUMPR) is not None:
            is_list = True
    formatting: list[str] = []
    first_r = p_elem.find(W_R)
    if first_r is not None:
        rpr = first_r.find(W_RPR)
        if rpr is not None:
            if _docx_prop_is_on(rpr, W_B):
                formatting.append("bold")
            if _docx_prop_is_on(rpr, W_I):
                formatting.append("italic")
            if _docx_prop_is_on(rpr, W_U):
                formatting.append("underline")
            sz = rpr.find(W_SZ)
            if sz is not None and sz.get(W_VAL):
                try:
                    formatting.append(f"{int(sz.get(W_VAL)) / 2:g}pt")
                except ValueError:
                    pass
            color = rpr.find(W_COLOR)
            if color is not None and color.get(W_VAL) and color.get(W_VAL) != "auto":
                formatting.append(f"#{color.get(W_VAL)}")
    return {"style": style_name, "is_list": is_list, "formatting": formatting}


def docx_set_paragraph_text(p_elem: ET.Element, new_text: str) -> None:
    """Replace text inside a <w:p> while preserving <w:pPr>, bookmarks, drawings, and first run's <w:rPr>."""
    ppr = p_elem.find(W_PPR)
    first_r = p_elem.find(W_R)
    base_rpr = copy.deepcopy(first_r.find(W_RPR)) if (first_r is not None and first_r.find(W_RPR) is not None) else None
    # Strip red guidance color (#ee0000 / #ff0000) if present in template table cells
    if base_rpr is not None:
        col_el = base_rpr.find(W_COLOR)
        if col_el is not None and (col_el.get(W_VAL) or "").lower() in ("ee0000", "ff0000"):
            col_el.set(W_VAL, "000000")

    preserved_drawings = [copy.deepcopy(d) for d in p_elem.findall(f".//{W_DRAWING}")]
    preserved_bookmarks = [
        copy.deepcopy(ch)
        for ch in list(p_elem)
        if ch.tag in (W_BOOKMARK_START, W_BOOKMARK_END)
    ]

    for child in list(p_elem):
        if child is not ppr:
            p_elem.remove(child)

    for bm in preserved_bookmarks:
        p_elem.append(bm)
    for drw in preserved_drawings:
        r_drw = ET.SubElement(p_elem, W_R)
        r_drw.append(drw)

    tokens = re.split(r"(\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)", new_text)
    for token in tokens:
        if not token:
            continue
        is_bi = token.startswith("***") and token.endswith("***") and len(token) > 6
        is_bold = is_bi or (token.startswith("**") and token.endswith("**") and len(token) > 4)
        is_italic = is_bi or (not is_bold and token.startswith("*") and token.endswith("*") and len(token) > 2)
        is_code = token.startswith("`") and token.endswith("`") and len(token) > 2
        if is_bi:
            content = token[3:-3]
        elif is_bold:
            content = token[2:-2]
        elif is_italic or is_code:
            content = token[1:-1]
        else:
            content = token

        r_elem = ET.SubElement(p_elem, W_R)
        rpr = copy.deepcopy(base_rpr) if base_rpr is not None else ET.Element(W_RPR)
        if is_bold:
            _docx_set_prop_on(rpr, W_B)
            _docx_set_prop_on(rpr, W_BCS)
        if is_italic:
            _docx_set_prop_on(rpr, W_I)
            _docx_set_prop_on(rpr, W_ICS)
        if is_code:
            rfonts = rpr.find(W_RFONTS)
            if rfonts is None:
                rfonts = ET.SubElement(rpr, W_RFONTS)
            rfonts.set(f"{{{NS['w']}}}ascii", "Courier New")
            rfonts.set(f"{{{NS['w']}}}hAnsi", "Courier New")
            rfonts.set(f"{{{NS['w']}}}cs", "Courier New")
        if len(rpr) > 0 or base_rpr is not None:
            r_elem.append(rpr)

        lines = content.split("\n")
        for line_idx, line in enumerate(lines):
            if line_idx > 0:
                ET.SubElement(r_elem, W_BR)
            tab_parts = line.split("\t")
            for tab_idx, part in enumerate(tab_parts):
                if tab_idx > 0:
                    ET.SubElement(r_elem, W_TAB)
                if part or (len(tab_parts) == 1 and len(lines) == 1):
                    t_elem = ET.SubElement(r_elem, W_T)
                    t_elem.set(XML_SPACE, "preserve")
                    t_elem.text = part


def docx_replace_placeholders_in_paragraph(p_elem: ET.Element, mapping: dict[str, str]) -> None:
    if not mapping:
        return
    full_text = docx_get_paragraph_text(p_elem)
    if not any(k in full_text for k in mapping):
        return

    t_elems = [elem for elem in p_elem.iter(W_T)]
    combined_t = "".join(t.text or "" for t in t_elems)
    if not any(k in combined_t for k in mapping):
        return

    for t in t_elems:
        if t.text:
            for k, v in mapping.items():
                if k in t.text:
                    t.text = t.text.replace(k, str(v))
                    t.set(XML_SPACE, "preserve")

    remaining_text = "".join(t.text or "" for t in t_elems)
    if any(k in remaining_text for k in mapping):
        updated_full = docx_get_paragraph_text(p_elem)
        for k, v in mapping.items():
            updated_full = updated_full.replace(k, str(v))
        docx_set_paragraph_text(p_elem, updated_full)


def inspect_docx(docx_path: Path) -> dict:
    blocks: list[dict] = []
    tables: list[dict] = []
    identity_fields: list[dict] = []
    placeholders: set[str] = set()

    with zipfile.ZipFile(docx_path, "r") as zf:
        xml_targets: list[tuple[str, str]] = [("body", "word/document.xml")]
        for name in sorted(zf.namelist()):
            if re.match(r"^word/header\d*\.xml$", name):
                xml_targets.append((Path(name).stem, name))
            elif re.match(r"^word/footer\d*\.xml$", name):
                xml_targets.append((Path(name).stem, name))

        b_counter = 0
        t_counter = 0

        for loc, xml_name in xml_targets:
            if xml_name not in zf.namelist():
                continue
            root = ET.fromstring(zf.read(xml_name))
            container = root.find(W_BODY) if loc == "body" else root
            if container is None:
                continue

            for child in list(container):
                if child.tag == W_P:
                    text = docx_get_paragraph_text(child)
                    for ph in PLACEHOLDER_RE.findall(text):
                        placeholders.add(ph)
                    block_id = f"b{b_counter}" if loc == "body" else f"{loc}_b{b_counter}"
                    b_counter += 1
                    info = docx_get_paragraph_style_info(child)
                    has_img = (child.find(f".//{W_DRAWING}") is not None) or (child.find(f".//{W_PICT}") is not None)
                    b_info = {
                        "id": block_id,
                        "location": loc,
                        "style": info["style"],
                        "is_list": info["is_list"],
                        "formatting": info["formatting"],
                        "text": text,
                    }
                    if has_img:
                        b_info["has_images"] = True
                    blocks.append(b_info)
                    id_field = detect_identity_in_block(block_id, text)
                    if id_field:
                        identity_fields.append(id_field)

                elif child.tag == W_TBL:
                    table_id = f"t{t_counter}" if loc == "body" else f"{loc}_t{t_counter}"
                    t_counter += 1
                    rows_data: list[list[str]] = []
                    for r_idx, tr in enumerate(child.findall(W_TR)):
                        row_cells: list[str] = []
                        for c_idx, tc in enumerate(tr.findall(W_TC)):
                            cell_paras = [docx_get_paragraph_text(p) for p in tc.findall(W_P)]
                            cell_text = "\n".join(cell_paras)
                            for ph in PLACEHOLDER_RE.findall(cell_text):
                                placeholders.add(ph)
                            row_cells.append(cell_text)
                        rows_data.append(row_cells)
                        if len(row_cells) >= 2 and is_identity_label(row_cells[0]):
                            val_col = 2 if (len(row_cells) >= 3 and row_cells[1].strip() == ":") else 1
                            identity_fields.append({
                                "id": f"{table_id}.r{r_idx}.c{val_col}",
                                "table_id": table_id,
                                "cell": f"{r_idx},{val_col}",
                                "label": row_cells[0].strip().rstrip(":").strip(),
                                "current_value": row_cells[val_col].strip(),
                                "raw_text": " | ".join(row_cells),
                            })
                    tables.append({
                        "id": table_id,
                        "location": loc,
                        "row_count": len(rows_data),
                        "col_count": max((len(r) for r in rows_data), default=0),
                        "rows": rows_data,
                    })

    return {
        "format": "docx",
        "source": str(docx_path),
        "placeholders": sorted(placeholders),
        "identity_fields": identity_fields,
        "blocks": blocks,
        "tables": tables,
    }


def _docx_populate_table_element(
    tbl_elem: ET.Element,
    new_rows: list[list[str]],
    col_widths_in: list[float] | None = None,
    total_width_in: float = 5.5,
) -> None:
    """Populate a DOCX <w:tbl> with new_rows, dynamically adjusting column count and widths."""
    if not new_rows:
        return
    num_cols = max(len(r) for r in new_rows)
    existing_trs = tbl_elem.findall(W_TR)
    if not existing_trs:
        return

    if col_widths_in and len(col_widths_in) == num_cols:
        widths_twips = [int(round(w * 1440)) for w in col_widths_in]
    else:
        col_weights: list[float] = []
        for c in range(num_cols):
            max_len = max((len(str(r[c])) for r in new_rows if c < len(r)), default=5)
            col_weights.append(max(min(max_len ** 0.55, 18.0), 2.2))
        total_w = sum(col_weights) or 1.0
        widths_twips = [int(round(total_width_in * 1440 * (w / total_w))) for w in col_weights]

    tbl_grid = tbl_elem.find(W_TBLGRID)
    if tbl_grid is not None:
        for gc in list(tbl_grid.findall(W_GRIDCOL)):
            tbl_grid.remove(gc)
        for tw in widths_twips:
            ET.SubElement(tbl_grid, W_GRIDCOL, {W_W: str(tw)})

    header_tr = existing_trs[0]
    template_tr = copy.deepcopy(existing_trs[1] if len(existing_trs) > 1 else existing_trs[0])

    def _adjust_tr_cells(tr: ET.Element) -> list[ET.Element]:
        tcs = tr.findall(W_TC)
        for tc in tcs:
            tcpr = tc.find(W_TCPR)
            if tcpr is not None:
                gs = tcpr.find(W_GRIDSPAN)
                if gs is not None:
                    tcpr.remove(gs)
        while len(tcs) < num_cols:
            tr.append(copy.deepcopy(tcs[-1]))
            tcs = tr.findall(W_TC)
        while len(tcs) > num_cols:
            tr.remove(tcs[-1])
            tcs = tr.findall(W_TC)
        for c_i, tc in enumerate(tcs):
            tcpr = tc.find(W_TCPR)
            if tcpr is not None:
                tcw = tcpr.find(W_TCW)
                if tcw is not None and c_i < len(widths_twips):
                    tcw.set(W_W, str(widths_twips[c_i]))
                    tcw.set(W_TYPE, "dxa")
        return tcs

    def _set_cell_text(tc: ET.Element, text_val: str) -> None:
        p_list = tc.findall(W_P)
        base_p = p_list[0] if p_list else ET.SubElement(tc, W_P)
        for extra_p in p_list[1:]:
            tc.remove(extra_p)
        paragraphs = str(text_val).split("\n\n")
        docx_set_paragraph_text(base_p, paragraphs[0])
        for extra_txt in paragraphs[1:]:
            extra_p = copy.deepcopy(base_p)
            docx_set_paragraph_text(extra_p, extra_txt)
            tc.append(extra_p)

    hdr_tcs = _adjust_tr_cells(header_tr)
    for c_idx, tc in enumerate(hdr_tcs):
        val = str(new_rows[0][c_idx]) if c_idx < len(new_rows[0]) else ""
        _set_cell_text(tc, val)

    for old_tr in existing_trs[1:]:
        tbl_elem.remove(old_tr)

    for row_vals in new_rows[1:]:
        new_tr = copy.deepcopy(template_tr)
        data_tcs = _adjust_tr_cells(new_tr)
        for c_idx, tc in enumerate(data_tcs):
            val = str(row_vals[c_idx]) if c_idx < len(row_vals) else ""
            _set_cell_text(tc, val)
        tbl_elem.append(new_tr)


def _docx_build_inline_drawing(rel_id: str, doc_pr_id: int, name: str, cx_emu: int, cy_emu: int) -> ET.Element:
    """Create a native OOXML <w:drawing><wp:inline> element for an embedded image."""
    drawing_xml = f"""<w:drawing xmlns:w="{NS['w']}" xmlns:wp="{NS['wp']}" xmlns:a="{NS['a']}" xmlns:pic="{NS['pic']}" xmlns:r="{NS['r']}">
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="{cx_emu}" cy="{cy_emu}"/>
        <wp:docPr id="{doc_pr_id}" name="{name}"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="{doc_pr_id}" name="{name}"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="{rel_id}"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="{cx_emu}" cy="{cy_emu}"/>
                </a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>"""
    return ET.fromstring(drawing_xml)


def _docx_build_rich_elements(
    item: str | dict,
    template_p: ET.Element,
    table_templates: dict[str, ET.Element],
    added_media: dict[str, tuple[bytes, str, str]],
    counter_state: dict[str, int],
) -> list[ET.Element]:
    """Convert a string or rich dict item into one or more OOXML <w:p> / <w:tbl> elements."""
    if isinstance(item, str):
        m_img = MD_IMAGE_RE.match(item.strip())
        if m_img:
            item = {"type": "image", "caption": m_img.group(1), "path": m_img.group(2)}
        else:
            cloned_p = copy.deepcopy(template_p)
            docx_set_paragraph_text(cloned_p, item)
            return [cloned_p]

    itype = item.get("type", "p")

    if itype in ("p", "heading"):
        p_el = copy.deepcopy(template_p)
        if itype == "heading":
            lvl = int(item.get("level", 2))
            ppr = p_el.find(W_PPR)
            if ppr is None:
                ppr = ET.SubElement(p_el, W_PPR)
            pstyle = ppr.find(W_PSTYLE)
            if pstyle is None:
                pstyle = ET.SubElement(ppr, W_PSTYLE)
            pstyle.set(W_VAL, f"Heading{lvl}")
        docx_set_paragraph_text(p_el, str(item.get("text", "")))
        return [p_el]

    if itype == "list":
        out_ps: list[ET.Element] = []
        for idx_i, sub_txt in enumerate(item.get("items", []), start=1):
            p_el = copy.deepcopy(template_p)
            docx_set_paragraph_text(p_el, f"{idx_i}. {sub_txt}")
            out_ps.append(p_el)
        return out_ps

    if itype == "image":
        img_path = Path(item["path"])
        data, mime, ext, w_px, h_px = get_image_meta(img_path)
        max_w = float(item.get("max_width_in", 5.4))
        max_h = float(item.get("max_height_in", 4.2))
        w_in, h_in = compute_display_dims_in(w_px, h_px, max_w, max_h)
        if "width_in" in item and "height_in" in item:
            w_in, h_in = float(item["width_in"]), float(item["height_in"])
        counter_state["img_id"] += 1
        img_num = counter_state["img_id"]
        rel_id = f"rIdDMImg{img_num}"
        media_rel_path = f"media/dm_image_{img_num}.{ext}"
        added_media[rel_id] = (data, mime, media_rel_path)

        cx_emu = int(round(w_in * 914400))
        cy_emu = int(round(h_in * 914400))
        p_img = ET.Element(W_P)
        ppr = ET.SubElement(p_img, W_PPR)
        ET.SubElement(ppr, W_JC, {W_VAL: "center"})
        r_img = ET.SubElement(p_img, W_R)
        r_img.append(_docx_build_inline_drawing(rel_id, img_num, f"Picture {img_num}", cx_emu, cy_emu))
        out_elems = [p_img]

        caption = item.get("caption")
        if caption:
            p_cap = copy.deepcopy(template_p)
            ppr_cap = p_cap.find(W_PPR)
            if ppr_cap is None:
                ppr_cap = ET.SubElement(p_cap, W_PPR)
            num_pr = ppr_cap.find(W_NUMPR)
            if num_pr is not None:
                ppr_cap.remove(num_pr)
            jc = ppr_cap.find(W_JC)
            if jc is None:
                jc = ET.SubElement(ppr_cap, W_JC)
            jc.set(W_VAL, "center")
            docx_set_paragraph_text(p_cap, f"**{caption}**" if not str(caption).startswith("**") else str(caption))
            out_elems.append(p_cap)
        return out_elems

    if itype == "table":
        out_elems = []
        caption = item.get("caption")
        if caption:
            p_cap = copy.deepcopy(template_p)
            ppr_cap = p_cap.find(W_PPR)
            if ppr_cap is None:
                ppr_cap = ET.SubElement(p_cap, W_PPR)
            num_pr = ppr_cap.find(W_NUMPR)
            if num_pr is not None:
                ppr_cap.remove(num_pr)
            jc = ppr_cap.find(W_JC)
            if jc is None:
                jc = ET.SubElement(ppr_cap, W_JC)
            jc.set(W_VAL, "center")
            docx_set_paragraph_text(p_cap, f"**{caption}**" if not str(caption).startswith("**") else str(caption))
            out_elems.append(p_cap)
        tpl_id = item.get("template_table", "t6")
        base_tbl = table_templates.get(tpl_id)
        if base_tbl is None and table_templates:
            base_tbl = next(iter(table_templates.values()))
        if base_tbl is not None:
            new_tbl = copy.deepcopy(base_tbl)
            _docx_populate_table_element(
                new_tbl,
                item.get("rows", []),
                col_widths_in=item.get("col_widths"),
                total_width_in=float(item.get("total_width_in", 5.5)),
            )
            out_elems.append(new_tbl)
        return out_elems

    if itype == "code_table":
        title = str(item.get("title", "Kode Sumber"))
        code_lines = str(item.get("code", "")).strip("\n").split("\n")
        nums_str = "\n".join(str(i) for i in range(1, len(code_lines) + 1))
        code_str = "\n".join(f"`{ln}`" if ln.strip() else " " for ln in code_lines)
        base_tbl = table_templates.get(item.get("template_table", "t5"))
        if base_tbl is None and table_templates:
            base_tbl = next(iter(table_templates.values()))
        if base_tbl is not None:
            new_tbl = copy.deepcopy(base_tbl)
            _docx_populate_table_element(
                new_tbl,
                [[f"**{title}**", ""], [nums_str, code_str]],
                col_widths_in=[0.45, 5.05],
            )
            return [new_tbl]

    cloned_p = copy.deepcopy(template_p)
    docx_set_paragraph_text(cloned_p, str(item.get("text", "")))
    return [cloned_p]


def ensure_docx_update_fields(docx_file: Path) -> None:
    """Inject <w:updateFields w:val="true"/> into word/settings.xml inside a .docx archive."""
    with zipfile.ZipFile(docx_file, "r") as zin:
        namelist = zin.namelist()
        if "word/settings.xml" in namelist:
            s_root = ET.fromstring(zin.read("word/settings.xml"))
            if s_root.find(W_UPDATE_FIELDS) is None:
                ET.SubElement(s_root, W_UPDATE_FIELDS, {W_VAL: "true"})
            settings_bytes = ET.tostring(s_root, encoding="utf-8", xml_declaration=True)
        else:
            settings_bytes = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:settings xmlns:w="{NS["w"]}"><w:updateFields w:val="true"/></w:settings>'
            ).encode("utf-8")
        all_files = {name: zin.read(name) for name in namelist}
    all_files["word/settings.xml"] = settings_bytes
    with zipfile.ZipFile(docx_file, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in all_files.items():
            zout.writestr(name, content)


def render_docx(docx_path: Path, spec: dict, out_docx_path: Path) -> None:
    placeholders: dict[str, str] = spec.get("placeholders", {})
    block_ops: dict[str, dict] = {op["id"]: op for op in spec.get("blocks", []) if "id" in op}
    table_ops: dict[str, dict] = {op["id"]: op for op in spec.get("tables", []) if "id" in op}

    for op in spec.get("blocks", []):
        bid = op.get("id", "")
        m = re.match(r"^(t\d+)\.r(\d+)\.c(\d+)$", bid)
        if m and "text" in op:
            tid, r_idx, c_idx = m.group(1), m.group(2), m.group(3)
            t_op = table_ops.setdefault(tid, {"id": tid})
            cells = t_op.setdefault("cells", {})
            cells[f"{r_idx},{c_idx}"] = op["text"]

    added_media: dict[str, tuple[bytes, str, str]] = {}
    counter_state = {"img_id": 500}

    with zipfile.ZipFile(docx_path, "r") as zin:
        xml_targets: dict[str, str] = {"word/document.xml": "body"}
        for name in sorted(zin.namelist()):
            if re.match(r"^word/header\d*\.xml$", name) or re.match(r"^word/footer\d*\.xml$", name):
                xml_targets[name] = Path(name).stem

        b_counter = 0
        t_counter = 0
        modified_xmls: dict[str, bytes] = {}

        ordered_names = ["word/document.xml"] + [
            n for n in sorted(zin.namelist())
            if re.match(r"^word/(header|footer)\d*\.xml$", n)
        ]

        for xml_name in ordered_names:
            if xml_name not in zin.namelist():
                continue
            loc = xml_targets[xml_name]
            root = ET.fromstring(zin.read(xml_name))
            container = root.find(W_BODY) if loc == "body" else root
            if container is None:
                continue

            table_templates: dict[str, ET.Element] = {}
            if loc == "body":
                t_snap = 0
                for ch in list(container):
                    if ch.tag == W_TBL:
                        table_templates[f"t{t_snap}"] = copy.deepcopy(ch)
                        t_snap += 1

            children_snapshot = list(container)
            for child in children_snapshot:
                if child.tag == W_P:
                    block_id = f"b{b_counter}" if loc == "body" else f"{loc}_b{b_counter}"
                    b_counter += 1
                    op = block_ops.get(block_id)
                    if op:
                        if op.get("delete"):
                            container.remove(child)
                            continue
                        if "items" in op and isinstance(op["items"], list):
                            items = op["items"]
                            if not items:
                                container.remove(child)
                                continue
                            idx = list(container).index(child)
                            container.remove(child)
                            insert_offset = 0
                            for item in items:
                                built = _docx_build_rich_elements(
                                    item, child, table_templates, added_media, counter_state
                                )
                                for el in built:
                                    container.insert(idx + insert_offset, el)
                                    insert_offset += 1
                            continue
                        elif "text" in op:
                            docx_set_paragraph_text(child, str(op["text"]))

                    if placeholders:
                        docx_replace_placeholders_in_paragraph(child, placeholders)

                elif child.tag == W_TBL:
                    table_id = f"t{t_counter}" if loc == "body" else f"{loc}_t{t_counter}"
                    t_counter += 1
                    t_op = table_ops.get(table_id)
                    if t_op:
                        if t_op.get("delete"):
                            container.remove(child)
                            continue
                        if "rows" in t_op and isinstance(t_op["rows"], list):
                            _docx_populate_table_element(
                                child,
                                t_op["rows"],
                                col_widths_in=t_op.get("col_widths"),
                                total_width_in=float(t_op.get("total_width_in", 5.5)),
                            )
                        if "cells" in t_op and isinstance(t_op["cells"], dict):
                            existing_trs = child.findall(W_TR)
                            for rc_key, val in t_op["cells"].items():
                                r_str, c_str = rc_key.split(",")
                                r_i, c_i = int(r_str.strip()), int(c_str.strip())
                                if 0 <= r_i < len(existing_trs):
                                    tcs = existing_trs[r_i].findall(W_TC)
                                    if 0 <= c_i < len(tcs):
                                        p_list = tcs[c_i].findall(W_P)
                                        if p_list:
                                            docx_set_paragraph_text(p_list[0], str(val))
                                            for extra_p in p_list[1:]:
                                                tcs[c_i].remove(extra_p)

                    if placeholders:
                        for p_in_tbl in child.iter(W_P):
                            docx_replace_placeholders_in_paragraph(p_in_tbl, placeholders)

            modified_xmls[xml_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        # If images were added in DOCX, update word/_rels/document.xml.rels and [Content_Types].xml
        if added_media:
            rels_path = "word/_rels/document.xml.rels"
            rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
            ET.register_namespace("", rels_ns)
            if rels_path in zin.namelist():
                rels_root = ET.fromstring(zin.read(rels_path))
            else:
                rels_root = ET.Element(f"{{{rels_ns}}}Relationships")
            for rel_id, (_, _, target_media) in added_media.items():
                ET.SubElement(rels_root, f"{{{rels_ns}}}Relationship", {
                    "Id": rel_id,
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                    "Target": target_media,
                })
            modified_xmls[rels_path] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

            ct_path = "[Content_Types].xml"
            if ct_path in zin.namelist():
                ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
                ET.register_namespace("", ct_ns)
                ct_root = ET.fromstring(zin.read(ct_path))
                existing_exts = {
                    (d.get("Extension") or "").lower()
                    for d in ct_root.findall(f"{{{ct_ns}}}Default")
                }
                for _, (_, mime, target_media) in added_media.items():
                    ext = Path(target_media).suffix.lstrip(".").lower()
                    if ext and ext not in existing_exts:
                        ET.SubElement(ct_root, f"{{{ct_ns}}}Default", {
                            "Extension": ext,
                            "ContentType": mime,
                        })
                        existing_exts.add(ext)
                modified_xmls[ct_path] = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(out_docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            written_names = set()
            for item in zin.infolist():
                if item.filename in modified_xmls:
                    zout.writestr(item, modified_xmls[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))
                written_names.add(item.filename)
            for extra_xml, xml_bytes in modified_xmls.items():
                if extra_xml not in written_names:
                    zout.writestr(extra_xml, xml_bytes)
            for _, (data, _, target_media) in added_media.items():
                zout.writestr(f"word/{target_media}", data)

    ensure_docx_update_fields(out_docx_path)


# ============================================================================
# ODT (OpenDocument Text) Inspection & Manipulation
# ============================================================================

def odt_get_paragraph_text(p_elem: ET.Element) -> str:
    parts: list[str] = []

    def walk(node: ET.Element) -> None:
        if node.tag in (ODF_SVG_DESC, ODF_SVG_TITLE):
            return
        if node.text:
            parts.append(node.text)
        for ch in node:
            if ch.tag in (ODF_SVG_DESC, ODF_SVG_TITLE):
                if ch.tail:
                    parts.append(ch.tail)
                continue
            if ch.tag == ODF_TAB:
                parts.append("\t")
            elif ch.tag == ODF_LINE_BREAK:
                parts.append("\n")
            elif ch.tag == ODF_S:
                count = int(ch.get(f"{{{NS['text']}}}c", "1"))
                parts.append(" " * count)
            else:
                walk(ch)
            if ch.tail:
                parts.append(ch.tail)

    walk(p_elem)
    return "".join(parts)


def _odt_ensure_inline_styles(auto_styles: ET.Element) -> None:
    """Ensure helper text styles exist in <office:automatic-styles> for bold, italic, code, and normal."""
    existing = {
        st.get(ODF_STYLE_NAME_ATTR)
        for st in auto_styles.findall(ODF_STYLE)
    }
    defs = [
        ("DM_T_NORM12", "text", {
            f"{{{NS['style']}}}font-name": "Calibri",
            f"{{{NS['fo']}}}font-size": "12pt",
            f"{{{NS['fo']}}}color": "#000000",
            f"{{{NS['fo']}}}font-weight": "normal",
            f"{{{NS['fo']}}}font-style": "normal",
        }),
        ("DM_T_BOLD12", "text", {
            f"{{{NS['style']}}}font-name": "Calibri",
            f"{{{NS['fo']}}}font-size": "12pt",
            f"{{{NS['fo']}}}color": "#000000",
            f"{{{NS['fo']}}}font-weight": "bold",
            f"{{{NS['style']}}}font-weight-asian": "bold",
            f"{{{NS['style']}}}font-weight-complex": "bold",
        }),
        ("DM_T_ITALIC12", "text", {
            f"{{{NS['style']}}}font-name": "Calibri",
            f"{{{NS['fo']}}}font-size": "12pt",
            f"{{{NS['fo']}}}color": "#000000",
            f"{{{NS['fo']}}}font-style": "italic",
            f"{{{NS['style']}}}font-style-asian": "italic",
            f"{{{NS['style']}}}font-style-complex": "italic",
        }),
        ("DM_T_BI12", "text", {
            f"{{{NS['style']}}}font-name": "Calibri",
            f"{{{NS['fo']}}}font-size": "12pt",
            f"{{{NS['fo']}}}color": "#000000",
            f"{{{NS['fo']}}}font-weight": "bold",
            f"{{{NS['fo']}}}font-style": "italic",
            f"{{{NS['style']}}}font-weight-asian": "bold",
            f"{{{NS['style']}}}font-style-asian": "italic",
        }),
        ("DM_T_CODE11", "text", {
            f"{{{NS['style']}}}font-name": "Courier New",
            f"{{{NS['fo']}}}font-size": "10.5pt",
            f"{{{NS['fo']}}}color": "#000000",
        }),
        ("DM_T_TBL_NORM10", "text", {
            f"{{{NS['style']}}}font-name": "Calibri",
            f"{{{NS['fo']}}}font-size": "10pt",
            f"{{{NS['fo']}}}color": "#000000",
            f"{{{NS['fo']}}}font-weight": "normal",
            f"{{{NS['fo']}}}font-style": "normal",
        }),
        ("DM_T_TBL_BOLD10", "text", {
            f"{{{NS['style']}}}font-name": "Calibri",
            f"{{{NS['fo']}}}font-size": "10pt",
            f"{{{NS['fo']}}}color": "#000000",
            f"{{{NS['fo']}}}font-weight": "bold",
            f"{{{NS['style']}}}font-weight-asian": "bold",
            f"{{{NS['style']}}}font-weight-complex": "bold",
        }),
        ("DM_T_TBL_ITALIC10", "text", {
            f"{{{NS['style']}}}font-name": "Calibri",
            f"{{{NS['fo']}}}font-size": "10pt",
            f"{{{NS['fo']}}}color": "#000000",
            f"{{{NS['fo']}}}font-style": "italic",
            f"{{{NS['style']}}}font-style-asian": "italic",
        }),
        ("DM_T_TBL_BI10", "text", {
            f"{{{NS['style']}}}font-name": "Calibri",
            f"{{{NS['fo']}}}font-size": "10pt",
            f"{{{NS['fo']}}}color": "#000000",
            f"{{{NS['fo']}}}font-weight": "bold",
            f"{{{NS['fo']}}}font-style": "italic",
        }),
        ("DM_T_TBL_CODE9", "text", {
            f"{{{NS['style']}}}font-name": "Courier New",
            f"{{{NS['fo']}}}font-size": "9pt",
            f"{{{NS['fo']}}}color": "#000000",
        }),
        ("DM_T_TBL_NORM85", "text", {
            f"{{{NS['style']}}}font-name": "Calibri",
            f"{{{NS['fo']}}}font-size": "8.5pt",
            f"{{{NS['fo']}}}color": "#000000",
            f"{{{NS['fo']}}}font-weight": "normal",
            f"{{{NS['fo']}}}font-style": "normal",
        }),
        ("DM_T_TBL_BOLD85", "text", {
            f"{{{NS['style']}}}font-name": "Calibri",
            f"{{{NS['fo']}}}font-size": "8.5pt",
            f"{{{NS['fo']}}}color": "#000000",
            f"{{{NS['fo']}}}font-weight": "bold",
            f"{{{NS['style']}}}font-weight-asian": "bold",
            f"{{{NS['style']}}}font-weight-complex": "bold",
        }),
        ("DM_T_TBL_ITALIC85", "text", {
            f"{{{NS['style']}}}font-name": "Calibri",
            f"{{{NS['fo']}}}font-size": "8.5pt",
            f"{{{NS['fo']}}}color": "#000000",
            f"{{{NS['fo']}}}font-style": "italic",
            f"{{{NS['style']}}}font-style-asian": "italic",
        }),
        ("DM_T_TBL_CODE8", "text", {
            f"{{{NS['style']}}}font-name": "Courier New",
            f"{{{NS['fo']}}}font-size": "8pt",
            f"{{{NS['fo']}}}color": "#000000",
        }),
    ]
    for sname, fam, tprops in defs:
        if sname not in existing:
            st = ET.SubElement(auto_styles, ODF_STYLE, {
                ODF_STYLE_NAME_ATTR: sname,
                ODF_STYLE_FAMILY_ATTR: fam,
            })
            ET.SubElement(st, ODF_TEXT_PROPS, tprops)

    p_defs = [
        ("DM_P_TBL_HDR", {
            f"{{{NS['fo']}}}margin-top": "0.02in",
            f"{{{NS['fo']}}}margin-bottom": "0.02in",
            f"{{{NS['fo']}}}line-height": "100%",
            f"{{{NS['fo']}}}text-align": "center",
            f"{{{NS['style']}}}justify-single-word": "false",
        }),
        ("DM_P_TBL_CELL", {
            f"{{{NS['fo']}}}margin-top": "0.02in",
            f"{{{NS['fo']}}}margin-bottom": "0.03in",
            f"{{{NS['fo']}}}line-height": "105%",
            f"{{{NS['fo']}}}text-align": "start",
            f"{{{NS['style']}}}justify-single-word": "false",
        }),
        ("DM_P_CENTER", {
            f"{{{NS['fo']}}}margin-top": "0.06in",
            f"{{{NS['fo']}}}margin-bottom": "0.06in",
            f"{{{NS['fo']}}}text-align": "center",
            f"{{{NS['style']}}}justify-single-word": "false",
        }),
        ("DM_P_CODE", {
            f"{{{NS['fo']}}}margin-top": "0in",
            f"{{{NS['fo']}}}margin-bottom": "0in",
            f"{{{NS['fo']}}}line-height": "100%",
            f"{{{NS['fo']}}}text-align": "start",
        }),
    ]
    for sname, pprops in p_defs:
        if sname not in existing:
            st = ET.SubElement(auto_styles, ODF_STYLE, {
                ODF_STYLE_NAME_ATTR: sname,
                ODF_STYLE_FAMILY_ATTR: "paragraph",
                f"{{{NS['style']}}}parent-style-name": "Standard",
            })
            ET.SubElement(st, ODF_PARA_PROPS, pprops)

    if "DM_FR_IMG" not in existing:
        st_fr = ET.SubElement(auto_styles, ODF_STYLE, {
            ODF_STYLE_NAME_ATTR: "DM_FR_IMG",
            ODF_STYLE_FAMILY_ATTR: "graphic",
            f"{{{NS['style']}}}parent-style-name": "Graphics",
        })
        ET.SubElement(st_fr, f"{{{NS['style']}}}graphic-properties", {
            f"{{{NS['style']}}}vertical-pos": "top",
            f"{{{NS['style']}}}vertical-rel": "baseline",
            f"{{{NS['style']}}}horizontal-pos": "center",
            f"{{{NS['style']}}}horizontal-rel": "paragraph",
        })


def _odt_append_text_with_breaks(target: ET.Element, text: str) -> None:
    lines = text.split("\n")
    last_elem: ET.Element | None = list(target)[-1] if len(target) > 0 else None
    for line_idx, line in enumerate(lines):
        if line_idx > 0:
            br = ET.SubElement(target, ODF_LINE_BREAK)
            last_elem = br
        tab_parts = line.split("\t")
        for tab_idx, part in enumerate(tab_parts):
            if tab_idx > 0:
                tb = ET.SubElement(target, ODF_TAB)
                last_elem = tb
            if last_elem is None:
                target.text = (target.text or "") + part
            else:
                last_elem.tail = (last_elem.tail or "") + part


def odt_set_paragraph_text(
    p_elem: ET.Element,
    new_text: str,
    *,
    in_table: bool = False,
    compact_table: bool = False,
    force_span_style: str | None = None,
    bookmark_name: str | None = None,
) -> None:
    """Replace text inside <text:p> or <text:h>, preserving bookmarks/page-breaks and rendering inline markdown."""
    first_span = p_elem.find(ODF_SPAN)
    base_span_style = force_span_style or (first_span.get(ODF_STYLE_NAME) if first_span is not None else None)
    # Replace red template guidance style T22 / P23 with normal black table style
    if base_span_style == "T22":
        base_span_style = "DM_T_TBL_NORM10" if in_table else "DM_T_NORM12"
    if p_elem.get(ODF_STYLE_NAME) == "P23":
        p_elem.set(ODF_STYLE_NAME, "DM_P_TBL_CELL" if in_table else "Standard")

    preserved_children = [
        copy.deepcopy(ch)
        for ch in list(p_elem)
        if ch.tag in (ODF_BOOKMARK, ODF_BOOKMARK_START, ODF_BOOKMARK_END, ODF_SOFT_PAGE_BREAK, ODF_DRAW_FRAME)
    ]
    for ch in preserved_children:
        ch.tail = None

    for ch in list(p_elem):
        p_elem.remove(ch)
    p_elem.text = None

    for ch in preserved_children:
        p_elem.append(ch)
    if bookmark_name and not any(
        ch.tag == ODF_BOOKMARK and ch.get(f"{{{NS['text']}}}name") == bookmark_name
        for ch in preserved_children
    ):
        ET.SubElement(p_elem, ODF_BOOKMARK, {f"{{{NS['text']}}}name": bookmark_name})

    has_md = bool(re.search(r"(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)", new_text))
    if not has_md:
        target = p_elem
        if base_span_style:
            target = ET.SubElement(p_elem, ODF_SPAN, {ODF_STYLE_NAME: base_span_style})
        _odt_append_text_with_breaks(target, new_text)
        return

    if compact_table:
        norm_style = base_span_style or "DM_T_TBL_NORM85"
        bold_style = "DM_T_TBL_BOLD85"
        italic_style = "DM_T_TBL_ITALIC85"
        bi_style = "DM_T_TBL_BOLD85"
        code_style = "DM_T_TBL_CODE8"
    else:
        norm_style = base_span_style or ("DM_T_TBL_NORM10" if in_table else "DM_T_NORM12")
        bold_style = "DM_T_TBL_BOLD10" if in_table else "DM_T_BOLD12"
        italic_style = "DM_T_TBL_ITALIC10" if in_table else "DM_T_ITALIC12"
        bi_style = "DM_T_TBL_BI10" if in_table else "DM_T_BI12"
        code_style = "DM_T_TBL_CODE9" if in_table else "DM_T_CODE11"

    tokens = re.split(r"(\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)", new_text)
    for token in tokens:
        if not token:
            continue
        is_bi = token.startswith("***") and token.endswith("***") and len(token) > 6
        is_bold = not is_bi and token.startswith("**") and token.endswith("**") and len(token) > 4
        is_italic = not is_bi and not is_bold and token.startswith("*") and token.endswith("*") and len(token) > 2
        is_code = token.startswith("`") and token.endswith("`") and len(token) > 2

        if is_bi:
            content, s_style = token[3:-3], bi_style
        elif is_bold:
            content, s_style = token[2:-2], bold_style
        elif is_italic:
            content, s_style = token[1:-1], italic_style
        elif is_code:
            content, s_style = token[1:-1], code_style
        else:
            content, s_style = token, norm_style

        sp = ET.SubElement(p_elem, ODF_SPAN, {ODF_STYLE_NAME: s_style})
        _odt_append_text_with_breaks(sp, content)


def inspect_odt(odt_path: Path) -> dict:
    blocks: list[dict] = []
    tables: list[dict] = []
    tocs: list[dict] = []
    identity_fields: list[dict] = []
    placeholders: set[str] = set()

    with zipfile.ZipFile(odt_path, "r") as zf:
        root = ET.fromstring(zf.read("content.xml"))
        office_text = root.find(f".//{ODF_OFFICE_TEXT}")
        if office_text is None:
            raise ValueError(f"No office:text element found in {odt_path}")

        b_counter = 0
        t_counter = 0
        toc_counter = 0

        for child in list(office_text):
            if child.tag == ODF_TOC:
                toc_id = f"toc{toc_counter}"
                toc_counter += 1
                entries = [odt_get_paragraph_text(p) for p in child.iter(ODF_P) if odt_get_paragraph_text(p).strip()]
                tocs.append({
                    "id": toc_id,
                    "entries": entries,
                })

            elif child.tag in (ODF_P, ODF_H):
                text = odt_get_paragraph_text(child)
                for ph in PLACEHOLDER_RE.findall(text):
                    placeholders.add(ph)
                block_id = f"b{b_counter}"
                b_counter += 1
                has_img = child.find(f".//{ODF_DRAW_IMAGE}") is not None
                b_info = {
                    "id": block_id,
                    "location": "body",
                    "style": child.get(ODF_STYLE_NAME, "Standard"),
                    "is_list": False,
                    "formatting": ["heading"] if child.tag == ODF_H else [],
                    "text": text,
                }
                if has_img:
                    b_info["has_images"] = True
                blocks.append(b_info)
                id_field = detect_identity_in_block(block_id, text)
                if id_field:
                    identity_fields.append(id_field)

            elif child.tag == ODF_LIST:
                for p in child.iter(ODF_P):
                    text = odt_get_paragraph_text(p)
                    for ph in PLACEHOLDER_RE.findall(text):
                        placeholders.add(ph)
                    block_id = f"b{b_counter}"
                    b_counter += 1
                    has_img = p.find(f".//{ODF_DRAW_IMAGE}") is not None
                    b_info = {
                        "id": block_id,
                        "location": "body",
                        "style": p.get(ODF_STYLE_NAME, "List"),
                        "is_list": True,
                        "formatting": [],
                        "text": text,
                    }
                    if has_img:
                        b_info["has_images"] = True
                    blocks.append(b_info)
                    id_field = detect_identity_in_block(block_id, text)
                    if id_field:
                        identity_fields.append(id_field)

            elif child.tag == ODF_TABLE:
                table_id = f"t{t_counter}"
                t_counter += 1
                rows_data: list[list[str]] = []
                for r_idx, tr in enumerate(child.findall(ODF_ROW)):
                    row_cells: list[str] = []
                    for c_idx, tc in enumerate(tr.findall(ODF_CELL)):
                        cell_paras = [odt_get_paragraph_text(p) for p in tc.findall(ODF_P)]
                        cell_text = "\n".join(cell_paras)
                        for ph in PLACEHOLDER_RE.findall(cell_text):
                            placeholders.add(ph)
                        row_cells.append(cell_text)
                    rows_data.append(row_cells)
                    if len(row_cells) >= 2 and is_identity_label(row_cells[0]):
                        val_col = 2 if (len(row_cells) >= 3 and row_cells[1].strip() == ":") else 1
                        identity_fields.append({
                            "id": f"{table_id}.r{r_idx}.c{val_col}",
                            "table_id": table_id,
                            "cell": f"{r_idx},{val_col}",
                            "label": row_cells[0].strip().rstrip(":").strip(),
                            "current_value": row_cells[val_col].strip(),
                            "raw_text": " | ".join(row_cells),
                        })
                tables.append({
                    "id": table_id,
                    "location": "body",
                    "row_count": len(rows_data),
                    "col_count": max((len(r) for r in rows_data), default=0),
                    "rows": rows_data,
                })

    res = {
        "format": "odt",
        "source": str(odt_path),
        "placeholders": sorted(placeholders),
        "identity_fields": identity_fields,
        "blocks": blocks,
        "tables": tables,
    }
    if tocs:
        res["tocs"] = tocs
    return res


def _odt_populate_table_element(
    tbl_elem: ET.Element,
    new_rows: list[list[str]],
    auto_styles: ET.Element,
    table_uid: str,
    col_widths_in: list[float] | None = None,
    total_width_in: float = 5.5049,
) -> None:
    """Populate an ODT <table:table> with new_rows, adjusting column count and widths dynamically."""
    if not new_rows:
        return
    num_cols = max(len(r) for r in new_rows)
    existing_cols = tbl_elem.findall(ODF_TABLE_COL)
    existing_trs = tbl_elem.findall(ODF_ROW)
    if not existing_trs:
        return

    # 1. Configure column styles & <table:table-column> elements
    if col_widths_in and len(col_widths_in) == num_cols:
        widths = col_widths_in
    elif len(existing_cols) == num_cols and not col_widths_in:
        widths = None
    else:
        # Compute smart proportional widths based on content length per column
        col_weights: list[float] = []
        for c in range(num_cols):
            max_len = max((len(str(r[c])) for r in new_rows if c < len(r)), default=5)
            col_weights.append(max(min(max_len ** 0.55, 18.0), 2.2))
        total_w = sum(col_weights) or 1.0
        widths = [round(total_width_in * (w / total_w), 4) for w in col_weights]

    if widths is not None:
        # Remove old <table:table-column> and insert new ones with exact widths
        first_col_idx = list(tbl_elem).index(existing_cols[0]) if existing_cols else 0
        for old_c in existing_cols:
            tbl_elem.remove(old_c)
        for c_i, w_val in enumerate(widths):
            col_style_name = f"{table_uid}_Col{c_i}"
            st = ET.SubElement(auto_styles, ODF_STYLE, {
                ODF_STYLE_NAME_ATTR: col_style_name,
                ODF_STYLE_FAMILY_ATTR: "table-column",
            })
            ET.SubElement(st, ODF_TABLE_COL_PROPS, {ODF_COL_WIDTH: f"{w_val:.4f}in"})
            col_el = ET.Element(ODF_TABLE_COL, {ODF_TABLE_STYLE_NAME: col_style_name})
            tbl_elem.insert(first_col_idx + c_i, col_el)

    # 2. Prepare header row and data row templates
    header_tr = existing_trs[0]
    template_tr = copy.deepcopy(existing_trs[1] if len(existing_trs) > 1 else existing_trs[0])

    def _adjust_tr_cells(tr: ET.Element, is_header: bool) -> list[ET.Element]:
        # Remove any covered-table-cell
        for cov in list(tr.findall(ODF_COVERED_CELL)):
            tr.remove(cov)
        tcs = tr.findall(ODF_CELL)
        for tc in tcs:
            if ODF_NUM_COLS_SPANNED in tc.attrib:
                del tc.attrib[ODF_NUM_COLS_SPANNED]
        while len(tcs) < num_cols:
            cloned_tc = copy.deepcopy(tcs[-1])
            tr.append(cloned_tc)
            tcs = tr.findall(ODF_CELL)
        while len(tcs) > num_cols:
            tr.remove(tcs[-1])
            tcs = tr.findall(ODF_CELL)
        return tcs

    is_compact = num_cols >= 6
    existing_styles = (
        {st.get(ODF_STYLE_NAME_ATTR) for st in auto_styles.findall(ODF_STYLE)}
        if auto_styles is not None
        else set()
    )

    def _set_cell_content(tc: ET.Element, text_val: str, is_header: bool) -> None:
        # Remove any nested lists inside cell if replacing with text
        for lst in list(tc.findall(ODF_LIST)):
            tc.remove(lst)
        p_list = tc.findall(ODF_P)
        p_style = "DM_P_TBL_HDR" if is_header else "DM_P_TBL_CELL"
        base_p = p_list[0] if p_list else ET.SubElement(tc, ODF_P, {ODF_STYLE_NAME: p_style})
        for extra_p in p_list[1:]:
            tc.remove(extra_p)
        if is_compact:
            default_span = "DM_T_TBL_BOLD85" if is_header else "DM_T_TBL_NORM85"
        elif "T12" in existing_styles and "T11" in existing_styles:
            default_span = "T12" if is_header else "T11"
        else:
            default_span = "DM_T_TBL_BOLD10" if is_header else "DM_T_TBL_NORM10"
        base_p.set(ODF_STYLE_NAME, p_style)
        paragraphs = str(text_val).split("\n\n")
        odt_set_paragraph_text(
            base_p, paragraphs[0], in_table=True, compact_table=is_compact, force_span_style=default_span
        )
        for extra_txt in paragraphs[1:]:
            extra_p = copy.deepcopy(base_p)
            odt_set_paragraph_text(
                extra_p, extra_txt, in_table=True, compact_table=is_compact, force_span_style=default_span
            )
            tc.append(extra_p)

    # Update header row
    hdr_tcs = _adjust_tr_cells(header_tr, is_header=True)
    for c_idx, tc in enumerate(hdr_tcs):
        val = str(new_rows[0][c_idx]) if c_idx < len(new_rows[0]) else ""
        _set_cell_content(tc, val, is_header=True)

    # Remove old data rows
    for old_tr in existing_trs[1:]:
        tbl_elem.remove(old_tr)

    # Append new data rows
    for row_vals in new_rows[1:]:
        new_tr = copy.deepcopy(template_tr)
        data_tcs = _adjust_tr_cells(new_tr, is_header=False)
        for c_idx, tc in enumerate(data_tcs):
            val = str(row_vals[c_idx]) if c_idx < len(row_vals) else ""
            _set_cell_content(tc, val, is_header=False)
        tbl_elem.append(new_tr)


def _odt_build_rich_elements(
    item: str | dict,
    template_p: ET.Element,
    table_templates: dict[str, ET.Element],
    auto_styles: ET.Element,
    added_media: dict[str, tuple[bytes, str]],
    counter_state: dict[str, int],
) -> list[ET.Element]:
    """Convert a string or rich dict item into one or more ODF top-level elements."""
    if isinstance(item, str):
        m_img = MD_IMAGE_RE.match(item.strip())
        if m_img:
            item = {"type": "image", "caption": m_img.group(1), "path": m_img.group(2)}
        else:
            cloned_p = copy.deepcopy(template_p)
            odt_set_paragraph_text(cloned_p, item)
            return [cloned_p]

    existing_styles = (
        {st.get(ODF_STYLE_NAME_ATTR) for st in auto_styles.findall(ODF_STYLE)}
        if auto_styles is not None
        else set()
    )
    itype = item.get("type", "p")

    if itype == "p":
        p_el = copy.deepcopy(template_p)
        if "style" in item:
            p_el.set(ODF_STYLE_NAME, item["style"])
        odt_set_paragraph_text(
            p_el,
            str(item.get("text", "")),
            force_span_style=item.get("span_style"),
            bookmark_name=item.get("bookmark"),
        )
        return [p_el]

    if itype == "heading":
        level = int(item.get("level", 2))
        list_style = item.get("list_style")
        if not list_style and "WWNum8" not in existing_styles and "style" not in item:
            h_el = ET.Element(ODF_H, {
                ODF_STYLE_NAME: template_p.get(ODF_STYLE_NAME, "Standard"),
                f"{{{NS['text']}}}outline-level": str(level),
            })
            odt_set_paragraph_text(
                h_el,
                str(item.get("text", "")),
                force_span_style=item.get("span_style", "DM_T_BOLD12"),
                bookmark_name=item.get("bookmark"),
            )
            return [h_el]
        list_style = list_style or "WWNum8"
        default_pstyle = {1: "P69", 2: "P75", 3: "P71"}.get(level, "P71")
        if default_pstyle not in existing_styles:
            default_pstyle = template_p.get(ODF_STYLE_NAME, "Standard")
        p_style = item.get("style", default_pstyle)
        continue_num = item.get("continue", True)
        counter_state["list_id"] += 1
        top_list = ET.Element(ODF_LIST, {
            XML_ID: f"list_dm_{counter_state['list_id']}",
            ODF_STYLE_NAME: list_style,
        })
        if continue_num:
            top_list.set(ODF_CONTINUE_NUMBERING, "true")
        curr_li = ET.SubElement(top_list, ODF_LIST_ITEM)
        for _ in range(1, level):
            sub_list = ET.SubElement(curr_li, ODF_LIST)
            curr_li = ET.SubElement(sub_list, ODF_LIST_ITEM)
        p_el = ET.SubElement(curr_li, ODF_P, {ODF_STYLE_NAME: p_style})
        odt_set_paragraph_text(
            p_el,
            str(item.get("text", "")),
            force_span_style=item.get("span_style"),
            bookmark_name=item.get("bookmark"),
        )
        return [top_list]

    if itype == "list":
        list_style = item.get("list_style", "WWNum7")
        items_list = item.get("items", [])
        fallback_p = template_p.get(ODF_STYLE_NAME, "Standard")
        first_style = item.get("first_style", item.get("item_style", "P35" if "P35" in existing_styles else fallback_p))
        mid_style = item.get("item_style", "P38" if "P38" in existing_styles else fallback_p)
        last_style = item.get("last_style", item.get("item_style", "P43" if "P43" in existing_styles else fallback_p))
        span_style = item.get("span_style", "T3" if "T3" in existing_styles else "DM_T_NORM12")
        counter_state["list_id"] += 1
        top_list = ET.Element(ODF_LIST, {
            XML_ID: f"list_dm_{counter_state['list_id']}",
            ODF_STYLE_NAME: list_style,
        })
        if item.get("continue"):
            top_list.set(ODF_CONTINUE_NUMBERING, "true")
        for idx_i, sub_txt in enumerate(items_list):
            li = ET.SubElement(top_list, ODF_LIST_ITEM)
            st_name = first_style if idx_i == 0 else (last_style if idx_i == len(items_list) - 1 else mid_style)
            p_el = ET.SubElement(li, ODF_P, {ODF_STYLE_NAME: st_name})
            odt_set_paragraph_text(p_el, str(sub_txt), force_span_style=span_style)
        return [top_list]

    if itype == "image":
        img_path = Path(item["path"])
        data, mime, ext, w_px, h_px = get_image_meta(img_path)
        max_w = float(item.get("max_width_in", 5.4))
        max_h = float(item.get("max_height_in", 4.2))
        w_in, h_in = compute_display_dims_in(w_px, h_px, max_w, max_h)
        if "width_in" in item and "height_in" in item:
            w_in, h_in = float(item["width_in"]), float(item["height_in"])
        digest = hashlib.sha1(data).hexdigest()[:16].upper()
        pkg_href = f"Pictures/DM_{digest}.{ext}"
        added_media[pkg_href] = (data, mime)

        default_center_p = "P18" if "P18" in existing_styles else "DM_P_CENTER"
        default_fr = "fr1" if "fr1" in existing_styles else "DM_FR_IMG"
        default_cap_span = "T4" if "T4" in existing_styles else "DM_T_BOLD12"

        counter_state["img_id"] += 1
        p_img = ET.Element(ODF_P, {ODF_STYLE_NAME: item.get("style", default_center_p)})
        frame = ET.SubElement(p_img, ODF_DRAW_FRAME, {
            f"{{{NS['draw']}}}style-name": default_fr,
            f"{{{NS['draw']}}}name": f"Picture_{counter_state['img_id']}",
            f"{{{NS['text']}}}anchor-type": "as-char",
            f"{{{NS['svg']}}}width": f"{w_in:.4f}in",
            f"{{{NS['svg']}}}height": f"{h_in:.4f}in",
            f"{{{NS['draw']}}}z-index": "0",
        })
        ET.SubElement(frame, ODF_DRAW_IMAGE, {
            f"{{{NS['xlink']}}}href": pkg_href,
            f"{{{NS['xlink']}}}type": "simple",
            f"{{{NS['xlink']}}}show": "embed",
            f"{{{NS['xlink']}}}actuate": "onLoad",
            f"{{{NS['loext']}}}mime-type": mime,
        })
        out_elems = [p_img]
        caption = item.get("caption")
        if caption:
            p_cap = ET.Element(ODF_P, {ODF_STYLE_NAME: item.get("caption_style", default_center_p)})
            odt_set_paragraph_text(
                p_cap,
                str(caption),
                force_span_style=item.get("caption_span", default_cap_span),
                bookmark_name=item.get("bookmark"),
            )
            out_elems.append(p_cap)
        return out_elems

    if itype == "table":
        out_elems = []
        default_center_p = "P18" if "P18" in existing_styles else "DM_P_CENTER"
        default_cap_span = "T4" if "T4" in existing_styles else "DM_T_BOLD12"
        caption = item.get("caption")
        if caption:
            p_cap = ET.Element(ODF_P, {ODF_STYLE_NAME: item.get("caption_style", default_center_p)})
            odt_set_paragraph_text(
                p_cap,
                str(caption),
                force_span_style=item.get("caption_span", default_cap_span),
                bookmark_name=item.get("bookmark"),
            )
            out_elems.append(p_cap)
        tpl_id = item.get("template_table", "t6")
        base_tbl = table_templates.get(tpl_id)
        if base_tbl is None and table_templates:
            base_tbl = next(iter(table_templates.values()))
        if base_tbl is not None:
            new_tbl = copy.deepcopy(base_tbl)
            counter_state["tbl_id"] += 1
            uid = f"DMTable_{counter_state['tbl_id']}"
            new_tbl.set(ODF_TABLE_NAME, uid)
            _odt_populate_table_element(
                new_tbl,
                item.get("rows", []),
                auto_styles,
                uid,
                col_widths_in=item.get("col_widths"),
                total_width_in=float(item.get("total_width_in", 5.5049)),
            )
            out_elems.append(new_tbl)
        return out_elems

    if itype == "code_table":
        tpl_id = item.get("template_table", "t5")
        base_tbl = table_templates.get(tpl_id)
        if base_tbl is None and table_templates:
            base_tbl = next(iter(table_templates.values()))
        if base_tbl is None:
            p_el = copy.deepcopy(template_p)
            odt_set_paragraph_text(p_el, str(item.get("code", "")))
            return [p_el]
        new_tbl = copy.deepcopy(base_tbl)
        counter_state["tbl_id"] += 1
        uid = f"DMCodeTable_{counter_state['tbl_id']}"
        new_tbl.set(ODF_TABLE_NAME, uid)
        title = str(item.get("title", "Kode Sumber"))
        code_lines = str(item.get("code", "")).strip("\n").split("\n")
        if tpl_id not in table_templates:
            nums_str = "\n".join(str(i) for i in range(1, len(code_lines) + 1))
            code_str = "\n".join(f"`{ln}`" if ln.strip() else " " for ln in code_lines)
            _odt_populate_table_element(
                new_tbl,
                [["#", title], [nums_str, code_str]],
                auto_styles,
                uid,
                col_widths_in=[0.45, 5.05],
            )
            return [new_tbl]
        code_p_style = "P11" if "P11" in existing_styles else "DM_P_CODE"
        code_span_style = "T13" if "T13" in existing_styles else "DM_T_TBL_CODE9"
        trs = new_tbl.findall(ODF_ROW)
        if len(trs) >= 1:
            tcs0 = trs[0].findall(ODF_CELL)
            if tcs0:
                p0 = tcs0[0].find(ODF_P)
                if p0 is not None:
                    odt_set_paragraph_text(p0, title, force_span_style=code_span_style)
        if len(trs) >= 2:
            tcs1 = trs[1].findall(ODF_CELL)
            if len(tcs1) >= 2:
                for ch in list(tcs1[0]):
                    tcs1[0].remove(ch)
                for ch in list(tcs1[1]):
                    tcs1[1].remove(ch)
                for l_idx, c_line in enumerate(code_lines, start=1):
                    p_num = ET.SubElement(tcs1[0], ODF_P, {ODF_STYLE_NAME: code_p_style})
                    sp_num = ET.SubElement(p_num, ODF_SPAN, {ODF_STYLE_NAME: code_span_style})
                    sp_num.text = str(l_idx)

                    p_code = ET.SubElement(tcs1[1], ODF_P, {ODF_STYLE_NAME: code_p_style})
                    sp_code = ET.SubElement(p_code, ODF_SPAN, {ODF_STYLE_NAME: code_span_style})
                    leading = len(c_line) - len(c_line.lstrip(" "))
                    stripped = c_line.lstrip(" ")
                    if leading > 0:
                        s_el = ET.SubElement(sp_code, ODF_S, {f"{{{NS['text']}}}c": str(leading)})
                        s_el.tail = stripped
                    else:
                        sp_code.text = stripped
        return [new_tbl]

    cloned_p = copy.deepcopy(template_p)
    odt_set_paragraph_text(cloned_p, str(item.get("text", "")))
    return [cloned_p]


def _odt_update_toc_index_body(toc_elem: ET.Element, entries: list[dict]) -> None:
    """Populate <text:index-body> inside a <text:table-of-content> element."""
    idx_body = toc_elem.find(ODF_INDEX_BODY)
    if idx_body is None:
        idx_body = ET.SubElement(toc_elem, ODF_INDEX_BODY)
    for ch in list(idx_body):
        idx_body.remove(ch)
    for ent in entries:
        p_style = ent.get("style", "P15")
        span_style = ent.get("span_style", "T3")
        href = ent.get("href", "")
        label = str(ent.get("text", ""))
        page = str(ent.get("page", ""))
        p_el = ET.SubElement(idx_body, ODF_P, {ODF_STYLE_NAME: p_style})
        container = p_el
        if href:
            container = ET.SubElement(p_el, ODF_A, {
                f"{{{NS['xlink']}}}type": "simple",
                f"{{{NS['xlink']}}}href": href,
                ODF_STYLE_NAME: "Index_20_Link",
                f"{{{NS['text']}}}visited-style-name": "Index_20_Link",
            })
        sp = ET.SubElement(container, ODF_SPAN, {ODF_STYLE_NAME: span_style})
        if page:
            sp.text = label
            tb = ET.SubElement(sp, ODF_TAB)
            tb.tail = page
        else:
            sp.text = label


def render_odt(odt_path: Path, spec: dict, out_odt_path: Path) -> None:
    placeholders: dict[str, str] = spec.get("placeholders", {})
    block_ops: dict[str, dict] = {op["id"]: op for op in spec.get("blocks", []) if "id" in op}
    table_ops: dict[str, dict] = {op["id"]: op for op in spec.get("tables", []) if "id" in op}
    toc_ops: dict[str, list[dict]] = spec.get("toc_updates", {})

    for op in spec.get("blocks", []):
        bid = op.get("id", "")
        m = re.match(r"^(t\d+)\.r(\d+)\.c(\d+)$", bid)
        if m and "text" in op:
            tid, r_idx, c_idx = m.group(1), m.group(2), m.group(3)
            t_op = table_ops.setdefault(tid, {"id": tid})
            cells = t_op.setdefault("cells", {})
            cells[f"{r_idx},{c_idx}"] = op["text"]

    added_media: dict[str, tuple[bytes, str]] = {}
    counter_state = {"list_id": 1000, "img_id": 100, "tbl_id": 100}

    with zipfile.ZipFile(odt_path, "r") as zin:
        root = ET.fromstring(zin.read("content.xml"))
        auto_styles = root.find(f".//{ODF_OFFICE_AUTO_STYLES}")
        if auto_styles is not None:
            _odt_ensure_inline_styles(auto_styles)

        office_text = root.find(f".//{ODF_OFFICE_TEXT}")
        if office_text is None:
            raise ValueError(f"No office:text element found in {odt_path}")

        # Snapshot all original template tables by id (t0..tN) so rich items can clone them
        table_templates: dict[str, ET.Element] = {}
        t_snap = 0
        for ch in list(office_text):
            if ch.tag == ODF_TABLE:
                table_templates[f"t{t_snap}"] = copy.deepcopy(ch)
                t_snap += 1

        # Find default body paragraph template (P33 if present)
        default_body_p = ET.Element(ODF_P, {ODF_STYLE_NAME: "P33"})
        for ch in list(office_text):
            if ch.tag == ODF_P and ch.get(ODF_STYLE_NAME) == "P33":
                default_body_p = copy.deepcopy(ch)
                break

        b_counter = 0
        t_counter = 0
        toc_counter = 0

        for child in list(office_text):
            if child.tag == ODF_TOC:
                toc_key = f"toc{toc_counter}"
                toc_counter += 1
                if toc_key in toc_ops:
                    if toc_ops[toc_key] is None or toc_ops[toc_key] == "delete":
                        office_text.remove(child)
                    else:
                        _odt_update_toc_index_body(child, toc_ops[toc_key])
                continue

            if child.tag in (ODF_P, ODF_H):
                block_id = f"b{b_counter}"
                b_counter += 1
                op = block_ops.get(block_id)
                if op:
                    if op.get("delete"):
                        office_text.remove(child)
                        continue
                    if "items" in op and isinstance(op["items"], list):
                        items = op["items"]
                        if not items:
                            office_text.remove(child)
                            continue
                        idx = list(office_text).index(child)
                        office_text.remove(child)
                        insert_offset = 0
                        for item in items:
                            built = _odt_build_rich_elements(
                                item, child, table_templates, auto_styles, added_media, counter_state
                            )
                            for el in built:
                                office_text.insert(idx + insert_offset, el)
                                insert_offset += 1
                        continue
                    elif "text" in op:
                        odt_set_paragraph_text(
                            child,
                            str(op["text"]),
                            force_span_style=op.get("span_style"),
                            bookmark_name=op.get("bookmark"),
                        )
                        if "insert_toc_after" in op and isinstance(op["insert_toc_after"], list):
                            existing_toc = office_text.find(ODF_TOC)
                            if existing_toc is not None:
                                new_toc = copy.deepcopy(existing_toc)
                                _odt_update_toc_index_body(new_toc, op["insert_toc_after"])
                                idx = list(office_text).index(child)
                                office_text.insert(idx + 1, new_toc)

                if placeholders:
                    txt = odt_get_paragraph_text(child)
                    if any(k in txt for k in placeholders):
                        for k, v in placeholders.items():
                            txt = txt.replace(k, str(v))
                        odt_set_paragraph_text(child, txt)

            elif child.tag == ODF_LIST:
                # Flatten all <text:p> inside this <text:list> in document order
                list_ps = list(child.iter(ODF_P))
                is_single_p_list = len(list_ps) == 1
                if is_single_p_list:
                    p = list_ps[0]
                    block_id = f"b{b_counter}"
                    b_counter += 1
                    op = block_ops.get(block_id)
                    if op:
                        if op.get("delete"):
                            office_text.remove(child)
                            continue
                        if "items" in op and isinstance(op["items"], list):
                            items = op["items"]
                            if not items:
                                office_text.remove(child)
                                continue
                            # Check if all items are simple strings (list items) or rich elements
                            if all(isinstance(x, str) and not MD_IMAGE_RE.match(x.strip()) for x in items):
                                li = child.find(ODF_LIST_ITEM)
                                if li is not None and li.find(ODF_P) is not None:
                                    odt_set_paragraph_text(p, str(items[0]))
                                    for offset, item_text in enumerate(items[1:], start=1):
                                        cloned_li = copy.deepcopy(li)
                                        cloned_p = cloned_li.find(ODF_P)
                                        if cloned_p is not None:
                                            odt_set_paragraph_text(cloned_p, str(item_text))
                                        child.insert(offset, cloned_li)
                                else:
                                    odt_set_paragraph_text(p, str(items[0]))
                            else:
                                idx = list(office_text).index(child)
                                office_text.remove(child)
                                insert_offset = 0
                                for item in items:
                                    built = _odt_build_rich_elements(
                                        item, default_body_p, table_templates, auto_styles, added_media, counter_state
                                    )
                                    for el in built:
                                        office_text.insert(idx + insert_offset, el)
                                        insert_offset += 1
                                continue
                        elif "text" in op:
                            odt_set_paragraph_text(
                                p,
                                str(op["text"]),
                                force_span_style=op.get("span_style"),
                                bookmark_name=op.get("bookmark"),
                            )
                            if "continue" in op:
                                if op["continue"]:
                                    child.set(ODF_CONTINUE_NUMBERING, "true")
                                elif ODF_CONTINUE_NUMBERING in child.attrib:
                                    del child.attrib[ODF_CONTINUE_NUMBERING]
                    if placeholders:
                        txt = odt_get_paragraph_text(p)
                        if any(k in txt for k in placeholders):
                            for k, v in placeholders.items():
                                txt = txt.replace(k, str(v))
                            odt_set_paragraph_text(p, txt)
                else:
                    # Multi-item or nested list (e.g. WWNum7, WWNum4, etc.)
                    remove_whole_list = False
                    parent_map = {c: p_node for p_node in child.iter() for c in p_node}
                    for li in list(child.findall(ODF_LIST_ITEM)):
                        for p in list(li.iter(ODF_P)):
                            block_id = f"b{b_counter}"
                            b_counter += 1
                            op = block_ops.get(block_id)
                            if op:
                                if op.get("delete"):
                                    p_parent = parent_map.get(p)
                                    if p_parent is not None and p in list(p_parent):
                                        p_parent.remove(p)
                                    continue
                                if "items" in op and isinstance(op["items"], list):
                                    items = op["items"]
                                    if all(isinstance(x, str) for x in items) and p in list(li):
                                        idx_li = list(child).index(li)
                                        odt_set_paragraph_text(p, str(items[0]))
                                        for offset, item_text in enumerate(items[1:], start=1):
                                            cloned_li = copy.deepcopy(li)
                                            cloned_p = cloned_li.find(ODF_P)
                                            if cloned_p is not None:
                                                odt_set_paragraph_text(cloned_p, str(item_text))
                                            child.insert(idx_li + offset, cloned_li)
                                    else:
                                        idx_body = list(office_text).index(child)
                                        insert_offset = 1
                                        for item in items:
                                            built = _odt_build_rich_elements(
                                                item, default_body_p, table_templates, auto_styles, added_media, counter_state
                                            )
                                            for el in built:
                                                office_text.insert(idx_body + insert_offset, el)
                                                insert_offset += 1
                                        remove_whole_list = True
                                elif "text" in op:
                                    odt_set_paragraph_text(
                                        p,
                                        str(op["text"]),
                                        force_span_style=op.get("span_style"),
                                        bookmark_name=op.get("bookmark"),
                                    )
                            if placeholders:
                                txt = odt_get_paragraph_text(p)
                                if any(k in txt for k in placeholders):
                                    for k, v in placeholders.items():
                                        txt = txt.replace(k, str(v))
                                    odt_set_paragraph_text(p, txt)
                        # Prune empty nested <text:list> / <text:list-item>
                        for sub in list(li.iter(ODF_LIST)):
                            if len(list(sub.iter(ODF_P))) == 0:
                                sub_parent = parent_map.get(sub)
                                if sub_parent is not None and sub in list(sub_parent):
                                    sub_parent.remove(sub)
                        if len(list(li.iter(ODF_P))) == 0 and li in list(child):
                            child.remove(li)
                    if remove_whole_list or len(list(child.iter(ODF_P))) == 0:
                        if child in list(office_text):
                            office_text.remove(child)

            elif child.tag == ODF_TABLE:
                table_id = f"t{t_counter}"
                t_counter += 1
                t_op = table_ops.get(table_id)
                if t_op:
                    if t_op.get("delete"):
                        office_text.remove(child)
                        continue
                    if "rows" in t_op and isinstance(t_op["rows"], list):
                        _odt_populate_table_element(
                            child,
                            t_op["rows"],
                            auto_styles,
                            f"Tbl_{table_id}",
                            col_widths_in=t_op.get("col_widths"),
                            total_width_in=float(t_op.get("total_width_in", 5.5049)),
                        )
                    if "cells" in t_op and isinstance(t_op["cells"], dict):
                        existing_trs = child.findall(ODF_ROW)
                        for rc_key, val in t_op["cells"].items():
                            r_str, c_str = rc_key.split(",")
                            r_i, c_i = int(r_str.strip()), int(c_str.strip())
                            if 0 <= r_i < len(existing_trs):
                                tcs = existing_trs[r_i].findall(ODF_CELL)
                                if 0 <= c_i < len(tcs):
                                    tc = tcs[c_i]
                                    p_list = tc.findall(ODF_P)
                                    if p_list:
                                        base_p = p_list[0]
                                        for extra_p in p_list[1:]:
                                            tc.remove(extra_p)
                                        paragraphs = str(val).split("\n\n")
                                        odt_set_paragraph_text(base_p, paragraphs[0], in_table=True, force_span_style="T11")
                                        for extra_txt in paragraphs[1:]:
                                            extra_p = copy.deepcopy(base_p)
                                            odt_set_paragraph_text(extra_p, extra_txt, in_table=True, force_span_style="T11")
                                            tc.append(extra_p)

                if placeholders:
                    for p_in_tbl in child.iter(ODF_P):
                        txt = odt_get_paragraph_text(p_in_tbl)
                        if any(k in txt for k in placeholders):
                            for k, v in placeholders.items():
                                txt = txt.replace(k, str(v))
                            odt_set_paragraph_text(p_in_tbl, txt, in_table=True)

        modified_content = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        # Update META-INF/manifest.xml if any images were added
        modified_manifest: bytes | None = None
        if added_media and "META-INF/manifest.xml" in zin.namelist():
            m_root = ET.fromstring(zin.read("META-INF/manifest.xml"))
            existing_paths = {
                fe.get(f"{{{NS['manifest']}}}full-path")
                for fe in m_root.findall(f"{{{NS['manifest']}}}file-entry")
            }
            for pkg_href, (_, mime) in added_media.items():
                if pkg_href not in existing_paths:
                    ET.SubElement(m_root, f"{{{NS['manifest']}}}file-entry", {
                        f"{{{NS['manifest']}}}full-path": pkg_href,
                        f"{{{NS['manifest']}}}media-type": mime,
                    })
            modified_manifest = ET.tostring(m_root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(out_odt_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "content.xml":
                    zout.writestr(item, modified_content)
                elif item.filename == "META-INF/manifest.xml" and modified_manifest is not None:
                    zout.writestr(item, modified_manifest)
                else:
                    zout.writestr(item, zin.read(item.filename))
            for pkg_href, (data, _) in added_media.items():
                zout.writestr(pkg_href, data)


# ============================================================================
# Unified Commands (inspect, render, self-test)
# ============================================================================

def inspect_document(input_path: Path) -> dict:
    ext = input_path.suffix.lower()
    if ext == ".docx":
        return inspect_docx(input_path)
    if ext == ".odt":
        return inspect_odt(input_path)
    if ext == ".doc":
        with tempfile.TemporaryDirectory(prefix="doc_maker_") as tmp_dir:
            converted = convert_with_soffice(input_path, "docx", Path(tmp_dir))
            info = inspect_docx(converted)
            info["format"] = "doc (via docx)"
            info["source"] = str(input_path)
            return info
    raise ValueError(f"Unsupported input format: {ext}. Expected .doc, .docx, or .odt")


def render_document(input_path: Path, spec: dict, outputs: list[Path]) -> list[str]:
    if not outputs:
        raise ValueError("At least one output path (-o) is required.")

    ext = input_path.suffix.lower()
    written: list[str] = []

    with tempfile.TemporaryDirectory(prefix="doc_maker_render_") as tmp_dir_str:
        tmp_dir = Path(tmp_dir_str)
        working_input = input_path
        working_ext = ext
        if ext == ".doc":
            working_input = convert_with_soffice(input_path, "docx", tmp_dir)
            working_ext = ".docx"

        native_rendered = tmp_dir / f"rendered{working_ext}"
        if working_ext == ".docx":
            render_docx(working_input, spec, native_rendered)
        elif working_ext == ".odt":
            render_odt(working_input, spec, native_rendered)
        else:
            raise ValueError(f"Unsupported working format: {working_ext}")

        for out_path in outputs:
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_ext = out_path.suffix.lower()
            if out_ext == working_ext:
                shutil.copy2(native_rendered, out_path)
            elif out_ext in (".docx", ".odt", ".pdf"):
                converted = convert_with_soffice(native_rendered, out_ext.lstrip("."), tmp_dir)
                shutil.copy2(converted, out_path)
                if out_ext == ".docx":
                    ensure_docx_update_fields(out_path)
            else:
                raise ValueError(f"Unsupported output extension: {out_ext}")
            written.append(str(out_path))

    return written


def create_minimal_sample_docx(path: Path) -> None:
    """Create a valid reference .docx with styles, identity fields, split placeholders, list, and table."""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>LAPORAN PRAKTIKUM {{course}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Nama</w:t><w:tab/><w:t>: Budi Santoso</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>NIM</w:t><w:tab/><w:t>: 21000123</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Age</w:t><w:tab/><w:t>: 20</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Sample bullet point from template</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Old extra paragraph to delete</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>No</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Topic</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Template Topic</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document_xml)


def run_self_test() -> None:
    with tempfile.TemporaryDirectory(prefix="doc_maker_test_") as tmp:
        tmp_dir = Path(tmp)
        sample_docx = tmp_dir / "template.docx"
        create_minimal_sample_docx(sample_docx)

        # Create a 1x1 PNG sample image for testing image embedding
        sample_png = tmp_dir / "sample.png"
        sample_png.write_bytes(
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
            b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\xcf\xc0"
            b"\x00\x00\x03\x01\x01\x00\x18\xdd\x8d\xb0\x00\x00\x00\x00IEND\xaeB`\x82"
        )

        # 1. Inspect .docx
        info = inspect_document(sample_docx)
        assert "{{course}}" in info["placeholders"], f"Missing placeholder: {info}"
        id_labels = {f["label"] for f in info["identity_fields"]}
        assert {"Nama", "NIM", "Age"}.issubset(id_labels), f"Missing identity fields: {id_labels}"

        # 2. Render to both .docx and .odt with rich items (image, dynamic 3-col table, inline markdown)
        spec = {
            "placeholders": {"{{course}}": "JARINGAN KOMPUTER"},
            "blocks": [
                {"id": "b1", "text": "Nama\t: Andi Wijaya"},
                {"id": "b2", "text": "NIM\t: 24099887"},
                {"id": "b3", "text": "Age\t: 21"},
                {
                    "id": "b4",
                    "items": [
                        "Bullet 1 with **bold** and `code`",
                        "Bullet 2",
                        {"type": "image", "path": str(sample_png), "caption": "Gambar 1.1 Arsitektur Uji"},
                        {
                            "type": "table",
                            "caption": "Tabel 1.2 Tabel Dinamis 3 Kolom",
                            "rows": [
                                ["Komponen", "Metrik", "Hasil"],
                                ["Qdrant", "QPS", "173.92"],
                            ],
                        },
                        "Bullet 3",
                    ],
                },
                {"id": "b5", "delete": True},
            ],
            "tables": [
                {
                    "id": "t0",
                    "rows": [
                        ["No", "Topic", "Status"],
                        ["1", "OSI Layer", "Done"],
                        ["2", "TCP/IP", "Done"],
                        ["3", "Routing", "Verified"],
                    ],
                }
            ],
        }
        out_docx = tmp_dir / "out.docx"
        outputs = [out_docx]
        if find_soffice():
            outputs.append(tmp_dir / "out.odt")

        render_document(sample_docx, spec, outputs)

        # 3. Verify rendered .docx
        rendered_info = inspect_document(out_docx)
        texts = [b["text"] for b in rendered_info["blocks"]]
        assert "LAPORAN PRAKTIKUM JARINGAN KOMPUTER" in texts[0]
        assert "Andi Wijaya" in texts[1]
        assert "24099887" in texts[2]
        assert any("Bullet 3" in t for t in texts)
        assert any("Gambar 1.1 Arsitektur Uji" in t for t in texts)
        assert any(b.get("has_images") for b in rendered_info["blocks"]), "Expected embedded image block in DOCX"
        assert "Old extra paragraph to delete" not in texts
        assert len(rendered_info["tables"]) == 2
        assert rendered_info["tables"][0]["rows"][1] == ["Qdrant", "QPS", "173.92"]
        assert rendered_info["tables"][1]["row_count"] == 4
        assert rendered_info["tables"][1]["col_count"] == 3
        assert rendered_info["tables"][1]["rows"][3] == ["3", "Routing", "Verified"]

        # 4. If LibreOffice is present, also test .odt native inspect & render with rich items
        if find_soffice():
            out_odt = tmp_dir / "out.odt"
            odt_info = inspect_document(out_odt)
            assert len(odt_info["tables"]) == 2
            out_odt2 = tmp_dir / "out2.odt"
            render_document(
                out_odt,
                {
                    "blocks": [
                        {"id": "b1", "text": "Nama\t: Citra Dewi"},
                        {
                            "id": "b2",
                            "items": [
                                {"type": "p", "text": "NIM\t: **23515040711005**"},
                                {"type": "image", "path": str(sample_png), "caption": "Gambar 2.1 Diagram ODT"},
                            ],
                        },
                    ]
                },
                [out_odt2],
            )
            odt2_info = inspect_document(out_odt2)
            assert "Citra Dewi" in odt2_info["blocks"][1]["text"]
            assert any(b.get("has_images") for b in odt2_info["blocks"]), "Expected embedded image block in ODT"

    print("SELF-TEST PASSED: All .docx / .odt inspection, identity detection, cloning, and rendering checks succeeded.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect and fill .doc, .docx, and .odt reference templates.")
    parser.add_argument("--self-test", action="store_true", help="Run built-in verification suite")
    sub = parser.add_subparsers(dest="cmd")

    p_inspect = sub.add_parser("inspect", help="Extract structure, identity fields, blocks, and tables as JSON")
    p_inspect.add_argument("input", type=Path, help="Path to reference .doc, .docx, or .odt")
    p_inspect.add_argument("-o", "--output", type=Path, default=None, help="Optional path to save UTF-8 JSON output directly to a file")

    p_render = sub.add_parser("render", help="Render new document(s) from reference template and JSON spec")
    p_render.add_argument("input", type=Path, help="Path to reference .doc, .docx, or .odt")
    p_render.add_argument("-d", "--data", type=Path, required=True, help="Path to JSON render spec")
    p_render.add_argument("-o", "--output", type=Path, action="append", required=True, help="Output file path(s) (.docx, .odt)")

    args = parser.parse_args()
    if args.self_test:
        run_self_test()
        return

    if args.cmd == "inspect":
        res = inspect_document(args.input)
        json_str = json.dumps(res, indent=2, ensure_ascii=False)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(json_str, encoding="utf-8")
            print(json.dumps({
                "status": "ok",
                "output": str(args.output),
                "format": res.get("format"),
                "blocks": len(res.get("blocks", [])),
                "tables": len(res.get("tables", [])),
                "identity_fields": len(res.get("identity_fields", [])),
            }, indent=2, ensure_ascii=False))
        else:
            print(json_str)
    elif args.cmd == "render":
        spec = json.loads(args.data.read_text(encoding="utf-8"))
        written = render_document(args.input, spec, args.output)
        print(json.dumps({"status": "ok", "outputs": written}, indent=2, ensure_ascii=False))
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
