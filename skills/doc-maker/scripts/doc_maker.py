#!/usr/bin/env python3
"""
doc_maker.py - Zero-dependency CLI engine to inspect and render .docx and .odt
documents using any existing .doc, .docx, or .odt file as a formatting template.
"""

from __future__ import annotations

import argparse
import copy
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

# XML Namespaces
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "xml": "http://www.w3.org/XML/1998/namespace",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
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
W_TBL = f"{{{NS['w']}}}tbl"
W_TR = f"{{{NS['w']}}}tr"
W_TC = f"{{{NS['w']}}}tc"
W_BODY = f"{{{NS['w']}}}body"
W_VAL = f"{{{NS['w']}}}val"
W_B = f"{{{NS['w']}}}b"
W_I = f"{{{NS['w']}}}i"
W_U = f"{{{NS['w']}}}u"
W_SZ = f"{{{NS['w']}}}sz"
W_COLOR = f"{{{NS['w']}}}color"
XML_SPACE = f"{{{NS['xml']}}}space"

ODF_P = f"{{{NS['text']}}}p"
ODF_H = f"{{{NS['text']}}}h"
ODF_SPAN = f"{{{NS['text']}}}span"
ODF_S = f"{{{NS['text']}}}s"
ODF_TAB = f"{{{NS['text']}}}tab"
ODF_LINE_BREAK = f"{{{NS['text']}}}line-break"
ODF_LIST = f"{{{NS['text']}}}list"
ODF_LIST_ITEM = f"{{{NS['text']}}}list-item"
ODF_STYLE_NAME = f"{{{NS['text']}}}style-name"
ODF_TABLE = f"{{{NS['table']}}}table"
ODF_ROW = f"{{{NS['table']}}}table-row"
ODF_CELL = f"{{{NS['table']}}}table-cell"
ODF_OFFICE_TEXT = f"{{{NS['office']}}}text"

PLACEHOLDER_RE = re.compile(r"(\{\{[^{}\n]+\}\}|\[\[[A-Za-z0-9_.\-\s]+\]\]|<[A-Za-z0-9_.\-]+>)")

IDENTITY_KEYWORDS = {
    "name", "full name", "student name", "author", "prepared by", "submitted by",
    "nama", "nama lengkap", "nama mahasiswa", "nama siswa", "penulis", "penyusun",
    "nim", "npm", "nrp", "nis", "nisn", "nik", "nip", "id", "student id", "employee id", "no induk", "nomor induk",
    "age", "umur", "usia",
    "ttl", "tempat tanggal lahir", "date of birth", "dob", "birth date",
    "gender", "sex", "jenis kelamin",
    "class", "kelas", "semester", "angkatan", "batch", "group", "kelompok",
    "department", "major", "program studi", "prodi", "jurusan", "faculty", "fakultas",
    "university", "universitas", "institution", "instansi", "school", "sekolah", "company", "perusahaan",
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
        # Also check "Label :" with empty value
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
            if rpr.find(W_B) is not None:
                formatting.append("bold")
            if rpr.find(W_I) is not None:
                formatting.append("italic")
            if rpr.find(W_U) is not None:
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
    """Replace text inside a <w:p> while preserving <w:pPr> and first run's <w:rPr>."""
    ppr = p_elem.find(W_PPR)
    first_r = p_elem.find(W_R)
    base_rpr = copy.deepcopy(first_r.find(W_RPR)) if (first_r is not None and first_r.find(W_RPR) is not None) else None

    # Remove all child elements except <w:pPr>
    for child in list(p_elem):
        if child is not ppr:
            p_elem.remove(child)

    # Parse inline **bold** and *italic* segments
    tokens = re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", new_text)
    for token in tokens:
        if not token:
            continue
        is_bold = token.startswith("**") and token.endswith("**") and len(token) > 4
        is_italic = not is_bold and token.startswith("*") and token.endswith("*") and len(token) > 2
        content = token[2:-2] if is_bold else (token[1:-1] if is_italic else token)

        r_elem = ET.SubElement(p_elem, W_R)
        rpr = copy.deepcopy(base_rpr) if base_rpr is not None else ET.Element(W_RPR)
        if is_bold and rpr.find(W_B) is None:
            ET.SubElement(rpr, W_B)
        if is_italic and rpr.find(W_I) is None:
            ET.SubElement(rpr, W_I)
        if len(rpr) > 0 or base_rpr is not None:
            r_elem.append(rpr)

        # Handle \n and \t inside content
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
    """Replace placeholders even when split across multiple <w:r><w:t> runs."""
    if not mapping:
        return
    full_text = docx_get_paragraph_text(p_elem)
    if not any(k in full_text for k in mapping):
        return

    # Fast & safe run-level replacement first; if split across runs, rebuild paragraph text
    t_elems = [elem for elem in p_elem.iter(W_T)]
    combined_t = "".join(t.text or "" for t in t_elems)
    if not any(k in combined_t for k in mapping):
        return

    # Replace in each <w:t> if placeholder is contained within a single <w:t>
    for t in t_elems:
        if t.text:
            for k, v in mapping.items():
                if k in t.text:
                    t.text = t.text.replace(k, str(v))
                    t.set(XML_SPACE, "preserve")

    # Check if any placeholder still remains due to split runs
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
                    blocks.append({
                        "id": block_id,
                        "location": loc,
                        "style": info["style"],
                        "is_list": info["is_list"],
                        "formatting": info["formatting"],
                        "text": text,
                    })
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
                        # Check 2-column or 3-column key-value identity rows in tables
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


def render_docx(docx_path: Path, spec: dict, out_docx_path: Path) -> None:
    placeholders: dict[str, str] = spec.get("placeholders", {})
    block_ops: dict[str, dict] = {op["id"]: op for op in spec.get("blocks", []) if "id" in op}
    table_ops: dict[str, dict] = {op["id"]: op for op in spec.get("tables", []) if "id" in op}

    # Support shorthand identity/table cell updates in blocks list like {"id": "t0.r1.c1", "text": "..."}
    for op in spec.get("blocks", []):
        bid = op.get("id", "")
        m = re.match(r"^(t\d+)\.r(\d+)\.c(\d+)$", bid)
        if m and "text" in op:
            tid, r_idx, c_idx = m.group(1), m.group(2), m.group(3)
            t_op = table_ops.setdefault(tid, {"id": tid})
            cells = t_op.setdefault("cells", {})
            cells[f"{r_idx},{c_idx}"] = op["text"]

    with zipfile.ZipFile(docx_path, "r") as zin:
        xml_targets: dict[str, str] = {"word/document.xml": "body"}
        for name in sorted(zin.namelist()):
            if re.match(r"^word/header\d*\.xml$", name) or re.match(r"^word/footer\d*\.xml$", name):
                xml_targets[name] = Path(name).stem

        b_counter = 0
        t_counter = 0
        modified_xmls: dict[str, bytes] = {}

        # Process in the exact same order as inspect_docx
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
                            docx_set_paragraph_text(child, str(items[0]))
                            for offset, item_text in enumerate(items[1:], start=1):
                                cloned = copy.deepcopy(child)
                                docx_set_paragraph_text(cloned, str(item_text))
                                container.insert(idx + offset, cloned)
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
                            new_rows = t_op["rows"]
                            existing_trs = child.findall(W_TR)
                            if existing_trs and new_rows:
                                header_tr = existing_trs[0]
                                template_tr = copy.deepcopy(existing_trs[1] if len(existing_trs) > 1 else existing_trs[0])
                                # Update header row (row 0)
                                for c_idx, tc in enumerate(header_tr.findall(W_TC)):
                                    val = str(new_rows[0][c_idx]) if c_idx < len(new_rows[0]) else ""
                                    p_list = tc.findall(W_P)
                                    if p_list:
                                        docx_set_paragraph_text(p_list[0], val)
                                        for extra_p in p_list[1:]:
                                            tc.remove(extra_p)
                                # Remove all data rows after header_tr
                                for old_tr in existing_trs[1:]:
                                    child.remove(old_tr)
                                # Append cloned data rows
                                for row_vals in new_rows[1:]:
                                    new_tr = copy.deepcopy(template_tr)
                                    for c_idx, tc in enumerate(new_tr.findall(W_TC)):
                                        val = str(row_vals[c_idx]) if c_idx < len(row_vals) else ""
                                        p_list = tc.findall(W_P)
                                        if p_list:
                                            docx_set_paragraph_text(p_list[0], val)
                                            for extra_p in p_list[1:]:
                                                tc.remove(extra_p)
                                    child.append(new_tr)
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

        with zipfile.ZipFile(out_docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in modified_xmls:
                    zout.writestr(item, modified_xmls[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))


# ============================================================================
# ODT (OpenDocument Text) Inspection & Manipulation
# ============================================================================

def odt_get_paragraph_text(p_elem: ET.Element) -> str:
    parts: list[str] = []

    def walk(node: ET.Element) -> None:
        if node.text:
            parts.append(node.text)
        for ch in node:
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


def odt_set_paragraph_text(p_elem: ET.Element, new_text: str) -> None:
    """Replace text inside <text:p> or <text:h> while keeping its style and first span's style."""
    first_span = p_elem.find(ODF_SPAN)
    span_style = first_span.get(ODF_STYLE_NAME) if first_span is not None else None

    for ch in list(p_elem):
        p_elem.remove(ch)
    p_elem.text = None

    # Clean inline markdown markers (**bold**, *italic*, `code`) for ODF text nodes
    clean_text = re.sub(r"\*\*([^*]+)\*\*", r"\1", new_text)
    clean_text = re.sub(r"\*([^*]+)\*", r"\1", clean_text)
    clean_text = re.sub(r"`([^`]+)`", r"\1", clean_text)

    target = p_elem
    if span_style:
        target = ET.SubElement(p_elem, ODF_SPAN, {ODF_STYLE_NAME: span_style})

    lines = clean_text.split("\n")
    last_elem: ET.Element | None = None
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


def inspect_odt(odt_path: Path) -> dict:
    blocks: list[dict] = []
    tables: list[dict] = []
    identity_fields: list[dict] = []
    placeholders: set[str] = set()

    with zipfile.ZipFile(odt_path, "r") as zf:
        root = ET.fromstring(zf.read("content.xml"))
        office_text = root.find(f".//{ODF_OFFICE_TEXT}")
        if office_text is None:
            raise ValueError(f"No office:text element found in {odt_path}")

        b_counter = 0
        t_counter = 0

        for child in list(office_text):
            if child.tag in (ODF_P, ODF_H):
                text = odt_get_paragraph_text(child)
                for ph in PLACEHOLDER_RE.findall(text):
                    placeholders.add(ph)
                block_id = f"b{b_counter}"
                b_counter += 1
                blocks.append({
                    "id": block_id,
                    "location": "body",
                    "style": child.get(ODF_STYLE_NAME, "Standard"),
                    "is_list": False,
                    "formatting": ["heading"] if child.tag == ODF_H else [],
                    "text": text,
                })
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
                    blocks.append({
                        "id": block_id,
                        "location": "body",
                        "style": p.get(ODF_STYLE_NAME, "List"),
                        "is_list": True,
                        "formatting": [],
                        "text": text,
                    })
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

    return {
        "format": "odt",
        "source": str(odt_path),
        "placeholders": sorted(placeholders),
        "identity_fields": identity_fields,
        "blocks": blocks,
        "tables": tables,
    }


def render_odt(odt_path: Path, spec: dict, out_odt_path: Path) -> None:
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

    with zipfile.ZipFile(odt_path, "r") as zin:
        root = ET.fromstring(zin.read("content.xml"))
        office_text = root.find(f".//{ODF_OFFICE_TEXT}")
        if office_text is None:
            raise ValueError(f"No office:text element found in {odt_path}")

        b_counter = 0
        t_counter = 0

        for child in list(office_text):
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
                        odt_set_paragraph_text(child, str(items[0]))
                        for offset, item_text in enumerate(items[1:], start=1):
                            cloned = copy.deepcopy(child)
                            odt_set_paragraph_text(cloned, str(item_text))
                            office_text.insert(idx + offset, cloned)
                    elif "text" in op:
                        odt_set_paragraph_text(child, str(op["text"]))

                if placeholders:
                    txt = odt_get_paragraph_text(child)
                    if any(k in txt for k in placeholders):
                        for k, v in placeholders.items():
                            txt = txt.replace(k, str(v))
                        odt_set_paragraph_text(child, txt)

            elif child.tag == ODF_LIST:
                for li in list(child.findall(ODF_LIST_ITEM)):
                    for p in list(li.findall(ODF_P)):
                        block_id = f"b{b_counter}"
                        b_counter += 1
                        op = block_ops.get(block_id)
                        if op:
                            if op.get("delete"):
                                child.remove(li)
                                break
                            if "items" in op and isinstance(op["items"], list):
                                items = op["items"]
                                if not items:
                                    child.remove(li)
                                    break
                                idx = list(child).index(li)
                                odt_set_paragraph_text(p, str(items[0]))
                                for offset, item_text in enumerate(items[1:], start=1):
                                    cloned_li = copy.deepcopy(li)
                                    cloned_p = cloned_li.find(ODF_P)
                                    if cloned_p is not None:
                                        odt_set_paragraph_text(cloned_p, str(item_text))
                                    child.insert(idx + offset, cloned_li)
                            elif "text" in op:
                                odt_set_paragraph_text(p, str(op["text"]))
                        if placeholders:
                            txt = odt_get_paragraph_text(p)
                            if any(k in txt for k in placeholders):
                                for k, v in placeholders.items():
                                    txt = txt.replace(k, str(v))
                                odt_set_paragraph_text(p, txt)

            elif child.tag == ODF_TABLE:
                table_id = f"t{t_counter}"
                t_counter += 1
                t_op = table_ops.get(table_id)
                if t_op:
                    if t_op.get("delete"):
                        office_text.remove(child)
                        continue
                    if "rows" in t_op and isinstance(t_op["rows"], list):
                        new_rows = t_op["rows"]
                        existing_trs = child.findall(ODF_ROW)
                        if existing_trs and new_rows:
                            header_tr = existing_trs[0]
                            template_tr = copy.deepcopy(existing_trs[1] if len(existing_trs) > 1 else existing_trs[0])
                            for c_idx, tc in enumerate(header_tr.findall(ODF_CELL)):
                                val = str(new_rows[0][c_idx]) if c_idx < len(new_rows[0]) else ""
                                p_list = tc.findall(ODF_P)
                                if p_list:
                                    odt_set_paragraph_text(p_list[0], val)
                            for old_tr in existing_trs[1:]:
                                child.remove(old_tr)
                            for row_vals in new_rows[1:]:
                                new_tr = copy.deepcopy(template_tr)
                                for c_idx, tc in enumerate(new_tr.findall(ODF_CELL)):
                                    val = str(row_vals[c_idx]) if c_idx < len(row_vals) else ""
                                    p_list = tc.findall(ODF_P)
                                    if p_list:
                                        odt_set_paragraph_text(p_list[0], val)
                                child.append(new_tr)
                    if "cells" in t_op and isinstance(t_op["cells"], dict):
                        existing_trs = child.findall(ODF_ROW)
                        for rc_key, val in t_op["cells"].items():
                            r_str, c_str = rc_key.split(",")
                            r_i, c_i = int(r_str.strip()), int(c_str.strip())
                            if 0 <= r_i < len(existing_trs):
                                tcs = existing_trs[r_i].findall(ODF_CELL)
                                if 0 <= c_i < len(tcs):
                                    p_list = tcs[c_i].findall(ODF_P)
                                    if p_list:
                                        odt_set_paragraph_text(p_list[0], str(val))

                if placeholders:
                    for p_in_tbl in child.iter(ODF_P):
                        txt = odt_get_paragraph_text(p_in_tbl)
                        if any(k in txt for k in placeholders):
                            for k, v in placeholders.items():
                                txt = txt.replace(k, str(v))
                            odt_set_paragraph_text(p_in_tbl, txt)

        modified_content = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(out_odt_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "content.xml":
                    zout.writestr(item, modified_content)
                else:
                    zout.writestr(item, zin.read(item.filename))


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

        # 1. Inspect .docx
        info = inspect_document(sample_docx)
        assert "{{course}}" in info["placeholders"], f"Missing placeholder: {info}"
        id_labels = {f["label"] for f in info["identity_fields"]}
        assert {"Nama", "NIM", "Age"}.issubset(id_labels), f"Missing identity fields: {id_labels}"

        # 2. Render to both .docx and .odt
        spec = {
            "placeholders": {"{{course}}": "JARINGAN KOMPUTER"},
            "blocks": [
                {"id": "b1", "text": "Nama\t: Andi Wijaya"},
                {"id": "b2", "text": "NIM\t: 24099887"},
                {"id": "b3", "text": "Age\t: 21"},
                {"id": "b4", "items": ["Bullet 1", "Bullet 2", "Bullet 3"]},
                {"id": "b5", "delete": True},
            ],
            "tables": [
                {
                    "id": "t0",
                    "rows": [
                        ["No", "Topic"],
                        ["1", "OSI Layer"],
                        ["2", "TCP/IP"],
                        ["3", "Routing"],
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
        assert "Bullet 3" in texts
        assert "Old extra paragraph to delete" not in texts
        assert rendered_info["tables"][0]["row_count"] == 4
        assert rendered_info["tables"][0]["rows"][3] == ["3", "Routing"]

        # 4. If LibreOffice is present, also test .odt native inspect & render
        if find_soffice():
            out_odt = tmp_dir / "out.odt"
            odt_info = inspect_document(out_odt)
            assert odt_info["tables"][0]["row_count"] == 4
            out_odt2 = tmp_dir / "out2.odt"
            render_document(out_odt, {"blocks": [{"id": "b1", "text": "Nama\t: Citra Dewi"}]}, [out_odt2])
            odt2_info = inspect_document(out_odt2)
            assert "Citra Dewi" in odt2_info["blocks"][1]["text"]

    print("SELF-TEST PASSED: All .docx / .odt inspection, identity detection, cloning, and rendering checks succeeded.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect and fill .doc, .docx, and .odt reference templates.")
    parser.add_argument("--self-test", action="store_true", help="Run built-in verification suite")
    sub = parser.add_subparsers(dest="cmd")

    p_inspect = sub.add_parser("inspect", help="Extract structure, identity fields, blocks, and tables as JSON")
    p_inspect.add_argument("input", type=Path, help="Path to reference .doc, .docx, or .odt")

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
        print(json.dumps(res, indent=2, ensure_ascii=False))
    elif args.cmd == "render":
        spec = json.loads(args.data.read_text(encoding="utf-8"))
        written = render_document(args.input, spec, args.output)
        print(json.dumps({"status": "ok", "outputs": written}, indent=2, ensure_ascii=False))
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
