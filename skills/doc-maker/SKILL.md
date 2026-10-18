---
name: doc-maker
description: >-
  Generate .docx and .odt documents based on an existing .doc, .docx, or .odt
  reference template. Use whenever the user wants to create, fill, or format a
  Word (.docx) or OpenDocument (.odt) file using an existing document as a
  template or style/format reference, or when the user invokes /doc-maker.
  Trigger: /doc-maker, "doc-maker".
argument-hint: "<reference_doc> [material_or_output]"
---

# Doc Maker (`doc-maker`)

Create `.docx` and `.odt` files from any reference `.doc`, `.docx`, or `.odt` document while preserving exact styles, fonts, margins, headers/footers, bullet lists, and table formatting.

The helper script is located at [`scripts/doc_maker.py`](./scripts/doc_maker.py) relative to this `SKILL.md` file.

---

## Mandatory Workflow

Follow these steps in order every time:

### Step 1: Inspect the Reference Document

Run `inspect` on the reference `.doc`, `.docx`, or `.odt` file to extract its structure, `placeholders`, `identity_fields`, `blocks`, `tables`, and `tocs` (use `-o <path.json>` for large templates to write UTF-8 JSON directly to disk without shell encoding or truncation issues):

```bash
python <path-to-skill>/scripts/doc_maker.py inspect "<reference_doc>" -o inspect.json
```

The JSON output contains:
- `placeholders`: Explicit tags like `{{name}}` or `[TITLE]`.
- `identity_fields`: Automatically detected personal/document identification fields (e.g., `Name`/`Nama`, `NIM`/`NPM`/`Student ID`, `Age`/`Umur`, `Class`/`Kelas`, `Department`/`Jurusan`/`Prodi`, `Lecturer`/`Dosen`, `Date`/`Tanggal`, `Email`, `Phone`, etc.) with their `id`, `label`, `current_value`, and `raw_text`.
- `blocks`: Every paragraph/heading in the document (`b0`, `b1`, ..., plus header/footer blocks `h0_b0`, `f0_b0`, ...) with its `style`, formatting summary, `has_images` flag (when the block contains an embedded image/logo), and `text`.
- `tables`: Every table (`t0`, `t1`, ...) with row/column counts and 2D `rows` cell array.
- `tocs`: Table of Contents / List of Figures / List of Tables sections (`toc0`, `toc1`, ...) with their index entries.

---

### Step 2: Check & Confirm Identification Fields (MANDATORY STOP-AND-ASK GATE)

Before generating any document, inspect all `identity_fields`, `placeholders`, cover page blocks, headers, and metadata tables for identification fields (such as **Name / Nama**, **NIM / NPM / ID**, **Age / Umur**, **Class / Kelas**, **Course / Mata Kuliah**, **Department / Prodi / Jurusan**, **Institution**, **Date / Tanggal**, etc.):

1. **Ask for Missing Identification**:
   - Compare the template's identification fields against the user's provided material.
   - If **any** identification field in the template is not mentioned in the given material, **STOP and ask the user** what value to fill for each missing field.
   - **NEVER** guess or fabricate personal data, and **NEVER** leave the reference document's original author/student data in place.
2. **Confirm Identification Before Generating**:
   - Even if all identification fields were provided (or right after the user supplies the missing ones), present a clear summary table/list of the resolved identification fields (e.g. `Name`, `NIM`, `Age`, etc.) to the user and ask:
     > *"Please confirm: is this identification information correct before I generate the document?"*
   - **Wait for the user's explicit confirmation** before running `render`.

---

### Step 3: Build the Render Spec JSON

Once the user confirms the identification details, create a temporary JSON specification file (e.g., `spec.json`) mapping the confirmed identity fields and new material onto the reference template:

```json
{
  "placeholders": {
    "{{nama}}": "Siti Aminah",
    "{{nim}}": "22010456"
  },
  "blocks": [
    {
      "id": "b1",
      "text": "Nama\t: Siti Aminah"
    },
    {
      "id": "b2",
      "text": "NIM\t: 22010456"
    },
    {
      "id": "b5",
      "items": [
        "First paragraph or bullet item using b5's exact XML style.",
        "Second paragraph or bullet item cloned from b5.",
        "Third paragraph with **bold** and *italic* inline runs."
      ]
    },
    {
      "id": "b6",
      "delete": true
    }
  ],
  "tables": [
    {
      "id": "t0",
      "rows": [
        ["No", "Component", "Status"],
        ["1", "Authentication Module", "Completed"],
        ["2", "Document Engine", "Completed"],
        ["3", "Unit Tests", "Verified"]
      ]
    },
    {
      "id": "t1",
      "cells": {
        "0,1": "Siti Aminah",
        "1,1": "22010456"
      }
    }
  ]
}
```

#### Spec Rules:
- **Markdown (`.md`) Material Support**:
  - When the user provides a `.md` file as material, strip block-level Markdown prefixes (`# `, `## `, `- `, `* `, `1. `) before placing text into `blocks` or `items` (since the reference document's XML styles already provide heading formatting and bullet/numbering markers).
  - You **may** keep inline `**bold**`, `*italic*`, and `` `code` `` inside `text` / `items` / table cells—`doc_maker.py` automatically converts them into native styled XML runs!
  - Map Markdown tables (`| A | B |`) directly into 2D arrays in `tables[].rows` or inline `{"type": "table", "rows": [...]}` inside `blocks[].items`.
- **`placeholders`**: Replaces matching substrings across all body paragraphs, tables, headers, and footers (handles tags split across multiple XML runs).
- **`blocks`**:
  - **Preserving Cover Logos / Embedded Images**: Blocks with `"has_images": true` in `inspect` output (such as cover-page logos) contain embedded drawings. **Never** delete (`"delete": true`) or replace them with `"items"` unless you intend to replace the image.
  - `{"id": "bX", "text": "..."}`: Replaces the text of block `bX` while keeping its exact paragraph style (`<w:pPr>` / `text:style-name`), run formatting (`<w:rPr>`), and bookmarks/page-breaks. If the original block used `\t` (tabs) around `:` (e.g. `Nama\t: Budi`), preserve `\t` in the replacement string so alignment stays identical.
  - `{"id": "bX", "items": [...]}`: Replaces `bX` with a sequence of strings or rich item objects:
    - Plain string `"Paragraph text"` or `{"type": "p", "style": "P33", "text": "..."}`: Clones `bX` (or uses the specified paragraph style).
    - `{"type": "image", "path": "C:/path/to/img.png", "caption": "Gambar 2.1 ...", "max_width_in": 5.4}` (or Markdown string `"![Gambar 2.1 ...](C:/path/to/img.png)"`): Embeds an image with binary header detection (`PNG`/`JPEG`), aspect-ratio scaling, optional caption paragraph, and `META-INF/manifest.xml` registration.
    - `{"type": "table", "caption": "Tabel 2.2 ...", "template_table": "t6", "col_widths": [1.3, 2.0, 2.2], "rows": [[...], [...]]}`: Clones a template table (`template_table`) and populates it with dynamic column count and widths.
    - `{"type": "code_table", "title": "Kode Sumber 2.1: ...", "code": "{\n  ...\n}"}`: Clones a 2-column line-numbered source code table (`t5` / `Table6`) with automatic line numbering and indentation preservation.
    - `{"type": "heading", "level": 3, "list_style": "WWNum8", "style": "P71", "continue": true, "text": "Sub-section Title"}`: Inserts a numbered heading list item matching the template's multilevel heading list.
    - `{"type": "list", "list_style": "WWNum7", "items": ["Item 1", "Item 2"]}`: Inserts a numbered or bulleted list.
  - `{"id": "bX", "delete": true}`: Deletes an unneeded block from the reference document (e.g. template instructions or sample sections).
- **`tables`**:
  - `{"id": "tX", "col_widths": [...], "rows": [[...], [...]]}`: Replaces the entire table's content and dynamically adjusts column count and widths if the new table has more or fewer columns than the template table.
  - `{"id": "tX", "cells": {"row,col": "New text"}}`: Updates specific cells by `row,col` index without changing the number of rows.
  - `{"id": "tX", "delete": true}`: Deletes an unneeded sample table from the template.
- **`toc_updates`** *(ODT templates with `<text:table-of-content>`)*:
  - Maps `"toc0"`, `"toc1"`, etc. (from `inspect` output's `"tocs"`) to a list of `{"style": "P15", "text": "Tabel 2.1 ...", "page": "8", "href": "#_bm_tbl_2_1"}` entries, or use `"insert_toc_after"` on a heading block (such as `DAFTAR ISI`) to insert a populated TOC section.

---

### Step 4: Render `.docx` and/or `.odt`

Run `render` with one or more `-o` output paths (`.docx`, `.odt`, or both):

```bash
python <path-to-skill>/scripts/doc_maker.py render "<reference_doc>" -d spec.json -o output.docx -o output.odt
```

---

### Step 5: Verify Output

Run `inspect` on the generated file(s) to verify that all confirmed identification fields and material sections were written properly and no stale reference text remains:

```bash
python <path-to-skill>/scripts/doc_maker.py inspect output.docx
```
