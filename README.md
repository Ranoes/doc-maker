# doc-maker

An AI Agent Skill & zero-dependency Python CLI tool that generates `.docx` and `.odt` documents using any existing `.doc`, `.docx`, or `.odt` file as a reference template—preserving exact styles, fonts, margins, headers/footers, bullet lists, and table formatting.

## Features

- **Use Any Document as a Template**: Works with both explicit `{{placeholder}}` templates and normal finished documents (reports, letters, assignments, proposals).
- **100% Style Preservation**: Clones and edits underlying OOXML (`word/document.xml`) and ODF (`content.xml`) nodes in-place so fonts, sizes, numbering/bullets, headers/footers, and table borders stay intact.
- **Rich Content & Markdown Material**: Supports inline Markdown (`**bold**`, `*italic*`, `` `code` ``), automatic image embedding (`PNG`/`JPEG` with aspect-ratio scaling and captions), dynamic table column resizing, and line-numbered source code tables across both `.docx` and `.odt`.
- **Smart Identification Gate**: Automatically detects personal/document identification fields (`Name`/`Nama`, `NIM`/`NPM`/`Student ID`, `Age`/`Umur`, `Class`/`Kelas`, `Department`/`Prodi`, `Date`, etc.). If any field is missing from the given material, the agent asks what to fill and always asks you to confirm the identification details before generating the final document.
- **Zero Pip Dependencies**: Uses only the Python 3 standard library (`zipfile`, `xml.etree.ElementTree`, `json`, `subprocess`, `struct`). Uses LibreOffice (`soffice --headless`) when converting legacy `.doc` files or cross-converting between `.docx` and `.odt`.

---

## Installation

### 1. Antigravity (`agy`)
```bash
agy plugin install https://github.com/Ranoes/doc-maker
```
Or clone directly into your global plugins folder (so it works in **every** workspace):
```bash
git clone https://github.com/Ranoes/doc-maker ~/.gemini/config/plugins/doc-maker
```

### 2. OpenAI Codex (CLI & Desktop App)
Install via Codex plugin marketplace:
```bash
codex plugin marketplace add Ranoes/doc-maker
codex plugin add doc-maker@doc-maker
```
Or install directly into your global Codex skills directory (`~/.codex/skills`):
```bash
git clone https://github.com/Ranoes/doc-maker ~/.codex/doc-maker-repo
mkdir -p ~/.codex/skills
cp -r ~/.codex/doc-maker-repo/skills/doc-maker ~/.codex/skills/doc-maker
```
*(On Windows PowerShell)*:
```powershell
git clone https://github.com/Ranoes/doc-maker "$HOME\.codex\doc-maker-repo"
New-Item -ItemType Directory -Force -Path "$HOME\.codex\skills" | Out-Null
Copy-Item -Recurse -Force "$HOME\.codex\doc-maker-repo\skills\doc-maker" "$HOME\.codex\skills\doc-maker"
```

### 3. Gemini CLI
```bash
gemini extensions install https://github.com/Ranoes/doc-maker
```

### 4. Claude Code
```bash
/plugin marketplace add Ranoes/doc-maker
/plugin install doc-maker@doc-maker
```

### 5. GitHub Copilot CLI
```bash
copilot plugin marketplace add Ranoes/doc-maker
copilot plugin install doc-maker@doc-maker
```

### 6. Project-Level Skill (Any Agent)
Copy `skills/doc-maker` into your project's `.agents/skills/doc-maker` or `.codex/skills/doc-maker` directory.

---

## How to Use in Chat (`/doc-maker`)

Once installed, you can invoke it directly with the **`/doc-maker`** slash command (or `$doc-maker` / natural language in Codex):

### Example Commands
- **Using `/doc-maker`**:
  ```text
  /doc-maker template_laporan.docx make a new .docx and .odt using @materi.md with images in @gambar/
  ```
  ```text
  /doc-maker reference.doc create output.docx from these notes: [paste material]
  ```
- **Or natural language (auto-detected)**:
  > *"Make a new `.docx` and `.odt` report based on `template_laporan.docx` using the material in `materi.md`."*

### What Happens Next (Automatic Workflow)
1. **Inspects Template**: The agent runs `doc_maker.py inspect "<template>" -o inspect.json` on your reference `.doc` / `.docx` / `.odt`.
2. **Asks & Confirms Identification**:
   - If the template has fields like `Nama`, `NIM`, `Age`, `Kelas`, `Tanggal`, etc. that aren't in your material, the agent **stops and asks you** what to fill.
   - Before generating the file, the agent displays the identification summary and asks: *"Is this identification information correct before I generate the document?"*
3. **Generates `.docx` / `.odt`**: Once you confirm, it generates the new document(s) with the exact formatting, styles, embedded images, bullets, and tables of the reference doc.

---

## Requirements

- **Python 3.8+** (no third-party `pip` packages required)
- **LibreOffice** *(optional)*: Only needed when reading legacy binary `.doc` files or converting between `.docx` and `.odt`. Same-format generation (`.docx` → `.docx` or `.odt` → `.odt`) runs in pure Python.

---

## Standalone CLI Usage

### 1. Inspect a Reference Document
```bash
python skills/doc-maker/scripts/doc_maker.py inspect reference.docx -o inspect.json
```

### 2. Render New `.docx` and `.odt` Documents
```bash
python skills/doc-maker/scripts/doc_maker.py render reference.docx -d spec.json -o output.docx -o output.odt
```

### 3. Run Self-Test
```bash
python skills/doc-maker/scripts/doc_maker.py --self-test
```
