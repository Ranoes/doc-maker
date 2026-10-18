---
description: Generate .docx and .odt documents based on an existing .doc, .docx, or .odt reference template
---

# `/doc-maker` Workflow

Read and follow the [`doc-maker` skill instructions](../../skills/doc-maker/SKILL.md):

1. **Inspect Reference Template**:
   Run `python skills/doc-maker/scripts/doc_maker.py inspect "<reference_doc>"` to extract `placeholders`, `identity_fields`, `blocks`, and `tables`.
2. **Missing Identification Check & Confirmation Gate (MANDATORY)**:
   - Check if any identification fields in the template (`Name`/`Nama`, `NIM`/`NPM`, `Age`/`Umur`, `Class`/`Kelas`, `Department`/`Prodi`, `Date`/`Tanggal`, etc.) are missing from the user's provided material. If missing, **ask the user what to fill**.
   - Display the complete resolved identification list to the user and ask: *"Is this identification information correct before I generate the document?"* Wait for confirmation.
3. **Render Document(s)**:
   Create the JSON render spec and run `python skills/doc-maker/scripts/doc_maker.py render "<reference_doc>" -d spec.json -o <output.docx> [-o <output.odt>]`.
4. **Verify**:
   Run `inspect` on the generated output file(s) to confirm all fields and sections are properly populated.
