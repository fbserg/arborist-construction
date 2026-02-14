# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

This repository contains arborist consulting project files. Claude instances working here will create and edit arborist reports in Word (.docx) format, maintaining consistent professional style across all documents.

## Content Guidelines

**Always read `guideline.md` at the start of any session involving report content.** It defines the data table schema, impact profiles (A–E), species tolerance modifiers, narrative templates, TPZ calculations, tone constraints, and post-removal notes. This file covers only the technical workflow; `guideline.md` governs what goes into the report.

## Available Skills

The **docx skill** (`/docx`) is installed and provides the workflows below for all Word document operations.

## Version Control

This repo is pushed to GitHub. **Only guidelines and documentation are tracked in git.** All binary/work files live on disk but are gitignored (see `.gitignore`).

**Tracked:** `CLAUDE.md`, `guideline.md`, `work/sample-reports.md`, `.gitignore`
**Ignored:** `*.docx`, `*.xlsx`, `*.ai`, `*.pdf`, `*.emf`, `**/.work/`, `.claude/`, `**/Visit/`, `**/media/`

## Project Folder Structure

```
Arborism/
├── CLAUDE.md                         # Technical workflow (this file)  [tracked]
├── guideline.md                      # Content rules for reports  [tracked]
├── .gitignore                        # Keeps work files out of git  [tracked]
├── work/                             # Staging: one active template at a time
│   ├── [Address] Report.docx         # Template placed here by user  [ignored]
│   └── sample-reports.md             # Reference documentation  [tracked]
├── [Address, City, Province, Country, Postal Code]/
│   ├── .work/                        # Working directory for document editing  [ignored]
│   │   ├── unpacked/                 # Unpacked OOXML for current document
│   │   ├── changelog.md              # Log of edits made to documents
│   │   └── [Document Name].md        # Readable markdown copy
│   ├── [YYYY.MM.DD Description]/     # Date-stamped project phases
│   │   ├── Work/                     # Internal working files
│   │   │   ├── Visit/                # Site visit photos
│   │   │   └── [reference PDFs, plans, drawings]
│   │   └── Client/                   # Deliverables sent to client
│   │       ├── [Address] Report.pdf
│   │       └── [Address] Plan.pdf
│   ├── [Address] Report.docx         # Current working report (master)  [ignored]
│   ├── [Address] Plan.ai             # Current working plan (Illustrator)  [ignored]
│   └── [Address] inventory.xlsx      # Tree inventory spreadsheet  [ignored]
```

`[project]` in paths below refers to the address folder (e.g., `123 Fake Street, Toronto, ON, Canada, M5V 1A1`).

## Tool Paths (Windows)

All commands in this file use **PowerShell syntax**. When calling from the Bash tool, wrap in `powershell -Command "..."`.

```
PYTHON    = C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe
PANDOC    = C:\Users\User\AppData\Local\Pandoc\pandoc.exe
SKILL_ROOT = C:\Users\User\.claude\plugins\cache\anthropic-agent-skills\document-skills\69c0b1a06741\skills\docx
```

**Session setup:** Run once at the start of each editing session to define shorthand variables. All commands below use these.
```powershell
$PY = "C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe"
$PANDOC = "C:\Users\User\AppData\Local\Pandoc\pandoc.exe"
$SKILL = "C:\Users\User\.claude\plugins\cache\anthropic-agent-skills\document-skills\69c0b1a06741\skills\docx"
$UNPACK = "$SKILL\ooxml\scripts\unpack.py"
$PACK = "$SKILL\ooxml\scripts\pack.py"
```

## Working with Reports

There are two modes of operation:

- **New project from template** (Section 0a): User places a template in `work/`. Claude copies it to the project folder, unpacks, fills in site-specific content using tracked changes per `guideline.md`, then packs.
- **Editing an existing report** (Sections 0b–4): Report already has content. Claude reads, locates text, and makes tracked-change edits.

### 0a. New Project from Template

Use when starting a new report. The user will place a template `.docx` into `work/`.

**Step 1: Copy template to project folder**
```powershell
Copy-Item "C:\Projects\Arborism\work\[Address] Report.docx" "[project]\[Address] Report.docx"
```

**Step 2: Create .work directory and unpack to temp**
```powershell
New-Item -ItemType Directory -Force "[project]\.work"
& $PY $UNPACK "[project]\[Address] Report.docx" "[project]\.work\unpacked\temp"
# Note the suggested RSID
```

**Step 3: Read guideline.md and user-supplied project data**

Consult `C:\Projects\Arborism\guideline.md` to determine:
- Impact profile (A–E) for the proposed work
- Species tolerance modifier per tree
- Narrative template (1, 2, or 3) per tree
- TPZ calculations and encroachment
- Applicable post-removal/conclusion notes

**Step 4: Fill in sections using tracked changes**

Use the Section 2 editing workflow. Replace placeholder/boilerplate text with site-specific content. All changes must be tracked (`<w:del>` on placeholder, `<w:ins>` on new text). Leave unchanged sections as-is.

**Step 5: Pack, verify, and log**

Follow Section 2 Steps 4–6.

### 0b. Before Reading a Document

**Step 0: Locate the document**
```powershell
Get-ChildItem -Recurse -Filter "*.docx" "C:\Projects\Arborism"
```
Use recursive search rather than assuming path depth.

Ask what the user needs before reading a document.

- **Specific edits** (typos, corrections, known changes): Skip full read. Use grep to locate text, then edit directly.
- **Injury/removal edits**: Read only Section 4 (tree data table) and the target section (e.g., Section 5 conclusion). Do not read the full document.
- **Review or audit**: Read the full document.
- **Unsure**: Ask clarifying questions.

### 1. Reading a Document

**Check for cached copy first:**
```powershell
ls "[project]\.work\"  # Check if markdown copy already exists
```

**Quick read (markdown output):**
```powershell
& $PANDOC "path\to\document.docx" -t markdown
```

**Save readable copy to .work (always do this when reading fully):**
```powershell
& $PANDOC "path\to\document.docx" -t markdown -o "[project]\.work\[Document Name].md"
```

**View tracked changes:**
```powershell
& $PANDOC --track-changes=all "path\to\document.docx" -t markdown
```

### 2. Editing a Document (Tracked Changes Workflow)

**Step 1: Unpack the document to temp**
```powershell
& $PY $UNPACK "path\to\document.docx" "[project]\.work\unpacked\temp"
# Note the suggested RSID (Revision Session ID - groups related tracked changes together)
```

**Step 2: Create edit script**

Standard script header (always use):
```python
import sys, os
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')
SKILL = r"C:\Users\User\.claude\plugins\cache\anthropic-agent-skills\document-skills\69c0b1a06741\skills\docx"
sys.path.insert(0, SKILL)
from scripts.document import Document

PROJECT = r"[project]"
doc = Document(PROJECT + r"\.work\unpacked\temp", author="Claude", rsid="SUGGESTED_RSID")
editor = doc["word/document.xml"]
```

Editor API reference:
```
editor.get_node(tag, contains)   — find node by tag name and text content
editor.replace_node(node, xml)   — replace node with new XML
editor.insert_after(node, xml)   — insert XML after element
editor.insert_before(node, xml)  — insert XML before element
editor.append_to(node, xml)      — append XML as child of element

All methods accept multi-element XML strings.
Auto-inject RSID attributes to <w:p> and <w:r> elements.
Auto-populate w:id/w:author/w:date on <w:ins>/<w:del> — but does NOT create them.
```

**Critical: `<w:ins>` wrapping is manual.** The API populates attributes on `<w:ins>`/`<w:del>` elements but does **not** auto-create them. All new content must be wrapped in `<w:ins>`:
```xml
<!-- WRONG — will not show as tracked change -->
<w:r><w:rPr>...</w:rPr><w:t>new text</w:t></w:r>

<!-- CORRECT — will show as tracked insertion -->
<w:ins><w:r><w:rPr>...</w:rPr><w:t>new text</w:t></w:r></w:ins>
```

**XML helper patterns** — use these in edit scripts to reduce boilerplate:
```python
# Tracked insertion paragraph (body text)
def ins_para(text, rpr):
    return f'''<w:p>
  <w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr>
  <w:ins><w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:ins>
</w:p>'''

# Standard 4-side border (for table cells)
BORDERS = '''<w:tcBorders>
  <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tcBorders>'''
```

For replacements, use `<w:del>` + `<w:ins>`:
```python
node = editor.get_node(tag="w:r", contains="text to find")
rpr = tags[0].toxml() if (tags := node.getElementsByTagName("w:rPr")) else ""
replacement = f'<w:del><w:r>{rpr}<w:delText>old</w:delText></w:r></w:del><w:ins><w:r>{rpr}<w:t>new</w:t></w:r></w:ins>'
editor.replace_node(node, replacement)

doc.save()
```

**Step 3: Run the edit script**
```powershell
& $PY "[project]\.work\edit_script.py"
```
UTF-8 encoding is handled in the script header (see Step 2 template). Do **not** rely on `$env:PYTHONIOENCODING` — it does not work through the bash→powershell bridge.

**Step 4: Pack back to docx**
```powershell
& $PY $PACK "[project]\.work\unpacked\temp" "path\to\output.docx"
```

**Step 5: Verify changes**
```powershell
& $PANDOC --track-changes=all "path\to\output.docx" -t markdown | Select-String -Pattern "expected change" -Context 2
```

**Step 6: Log edits**

Append to `[project]\.work\changelog.md`:
```markdown
## YYYY-MM-DD
- Changed "old text" → "new text" in Section X
```

**Step 7: Promote temp to current**
```powershell
Remove-Item -Recurse -Force "[project]\.work\unpacked\current" -ErrorAction SilentlyContinue
Rename-Item "[project]\.work\unpacked\temp" "current"
```

### 3. Key Editing Principles

- **Always track changes**: Every edit to a Word document must use tracked changes (`<w:del>`/`<w:ins>` tags). Never make untracked modifications. All new/inserted content must be wrapped in `<w:ins>` — the API does not create these automatically.
- **Scope lock**: Only edit the section specified by the user. Do not modify any other section under any circumstances.
- **Minimal edits**: Only wrap changed text in `<w:del>`/`<w:ins>` tags
- **Preserve formatting**: Extract and reuse `<w:rPr>` from original nodes
- **Batch changes**: Group 3-10 related edits per script
- **Grep first**: Always check `word/document.xml` line numbers before editing
- **Validation output**: The `Paragraphs: X → Y` count in save output may show misleading numbers (e.g., large drops). This is normal when inserting tables. Verify content with pandoc output instead.

### 3a. Section 5 (Conclusion) Layout

Section 5 follows this order. Do not reorder or skip structural elements:

1. **Permit summary** — Bullet list of removal/injury permit counts
2. **Protection note** — Standard paragraph re: TPZ fencing
3. **No other trees note** — Standard paragraph
4. **Impact summary table** — Tree #, Source of impact, Direction (one row per tree, multi-row sources merged)
5. **"Injuries" subheading**
   - Per injured tree:
     a. **Mini data table** — TREE #, Species, DBH, Condition, Ownership, Direction, TPZ, % Encroachment, Permit
     b. **Injury detail table** — Injury source, Closest point of impact, Max depth, Impact to condition (one row per injury source)
     c. **Narrative paragraphs** — Written per `guideline.md` Template 1
   - After all injury narratives: RSE boilerplate paragraphs
6. **"Removals" subheading** (if applicable)
   - Per removed tree: narrative per Template 2 or 3
7. **"General Notes"** — Standard closing paragraphs
8. **Signature block**

Injury narratives go **after** the injury detail table, **before** the RSE boilerplate. They do NOT go after the RSE boilerplate or in the empty paragraphs between RSE and Removals.

**Standard XML anchors** (use with `editor.get_node` to locate insertion points):
| Anchor text | Locates |
|---|---|
| `"Injuries"` | Injuries subheading paragraph |
| `"Injury source"` | Injury detail table header row |
| `"Excavation should not be deeper"` | RSE boilerplate (insert narrative before this) |
| `"Removals"` | Removals subheading |

### 4. Reset Working Directory

To start fresh or switch documents:
```powershell
Remove-Item -Recurse -Force "[project]\.work\unpacked" -ErrorAction SilentlyContinue
& $PY $UNPACK "path\to\new-document.docx" "[project]\.work\unpacked\temp"
```

When re-doing edits, always unpack from the original source in `work/` (the clean template), not from the project folder (which may contain previous edits).

## Dependencies

Required (installed):
- **Python 3.12**: `C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe`
- **pandoc**: `C:\Users\User\AppData\Local\Pandoc\pandoc.exe`
- **defusedxml**: `pip install defusedxml`
- **lxml**: `pip install lxml`

Optional:
- **LibreOffice**: For PDF conversion (`soffice --headless --convert-to pdf`)
- **Poppler**: For PDF to image (`pdftoppm`)
