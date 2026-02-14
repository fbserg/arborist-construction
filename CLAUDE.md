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

**Tracked:** `CLAUDE.md`, `guideline.md`, `docs/`, `work/sample-reports.md`, `.gitignore`
**Ignored:** `*.docx`, `*.xlsx`, `*.ai`, `*.pdf`, `*.emf`, `temp/`, `client/`, `.claude/`

## Project Folder Structure

```
Arborism/
├── CLAUDE.md                         # Technical workflow (this file)  [tracked]
├── guideline.md                      # Content rules for reports  [tracked]
├── .gitignore                        # Keeps work files out of git  [tracked]
├── docs/                             # Detailed reference (read on demand)  [tracked]
│   ├── editing-workflow.md           # Script templates, API ref, XML patterns
│   └── section5-layout.md           # Section 5 structure and XML anchors
├── work/                             # Staging: user drops templates here
│   ├── [Address] Report.docx         # Template placed here by user  [ignored]
│   └── sample-reports.md             # Reference documentation  [tracked]
├── temp/                             # Working directory for document editing  [ignored]
│   └── [# Word]/                     # Short signifier (e.g., 71 Lloyd, 45 King)
│       ├── unpacked/                 # Unpacked OOXML for current document
│       ├── changelog.md              # Log of edits made to documents
│       └── [Document Name].md        # Readable markdown copy
├── client/                           # Re-compiled reports for manual review  [ignored]
│   ├── [Address] Report.docx         # Packed output, flat (no subfolders)
│   └── ...
```

`[project]` in paths below refers to the temp signifier folder (e.g., `temp\71 Lloyd`).

## Tool Paths (Windows)

All commands in this file use **PowerShell syntax**. Run them directly in the Bash tool (which already runs PowerShell). Do **not** wrap in `powershell -Command "..."`.

**Session setup:** Run once at the start of each editing session to define shorthand variables.
```powershell
$PY = "C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe"
$PANDOC = "C:\Users\User\AppData\Local\Pandoc\pandoc.exe"
$SKILL = "C:\Users\User\.claude\plugins\cache\anthropic-agent-skills\document-skills\69c0b1a06741\skills\docx"
$UNPACK = "$SKILL\ooxml\scripts\unpack.py"
$PACK = "$SKILL\ooxml\scripts\pack.py"
```

## Working with Reports

Two modes of operation:

- **New project from template** (Section 0a): User places a template in `work/`. Claude unpacks to `temp/[signifier]`, fills in site-specific content using tracked changes per `guideline.md`, then packs to `client/`.
- **Editing an existing report** (Sections 0b–4): Report already has content. Claude reads, locates text, and makes tracked-change edits.

### 0a. New Project from Template

Use when starting a new report. The user will place a template `.docx` into `work/`.

**Step 1: Create project temp folder and unpack**
```powershell
New-Item -ItemType Directory -Force "C:\Projects\Arborism\temp\[signifier]"
& $PY $UNPACK "C:\Projects\Arborism\work\[Address] Report.docx" "[project]\unpacked\temp"
# Note the suggested RSID
```

**Step 2: Read guideline.md and user-supplied project data**

Consult `C:\Projects\Arborism\guideline.md` for impact profile, species tolerance, narrative template, TPZ calculations, and post-removal notes.

**Step 3: Fill in sections using tracked changes**

Use the Section 2 editing workflow. All changes must be tracked (`<w:del>` on placeholder, `<w:ins>` on new text). Leave unchanged sections as-is.

**Step 4: Pack, verify, and log**

Follow Section 2 Steps 4–6.

### 0b. Before Reading a Document

```powershell
Get-ChildItem -Recurse -Filter "*.docx" "C:\Projects\Arborism"
```

Ask what the user needs before reading:
- **Specific edits**: Skip full read. Grep to locate text, then edit directly.
- **Injury/removal edits**: Read only Section 4 (tree data table) and the target section.
- **Review or audit**: Read the full document.
- **Unsure**: Ask clarifying questions.

### 1. Reading a Document

```powershell
ls "[project]\"                                                          # Check for cached markdown copy
& $PANDOC "path\to\document.docx" -t markdown                           # Quick read
& $PANDOC "path\to\document.docx" -t markdown -o "[project]\[Name].md"  # Save readable copy
& $PANDOC --track-changes=all "path\to\document.docx" -t markdown       # View tracked changes
```

### 2. Editing a Document (Tracked Changes Workflow)

**Before editing, read `docs/editing-workflow.md`** for the script template, editor API, and XML patterns.

**Step 1:** Unpack: `& $PY $UNPACK "path\to\document.docx" "[project]\unpacked\temp"` — note the suggested RSID.
**Step 2:** Create edit script in `[project]\edit_script.py` using the template from `docs/editing-workflow.md`.
**Step 3:** Run: `& $PY "[project]\edit_script.py"`
**Step 4:** Pack: `& $PY $PACK "[project]\unpacked\temp" "C:\Projects\Arborism\client\[Address] Report.docx"`
**Step 5:** Verify: `& $PANDOC --track-changes=all "C:\Projects\Arborism\client\[Address] Report.docx" -t markdown | Select-String -Pattern "expected change" -Context 2`
**Step 6:** Log edits in `[project]\changelog.md`.
**Step 7:** Promote: `Remove-Item -Recurse -Force "[project]\unpacked\current" -ErrorAction SilentlyContinue; Rename-Item "[project]\unpacked\temp" "current"`

### 3. Key Editing Principles

- **Always track changes**: Use `<w:del>`/`<w:ins>` tags. Never make untracked modifications. All new content must be wrapped in `<w:ins>` — the API does not create these automatically.
- **Scope lock**: Only edit the section specified by the user.
- **Minimal edits**: Only wrap changed text in `<w:del>`/`<w:ins>` tags.
- **Preserve formatting**: Extract and reuse `<w:rPr>` from original nodes.
- **Batch changes**: Group 3-10 related edits per script.
- **Grep first**: Always check `word/document.xml` line numbers before editing.
- **Section 5 edits**: Read `docs/section5-layout.md` before editing the conclusion section.

### 4. Reset Working Directory

```powershell
Remove-Item -Recurse -Force "[project]\unpacked" -ErrorAction SilentlyContinue
& $PY $UNPACK "path\to\new-document.docx" "[project]\unpacked\temp"
```

Always unpack from the original source in `work/` (the clean template), not from the project folder.

## Dependencies

Required (installed):
- **Python 3.12**: `C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe`
- **pandoc**: `C:\Users\User\AppData\Local\Pandoc\pandoc.exe`
- **defusedxml** and **lxml** (pip installed)

Optional: **LibreOffice** (PDF conversion), **Poppler** (PDF to image)
