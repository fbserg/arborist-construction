---
name: editing-arborist-reports
description: "Creates and edits arborist consulting reports in .docx format using tracked changes. Use when creating a new report from a template, making scoped edits to an existing report, or reading report content."
---

# Editing Arborist Reports

## Contents

- Session setup
- Decision tree (choose workflow)
- New report from template
- Editing an existing report
- Reading a report
- Key editing principles
- Verification and logging
- Reference files (load as needed)

## Session Setup

Run once at the start of each editing session:

```powershell
. "C:\Projects\Arborism\.agents\skills\editing-arborist-reports\scripts\setup.ps1"
```

## Project Folder Structure

```
Arborism/
├── work/           # User drops templates here
├── temp/           # Working directory
│   └── [signifier]/    # e.g., 71 Lloyd
│       ├── unpacked/   # Unpacked OOXML
│       ├── changelog.md
│       └── [Name].md   # Readable markdown copy
├── client/         # Packed output for review
```

`[project]` in paths below = the temp signifier folder (e.g., `temp\71 Lloyd`).

## Decision Tree

Ask what the user needs:

1. **New report from template** → Section: New Report
2. **Specific edits** → Skip full read. Grep to locate text, then edit directly
3. **Injury/removal edits** → Read only Section 4 (tree data table) and the target section
4. **Review or audit** → Read the full document
5. **Unsure** → Ask clarifying questions

## New Report from Template

**Step 1:** Create project folder and unpack:
```powershell
New-Item -ItemType Directory -Force "C:\Projects\Arborism\temp\[signifier]"
& $PY $UNPACK "C:\Projects\Arborism\work\[Address] Report.docx" "[project]\unpacked\temp"
```
Note the suggested RSID.

**Step 2:** Read content rules. Load `C:\Projects\Arborism\guideline.md` for impact profiles, species tolerance, narrative templates, TPZ calculations, and post-removal notes.

**Step 3:** Fill sections using tracked changes per the Editing workflow below.

**Step 4:** Pack, verify, and log per the Verification section below.

## Reading a Report

```powershell
ls "[project]\"                                                          # Check for cached markdown
& $PANDOC "path\to\document.docx" -t markdown                           # Quick read
& $PANDOC "path\to\document.docx" -t markdown -o "[project]\[Name].md"  # Save readable copy
& $PANDOC --track-changes=all "path\to\document.docx" -t markdown       # View tracked changes
```

## Editing a Report (Tracked Changes)

**Before editing:** Load `reference/editing-tracked-changes.md` for the script template, Editor API, and XML patterns.

**Step 1 — Unpack:**
```powershell
& $PY $UNPACK "path\to\document.docx" "[project]\unpacked\temp"
```
Note the suggested RSID.

**Step 2 — Create edit script** in `[project]\edit_script.py` using the template from `reference/editing-tracked-changes.md`.

**Step 3 — Run:**
```powershell
& $PY "[project]\edit_script.py"
```

**Step 4–7:** Follow Verification section below.

## Key Editing Principles

- **Always track changes**: Use `<w:del>`/`<w:ins>` tags. Never make untracked modifications. All new content must be wrapped in `<w:ins>` — the API does not create these automatically.
- **Scope lock**: Only edit the section specified by the user.
- **Minimal edits**: Only wrap changed text in `<w:del>`/`<w:ins>` tags.
- **Preserve formatting**: Extract and reuse `<w:rPr>` from original nodes.
- **Batch changes**: Group 3-10 related edits per script.
- **Grep first**: Always check `word/document.xml` line numbers before editing.
- **Section 5 edits**: Load `reference/section5-layout.md` before editing the conclusion section.

## Verification and Logging

**Step 4 — Pack:**
```powershell
& $PY $PACK "[project]\unpacked\temp" "C:\Projects\Arborism\client\[Address] Report.docx"
```

**Step 5 — Verify:**
```powershell
& $PANDOC --track-changes=all "C:\Projects\Arborism\client\[Address] Report.docx" -t markdown | Select-String -Pattern "expected change" -Context 2
```

**Step 6 — Log** edits in `[project]\changelog.md`.

**Step 7 — Promote:**
```powershell
Remove-Item -Recurse -Force "[project]\unpacked\current" -ErrorAction SilentlyContinue
Rename-Item "[project]\unpacked\temp" "current"
```

## Reset Working Directory

```powershell
Remove-Item -Recurse -Force "[project]\unpacked" -ErrorAction SilentlyContinue
& $PY $UNPACK "path\to\new-document.docx" "[project]\unpacked\temp"
```

Always unpack from the original source in `work/`, not from the project folder.

## Reference Files

**Load only the reference required for the current task. Do not read all references upfront.**

| Reference | Load when... |
|---|---|
| `C:\Projects\Arborism\guideline.md` | Writing or editing narrative content (impact profiles, templates, tone rules) |
| `reference/editing-tracked-changes.md` | Making tracked-change edits to any document |
| `reference/section5-layout.md` | Editing the conclusion section (Section 5) |
| `reference/exploring-sample-reports.md` | Needing to check formatting against sample reports |

For OOXML details beyond what `reference/editing-tracked-changes.md` covers, consult the installed docx skill's `ooxml.md`.
