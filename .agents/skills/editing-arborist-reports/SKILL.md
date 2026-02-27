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

Key variables for all editing commands:

```bash
source "$PROJECT_ROOT/.agents/skills/editing-arborist-reports/scripts/setup.sh"
# Or set variables directly — see setup.sh for canonical values
```

Canonical values:
- `PROJECT_ROOT` = `/home/serg/projects/arborist-construction`
- `SKILL_OFFICE` = `~/.claude/plugins/cache/anthropic-agent-skills/document-skills/1ed29a03dc85/skills/docx/scripts/office`

## Project Folder Structure

```
arborist-construction/
├── new/            # Incoming — files needing work
├── work/           # Active processing
│   ├── [Client]/       # e.g., 71 Lloyd Manor
│   │   ├── .work/          # Edit artifacts (scripts, unpacked, changelog)
│   │   └── [Name].docx     # Source report
│   └── [Sample] Report.docx  # Sample reports for reference
├── complete/       # Delivered, closed files
```

Revisions loop back: `complete/ → new/ → work/ → complete/`

`[project]` in paths below = the client folder under work (e.g., `work/71 Lloyd Manor`).

## Decision Tree

Ask what the user needs:

1. **New report from template** → Section: New Report
2. **Specific edits** → Skip full read. Grep to locate text, then edit directly
3. **Injury/removal edits** → Read only Section 4 (tree data table) and the target section
4. **Review or audit** → Read the full document
5. **Unsure** → Ask clarifying questions

## New Report from Template

**Step 1:** Create project folder and unpack:
```bash
mkdir -p "$PROJECT_ROOT/work/[Client]/.work"
cd "$SKILL_OFFICE" && python3 unpack.py "$PROJECT_ROOT/new/[Address] Report.docx" "[project]/.work/unpacked/temp"
```
Generate an 8-hex-digit RSID (e.g. `00F2A1B3`) — use it as `w:rsidR` on all new runs in the edit script. (unpack.py does not output a suggested RSID.)

**Step 2:** Read content rules. Load `$PROJECT_ROOT/guideline.md` for impact profiles, species tolerance, narrative templates, TPZ calculations, and post-removal notes.

**Step 3:** Fill sections using tracked changes per the Editing workflow below.

**Step 4:** Pack, verify, and log per the Verification section below.

## Reading a Report

```bash
pandoc "path/to/document.docx" -t markdown     # or -o "[project]/[Name].md" to save
# Add --track-changes=all to include tracked changes in output
```

## Editing a Report (Tracked Changes)

**Before editing:** Load `reference/editing-tracked-changes.md` for the script template, helper functions, and XML patterns.

**Step 1 — Unpack:**
```bash
cd "$SKILL_OFFICE" && python3 unpack.py "path/to/document.docx" "[project]/.work/unpacked/temp"
```
Generate an 8-hex-digit RSID (e.g. `00F2A1B3`) — use it as `w:rsidR` on all new runs in the edit script. (unpack.py does not output a suggested RSID.)

**Step 1b — Extract tree data** (if not already done for this project):
```bash
python3 "$PROJECT_ROOT/.agents/skills/editing-arborist-reports/scripts/extract_trees.py" "path/to/document.docx"
```
Outputs `[project]/.work/tree_data.json` — read this instead of re-reading the full document for tree species, DBH, TPZ, condition, etc.

**Step 1c — Map report structure** (replaces manual grepping for paraIds and table schemas):
```bash
python3 "$PROJECT_ROOT/.agents/skills/editing-arborist-reports/scripts/map_report.py" "[project]/.work"
```
Outputs heading list, tree sections with paraIds + line numbers + previews, table column schemas, and summary paragraphs. JSON saved to `[project]/.work/map.json` — reference it when writing the edit script instead of grepping `document.xml` by hand.

**Step 1d — Schema inventory** (required before any insert-heavy edit):

Walk every insert operation in the plan and build a table:

| Insert operation | XML structure needed | In get_schema.py? |
|---|---|---|
| New table row (any table) | `w:tr` from that specific table | ✓/✗ |
| New floating mini-table (tree entry) | Full `w:tbl` + data `w:tr` | ✓/✗ |
| New narrative paragraph | `ins_para()` helper — no extract needed | — |

Write `get_schema.py` to extract one live example of every ✗ row. Run it. Verify output covers the full inventory. Fix and re-run until all ✗ become ✓.

**get_schema.py must also validate RPR constants.** For every table type that uses a builder function (`sec4_row`, `impact_row`, `injury_row`, `injury_detail_table`), extract the `w:rPr` from an existing data row and print it alongside the hardcoded `RPR_*` constant. If they differ (font size, language, bold, color), pass the extracted rPr to the builder via `rpr=` override in the edit script. The hardcoded constants (`RPR_SEC4`, `RPR_INJURY`, etc.) are fallbacks — they were extracted from one sample report and may not match the current document.

**After get_schema.py succeeds: document.xml is sealed.** Any schema gap found while writing edit_script.py → fix get_schema.py and re-run. Never use sed/grep on document.xml for schema discovery after this point.

**Step 2 — Discover scope** before writing the script:
- Use `map.json` (from Step 1c) for paraIds, table schemas, and section structure — do not assume schema from templates.
- For each target narrative paragraph: read the paraId and preview from the map output. Confirm the paragraph exists and note its first ~10 words before finalizing what to change.
- Cross-check: confirm every target cell/paragraph actually exists.

**Step 3 — Create edit script** in `[project]/.work/edit_script.py` using the template from `reference/editing-tracked-changes.md`. If editing narrative content, also load `$PROJECT_ROOT/guideline.md`.

**Step 4 — Run:**
```bash
python3 "[project]/.work/edit_script.py"
```

**Step 5–8:** Follow Verification section below.

## Key Editing Principles

- **Always track changes**: Use `<w:del>`/`<w:ins>` tags. Never make untracked modifications. All new content must be wrapped in `<w:ins>` — the helpers do not create these automatically.
- **Scope lock**: Only edit the section specified by the user.
- **Minimal edits**: Only wrap changed text in `<w:del>`/`<w:ins>` tags.
- **Preserve formatting**: Extract and reuse `<w:rPr>` from original nodes.
- **Batch changes**: Group 3-10 related edits per script.
- **Surgical phrase edits**: For small phrase changes (≤3 words) within long paragraphs, use `s.replace_phrase_in_run(run, phrase, replacement)` instead of `replace_text()` — it marks only the changed words as del/ins.
- **Section 5 edits**: Load `reference/section5-layout.md` before editing the conclusion section.
- **Boilerplate paragraphs — do not edit**: These Section 5 paragraphs are standard text that applies regardless of tree count or type. Never modify them when adding/editing trees:
  - "All bylaw protected trees not slated for removal or injury are to be protected by fencing..."
  - "No other municipally owned trees of any size, private trees, or neighbouring trees..."
  - All RSE (Root Sensitive Excavation) boilerplate paragraphs
  - General Notes and signature block
- **Consistency check**: After completing scoped edits, scan the Summary section and other sections for contradictions with the new content before packing (e.g., source-of-impact wording must match across Summary, impact table, and narrative).
- **paraId anchoring**: For narrative paragraphs, prefer `s.find_para("XXXXXXXX")` over line-range text matching — paraIds are stable across content changes.

### document.xml is sealed — no shell access, ever

`document.xml` is ~16,000 lines. **No grep. No sed. No cat. No Read tool. Ever.**

The only thing that reads document.xml is Python code (`load_document()` in edit_helpers.py, get_schema.py). Shell commands on document.xml are always wrong. Every need is already covered:

| Need | Tool |
|---|---|
| ParaId exists? | `map.json` |
| Section structure? | `map.json` headings list (absence of a heading IS the answer) |
| Row XML schema? | `get_schema.py` |
| Find text at runtime? | `find_run(text, lo, hi)` in the edit script |
| Tree data? | `tree_data.json` |

**Before any grep/sed on document.xml, ask:** does map.json, tree_data.json, or get_schema.py output already answer this? If yes, stop.

### The get_schema.py seal

Before writing `get_schema.py`, produce a **schema inventory**: walk every insert operation in the plan and list each XML row/table type that edit_script.py will INSERT. Write get_schema.py to extract one live example of each. Run it. Compare output against the inventory. If anything is missing → fix get_schema.py and re-run.

**After get_schema.py succeeds: document.xml is sealed.** Schema gap found while writing edit_script.py → fix get_schema.py and re-run. No shell access to document.xml.

Example inventory table:

| Insert operation | XML structure needed | In get_schema.py? |
|---|---|---|
| New impact table row | `w:tr` (3-col pct table) | ✓/✗ |
| New injury detail row | `w:tr` (4-col pct table) | ✓/✗ |
| New Section 4 data row | `w:tr` (10-col dxa table) | ✓/✗ |
| New mini-table (tree entry) | Full `w:tbl` + data `w:tr` | ✓/✗ |
| New narrative paragraph | `ins_para()` — no extract needed | — |

## Verification and Logging

**Step 5 — Pack:**
```bash
cd "$SKILL_OFFICE" && python3 pack.py "[project]/.work/unpacked/temp" "$PROJECT_ROOT/complete/[Address] Report.docx" --validate false
```

**Step 6 — Verify:**
```bash
# Primary — confirm final accepted state (acceptance test):
pandoc --track-changes=accept "$PROJECT_ROOT/complete/[Address] Report.docx" -t plain \
    | grep -i "new text"

# Secondary — inspect markup (shows both deleted and inserted inline):
pandoc --track-changes=all "$PROJECT_ROOT/complete/[Address] Report.docx" -t plain \
    | grep -i "target text"
```
The `--track-changes=all` output includes deleted text inline — don't mistake it for surviving text. Always confirm the acceptance-test output first.

**Pandoc limitations:** Pandoc verification confirms text content only — it cannot detect font size mismatches, table positioning errors (floating vs. in-flow), or styling inconsistencies between original and inserted elements. For formatting-sensitive edits (new table rows, mini-tables), also verify via the get_schema.py RPR validation output or manual review in Word/LibreOffice.

**Step 7 — Log** edits in `[project]/.work/changelog.md`. Required fields:
- **Date** and **description** of changes
- **Change IDs** used (range, e.g. 1–16)
- **RSID** used for the session
- **Author** and **date stamp** on tracked changes

**Step 8 — Promote:**
```bash
rm -rf "[project]/.work/unpacked/current"
mv "[project]/.work/unpacked/temp" "[project]/.work/unpacked/current"
```

**Step 9 — Review** (optional but recommended):
- Cross-section consistency: do Summary, impact table, and narrative all agree?
- Were all SKILL.md steps followed (RSID captured, references loaded, guideline loaded for narrative)?
- Any workarounds used that indicate tooling gaps? Note in memory.
- Changelog updated with all required fields?

## Reset Working Directory

```bash
rm -rf "[project]/.work/unpacked"
cd "$SKILL_OFFICE" && python3 unpack.py "path/to/new-document.docx" "[project]/.work/unpacked/temp"
```

Always unpack from the source docx (in `work/[Client]/` or `new/`), not from the unpacked folder.

## Reference Files

**Load only the reference required for the current task. Do not read all references upfront.**

| Reference | Load when... |
|---|---|
| `$PROJECT_ROOT/guideline.md` | Writing or editing narrative content (impact profiles, templates, tone rules) |
| `reference/editing-tracked-changes.md` | Making tracked-change edits to any document |
| `reference/section5-layout.md` | Editing the conclusion section (Section 5) |
| `reference/exploring-sample-reports.md` | Needing to check formatting against sample reports |
