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
cd "$SKILL_OFFICE" && python3 unpack.py "$PROJECT_ROOT/new/[Address] Report.docx" "$PROJECT_ROOT/[project]/.work/unpacked/temp"
```
Generate an RSID for this editing session:
```bash
python3 -c "import secrets; print(secrets.token_hex(4).upper())"
```
Use it as `w:rsidR` on all new runs in the edit script.

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
cd "$SKILL_OFFICE" && python3 unpack.py "$PROJECT_ROOT/[project]/[Name].docx" "$PROJECT_ROOT/[project]/.work/unpacked/temp"
```
Generate an RSID for this editing session:
```bash
python3 -c "import secrets; print(secrets.token_hex(4).upper())"
```
Use it as `w:rsidR` on all new runs in the edit script.

**Step 1b — Extract tree data** (if not already done for this project):
```bash
python3 "$PROJECT_ROOT/.agents/skills/editing-arborist-reports/scripts/extract_trees.py" "path/to/document.docx"
```
Outputs `[project]/.work/tree_data.json` — read this instead of re-reading the full document for tree species, DBH, TPZ, condition, etc.

**Step 1b-alt — Diff against updated Excel** (if an updated `.xlsx` or pasted TSV is provided):
```bash
# Excel file:
python3 "$PROJECT_ROOT/.agents/skills/editing-arborist-reports/scripts/diff_trees.py" \
    "$PROJECT_ROOT/[project]/[Name].docx" \
    "$PROJECT_ROOT/new/[Address].xlsx"
# Or TSV from stdin (pasted tab-separated data):
echo -e "TREE #\tDirection\tTPZ (m)\n2\tInjury\t2.4" | \
    python3 "$PROJECT_ROOT/.agents/skills/editing-arborist-reports/scripts/diff_trees.py" \
        "$PROJECT_ROOT/[project]/[Name].docx" --stdin
```
Outputs `[project]/.work/diff.json` — per-tree field changes with old/new values. Only compares fields present in the source data (partial TSV input is fine). Use this diff to drive the edit script instead of manually reading the Excel.

**Step 1c — Map report structure** (replaces manual grepping for paraIds and table schemas):
```bash
python3 "$PROJECT_ROOT/.agents/skills/editing-arborist-reports/scripts/map_report.py" "[project]/.work"
```
Outputs to `[project]/.work/map.json`:
- Heading list with paraIds + line numbers
- Tables with **cell-level data**: each row has `row_para_id` and per-cell `{para_id, text, col_name}` — use these for `find_para()` calls instead of grepping `document.xml`
- Tree sections with paraIds + line numbers + previews
- Summary paragraphs
- **Permit bullets**: Section 5 permit count paragraphs with paraIds (e.g. "Removal Permit requirement: 6 trees")
- **Section paragraphs** (`section_paras`): all paragraphs under each Heading1, keyed by heading paraId. Includes paraId, text preview, and line number. Tables appear as `{type: "table"}`. Eliminates manual DOM walks for Section 5, Addendum 0, etc.

**Step 1d — Extract schema** (required before any insert-heavy edit):
```bash
python3 "$PROJECT_ROOT/.agents/skills/editing-arborist-reports/scripts/get_schema.py" "[project]/.work"
```
Reads `map.json` + `document.xml`. Outputs `[project]/.work/schema.json` with:
- Live rPr (inner XML) for every table type (sec4, impact, mini, injury_detail, removal, replanting)
- Table positioning (floating vs in-flow, tblpPr attributes)
- **Column widths** (`column_widths` array per table type) — use `schema['tables']['removal']['column_widths']` instead of hardcoding width arrays
- RPR constant validation (compares live values against edit_helpers defaults)
- Warnings for missing table types (expected: sec4, impact, summary; optional: injury_detail, mini, removal, replanting)

**Use schema.json rPr in edit scripts.** Pass `rpr=schema['rpr']['mini_data']` to builder functions. The hardcoded constants (`RPR_SEC4`, `RPR_INJURY`, etc.) are fallbacks — they were extracted from one sample report and may not match the current document.

**After get_schema.py succeeds: document.xml is sealed.** Never use sed/grep on document.xml for schema discovery after this point.

**Step 2 — Discover scope** before writing the script:
- Use `map.json` (from Step 1c) for paraIds — it has **cell-level paraIds** for every table cell. Use `table['rows'][N]['cells'][M]['para_id']` for `find_para()` calls.
- Use `map.json` `permit_bullets` for permit count paragraph paraIds.
- Use `schema.json` (from Step 1d) for live rPr values — pass to builders via `rpr=`.
- For each target narrative paragraph: read the paraId and preview from the map output.
- Cross-check: confirm every target cell/paragraph actually exists.

**Step 3 — Create edit script** in `[project]/.work/edit_script.py` using the template from `reference/editing-tracked-changes.md`. If editing narrative content, also load `$PROJECT_ROOT/guideline.md`.

**Required in every edit script:** call `s.validate_targets()` before any mutations:
```python
s.validate_targets([
    ("65486C2D", "w:p", "Sec4 tree 2 direction cell"),
    ("5EC182F9", "w:tr", "Impact table row for trees 2,3"),
    # ... all paraIds the script will target
])
```
This catches paraId type confusion (w:tr vs w:p) before any edits are made. If a target is missing, the script fails cleanly with no partial changes (auto-backup restores document.xml).

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
| ParaId exists? | `map.json` (cell-level paraIds in `rows[N].cells[M].para_id`) |
| Section structure? | `map.json` headings list (absence of a heading IS the answer) |
| Permit bullet paraIds? | `map.json` `permit_bullets` section |
| Live rPr for builders? | `schema.json` (from standard `get_schema.py`) |
| Table positioning? | `schema.json` `tables[type].positioning` |
| Find text at runtime? | `find_run(text, lo, hi)` in the edit script |
| Tree data? | `tree_data.json` |
| What changed in Excel? | `diff.json` (from `diff_trees.py`) |

**Before any grep/sed on document.xml, ask:** does map.json, tree_data.json, or get_schema.py output already answer this? If yes, stop.

### The schema seal

Run the standard `get_schema.py` (Step 1d). Review its output — check that all table types your edit needs are classified and have live rPr extracted. If a table type is missing (e.g., the report has no existing injury tables), the builder functions will use their RPR defaults — check `schema.json` `rpr_validation` for mismatches and override accordingly.

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
cd "$SKILL_OFFICE" && python3 pack.py "$PROJECT_ROOT/[project]/.work/unpacked/temp" "$PROJECT_ROOT/complete/[Address] Report.docx" --validate false
```
`--validate false` is required: the built-in `RedliningValidator` hardcodes author `"Claude"` but our tracked changes use `"Arborist"`, so validation flags every change as invalid.

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
rm -rf "$PROJECT_ROOT/[project]/.work/unpacked"
cd "$SKILL_OFFICE" && python3 unpack.py "$PROJECT_ROOT/[project]/[Name].docx" "$PROJECT_ROOT/[project]/.work/unpacked/temp"
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
