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
Note the suggested RSID.

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
Note the suggested RSID.

**Step 2 — Create edit script** in `[project]/.work/edit_script.py` using the template from `reference/editing-tracked-changes.md`.

**Step 3 — Run:**
```bash
python3 "[project]/.work/edit_script.py"
```

**Step 4–7:** Follow Verification section below.

## Key Editing Principles

- **Always track changes**: Use `<w:del>`/`<w:ins>` tags. Never make untracked modifications. All new content must be wrapped in `<w:ins>` — the helpers do not create these automatically.
- **Scope lock**: Only edit the section specified by the user.
- **Minimal edits**: Only wrap changed text in `<w:del>`/`<w:ins>` tags.
- **Preserve formatting**: Extract and reuse `<w:rPr>` from original nodes.
- **Batch changes**: Group 3-10 related edits per script.
- **Grep first**: Always check `word/document.xml` line numbers before editing.
- **Section 5 edits**: Load `reference/section5-layout.md` before editing the conclusion section.

## Verification and Logging

**Step 4 — Pack:**
```bash
cd "$SKILL_OFFICE" && python3 pack.py "[project]/.work/unpacked/temp" "$PROJECT_ROOT/complete/[Address] Report.docx" --validate false
```

**Step 5 — Verify:**
```bash
pandoc --track-changes=all "$PROJECT_ROOT/complete/[Address] Report.docx" -t plain \
    | grep -i "expected change"
```

**Step 6 — Log** edits in `[project]/.work/changelog.md`.

**Step 7 — Promote:**
```bash
rm -rf "[project]/.work/unpacked/current"
mv "[project]/.work/unpacked/temp" "[project]/.work/unpacked/current"
```

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
