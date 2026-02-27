# Editing with Tracked Changes

Detailed reference for making tracked-change edits to `.docx` files. Load this before creating edit scripts.

## Contents

- Edit script template
- Helper functions
- Critical: manual `<w:ins>` wrapping
- XML helper patterns
- Getting nodes
- Saving

## Edit Script Template

```python
import sys, json
sys.path.insert(0, "/home/serg/projects/arborist-construction/.agents/skills/editing-arborist-reports/scripts")
from edit_helpers import (EditSession, insert_xml_after, replace_node_with_xml, get_text,
                          prev_element_sibling, next_element_sibling, extract_rpr,
                          RPR_NORMAL, RPR_BOLD)
# Builder functions for table rows/mini-tables (import only what the script needs):
from edit_helpers import tc, impact_row, injury_row, sec4_row, mini_table, injury_detail_table

# EditSession handles encoding, document loading, and auto-incrementing change IDs.
# start_id auto-detects from existing tracked changes (max existing ID + 1).
# Override with start_id=N if needed (check changelog for last used ID).
# rsid: use python3 -c "import secrets; print(secrets.token_hex(4).upper())" to generate.
# Auto-backup: document.xml is backed up on init; auto-rollback on exception via context manager.
with EditSession("work/[Client]/.work", "2026-02-25", "Arborist", rsid="00AB12CD") as s:

    # Load live rPr from schema.json (always use instead of hardcoded RPR_ constants):
    with open("work/[Client]/.work/schema.json") as f:
        schema = json.load(f)
    LIVE_MINI_RPR = schema['rpr'].get('mini_data')
    LIVE_MINI_HDR = schema['rpr'].get('mini_hdr')

    # Pre-flight validation — catches paraId type confusion before any mutations:
    s.validate_targets([
        ("65486C2D", "w:p", "target cell description"),
        ("5EC182F9", "w:tr", "target row description"),
    ])

    # ── Find-and-replace in a paragraph cell (one-liner) ──
    s.replace_in_para(s.find_para("65486C2D"), "Remove", "Injury")

    # ── Surgical phrase replacement (marks only changed words — use for ≤3 word changes) ──
    node = s.find_run("The tree in question is significant", 1200, 1220)
    s.replace_phrase_in_run(node, "tree in question", "subject tree")

    # ── Track-delete all runs in a paragraph ──
    s.delete_para(s.find_para("2F097DE3"))

    # ── Track-delete an entire table row ──
    s.delete_row(s.find_tr("67C88484"))

    # ── Sibling navigation (skips whitespace text nodes) ──
    prev = prev_element_sibling(s.find_para("403C8662"))
    next_ = next_element_sibling(s.find_para("403C8662"))

    # ── Insert paragraphs after a reference node ──
    para = s.find_para("60180DE4")
    insert_xml_after(s.dom, para,
        s.ins_para("Paragraph 1 text.") +
        s.ins_para("Paragraph 2 text.")
    )

    # ── Generate del/ins XML directly (for complex cases) ──
    xml = s.del_run("old") + s.ins_run("new", RPR_BOLD)

    # ── Builder functions (for insert-heavy edits) ──
    # All builders accept optional rpr= to override hardcoded RPR_ constants.
    # IMPORTANT: Extract actual rPr from the live document via get_schema.py —
    # the hardcoded constants (RPR_SEC4, RPR_INJURY, etc.) may not match every report.

    # Impact table row (3-col):
    anchor_tr = s.find_tr("16321D29")
    insert_xml_after(s.dom, anchor_tr, impact_row(s, 5, "Removal — condition-based", "Removal"))

    # Injury detail row (4-col):
    insert_xml_after(s.dom, anchor_tr, injury_row(s, "Front walkway", "3.1m", '4"', "Moderate"))

    # Section 4 data row (10-col) — pass rpr= extracted from existing data row:
    insert_xml_after(s.dom, anchor_tr, sec4_row(s, ["15", "Silver Maple", "Acer saccharinum", "22", "Good", "", "Private", "Injury", "4.4", "Yes"], rpr=LIVE_MINI_RPR))

    # Mini-table (floating 8-col tree summary — tblpX/tblpY from schema.json, rPr overrides):
    insert_xml_after(s.dom, para, mini_table(s, 15, "Silver Maple", 22, "Good", "Good health", "Private", "Injury", 4.4, tblpX="1513", tblpY="2322", hdr_rpr=LIVE_MINI_HDR, data_rpr=LIVE_MINI_RPR))

    # Injury detail table (header + rows — tblpX/tblpY for floating, rpr overrides):
    insert_xml_after(s.dom, para, injury_detail_table(s, [("Front driveway", "2.1m", '4"', "Moderate")], tblpX="1513", tblpY="5500"))

    # Column widths are now in schema.json — use schema['tables']['removal']['column_widths']
    # instead of hardcoding width arrays.

    s.save()  # Prints IDs used range; auto-removes backup on success
# If exception occurs before save(), context manager auto-rolls back document.xml
```

**Standalone helpers** are also importable for scripts that need DOM access without EditSession:
`from edit_helpers import load_document, get_text, find_run_in_line_range, find_para_by_para_id, insert_xml_after, replace_node_with_xml, prev_element_sibling, next_element_sibling, extract_rpr`

**Full module:** `.agents/skills/editing-arborist-reports/scripts/edit_helpers.py`

## Helper Quick Reference

| Need | Method |
|---|---|
| Replace run text (full match) | `s.replace_text(node, old, new)` |
| Replace first matching run in paragraph | `s.replace_in_para(para, old, new)` |
| Surgical phrase change (≤3 words) | `s.replace_phrase_in_run(run, phrase, new)` |
| Track-delete paragraph content | `s.delete_para(para)` |
| Track-delete table row | `s.delete_row(tr)` |
| Navigate siblings | `prev_element_sibling(node)` / `next_element_sibling(node)` |
| Date without time | `s.date_short` (returns "2026-02-25") |
| Generate insert paragraph | `s.ins_para(text)` |
| Generate del/ins XML | `s.del_run(text)` / `s.ins_run(text)` |

## Critical: `<w:ins>` Wrapping

**EditSession methods** (`del_run`, `ins_run`, `replace_text`, `ins_para`) auto-create `<w:ins>`/`<w:del>` wrappers with correct IDs.

**Manual XML** (for complex cases like table cells) still requires explicit wrapping:
```xml
<!-- WRONG — will not show as tracked change -->
<w:r><w:rPr>...</w:rPr><w:t>new text</w:t></w:r>

<!-- CORRECT — will show as tracked insertion -->
<w:ins w:id="30" w:author="Arborist" w:date="2026-02-24T00:00:00Z">
  <w:r><w:rPr>...</w:rPr><w:t>new text</w:t></w:r>
</w:ins>
```

## XML Helper Patterns

```python
# Tracked insertion paragraph (body text) — now just use s.ins_para():
xml = s.ins_para("New paragraph text.")

# Standard 4-side border (for table cells) — available as BORDERS constant:
from edit_helpers import BORDERS
```

## Getting Nodes

```python
# By paraId (most stable — get paraIds from map.json):
para = s.find_para('0213C74E')

# By text + line range (disambiguate repeated text):
node = s.find_run("Injury", 8350, 8420)

# Table row by paraId:
tr = s.find_tr('5EC182F9')
```

## Saving

```python
s.save()  # Writes document.xml, removes backup, prints ID range
```

Or use context manager (recommended):
```python
with EditSession(...) as s:
    # edits
    s.save()
# auto-rollback if exception before save
```
