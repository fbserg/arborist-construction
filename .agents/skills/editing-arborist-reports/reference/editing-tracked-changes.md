# Editing with Tracked Changes

Detailed reference for making tracked-change edits to `.docx` files. Load this before creating edit scripts.

## Contents

- Edit script template
- Helper functions
- Critical: manual `<w:ins>` wrapping
- XML helper patterns
- Replacement pattern
- Getting nodes
- Saving

## Edit Script Template

```python
import sys
sys.path.insert(0, "/home/serg/projects/arborist-construction/.agents/skills/editing-arborist-reports/scripts")
from edit_helpers import EditSession, insert_xml_after, RPR_NORMAL, RPR_BOLD
# Builder functions for table rows/mini-tables (import only what the script needs):
from edit_helpers import tc, impact_row, injury_row, sec4_row, mini_table, injury_detail_table

# EditSession handles encoding, document loading, and auto-incrementing change IDs.
# start_id auto-detects from existing tracked changes (max existing ID + 1).
# Override with start_id=N if needed (check changelog for last used ID).
# rsid: record the RSID from unpack output — used as w:rsidR on new runs/rows.
s = EditSession("work/[Client]/.work", "2026-02-25", "Arborist", rsid="00AB12CD")

# Find and replace text (auto del+ins, auto rPr extraction):
node = s.find_run("old text", 4455, 4465)
s.replace_text(node, "old text", "new text")

# Surgical phrase replacement (marks only changed words — use for ≤3 word changes):
node = s.find_run("The tree in question is significant", 1200, 1220)
s.replace_phrase_in_run(node, "tree in question", "subject tree")

# Replace with explicit rPr:
s.replace_text(node, "old text", "new text", RPR_BOLD)

# Insert paragraphs after a reference node:
para = s.find_para("60180DE4")
insert_xml_after(s.dom, para,
    s.ins_para("Paragraph 1 text.") +
    s.ins_para("Paragraph 2 text.")
)

# Generate del/ins XML directly (for complex cases):
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
insert_xml_after(s.dom, anchor_tr, sec4_row(s, ["15", "Silver Maple", "Acer saccharinum", "22", "Good", "", "Private", "Injury", "4.4", "Yes"], rpr=live_rpr))

# Mini-table (floating 8-col tree summary — tblpX/tblpY from get_schema.py):
insert_xml_after(s.dom, para, mini_table(s, 15, "Silver Maple", 22, "Good", "Good health", "Private", "Injury", 4.4, tblpX="1513", tblpY="2322"))

# Injury detail table (header + rows — tblpX/tblpY for floating, rpr overrides):
insert_xml_after(s.dom, para, injury_detail_table(s, [("Front driveway", "2.1m", '4"', "Moderate")], tblpX="1513", tblpY="5500"))

s.save()  # Prints IDs used range
```

**Standalone helpers** are also importable for scripts that need DOM access without EditSession:
`from edit_helpers import load_document, get_text, find_run_in_line_range, find_para_by_para_id, insert_xml_after, replace_node_with_xml`

**Full module:** `.agents/skills/editing-arborist-reports/scripts/edit_helpers.py`

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
# Tracked insertion paragraph (body text)
def ins_para(text, rpr):
    return f'''<w:p>
  <w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr>
  <w:ins w:id="30" w:author="Arborist" w:date="2026-02-24T00:00:00Z"><w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:ins>
</w:p>'''

# Standard 4-side border (for table cells)
BORDERS = '''<w:tcBorders>
  <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tcBorders>'''
```

## Replacement Pattern

Use `<w:del>` + `<w:ins>`:
```python
node = find_run_in_line_range(content, dom, "Injury", 8350, 8420)
rpr_nodes = node.getElementsByTagName('w:rPr')
rpr = rpr_nodes[0].toxml() if rpr_nodes else ''

replace_node_with_xml(dom, node,
    f'<w:del w:id="29" w:author="Arborist" w:date="2026-02-24T00:00:00Z">'
    f'<w:r>{rpr}<w:delText>Injury</w:delText></w:r></w:del>'
    f'<w:ins w:id="30" w:author="Arborist" w:date="2026-02-24T00:00:00Z">'
    f'<w:r>{rpr}<w:t>Removal</w:t></w:r></w:ins>'
)
```

**Minimal edits principle** — only mark text that actually changes:
```python
# BAD - Replaces entire sentence
'<w:del w:id="1" w:author="Arborist" w:date="2026-02-24T00:00:00Z"><w:r><w:delText>The term is 30 days.</w:delText></w:r></w:del><w:ins w:id="2" w:author="Arborist" w:date="2026-02-24T00:00:00Z"><w:r><w:t>The term is 60 days.</w:t></w:r></w:ins>'

# GOOD - Only marks what changed
'<w:r w:rsidR="00AB12CD"><w:t>The term is </w:t></w:r><w:del w:id="1" w:author="Arborist" w:date="2026-02-24T00:00:00Z"><w:r><w:delText>30</w:delText></w:r></w:del><w:ins w:id="2" w:author="Arborist" w:date="2026-02-24T00:00:00Z"><w:r><w:t>60</w:t></w:r></w:ins><w:r w:rsidR="00AB12CD"><w:t> days.</w:t></w:r>'
```

## Getting Nodes

```python
# By text + line range (disambiguate repeated text):
node = find_run_in_line_range(content, dom, "Injury", 8350, 8420)

# By paraId (most stable — use Grep/Read to find paraId in document.xml):
para = find_para_by_para_id(dom, '0213C74E')

# By unique text (if text only appears once):
runs = [r for r in dom.getElementsByTagName('w:r') if get_text(r) == 'unique text']
```

## Saving

```python
# Save
with open(DOC_PATH, 'wb') as f:
    f.write(dom.toxml(encoding='utf-8'))
print("Saved.")
```

`dom.toxml(encoding='utf-8')` returns `bytes` — open the output file in `'wb'` mode.
