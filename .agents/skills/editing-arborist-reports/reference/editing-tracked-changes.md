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
from edit_helpers import EditSession, find_run_in_line_range, find_para_by_para_id, insert_xml_after, extract_rpr, RPR_NORMAL, RPR_BOLD

# EditSession handles encoding, document loading, and auto-incrementing change IDs.
# Check [project]/.work/changelog.md for last used ID and pass start_id accordingly.
s = EditSession("work/[Client]/.work", "2026-02-25", "Arborist", start_id=1)

# Find and replace text (auto del+ins, auto rPr extraction):
node = s.find_run("old text", 4455, 4465)
s.replace_text(node, "old text", "new text")

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
