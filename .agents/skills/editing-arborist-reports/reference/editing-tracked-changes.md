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
import sys, os, xml.dom.minidom as minidom
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')

PROJECT = "/home/serg/projects/arborist-construction/[client]/.work"
DOC_PATH = PROJECT + "/unpacked/temp/word/document.xml"

with open(DOC_PATH, 'r', encoding='utf-8') as f:
    content = f.read()

dom = minidom.parseString(content.encode('utf-8'))
```

## Helper Functions

```python
# ─── Helpers ────────────────────────────────────────────────────────────

def get_text(node):
    """Concatenated text from all w:t and w:delText children."""
    parts = []
    for tag in ('w:t', 'w:delText'):
        for child in node.getElementsByTagName(tag):
            if child.firstChild:
                parts.append(child.firstChild.data)
    return ''.join(parts)

def find_run_in_line_range(content, dom, text, line_lo, line_hi):
    """Find w:r whose text matches AND whose occurrence falls within line range.
    Uses two-pass: line-scan to get occurrence index, then picks that DOM node."""
    lines = content.splitlines()
    occ_lines = [i+1 for i, ln in enumerate(lines) if f'<w:t>{text}</w:t>' in ln]
    target = [ln for ln in occ_lines if line_lo <= ln <= line_hi]
    if len(target) != 1:
        raise ValueError(f"Expected 1 match in {line_lo}-{line_hi}, got {target}")
    idx = occ_lines.index(target[0])
    runs = [r for r in dom.getElementsByTagName('w:r') if get_text(r) == text]
    return runs[idx]

def find_para_by_para_id(dom, para_id):
    """Find w:p with matching w14:paraId attribute."""
    for p in dom.getElementsByTagName('w:p'):
        if p.getAttribute('w14:paraId') == para_id:
            return p
    raise ValueError(f"Paragraph with paraId={para_id} not found")

def insert_xml_after(dom, ref_node, xml_string):
    """Insert parsed XML nodes immediately after ref_node."""
    parent = ref_node.parentNode
    wrapper = minidom.parseString(
        f'<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
        f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'{xml_string}</root>'
    )
    ref_sibling = ref_node.nextSibling
    for node in wrapper.documentElement.childNodes:
        parent.insertBefore(dom.importNode(node, True), ref_sibling)

def replace_node_with_xml(dom, old_node, xml_string):
    """Replace old_node in its parent with the parsed XML nodes."""
    parent = old_node.parentNode
    wrapper = minidom.parseString(
        f'<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'{xml_string}</root>'
    )
    ref = old_node.nextSibling
    for node in wrapper.documentElement.childNodes:
        parent.insertBefore(dom.importNode(node, True), ref)
    parent.removeChild(old_node)
```

## Critical: `<w:ins>` Wrapping is Manual

The helpers do **not** auto-create `<w:ins>`/`<w:del>` elements. All new content must be wrapped explicitly:
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
