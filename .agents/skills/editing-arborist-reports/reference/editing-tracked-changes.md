# Editing with Tracked Changes

Detailed reference for making tracked-change edits to `.docx` files. Load this before creating edit scripts.

## Contents

- Edit script template
- Editor API
- Critical: manual `<w:ins>` wrapping
- XML helper patterns
- Replacement pattern
- Getting nodes
- Tracked change XML patterns
- Saving and validation
- UTF-8 note

## Edit Script Template

```python
import sys, os
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')
SKILL = r"C:\Users\User\.claude\plugins\cache\anthropic-agent-skills\document-skills\69c0b1a06741\skills\docx"
sys.path.insert(0, SKILL)
from scripts.document import Document

PROJECT = r"[project]"
doc = Document(PROJECT + r"\unpacked\temp", author="Arborist", rsid="SUGGESTED_RSID")
editor = doc["word/document.xml"]
```

## Editor API

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

## Critical: `<w:ins>` Wrapping is Manual

The API populates attributes on `<w:ins>`/`<w:del>` elements but does **not** auto-create them. All new content must be wrapped in `<w:ins>`:
```xml
<!-- WRONG — will not show as tracked change -->
<w:r><w:rPr>...</w:rPr><w:t>new text</w:t></w:r>

<!-- CORRECT — will show as tracked insertion -->
<w:ins><w:r><w:rPr>...</w:rPr><w:t>new text</w:t></w:r></w:ins>
```

## XML Helper Patterns

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

## Replacement Pattern

Use `<w:del>` + `<w:ins>`:
```python
node = editor.get_node(tag="w:r", contains="text to find")
rpr = tags[0].toxml() if (tags := node.getElementsByTagName("w:rPr")) else ""
replacement = f'<w:del><w:r>{rpr}<w:delText>old</w:delText></w:r></w:del><w:ins><w:r>{rpr}<w:t>new</w:t></w:r></w:ins>'
editor.replace_node(node, replacement)

doc.save()
```

**Minimal edits principle** — only mark text that actually changes:
```python
# BAD - Replaces entire sentence
'<w:del><w:r><w:delText>The term is 30 days.</w:delText></w:r></w:del><w:ins><w:r><w:t>The term is 60 days.</w:t></w:r></w:ins>'

# GOOD - Only marks what changed
'<w:r w:rsidR="00AB12CD"><w:t>The term is </w:t></w:r><w:del><w:r><w:delText>30</w:delText></w:r></w:del><w:ins><w:r><w:t>60</w:t></w:r></w:ins><w:r w:rsidR="00AB12CD"><w:t> days.</w:t></w:r>'
```

## Getting Nodes

```python
# By text content
node = editor.get_node(tag="w:p", contains="specific text")

# By line range
para = editor.get_node(tag="w:p", line_number=range(100, 150))

# By exact line number
para = editor.get_node(tag="w:p", line_number=42)

# By attributes
node = editor.get_node(tag="w:del", attrs={"w:id": "1"})

# Combine filters (disambiguate repeated text)
node = editor.get_node(tag="w:r", contains="Section", line_number=range(2400, 2500))
```

## Tracked Change XML Patterns

**Text Insertion:**
```xml
<w:ins>
  <w:r><w:t>inserted text</w:t></w:r>
</w:ins>
```

**Text Deletion:**
```xml
<w:del>
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

**Deleting another author's insertion (nested structure required):**
```xml
<w:ins w:author="Jane Smith" w:id="16">
  <w:del w:author="Arborist" w:id="40">
    <w:r><w:delText>monthly</w:delText></w:r>
  </w:del>
</w:ins>
<w:ins w:author="Arborist" w:id="41">
  <w:r><w:t>weekly</w:t></w:r>
</w:ins>
```

**Delete entire run:**
```python
node = editor.get_node(tag="w:r", contains="text to delete")
editor.suggest_deletion(node)
```

**Delete entire paragraph:**
```python
para = editor.get_node(tag="w:p", contains="paragraph to delete")
editor.suggest_deletion(para)
```

**New tracked paragraph insertion:**
```python
from scripts.document import DocxXMLEditor
target_para = editor.get_node(tag="w:p", contains="existing text")
pPr = tags[0].toxml() if (tags := target_para.getElementsByTagName("w:pPr")) else ""
new_item = f'<w:p>{pPr}<w:r><w:t>New item</w:t></w:r></w:p>'
tracked_para = DocxXMLEditor.suggest_paragraph(new_item)
editor.insert_after(target_para, tracked_para)
```

## Saving and Validation

```python
doc.save()  # Validates by default, raises error if validation fails
doc.save(validate=False)  # Skip validation (debugging only)
```

The validator checks that document text matches the original after reverting tracked changes. Rules:
- Never modify text inside another author's `<w:ins>` or `<w:del>` tags
- Always use nested deletions to remove another author's insertions
- Every edit must be properly tracked

## UTF-8 Note

UTF-8 encoding is handled in the script header. Do not rely on `$env:PYTHONIOENCODING`.
