# Editing Workflow Reference

Detailed reference for document editing. Read this when performing tracked-change edits on `.docx` files. For overview and folder structure, see `CLAUDE.md`.

`[project]` refers to the temp signifier folder (e.g., `temp\71 Lloyd`).

## Edit Script Template

Standard script header (always use):
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

## UTF-8 Note

UTF-8 encoding is handled in the script header. Do **not** rely on `$env:PYTHONIOENCODING` — it does not work through the bash→powershell bridge.
