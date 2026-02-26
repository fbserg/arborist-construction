"""Shared helpers for arborist report tracked-change editing.

Usage:
    sys.path.insert(0, "/home/serg/projects/arborist-construction/.agents/skills/editing-arborist-reports/scripts")
    from edit_helpers import EditSession, find_run_in_line_range, find_para_by_para_id, insert_xml_after
"""

import sys, os, xml.dom.minidom as minidom

# ─── Encoding setup ──────────────────────────────────────────────────────────

os.environ["PYTHONIOENCODING"] = "utf-8"
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# ─── DOM helpers ─────────────────────────────────────────────────────────────

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

    Two-pass: line-scan to get occurrence index, then picks that DOM node.
    Falls back to DOM text matching when raw scan misses (XML entity encoding)."""
    lines = content.splitlines()
    occ_lines = [i+1 for i, ln in enumerate(lines)
                 if f'<w:t>{text}</w:t>' in ln
                 or f'<w:t xml:space="preserve">{text}</w:t>' in ln]
    target = [ln for ln in occ_lines if line_lo <= ln <= line_hi]

    if len(target) == 1:
        idx = occ_lines.index(target[0])
        runs = [r for r in dom.getElementsByTagName('w:r') if get_text(r) == text]
        return runs[idx]

    # Fallback: raw scan missed (XML entities, etc.) — use DOM text matching only
    runs = [r for r in dom.getElementsByTagName('w:r') if get_text(r) == text]
    if len(runs) == 1:
        return runs[0]
    if not runs:
        raise ValueError(f"No run found with text '{text}' (raw scan: {occ_lines})")
    raise ValueError(
        f"Ambiguous: {len(runs)} DOM matches for '{text}' but raw line scan "
        f"found {len(target)} in range {line_lo}-{line_hi} (entity encoding?). "
        f"Use find_para_by_para_id() or narrow the line range."
    )


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
    """Replace old_node in its parent with parsed XML nodes."""
    parent = old_node.parentNode
    wrapper = minidom.parseString(
        f'<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'{xml_string}</root>'
    )
    ref = old_node.nextSibling
    for node in wrapper.documentElement.childNodes:
        parent.insertBefore(dom.importNode(node, True), ref)
    parent.removeChild(old_node)


def extract_rpr(node):
    """Pull <w:rPr> XML string from an existing run node. Returns '' if none."""
    rpr_nodes = node.getElementsByTagName('w:rPr')
    return rpr_nodes[0].toxml() if rpr_nodes else ''


def load_document(work_path):
    """Load document.xml from an unpacked .work directory.

    Args:
        work_path: Path to [project]/.work (e.g. "work/27 Shudell Avenue/.work")
                   Can be absolute or relative to PROJECT_ROOT.

    Returns:
        (dom, content) — parsed minidom Document and raw XML string.
    """
    if not os.path.isabs(work_path):
        work_path = os.path.join("/home/serg/projects/arborist-construction", work_path)
    doc_path = os.path.join(work_path, "unpacked/temp/word/document.xml")
    with open(doc_path, 'r', encoding='utf-8') as f:
        content = f.read()
    dom = minidom.parseString(content.encode('utf-8'))
    return dom, content


# ─── RPR constants ───────────────────────────────────────────────────────────

RPR_NORMAL = '<w:rPr><w:rFonts w:cs="Helvetica"/><w:szCs w:val="22"/></w:rPr>'
RPR_BOLD = '<w:rPr><w:rFonts w:cs="Helvetica"/><w:b/><w:bCs/><w:szCs w:val="22"/></w:rPr>'
PPR = '<w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr>'

# Standard 4-side border for table cells
BORDERS = '''<w:tcBorders>
  <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tcBorders>'''


# ─── EditSession ─────────────────────────────────────────────────────────────

class EditSession:
    """Manages tracked-change IDs, XML generation, and document save.

    Usage:
        s = EditSession("work/27 Shudell Avenue/.work", "2026-02-25", "Arborist")
        node = s.find_run("Garage slab install", 4455, 4465)
        s.replace(node, "Garage slab install", "Walkway construction")
        s.save()
    """

    def __init__(self, work_path, date, author="Arborist", start_id=None):
        """
        Args:
            work_path: Path to [project]/.work directory.
            date: Date string, e.g. "2026-02-25" (T00:00:00Z appended automatically).
            author: Tracked change author name.
            start_id: First change ID to use. Defaults to None = auto-detect from DOM
                      (max existing w:del/w:ins id + 1). Pass an explicit int to override.
        """
        if not os.path.isabs(work_path):
            work_path = os.path.join("/home/serg/projects/arborist-construction", work_path)
        self._work_path = work_path
        self._doc_path = os.path.join(work_path, "unpacked/temp/word/document.xml")

        # Normalize date — accept "2026-02-25" or full ISO
        self.date = date if 'T' in date else f"{date}T00:00:00Z"
        self.author = author

        # Load document
        with open(self._doc_path, 'r', encoding='utf-8') as f:
            self.content = f.read()
        self.dom = minidom.parseString(self.content.encode('utf-8'))

        # Auto-detect start_id from existing tracked changes if not specified
        if start_id is None:
            max_id = 0
            for tag in ('w:del', 'w:ins'):
                for elem in self.dom.getElementsByTagName(tag):
                    id_val = elem.getAttribute('w:id')
                    if id_val:
                        try:
                            max_id = max(max_id, int(id_val))
                        except ValueError:
                            pass
            start_id = max_id + 1
            if max_id > 0:
                print(f"Auto-detected start_id={start_id} (max existing ID was {max_id})")
        self._start_id = start_id
        self._next_id = start_id

    # ── ID management ──

    def next_id(self):
        """Return next tracked-change ID and increment counter."""
        id_ = self._next_id
        self._next_id += 1
        return id_

    @property
    def last_id(self):
        """Last ID that was issued (for logging)."""
        return self._next_id - 1

    # ── XML generators ──

    def del_run(self, text, rpr=None):
        """Generate <w:del> XML wrapping a deletion run."""
        if rpr is None:
            rpr = RPR_NORMAL
        id_ = self.next_id()
        return (f'<w:del w:id="{id_}" w:author="{self.author}" w:date="{self.date}">'
                f'<w:r>{rpr}<w:delText xml:space="preserve">{text}</w:delText></w:r>'
                f'</w:del>')

    def ins_run(self, text, rpr=None):
        """Generate <w:ins> XML wrapping an insertion run."""
        if rpr is None:
            rpr = RPR_NORMAL
        id_ = self.next_id()
        return (f'<w:ins w:id="{id_}" w:author="{self.author}" w:date="{self.date}">'
                f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'
                f'</w:ins>')

    def replace_text(self, node, old_text, new_text, rpr=None):
        """Replace a run node: <w:del>old</w:del><w:ins>new</w:ins>.

        Uses the node's own rPr if rpr is not specified."""
        if rpr is None:
            rpr = extract_rpr(node) or RPR_NORMAL
        xml = self.del_run(old_text, rpr) + self.ins_run(new_text, rpr)
        replace_node_with_xml(self.dom, node, xml)

    def ins_para(self, text, rpr=None, ppr=None):
        """Generate a full tracked-insertion paragraph XML string."""
        if rpr is None:
            rpr = RPR_NORMAL
        if ppr is None:
            ppr = PPR
        id_ = self.next_id()
        return (f'<w:p>{ppr}'
                f'<w:ins w:id="{id_}" w:author="{self.author}" w:date="{self.date}">'
                f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'
                f'</w:ins></w:p>')

    # ── Node finding (convenience wrappers) ──

    def find_run(self, text, line_lo=None, line_hi=None):
        """Find a run by text. If line_lo/line_hi given, disambiguate by line range."""
        if line_lo is not None and line_hi is not None:
            return find_run_in_line_range(self.content, self.dom, text, line_lo, line_hi)
        runs = [r for r in self.dom.getElementsByTagName('w:r') if get_text(r) == text]
        if len(runs) == 1:
            return runs[0]
        if not runs:
            raise ValueError(f"No run found with text '{text}'")
        raise ValueError(f"Ambiguous: {len(runs)} runs match '{text}' — provide line_lo/line_hi")

    def find_para(self, para_id):
        """Find paragraph by w14:paraId."""
        return find_para_by_para_id(self.dom, para_id)

    # ── Phrase-level editing ──

    def replace_phrase_in_run(self, run_node, phrase, replacement, rpr=None):
        """Surgically replace a phrase within a run, tracking only the changed words.

        Splits run into: prefix + <w:del>phrase</w:del><w:ins>replacement</w:ins> + suffix.
        Use for small phrase changes (tone fixes, value swaps) in long paragraphs.
        For full-paragraph rewrites (>50% changed), use replace_text() instead."""
        if rpr is None:
            rpr = extract_rpr(run_node) or RPR_NORMAL
        full_text = get_text(run_node)
        idx = full_text.find(phrase)
        if idx == -1:
            raise ValueError(f"Phrase '{phrase}' not found in run text '{full_text[:80]}'")
        prefix = full_text[:idx]
        suffix = full_text[idx + len(phrase):]
        rsid = run_node.getAttribute('w:rsidR')
        rsid_attr = f' w:rsidR="{rsid}"' if rsid else ''
        parts = []
        if prefix:
            parts.append(f'<w:r{rsid_attr}>{rpr}<w:t xml:space="preserve">{prefix}</w:t></w:r>')
        parts.append(self.del_run(phrase, rpr))
        parts.append(self.ins_run(replacement, rpr))
        if suffix:
            parts.append(f'<w:r{rsid_attr}>{rpr}<w:t xml:space="preserve">{suffix}</w:t></w:r>')
        replace_node_with_xml(self.dom, run_node, ''.join(parts))

    # ── Save ──

    def save(self):
        """Write modified document.xml back to disk."""
        with open(self._doc_path, 'wb') as f:
            f.write(self.dom.toxml(encoding='utf-8'))
        print(f"Saved. Change IDs used: {self._start_id}–{self.last_id}")
