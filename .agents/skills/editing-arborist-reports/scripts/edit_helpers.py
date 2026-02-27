"""Shared helpers for arborist report tracked-change editing.

Usage:
    sys.path.insert(0, "/home/serg/projects/arborist-construction/.agents/skills/editing-arborist-reports/scripts")
    from edit_helpers import EditSession, insert_xml_after, replace_node_with_xml, get_text
    from edit_helpers import prev_element_sibling, next_element_sibling, extract_rpr
    from edit_helpers import tc, impact_row, injury_row, sec4_row, mini_table, injury_detail_table
    from edit_helpers import RPR_NORMAL, RPR_BOLD
"""

import sys, os, shutil, xml.dom.minidom as minidom

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
        f'<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
        f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
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


def prev_element_sibling(node):
    """Walk to previous sibling, skipping whitespace text nodes."""
    sib = node.previousSibling
    while sib and sib.nodeType == sib.TEXT_NODE:
        sib = sib.previousSibling
    return sib


def next_element_sibling(node):
    """Walk to next sibling, skipping whitespace text nodes."""
    sib = node.nextSibling
    while sib and sib.nodeType == sib.TEXT_NODE:
        sib = sib.nextSibling
    return sib


def load_document(work_path):
    """Load document.xml from an unpacked .work directory.

    Args:
        work_path: Path to [project]/.work (e.g. "work/27 Shudell Avenue/.work")
                   Can be absolute or relative to PROJECT_ROOT.

    Returns:
        (dom, content) — parsed minidom Document and raw XML string.
    """
    if not os.path.isabs(work_path):
        work_path = os.path.join(os.environ.get("PROJECT_ROOT", "/home/serg/projects/arborist-construction"), work_path)
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

    def __init__(self, work_path, date, author="Arborist", start_id=None, rsid=None):
        """
        Args:
            work_path: Path to [project]/.work directory.
            date: Date string, e.g. "2026-02-25" (T00:00:00Z appended automatically).
            author: Tracked change author name.
            start_id: First change ID to use. Defaults to None = auto-detect from DOM
                      (max existing w:del/w:ins id + 1). Pass an explicit int to override.
            rsid: RSID from unpack output, used as w:rsidR on new runs/rows.
        """
        if not os.path.isabs(work_path):
            work_path = os.path.join(os.environ.get("PROJECT_ROOT", "/home/serg/projects/arborist-construction"), work_path)
        self._work_path = work_path
        self._doc_path = os.path.join(work_path, "unpacked/temp/word/document.xml")

        # Normalize date — accept "2026-02-25" or full ISO
        self.date = date if 'T' in date else f"{date}T00:00:00Z"
        self.author = author
        self.rsid = rsid or ""

        # Backup document.xml before any edits
        self._backup_path = self._doc_path + '.bak'
        shutil.copy2(self._doc_path, self._backup_path)

        # Load document
        with open(self._doc_path, 'r', encoding='utf-8') as f:
            self.content = f.read()
        self.dom = minidom.parseString(self.content.encode('utf-8'))

        # Collect existing paraIds for collision detection
        self._used_para_ids = set()
        for p in self.dom.getElementsByTagName('w:p'):
            pid = p.getAttribute('w14:paraId')
            if pid:
                self._used_para_ids.add(pid)
        for tr in self.dom.getElementsByTagName('w:tr'):
            pid = tr.getAttribute('w14:paraId')
            if pid:
                self._used_para_ids.add(pid)

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

    @property
    def date_short(self):
        """Date without time component, e.g. '2026-02-25'."""
        return self.date.split("T")[0]

    def generate_para_id(self, prefix, suffix=""):
        """Generate a paraId with collision detection.

        Tries prefix+suffix first; if it collides, appends incrementing hex digits.
        Registers the result so future calls won't collide either."""
        candidate = f"{prefix}{suffix}"
        counter = 0
        while candidate in self._used_para_ids:
            counter += 1
            candidate = f"{prefix}{counter:X}{suffix}"
        self._used_para_ids.add(candidate)
        return candidate

    # ── XML generators ──

    def del_run(self, text, rpr=None):
        """Generate <w:del> XML wrapping a deletion run."""
        if rpr is None:
            rpr = RPR_NORMAL
        id_ = self.next_id()
        text_esc = _escape_xml(text)
        return (f'<w:del w:id="{id_}" w:author="{self.author}" w:date="{self.date}">'
                f'<w:r>{rpr}<w:delText xml:space="preserve">{text_esc}</w:delText></w:r>'
                f'</w:del>')

    def ins_run(self, text, rpr=None):
        """Generate <w:ins> XML wrapping an insertion run."""
        if rpr is None:
            rpr = RPR_NORMAL
        id_ = self.next_id()
        text_esc = _escape_xml(text)
        return (f'<w:ins w:id="{id_}" w:author="{self.author}" w:date="{self.date}">'
                f'<w:r>{rpr}<w:t xml:space="preserve">{text_esc}</w:t></w:r>'
                f'</w:ins>')

    def replace_text(self, node, old_text, new_text, rpr=None):
        """Replace a run node: <w:del>old</w:del><w:ins>new</w:ins>.

        Uses the node's own rPr if rpr is not specified.
        Raises ValueError if node text doesn't match old_text."""
        actual = get_text(node)
        if actual != old_text:
            raise ValueError(
                f"replace_text mismatch: expected '{old_text}' but node contains '{actual[:80]}'"
            )
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
        text_esc = _escape_xml(text)
        return (f'<w:p>{ppr}'
                f'<w:ins w:id="{id_}" w:author="{self.author}" w:date="{self.date}">'
                f'<w:r>{rpr}<w:t xml:space="preserve">{text_esc}</w:t></w:r>'
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
            parts.append(f'<w:r{rsid_attr}>{rpr}<w:t xml:space="preserve">{_escape_xml(prefix)}</w:t></w:r>')
        parts.append(self.del_run(phrase, rpr))
        parts.append(self.ins_run(replacement, rpr))
        if suffix:
            parts.append(f'<w:r{rsid_attr}>{rpr}<w:t xml:space="preserve">{_escape_xml(suffix)}</w:t></w:r>')
        replace_node_with_xml(self.dom, run_node, ''.join(parts))

    # ── Bulk editing ──

    def delete_para(self, para_node):
        """Track-delete all runs in a paragraph, preserving the w:p element and pPr."""
        for r in list(para_node.getElementsByTagName('w:r')):
            txt = get_text(r)
            if txt:
                rpr = extract_rpr(r) or None
                del_xml = self.del_run(txt, rpr)
                replace_node_with_xml(self.dom, r, del_xml)

    def delete_row(self, tr_node):
        """Track-delete an entire table row (w:del in trPr + del_run on all cell runs)."""
        trpr_nodes = tr_node.getElementsByTagName('w:trPr')
        if trpr_nodes:
            trpr = trpr_nodes[0]
        else:
            trpr = self.dom.createElement('w:trPr')
            tr_node.insertBefore(trpr, tr_node.firstChild)
        del_id = self.next_id()
        del_elem = self.dom.createElement('w:del')
        del_elem.setAttribute('w:id', str(del_id))
        del_elem.setAttribute('w:author', self.author)
        del_elem.setAttribute('w:date', self.date)
        trpr.appendChild(del_elem)
        for tc_node in tr_node.getElementsByTagName('w:tc'):
            for p in tc_node.getElementsByTagName('w:p'):
                for r in list(p.getElementsByTagName('w:r')):
                    rpr = extract_rpr(r) or None
                    txt = get_text(r)
                    if txt:
                        del_xml = self.del_run(txt, rpr)
                        replace_node_with_xml(self.dom, r, del_xml)

    def replace_in_para(self, para_node, old_text, new_text):
        """Find-and-replace the first run in a paragraph whose text matches old_text."""
        for r in para_node.getElementsByTagName('w:r'):
            if get_text(r) == old_text:
                self.replace_text(r, old_text, new_text)
                return
        raise ValueError(f"No run with text '{old_text}' in paragraph {para_node.getAttribute('w14:paraId')}")

    # ── Row/table finding ──

    def find_tr(self, para_id):
        """Find w:tr with matching w14:paraId attribute."""
        for tr in self.dom.getElementsByTagName('w:tr'):
            if tr.getAttribute('w14:paraId') == para_id:
                return tr
        raise ValueError(f"Table row with paraId={para_id} not found")

    # ── Validation ──

    def validate_targets(self, targets):
        """Pre-flight check: verify all edit targets exist before mutations.

        Args:
            targets: list of (paraId, expected_tag, label) tuples.
                expected_tag: 'w:p' for paragraphs, 'w:tr' for table rows.
                label: human-readable description for error messages.

        Raises:
            ValueError listing all missing targets (not just the first).
        """
        missing = []
        for para_id, expected_tag, label in targets:
            found = False
            for node in self.dom.getElementsByTagName(expected_tag):
                if node.getAttribute('w14:paraId') == para_id:
                    found = True
                    break
            if not found:
                missing.append(f"  {label}: paraId={para_id} expected {expected_tag}")
        if missing:
            raise ValueError(
                f"validate_targets failed — {len(missing)} target(s) not found:\n"
                + "\n".join(missing)
            )

    # ── Context manager ──

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None and os.path.exists(self._backup_path):
            self.rollback()
            print(f"Auto-rolled back due to {exc_type.__name__}: {exc_val}")
        return False  # don't suppress the exception

    # ── Save / Rollback ──

    def rollback(self):
        """Restore document.xml from backup and re-parse DOM."""
        if not os.path.exists(self._backup_path):
            raise FileNotFoundError("No backup file to rollback to")
        shutil.copy2(self._backup_path, self._doc_path)
        with open(self._doc_path, 'r', encoding='utf-8') as f:
            self.content = f.read()
        self.dom = minidom.parseString(self.content.encode('utf-8'))
        self._next_id = self._start_id
        print(f"Rolled back to backup. IDs reset to {self._start_id}.")

    def save(self):
        """Write modified document.xml back to disk."""
        with open(self._doc_path, 'wb') as f:
            f.write(self.dom.toxml(encoding='utf-8'))
        # Clean up backup after successful save
        if os.path.exists(self._backup_path):
            os.remove(self._backup_path)
        print(f"Saved. Change IDs used: {self._start_id}–{self.last_id}")


# ─── Table-type RPR constants ───────────────────────────────────────────────
# These match the formatting found in real report tables via get_schema.py.

RPR_IMPACT = '<w:rFonts w:cs="Helvetica"/><w:szCs w:val="22"/>'
RPR_INJURY = '<w:rFonts w:cs="Helvetica"/><w:szCs w:val="22"/>'
RPR_SEC4 = '<w:rFonts w:cs="Helvetica"/><w:color w:val="000000"/><w:sz w:val="16"/><w:szCs w:val="16"/><w:lang w:val="en-PH" w:eastAsia="en-PH"/>'
RPR_MINI = '<w:rFonts w:cs="Helvetica"/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-PH" w:eastAsia="en-PH"/>'
RPR_MINI_HDR = '<w:rFonts w:cs="Helvetica"/><w:b/><w:bCs/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-PH" w:eastAsia="en-PH"/>'


# ─── Builder functions ──────────────────────────────────────────────────────
# These generate tracked-insertion XML for table rows and mini-tables.
# All take an EditSession as first arg; IDs come from sess.next_id().

def _escape_xml(text):
    """Escape XML special characters in text content."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def tc(width, wtype, text, rpr, centered=True, borders="single", fill="none",
       ins_id=None, author="Arborist", date="2026-02-26", left_border="single",
       top_border="single"):
    """Build a <w:tc> table cell with optional tracked insertion on the text run."""
    jc = '<w:jc w:val="center"/>' if centered else ''
    top = f'<w:top w:val="{top_border}" w:sz="4" w:space="0" w:color="auto"/>' if top_border != "nil" else '<w:top w:val="nil"/>'
    left = f'<w:left w:val="{left_border}" w:sz="4" w:space="0" w:color="auto"/>' if left_border != "nil" else '<w:left w:val="nil"/>'
    shd = f'<w:shd w:val="clear" w:color="000000" w:fill="{fill}"/>' if fill != "none" else ''
    ins_open = f'<w:ins w:id="{ins_id}" w:author="{author}" w:date="{date}T00:00:00Z">' if ins_id else ''
    ins_close = '</w:ins>' if ins_id else ''
    text_esc = _escape_xml(text)
    return f'''<w:tc>
  <w:tcPr>
    <w:tcW w:w="{width}" w:type="{wtype}"/>
    <w:tcBorders>
      {top}
      {left}
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tcBorders>
    {shd}
    <w:vAlign w:val="center"/>
    <w:hideMark/>
  </w:tcPr>
  <w:p>
    <w:pPr>{jc}<w:rPr>{rpr}</w:rPr></w:pPr>
    {ins_open}<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text_esc}</w:t></w:r>{ins_close}
  </w:p>
</w:tc>'''


def impact_row(sess, tree_num, source, result, rpr=None):
    """Build a 3-col impact table row (pct widths: 528/3563/909) as tracked insertion.

    rpr: override run properties (extracted from live document via get_schema.py).
         Falls back to RPR_IMPACT if not provided.
    """
    rpr = rpr or RPR_IMPACT
    row_id = sess.next_id()
    pid = sess.generate_para_id(f"1A0{tree_num:04X}", "0")
    return (
        f'<w:tr w:rsidR="{sess.rsid}" w14:paraId="{pid}" w14:textId="77777777" w:rsidTr="{sess.rsid}">'
        f'<w:trPr><w:trHeight w:val="679"/>'
        f'<w:ins w:id="{row_id}" w:author="{sess.author}" w:date="{sess.date}"/></w:trPr>'
        + tc("528",  "pct", str(tree_num), rpr, centered=True, ins_id=sess.next_id(), author=sess.author, date=sess.date.split("T")[0])
        + tc("3563", "pct", source,        rpr, centered=True, ins_id=sess.next_id(), author=sess.author, date=sess.date.split("T")[0])
        + tc("909",  "pct", result,        rpr, centered=True, ins_id=sess.next_id(), author=sess.author, date=sess.date.split("T")[0])
        + '</w:tr>'
    )


def injury_row(sess, source, distance, depth, rating, rpr=None):
    """Build a 4-col injury detail table row (pct widths: 1702/1096/853/1349) as tracked insertion.

    rpr: override run properties. Falls back to RPR_INJURY if not provided.
    """
    rpr = rpr or RPR_INJURY
    row_id = sess.next_id()
    pid = sess.generate_para_id(f"2A0{row_id:04X}", "0")
    return (
        f'<w:tr w:rsidR="{sess.rsid}" w14:paraId="{pid}" w14:textId="77777777" w:rsidTr="{sess.rsid}">'
        f'<w:trPr><w:trHeight w:val="763"/>'
        f'<w:ins w:id="{row_id}" w:author="{sess.author}" w:date="{sess.date}"/></w:trPr>'
        + tc("1702", "pct", source,   rpr, centered=True, ins_id=sess.next_id(), author=sess.author, date=sess.date.split("T")[0])
        + tc("1096", "pct", distance, rpr, centered=True, ins_id=sess.next_id(), author=sess.author, date=sess.date.split("T")[0])
        + tc("853",  "pct", depth,    rpr, centered=True, ins_id=sess.next_id(), author=sess.author, date=sess.date.split("T")[0])
        + tc("1349", "pct", rating,   rpr, centered=True, ins_id=sess.next_id(), author=sess.author, date=sess.date.split("T")[0])
        + '</w:tr>'
    )


def sec4_row(sess, cols, fill="FFFFFF", rpr=None):
    """Build a 10-col Section 4 data row (dxa widths).

    fill: "FFFFFF" (white, nil top/left borders) or "F2F2F2" (gray, single all sides)
    cols: list of 10 text values matching Section 4 column order.
    rpr: override run properties. Falls back to RPR_SEC4 if not provided.
         IMPORTANT: RPR_SEC4 uses sz=16/en-PH which may not match all reports.
         Extract the actual rPr from an existing data row via get_schema.py.
    """
    rpr = rpr or RPR_SEC4
    widths = [789, 1158, 1300, 675, 1163, 4281, 1217, 1062, 779, 1393]
    row_id = sess.next_id()
    pid = sess.generate_para_id(f"3A0{row_id:04X}", "0")
    date_short = sess.date.split("T")[0]
    row_xml = (
        f'<w:tr w:rsidR="{sess.rsid}" w14:paraId="{pid}" w14:textId="77777777" w:rsidTr="{sess.rsid}">'
        f'<w:trPr><w:trHeight w:val="347"/>'
        f'<w:ins w:id="{row_id}" w:author="{sess.author}" w:date="{sess.date}"/></w:trPr>'
    )
    for i, (w, text) in enumerate(zip(widths, cols)):
        centered = (i != 5)
        top_b = "nil" if fill == "FFFFFF" else "single"
        lft_b = "nil" if fill == "FFFFFF" else "single"
        row_xml += tc(str(w), "dxa", text, rpr, centered=centered,
                      fill=fill, ins_id=sess.next_id(), top_border=top_b, left_border=lft_b,
                      author=sess.author, date=date_short)
    row_xml += '</w:tr>'
    return row_xml


def mini_table(sess, tree_num, species, dbh, condition, comments, ownership, direction, tpz,
               tblpX="1513", tblpY="2322", hdr_rpr=None, data_rpr=None):
    """Build a full 8-col floating mini-table (header row + data row) as tracked insertion.

    tblpX/tblpY: absolute page-position coordinates for the floating table.
    Must be extracted from an existing mini-table via get_schema.py for each report.
    hdr_rpr: override header run properties. Falls back to RPR_MINI_HDR.
    data_rpr: override data row run properties. Falls back to RPR_MINI.
    """
    hdr_rpr = hdr_rpr or RPR_MINI_HDR
    data_rpr = data_rpr or RPR_MINI
    hdr_cols = ["TREE #", "Species", "DBH (cm)", "Condition Rating",
                "Comments (* denotes approx DBH)", "Ownership\nCategory", "Direction", "TPZ (m)"]
    hdr_widths = [839, 1129, 707, 1271, 2789, 1275, 1083, 678]
    date_short = sess.date.split("T")[0]

    hdr_row_id = sess.next_id()
    hdr_pid = sess.generate_para_id(f"4A0{hdr_row_id:04X}", "0")
    hdr_row = (
        f'<w:tr w:rsidR="{sess.rsid}" w14:paraId="{hdr_pid}" w14:textId="77777777" w:rsidTr="{sess.rsid}">'
        f'<w:trPr><w:trHeight w:val="809"/>'
        f'<w:ins w:id="{hdr_row_id}" w:author="{sess.author}" w:date="{sess.date}"/></w:trPr>'
    )
    for i, (w, text) in enumerate(zip(hdr_widths, hdr_cols)):
        lft = "single" if i == 0 else "nil"
        if "\n" in text:
            parts = text.split("\n")
            iid = sess.next_id()
            hdr_row += (
                f'<w:tc><w:tcPr>'
                f'<w:tcW w:w="{w}" w:type="dxa"/>'
                f'<w:tcBorders>'
                f'<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'<w:left w:val="{lft}" w:sz="4" w:space="0" w:color="auto"/>'
                f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'</w:tcBorders>'
                f'<w:shd w:val="clear" w:color="000000" w:fill="F2F2F2"/>'
                f'<w:vAlign w:val="center"/><w:hideMark/></w:tcPr>'
                f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{hdr_rpr}</w:rPr></w:pPr>'
                f'<w:ins w:id="{iid}" w:author="{sess.author}" w:date="{sess.date}">'
                f'<w:r><w:rPr>{hdr_rpr}</w:rPr>'
                f'<w:t>{parts[0]}</w:t><w:br/><w:t>{parts[1]}</w:t>'
                f'</w:r></w:ins></w:p></w:tc>'
            )
        else:
            hdr_row += tc(str(w), "dxa", text, hdr_rpr, centered=True,
                          fill="F2F2F2", ins_id=sess.next_id(),
                          top_border="single", left_border=lft,
                          author=sess.author, date=date_short)
    hdr_row += '</w:tr>'

    data_cols = [str(tree_num), species, str(dbh), condition, comments, ownership, direction, str(tpz)]
    data_row_id = sess.next_id()
    data_pid = sess.generate_para_id(f"4A0{data_row_id:04X}", "1")
    data_row = (
        f'<w:tr w:rsidR="{sess.rsid}" w14:paraId="{data_pid}" w14:textId="77777777" w:rsidTr="{sess.rsid}">'
        f'<w:trPr><w:trHeight w:val="609"/>'
        f'<w:ins w:id="{data_row_id}" w:author="{sess.author}" w:date="{sess.date}"/></w:trPr>'
    )
    for i, (w, text) in enumerate(zip(hdr_widths, data_cols)):
        lft = "single" if i == 0 else "nil"
        data_row += tc(str(w), "dxa", text, data_rpr, centered=(i != 4),
                       fill="F2F2F2", ins_id=sess.next_id(), top_border="nil", left_border=lft,
                       author=sess.author, date=date_short)
    data_row += '</w:tr>'

    return (
        f'<w:tbl><w:tblPr>'
        f'<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="page" w:horzAnchor="page" w:tblpX="{tblpX}" w:tblpY="{tblpY}"/>'
        f'<w:tblW w:w="9771" w:type="dxa"/>'
        f'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        f'</w:tblPr>'
        f'<w:tblGrid>'
        f'<w:gridCol w:w="839"/><w:gridCol w:w="1129"/><w:gridCol w:w="707"/>'
        f'<w:gridCol w:w="1271"/><w:gridCol w:w="2789"/><w:gridCol w:w="1275"/>'
        f'<w:gridCol w:w="1083"/><w:gridCol w:w="678"/>'
        f'</w:tblGrid>'
        + hdr_row + data_row + '</w:tbl>'
    )


def injury_detail_table(sess, rows_data, hdr_rpr=None, data_rpr=None,
                        tblpX=None, tblpY=None):
    """Build a complete 4-col injury detail table with header + data rows.

    rows_data: list of (source, distance, depth, rating) tuples.
    hdr_rpr: override header run properties. Falls back to RPR_MINI_HDR.
    data_rpr: override data row run properties. Falls back to RPR_INJURY.
              Passed through to injury_row() calls.
    tblpX/tblpY: if provided, makes the table floating (absolute page position).
                 Extract from existing injury detail table via get_schema.py.
                 If omitted, table is in-flow (no positioning).
    """
    hdr_rpr = hdr_rpr or RPR_MINI_HDR
    date_short = sess.date.split("T")[0]
    hdr_row_id = sess.next_id()
    hdr_pid = sess.generate_para_id(f"5A0{hdr_row_id:04X}", "0")
    hdr_row = (
        f'<w:tr w:rsidR="{sess.rsid}" w14:paraId="{hdr_pid}" w14:textId="77777777" w:rsidTr="{sess.rsid}">'
        f'<w:trPr><w:trHeight w:val="809"/>'
        f'<w:ins w:id="{hdr_row_id}" w:author="{sess.author}" w:date="{sess.date}"/></w:trPr>'
        + tc("1702", "pct", "Injury source",               hdr_rpr, centered=True, fill="F2F2F2", ins_id=sess.next_id(), author=sess.author, date=date_short)
        + tc("1096", "pct", "Closest point of impact (m)", hdr_rpr, centered=True, fill="F2F2F2", ins_id=sess.next_id(), author=sess.author, date=date_short)
        + tc("853",  "pct", "Max depth of excavation",     hdr_rpr, centered=True, fill="F2F2F2", ins_id=sess.next_id(), author=sess.author, date=date_short)
        + tc("1349", "pct", "Impact to condition",         hdr_rpr, centered=True, fill="F2F2F2", ins_id=sess.next_id(), author=sess.author, date=date_short)
        + '</w:tr>'
    )
    data_rows = "".join(injury_row(sess, src, dist, dep, rat, rpr=data_rpr) for src, dist, dep, rat in rows_data)
    tblp = ''
    if tblpX is not None and tblpY is not None:
        tblp = (f'<w:tblpPr w:leftFromText="180" w:rightFromText="180"'
                f' w:vertAnchor="page" w:horzAnchor="page"'
                f' w:tblpX="{tblpX}" w:tblpY="{tblpY}"/>')
    return (
        f'<w:tbl><w:tblPr>'
        f'{tblp}'
        f'<w:tblW w:w="5000" w:type="pct"/>'
        f'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        f'</w:tblPr>'
        f'<w:tblGrid>'
        f'<w:gridCol w:w="1702"/><w:gridCol w:w="1096"/>'
        f'<w:gridCol w:w="853"/><w:gridCol w:w="1349"/>'
        f'</w:tblGrid>'
        + hdr_row + data_rows + '</w:tbl>'
    )
