"""Map an arborist report's document structure for edit-script writing.

Usage:
    python3 map_report.py <work_path>

work_path: path to [project]/.work directory (e.g. "work/8 Glen Agar Drive/.work")
Looks for unpacked/temp/word/document.xml first; falls back to unpacked/current/.
"""

import sys, os, json, re
import xml.dom.minidom as minidom

PROJECT_ROOT = os.environ.get("PROJECT_ROOT", "/home/serg/projects/arborist-construction")
sys.path.insert(0, os.path.join(PROJECT_ROOT, ".agents/skills/editing-arborist-reports/scripts"))
from edit_helpers import get_text

HEADING_STYLES = {'Heading1', 'Heading2', 'Heading3', 'Heading4', 'pagebreak2'}


# ── helpers ──────────────────────────────────────────────────────────────────

def _resolve_doc_path(work_path):
    if not os.path.isabs(work_path):
        work_path = os.path.join(PROJECT_ROOT, work_path)
    for sub in ("temp", "current"):
        p = os.path.join(work_path, "unpacked", sub, "word", "document.xml")
        if os.path.exists(p):
            return work_path, p
    raise FileNotFoundError(f"No document.xml under {work_path}/unpacked/temp|current/")


def _load(doc_path):
    with open(doc_path, encoding='utf-8') as f:
        content = f.read()
    dom = minidom.parseString(content.encode('utf-8'))
    return dom, content


def _build_line_index(content):
    """Map w14:paraId → opening line number; also collect <w:tbl> line numbers."""
    para_id_to_line = {}
    tbl_lines = []
    for i, line in enumerate(content.splitlines(), 1):
        m = re.search(r'<w:p\b[^>]*w14:paraId="([^"]+)"', line)
        if m:
            para_id_to_line[m.group(1)] = i
        if '<w:tbl>' in line or '<w:tbl ' in line:
            tbl_lines.append(i)
    return para_id_to_line, tbl_lines


def _para_style(p):
    for pPr in p.getElementsByTagName('w:pPr'):
        for node in pPr.childNodes:
            if node.nodeName == 'w:pStyle':
                return node.getAttribute('w:val')
    return None


def _preview(text, n=15):
    words = text.split()
    return ' '.join(words[:n]) + ('...' if len(words) > n else '')


def _nearest_before(line, items, key='line'):
    """Return text of the last item whose line < given line."""
    best = None
    for item in items:
        if item.get(key) and item[key] < line:
            best = item['text']
    return best


# ── extractors ───────────────────────────────────────────────────────────────

def _extract_headings(body, line_idx):
    headings = []
    for node in body.childNodes:
        if node.nodeName != 'w:p':
            continue
        style = _para_style(node)
        if style not in HEADING_STYLES:
            continue
        text = get_text(node).strip()
        if not text:
            continue
        pid = node.getAttribute('w14:paraId')
        headings.append({'text': text, 'style': style, 'para_id': pid,
                         'line': line_idx.get(pid)})
    return headings


def _extract_tables(body, tbl_lines, headings):
    tables = []
    tbl_nodes = [c for c in body.childNodes if c.nodeName == 'w:tbl']
    assert len(tbl_nodes) == len(tbl_lines), (
        f"Table count mismatch: {len(tbl_nodes)} DOM nodes vs {len(tbl_lines)} line-scanned "
        f"<w:tbl> tags. Nested tables may be causing misalignment."
    )
    for idx, tbl in enumerate(tbl_nodes):
        line = tbl_lines[idx] if idx < len(tbl_lines) else None
        rows = tbl.getElementsByTagName('w:tr')
        if not rows:
            continue
        cols = [get_text(tc).strip() for tc in rows[0].getElementsByTagName('w:tc')]
        # Cell-level data for all rows
        row_data = []
        for ri, row in enumerate(rows):
            row_pid = row.getAttribute('w14:paraId')
            cells = row.getElementsByTagName('w:tc')
            cell_data = []
            for ci, cell in enumerate(cells):
                paras = cell.getElementsByTagName('w:p')
                cell_pid = paras[0].getAttribute('w14:paraId') if paras else None
                cell_text = get_text(cell).strip()
                entry = {'col': ci, 'text': cell_text[:60], 'para_id': cell_pid}
                if ci < len(cols):
                    entry['col_name'] = cols[ci]
                cell_data.append(entry)
            row_data.append({'row_index': ri, 'row_para_id': row_pid, 'cells': cell_data})
        tables.append({
            'table_index': idx,
            'line': line,
            'context_heading': _nearest_before(line or 0, headings),
            'columns': cols,
            'row_count': len(rows),
            'rows': row_data,
        })
    return tables


def _collect_after(body, anchor_pid, line_idx, stop_styles, stop_on_tree=False):
    nodes = list(body.childNodes)
    anchor_i = next((i for i, n in enumerate(nodes)
                     if n.nodeName == 'w:p' and n.getAttribute('w14:paraId') == anchor_pid), None)
    if anchor_i is None:
        return []
    result = []
    for node in nodes[anchor_i + 1:]:
        if node.nodeName == 'w:tbl':
            result.append({'type': 'table', 'para_id': None,
                           'line_lo': None, 'line_hi': None, 'preview': '[TABLE]',
                           'has_tracked_changes': False})
            continue
        if node.nodeName != 'w:p':
            continue
        style = _para_style(node)
        if style in stop_styles:
            break
        text = get_text(node).strip()
        if stop_on_tree and re.fullmatch(r'Tree\s+\d+\s*:?', text, re.IGNORECASE):
            break
        if not text:
            continue
        pid = node.getAttribute('w14:paraId')
        line = line_idx.get(pid)
        has_chg = bool(node.getElementsByTagName('w:del') or node.getElementsByTagName('w:ins'))
        result.append({'para_id': pid, 'line_lo': line, 'line_hi': line,
                       'preview': _preview(text), 'has_tracked_changes': has_chg})
        if len(result) >= 20:
            break
    # Fix line_hi: each paragraph ends just before the next one starts
    paras = [p for p in result if p.get('line_lo')]
    for i, p in enumerate(paras[:-1]):
        p['line_hi'] = (paras[i + 1]['line_lo'] - 1) if paras[i + 1]['line_lo'] else p['line_lo']
    return result


def _extract_tree_sections(body, line_idx, headings):
    results = []
    # Phase 1: Heading2/3 labelled tree sections (Section 3)
    for h in headings:
        m = re.search(r'\bTree\s+(\d+)\b', h['text'], re.IGNORECASE)
        if m and h['style'] in ('Heading2', 'Heading3', 'Heading4'):
            paras = _collect_after(body, h['para_id'], line_idx, HEADING_STYLES)
            ctx = next((hh['text'] for hh in reversed(headings)
                        if hh['style'] == 'Heading1' and (hh['line'] or 0) < (h['line'] or 0)), None)
            results.append({'tree_num': int(m.group(1)), 'heading_text': h['text'],
                            'heading_para_id': h['para_id'], 'heading_line': h['line'],
                            'section_context': ctx, 'paragraphs': paras})
    # Phase 2: Bold "Tree N:" label paragraphs (Section 5/6)
    for node in body.childNodes:
        if node.nodeName != 'w:p':
            continue
        style = _para_style(node)
        if style in HEADING_STYLES:
            continue
        text = get_text(node).strip()
        m = re.fullmatch(r'Tree\s+(\d+)\s*:?', text, re.IGNORECASE)
        if not m:
            continue
        if not any(r.getAttribute('w:val') == 'Strong'
                   for r in node.getElementsByTagName('w:rStyle')):
            continue
        pid = node.getAttribute('w14:paraId')
        line = line_idx.get(pid)
        paras = _collect_after(body, pid, line_idx,
                               HEADING_STYLES | {'pagebreak2'}, stop_on_tree=True)
        ctx = _nearest_before(line or 0, [h for h in headings if h['style'] == 'Heading1'])
        results.append({'tree_num': int(m.group(1)), 'heading_text': text,
                        'heading_para_id': pid, 'heading_line': line,
                        'section_context': ctx, 'paragraphs': paras})
    results.sort(key=lambda x: x['heading_line'] or 0)
    return results


def _extract_summary(body, headings, line_idx):
    h = next((h for h in headings
               if re.search(r'\bsummary\b', h['text'], re.IGNORECASE) and h['style'] == 'Heading1'), None)
    if not h:
        return None
    return {'heading_para_id': h['para_id'], 'heading_line': h['line'],
            'paragraphs': _collect_after(body, h['para_id'], line_idx, HEADING_STYLES)}


def _extract_permit_bullets(body, headings, line_idx):
    """Extract permit summary bullets from Section 5 opening.

    Returns list of {para_id, text, line} for paragraphs matching
    permit count patterns (e.g. "Removal Permit requirement: 6 trees").
    """
    # Find Section 5 heading
    sec5 = next((h for h in headings
                 if re.search(r'Section\s*5', h['text'], re.IGNORECASE)), None)
    if not sec5:
        return []

    bullets = []
    nodes = list(body.childNodes)
    anchor_i = next((i for i, n in enumerate(nodes)
                     if n.nodeName == 'w:p' and n.getAttribute('w14:paraId') == sec5['para_id']), None)
    if anchor_i is None:
        return []

    # Scan paragraphs after Section 5 heading until we hit a table or another heading
    PERMIT_PATTERNS = [
        r'(?i)permit\s+requirement',
        r'(?i)to be removed',
        r'(?i)to be injured',
    ]
    for node in nodes[anchor_i + 1:]:
        if node.nodeName == 'w:tbl':
            break  # stop at the impact summary table
        if node.nodeName != 'w:p':
            continue
        style = _para_style(node)
        if style in HEADING_STYLES:
            break
        text = get_text(node).strip()
        if any(re.search(pat, text) for pat in PERMIT_PATTERNS):
            pid = node.getAttribute('w14:paraId')
            bullets.append({'para_id': pid, 'text': text, 'line': line_idx.get(pid)})
    return bullets


# ── public API ────────────────────────────────────────────────────────────────

def map_report(work_path) -> dict:
    work_path, doc_path = _resolve_doc_path(work_path)
    dom, content = _load(doc_path)
    body = dom.getElementsByTagName('w:body')[0]
    line_idx, tbl_lines = _build_line_index(content)
    headings = _extract_headings(body, line_idx)
    tables = _extract_tables(body, tbl_lines, headings)
    tree_sections = _extract_tree_sections(body, line_idx, headings)
    summary = _extract_summary(body, headings, line_idx)
    permit_bullets = _extract_permit_bullets(body, headings, line_idx)
    result = {'source_path': doc_path, 'headings': headings,
              'tables': tables, 'tree_sections': tree_sections,
              'summary': summary, 'permit_bullets': permit_bullets}
    out = os.path.join(work_path, 'map.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    return result


def _print_map(result):
    print("=== DOCUMENT MAP ===\n")
    print("── HEADINGS ──")
    indent_map = {'Heading1': '', 'Heading2': '  ', 'Heading3': '    ',
                  'Heading4': '      ', 'pagebreak2': '  [sub] '}
    for h in result['headings']:
        ind = indent_map.get(h['style'], '')
        print(f"  {ind}{h['style']:12}  L{h['line'] or '?':>5}  [{h['para_id']}]  {h['text']}")
    print("\n── TABLES ──")
    for t in result['tables']:
        print(f"  Table #{t['table_index']}  L{t['line']}  ({t['row_count']} rows)  after: {t['context_heading']!r}")
        print(f"    Columns: {t['columns']}")
        if t.get('rows'):
            for row in t['rows'][:2]:  # show header + first data row
                cells_preview = [(c['col_name'], c['text'][:15], c['para_id']) for c in row['cells'] if 'col_name' in c]
                if cells_preview:
                    print(f"    Row {row['row_index']} [{row['row_para_id']}]: {[(n, t) for n, t, _ in cells_preview[:4]]}...")
            if t['row_count'] > 2:
                print(f"    ... ({t['row_count'] - 2} more data rows)")
    print("\n── TREE SECTIONS ──")
    for ts in result['tree_sections']:
        print(f"  Tree {ts['tree_num']:>2}  [{ts['heading_para_id']}]  L{ts['heading_line'] or '?'}  in: {ts['section_context']!r}")
        for p in ts['paragraphs']:
            chg = ' *' if p.get('has_tracked_changes') else ''
            lo, hi = p.get('line_lo', '?'), p.get('line_hi', '?')
            print(f"    [{p['para_id'] or '—':>8}]  L{lo}-{hi}{chg}  {p['preview']}")
    print("\n── PERMIT BULLETS ──")
    for b in result.get('permit_bullets', []):
        print(f"  [{b['para_id']}]  L{b.get('line', '?')}  {b['text']}")

    print("\n── SUMMARY ──")
    s = result.get('summary')
    if s:
        print(f"  Heading [{s['heading_para_id']}]  L{s['heading_line']}")
        for p in s['paragraphs']:
            print(f"    [{p['para_id']}]  L{p['line_lo']}-{p['line_hi']}  {p['preview']}")


def main():
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <work_path>", file=sys.stderr)
        sys.exit(1)
    result = map_report(sys.argv[1])
    _print_map(result)
    print(f"\nJSON: {result['source_path'].replace('/unpacked/', '/map.json ← ')}")


if __name__ == '__main__':
    main()
