"""Standard schema extraction for arborist report editing.

Usage:
    python3 get_schema.py <work_path>

Reads map.json and document.xml from the work directory.
Outputs schema.json with:
  - Live rPr (inner XML) from each table type
  - Table positioning (floating vs in-flow, tblpPr attributes)
  - RPR constant validation (compares live values against edit_helpers defaults)

Run this once after unpack + map_report. The edit script imports schema.json
instead of hardcoding rPr or writing custom extraction scripts.
"""

import sys, os, json
import xml.dom.minidom as minidom

PROJECT_ROOT = os.environ.get("PROJECT_ROOT", "/home/serg/projects/arborist-construction")
sys.path.insert(0, os.path.join(PROJECT_ROOT, ".agents/skills/editing-arborist-reports/scripts"))
from edit_helpers import (load_document, get_text,
                          RPR_IMPACT, RPR_INJURY, RPR_SEC4, RPR_MINI, RPR_MINI_HDR)


def _inner_rpr(run_node):
    """Extract inner XML of w:rPr (children only, no wrapper tag)."""
    rpr_nodes = run_node.getElementsByTagName('w:rPr')
    if not rpr_nodes:
        return None
    return ''.join(
        child.toxml() for child in rpr_nodes[0].childNodes
        if child.nodeType != child.TEXT_NODE
    )


def _extract_table_rpr(dom, table_index):
    """Extract rPr from first data row of a table (skipping header)."""
    tables = dom.getElementsByTagName('w:tbl')
    if table_index >= len(tables):
        return None, None
    tbl = tables[table_index]
    rows = tbl.getElementsByTagName('w:tr')

    # Header rPr
    hdr_rpr = None
    if rows:
        for cell in rows[0].getElementsByTagName('w:tc'):
            for run in cell.getElementsByTagName('w:r'):
                hdr_rpr = _inner_rpr(run)
                if hdr_rpr:
                    break
            if hdr_rpr:
                break

    # Data rPr (from first data row)
    data_rpr = None
    if len(rows) > 1:
        for cell in rows[1].getElementsByTagName('w:tc'):
            for run in cell.getElementsByTagName('w:r'):
                data_rpr = _inner_rpr(run)
                if data_rpr:
                    break
            if data_rpr:
                break

    return hdr_rpr, data_rpr


def _extract_table_positioning(dom, table_index):
    """Extract tblpPr positioning from a table."""
    tables = dom.getElementsByTagName('w:tbl')
    if table_index >= len(tables):
        return None
    tbl = tables[table_index]
    tblpr = tbl.getElementsByTagName('w:tblPr')
    if not tblpr:
        return {'type': 'unknown'}
    tblp = tblpr[0].getElementsByTagName('w:tblpPr')
    if not tblp:
        return {'type': 'in-flow'}
    # Extract all attributes
    attrs = {}
    for attr in tblp[0].attributes.keys():
        attrs[attr] = tblp[0].getAttribute(attr)
    return {'type': 'floating', **attrs}


def extract_schema(work_path):
    """Extract schema from an unpacked report.

    Returns schema dict and writes schema.json to work directory.
    """
    if not os.path.isabs(work_path):
        work_path = os.path.join(PROJECT_ROOT, work_path)

    # Load map.json for table structure
    map_path = os.path.join(work_path, 'map.json')
    if not os.path.exists(map_path):
        raise FileNotFoundError(f"map.json not found at {map_path} — run map_report.py first")
    with open(map_path) as f:
        report_map = json.load(f)

    # Load document (try temp first, fall back to current)
    for sub in ("temp", "current"):
        doc_path = os.path.join(work_path, "unpacked", sub, "word", "document.xml")
        if os.path.exists(doc_path):
            break
    else:
        raise FileNotFoundError(f"No document.xml under {work_path}/unpacked/temp|current/")
    with open(doc_path, 'r', encoding='utf-8') as f:
        content = f.read()
    dom = minidom.parseString(content.encode('utf-8'))

    # Classify tables by context heading
    table_types = {}
    for t in report_map['tables']:
        ctx = (t.get('context_heading') or '').lower()
        cols = t.get('columns', [])
        idx = t['table_index']

        if 'section 4' in ctx and len(cols) >= 8:
            table_types['sec4'] = idx
        elif 'section 5' in ctx or 'conclusion' in ctx:
            col_str = '|'.join(cols).lower()
            if 'source of impact' in col_str:
                table_types['impact'] = idx
            elif 'injury source' in col_str:
                table_types['injury_detail'] = idx
            elif len(cols) >= 7 and 'direction' in col_str:
                # Mini-table (tree summary in Section 5)
                if 'mini' not in table_types:
                    table_types['mini'] = idx
        elif 'removal' in ctx:
            if 'removal' not in table_types:
                table_types['removal'] = idx
        elif 'replanting' in ctx or 'addendum 0' in ctx:
            table_types['replanting'] = idx

        # Summary tables
        if 'summary' in ctx:
            if 'replacements' in '|'.join(cols).lower() and len(cols) == 2:
                table_types['replacement'] = idx
            elif 'permits' in '|'.join(cols).lower():
                table_types['summary'] = idx

    # Extract rPr and positioning for each classified table
    schema = {'tables': {}, 'rpr': {}, 'rpr_validation': []}

    RPR_DEFAULTS = {
        'impact': ('RPR_IMPACT', RPR_IMPACT),
        'sec4': ('RPR_SEC4', RPR_SEC4),
        'mini': ('RPR_MINI', RPR_MINI),
        'injury_detail': ('RPR_INJURY', RPR_INJURY),
    }
    HDR_DEFAULTS = {
        'mini': ('RPR_MINI_HDR', RPR_MINI_HDR),
        'injury_detail': ('RPR_MINI_HDR', RPR_MINI_HDR),
    }

    for label, idx in table_types.items():
        hdr_rpr, data_rpr = _extract_table_rpr(dom, idx)
        pos = _extract_table_positioning(dom, idx)

        schema['tables'][label] = {
            'table_index': idx,
            'positioning': pos,
        }
        if data_rpr:
            schema['rpr'][f'{label}_data'] = data_rpr
        if hdr_rpr:
            schema['rpr'][f'{label}_hdr'] = hdr_rpr

        # RPR validation
        if label in RPR_DEFAULTS and data_rpr:
            const_name, const_val = RPR_DEFAULTS[label]
            match = (data_rpr.replace(' ', '').replace('\n', '') ==
                     const_val.replace(' ', '').replace('\n', ''))
            schema['rpr_validation'].append({
                'constant': const_name,
                'match': match,
                'live': data_rpr,
                'hardcoded': const_val,
            })
        if label in HDR_DEFAULTS and hdr_rpr:
            const_name, const_val = HDR_DEFAULTS[label]
            match = (hdr_rpr.replace(' ', '').replace('\n', '') ==
                     const_val.replace(' ', '').replace('\n', ''))
            schema['rpr_validation'].append({
                'constant': const_name + ' (hdr)',
                'match': match,
                'live': hdr_rpr,
                'hardcoded': const_val,
            })

    # Write schema.json
    out_path = os.path.join(work_path, 'schema.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(schema, f, indent=2, ensure_ascii=False)

    return schema


def _print_schema(schema):
    print("=== SCHEMA ===\n")

    print("── TABLE CLASSIFICATION ──")
    for label, info in schema['tables'].items():
        pos = info['positioning']
        pos_str = pos['type']
        if pos_str == 'floating':
            pos_str += f" (tblpX={pos.get('w:tblpX', '?')}, tblpY={pos.get('w:tblpY', '?')})"
        print(f"  {label:20} table #{info['table_index']:>2}  {pos_str}")

    print("\n── LIVE RPR ──")
    for key, val in schema['rpr'].items():
        print(f"  {key:20} {val}")

    print("\n── RPR VALIDATION ──")
    for v in schema['rpr_validation']:
        status = "MATCH" if v['match'] else "MISMATCH"
        print(f"  {v['constant']:20} {status}")
        if not v['match']:
            print(f"    live:      {v['live']}")
            print(f"    hardcoded: {v['hardcoded']}")


def main():
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <work_path>", file=sys.stderr)
        sys.exit(1)
    schema = extract_schema(sys.argv[1])
    _print_schema(schema)
    print(f"\nJSON: {sys.argv[1]}/schema.json")


if __name__ == '__main__':
    main()
