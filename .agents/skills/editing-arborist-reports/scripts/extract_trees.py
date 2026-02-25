"""Extract tree data from arborist report Section 4 table.

Usage:
    python3 extract_trees.py <report.docx> [output.json]

If output.json is omitted, writes to <report_dir>/.work/tree_data.json.
Handles pandoc's fixed-width grid tables (space-aligned columns with --- separators).
"""

import sys, os, json, re, subprocess


def extract_trees(docx_path, output_path=None):
    """Convert docx to markdown via pandoc, parse Section 4 table and impact summary."""

    # Default output path: sibling .work directory
    if output_path is None:
        report_dir = os.path.dirname(os.path.abspath(docx_path))
        work_dir = os.path.join(report_dir, ".work")
        os.makedirs(work_dir, exist_ok=True)
        output_path = os.path.join(work_dir, "tree_data.json")

    # Convert to markdown
    result = subprocess.run(
        ["pandoc", docx_path, "-t", "markdown"],
        capture_output=True, text=True, encoding='utf-8'
    )
    if result.returncode != 0:
        print(f"pandoc error: {result.stderr}", file=sys.stderr)
        sys.exit(1)
    md = result.stdout

    trees = parse_tree_table(md)
    impacts = parse_impact_table(md)

    data = {"trees": trees}
    if impacts:
        data["impacts"] = impacts

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"Wrote {len(trees)} trees to {output_path}")
    if impacts:
        print(f"  + {len(impacts)} impact entries")

    return data


# ─── Grid table parser ──────────────────────────────────────────────────────

def _find_columns(sep_line):
    """Parse a --- separator line into column (start, end) positions.

    Example: "  -------- ------------- -------- -------"
    Returns: [(2,10), (11,24), (25,33), (34,41)]
    """
    cols = []
    i = 0
    while i < len(sep_line):
        if sep_line[i] == '-':
            start = i
            while i < len(sep_line) and sep_line[i] == '-':
                i += 1
            cols.append((start, i))
        else:
            i += 1
    return cols


def _extract_cell(line, start, end):
    """Extract cell text from a fixed-width line at given column positions."""
    if start >= len(line):
        return ''
    return line[start:min(end, len(line))].strip()


def _parse_grid_table(lines):
    """Parse a pandoc fixed-width grid table into header names and row dicts.

    Returns: (headers: list[str], rows: list[dict[str, str]])
    Each row maps header name -> cell value (multi-line cells joined with space).
    """
    # Find the inner separator (between header and data) — it's the second --- line
    sep_indices = []
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped and all(c in '-  ' for c in stripped) and '--' in stripped:
            sep_indices.append(i)

    if len(sep_indices) < 2:
        return [], []

    # Column positions from the inner separator (most reliable)
    cols = _find_columns(lines[sep_indices[1]])

    # Header: lines between first and second separator
    header_lines = lines[sep_indices[0]+1:sep_indices[1]]
    headers = []
    for col_start, col_end in cols:
        parts = []
        for hl in header_lines:
            cell = _extract_cell(hl, col_start, col_end)
            if cell:
                parts.append(cell)
        raw = ' '.join(parts)
        # Strip markdown bold markers and backslash escapes
        clean = re.sub(r'\*\*', '', raw)
        clean = clean.replace('\\#', '#').replace('\\*', '*').replace('\\', '')
        headers.append(clean.strip())

    # Data rows: lines between second separator and final separator
    # Rows are separated by blank lines (or the closing --- line)
    data_start = sep_indices[1] + 1
    data_end = sep_indices[2] if len(sep_indices) > 2 else len(lines)

    rows = []
    current_cells = None

    for line in lines[data_start:data_end]:
        stripped = line.strip()
        # Skip closing separator
        if stripped and all(c in '-  ' for c in stripped) and '--' in stripped:
            break

        # Check if this line starts a new row (first column has content)
        first_cell = _extract_cell(line, cols[0][0], cols[0][1]) if cols else ''

        if first_cell and first_cell != '':
            # Save previous row
            if current_cells is not None:
                rows.append(current_cells)
            # Start new row
            current_cells = {}
            for j, (cs, ce) in enumerate(cols):
                current_cells[headers[j]] = _extract_cell(line, cs, ce)
        elif current_cells is not None:
            # Continuation line — append to existing cells
            for j, (cs, ce) in enumerate(cols):
                extra = _extract_cell(line, cs, ce)
                if extra:
                    current_cells[headers[j]] += ' ' + extra

        # Blank line within data = end of a row's continuation lines
        if not stripped and current_cells is not None:
            rows.append(current_cells)
            current_cells = None

    # Don't forget last row
    if current_cells is not None:
        rows.append(current_cells)

    return headers, rows


def _find_section_lines(md, start_pat, end_pat):
    """Extract lines between start_pat and end_pat regex matches."""
    lines = md.splitlines()
    start = None
    for i, line in enumerate(lines):
        if start is None:
            if re.search(start_pat, line, re.IGNORECASE):
                start = i
        elif re.search(end_pat, line, re.IGNORECASE):
            return lines[start:i]
    if start is not None:
        return lines[start:]
    return []


# ─── Section 4 parser ───────────────────────────────────────────────────────

# Map from (lowercased) header substrings to output field names
_TREE_FIELD_MAP = {
    'tree': ('id', 'int'),
    'species': ('species', 'str'),
    'botanical': ('botanical', 'str'),
    'dbh': ('dbh_cm', 'num'),
    'condition': ('condition', 'str'),
    'comment': ('comments', 'str'),
    'ownership': ('ownership', 'str'),
    'direction': ('direction', 'str'),
    'tpz': ('tpz_m', 'num'),
    'encroach': ('encroachment_pct', 'num'),
    'permit': ('permit', 'bool'),
    'crown': ('crown_dia', 'num'),
}


def _map_header(header_text):
    """Map a header string to (field_name, type) or None."""
    h = header_text.lower()
    # 'species' in header but not 'botanical' -> species field
    if 'species' in h and 'botanical' not in h:
        return _TREE_FIELD_MAP['species']
    # Skip 'comment' headers — they contain noise words like "DBH"
    if 'comment' in h:
        return _TREE_FIELD_MAP['comment']
    for key, val in _TREE_FIELD_MAP.items():
        if key in h:
            if key == 'species' or key == 'comment':
                continue
            return val
    return None


def _coerce(val, type_name):
    """Coerce a string value to the given type."""
    val = val.strip()
    # Strip markdown italic/bold markers
    val = re.sub(r'[*_]+', '', val).strip()
    if not val or val == 'N/A':
        return None
    if type_name == 'int':
        m = re.search(r'\d+', val)
        return int(m.group()) if m else val
    if type_name == 'num':
        m = re.search(r'[\d.]+', val)
        return float(m.group()) if m else val
    if type_name == 'bool':
        return val.lower() in ('yes', 'y', 'true')
    return val


def parse_tree_table(md):
    """Parse Section 4 tree inventory grid table."""
    section = _find_section_lines(md, r'#.*Section\s*4', r'#.*Section\s*5')
    if not section:
        return []

    headers, rows = _parse_grid_table(section)
    if not rows:
        return []

    # Build field mapping: column header -> (output_field, type)
    field_map = {}  # header_name -> (field, type)
    for h in headers:
        mapped = _map_header(h)
        if mapped:
            field_map[h] = mapped

    trees = []
    for row in rows:
        tree = {}
        for h, (field, type_name) in field_map.items():
            if h in row:
                val = _coerce(row[h], type_name)
                if val is not None:
                    tree[field] = val
        if tree:
            trees.append(tree)
    return trees


# ─── Impact summary parser ──────────────────────────────────────────────────

def parse_impact_table(md):
    """Parse Section 5 impact summary table if present."""
    lines = md.splitlines()

    # Find lines around "Source of impact"
    source_idx = None
    for i, line in enumerate(lines):
        if re.search(r'source\s+of\s+impact', line, re.IGNORECASE):
            source_idx = i
            break
    if source_idx is None:
        return []

    # Grab surrounding table block — look back for opening --- separator
    table_start = source_idx
    for j in range(source_idx - 1, max(0, source_idx - 10), -1):
        stripped = lines[j].strip()
        if stripped and all(c in '-  ' for c in stripped) and '--' in stripped:
            table_start = j
            break

    # Find closing --- separator after the data
    table_end = len(lines)
    sep_count = 0
    for j in range(table_start, len(lines)):
        stripped = lines[j].strip()
        if stripped and all(c in '-  ' for c in stripped) and '--' in stripped:
            sep_count += 1
            if sep_count >= 3:  # opening + inner + closing
                table_end = j + 1
                break

    table_lines = lines[table_start:table_end]
    headers, rows = _parse_grid_table(table_lines)
    if not rows:
        return []

    # Map headers
    tree_h = source_h = dir_h = None
    for h in headers:
        hl = h.lower()
        if 'tree' in hl:
            tree_h = h
        elif 'source' in hl:
            source_h = h
        elif 'direction' in hl:
            dir_h = h

    impacts = []
    for row in rows:
        entry = {}
        if tree_h and tree_h in row:
            m = re.search(r'\d+', row[tree_h])
            if m:
                entry['tree_id'] = int(m.group())
        if source_h and source_h in row:
            val = row[source_h].strip()
            if val:
                entry['source'] = val
        if dir_h and dir_h in row:
            val = row[dir_h].strip()
            if val:
                entry['direction'] = val
        if entry:
            impacts.append(entry)
    return impacts


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <report.docx> [output.json]", file=sys.stderr)
        sys.exit(1)
    docx = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    extract_trees(docx, out)
