import sys, os
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, r"C:\Users\User\.claude\plugins\cache\anthropic-agent-skills\document-skills\69c0b1a06741\skills\docx")
from scripts.document import Document

PROJECT = r"C:\Projects\Arborism\71 Lloyd Manor, Toronto, ON, Canada"
doc = Document(PROJECT + r"\.work\unpacked\temp", author="Claude", rsid="8F693E63")
editor = doc["word/document.xml"]

# ============================================================
# EDIT 1: Fill in the first empty row of the injury detail table
# Columns: Injury source | Closest point of impact | Max depth | Impact to condition
# ============================================================

header_node = editor.get_node(tag="w:r", contains="Injury source")
header_tr = header_node.parentNode
while header_tr.tagName != "w:tr":
    header_tr = header_tr.parentNode

data_row_1 = header_tr.nextSibling
while data_row_1 and data_row_1.nodeType != 1:
    data_row_1 = data_row_1.nextSibling

cells = [n for n in data_row_1.childNodes if n.nodeType == 1 and n.tagName == "w:tc"]

rpr = '<w:rPr><w:rFonts w:cs="Helvetica"/><w:szCs w:val="22"/></w:rPr>'

cell_values = ["Driveway", "2.5m", '6"', "Moderate"]
for i, val in enumerate(cell_values):
    p_nodes = [n for n in cells[i].childNodes if n.nodeType == 1 and n.tagName == "w:p"]
    if p_nodes:
        empty_p = p_nodes[0]
        # Wrap run in <w:ins> for tracked changes
        run_xml = f'<w:ins><w:r>{rpr}<w:t>{val}</w:t></w:r></w:ins>'
        editor.append_to(empty_p, run_xml)

print("Edit 1 done: Injury detail table first row filled (tracked)")

# ============================================================
# EDIT 2: Insert Tree 2 mini data table after "Injuries" heading
# ============================================================

injuries_node = editor.get_node(tag="w:r", contains="Injuries")
injuries_p = injuries_node.parentNode
while injuries_p.tagName != "w:p":
    injuries_p = injuries_p.parentNode

next_p = injuries_p.nextSibling
while next_p and (next_p.nodeType != 1 or next_p.tagName != "w:p"):
    next_p = next_p.nextSibling

def hdr_cell(text, width):
    return f'''<w:tc>
  <w:tcPr>
    <w:tcW w:w="{width}" w:type="dxa"/>
    <w:tcBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tcBorders>
    <w:vAlign w:val="center"/>
  </w:tcPr>
  <w:p>
    <w:pPr><w:jc w:val="center"/></w:pPr>
    <w:ins><w:r>
      <w:rPr>
        <w:rFonts w:cs="Helvetica"/>
        <w:b/>
        <w:bCs/>
        <w:color w:val="000000"/>
        <w:sz w:val="16"/>
        <w:szCs w:val="16"/>
      </w:rPr>
      <w:t>{text}</w:t>
    </w:r></w:ins>
  </w:p>
</w:tc>'''

def data_cell(text, width):
    return f'''<w:tc>
  <w:tcPr>
    <w:tcW w:w="{width}" w:type="dxa"/>
    <w:tcBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tcBorders>
    <w:shd w:val="clear" w:color="000000" w:fill="F2F2F2"/>
    <w:vAlign w:val="center"/>
  </w:tcPr>
  <w:p>
    <w:pPr><w:jc w:val="center"/></w:pPr>
    <w:ins><w:r>
      <w:rPr>
        <w:rFonts w:cs="Helvetica"/>
        <w:color w:val="000000"/>
        <w:sz w:val="16"/>
        <w:szCs w:val="16"/>
      </w:rPr>
      <w:t>{text}</w:t>
    </w:r></w:ins>
  </w:p>
</w:tc>'''

cols = [
    ("TREE #", "2", 662),
    ("Species", "Norway Maple", 1200),
    ("DBH (cm)", "91", 700),
    ("Condition Rating", "Good", 1000),
    ("Ownership", "Private", 900),
    ("Direction", "Injury", 900),
    ("TPZ (m)", "6", 700),
    ("% Encroachment", "20%", 900),
    ("Permit", "Yes", 700),
]

header_cells = "".join(hdr_cell(h, w) for h, _, w in cols)
data_cells = "".join(data_cell(v, w) for _, v, w in cols)

mini_table_xml = f'''<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="5000" w:type="pct"/>
    <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
  </w:tblPr>
  <w:tblGrid>
    {"".join(f'<w:gridCol w:w="{w}"/>' for _, _, w in cols)}
  </w:tblGrid>
  <w:tr>
    <w:trPr><w:trHeight w:val="400"/><w:ins/></w:trPr>
    {header_cells}
  </w:tr>
  <w:tr>
    <w:trPr><w:trHeight w:val="400"/><w:ins/></w:trPr>
    {data_cells}
  </w:tr>
</w:tbl>'''

editor.insert_after(next_p, mini_table_xml)

print("Edit 2 done: Tree 2 mini data table inserted (tracked)")

# ============================================================
# EDIT 3: Insert narrative paragraphs after the injury detail table
# ============================================================

rse_node = editor.get_node(tag="w:r", contains="Excavation should not be deeper")
rse_p = rse_node.parentNode
while rse_p.tagName != "w:p":
    rse_p = rse_p.parentNode

body_rpr = '<w:rPr><w:rFonts w:cs="Helvetica"/><w:lang w:val="en-US"/></w:rPr>'

para1 = f'''<w:p>
  <w:pPr>
    <w:spacing w:after="160" w:line="276" w:lineRule="auto"/>
    <w:rPr><w:rFonts w:cs="Helvetica"/><w:lang w:val="en-US"/></w:rPr>
  </w:pPr>
  <w:ins><w:r>{body_rpr}<w:t xml:space="preserve">Tree #2 is a 91cm mature Norway Maple in private ownership. The subject tree is in good condition. Good crown spread and vitality were noted. These attributes are not of immediate concern.</w:t></w:r></w:ins>
</w:p>'''

para2 = f'''<w:p>
  <w:pPr>
    <w:spacing w:after="160" w:line="276" w:lineRule="auto"/>
    <w:rPr><w:rFonts w:cs="Helvetica"/><w:lang w:val="en-US"/></w:rPr>
  </w:pPr>
  <w:ins><w:r>{body_rpr}<w:t xml:space="preserve">The subject tree will require moderate TPZ encroachment and root injuries to allow for the construction of the proposed driveway. Footprints of the driveway are located approximately 2.5m away from the trunk at the closest point, necessitating 3.5m encroachment on its 6.0m TPZ.</w:t></w:r></w:ins>
</w:p>'''

para3 = f'''<w:p>
  <w:pPr>
    <w:spacing w:after="160" w:line="276" w:lineRule="auto"/>
    <w:rPr><w:rFonts w:cs="Helvetica"/><w:lang w:val="en-US"/></w:rPr>
  </w:pPr>
  <w:ins><w:r>{body_rpr}<w:t xml:space="preserve">The driveway will be constructed as a non-permeable surface requiring 6 inches excavation to allow for a 3 inch gravel base and 2-3 inch asphalt layer. Excavation deeper than 6 inches is not allowed.</w:t></w:r></w:ins>
</w:p>'''

para4 = f'''<w:p>
  <w:pPr>
    <w:spacing w:after="160" w:line="276" w:lineRule="auto"/>
    <w:rPr><w:rFonts w:cs="Helvetica"/><w:lang w:val="en-US"/></w:rPr>
  </w:pPr>
  <w:ins><w:r>{body_rpr}<w:t xml:space="preserve">Footprints of the proposal overlap the TPZ by approximately 20%, and at this distance, it is likely that minor to moderate amounts of small and fibrous roots will be discovered. Roots larger than 5cm must be retained. If uncovered, the envelope should be backfilled and the footprint reduced. Pruning minor amounts of small roots is expected to have no effect on structural integrity.</w:t></w:r></w:ins>
</w:p>'''

para5 = f'''<w:p>
  <w:pPr>
    <w:spacing w:after="160" w:line="276" w:lineRule="auto"/>
    <w:rPr><w:rFonts w:cs="Helvetica"/><w:lang w:val="en-US"/></w:rPr>
  </w:pPr>
  <w:ins><w:r>{body_rpr}<w:t xml:space="preserve">Provided the mitigation measures are followed, the proposed injury is within the species' tolerance levels and the tree is expected to remain viable.</w:t></w:r></w:ins>
</w:p>'''

all_narrative = para1 + para2 + para3 + para4 + para5
editor.insert_before(rse_p, all_narrative)

print("Edit 3 done: Narrative paragraphs inserted (tracked)")

# ============================================================
# SAVE
# ============================================================
doc.save()
print("All edits saved successfully")
