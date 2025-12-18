import streamlit as st
import re, random
from io import BytesIO
from copy import deepcopy
from collections import defaultdict
from docx import Document
from docx.shared import Inches
from docx.table import _Cell
import zipfile

st.set_page_config(page_title="EndSem QB Generator (With Bloom's Level)", layout="wide")

# -----------------------
# Regex & XML Helpers
# -----------------------
CO_RE = re.compile(r'CO\s*[_:]?\s*(\d+)', re.I)
# Updated Bloom's regex to capture K levels (K1-K6)
BLOOM_RE = re.compile(r'K\s*([1-6])', re.I)
NUM_PREFIX_RE = re.compile(r'^\s*(\d+)\s*[\.\)]')

def _copy_element(elem):
    return deepcopy(elem)

def replace_cell_with_cell(target_cell, src_cell):
    t_tc = target_cell._tc
    s_tc = src_cell._tc
    for child in list(t_tc):
        t_tc.remove(child)
    for child in list(s_tc):
        t_tc.append(_copy_element(child))

def extract_images_from_cell(cell):
    images = []
    for blip in cell._element.xpath(".//*[local-name()='blip']"):
        rId = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        if rId:
            part = cell.part.related_parts[rId]
            images.append(BytesIO(part.blob))
    return images

def find_or_pairs(slots):
    pairs = []
    for i, s in enumerate(slots):
        if s.get("is_or"):
            left = right = None
            for j in range(i-1, -1, -1):
                if slots[j].get("slot_num"):
                    left = slots[j]
                    break
            for k in range(i+1, len(slots)):
                if slots[k].get("slot_num"):
                    right = slots[k]
                    break
            if left and right:
                pairs.append((left, right))
    return pairs

# -----------------------
# Mapping & Extraction
# -----------------------
def get_tag_for_slot(slot_num, part_c_unit):
    if 1 <= slot_num <= 2: return "1A"
    if 3 <= slot_num <= 4: return "2A"
    if 5 <= slot_num <= 6: return "3A"
    if 7 <= slot_num <= 8: return "4A"
    if 9 <= slot_num <= 10: return "5A"
    if 11 <= slot_num <= 12: return "1B"
    if 13 <= slot_num <= 14: return "2B"
    if 15 <= slot_num <= 16: return "3B"
    if 17 <= slot_num <= 18: return "4B"
    if 19 <= slot_num <= 20: return "5B"
    if 21 <= slot_num <= 22: return f"{part_c_unit}C"
    return None

def extract_questions_from_bank_docx(uploaded_file):
    doc = Document(BytesIO(uploaded_file.read()))
    questions = []
    for table in doc.tables:
        for row in table.rows:
            try:
                cells_text = [c.text.strip() for c in row.cells]
            except: continue
            if not any(cells_text): continue
            
            tag = None
            for txt in cells_text:
                if re.match(r'^[1-5][ABC]$', txt.upper()):
                    tag = txt.upper()
                    break
            if not tag: continue

            cm = CO_RE.search(" ".join(cells_text))
            # Extract Bloom's Level from the row
            bm = BLOOM_RE.search(" ".join(cells_text))
            
            main_idx = max(range(len(row.cells)), key=lambda i: len(row.cells[i].text))
            
            questions.append({
                "id": random.random(),
                "tag": tag,
                "co": cm.group(1) if cm else "",
                "bloom": f"K{bm.group(1)}" if bm else "", # Store the K level
                "cell": row.cells[main_idx], 
                "images": extract_images_from_cell(row.cells[main_idx])
            })
    return questions

# -----------------------
# Assembly Logic
# -----------------------
def assemble_doc(template_file_bytes, selected_map):
    doc = Document(BytesIO(template_file_bytes))
    for (ti, ri, ci), q in selected_map.items():
        table = doc.tables[ti]
        tr = table.rows[ri]._tr
        cells_xml = tr.tc_lst
        
        # Mapping: Col 0: No., Col 1: Question, Col 2: CO, Col 3: Bloom's Level
        q_cell = _Cell(cells_xml[1], table)
        co_cell = _Cell(cells_xml[2], table)
        
        q_cell.text = ""
        co_cell.text = ""

        if q == "TAG_NOT_FOUND":
            q_cell.text = "Error: Tag not found in QB"
        else:
            replace_cell_with_cell(q_cell, q["cell"])
            # Clear drawing errors
            for p in q_cell.paragraphs:
                for run in list(p.runs):
                    if run._element.find('.//w:drawing', namespaces={'w':"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}) is not None:
                        p._element.remove(run._element)
            # Re-insert images
            for img in q.get("images", []):
                img.seek(0)
                q_cell.add_paragraph().add_run().add_picture(img, width=Inches(2.5))
            
            co_cell.text = f"CO{q.get('co','')}"
            
            # THE FIX: Fill the 4th column (Bloom's Level) with the K value
            if len(cells_xml) > 3:
                k_cell = _Cell(cells_xml[3], table)
                k_cell.text = q.get('bloom', '')

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# -----------------------
# Standard Parsing & UI
# -----------------------
def parse_template_slots(template_file_bytes):
    doc = Document(BytesIO(template_file_bytes))
    slots = []
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            tr = row._tr
            for ci, tc in enumerate(tr.tc_lst):
                cell = _Cell(tc, table)
                txt = cell.text.strip()
                if txt.upper() in ["OR", "(OR)", "( OR )"]:
                    slots.append({"table_index": ti, "row_index": ri, "cell_index": ci, "slot_num": None, "is_or": True})
                else:
                    m = NUM_PREFIX_RE.match(txt)
                    if m:
                        slots.append({"table_index": ti, "row_index": ri, "cell_index": ci, "slot_num": int(m.group(1)), "is_or": False})
    return slots

def select_questions_by_tag(entries, or_pairs, questions, part_c_unit):
    selected = {}
    used_q_ids = set() 
    by_tag = defaultdict(list)
    for q in questions:
        by_tag[q["tag"]].append(q)

    def pick_for_slot(slot_num):
        target_tag = get_tag_for_slot(slot_num, part_c_unit)
        pool = by_tag.get(target_tag, [])
        available = [q for q in pool if q["id"] not in used_q_ids]
        if not available: return "TAG_NOT_FOUND" if not pool else random.choice(pool)
        chosen_q = random.choice(available)
        used_q_ids.add(chosen_q["id"])
        return chosen_q

    for left, right in or_pairs:
        selected[(left["table_index"], left["row_index"], left["cell_index"])] = pick_for_slot(left["slot_num"])
        selected[(right["table_index"], right["row_index"], right["cell_index"])] = pick_for_slot(right["slot_num"])

    for e in entries:
        coord = (e["table_index"], e["row_index"], e["cell_index"])
        if coord not in selected:
            selected[coord] = pick_for_slot(e["slot_num"])

    return selected

st.title("📑 QP Generator (Including K-Levels)")
with st.sidebar:
    t_file = st.file_uploader("Template", type="docx")
    b_file = st.file_uploader("Question Bank", type="docx")
    n_sets = st.number_input("Sets", 1, 10, 1)
    part_c_unit = st.selectbox("Part C Unit", [1, 2, 3, 4, 5])

if st.button("Generate") and t_file and b_file:
    qb_questions = extract_questions_from_bank_docx(b_file)
    template_bytes = t_file.read()
    slots = parse_template_slots(template_bytes)
    entries = [s for s in slots if s.get("slot_num")]
    or_pairs = find_or_pairs(slots)
    
    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w") as z:
        for i in range(n_sets):
            selected = select_questions_by_tag(entries, or_pairs, qb_questions, part_c_unit)
            final_doc = assemble_doc(template_bytes, selected)
            z.writestr(f"Set_{i+1}.docx", final_doc.getvalue())
            st.download_button(f"Download Set {i+1}", final_doc, f"Set_{i+1}.docx")
    
    zip_buf.seek(0)
    st.download_button("Download All (ZIP)", zip_buf, "All_Sets.zip")
