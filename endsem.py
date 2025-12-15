import streamlit as st
import re, random
from io import BytesIO
from copy import deepcopy
from collections import defaultdict
from docx import Document
from docx.shared import Inches
import zipfile
from lxml import etree 

st.set_page_config(page_title="EndSem QB Generator (Stable Multi-Set)", layout="wide")

# -----------------------
# Regex helpers
# -----------------------
BLOOM_RE = re.compile(r'K\s*([1-6])', re.I)
CO_RE = re.compile(r'CO\s*[_:]?\s*(\d+)', re.I)
UNIT_RE = re.compile(r'Unit\s*[-:]?\s*(\d+)', re.I)
NUM_PREFIX_RE = re.compile(r'^\s*(\d+)\s*[\.\)]')

# -----------------------
# XML helpers
# -----------------------
def _copy_element(elem):
    return deepcopy(elem)

def replace_cell_with_cell(target_cell, src_cell):
    """Copy entire DOCX cell XML (text + equations)"""
    t_tc = target_cell._tc
    s_tc = src_cell._tc
    for child in list(t_tc):
        t_tc.remove(child)
    for child in list(s_tc):
        t_tc.append(_copy_element(child))

def extract_images_from_cell(cell):
    """Extract images as binary blobs"""
    images = []
    for blip in cell._element.xpath(".//*[local-name()='blip']"):
        rId = blip.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
        )
        if rId:
            part = cell.part.related_parts[rId]
            images.append(BytesIO(part.blob))
    return images

# -----------------------
# Parse Question Bank
# -----------------------
def extract_questions_from_bank_docx(uploaded_file):
    doc = Document(BytesIO(uploaded_file.read()))
    questions = []
    qid = 0

    for table in doc.tables:
        for row in table.rows:
            texts = [c.text.strip() for c in row.cells]
            if all(not t for t in texts):
                continue

            joined = " ".join(texts)
            bm = BLOOM_RE.search(joined)
            if not bm:
                continue

            um = UNIT_RE.search(joined)
            cm = CO_RE.search(joined)

            main_idx = max(range(len(row.cells)), key=lambda i: len(row.cells[i].text))
            main_cell = row.cells[main_idx]

            qid += 1
            questions.append({
                "id": qid,
                "unit": int(um.group(1)) if um else None,
                "co": int(cm.group(1)) if cm else None,
                "bloom": int(bm.group(1)),
                "cell": main_cell,  # FULL CELL (equation-safe)
                "images": extract_images_from_cell(main_cell)
            })

    return questions

# -----------------------
# Parse Template
# -----------------------
def parse_template_slots(template_file):
    doc = Document(template_file)
    slots = []

    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                txt = cell.text.strip()
                if not txt:
                    continue

                if txt.upper() in ["OR", "(OR)", "( OR )"]:
                    slots.append({
                        "table_index": ti,
                        "row_index": ri,
                        "cell_index": ci,
                        "slot_num": None,
                        "is_or": True
                    })
                    continue

                m = NUM_PREFIX_RE.match(txt)
                if m:
                    um = UNIT_RE.search(txt)
                    slots.append({
                        "table_index": ti,
                        "row_index": ri,
                        "cell_index": ci,
                        "slot_num": int(m.group(1)),
                        "is_or": False,
                        "unit_in_template": int(um.group(1)) if um else None
                    })
    return slots

# -----------------------
# Logic helpers
# -----------------------
def allowed_blooms_for_slot_num(slot_num):
    if 1 <= slot_num <= 10:
        return [1, 2, 3]
    if 11 <= slot_num <= 20:
        return [4, 5]
    if 21 <= slot_num <= 22:
        return [6]
    return [1, 2, 3, 4, 5, 6]

def map_slots(slots):
    return sorted([s for s in slots if s.get("slot_num")], key=lambda x: x["slot_num"])

def find_or_pairs(slots):
    pairs = []
    for i, s in enumerate(slots):
        if s["is_or"]:
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
# Select Questions (Cleaned: No streamli warnings)
# -----------------------
def select_questions(entries, or_pairs, questions):
    selected = {}
    used_q_ids = set() # Track questions used in THIS SET

    by_unit_bloom = defaultdict(lambda: defaultdict(list))
    by_bloom = defaultdict(list)

    for q in questions:
        by_bloom[q["bloom"]].append(q)
        if q["unit"]:
            by_unit_bloom[q["unit"]][q["bloom"]].append(q)
            
    # Helper to select a unique question
    def get_unique_question(pool):
        # 1. Filter by uniqueness check
        available = [q for q in pool if q["id"] not in used_q_ids]
        
        if not available:
            # Fallback: if no unique question is available, we have to repeat
            # NOTE: st.warning() is REMOVED here to prevent messages
            chosen_q = random.choice(pool) if pool else None
        else:
            chosen_q = random.choice(available)
        
        if chosen_q:
            used_q_ids.add(chosen_q["id"]) # Mark as used
        return chosen_q

    # Handle OR pairs first (where 2 questions are chosen)
    for left, right in or_pairs:
        lb = allowed_blooms_for_slot_num(left["slot_num"])
        rb = allowed_blooms_for_slot_num(right["slot_num"])
        unit = left.get("unit_in_template")

        # Create pools based on Bloom/Unit criteria
        lp_all = [q for b in lb for q in by_unit_bloom.get(unit, {}).get(b, [])] or \
                 [q for b in lb for q in by_bloom[b]]

        rp_all = [q for b in rb for q in by_unit_bloom.get(unit, {}).get(b, [])] or \
                 [q for b in rb for q in by_bloom[b]]

        # Select the 'left' question
        q_left = get_unique_question(lp_all)
        if q_left:
            selected[(left["table_index"], left["row_index"], left["cell_index"])] = q_left

        # Select the 'right' question (must be unique from 'left' and all others)
        q_right = get_unique_question(rp_all)
        if q_right:
            selected[(right["table_index"], right["row_index"], right["cell_index"])] = q_right
            
    # Handle single entries
    for e in entries:
        coord = (e["table_index"], e["row_index"], e["cell_index"])
        
        # Skip if already handled as part of an OR pair
        if coord in selected:
            continue

        blooms = allowed_blooms_for_slot_num(e["slot_num"])
        unit = e.get("unit_in_template")

        pool_all = [q for b in blooms for q in by_unit_bloom.get(unit, {}).get(b, [])] or \
                   [q for b in blooms for q in by_bloom[b]]

        chosen_q = get_unique_question(pool_all)
        if chosen_q:
            selected[coord] = chosen_q

    return selected, used_q_ids

# -----------------------
# Assemble DOCX (OMML & Image Safe Hybrid)
# -----------------------
def assemble_doc(template_file, selected_map):
    doc = Document(template_file)

    for (ti, ri, ci), q in selected_map.items():
        row = doc.tables[ti].rows[ri]
        if len(row.cells) < 4:
            continue

        q_cell = row.cells[1]
        co_cell = row.cells[2]
        k_cell = row.cells[3]

        # Clear existing content in target cells
        q_cell.text = ""
        co_cell.text = ""
        k_cell.text = ""

        # 1️⃣ Copy text and OMML Equations
        replace_cell_with_cell(q_cell, q["cell"])

        # 2️⃣ Remove broken/duplicate image XML from the target cell
        for p in q_cell.paragraphs:
            for run in list(p.runs):
                # Search for 'w:drawing' element within the run's XML
                drawing = run._element.find('.//w:drawing', namespaces={'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                if drawing is not None:
                    # Remove the run element containing the drawing
                    p._element.remove(run._element)
        
        # 3️⃣ Re-insert images explicitly
        for img in q.get("images", []):
            img.seek(0)
            q_cell.add_paragraph().add_run().add_picture(img, width=Inches(2.5))
        
        # 4️⃣ Add CO and K values
        co_cell.add_paragraph(f"CO{q.get('co','')}")
        k_cell.add_paragraph(f"K{q.get('bloom','')}")

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# -----------------------
# Streamlit UI (Cleaned: No repetition warnings)
# -----------------------
st.title("📘 End-Sem Question Paper Generator (OMML & Image Safe)")

template = st.file_uploader("Upload Template DOCX", type="docx")
bank = st.file_uploader("Upload Question Bank DOCX", type="docx")
n_sets = st.number_input("Number of Sets", 1, 10, 2)

if st.button("Generate Question Papers"):
    if not template or not bank:
        st.error("Upload both template and question bank")
        st.stop()
    if n_sets < 2:
        # st.warning("Generate at least 2 sets to calculate repetition percentage.") # REMOVED
        pass
        
    # --- Generation & Tracking Logic ---
    questions = extract_questions_from_bank_docx(bank)
    st.success(f"✅ {len(questions)} questions loaded from question bank.")

    slots = parse_template_slots(template)
    entries = map_slots(slots)
    or_pairs = find_or_pairs(slots)
    
    if not entries:
        st.warning("No question slots (like '1.', '2.', '3.') found in the template.")
        st.stop()
        
    buffers = {}
    all_used_q_ids = []
    
    # Check if enough unique questions exist for one set
    questions_required_per_set = len(entries) + len(or_pairs)
    if len(questions) < questions_required_per_set:
        st.error(f"❌ Error: Your bank has {len(questions)} questions, but {questions_required_per_set} are needed per set. Repetition is UNAVOIDABLE for uniqueness in Set 1.")


    for i in range(n_sets):
        selected, used_q_ids_in_set = select_questions(entries, or_pairs, questions)
        
        # Store used IDs for inter-set analysis
        all_used_q_ids.extend(list(used_q_ids_in_set))

        # Reset template file pointer for each set generation
        template.seek(0) 
        buf = assemble_doc(template, selected)
        buffers[f"Set_{i+1}.docx"] = buf

        st.download_button(
            f"📄 Download Set {i+1}",
            buf,
            file_name=f"Set_{i+1}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    # --- Repetition Analysis ---
    if n_sets >= 2 and all_used_q_ids:
        total_questions_generated = len(all_used_q_ids)
        unique_questions_used = len(set(all_used_q_ids))
        
        # Repetitions = Total generated questions - Unique questions used
        total_repetitions = total_questions_generated - unique_questions_used
        
        # Percentage of Repetition = (Total Repetitions / Total Questions Generated) * 100
        repetition_percentage = (total_repetitions / total_questions_generated) * 100
        
        st.markdown("---")
        st.header("📊 Repetition Analysis (Across Sets)")
        st.info(f"""
        * **Total Questions Generated:** {total_questions_generated}
        * **Unique Questions Used:** {unique_questions_used}
        * **Total Question Repetitions:** {total_repetitions}
        
        ### **Inter-Set Repetition Percentage: {repetition_percentage:.2f}%**
        
        *(This percentage represents how often a question from the bank was reused across all generated sets.)*
        """)


    # --- Zip Download ---
    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w") as z:
        for name, buf in buffers.items():
            z.writestr(name, buf.getvalue())
    zip_buf.seek(0)

    st.download_button("🗂 Download All Sets (ZIP)", zip_buf, "All_Sets.zip")