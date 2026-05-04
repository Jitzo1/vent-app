import streamlit as st
import math
from docx import Document
import io

st.set_page_config(page_title="NI Ventilation Surveyor", layout="wide")

# --- DATA TABLES ---
# Background Vent Table (Table 1)
bg_vent_table = {
    "50": [35000, 40000, 50000, 60000, 65000], "60": [35000, 40000, 50000, 60000, 65000],
    "70": [45000, 45000, 50000, 60000, 65000], "80": [50000, 50000, 50000, 60000, 65000],
    "90": [55000, 60000, 60000, 60000, 65000], "100": [65000, 65000, 65000, 65000, 65000]
}
# Whole Dwelling Table (Table 2.2)
whole_dwelling_flow_table = {1: 13, 2: 17, 3: 21, 4: 25, 5: 29}

# --- STATE MANAGEMENT ---
if 'vents' not in st.session_state:
    st.session_state.vents = [{'size': 0.0, 'is_open': True}]

def add_vent(): st.session_state.vents.append({'size': 0.0, 'is_open': True})
def remove_vent(i): 
    if len(st.session_state.vents) > 1: st.session_state.vents.pop(i)

# --- SIDEBAR ---
with st.sidebar:
    st.header("Property Details")
    floor_area = st.number_input("Total Floor Area (m²)", min_value=1.0, value=72.0, step=0.1)
    bedrooms = st.selectbox("Number of Bedrooms", options=[1, 2, 3, 4, 5])
    st.info("Passive Logic: (Size × 10) × 0.8")

st.title("🪟 Ventilation Surveyor & Report Generator")

# --- 1. VENT SURVEY SECTION ---
st.write("### 1. Trickle Vent Survey")
total_available_area = 0.0
total_actual_area = 0.0

for i, vent in enumerate(st.session_state.vents):
    c1, c2, c3, c4 = st.columns([3, 2, 3, 1])
    size = c1.number_input(f"Vent {i+1} Size", value=vent['size'], key=f"s_{i}")
    st.session_state.vents[i]['size'] = size
    is_open = c2.checkbox("Open?", value=vent['is_open'], key=f"o_{i}")
    st.session_state.vents[i]['is_open'] = is_open
    
    equiv = (size * 10) * 0.8
    total_available_area += equiv
    if is_open: total_actual_area += equiv
    
    status_text = "✅ Active" if is_open else "❌ Closed"
    c3.write(f"{equiv:,.0f} mm² ({status_text})")
    if c4.button("🗑️", key=f"d_{i}"):
        remove_vent(i)
        st.rerun()

st.button("➕ Add Another Vent", on_click=add_vent)

# --- 2. CALCULATIONS ---
# Passive Background Required (Table 1)
key = "Over" if floor_area > 100 else str(int(math.ceil(floor_area / 10.0) * 10))
if floor_area <= 50: key = "50"
if key == "Over":
    required_mm2 = 65000 + (math.ceil((floor_area - 100) / 10) * 7000)
else:
    required_mm2 = bg_vent_table[key][bedrooms - 1]

# Flow Rate Requirement (Table 2.2)
table_val = whole_dwelling_flow_table[bedrooms]
calc_val = round(floor_area * 0.3, 1)
final_val = math.ceil(max(table_val, calc_val))

# --- 3. GENERATE SURVEY COMMENTS ---
closed_count = sum(1 for v in st.session_state.vents if not v['is_open'])

# Line 1: Maximum Available
summary_line_1 = f"Total available passive ventilation via trickle vents = {total_available_area:,.0f} mm²."

# Line 2: Actual at time of survey
if closed_count > 0:
    summary_line_2 = f"At the time of survey, due to trickle vents being closed, the total amount of passive ventilation was {total_actual_area:,.0f} mm²."
else:
    summary_line_2 = f"At the time of survey, all trickle vents were open, providing {total_actual_area:,.0f} mm²."

# Line 3: Final Comparison Comment
status = "COMPLIANT" if total_actual_area >= required_mm2 else "NON-COMPLIANT"
comparison_comment = f"The required amount of passive ventilation is {required_mm2:,.0f} mm² versus what was actually occurring at the time of survey ({total_actual_area:,.0f} mm²)."

if total_actual_area < required_mm2 and closed_count > 0:
    comparison_comment += " Due to a number of Trickle vents being closed, which were therefore not allowing any passive ventilation into the property."

# Combine into a single block for the Word Doc
full_compliance_block = f"{summary_line_1} {summary_line_2} Status: {status}. {comparison_comment}"

# --- DISPLAY ASSESSMENT ---
st.divider()
st.header("Assessment Results")
st.write(f"**{summary_line_1}**")
if total_actual_area >= required_mm2:
    st.success(f"{summary_line_2} {comparison_comment}")
else:
    st.error(f"{summary_line_2} {comparison_comment}")

with st.expander("Show Flow Rate Logic"):
    st.write(f"- Table 2.2 Requirement: {table_val} L/s")
    st.write(f"- Area Calculation ({floor_area} x 0.3): {calc_val} L/s")
    st.write(f"- **Final PIV Setting: {final_val} L/s**")

# --- 4. WORD DOCUMENT GENERATOR ---
st.header("📁 Generate Word Report")
uploaded_file = st.file_uploader("Upload Master Template", type="docx")

def docx_replace(doc, data):
    for p in doc.paragraphs:
        for key, value in data.items():
            if key in p.text:
                p.text = p.text.replace(key, value)

if uploaded_file is not None:
    if st.button("Generate & Download"):
        doc = Document(uploaded_file)
        replacements = {
            "{{AREA}}": str(floor_area),
            "{{BEDROOMS}}": str(bedrooms),
            "{{TABLE_VAL}}": str(table_val),
            "{{CALC_VAL}}": str(calc_val),
            "{{FINAL_VAL}}": str(final_val),
            "{{TOTAL_AVAILABLE}}": f"{total_available_area:,.0f}",
            "{{ACTUAL_VENT}}": f"{total_actual_area:,.0f}",
            "{{REQUIRED_MM2}}": f"{required_mm2:,.0f}",
            "{{COMPLIANCE_TEXT}}": full_compliance_block
        }
        docx_replace(doc, replacements)
        buffer = io.BytesIO()
        doc.save(buffer)
        st.download_button("💾 Download Report", data=buffer.getvalue(), file_name="Survey_Report.docx")