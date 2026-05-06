import streamlit as st
import math
from docx import Document
import io
import json

st.set_page_config(page_title="NI Ventilation Surveyor", layout="wide")

# --- DATA TABLES ---
bg_vent_table = {
    "50": [35000, 40000, 50000, 60000, 65000], "60": [35000, 40000, 50000, 60000, 65000],
    "70": [45000, 45000, 50000, 60000, 65000], "80": [50000, 50000, 50000, 60000, 65000],
    "90": [55000, 60000, 60000, 60000, 65000], "100": [65000, 65000, 65000, 65000, 65000]
}
whole_dwelling_flow_table = {1: 13, 2: 17, 3: 21, 4: 25, 5: 29}

# --- URL PERSISTENCE LOGIC ---
params = st.query_params

if 'floor_area' not in st.session_state:
    st.session_state.floor_area = int(params.get("area", 72))
if 'bedrooms' not in st.session_state:
    st.session_state.bedrooms = int(params.get("beds", 2))
if 'vents' not in st.session_state:
    url_vents = params.get("vdata", None)
    if url_vents:
        st.session_state.vents = json.loads(url_vents)
    else:
        st.session_state.vents = [{'size': 0, 'is_open': True}]

def update_url():
    v_json = json.dumps(st.session_state.vents)
    st.query_params.update(
        area=st.session_state.floor_area,
        beds=st.session_state.bedrooms,
        vdata=v_json
    )

def add_vent():
    st.session_state.vents.append({'size': 0, 'is_open': True})
    update_url()

def remove_vent(i):
    if len(st.session_state.vents) > 1:
        st.session_state.vents.pop(i)
        update_url()

# --- SIDEBAR ---
with st.sidebar:
    st.header("Property Details")
    # Step=1 and value as int removes the decimal points
    st.number_input("Total Floor Area (m²)", min_value=1, step=1, key="floor_area", on_change=update_url)
    st.selectbox("Number of Bedrooms", options=[1, 2, 3, 4, 5], key="bedrooms", on_change=update_url)
    st.divider()
    if st.button("Reset Survey"):
        st.session_state.vents = [{'size': 0, 'is_open': True}]
        st.session_state.floor_area = 72
        st.session_state.bedrooms = 2
        st.query_params.clear()
        st.rerun()

st.title("🪟 Ventilation Surveyor")

# --- 1. VENT SURVEY SECTION ---
st.write("### 1. Trickle Vent Survey")
total_available_area = 0.0
total_actual_area = 0.0

for i, vent in enumerate(st.session_state.vents):
    c1, c2, c3, c4 = st.columns([3, 2, 3, 1])
    
    # size is now an integer (no decimals)
    size = c1.number_input(f"Vent {i+1} Size", value=int(vent['size']), step=1, key=f"s_{i}", on_change=update_url)
    st.session_state.vents[i]['size'] = size
    
    is_open = c2.checkbox("Open?", value=vent['is_open'], key=f"o_{i}", on_change=update_url)
    st.session_state.vents[i]['is_open'] = is_open
    
    # Calculation remains: (Size * 10) * 0.8
    equiv = (size * 10) * 0.8
    total_available_area += equiv
    if is_open: total_actual_area += equiv
    
    c3.write(f"{equiv:,.0f} mm² ({'✅ Active' if is_open else '❌ Closed'})")
    if c4.button("🗑️", key=f"d_{i}"):
        remove_vent(i)
        st.rerun()

st.button("➕ Add Another Vent", on_click=add_vent)

# --- CALCULATIONS ---
f_area = st.session_state.floor_area
beds = st.session_state.bedrooms

key = "Over" if f_area > 100 else str(int(math.ceil(f_area / 10.0) * 10))
if f_area <= 50: key = "50"
if key == "Over":
    required_mm2 = 65000 + (math.ceil((f_area - 100) / 10) * 7000)
else:
    required_mm2 = bg_vent_table[key][beds - 1]

table_val = whole_dwelling_flow_table[beds]
calc_val = round(f_area * 0.3, 1)
final_val = math.ceil(max(table_val, calc_val))

# --- SUMMARY GENERATION ---
closed_count = sum(1 for v in st.session_state.vents if not v['is_open'])
summary_line_1 = f"Total available passive ventilation via trickle vents = {total_available_area:,.0f} mm²."
if closed_count > 0:
    summary_line_2 = f"At the time of survey, due to trickle vents being closed, the total amount of passive ventilation was {total_actual_area:,.0f} mm²."
else:
    summary_line_2 = f"At the time of survey, all trickle vents were open, providing {total_actual_area:,.0f} mm²."

status = "COMPLIANT" if total_actual_area >= required_mm2 else "NON-COMPLIANT"
comparison_comment = f"The required amount of passive ventilation is {required_mm2:,.0f} mm² versus what was actually occurring at the time of survey ({total_actual_area:,.0f} mm²)."
if total_actual_area < required_mm2 and closed_count > 0:
    comparison_comment += " Due to a number of Trickle vents being closed, which were therefore not allowing any passive ventilation into the property."

full_compliance_block = f"{summary_line_1} {summary_line_2} Status: {status}. {comparison_comment}"

# --- DISPLAY ASSESSMENT ---
st.divider()
st.header("Assessment Results")
st.write(f"**{summary_line_1}**")
if total_actual_area >= required_mm2:
    st.success(f"{summary_line_2} {comparison_comment}")
else:
    st.error(f"{summary_line_2} {comparison_comment}")

st.write(f"**Final Extraction Requirement: {final_val} L/s**")

# --- WORD DOCUMENT GENERATOR ---
st.header("📁 Generate Word Report")
uploaded_file = st.file_uploader("Upload Master Template", type="docx")

def docx_replace(doc, data):
    for p in doc.paragraphs:
        for k, v in data.items():
            if k in p.text:
                p.text = p.text.replace(k, v)

if uploaded_file is not None:
    if st.button("Generate & Download Report"):
        doc = Document(uploaded_file)
        replacements = {
            "{{AREA}}": str(f_area),
            "{{BEDROOMS}}": str(beds),
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
        st.download_button("💾 Download Word Report", data=buffer.getvalue(), file_name="Survey_Report.docx")
