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

# UPDATED ROOM LIST
ROOM_OPTIONS = [
    "Lounge", "Living Room", "Dining Room", "Kitchen", "Utility room", 
    "Hallway", "Bathroom", "Bathroom 2", "Bathroom 3", "Ensuite", 
    "Bed 1", "Bed 2", "Bed 3", "Bed 4", "Bedroom 5"
]

# --- URL PERSISTENCE LOGIC ---
params = st.query_params

if 'address' not in st.session_state:
    st.session_state.address = params.get("addr", "")
if 'floor_area' not in st.session_state:
    st.session_state.floor_area = int(params.get("area", 72))
if 'bedrooms' not in st.session_state:
    st.session_state.bedrooms = int(params.get("beds", 2))
if 'vents' not in st.session_state:
    url_vents = params.get("vdata", None)
    if url_vents:
        st.session_state.vents = json.loads(url_vents)
    else:
        st.session_state.vents = [{'size': 0, 'is_open': True, 'room': 'Lounge'}]

def update_url():
    v_json = json.dumps(st.session_state.vents)
    st.query_params.update(
        addr=st.session_state.address,
        area=st.session_state.floor_area,
        beds=st.session_state.bedrooms,
        vdata=v_json
    )

def add_vent():
    st.session_state.vents.append({'size': 0, 'is_open': True, 'room': 'Lounge'})
    update_url()

def remove_vent(i):
    if len(st.session_state.vents) > 1:
        st.session_state.vents.pop(i)
        update_url()

# --- SIDEBAR ---
with st.sidebar:
    st.header("Property Details")
    st.text_input("Property Address", key="address", on_change=update_url)
    st.number_input("Total Floor Area (m²)", min_value=1, step=1, key="floor_area", on_change=update_url)
    st.selectbox("Number of Bedrooms", options=[1, 2, 3, 4, 5], key="bedrooms", on_change=update_url)
    st.divider()
    if st.button("Reset Survey"):
        st.session_state.vents = [{'size': 0, 'is_open': True, 'room': 'Lounge'}]
        st.session_state.floor_area = 72
        st.session_state.bedrooms = 2
        st.session_state.address = ""
        st.query_params.clear()
        st.rerun()

st.title("🪟 Ventilation Surveyor")

# --- 1. VENT SURVEY SECTION ---
st.write("### 1. Trickle Vent Survey")
total_available_area = 0.0
total_actual_area = 0.0

for i, vent in enumerate(st.session_state.vents):
    col_room, col_size, col_open, col_del = st.columns([3, 2, 2, 1])
    
    # Room Dropdown (with safety check for old room names)
    current_room = vent.get('room', 'Lounge')
    if current_room not in ROOM_OPTIONS: current_room = 'Lounge'
    
    room = col_room.selectbox(f"Room {i+1}", options=ROOM_OPTIONS, index=ROOM_OPTIONS.index(current_room), key=f"r_{i}", on_change=update_url)
    st.session_state.vents[i]['room'] = room

    # Size Input
    size = col_size.number_input(f"Size {i+1}", value=int(vent['size']), step=1, key=f"s_{i}", on_change=update_url)
    st.session_state.vents[i]['size'] = size
    
    # Checkbox
    is_open = col_open.checkbox("Open?", value=vent['is_open'], key=f"o_{i}", on_change=update_url)
    st.session_state.vents[i]['is_open'] = is_open
    
    equiv = (size * 10) * 0.8
    total_available_area += equiv
    if is_open: total_actual_area += equiv
    
    if col_del.button("🗑️", key=f"d_{i}"):
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

# --- SUMMARY & BREAKDOWN GENERATION ---
vent_list = []
for v in st.session_state.vents:
    v_status = "Open" if v['is_open'] else "Closed"
    v_area_mm = (v['size'] * 10) * 0.8
    vent_list.append(f"- {v['room']}: {v['size']}mm ({v_area_mm:,.0f} mm²) - Status: {v_status}")
vent_breakdown_string = "\n".join(vent_list)

closed_count = sum(1 for v in st.session_state.vents if not v['is_open'])
comparison_comment = f"The required amount of passive ventilation is {required_mm2:,.0f} mm² versus what was actually occurring at the time of survey ({total_actual_area:,.0f} mm²)."
if total_actual_area < required_mm2 and closed_count > 0:
    comparison_comment += " Due to a number of Trickle vents being closed, which were therefore not allowing any passive ventilation into the property."

full_compliance_block = f"Total available passive ventilation = {total_available_area:,.0f} mm². Status: {'COMPLIANT' if total_actual_area >= required_mm2 else 'NON-COMPLIANT'}. {comparison_comment}"

# --- DISPLAY RESULTS ---
st.divider()
st.header("Assessment Results")
st.write(f"**Address:** {st.session_state.address if st.session_state.address else 'Not Entered'}")
if total_actual_area >= required_mm2:
    st.success(full_compliance_block)
else:
    st.error(full_compliance_block)

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
            "{{ADDRESS}}": st.session_state.address,
            "{{AREA}}": str(f_area),
            "{{BEDROOMS}}": str(beds),
            "{{TABLE_VAL}}": str(table_val),
            "{{CALC_VAL}}": str(calc_val),
            "{{FINAL_VAL}}": str(final_val),
            "{{REQUIRED_MM2}}": f"{required_mm2:,.0f}",
            "{{COMPLIANCE_TEXT}}": full_compliance_block,
            "{{VENT_BREAKDOWN}}": vent_breakdown_string
        }
        docx_replace(doc, replacements)
        buffer = io.BytesIO()
        doc.save(buffer)
        st.download_button("💾 Download Word Report", data=buffer.getvalue(), file_name=f"Survey_{st.session_state.address}.docx")
