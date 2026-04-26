import streamlit as st
from datetime import datetime
import tempfile
import os
import pandas as pd

from scheduler import build_schedule, create_excel

# -----------------------
# PAGE CONFIG
# -----------------------
st.set_page_config(page_title="DC Scheduler", layout="wide", page_icon="🚀")

# -----------------------
# STYLING
# -----------------------
st.markdown("""
<style>
.main {background-color: #0f172a; color: white;}
h1, h2, h3 {color: #38bdf8;}
.stButton>button {border-radius: 10px; height: 3em; font-weight: 600;}
</style>
""", unsafe_allow_html=True)

st.title("🚀 Development Center Scheduler")

# -----------------------
# SESSION STATE
# -----------------------
if "data" not in st.session_state:
    st.session_state.data = {}

data = st.session_state.data

steps = ["⚙️ Setup", "👥 Names", "🧩 Tools", "📊 Review"]

# -----------------------
# SIDEBAR NAVIGATION (SIMPLE & STABLE)
# -----------------------
step = st.sidebar.radio("Navigation", steps)

# -----------------------
# STEP 1: SETUP
# -----------------------
if step == "⚙️ Setup":
    col1, col2, col3 = st.columns(3)

    with col1:
        data["candidates"] = st.number_input("Candidates", 1, 50, 6)
        data["assessors"] = st.number_input("Assessors", 2, 20, 2)

    with col2:
        data["start_time"] = st.text_input("Start Time", "09:00")
        data["end_time"] = st.text_input("End Time", "18:00")

    with col3:
        data["context_minutes"] = st.number_input("Context (mins)", 0, 120, 30)

    data["file_name"] = st.text_input("Output File Name", "DC_Schedule")

# -----------------------
# STEP 2: NAMES
# -----------------------
elif step == "👥 Names":
    use_names = st.toggle("Add Names", value=True)

    col1, col2 = st.columns(2)

    if use_names:
        with col1:
            data["assessor_names"] = [
                st.text_input(f"Assessor {i+1}", key=f"a{i}")
                for i in range(data.get("assessors", 2))
            ]

        with col2:
            data["participant_names"] = [
                st.text_input(f"Participant {i+1}", key=f"p{i}")
                for i in range(data.get("candidates", 6))
            ]
    else:
        data["assessor_names"] = [""] * data.get("assessors", 2)
        data["participant_names"] = [""] * data.get("candidates", 6)

# -----------------------
# STEP 3: TOOLS
# -----------------------
elif step == "🧩 Tools":
    tool_count = st.slider("Number of tools", 2, 10, 2)
    tools = []

    for i in range(int(tool_count)):
        st.markdown(f"### Tool {i+1}")

        tool_type = st.selectbox(
            "Type",
            ["Standard", "BEI", "Group Discussion", "Case Study", "Role Play", "Custom"],
            key=f"type{i}"
        )

        # Defaults
        if tool_type == "BEI":
            defaults = (0, 0, 60, 10)
        elif tool_type == "Group Discussion":
            defaults = (10, 60, 30, 10)
        elif tool_type == "Case Study":
            defaults = (10, 90, 30, 10)
        elif tool_type == "Role Play":
            defaults = (10, 30, 30, 10)
        else:
            defaults = (10, 15, 30, 10)

        # Initialize ONCE (prevents reset on navigation)
        if f"instr{i}" not in st.session_state:
            st.session_state[f"instr{i}"], st.session_state[f"prep{i}"], st.session_state[f"exec{i}"], st.session_state[f"score{i}"] = defaults

        # Detect actual dropdown change
        prev_type_key = f"prev_type{i}"

        if prev_type_key not in st.session_state:
            st.session_state[prev_type_key] = tool_type

        if st.session_state[prev_type_key] != tool_type:
            st.session_state[f"instr{i}"], st.session_state[f"prep{i}"], st.session_state[f"exec{i}"], st.session_state[f"score{i}"] = defaults
            st.session_state[prev_type_key] = tool_type

        name = st.text_input("Tool Name", key=f"name{i}")

        c1, c2, c3, c4 = st.columns(4)

        instruction = c1.number_input("Instruction", 0, 120, key=f"instr{i}")
        prep = c2.number_input("Preparation", 0, 120, key=f"prep{i}")
        execution = c3.number_input("Execution", 0, 180, key=f"exec{i}")
        scoring = c4.number_input("Scoring", 0, 60, key=f"score{i}")

        tools.append({
            "index": i,
            "name": name if name else f"Tool {i+1}",
            "instruction_slots": instruction // 5,
            "preparation_slots": prep // 5,
            "execution_slots": execution // 5,
            "scoring_slots": scoring // 5,
            "instruction_minutes": instruction,
            "preparation_minutes": prep,
            "execution_minutes": execution,
            "scoring_minutes": scoring,
        })

    data["tools"] = tools

# -----------------------
# STEP 4: REVIEW
# -----------------------
elif step == "📊 Review":
    st.subheader("Summary")
    st.write(data)

    try:
        start_dt = datetime.strptime(data["start_time"], "%H:%M")
        end_dt = datetime.strptime(data["end_time"], "%H:%M")

        inputs = {
            "candidates": data["candidates"],
            "assessors": data["assessors"],
            "start_time": start_dt,
            "end_time": end_dt,
            "start_minute": start_dt.hour * 60 + start_dt.minute,
            "end_minute": end_dt.hour * 60 + end_dt.minute,
            "slots_per_day": (end_dt.hour*60 + end_dt.minute - (start_dt.hour*60 + start_dt.minute)) // 5,
            "context_slots": data["context_minutes"] // 5,
            "context_minutes": data["context_minutes"],
            "tools": data["tools"],
            "assessor_names": data["assessor_names"],
            "participant_names": data["participant_names"],
        }

        col1, col2 = st.columns(2)

        with col1:
            if st.button("Preview Schedule"):
                result = build_schedule(inputs)

                preview_data = []
                for slot in range(result["max_slot"]):
                    minute = inputs["start_minute"] + slot * 5
                    time_label = f"{minute//60:02d}:{minute%60:02d}"

                    row = {"Time": time_label}

                    for c in range(1, data["candidates"] + 1):
                        row[f"P{c}"] = result["schedule"][c].get(slot, "")

                    preview_data.append(row)

                df = pd.DataFrame(preview_data)
                st.dataframe(df, use_container_width=True)

        with col2:
            if st.button("Generate Excel"):
                result = build_schedule(inputs)

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    path = tmp.name

                create_excel(inputs, result, path)

                with open(path, "rb") as f:
                    file = f.read()

                os.remove(path)

                file_name = data.get("file_name", "DC_Schedule")
                st.download_button("Download Excel", data=file, file_name=f"{file_name}.xlsx")

    except Exception as e:
        st.error(str(e))
