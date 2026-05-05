import math
import os
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st

from scheduler import create_excel, find_fastest_schedule


SLOT_MINUTES = 5


def minutes_to_slots(minutes):
    if minutes <= 0:
        return 0
    return int(math.ceil(minutes / SLOT_MINUTES))


def estimate_permutations(candidate_count, assessor_count, tool_count):
    group_count = math.ceil(candidate_count / assessor_count)
    tool_orders = math.factorial(tool_count)
    group_rotations = tool_count ** max(group_count - 1, 0)
    priority_orders = math.factorial(group_count) if group_count <= 4 else 2
    priority_modes = 2
    return tool_orders * group_rotations * priority_orders * priority_modes


# -----------------------
# PAGE CONFIG
# -----------------------
st.set_page_config(page_title="DC Scheduler", layout="wide", page_icon="🚀")

# -----------------------
# STYLING
# -----------------------
st.markdown(
    """
<style>
.main {background-color: #0f172a; color: white;}
h1, h2, h3 {color: #38bdf8;}
.stButton>button {border-radius: 10px; height: 3em; font-weight: 600;}
.card {
    padding: 15px;
    border-radius: 12px;
    background-color: #111827;
    margin-bottom: 10px;
}
.metric-card {
    padding: 18px;
    border-radius: 14px;
    background: #ffffff;
    border: 1px solid #e5e7eb;
    margin-bottom: 12px;
    color: #111827;
}

.metric-card h4,
.metric-card p {
    color: #111827;
}
</style>
""",
    unsafe_allow_html=True,
)

st.title("🚀 Development Center Scheduler")

# -----------------------
# SESSION STATE
# -----------------------
if "data" not in st.session_state:
    st.session_state.data = {}

if "tool_count" not in st.session_state:
    st.session_state.tool_count = 2

data = st.session_state.data

steps = ["⚙️ Setup", "🧩 Tools", "📊 Review"]

# -----------------------
# SIDEBAR NAVIGATION
# -----------------------
step = st.sidebar.radio("Navigation", steps)

# -----------------------
# STEP 1: SETUP
# -----------------------
if step == "⚙️ Setup":
    st.markdown("### ⚙️ Basic Configuration")
    st.info("""
Configure participants, assessors, and DC timing.

**General disclaimer: No confidential/personal data needs to be uploaded**
""")

    col1, col2, col3 = st.columns(3)

    with col1:
        data["candidates"] = st.number_input("Participants", 1, 50, 6)
        data["assessors"] = st.number_input("Assessors", 2, 20, 2)

    with col2:
        data["start_time"] = st.text_input("Start Time", "09:00")
        data["end_time"] = st.text_input("End Time", "20:00")

    with col3:
        data["context_minutes"] = st.number_input("Context Setting Time (mins)", 0, 120, 30)

    data["file_name"] = st.text_input("Output File Name", "DC_Schedule")

# -----------------------
# STEP 2: NAMES
# -----------------------
elif step == "👥 Names":
    st.markdown("### 👥 Add Participants & Assessors")
    st.info("Optional: Add names for better output clarity.")

    use_names = st.toggle("Add Names", value=False)

    col1, col2 = st.columns(2)

    if use_names:
        with col1:
            data["assessor_names"] = [
                st.text_input(f"Assessor {i + 1}", key=f"a{i}")
                for i in range(data.get("assessors", 2))
            ]

        with col2:
            data["participant_names"] = [
                st.text_input(f"Participant {i + 1}", key=f"p{i}")
                for i in range(data.get("candidates", 6))
            ]
    else:
        data["assessor_names"] = [""] * data.get("assessors", 2)
        data["participant_names"] = [""] * data.get("candidates", 6)

# -----------------------
# STEP 3: TOOLS
# -----------------------
elif step == "🧩 Tools":
    st.markdown("### 🧩 Configure Assessment Tools")
    st.info("Define each assessment tool with their respective timing and type.")

    tool_count = st.number_input(
        "Number of Assessment Tools",
        min_value=2,
        max_value=10,
        step=1,
        key="tool_count",
    )

    tools = []

    for i in range(int(tool_count)):
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(f"#### Assessment Tool {i + 1}")

        tool_type = st.selectbox(
            "Type",
            ["Standard", "BEI", "Group Discussion", "Case Study", "Role Play", "Custom"],
            key=f"type{i}",
        )

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

        if f"instr{i}" not in st.session_state:
            (
                st.session_state[f"instr{i}"],
                st.session_state[f"prep{i}"],
                st.session_state[f"exec{i}"],
                st.session_state[f"score{i}"],
            ) = defaults

        prev_type_key = f"prev_type{i}"

        if prev_type_key not in st.session_state:
            st.session_state[prev_type_key] = tool_type

        if st.session_state[prev_type_key] != tool_type:
            (
                st.session_state[f"instr{i}"],
                st.session_state[f"prep{i}"],
                st.session_state[f"exec{i}"],
                st.session_state[f"score{i}"],
            ) = defaults
            st.session_state[prev_type_key] = tool_type

        name = st.text_input("Exercise Name", key=f"name{i}")

        c1, c2, c3, c4 = st.columns(4)

        instruction = c1.number_input("Instruction", 0, 180, key=f"instr{i}")
        prep = c2.number_input("Preparation", 0, 240, key=f"prep{i}")
        execution = c3.number_input("Execution", 0, 240, key=f"exec{i}")
        scoring = c4.number_input("Scoring", 0, 120, key=f"score{i}")

        st.markdown("</div>", unsafe_allow_html=True)

        tools.append(
            {
                "index": i,
                "name": name if name else tool_type,
                "instruction_slots": minutes_to_slots(instruction),
                "preparation_slots": minutes_to_slots(prep),
                "execution_slots": minutes_to_slots(execution),
                "scoring_slots": minutes_to_slots(scoring),
                "instruction_minutes": minutes_to_slots(instruction) * SLOT_MINUTES,
                "preparation_minutes": minutes_to_slots(prep) * SLOT_MINUTES,
                "execution_minutes": minutes_to_slots(execution) * SLOT_MINUTES,
                "scoring_minutes": minutes_to_slots(scoring) * SLOT_MINUTES,
            }
        )

    data["tools"] = tools

# -----------------------
# STEP 4: REVIEW
# -----------------------
elif step == "📊 Review":
    st.markdown("### 📊 Review & Generate Schedule")

    try:
        start_dt = datetime.strptime(data["start_time"], "%H:%M")
        end_dt = datetime.strptime(data["end_time"], "%H:%M")

        if start_dt >= end_dt:
            st.error("End time must be after start time.")
            st.stop()

        if not data.get("tools"):
            st.warning("Please add at least one exercise.")
            st.stop()

        start_minute = start_dt.hour * 60 + start_dt.minute
        end_minute = end_dt.hour * 60 + end_dt.minute

        inputs = {
            "candidates": data["candidates"],
            "assessors": data["assessors"],
            "start_time": start_dt,
            "end_time": end_dt,
            "start_minute": start_minute,
            "end_minute": end_minute,
            "slots_per_day": (end_minute - start_minute) // SLOT_MINUTES,
            "context_slots": minutes_to_slots(data["context_minutes"]),
            "context_minutes": minutes_to_slots(data["context_minutes"]) * SLOT_MINUTES,
            "tools": data["tools"],
            "assessor_names": data.get("assessor_names", [""] * data["assessors"]),
            "participant_names": data.get("participant_names", [""] * data["candidates"]),
        }

        permutation_count = estimate_permutations(
            data["candidates"],
            data["assessors"],
            len(data["tools"]),
        )

        st.markdown(
            f"""
<div class="metric-card">
    <h4>Optimization Scope</h4>
    <p><b>{permutation_count:,}</b> schedule combinations will be checked for this setup.</p>
</div>
""",
            unsafe_allow_html=True,
        )

        with st.expander("Review input data"):
            st.write(data)

        col1, col2 = st.columns(2)

        with col1:
            if st.button("Preview Schedule"):
                with st.status("Optimizing schedule...", expanded=True) as status:
                    st.write(f"Checking {permutation_count:,} combinations...")
                    best_inputs, result = find_fastest_schedule(inputs)
                    status.update(label="Optimization complete", state="complete", expanded=False)

                preview_data = []
                include_day = result.get("total_days", 1) > 1

                for slot in range(result["max_slot"]):
                    local_slot = slot % best_inputs["slots_per_day"]
                    day = slot // best_inputs["slots_per_day"]
                    minute = best_inputs["start_minute"] + local_slot * SLOT_MINUTES
                    time_label = f"{minute // 60:02d}:{minute % 60:02d}"

                    if include_day:
                        time_label = f"Day {day + 1} {time_label}"

                    row = {"Time": time_label}

                    for c in range(1, data["candidates"] + 1):
                        row[f"P{c}"] = result["schedule"][c].get(slot, "")

                    preview_data.append(row)

                df = pd.DataFrame(preview_data)
                st.dataframe(df, use_container_width=True)

        with col2:
            if st.button("Generate Excel"):
                with st.status("Optimizing and creating Excel...", expanded=True) as status:
                    st.write(f"Checking {permutation_count:,} combinations...")
                    best_inputs, result = find_fastest_schedule(inputs)
                    st.write("Best schedule selected.")
                    st.write("Building Excel workbook...")

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        path = tmp.name

                    create_excel(best_inputs, result, path)
                    status.update(label="Excel ready", state="complete", expanded=False)

                with open(path, "rb") as f:
                    file = f.read()

                os.remove(path)

                file_name = data.get("file_name", "DC_Schedule")
                st.success(f"Done. Checked {permutation_count:,} combinations.")
                st.download_button(
                    "Download Excel",
                    data=file,
                    file_name=f"{file_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    except Exception as e:
        st.error(str(e))
