import json
import math
import os
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st

from scheduler import create_excel, find_fastest_schedule


SLOT_MINUTES = 5
PROJECTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dc_projects.json")

TOOL_TYPES = ["Standard", "BEI", "Group Discussion", "Case Study", "Role Play", "Custom"]
TYPE_DEFAULTS = {
    "BEI": (0, 0, 60, 10),
    "Group Discussion": (10, 60, 30, 10),
    "Case Study": (10, 90, 30, 10),
    "Role Play": (10, 30, 30, 10),
}
DEFAULT_TYPE_TIMES = (10, 15, 30, 10)


def minutes_to_slots(minutes):
    if minutes <= 0:
        return 0
    return int(math.ceil(minutes / SLOT_MINUTES))


def parse_time(text, label):
    try:
        return datetime.strptime(str(text).strip(), "%H:%M")
    except ValueError:
        raise ValueError(f"{label} must be in HH:MM format (e.g. 09:30).")


def minute_of(dt):
    return dt.hour * 60 + dt.minute


def clock(minute):
    minute %= 24 * 60
    return f"{minute // 60:02d}:{minute % 60:02d}"


def estimate_permutations(candidate_count, assessor_count, tool_count):
    group_count = math.ceil(candidate_count / assessor_count)
    tool_orders = math.factorial(tool_count)
    group_rotations = tool_count ** max(group_count - 1, 0)
    priority_orders = math.factorial(group_count) if group_count <= 4 else 2
    priority_modes = 2
    return tool_orders * group_rotations * priority_orders * priority_modes


# -----------------------
# PROJECT PERSISTENCE
# -----------------------
def load_projects():
    try:
        with open(PROJECTS_FILE, "r", encoding="utf-8") as handle:
            return json.load(handle)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_projects(projects):
    with open(PROJECTS_FILE, "w", encoding="utf-8") as handle:
        json.dump(projects, handle, indent=2)


def current_config():
    ss = st.session_state
    tools = []
    for i in range(int(ss.get("tool_count", 2))):
        tools.append(
            {
                "type": ss.get(f"type{i}", "Standard"),
                "name": ss.get(f"name{i}", ""),
                "instruction": ss.get(f"instr{i}", 10),
                "preparation": ss.get(f"prep{i}", 15),
                "execution": ss.get(f"exec{i}", 30),
                "scoring": ss.get(f"score{i}", 10),
            }
        )
    return {
        "participants": ss.get("participants", 6),
        "assessors": ss.get("assessors", 2),
        "start_time": ss.get("start_time", "09:30"),
        "end_time": ss.get("end_time", "20:00"),
        "context_minutes": ss.get("context_minutes", 30),
        "file_name": ss.get("file_name", "DC_Schedule"),
        "lunch_start": ss.get("lunch_start", "12:00"),
        "lunch_end": ss.get("lunch_end", "16:00"),
        "integration_on": ss.get("integration_on", False),
        "integration_minutes": ss.get("integration_minutes", 60),
        "secondary_tools": ss.get("secondary_tools", []),
        "tool_count": int(ss.get("tool_count", 2)),
        "tools": tools,
    }


def apply_config(config):
    ss = st.session_state
    for key in (
        "participants", "assessors", "start_time", "end_time", "context_minutes",
        "file_name", "lunch_start", "lunch_end", "integration_on", "integration_minutes",
    ):
        if key in config:
            ss[key] = config[key]

    ss["secondary_tools"] = config.get("secondary_tools", [])
    ss["tool_count"] = int(config.get("tool_count", len(config.get("tools", [])) or 2))

    for i, tool in enumerate(config.get("tools", [])):
        ss[f"type{i}"] = tool.get("type", "Standard")
        ss[f"name{i}"] = tool.get("name", "")
        ss[f"instr{i}"] = tool.get("instruction", 10)
        ss[f"prep{i}"] = tool.get("preparation", 15)
        ss[f"exec{i}"] = tool.get("execution", 30)
        ss[f"score{i}"] = tool.get("scoring", 10)
        ss[f"prev_type{i}"] = tool.get("type", "Standard")


# -----------------------
# PAGE CONFIG + THEME
# -----------------------
st.set_page_config(page_title="DC Scheduler", layout="wide", page_icon="🗓️")

st.markdown(
    """
<style>
:root { --accent:#38bdf8; --accent2:#818cf8; --ink:#e2e8f0; --panel:#111a2e; }
.stApp { background: radial-gradient(1200px 600px at 15% -10%, #16233f 0%, #0b1120 55%); }
.block-container { padding-top: 2rem; }
h1, h2, h3 { color: var(--accent); letter-spacing: .2px; }
.hero {
    background: linear-gradient(120deg, rgba(56,189,248,.14), rgba(129,140,248,.14));
    border: 1px solid rgba(56,189,248,.25); border-radius: 18px;
    padding: 20px 24px; margin-bottom: 18px;
}
.hero h1 { margin: 0; font-size: 1.9rem; }
.hero p { margin: 6px 0 0; color: #9fb3c8; }
.section {
    background: var(--panel); border: 1px solid rgba(148,163,184,.14);
    border-radius: 16px; padding: 18px 20px; margin-bottom: 16px;
}
.section h3 { margin-top: 0; }
.stButton>button {
    border-radius: 12px; height: 3em; font-weight: 700; border: 0;
    background: linear-gradient(120deg, var(--accent), var(--accent2)); color: #041018;
}
.stButton>button:hover { filter: brightness(1.08); }
.stDownloadButton>button {
    border-radius: 12px; height: 3em; font-weight: 700;
    background: linear-gradient(120deg,#34d399,#10b981); color:#04120c; border:0;
}
.card {
    padding: 14px 16px; border-radius: 14px; background: #0e1729;
    border: 1px solid rgba(148,163,184,.16); margin-bottom: 12px;
}
.timeline {
    display:flex; gap:10px; flex-wrap:wrap; margin-top:6px;
}
.chip {
    background:#0e1729; border:1px solid rgba(56,189,248,.35); color:#cfe8fb;
    border-radius:999px; padding:6px 14px; font-size:.86rem;
}
.metric-card {
    padding: 16px 18px; border-radius: 14px; background:#ffffff;
    border:1px solid #e5e7eb; color:#0f172a;
}
.metric-card h4, .metric-card p { color:#0f172a; margin:.2rem 0; }
</style>
""",
    unsafe_allow_html=True,
)

# Apply a pending project load BEFORE any widget renders.
if "_pending_project" in st.session_state:
    apply_config(st.session_state.pop("_pending_project"))

for key, default in (
    ("participants", 6), ("assessors", 2), ("start_time", "09:30"), ("end_time", "20:00"),
    ("context_minutes", 30), ("file_name", "DC_Schedule"),
    ("lunch_start", "12:00"), ("lunch_end", "16:00"),
    ("integration_on", False), ("integration_minutes", 60),
    ("tool_count", 2), ("secondary_tools", []),
):
    st.session_state.setdefault(key, default)

st.markdown(
    '<div class="hero"><h1>🗓️ Development Center Scheduler</h1>'
    '<p>Build assessor & participant day plans in three steps — Setup, Tools, Review.</p></div>',
    unsafe_allow_html=True,
)

# -----------------------
# SIDEBAR: NAVIGATION + PROJECTS
# -----------------------
steps = ["1 · Setup", "2 · Tools", "3 · Review"]
step = st.sidebar.radio("Steps", steps, label_visibility="collapsed")

st.sidebar.markdown("---")
st.sidebar.markdown("### 💾 Projects")
st.sidebar.caption("Save a design once, reload it later, change only the numbers.")

projects = load_projects()

st.sidebar.text_input("Project name", key="project_name", placeholder="e.g. BPCL Senior DC")
if st.sidebar.button("💾 Save project", use_container_width=True):
    name = (st.session_state.get("project_name") or "").strip()
    if not name:
        st.sidebar.error("Enter a project name first.")
    else:
        projects[name] = current_config()
        save_projects(projects)
        st.sidebar.success(f"Saved '{name}'.")

if projects:
    chosen = st.sidebar.selectbox("Load a saved project", ["—"] + sorted(projects.keys()))
    col_load, col_del = st.sidebar.columns(2)
    if col_load.button("Load", use_container_width=True) and chosen != "—":
        st.session_state["_pending_project"] = projects[chosen]
        st.rerun()
    if col_del.button("Delete", use_container_width=True) and chosen != "—":
        projects.pop(chosen, None)
        save_projects(projects)
        st.rerun()
else:
    st.sidebar.caption("No saved projects yet.")

# -----------------------
# STEP 1: SETUP
# -----------------------
if step == "1 · Setup":
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown("### ⚙️ Basic configuration")
    st.caption("No confidential or personal data is needed here.")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.number_input("Participants", 1, 50, key="participants")
        st.number_input("Assessors", 1, 20, key="assessors")
    with col2:
        st.text_input("Start time", key="start_time", help="When the DC day begins, HH:MM.")
        st.text_input("End time (before integration)", key="end_time",
                      help="When the ASSESSMENTS end. If integration is on, it runs after this — "
                           "End time is not the end of the day.")
    with col3:
        st.number_input("Context setting (mins)", 0, 120, key="context_minutes")
        st.text_input("Output file name", key="file_name")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown("### 🍴 Lunch window")
    st.caption("A 30-minute lunch is placed inside this window. Keep it tight (e.g. 12:00–12:30) to "
               "send everyone to lunch together, or widen it (e.g. 12:00–13:30) to let groups "
               "stagger — one preps while others eat.")
    lc1, lc2 = st.columns(2)
    with lc1:
        st.text_input("Window start", key="lunch_start")
    with lc2:
        st.text_input("Window end", key="lunch_end")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown("### 🤝 Integration")
    st.caption("An all-hands session that runs straight after the assessments — its minutes are added "
               "after End time, on the same day.")
    ic1, ic2 = st.columns([1.3, 1])
    with ic1:
        st.toggle("Add an integration session at the end", key="integration_on")
    with ic2:
        st.number_input("Duration (mins)", 0, 240, step=5, key="integration_minutes")
    st.markdown("</div>", unsafe_allow_html=True)

    # Live day timeline hint
    try:
        s = minute_of(parse_time(st.session_state.start_time, "Start"))
        e = minute_of(parse_time(st.session_state.end_time, "End"))
        integ = int(st.session_state.integration_minutes) if st.session_state.integration_on else 0
        chips = [f"Day starts {clock(s)}", f"Assessments end {clock(e)}"]
        if integ:
            chips.append(f"Integration {clock(e)}–{clock(e + integ)}")
            chips.append(f"Day ends {clock(e + integ)}")
        else:
            chips.append(f"Day ends {clock(e)}")
        st.markdown('<div class="timeline">' + "".join(f'<span class="chip">{c}</span>' for c in chips)
                    + "</div>", unsafe_allow_html=True)
    except ValueError:
        pass

# -----------------------
# STEP 2: TOOLS
# -----------------------
elif step == "2 · Tools":
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown("### 🧩 Assessment tools")
    st.caption("Define each tool's type and phase timings. Picking a type pre-fills typical timings — "
               "you can override any of them.")
    st.number_input("Number of tools", min_value=1, max_value=10, step=1, key="tool_count")
    st.markdown("</div>", unsafe_allow_html=True)

    tool_names = []
    for i in range(int(st.session_state.tool_count)):
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(f"**Tool {i + 1}**")

        top = st.columns([1, 1.4])
        tool_type = top[0].selectbox("Type", TOOL_TYPES, key=f"type{i}")
        defaults = TYPE_DEFAULTS.get(tool_type, DEFAULT_TYPE_TIMES)

        if f"instr{i}" not in st.session_state:
            (st.session_state[f"instr{i}"], st.session_state[f"prep{i}"],
             st.session_state[f"exec{i}"], st.session_state[f"score{i}"]) = defaults

        prev_key = f"prev_type{i}"
        if st.session_state.get(prev_key) != tool_type:
            if prev_key in st.session_state:
                (st.session_state[f"instr{i}"], st.session_state[f"prep{i}"],
                 st.session_state[f"exec{i}"], st.session_state[f"score{i}"]) = defaults
            st.session_state[prev_key] = tool_type

        top[1].text_input("Exercise name", key=f"name{i}", placeholder=tool_type)

        c1, c2, c3, c4 = st.columns(4)
        c1.number_input("Instruction", 0, 180, key=f"instr{i}")
        c2.number_input("Preparation", 0, 240, key=f"prep{i}")
        c3.number_input("Execution", 0, 240, key=f"exec{i}")
        c4.number_input("Scoring", 0, 120, key=f"score{i}")
        st.markdown("</div>", unsafe_allow_html=True)

        tool_names.append(st.session_state.get(f"name{i}") or tool_type)

    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown("### 👤 Assessor mapping — optional")
    st.caption("Leave empty for automatic mapping. To set it by hand, pick the tools the SECONDARY "
               "assessor runs; everything else goes to the primary. The rule of thumb is that the "
               "primary does more tools.")
    st.session_state["secondary_tools"] = [
        name for name in st.session_state.get("secondary_tools", []) if name in tool_names
    ]
    st.multiselect("Tools run by the secondary assessor", options=tool_names, key="secondary_tools")
    st.markdown("</div>", unsafe_allow_html=True)

# -----------------------
# STEP 3: REVIEW
# -----------------------
elif step == "3 · Review":
    st.markdown("### 📊 Review & generate")

    try:
        start_dt = parse_time(st.session_state.start_time, "Start time")
        end_dt = parse_time(st.session_state.end_time, "End time")
        if start_dt >= end_dt:
            st.error("End time must be after start time.")
            st.stop()

        lunch_start_dt = parse_time(st.session_state.lunch_start, "Lunch window start")
        lunch_end_dt = parse_time(st.session_state.lunch_end, "Lunch window end")
        if minute_of(lunch_end_dt) - minute_of(lunch_start_dt) < 30:
            st.error("The lunch window must be at least 30 minutes wide.")
            st.stop()

        tool_count = int(st.session_state.tool_count)
        if tool_count < 1:
            st.warning("Add at least one tool first.")
            st.stop()

        tools = []
        for i in range(tool_count):
            name = st.session_state.get(f"name{i}") or st.session_state.get(f"type{i}", "Standard")
            instr = st.session_state.get(f"instr{i}", 0)
            prep = st.session_state.get(f"prep{i}", 0)
            exe = st.session_state.get(f"exec{i}", 0)
            score = st.session_state.get(f"score{i}", 0)
            tools.append({
                "index": i,
                "name": name,
                "instruction_slots": minutes_to_slots(instr),
                "preparation_slots": minutes_to_slots(prep),
                "execution_slots": minutes_to_slots(exe),
                "scoring_slots": minutes_to_slots(score),
                "instruction_minutes": minutes_to_slots(instr) * SLOT_MINUTES,
                "preparation_minutes": minutes_to_slots(prep) * SLOT_MINUTES,
                "execution_minutes": minutes_to_slots(exe) * SLOT_MINUTES,
                "scoring_minutes": minutes_to_slots(score) * SLOT_MINUTES,
            })

        candidates = int(st.session_state.participants)
        assessors = int(st.session_state.assessors)
        start_minute = minute_of(start_dt)
        end_minute = minute_of(end_dt)  # assessments end here
        integration_minutes = (
            int(st.session_state.integration_minutes) if st.session_state.integration_on else 0
        )
        secondary_tools = [
            t["name"] for t in tools if t["name"] in st.session_state.get("secondary_tools", [])
        ]

        assessment_slots = (end_minute - start_minute) // SLOT_MINUTES
        full_day_slots = (end_minute + integration_minutes - start_minute) // SLOT_MINUTES

        inputs = {
            "candidates": candidates,
            "assessors": assessors,
            "start_time": start_dt,
            "end_time": end_dt,
            "start_minute": start_minute,
            "end_minute": end_minute,
            # Full day includes integration (for rendering); tools are capped at End time.
            "slots_per_day": full_day_slots,
            "assessment_slots_per_day": assessment_slots,
            "context_slots": minutes_to_slots(st.session_state.context_minutes),
            "context_minutes": minutes_to_slots(st.session_state.context_minutes) * SLOT_MINUTES,
            "tools": tools,
            "assessor_names": [""] * assessors,
            "participant_names": [""] * candidates,
            "lunch_window_start_minute": minute_of(lunch_start_dt),
            "lunch_window_end_minute": minute_of(lunch_end_dt),
            "integration_minutes": integration_minutes,
            "secondary_tools": secondary_tools,
        }

        permutation_count = estimate_permutations(candidates, assessors, tool_count)
        mapping_note = (
            "Manual — secondary: " + ", ".join(secondary_tools) if secondary_tools else "Automatic"
        )

        m1, m2, m3 = st.columns(3)
        m1.metric("Combinations checked", f"{permutation_count:,}")
        m2.metric("Assessments end", clock(end_minute))
        m3.metric("Integration", f"{integration_minutes} min" if integration_minutes else "Off")

        st.markdown(
            f'<div class="card">Lunch window <b>{st.session_state.lunch_start}–{st.session_state.lunch_end}</b>'
            f' &nbsp;•&nbsp; Mapping: <b>{mapping_note}</b>'
            f' &nbsp;•&nbsp; Participants <b>{candidates}</b>, Assessors <b>{assessors}</b>,'
            f' Tools <b>{tool_count}</b></div>',
            unsafe_allow_html=True,
        )

        with st.expander("Inspect the raw input the engine will receive"):
            st.write(inputs)

        col1, col2 = st.columns(2)

        with col1:
            if st.button("👁️ Preview schedule", use_container_width=True):
                with st.status("Optimizing…", expanded=True) as status:
                    st.write(f"Checking {permutation_count:,} combinations…")
                    best_inputs, result = find_fastest_schedule(inputs)
                    status.update(label="Optimization complete", state="complete", expanded=False)

                include_day = result.get("total_days", 1) > 1
                rows = []
                for slot in range(result["max_slot"]):
                    local = slot % best_inputs["slots_per_day"]
                    day = slot // best_inputs["slots_per_day"]
                    minute = best_inputs["start_minute"] + local * SLOT_MINUTES
                    label = clock(minute)
                    if include_day:
                        label = f"Day {day + 1} {label}"
                    row = {"Time": label}
                    for c in range(1, candidates + 1):
                        row[f"P{c}"] = result["schedule"][c].get(slot, "")
                    rows.append(row)
                st.dataframe(pd.DataFrame(rows), use_container_width=True, height=460)

        with col2:
            if st.button("📥 Generate Excel", use_container_width=True):
                with st.status("Optimizing and building the workbook…", expanded=True) as status:
                    st.write(f"Checking {permutation_count:,} combinations…")
                    best_inputs, result = find_fastest_schedule(inputs)
                    st.write("Best schedule selected. Writing sheets…")
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        path = tmp.name
                    create_excel(best_inputs, result, path)
                    status.update(label="Excel ready", state="complete", expanded=False)

                with open(path, "rb") as handle:
                    file_bytes = handle.read()
                os.remove(path)

                file_name = st.session_state.get("file_name", "DC_Schedule") or "DC_Schedule"
                st.success(f"Done — {result['total_days']} day(s), checked {permutation_count:,} combinations.")
                st.download_button(
                    "⬇️ Download Excel",
                    data=file_bytes,
                    file_name=f"{file_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    except ValueError as error:
        st.error(str(error))
    except Exception as error:  # surface engine errors to the user
        st.error(str(error))
