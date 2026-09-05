import copy
import json
import math
import os
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st

from scheduler import (
    ALLOWED_SLOT_MINUTES,
    WIDENED_SEARCH_LIMIT,
    create_excel,
    find_fastest_schedule,
    group_start_spread,
    set_slot_minutes,
)


DEFAULT_SLOT_MINUTES = 5
PROJECTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dc_projects.json")

TOOL_TYPES = ["Standard", "BEI", "Group Discussion", "Case Study", "Role Play", "Custom"]
TYPE_DEFAULTS = {
    "BEI": (0, 0, 60, 10),
    "Group Discussion": (10, 60, 30, 10),
    "Case Study": (10, 90, 30, 10),
    "Role Play": (10, 30, 30, 10),
}
DEFAULT_TYPE_TIMES = (10, 15, 30, 10)

STEPS = ["1 · Setup", "2 · Tools", "3 · Review"]


def default_tool():
    return {"type": "Standard", "name": "", "instruction": 10,
            "preparation": 15, "execution": 30, "scoring": 10}


def default_config():
    return {
        "participants": 6,
        "assessors": 2,
        "start_time": "09:30",
        "end_time": "18:00",
        "context_minutes": 30,
        "file_name": "DC_Schedule",
        "lunch_start": "12:00",
        "lunch_end": "16:00",
        "integration_on": False,
        "integration_minutes": 60,
        "slot_minutes": DEFAULT_SLOT_MINUTES,
        "tool_count": 2,
        "tools": [default_tool(), default_tool()],
        "secondary_tools": [],
    }


def minutes_to_slots(minutes, slot_minutes=DEFAULT_SLOT_MINUTES):
    if minutes <= 0:
        return 0
    return int(math.ceil(minutes / slot_minutes))


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
    priority_orders = math.factorial(group_count) if group_count <= 4 else 2
    priority_modes = 2

    # The engine searches every group's rotation when that stays quick, and
    # falls back to holding the first group fixed when it would not. Each
    # structure is also tried with two lunch strategies, and the strongest few
    # then get a lunch-placement refinement pass on top.
    lunch_strategies = 2
    priority_count = priority_orders * priority_modes

    # Each group can take its own tool order when that is affordable; otherwise
    # the engine falls back to rotations of one shared order.
    independent = (tool_orders ** group_count) * priority_count
    if independent <= WIDENED_SEARCH_LIMIT:
        return independent * lunch_strategies

    widened = tool_orders * (tool_count ** group_count) * priority_count
    if widened <= WIDENED_SEARCH_LIMIT:
        return widened * lunch_strategies
    return (tool_orders * (tool_count ** max(group_count - 1, 0))
            * priority_count * lunch_strategies)


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


def clear_widget_state():
    """Drop the transient widget keys so they re-seed from cfg (used on Load)."""
    for key in list(st.session_state.keys()):
        if key == "cfg" or key == "nav" or key == "project_name":
            continue
        del st.session_state[key]


# -----------------------
# SMALL WIDGET HELPERS
# cfg is the single source of truth. Each helper seeds its widget key from cfg
# once, then writes the widget value back to cfg every run. Because cfg is a
# plain (non-widget) session_state entry, it survives page navigation.
# -----------------------
def num(label, key, lo, hi, cfg, field, step=1, help=None):
    st.session_state.setdefault(key, int(cfg[field]))
    value = st.number_input(label, lo, hi, step=step, key=key, help=help)
    cfg[field] = value
    return value


def txt(label, key, cfg, field, help=None, placeholder=None):
    st.session_state.setdefault(key, str(cfg[field]))
    value = st.text_input(label, key=key, help=help, placeholder=placeholder)
    cfg[field] = value
    return value


def tog(label, key, cfg, field):
    st.session_state.setdefault(key, bool(cfg[field]))
    value = st.toggle(label, key=key)
    cfg[field] = value
    return value


# -----------------------
# PAGE CONFIG + THEME
# -----------------------
st.set_page_config(page_title="DC Scheduler", layout="wide", page_icon="🗓️")

st.markdown(
    """
<style>
:root { --accent:#0284c7; --accent2:#4f46e5; --ink:#0f172a; --muted:#5b6b7f; }
.stApp { background:#f5f7fb; }
.block-container { padding-top: 2rem; }
h1, h2, h3 { color: var(--ink); letter-spacing:.2px; }
.hero {
    background: linear-gradient(120deg, #e0f2fe, #e8e7fd);
    border: 1px solid #cfe0f5; border-radius:18px; padding:20px 24px; margin-bottom:18px;
}
.hero h1 { margin:0; font-size:1.9rem; color:#0b3a5b; }
.hero p { margin:6px 0 0; color:#41607a; }
.section {
    background:#ffffff; border:1px solid #e3e8ef;
    border-radius:16px; padding:18px 20px; margin-bottom:16px;
    box-shadow: 0 1px 2px rgba(16,24,40,.04);
}
.section h3 { margin-top:0; color:#12304a; }
.stButton>button {
    border-radius:12px; height:3em; font-weight:700; border:0;
    background: linear-gradient(120deg, var(--accent), var(--accent2)); color:#ffffff;
}
.stButton>button:hover { filter:brightness(1.06); color:#ffffff; }
.stDownloadButton>button {
    border-radius:12px; height:3em; font-weight:700;
    background: linear-gradient(120deg,#059669,#047857); color:#ffffff; border:0;
}
.card {
    padding:14px 16px; border-radius:14px; background:#ffffff;
    border:1px solid #e3e8ef; margin-bottom:12px; color:#1f2937;
    box-shadow: 0 1px 2px rgba(16,24,40,.04);
}
.timeline { display:flex; gap:10px; flex-wrap:wrap; margin-top:6px; }
.chip {
    background:#eef6ff; border:1px solid #bcd9f5; color:#12496f;
    border-radius:999px; padding:6px 14px; font-size:.86rem;
}
.metric-card { padding:16px 18px; border-radius:14px; background:#fff; color:#0f172a; }
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------
# STATE BOOTSTRAP
# -----------------------
if "cfg" not in st.session_state:
    st.session_state.cfg = default_config()

if "_pending_project" in st.session_state:
    loaded = st.session_state.pop("_pending_project")
    merged = default_config()
    merged.update(loaded)
    merged["tools"] = [dict(default_tool(), **t) for t in loaded.get("tools", merged["tools"])]
    merged["tool_count"] = len(merged["tools"]) or merged["tool_count"]
    st.session_state.cfg = merged
    clear_widget_state()

if "_goto" in st.session_state:
    st.session_state["nav"] = st.session_state.pop("_goto")

cfg = st.session_state.cfg

st.markdown(
    '<div class="hero"><h1>🗓️ Development Center Scheduler</h1>'
    '<p>Build assessor & participant day plans in three steps — Setup, Tools, Review. '
    'Your entries are kept as you move between steps.</p></div>',
    unsafe_allow_html=True,
)

# -----------------------
# SIDEBAR: NAV + PROJECTS
# -----------------------
step = st.sidebar.radio("Steps", STEPS, key="nav", label_visibility="collapsed")

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
        projects[name] = copy.deepcopy(cfg)
        save_projects(projects)
        st.sidebar.success(f"Saved '{name}' with all {cfg['tool_count']} tool(s) and settings.")

if projects:
    chosen = st.sidebar.selectbox("Load a saved project", ["—"] + sorted(projects.keys()))
    cload, cdel = st.sidebar.columns(2)
    if cload.button("Load", use_container_width=True) and chosen != "—":
        st.session_state["_pending_project"] = copy.deepcopy(projects[chosen])
        st.rerun()
    if cdel.button("Delete", use_container_width=True) and chosen != "—":
        projects.pop(chosen, None)
        save_projects(projects)
        st.rerun()
else:
    st.sidebar.caption("No saved projects yet.")

# =======================================================
# STEP 1: SETUP
# =======================================================
if step == "1 · Setup":
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown("### ⚙️ Basic configuration")
    st.caption("No confidential or personal data is needed here.")
    c1, c2, c3 = st.columns(3)
    with c1:
        num("Participants", "participants", 1, 50, cfg, "participants")
        num("Assessors", "assessors", 1, 20, cfg, "assessors")
    with c2:
        txt("Start time", "start_time", cfg, "start_time", help="When the DC day begins (HH:MM).")
        txt("End time (before integration)", "end_time", cfg, "end_time",
            help="When the ASSESSMENTS end. Integration, if on, runs AFTER this — End time is not "
                 "the end of the day.")
    with c3:
        num("Context setting (mins)", "context_minutes", 0, 120, cfg, "context_minutes")
        txt("Output file name", "file_name", cfg, "file_name")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown("### 🍴 Lunch window")
    st.caption("A 30-minute lunch is placed inside this window. Keep it tight (e.g. 12:00–12:30) to "
               "send everyone to lunch together, or widen it (e.g. 12:00–13:30) to let groups stagger "
               "— one preps while others eat.")
    lc1, lc2 = st.columns(2)
    with lc1:
        txt("Window start", "lunch_start", cfg, "lunch_start")
    with lc2:
        txt("Window end", "lunch_end", cfg, "lunch_end")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown("### 🤝 Integration")
    st.caption("An all-hands session that runs straight after the assessments — its minutes are added "
               "after End time, on the same day.")
    ic1, ic2 = st.columns([1.3, 1])
    with ic1:
        tog("Add an integration session at the end", "integration_on", cfg, "integration_on")
    with ic2:
        num("Duration (mins)", "integration_minutes", 0, 240, cfg, "integration_minutes", step=5)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown("### ⏱️ Time slot size")
    st.caption("How finely the day is divided. 5 minutes gives the tightest schedule; 10 or 15 "
               "give rounder, easier-to-read times at the cost of a little padding, and solve faster.")
    st.session_state.setdefault("slot_minutes", int(cfg["slot_minutes"]))
    cfg["slot_minutes"] = st.radio(
        "Slot size",
        ALLOWED_SLOT_MINUTES,
        key="slot_minutes",
        horizontal=True,
        format_func=lambda value: f"{value} minutes",
        label_visibility="collapsed",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # Live day timeline
    try:
        s = minute_of(parse_time(cfg["start_time"], "Start"))
        e = minute_of(parse_time(cfg["end_time"], "End"))
        integ = int(cfg["integration_minutes"]) if cfg["integration_on"] else 0
        chips = [f"Day starts {clock(s)}", f"Assessments end {clock(e)}"]
        chips += [f"Integration {clock(e)}–{clock(e + integ)}", f"Day ends {clock(e + integ)}"] if integ \
            else [f"Day ends {clock(e)}"]
        st.markdown('<div class="timeline">' + "".join(f'<span class="chip">{c}</span>' for c in chips)
                    + "</div>", unsafe_allow_html=True)
    except ValueError:
        pass

    st.write("")
    if st.button("Save & continue to Tools →", use_container_width=True):
        st.session_state["_goto"] = "2 · Tools"
        st.rerun()

# =======================================================
# STEP 2: TOOLS
# =======================================================
elif step == "2 · Tools":
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown("### 🧩 Assessment tools")
    st.caption("Define each tool's type and phase timings. Choosing a type pre-fills typical timings — "
               "override any of them.")
    new_count = num("Number of tools", "tool_count", 1, 10, cfg, "tool_count")
    st.markdown("</div>", unsafe_allow_html=True)

    # Keep the tools list in sync with the count.
    tools = cfg["tools"]
    while len(tools) < new_count:
        tools.append(default_tool())
    while len(tools) > new_count:
        tools.pop()

    tool_names = []
    for i in range(new_count):
        tool = tools[i]
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(f"**Tool {i + 1}**")

        type_key = f"t{i}_type"
        st.session_state.setdefault(type_key, tool.get("type", "Standard"))
        top = st.columns([1, 1.4])
        new_type = top[0].selectbox("Type", TOOL_TYPES, key=type_key)

        # When the type changes, refresh phase timings to that type's defaults
        # by resetting the phase widget keys BEFORE those widgets render.
        if new_type != tool.get("type"):
            d = TYPE_DEFAULTS.get(new_type, DEFAULT_TYPE_TIMES)
            st.session_state[f"t{i}_instr"] = d[0]
            st.session_state[f"t{i}_prep"] = d[1]
            st.session_state[f"t{i}_exec"] = d[2]
            st.session_state[f"t{i}_score"] = d[3]
            tool["instruction"], tool["preparation"], tool["execution"], tool["scoring"] = d
        tool["type"] = new_type

        st.session_state.setdefault(f"t{i}_name", tool.get("name", ""))
        tool["name"] = top[1].text_input("Exercise name", key=f"t{i}_name", placeholder=new_type)

        c1, c2, c3, c4 = st.columns(4)
        st.session_state.setdefault(f"t{i}_instr", int(tool["instruction"]))
        st.session_state.setdefault(f"t{i}_prep", int(tool["preparation"]))
        st.session_state.setdefault(f"t{i}_exec", int(tool["execution"]))
        st.session_state.setdefault(f"t{i}_score", int(tool["scoring"]))
        tool["instruction"] = c1.number_input("Instruction", 0, 180, key=f"t{i}_instr")
        tool["preparation"] = c2.number_input("Preparation", 0, 240, key=f"t{i}_prep")
        tool["execution"] = c3.number_input("Execution", 0, 240, key=f"t{i}_exec")
        tool["scoring"] = c4.number_input("Scoring", 0, 120, key=f"t{i}_score")
        st.markdown("</div>", unsafe_allow_html=True)

        tool_names.append(tool["name"] or new_type)

    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown("### 👤 Assessor mapping — optional")
    st.caption("Leave empty for automatic mapping. To set it by hand, pick the tools the SECONDARY "
               "assessor runs; everything else goes to the primary. The rule of thumb is that the "
               "primary does more tools.")
    valid = [n for n in st.session_state.get("secondary_tools", cfg.get("secondary_tools", []))
             if n in tool_names]
    st.session_state["secondary_tools"] = valid
    cfg["secondary_tools"] = st.multiselect(
        "Tools run by the secondary assessor", options=tool_names, key="secondary_tools"
    )
    st.markdown("</div>", unsafe_allow_html=True)

    nav1, nav2 = st.columns(2)
    if nav1.button("← Back to Setup", use_container_width=True):
        st.session_state["_goto"] = "1 · Setup"
        st.rerun()
    if nav2.button("Save & continue to Review →", use_container_width=True):
        st.session_state["_goto"] = "3 · Review"
        st.rerun()

# =======================================================
# STEP 3: REVIEW
# =======================================================
elif step == "3 · Review":
    st.markdown("### 📊 Review & generate")

    try:
        start_dt = parse_time(cfg["start_time"], "Start time")
        end_dt = parse_time(cfg["end_time"], "End time")
        if start_dt >= end_dt:
            st.error("End time must be after start time.")
            st.stop()

        lunch_start_dt = parse_time(cfg["lunch_start"], "Lunch window start")
        lunch_end_dt = parse_time(cfg["lunch_end"], "Lunch window end")
        if minute_of(lunch_end_dt) - minute_of(lunch_start_dt) < 30:
            st.error("The lunch window must be at least 30 minutes wide.")
            st.stop()

        # Apply the chosen granularity before anything is converted to slots.
        slot_minutes = int(cfg.get("slot_minutes", DEFAULT_SLOT_MINUTES))
        set_slot_minutes(slot_minutes)

        tool_count = int(cfg["tool_count"])
        tools = []
        for i in range(tool_count):
            t = cfg["tools"][i]
            name = t["name"] or t["type"]
            instr, prep = t["instruction"], t["preparation"]
            exe, score = t["execution"], t["scoring"]
            tools.append({
                "index": i, "name": name,
                "instruction_slots": minutes_to_slots(instr, slot_minutes),
                "preparation_slots": minutes_to_slots(prep, slot_minutes),
                "execution_slots": minutes_to_slots(exe, slot_minutes),
                "scoring_slots": minutes_to_slots(score, slot_minutes),
                "instruction_minutes": minutes_to_slots(instr, slot_minutes) * slot_minutes,
                "preparation_minutes": minutes_to_slots(prep, slot_minutes) * slot_minutes,
                "execution_minutes": minutes_to_slots(exe, slot_minutes) * slot_minutes,
                "scoring_minutes": minutes_to_slots(score, slot_minutes) * slot_minutes,
            })

        candidates = int(cfg["participants"])
        assessors = int(cfg["assessors"])
        start_minute = minute_of(start_dt)
        end_minute = minute_of(end_dt)
        integration_minutes = int(cfg["integration_minutes"]) if cfg["integration_on"] else 0
        secondary_tools = [t["name"] for t in tools if t["name"] in cfg.get("secondary_tools", [])]

        inputs = {
            "candidates": candidates, "assessors": assessors,
            "start_time": start_dt, "end_time": end_dt,
            "start_minute": start_minute, "end_minute": end_minute,
            "slots_per_day": (end_minute + integration_minutes - start_minute) // slot_minutes,
            "assessment_slots_per_day": (end_minute - start_minute) // slot_minutes,
            "context_slots": minutes_to_slots(cfg["context_minutes"], slot_minutes),
            "context_minutes": minutes_to_slots(cfg["context_minutes"], slot_minutes) * slot_minutes,
            "tools": tools,
            "assessor_names": [""] * assessors,
            "participant_names": [""] * candidates,
            "lunch_window_start_minute": minute_of(lunch_start_dt),
            "lunch_window_end_minute": minute_of(lunch_end_dt),
            "integration_minutes": integration_minutes,
            "secondary_tools": secondary_tools,
        }

        permutation_count = estimate_permutations(candidates, assessors, tool_count)
        mapping_note = "Manual — secondary: " + ", ".join(secondary_tools) if secondary_tools else "Automatic"

        m1, m2, m3 = st.columns(3)
        m1.metric("Combinations checked", f"{permutation_count:,}")
        m2.metric("Assessments end", clock(end_minute))
        m3.metric("Integration", f"{integration_minutes} min" if integration_minutes else "Off")

        st.markdown(
            f'<div class="card">Lunch window <b>{cfg["lunch_start"]}–{cfg["lunch_end"]}</b>'
            f' &nbsp;•&nbsp; Mapping: <b>{mapping_note}</b>'
            f' &nbsp;•&nbsp; Participants <b>{candidates}</b>, Assessors <b>{assessors}</b>,'
            f' &nbsp;•&nbsp; Tools <b>{tool_count}</b> ({", ".join(t["name"] for t in tools)})'
            f' &nbsp;•&nbsp; Slot size <b>{slot_minutes} min</b></div>',
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
                    minute = best_inputs["start_minute"] + local * slot_minutes
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
                file_name = cfg.get("file_name") or "DC_Schedule"
                spread = group_start_spread(result, best_inputs) * slot_minutes
                st.success(
                    f"Done — {result['total_days']} day(s), all groups start within "
                    f"{spread} minutes of each other."
                )
                st.download_button(
                    "⬇️ Download Excel", data=file_bytes, file_name=f"{file_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        if st.button("← Back to Tools", use_container_width=True):
            st.session_state["_goto"] = "2 · Tools"
            st.rerun()

    except ValueError as error:
        st.error(str(error))
    except Exception as error:  # surface engine errors to the user
        st.error(str(error))
