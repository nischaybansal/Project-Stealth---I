from collections import defaultdict
from datetime import datetime, timedelta
import math

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


SLOT_MINUTES = 5
LUNCH_MINUTES = 30
LUNCH_START_MINUTE = 12 * 60
LUNCH_END_MINUTE = 14 * 60 + 30
PREFERRED_ASSESSOR_LUNCH_MINUTE = 13 * 60
OUTPUT_FILE = "DC_Schedule.xlsx"


ASSESSOR_COLORS = [
    "FFE070",  # yellow
    "9DC3E6",  # blue
    "A9D18E",  # green
    "C6A0DC",  # purple
    "F4B183",  # orange
    "B7DEE8",  # aqua
    "D9EAD3",  # pale green
    "F8CBAD",  # peach
]


def round_up_to_slots(minutes):
    if minutes <= 0:
        return 0
    return int(math.ceil(minutes / SLOT_MINUTES))


def slots_to_minutes(slots):
    return slots * SLOT_MINUTES


def rounded_duration(minutes):
    slots = round_up_to_slots(minutes)
    return slots, slots_to_minutes(slots)


def minutes_from_time(value):
    return value.hour * 60 + value.minute


def read_positive_int(prompt):
    while True:
        value = input(prompt).strip()
        try:
            number = int(value)
            if number > 0:
                return number
        except ValueError:
            pass
        print("Please enter a positive whole number.")


def read_minimum_int(prompt, minimum):
    while True:
        value = input(prompt).strip()
        try:
            number = int(value)
            if number >= minimum:
                return number
        except ValueError:
            pass
        print(f"Please enter a whole number of at least {minimum}.")


def read_non_negative_int(prompt):
    while True:
        value = input(prompt).strip()
        try:
            number = int(value)
            if number >= 0:
                return number
        except ValueError:
            pass
        print("Please enter zero or a positive whole number.")


def read_time(prompt):
    while True:
        value = input(prompt).strip()
        try:
            return datetime.strptime(value, "%H:%M")
        except ValueError:
            print("Please enter time in HH:MM format, for example 09:30.")


def read_end_time(start_time):
    start_minute = minutes_from_time(start_time)
    while True:
        end_time = read_time("DC end time (HH:MM): ")
        if minutes_from_time(end_time) > start_minute:
            return end_time
        print("End time must be later than the start time.")


def read_yes_no(prompt):
    while True:
        value = input(prompt).strip().lower()
        if value in ("y", "yes"):
            return True
        if value in ("n", "no"):
            return False
        print("Please enter y or n.")


def collect_optional_names(count, label, prefix):
    if not read_yes_no(f"Do you want to enter {label} names? (y/n): "):
        return [""] * count

    names = []
    for index in range(1, count + 1):
        names.append(input(f"Enter {label} name for {prefix}{index}: ").strip())
    return names


def collect_inputs():
    print("Enter Development Center schedule details")
    candidates = read_positive_int("Number of candidates: ")
    assessors = read_minimum_int("Number of assessors: ", 2)
    start_time = read_time("Start time (HH:MM): ")
    end_time = read_end_time(start_time)
    context_minutes = read_non_negative_int("Context setting duration (minutes): ")
    tool_count = read_minimum_int("Number of tools: ", 2)
    assessor_names = collect_optional_names(assessors, "assessor", "A")
    participant_names = collect_optional_names(candidates, "participant", "P")

    start_minute = minutes_from_time(start_time)
    end_minute = minutes_from_time(end_time)
    slots_per_day = (end_minute - start_minute) // SLOT_MINUTES
    if slots_per_day <= 0:
        raise ValueError("The DC day is too short for a 5-minute schedule.")

    tools = []
    for index in range(1, tool_count + 1):
        print(f"\nTool {index}")
        name = input("Tool name: ").strip()
        while not name:
            print("Tool name cannot be empty.")
            name = input("Tool name: ").strip()

        if "bei" in name.lower():
            instruction_minutes = 0
            print("Instruction time set to 0 because tool name contains 'bei'.")
        else:
            instruction_minutes = read_non_negative_int("Instruction time (minutes): ")

        preparation_minutes = read_non_negative_int("Preparation time (minutes): ")
        execution_minutes = read_non_negative_int("Execution time (minutes): ")
        scoring_minutes = read_non_negative_int("Scoring time (minutes): ")

        instruction_slots, rounded_instruction_minutes = rounded_duration(instruction_minutes)
        preparation_slots, rounded_preparation_minutes = rounded_duration(preparation_minutes)
        execution_slots, rounded_execution_minutes = rounded_duration(execution_minutes)
        scoring_slots, rounded_scoring_minutes = rounded_duration(scoring_minutes)

        tools.append(
            {
                "index": index - 1,
                "name": name,
                "instruction_slots": instruction_slots,
                "preparation_slots": preparation_slots,
                "execution_slots": execution_slots,
                "scoring_slots": scoring_slots,
                "instruction_minutes": rounded_instruction_minutes,
                "preparation_minutes": rounded_preparation_minutes,
                "execution_minutes": rounded_execution_minutes,
                "scoring_minutes": rounded_scoring_minutes,
            }
        )

    context_slots, rounded_context_minutes = rounded_duration(context_minutes)
    if context_slots > slots_per_day:
        raise ValueError("Context setting duration is longer than one DC day.")

    return {
        "candidates": candidates,
        "assessors": assessors,
        "start_time": start_time,
        "end_time": end_time,
        "start_minute": start_minute,
        "end_minute": end_minute,
        "slots_per_day": slots_per_day,
        "context_slots": context_slots,
        "context_minutes": rounded_context_minutes,
        "tools": tools,
        "assessor_names": assessor_names,
        "participant_names": participant_names,
    }


def participant_code(number):
    return f"P{number}"


def assessor_code(index):
    return f"A{index + 1}"


def assessor_display_name(index, assessor_names):
    name = assessor_names[index].strip()
    if name:
        return f"{assessor_code(index)} - {name}"
    return assessor_code(index)


def assessor_fill(index):
    return PatternFill("solid", fgColor=ASSESSOR_COLORS[index % len(ASSESSOR_COLORS)])


def lunch_fill():
    return PatternFill("solid", fgColor="F2F2F2")


def phase_sequence(tool):
    return [
        ("Instr", tool["instruction_slots"]),
        ("Prep", tool["preparation_slots"]),
        ("Exec", tool["execution_slots"]),
        ("Score", tool["scoring_slots"]),
    ]


def tool_total_slots(tool):
    return sum(duration for _, duration in phase_sequence(tool))


def phase_needs_assessor(phase_name):
    return phase_name in ("Exec", "Score")


def create_groups(candidate_count, assessor_count):
    groups = []
    candidate_number = 1
    group_index = 0

    while candidate_number <= candidate_count:
        group_size = min(assessor_count, candidate_count - candidate_number + 1)
        members = list(range(candidate_number, candidate_number + group_size))
        candidate_number += group_size
        primary_assessor = group_index % assessor_count

        groups.append(
            {
                "name": f"G{group_index + 1}",
                "index": group_index,
                "members": members,
                "primary_assessor": primary_assessor,
                "secondary_assessor": paired_assessor_index(primary_assessor, assessor_count),
                "ready_slot": 0,
                "next_tool_index": 0,
                "tool_order": [],
                "lunch_days": set(),
            }
        )
        group_index += 1

    return groups


def paired_assessor_index(index, assessor_count):
    if assessor_count % 2 == 0:
        return index + 1 if index % 2 == 0 else index - 1
    return (index + 1) % assessor_count


def rotate_tools_for_groups(groups, tools):
    for index, group in enumerate(groups):
        rotation = index % len(tools)
        group["tool_order"] = tools[rotation:] + tools[:rotation]


def assigned_assessor_for_tool(group, tool):
    if tool["index"] % 2 == 0:
        return group["primary_assessor"]
    return group["secondary_assessor"]


def participant_assessor_pair(participant, assessor_count):
    primary = (participant - 1) % assessor_count
    secondary = paired_assessor_index(primary, assessor_count)
    return primary, secondary


def assigned_assessor_for_participant(participant, tool, assessor_count):
    primary, secondary = participant_assessor_pair(participant, assessor_count)
    if tool["index"] % 2 == 0:
        return primary
    return secondary


def group_assessor_assignments(group, tool, inputs):
    return {
        participant: assigned_assessor_for_participant(participant, tool, inputs["assessors"])
        for participant in group["members"]
    }


def day_index(slot, inputs):
    return slot // inputs["slots_per_day"]


def local_slot(slot, inputs):
    return slot % inputs["slots_per_day"]


def day_start_slot(day, inputs):
    return day * inputs["slots_per_day"]


def minute_for_slot(slot, inputs):
    return inputs["start_minute"] + local_slot(slot, inputs) * SLOT_MINUTES


def time_text_from_minute(minute):
    hour = minute // 60
    minute_part = minute % 60
    return f"{hour:02d}:{minute_part:02d}"


def format_time(inputs, slot, include_day=False):
    label = time_text_from_minute(minute_for_slot(slot, inputs))
    if include_day:
        return f"Day {day_index(slot, inputs) + 1} {label}"
    return label


def format_boundary_time(inputs, slot, include_day=False):
    if slot > 0 and slot % inputs["slots_per_day"] == 0:
        boundary_day = (slot // inputs["slots_per_day"]) - 1
        label = time_text_from_minute(inputs["end_minute"])
        if include_day:
            return f"Day {boundary_day + 1} {label}"
        return label
    return format_time(inputs, slot, include_day)


def format_time_range(inputs, slot, include_day=False):
    start_label = time_text_from_minute(minute_for_slot(slot, inputs))
    end_minute = minute_for_slot(slot, inputs) + SLOT_MINUTES
    end_label = time_text_from_minute(end_minute)
    if include_day:
        return f"Day {day_index(slot, inputs) + 1} {start_label}-{end_label}"
    return f"{start_label}-{end_label}"


def normalize_start_slot(slot, total_slots, inputs):
    local = local_slot(slot, inputs)
    if local + total_slots > inputs["slots_per_day"]:
        return day_start_slot(day_index(slot, inputs) + 1, inputs)
    return slot


def participant_lunch_window_slots(inputs, day):
    window_start = max(inputs["start_minute"], LUNCH_START_MINUTE)
    window_end = min(inputs["end_minute"], LUNCH_END_MINUTE)
    if window_end - window_start < LUNCH_MINUTES:
        return None

    start_local = int(math.ceil((window_start - inputs["start_minute"]) / SLOT_MINUTES))
    end_local = int(math.floor((window_end - inputs["start_minute"]) / SLOT_MINUTES))
    return day_start_slot(day, inputs) + start_local, day_start_slot(day, inputs) + end_local


def assessor_lunch_local_range(inputs):
    window_start = max(inputs["start_minute"], LUNCH_START_MINUTE)
    window_end = min(inputs["end_minute"], LUNCH_END_MINUTE)
    if window_end - window_start < LUNCH_MINUTES:
        return None

    lunch_start = min(
        max(PREFERRED_ASSESSOR_LUNCH_MINUTE, window_start),
        window_end - LUNCH_MINUTES,
    )
    start_local = int(math.ceil((lunch_start - inputs["start_minute"]) / SLOT_MINUTES))
    lunch_slots = round_up_to_slots(LUNCH_MINUTES)
    return start_local, start_local + lunch_slots


def is_assessor_lunch_slot(slot, inputs):
    lunch_range = inputs.get("assessor_lunch_local_range")
    if lunch_range is None:
        return False

    start_local, end_local = lunch_range
    local = local_slot(slot, inputs)
    return start_local <= local < end_local


def assessor_lunch_text(inputs):
    lunch_range = inputs.get("assessor_lunch_local_range")
    if lunch_range is None:
        return "Not available inside DC hours"

    start_local, end_local = lunch_range
    start_minute = inputs["start_minute"] + start_local * SLOT_MINUTES
    end_minute = inputs["start_minute"] + end_local * SLOT_MINUTES
    return f"{time_text_from_minute(start_minute)}-{time_text_from_minute(end_minute)}"


def schedule_context(schedule, candidate_count, context_slots):
    for candidate in range(1, candidate_count + 1):
        for slot in range(context_slots):
            schedule[candidate][slot] = "Context Setting"


def maybe_schedule_group_lunch(schedule, group, inputs, next_tool=None, warnings=None):
    current_day = day_index(group["ready_slot"], inputs)
    if current_day in group["lunch_days"]:
        return False

    window = participant_lunch_window_slots(inputs, current_day)
    if window is None:
        if warnings is not None:
            warnings.add("No 30-minute lunch window exists inside the DC hours.")
        return False

    lunch_start, lunch_end = window
    lunch_slots = round_up_to_slots(LUNCH_MINUTES)
    ready_slot = group["ready_slot"]

    if ready_slot > lunch_end - lunch_slots:
        return False

    should_lunch_now = lunch_start <= ready_slot <= lunch_end - lunch_slots
    if not should_lunch_now and ready_slot < lunch_start and next_tool is not None:
        latest_possible_end = ready_slot + tool_total_slots(next_tool)
        should_lunch_now = latest_possible_end > lunch_end - lunch_slots
        if should_lunch_now:
            ready_slot = lunch_start

    if not should_lunch_now:
        return False

    if local_slot(ready_slot, inputs) + lunch_slots > inputs["slots_per_day"]:
        return False

    for slot in range(ready_slot, ready_slot + lunch_slots):
        for candidate in group["members"]:
            schedule[candidate][slot] = "Lunch"

    group["ready_slot"] = ready_slot + lunch_slots
    group["lunch_days"].add(current_day)
    return True


def block_feasible(
    start_slot,
    tool,
    assignments,
    exec_counts,
    score_counts,
    assessor_busy,
    inputs,
):
    if normalize_start_slot(start_slot, tool_total_slots(tool), inputs) != start_slot:
        return False

    current_slot = start_slot
    for phase_name, duration_slots in phase_sequence(tool):
        for offset in range(duration_slots):
            slot = current_slot + offset

            if phase_needs_assessor(phase_name):
                if is_assessor_lunch_slot(slot, inputs):
                    return False
                for assessor in assignments.values():
                    if assessor_busy[assessor].get(slot):
                        return False

            if phase_name == "Exec":
                if exec_counts.get(slot, 0) + len(assignments) > inputs["assessors"]:
                    return False
                if score_counts.get(slot, 0) > 0:
                    return False

            if phase_name == "Score":
                if exec_counts.get(slot, 0) > 0:
                    return False
                if score_counts.get(slot, 0) + len(assignments) > inputs["assessors"]:
                    return False

        current_slot += duration_slots

    return True


def apply_block(
    schedule,
    schedule_assessors,
    group,
    tool,
    assignments,
    start_slot,
    exec_counts,
    score_counts,
    assessor_busy,
    summary,
    assessor_mapping,
    participant_assessors,
    inputs,
):
    current_slot = start_slot
    block_start = start_slot

    for phase_name, duration_slots in phase_sequence(tool):
        for offset in range(duration_slots):
            slot = current_slot + offset
            value = f"{tool['name']} {phase_name}"

            for candidate in group["members"]:
                schedule[candidate][slot] = value
                if phase_needs_assessor(phase_name):
                    schedule_assessors[candidate][slot] = assignments[candidate]

            if phase_needs_assessor(phase_name):
                for candidate, assessor in assignments.items():
                    assessor_busy[assessor][slot] = (
                        f"{group['name']} {participant_code(candidate)} {tool['name']} {phase_name}"
                    )

            if phase_name == "Exec":
                exec_counts[slot] += len(assignments)
            elif phase_name == "Score":
                score_counts[slot] += len(assignments)

        current_slot += duration_slots

    for candidate in group["members"]:
        assessor = assignments[candidate]
        assessor_mapping[assessor][tool["index"]].append(candidate)
        participant_assessors[candidate].add(assessor)

    include_day = current_slot > inputs["slots_per_day"]
    assessor_text = ", ".join(
        f"{participant_code(candidate)}:{assessor_display_name(assignments[candidate], inputs['assessor_names'])}"
        for candidate in group["members"]
    )
    summary.append(
        {
            "group": group["name"],
            "members": ", ".join(participant_code(member) for member in group["members"]),
            "tool": tool["name"],
            "assessor": assessor_text,
            "start": format_time(inputs, block_start, include_day),
            "end": format_boundary_time(inputs, current_slot, include_day),
        }
    )

    group["ready_slot"] = current_slot
    group["next_tool_index"] += 1


def build_schedule(inputs):
    candidate_count = inputs["candidates"]
    tools = inputs["tools"]
    context_slots = inputs["context_slots"]

    schedule = defaultdict(dict)
    schedule_assessors = defaultdict(dict)
    exec_counts = defaultdict(int)
    score_counts = defaultdict(int)
    assessor_busy = defaultdict(dict)
    summary = []
    warnings = set()
    assessor_mapping = {
        assessor_index: {tool["index"]: [] for tool in tools}
        for assessor_index in range(inputs["assessors"])
    }
    participant_assessors = defaultdict(set)

    inputs["assessor_lunch_local_range"] = assessor_lunch_local_range(inputs)
    if inputs["assessor_lunch_local_range"] is None:
        warnings.add("No 30-minute assessor lunch window exists inside the DC hours.")

    schedule_context(schedule, candidate_count, context_slots)

    groups = create_groups(candidate_count, inputs["assessors"])
    rotate_tools_for_groups(groups, tools)
    for group in groups:
        group["ready_slot"] = context_slots

    for tool in tools:
        if tool_total_slots(tool) > inputs["slots_per_day"]:
            raise ValueError(
                f"{tool['name']} is longer than one DC day after rounding to {SLOT_MINUTES}-minute slots."
            )

    while any(group["next_tool_index"] < len(tools) for group in groups):
        groups.sort(key=lambda item: (item["ready_slot"], item["index"]))

        for group in groups:
            if group["next_tool_index"] >= len(tools):
                continue

            tool = group["tool_order"][group["next_tool_index"]]
            maybe_schedule_group_lunch(schedule, group, inputs, tool, warnings)

            assignments = group_assessor_assignments(group, tool, inputs)
            start_slot = group["ready_slot"]
            total_slots = tool_total_slots(tool)
            search_limit = start_slot + inputs["slots_per_day"] * 30

            while True:
                start_slot = normalize_start_slot(start_slot, total_slots, inputs)
                if start_slot > search_limit:
                    raise ValueError(
                        f"Could not schedule {group['name']} for {tool['name']} within 30 DC days."
                    )
                if block_feasible(
                    start_slot,
                    tool,
                    assignments,
                    exec_counts,
                    score_counts,
                    assessor_busy,
                    inputs,
                ):
                    break
                start_slot += 1

            group["ready_slot"] = start_slot
            maybe_schedule_group_lunch(schedule, group, inputs, tool, warnings)
            start_slot = group["ready_slot"]

            while True:
                start_slot = normalize_start_slot(start_slot, total_slots, inputs)
                if start_slot > search_limit:
                    raise ValueError(
                        f"Could not schedule {group['name']} for {tool['name']} within 30 DC days."
                    )
                if block_feasible(
                    start_slot,
                    tool,
                    assignments,
                    exec_counts,
                    score_counts,
                    assessor_busy,
                    inputs,
                ):
                    break
                start_slot += 1

            apply_block(
                schedule,
                schedule_assessors,
                group,
                tool,
                assignments,
                start_slot,
                exec_counts,
                score_counts,
                assessor_busy,
                summary,
                assessor_mapping,
                participant_assessors,
                inputs,
            )

    for group in groups:
        maybe_schedule_group_lunch(schedule, group, inputs, None, warnings)

    max_slot = context_slots
    for candidate_slots in schedule.values():
        if candidate_slots:
            max_slot = max(max_slot, max(candidate_slots) + 1)

    return {
        "schedule": schedule,
        "schedule_assessors": schedule_assessors,
        "groups": groups,
        "summary": summary,
        "exec_counts": exec_counts,
        "score_counts": score_counts,
        "assessor_busy": assessor_busy,
        "assessor_mapping": assessor_mapping,
        "participant_assessors": participant_assessors,
        "warnings": sorted(warnings),
        "max_slot": max_slot,
    }


def validate_schedule(result, inputs):
    errors = quality_check_schedule(result, inputs)
    if errors:
        formatted_errors = "\n".join(f"- {error}" for error in errors)
        raise ValueError(f"Schedule failed quality checks:\n{formatted_errors}")


def quality_check_schedule(result, inputs):
    errors = []
    all_slots = set(result["exec_counts"].keys()) | set(result["score_counts"].keys())

    for group in result["groups"]:
        if len(group["members"]) > inputs["assessors"]:
            errors.append(
                f"{group['name']} has {len(group['members'])} participants. "
                f"Groups cannot be larger than the assessor count ({inputs['assessors']})."
            )

    ordered_groups = sorted(result["groups"], key=lambda item: item["index"])
    expected_candidates = [
        candidate
        for group in ordered_groups
        for candidate in group["members"]
    ]
    if expected_candidates != list(range(1, inputs["candidates"] + 1)):
        errors.append("Groups do not cover participants in sequential pair order.")

    for slot in all_slots:
        if result["exec_counts"][slot] > inputs["assessors"]:
            errors.append(
                f"Execution constraint failed at slot {slot}: "
                f"{result['exec_counts'][slot]} participants for {inputs['assessors']} assessors."
            )

        if result["score_counts"][slot] > inputs["assessors"]:
            errors.append(
                f"Scoring constraint failed at slot {slot}: "
                f"{result['score_counts'][slot]} participants for {inputs['assessors']} assessors."
            )

        if result["exec_counts"][slot] > 0 and result["score_counts"][slot] > 0:
            errors.append(f"Execution and scoring overlap at slot {slot}.")

    for slot in range(result["max_slot"]):
        active_exec_groups = []
        active_assessor_groups = []
        for group in result["groups"]:
            group_values = {
                result["schedule"][candidate].get(slot)
                for candidate in group["members"]
            }
            group_values.discard(None)
            group_values.discard("")

            if any(isinstance(value, str) and value.endswith(" Exec") for value in group_values):
                active_exec_groups.append(group["name"])

            if any(
                isinstance(value, str)
                and (value.endswith(" Exec") or value.endswith(" Score"))
                for value in group_values
            ):
                active_assessor_groups.append(group["name"])

        if len(active_exec_groups) > inputs["assessors"]:
            errors.append(
                f"Too many execution groups at {format_time_range(inputs, slot, True)}: "
                f"{', '.join(active_exec_groups)}."
            )

        if len(active_assessor_groups) > inputs["assessors"]:
            errors.append(
                f"Too many assessor-required groups at {format_time_range(inputs, slot, True)}: "
                f"{', '.join(active_assessor_groups)}."
            )

    for assessor_index, busy_slots in result["assessor_busy"].items():
        seen_slots = set()
        for slot in busy_slots:
            if slot in seen_slots:
                errors.append(f"{assessor_code(assessor_index)} is double-booked at slot {slot}.")
            seen_slots.add(slot)

    for participant in range(1, inputs["candidates"] + 1):
        assessor_total = len(result["participant_assessors"][participant])
        if assessor_total != 2:
            errors.append(
                f"{participant_code(participant)} is mapped to {assessor_total} assessor(s), "
                "but must be mapped to exactly 2."
            )

    return errors


def tool_name_from_activity(activity, tools):
    for tool in tools:
        if activity.startswith(f"{tool['name']} "):
            return tool["name"]
    return None


def write_schedule_sheet(workbook, inputs, result):
    ws = workbook.active
    ws.title = "Schedule"

    candidate_count = inputs["candidates"]
    schedule_column_count = candidate_count + 1
    mapping_start_column = schedule_column_count + 3
    mapping_column_count = len(inputs["tools"]) + 2
    mapping_end_column = mapping_start_column + mapping_column_count - 1
    participant_table_start_row = 3 + inputs["assessors"] + 2
    participant_table_end_row = participant_table_start_row + candidate_count
    used_max_column = mapping_end_column
    max_slot = result["max_slot"]
    schedule_end_row = max_slot + 3
    used_max_row = max(schedule_end_row, participant_table_end_row)
    include_day = max_slot > inputs["slots_per_day"]

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=used_max_column)
    ws["A1"] = "DC Schedule"

    tool_names = ", ".join(tool["name"] for tool in inputs["tools"])
    summary_text = (
        f"Start: {inputs['start_time'].strftime('%H:%M')} | "
        f"End: {inputs['end_time'].strftime('%H:%M')} | "
        f"Candidates: {inputs['candidates']} | "
        f"Assessors: {inputs['assessors']} | "
        f"Slot Size: {SLOT_MINUTES} minutes | "
        f"Assessor Lunch: {assessor_lunch_text(inputs)} | "
        f"Tools: {tool_names}"
    )
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=used_max_column)
    ws["A2"] = summary_text

    ws.cell(row=3, column=1, value="Time")
    for candidate in range(1, candidate_count + 1):
        ws.cell(row=3, column=candidate + 1, value=participant_code(candidate))

    for slot in range(max_slot):
        row = slot + 4
        ws.cell(row=row, column=1, value=format_time_range(inputs, slot, include_day))

        for candidate in range(1, candidate_count + 1):
            cell = ws.cell(row=row, column=candidate + 1)
            value = result["schedule"][candidate].get(slot, "")
            cell.value = value
            if value == "Lunch":
                cell.fill = lunch_fill()
            elif isinstance(value, str) and value.endswith(" Score"):
                cell.fill = PatternFill("solid", fgColor="BFBFBF")
            elif isinstance(value, str) and value.endswith(" Exec") and slot in result["schedule_assessors"][candidate]:
                cell.fill = assessor_fill(result["schedule_assessors"][candidate][slot])

    write_mapping_box(ws, inputs, result, 3, mapping_start_column)
    write_participant_table(ws, inputs, participant_table_start_row, mapping_start_column)

    format_schedule_sheet(ws, candidate_count, max_slot, used_max_column, used_max_row)
    format_mapping_box(ws, inputs, 3, mapping_start_column)
    format_participant_table(ws, inputs, participant_table_start_row, mapping_start_column)
    merge_repeated_activities(ws, candidate_count, max_slot)


def write_mapping_box(ws, inputs, result, start_row, start_column):
    headers = ["Assessors", "Room Name"] + [tool["name"] for tool in inputs["tools"]]
    for offset, header in enumerate(headers):
        ws.cell(row=start_row, column=start_column + offset, value=header)

    for assessor_index in range(inputs["assessors"]):
        row = start_row + assessor_index + 1
        ws.cell(
            row=row,
            column=start_column,
            value=assessor_display_name(assessor_index, inputs["assessor_names"]),
        )
        ws.cell(row=row, column=start_column + 1, value="")

        for tool in inputs["tools"]:
            participants = sorted(set(result["assessor_mapping"][assessor_index][tool["index"]]))
            ws.cell(
                row=row,
                column=start_column + 2 + tool["index"],
                value=",".join(str(participant) for participant in participants),
            )


def write_participant_table(ws, inputs, start_row, start_column):
    headers = ["P.NO", "Participant Name"]
    for offset, header in enumerate(headers):
        ws.cell(row=start_row, column=start_column + offset, value=header)

    for participant in range(1, inputs["candidates"] + 1):
        row = start_row + participant
        ws.cell(row=row, column=start_column, value=participant_code(participant))
        ws.cell(
            row=row,
            column=start_column + 1,
            value=inputs["participant_names"][participant - 1],
        )


def format_schedule_sheet(ws, candidate_count, max_slot, used_max_column, used_max_row):
    schedule_max_column = candidate_count + 1
    schedule_max_row = max_slot + 3
    dark_blue = PatternFill("solid", fgColor="1F4E78")
    light_blue = PatternFill("solid", fgColor="D9EAF7")
    thin_side = Side(style="thin", color="808080")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws["A1"].fill = dark_blue
    ws["A1"].font = Font(color="FFFFFF", bold=True, size=16)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["A2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["A2"].font = Font(bold=True)

    for row in (1, 2):
        for column in range(1, used_max_column + 1):
            ws.cell(row=row, column=column).border = thin_border

    for column in range(1, schedule_max_column + 1):
        header_cell = ws.cell(row=3, column=column)
        header_cell.fill = light_blue
        header_cell.font = Font(bold=True)
        header_cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in range(3, schedule_max_row + 1):
        for column in range(1, schedule_max_column + 1):
            cell = ws.cell(row=row, column=column)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 20
    for column in range(2, schedule_max_column + 1):
        ws.column_dimensions[get_column_letter(column)].width = 18

    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 42
    ws.row_dimensions[3].height = 22
    for row in range(4, used_max_row + 1):
        ws.row_dimensions[row].height = 24

    ws.freeze_panes = "A4"


def format_mapping_box(ws, inputs, start_row, start_column):
    header_fill = PatternFill("solid", fgColor="5B6770")
    thin_side = Side(style="thin", color="000000")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    column_count = len(inputs["tools"]) + 2

    for offset in range(column_count):
        cell = ws.cell(row=start_row, column=start_column + offset)
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for assessor_index in range(inputs["assessors"]):
        row = start_row + assessor_index + 1
        for offset in range(column_count):
            cell = ws.cell(row=row, column=start_column + offset)
            cell.fill = assessor_fill(assessor_index)
            cell.font = Font(bold=True)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    widths = [26, 18] + [24] * len(inputs["tools"])
    for offset, width in enumerate(widths):
        ws.column_dimensions[get_column_letter(start_column + offset)].width = width


def format_participant_table(ws, inputs, start_row, start_column):
    header_fill = PatternFill("solid", fgColor="5B6770")
    thin_side = Side(style="thin", color="000000")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    max_row = start_row + inputs["candidates"]

    for offset in range(2):
        cell = ws.cell(row=start_row, column=start_column + offset)
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in range(start_row + 1, max_row + 1):
        for offset in range(2):
            cell = ws.cell(row=row, column=start_column + offset)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if offset == 0:
                cell.font = Font(bold=True)


def merge_repeated_activities(ws, candidate_count, max_slot):
    first_schedule_row = 4
    last_schedule_row = max_slot + 3

    for column in range(2, candidate_count + 2):
        merge_start = first_schedule_row
        current_value = ws.cell(row=merge_start, column=column).value

        for row in range(first_schedule_row + 1, last_schedule_row + 2):
            next_value = ws.cell(row=row, column=column).value if row <= last_schedule_row else None

            if next_value != current_value:
                if current_value not in (None, "") and row - merge_start > 1:
                    ws.merge_cells(
                        start_row=merge_start,
                        start_column=column,
                        end_row=row - 1,
                        end_column=column,
                    )
                    ws.cell(row=merge_start, column=column).alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=True,
                    )

                merge_start = row
                current_value = next_value


def write_summary_sheet(workbook, inputs, result):
    ws = workbook.create_sheet("Group Summary")
    headers = ["Group", "Members", "Tool", "Assessor", "Start", "End"]

    for column, header in enumerate(headers, start=1):
        ws.cell(row=1, column=column, value=header)

    for row_index, item in enumerate(result["summary"], start=2):
        ws.cell(row=row_index, column=1, value=item["group"])
        ws.cell(row=row_index, column=2, value=item["members"])
        ws.cell(row=row_index, column=3, value=item["tool"])
        ws.cell(row=row_index, column=4, value=item["assessor"])
        ws.cell(row=row_index, column=5, value=item["start"])
        ws.cell(row=row_index, column=6, value=item["end"])

    format_summary_sheet(ws, len(result["summary"]) + 1, len(headers))


def format_summary_sheet(ws, max_row, max_column):
    light_blue = PatternFill("solid", fgColor="D9EAF7")
    thin_side = Side(style="thin", color="808080")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for column in range(1, max_column + 1):
        cell = ws.cell(row=1, column=column)
        cell.fill = light_blue
        cell.font = Font(bold=True)

    for row in range(1, max_row + 1):
        for column in range(1, max_column + 1):
            cell = ws.cell(row=row, column=column)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    widths = [12, 28, 24, 24, 18, 18]
    for column, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(column)].width = width

    ws.freeze_panes = "A2"


def create_excel(inputs, result, output_file=OUTPUT_FILE):
    validate_schedule(result, inputs)
    workbook = Workbook()
    write_schedule_sheet(workbook, inputs, result)
    write_summary_sheet(workbook, inputs, result)
    workbook.save(output_file)


def print_warnings(result):
    for warning in result["warnings"]:
        print(f"Warning: {warning}")


def main():
    inputs = collect_inputs()
    result = build_schedule(inputs)
    create_excel(inputs, result)
    print_warnings(result)
    print(f"\nSuccess: schedule saved as {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
