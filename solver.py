"""Constraint-based scheduling core for the DC Scheduler.

The original engine placed every tool at the earliest slot it could and searched
over tool orders to compensate. That is greedy and myopic: it can never hold a
group back so a long assessor-bound tool lands better later, which is exactly
what a consultant does by hand. No amount of searching over orderings recovers
that, because the limitation lives in the placer, not in the order.

This module states the day as constraints and lets a solver (OR-Tools CP-SAT)
search properly, with backtracking. Delaying a tool to open a window later is a
move the solver makes natively.

The result dictionary is the same shape the Excel writer already consumes, so
the reporting, mapping and app layers are unchanged.
"""

import collections

from ortools.sat.python import cp_model

import scheduler as S


# The solver is given a bounded amount of time; DC days are small, so this is
# generous. If it cannot prove optimality it still returns the best schedule
# found, which is what a consultant would want.
SOLVE_TIME_SECONDS = 30

# Groups should begin within this of each other. Participants feel the wait
# before they start far more than a slightly later finish, so this is expressed
# as a hard constraint and the finish time is what gets optimised. If no
# schedule satisfies it, it is relaxed step by step rather than abandoned.
DEFAULT_MAX_START_SPREAD_MINUTES = 60
SPREAD_RELAXATION_STEPS = (60, 90, 120, 180, 240, 360, 600)


def _tool_phases(tool):
    return [
        ("Instr", tool["instruction_slots"]),
        ("Prep", tool["preparation_slots"]),
        ("Exec", tool["execution_slots"]),
        ("Score", tool["scoring_slots"]),
    ]


def _assessor_assignments(groups, tools, inputs):
    """Which assessor takes which participant, per tool.

    Delegates to the existing mapping logic so the primary/secondary rotation
    and the manual override behave exactly as before.
    """
    assignments = {}
    for group in groups:
        for tool in tools:
            assignments[(group["index"], tool["index"])] = S.group_assessor_assignments(
                group, tool, inputs
            )
    return assignments


def _build_model(inputs, groups, tools, assignments, horizon, max_spread_slots,
                 excluded=None):
    model = cp_model.CpModel()

    context_slots = inputs.get("context_slots", 0)
    lunch_len = S.lunch_slots()

    window = S.lunch_window_slot_range(0, inputs)
    if window is None:
        raise ValueError("No lunch window fits inside the day.")
    lunch_earliest, lunch_latest = window

    starts = {}
    blocks = {}
    exec_blocks = {}
    group_intervals = collections.defaultdict(list)
    assessor_intervals = collections.defaultdict(list)

    for group in groups:
        for tool in tools:
            key = (group["index"], tool["index"])
            total = S.tool_total_slots(tool)

            start = model.NewIntVar(context_slots, horizon, f"s{key}")
            end = model.NewIntVar(context_slots, horizon, f"e{key}")
            interval = model.NewIntervalVar(start, total, end, f"i{key}")

            starts[key] = start
            blocks[key] = (start, end, total)
            group_intervals[group["index"]].append(interval)

            # Only Exec and Scoring occupy an assessor. Instructions and
            # preparation are unsupervised, which is what lets other groups work
            # while one group has all the assessors.
            lead = tool["instruction_slots"] + tool["preparation_slots"]
            busy_len = tool["execution_slots"] + tool["scoring_slots"]

            if busy_len <= 0:
                continue

            exec_start = model.NewIntVar(context_slots, horizon, f"xs{key}")
            exec_end = model.NewIntVar(context_slots, horizon, f"xe{key}")
            model.Add(exec_start == start + lead)
            exec_interval = model.NewIntervalVar(exec_start, busy_len, exec_end, f"xi{key}")
            exec_blocks[key] = (exec_start, exec_end, busy_len)

            used = {assignments[key][member] for member in group["members"]}
            for assessor in used:
                assessor_intervals[assessor].append(exec_interval)

    # A group does one thing at a time, lunch included.
    group_lunch = {}
    for group in groups:
        lunch_start = model.NewIntVar(lunch_earliest, lunch_latest, f"gl{group['index']}")
        lunch_end = model.NewIntVar(0, horizon, f"gle{group['index']}")
        lunch_interval = model.NewIntervalVar(lunch_start, lunch_len, lunch_end, f"gli{group['index']}")
        group_lunch[group["index"]] = (lunch_start, lunch_end)
        group_intervals[group["index"]].append(lunch_interval)
        model.AddNoOverlap(group_intervals[group["index"]])

    # An assessor cannot run two executions at once, and needs a lunch of their
    # own inside the window.
    assessor_lunch = {}
    for assessor in range(inputs["assessors"]):
        lunch_start = model.NewIntVar(lunch_earliest, lunch_latest, f"al{assessor}")
        lunch_end = model.NewIntVar(0, horizon, f"ale{assessor}")
        lunch_interval = model.NewIntervalVar(lunch_start, lunch_len, lunch_end, f"ali{assessor}")
        assessor_lunch[assessor] = (lunch_start, lunch_end)
        assessor_intervals[assessor].append(lunch_interval)
        model.AddNoOverlap(assessor_intervals[assessor])

    # When a group is at lunch its assessors are free to be elsewhere, but the
    # group's own tools obviously cannot run — already covered above.

    # Every group starts within max_spread of every other.
    first_activity = {}
    for group in groups:
        first = model.NewIntVar(context_slots, horizon, f"f{group['index']}")
        model.AddMinEquality(first, [starts[(group["index"], tool["index"])] for tool in tools])
        first_activity[group["index"]] = first

    earliest_first = model.NewIntVar(context_slots, horizon, "ef")
    latest_first = model.NewIntVar(context_slots, horizon, "lf")
    model.AddMinEquality(earliest_first, list(first_activity.values()))
    model.AddMaxEquality(latest_first, list(first_activity.values()))
    model.Add(latest_first - earliest_first <= max_spread_slots)

    # Rule out schedules already shown, so "another option" returns a genuinely
    # different arrangement rather than the same one again. A solution differs
    # if any single tool starts at a different time.
    for index, previous in enumerate(excluded or []):
        same = []
        for key, var in starts.items():
            if key not in previous:
                continue
            flag = model.NewBoolVar(f"same{index}_{key}")
            model.Add(var == previous[key]).OnlyEnforceIf(flag)
            model.Add(var != previous[key]).OnlyEnforceIf(flag.Not())
            same.append(flag)
        if same:
            model.Add(sum(same) <= len(same) - 1)

    # Finish as early as possible; that is what shortens the day.
    makespan = model.NewIntVar(context_slots, horizon, "makespan")
    model.AddMaxEquality(makespan, [end for _, end, _ in blocks.values()])
    model.Minimize(makespan * 1000 + (latest_first - earliest_first))

    return model, starts, group_lunch, assessor_lunch, makespan


def _extract(inputs, groups, tools, assignments, solved_starts, group_lunch, assessor_lunch):
    """Turn solved start times into the result structure the writer expects."""
    schedule = {c: {} for c in range(1, inputs["candidates"] + 1)}
    schedule_assessors = {c: {} for c in range(1, inputs["candidates"] + 1)}
    assessor_busy = {a: {} for a in range(inputs["assessors"])}
    exec_counts = collections.Counter()
    score_counts = collections.Counter()
    assessor_mapping = {
        a: collections.defaultdict(list) for a in range(inputs["assessors"])
    }
    participant_assessors = {c: set() for c in range(1, inputs["candidates"] + 1)}
    summary = []

    context_slots = inputs.get("context_slots", 0)
    for slot in range(context_slots):
        for candidate in schedule:
            schedule[candidate][slot] = "Context Setting"

    ordered = []
    for group in groups:
        for tool in tools:
            ordered.append((solved_starts[(group["index"], tool["index"])], group, tool))
    ordered.sort(key=lambda item: (item[0], item[1]["index"]))

    for start_slot, group, tool in ordered:
        key = (group["index"], tool["index"])
        assign = assignments[key]
        current = start_slot

        for phase_name, duration in _tool_phases(tool):
            for offset in range(duration):
                slot = current + offset
                value = f"{tool['name']} {phase_name}"

                for member in group["members"]:
                    schedule[member][slot] = value
                    if phase_name in ("Exec", "Score"):
                        schedule_assessors[member][slot] = assign[member]

                if phase_name in ("Exec", "Score"):
                    for member, assessor in assign.items():
                        assessor_busy[assessor][slot] = (
                            f"{group['name']} {S.participant_code(member)} "
                            f"{tool['name']} {phase_name}"
                        )
                    if phase_name == "Exec":
                        exec_counts[slot] += len(assign)
                    else:
                        score_counts[slot] += len(assign)

            current += duration

        for member in group["members"]:
            assessor_mapping[assign[member]][tool["index"]].append(member)
            participant_assessors[member].add(assign[member])

        include_day = current > inputs["slots_per_day"]
        summary.append({
            "group": group["name"],
            "members": ", ".join(S.participant_code(m) for m in group["members"]),
            "tool": tool["name"],
            "assessor": ", ".join(
                f"{S.participant_code(m)}:"
                f"{S.assessor_display_name(assign[m], inputs['assessor_names'])}"
                for m in group["members"]
            ),
            "start": S.format_time(inputs, start_slot, include_day),
            "end": S.format_boundary_time(inputs, current, include_day),
        })

    # Tool order per group, read back from when things actually happened.
    for group in groups:
        group["tool_order"] = [
            tool for _, tool in sorted(
                ((solved_starts[(group["index"], t["index"])], t) for t in tools),
                key=lambda item: item[0],
            )
        ]
        group["lunch_days"] = {0}

    lunch_len = S.lunch_slots()
    for group in groups:
        start = group_lunch[group["index"]]
        for slot in range(start, start + lunch_len):
            for member in group["members"]:
                schedule[member][slot] = "Lunch"

    assessor_lunches = []
    for assessor, start in assessor_lunch.items():
        for slot in range(start, start + lunch_len):
            assessor_busy[assessor][slot] = "Lunch"
        assessor_lunches.append({
            "day": 0,
            "assessor": assessor,
            "start_slot": start,
            "end_slot": start + lunch_len,
        })

    max_slot = max(
        (slot + 1 for c in schedule for slot, v in schedule[c].items() if v),
        default=context_slots,
    )

    integration_slots = S.round_up_to_slots(inputs.get("integration_minutes", 0))
    if integration_slots > 0:
        for slot in range(max_slot, max_slot + integration_slots):
            for candidate in schedule:
                schedule[candidate][slot] = "Integration"
        max_slot += integration_slots

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
        "assessor_lunches": assessor_lunches,
        "warnings": [],
        "max_slot": max_slot,
        "total_days": S.total_days_for_slot_count(max_slot, inputs),
    }


def solve_schedule(inputs, max_spread_minutes=DEFAULT_MAX_START_SPREAD_MINUTES,
                   time_limit=SOLVE_TIME_SECONDS, excluded=None):
    """Solve one DC day. Returns (inputs, result) or raises ValueError.

    Pass `excluded` -- a list of solution keys returned as result["solution_key"]
    -- to get a different schedule from the ones already shown.
    """
    tools = inputs["tools"]
    groups = S.create_groups(inputs["candidates"], inputs["assessors"])
    assignments = _assessor_assignments(groups, tools, inputs)

    # Tools may run a little past End Time rather than pushing the day into a
    # second one; integration follows whatever the real finish turns out to be.
    horizon = S.assessment_slots_with_grace(inputs)

    steps = [max_spread_minutes] + [
        step for step in SPREAD_RELAXATION_STEPS if step > max_spread_minutes
    ]

    for spread_minutes in steps:
        max_spread_slots = S.round_up_to_slots(spread_minutes)
        model, starts, group_lunch, assessor_lunch, _ = _build_model(
            inputs, groups, tools, assignments, horizon, max_spread_slots, excluded
        )

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit
        solver.parameters.num_search_workers = 8
        status = solver.Solve(model)

        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            continue

        solved_starts = {key: solver.Value(var) for key, var in starts.items()}
        solved_group_lunch = {g: solver.Value(v[0]) for g, v in group_lunch.items()}
        solved_assessor_lunch = {a: solver.Value(v[0]) for a, v in assessor_lunch.items()}

        result = _extract(
            inputs, groups, tools, assignments,
            solved_starts, solved_group_lunch, solved_assessor_lunch,
        )
        # Keep the raw start times so this exact schedule can be excluded when
        # the user asks for another option.
        result["solution_key"] = solved_starts
        return inputs, result

    if excluded:
        raise ValueError(
            "No further schedules are meaningfully different from the ones "
            "already shown for this setup."
        )

    # Adding assessors in proportion to participants does not help: a group's
    # execution occupies every assessor either way, so the groups just get
    # smaller. What decides whether a day fits is the assessor-bound time per
    # group multiplied by the number of groups.
    locked = sum(
        S.slots_to_minutes(tool["execution_slots"] + tool["scoring_slots"])
        for tool in tools
    )
    group_count = len(groups)
    needed = locked * group_count + S.LUNCH_MINUTES
    available = (
        inputs["end_minute"] - inputs["start_minute"]
        - S.slots_to_minutes(inputs.get("context_slots", 0))
    )

    raise ValueError(
        f"This day does not fit. The tools need {needed} minutes of assessor time "
        f"({locked} min per group x {group_count} groups, plus lunch) but the day "
        f"has {available} minutes after context setting. Shorten an "
        f"assessor-led tool, drop a tool, or extend the day. Adding assessors and "
        f"participants together will not help."
    )
