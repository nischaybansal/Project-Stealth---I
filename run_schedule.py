"""Command-line runner for the DC Scheduler.

Build a schedule without the app and without answering prompts, so a
configuration can be re-run identically as many times as you like.

Examples
--------
Nine participants, three assessors, three tools:

    python run_schedule.py --participants 9 --assessors 3 \\
        --start 09:00 --end 19:00 --integration 60 \\
        --tool "BEI:0,0,90,10" \\
        --tool "Case Study:10,70,30,10" \\
        --tool "Role Play:10,30,30,10"

Each --tool is  NAME:instruction,preparation,execution,scoring  in minutes.

Useful extras:
    --slot 5|10|15        time slot size (default 5)
    --lunch 12:00-13:30   window the 30-minute lunch is placed inside
    --context 30          context setting minutes
    --secondary "Role Play"   tool(s) the secondary assessor runs (repeatable)
    --out MySchedule.xlsx     output file name
    --legacy              use the old search engine instead of the solver
    --quiet               print the summary only
"""

import argparse
import sys

import scheduler as S


def parse_tool(text, index):
    if ":" not in text:
        raise ValueError(f"--tool needs NAME:instr,prep,exec,score (got {text!r})")

    name, _, numbers = text.partition(":")
    parts = [piece.strip() for piece in numbers.split(",")]

    if len(parts) != 4:
        raise ValueError(f"--tool {name!r} needs exactly 4 numbers, got {len(parts)}")

    try:
        instruction, preparation, execution, scoring = (int(piece) for piece in parts)
    except ValueError:
        raise ValueError(f"--tool {name!r} has a non-numeric duration")

    slots = [S.round_up_to_slots(value) for value in
             (instruction, preparation, execution, scoring)]

    return {
        "index": index,
        "name": name.strip() or f"Tool {index + 1}",
        "instruction_slots": slots[0],
        "preparation_slots": slots[1],
        "execution_slots": slots[2],
        "scoring_slots": slots[3],
        "instruction_minutes": slots[0] * S.SLOT_MINUTES,
        "preparation_minutes": slots[1] * S.SLOT_MINUTES,
        "execution_minutes": slots[2] * S.SLOT_MINUTES,
        "scoring_minutes": slots[3] * S.SLOT_MINUTES,
    }


def minutes_of(text, label):
    try:
        hours, minutes = text.strip().split(":")
        return int(hours) * 60 + int(minutes)
    except ValueError:
        raise ValueError(f"{label} must look like 09:00 (got {text!r})")


def describe(inputs, result, quiet=False):
    slot = S.SLOT_MINUTES
    start_minute = inputs["start_minute"]
    per_day = inputs["slots_per_day"]

    def clock(absolute_slot):
        day = absolute_slot // per_day
        minute = start_minute + (absolute_slot % per_day) * slot
        stamp = f"{minute // 60:02d}:{minute % 60:02d}"
        return f"Day {day + 1} {stamp}" if result["total_days"] > 1 else stamp

    print()
    print(f"Days: {result['total_days']}")
    print(f"Groups start within: "
          f"{S.slots_to_minutes(S.group_start_spread(result, inputs))} minutes")

    last = max(
        (s for c in result["schedule"] for s, v in result["schedule"][c].items()
         if v and v != "Integration"),
        default=0,
    )
    print(f"Assessments end: {clock(last + 1)}")
    if inputs.get("integration_minutes"):
        print(f"Integration: {clock(last + 1)} - {clock(result['max_slot'])}")

    if quiet:
        return

    print()
    for group in sorted(result["groups"], key=lambda item: item["index"]):
        busy = sorted({
            s for member in group["members"]
            for s, v in result["schedule"][member].items()
            if v and v not in ("Lunch", "Integration", "Context Setting")
        })
        lunch = sorted(
            s for s, v in result["schedule"][group["members"][0]].items() if v == "Lunch"
        )
        order = " -> ".join(tool["name"] for tool in group["tool_order"])
        print(f"{group['name']} ({', '.join(S.participant_code(m) for m in group['members'])})")
        print(f"   order  : {order}")
        print(f"   starts : {clock(busy[0])}   ends: {clock(busy[-1] + 1)}")
        print(f"   lunch  : {clock(lunch[0]) if lunch else '-'}")

    print()
    for assessor in range(inputs["assessors"]):
        worked = [s for s, v in result["assessor_busy"][assessor].items()
                  if v and v != "Lunch"]
        if not worked:
            continue
        span = max(worked) - min(worked) + 1
        used = len(worked)
        print(f"A{assessor + 1}: busy {S.slots_to_minutes(used)} min of "
              f"{S.slots_to_minutes(span)} min on duty "
              f"({used / span * 100:.0f}% utilised)")


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Build a DC schedule from the command line.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--participants", type=int, required=True)
    parser.add_argument("--assessors", type=int, required=True)
    parser.add_argument("--start", default="09:00")
    parser.add_argument("--end", default="19:00", help="when ASSESSMENTS end")
    parser.add_argument("--context", type=int, default=30)
    parser.add_argument("--integration", type=int, default=0)
    parser.add_argument("--lunch", default="12:00-13:30")
    parser.add_argument("--slot", type=int, default=5, choices=S.ALLOWED_SLOT_MINUTES)
    parser.add_argument("--tool", action="append", required=True,
                        help="NAME:instr,prep,exec,score (repeat per tool)")
    parser.add_argument("--secondary", action="append", default=[],
                        help="tool run by the SECONDARY assessor (repeatable)")
    parser.add_argument("--out", default="DC_Schedule.xlsx")
    parser.add_argument("--legacy", action="store_true",
                        help="use the old search engine instead of the solver")
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args(argv)

    S.set_slot_minutes(args.slot)

    tools = [parse_tool(text, index) for index, text in enumerate(args.tool)]

    start_minute = minutes_of(args.start, "--start")
    end_minute = minutes_of(args.end, "--end")
    if end_minute <= start_minute:
        raise ValueError("--end must be after --start")

    if "-" not in args.lunch:
        raise ValueError("--lunch must look like 12:00-13:30")
    lunch_from, lunch_to = args.lunch.split("-", 1)
    lunch_start = minutes_of(lunch_from, "--lunch start")
    lunch_end = minutes_of(lunch_to, "--lunch end")
    if lunch_end - lunch_start < S.LUNCH_MINUTES:
        raise ValueError("--lunch window must be at least 30 minutes wide")

    # The rendered day holds the assessments, the overrun buffer and then the
    # integration that follows them.
    full_day = end_minute + S.OVERRUN_GRACE_MINUTES + args.integration - start_minute

    inputs = {
        "candidates": args.participants,
        "assessors": args.assessors,
        "start_time": S.time_from_minutes(start_minute)
        if hasattr(S, "time_from_minutes") else __import__("datetime").datetime(
            1900, 1, 1, start_minute // 60, start_minute % 60),
        "end_time": __import__("datetime").datetime(
            1900, 1, 1, end_minute // 60, end_minute % 60),
        "start_minute": start_minute,
        "end_minute": end_minute,
        # Rounded up, so a schedule finishing in the last partial slot is not
        # reported as a phantom second day.
        "slots_per_day": -(-full_day // args.slot),
        "assessment_slots_per_day": (end_minute - start_minute) // args.slot,
        "context_slots": S.round_up_to_slots(args.context),
        "context_minutes": S.round_up_to_slots(args.context) * args.slot,
        "tools": tools,
        "assessor_names": [""] * args.assessors,
        "participant_names": [""] * args.participants,
        "lunch_window_start_minute": lunch_start,
        "lunch_window_end_minute": lunch_end,
        "integration_minutes": args.integration,
        "secondary_tools": list(args.secondary),
        "force_legacy_search": args.legacy,
    }

    best_inputs, result = S.find_fastest_schedule(inputs)
    describe(best_inputs, result, quiet=args.quiet)

    S.create_excel(best_inputs, result, args.out)
    print()
    print(f"Saved: {args.out}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except ValueError as error:
        print(f"Error: {error}", file=sys.stderr)
        sys.exit(1)
