"""
Shift Scheduler Engine
חלוקת משמרות אוטומטית עם אילוצים
"""
import random
from datetime import date, timedelta
from typing import Dict, List, Optional, Set, Tuple

# Shift types
MORNING = "בוקר"      # 8:00-16:00
EVENING = "ערב"       # 14:00-22:00
NIGHT = "לילה"        # 20:00-9:00

SHIFTS = [MORNING, EVENING, NIGHT]
SHIFT_HOURS = {
    MORNING: "8:00-16:00",
    EVENING: "14:00-22:00",
    NIGHT: "20:00-9:00",
}

# Night shift costs 2 points, others cost 1
SHIFT_COST = {
    MORNING: 1,
    EVENING: 1,
    NIGHT: 2,
}

MAX_WEEKLY_POINTS = 5

DAYS_HEB = {
    0: "שני",
    1: "שלישי",
    2: "רביעי",
    3: "חמישי",
    4: "שישי",
    5: "שבת",
    6: "ראשון",
}


def get_day_name(d: date) -> str:
    return DAYS_HEB[d.weekday()]


def get_week_number(d: date, start_date: date) -> int:
    """Return which week (0-based) a date falls in relative to start_date."""
    return (d - start_date).days // 7


def generate_schedule(
    employees: List[str],
    start_date: date,
    num_days: int,
    constraints: Optional[Dict[str, List[Tuple[date, str]]]] = None,
    max_weekly_points_override: Optional[Dict[str, int]] = None,
) -> Dict[Tuple[date, str], str]:
    """
    Generate a shift schedule.

    Args:
        employees: List of employee names
        start_date: First day of the schedule
        num_days: Number of days to schedule
        constraints: Dict mapping employee name -> list of (date, shift) they CANNOT do
        max_weekly_points_override: Dict mapping employee name -> custom max weekly points

    Returns:
        Dict mapping (date, shift_type) -> employee_name
    """
    if constraints is None:
        constraints = {}
    if max_weekly_points_override is None:
        max_weekly_points_override = {}

    # Build constraint set for quick lookup: (employee, date, shift)
    constraint_set: Set[Tuple[str, date, str]] = set()
    for emp, blocked in constraints.items():
        for d, s in blocked:
            constraint_set.add((emp, d, s))

    dates = [start_date + timedelta(days=i) for i in range(num_days)]

    # Track assignments
    schedule: Dict[Tuple[date, str], str] = {}

    # Track weekly points per employee: week_number -> employee -> points
    weekly_points: Dict[int, Dict[str, int]] = {}

    # Track who did night shift on which date
    night_assignments: Dict[date, str] = {}

    # Track total points for fairness
    total_points: Dict[str, int] = {emp: 0 for emp in employees}

    # Schedule day by day, shift by shift
    for d in dates:
        week = get_week_number(d, start_date)
        if week not in weekly_points:
            weekly_points[week] = {emp: 0 for emp in employees}

        for shift in SHIFTS:
            cost = SHIFT_COST[shift]
            candidates = []

            for emp in employees:
                # Check explicit constraints
                if (emp, d, shift) in constraint_set:
                    continue

                # Check weekly points limit (per-employee override or global)
                emp_max = max_weekly_points_override.get(emp, MAX_WEEKLY_POINTS)
                if weekly_points[week][emp] + cost > emp_max:
                    continue

                # Night shift rule: on the day of night shift, no other shifts
                if shift != NIGHT and night_assignments.get(d) == emp:
                    continue

                # Night shift rule: if this IS a night shift, employee can't have
                # done morning or evening today
                if shift == NIGHT:
                    if schedule.get((d, MORNING)) == emp or schedule.get((d, EVENING)) == emp:
                        continue

                # After night shift: no shifts the next day
                yesterday = d - timedelta(days=1)
                if night_assignments.get(yesterday) == emp:
                    continue

                # Night shift rule: on the day of night shift, no other shifts
                # Check if employee is already assigned to night tonight
                # (night is last in SHIFTS order, so check if assigned to morning/evening)
                if shift == NIGHT:
                    for other_shift in [MORNING, EVENING]:
                        if schedule.get((d, other_shift)) == emp:
                            break
                    else:
                        candidates.append(emp)
                    continue

                candidates.append(emp)

            if not candidates:
                # Fallback: try anyone (relaxing weekly limit)
                for emp in employees:
                    can_do = True
                    if (emp, d, shift) in constraint_set:
                        can_do = False
                    if shift != NIGHT and night_assignments.get(d) == emp:
                        can_do = False
                    if shift == NIGHT and (schedule.get((d, MORNING)) == emp or schedule.get((d, EVENING)) == emp):
                        can_do = False
                    yesterday = d - timedelta(days=1)
                    if night_assignments.get(yesterday) == emp:
                        can_do = False
                    if can_do:
                        candidates.append(emp)

            if not candidates:
                schedule[(d, shift)] = "❌ אין זמין"
                continue

            # Pick candidate with lowest total points (fairness), with randomness for ties
            random.shuffle(candidates)
            candidates.sort(key=lambda e: total_points[e])
            chosen = candidates[0]

            schedule[(d, shift)] = chosen
            weekly_points[week][chosen] += cost
            total_points[chosen] += cost

            if shift == NIGHT:
                night_assignments[d] = chosen

    return schedule


def print_schedule(
    schedule: Dict[Tuple[date, str], str],
    start_date: date,
    num_days: int,
):
    """Print schedule to console for debugging."""
    dates = [start_date + timedelta(days=i) for i in range(num_days)]

    # Header
    print(f"{'משמרת':<12}", end="")
    for d in dates:
        print(f"{d.strftime('%d.%m')} {get_day_name(d):<10}", end="")
    print()

    for shift in SHIFTS:
        print(f"{shift} ({SHIFT_HOURS[shift]})", end="  ")
        for d in dates:
            emp = schedule.get((d, shift), "—")
            print(f"{emp:<16}", end="")
        print()
