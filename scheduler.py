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

            def _can_work(emp, d, shift):
                """Check non-points constraints (night rules, explicit blocks)."""
                if (emp, d, shift) in constraint_set:
                    return False
                if shift != NIGHT and night_assignments.get(d) == emp:
                    return False
                if shift == NIGHT:
                    if schedule.get((d, MORNING)) == emp or schedule.get((d, EVENING)) == emp:
                        return False
                yesterday = d - timedelta(days=1)
                if night_assignments.get(yesterday) == emp:
                    return False
                return True

            # Tier 1: employees under their weekly target (5 or override)
            # Tier 2: employees at target but can stretch to target+1 (6 or override+1)
            # Tier 3: fallback — anyone who can physically work (ignore points)
            for emp in employees:
                if not _can_work(emp, d, shift):
                    continue
                emp_max = max_weekly_points_override.get(emp, MAX_WEEKLY_POINTS)
                if weekly_points[week][emp] + cost <= emp_max:
                    candidates.append(emp)

            if not candidates:
                # Tier 2: allow +1 over target (e.g., 6 for regular, override+1 for custom)
                for emp in employees:
                    if not _can_work(emp, d, shift):
                        continue
                    emp_max = max_weekly_points_override.get(emp, MAX_WEEKLY_POINTS)
                    if weekly_points[week][emp] + cost <= emp_max + 1:
                        candidates.append(emp)

            if not candidates:
                # Tier 3: last resort — ignore points entirely
                for emp in employees:
                    if _can_work(emp, d, shift):
                        candidates.append(emp)

            if not candidates:
                schedule[(d, shift)] = "❌ אין זמין"
                continue

            # Pick candidate with lowest weekly points first (enforce 5 before allowing 6),
            # then by total points for overall fairness
            random.shuffle(candidates)
            candidates.sort(key=lambda e: (weekly_points[week][e], total_points[e]))
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
