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
    allow_double_shift: bool = False,
    no_double_shift_weekday: Optional[int] = None,
    ignore_constraints: bool = False,
    max_weekly_nights: int = 1,
    night_overflow_preference: Optional[List[str]] = None,
    gap_fill_employees: Optional[List[str]] = None,
) -> Dict[Tuple[date, str], str]:
    """
    Generate a shift schedule.

    Args:
        employees: List of employee names
        start_date: First day of the schedule
        num_days: Number of days to schedule
        constraints: Dict mapping employee name -> list of (date, shift) they CANNOT do
        max_weekly_points_override: Dict mapping employee name -> custom max weekly points
        allow_double_shift: If True, same employee can do 2 shifts in one day (e.g. morning+evening)
        no_double_shift_weekday: weekday number (0=Mon..6=Sun) on which double shifts are forbidden
        ignore_constraints: If True, employee constraints are ignored (used for "לשלישות" sheet)
        max_weekly_nights: Max night shifts per employee per week (default 1, hard cap 2)
        night_overflow_preference: Employees preferred for a 2nd night if needed (e.g. ["אורי"])
        gap_fill_employees: Employees allowed to fill open slots when all others are at their target
                            (e.g. ["יהב"] — used only when there is no other option)

    Returns:
        Dict mapping (date, shift_type) -> employee_name
    """
    if constraints is None:
        constraints = {}
    if max_weekly_points_override is None:
        max_weekly_points_override = {}
    if night_overflow_preference is None:
        night_overflow_preference = []
    if gap_fill_employees is None:
        gap_fill_employees = []

    # Build constraint set for quick lookup: (employee, date, shift)
    constraint_set: Set[Tuple[str, date, str]] = set()
    if not ignore_constraints:
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

    # Track weekly night count per employee: week_number -> employee -> nights
    weekly_nights: Dict[int, Dict[str, int]] = {}

    # Track total points for fairness
    total_points: Dict[str, int] = {emp: 0 for emp in employees}

    def _can_work(emp, d, shift):
        """Check non-points constraints (night rules, explicit blocks)."""
        if (emp, d, shift) in constraint_set:
            return False

        # Night shift rules always apply (safety)
        if shift != NIGHT and night_assignments.get(d) == emp:
            return False
        if shift == NIGHT:
            if schedule.get((d, MORNING)) == emp or schedule.get((d, EVENING)) == emp:
                return False
        yesterday = d - timedelta(days=1)
        if night_assignments.get(yesterday) == emp:
            return False

        # Double-shift guard: same employee can't do 2 non-night shifts in one day
        # unless allow_double_shift is on (and not on the restricted weekday)
        if not allow_double_shift or (no_double_shift_weekday is not None and d.weekday() == no_double_shift_weekday):
            if shift != NIGHT:
                for other_shift in (MORNING, EVENING):
                    if other_shift != shift and schedule.get((d, other_shift)) == emp:
                        return False

        return True

    def _remaining_slots(emp, current_date, week):
        """Count how many shift slots remain for this employee from current_date to end of week."""
        slots = 0
        week_start = start_date + timedelta(days=week * 7)
        week_end_day = min(week_start + timedelta(days=6), start_date + timedelta(days=num_days - 1))
        d = current_date
        while d <= week_end_day:
            for s in SHIFTS:
                if _can_work(emp, d, s) and (d, s) not in schedule:
                    slots += SHIFT_COST[s]
            d += timedelta(days=1)
        return slots

    def _night_ok(emp, week):
        """Check whether emp is allowed another night this week."""
        nights_done = weekly_nights.get(week, {}).get(emp, 0)
        if nights_done == 0:
            return True
        if nights_done >= 2:
            return False
        # nights_done == 1 → allowed only if emp is in overflow preference list
        return emp in night_overflow_preference

    # Schedule day by day, shift by shift
    for d in dates:
        week = get_week_number(d, start_date)
        if week not in weekly_points:
            weekly_points[week] = {emp: 0 for emp in employees}
        if week not in weekly_nights:
            weekly_nights[week] = {emp: 0 for emp in employees}

        for shift in SHIFTS:
            cost = SHIFT_COST[shift]
            candidates = []

            # For night shifts also enforce the weekly night cap
            def _eligible_for_shift(emp):
                if not _can_work(emp, d, shift):
                    return False
                if shift == NIGHT and not _night_ok(emp, week):
                    return False
                return True

            # Tier 1: employees under their weekly points target (5 or override)
            # Gap-fill employees (e.g. יהב) are excluded from Tier 1 unless
            # everyone else is already at their target (handled in Tier 1b below).
            regular_employees = [e for e in employees if e not in gap_fill_employees]

            for emp in regular_employees:
                if not _eligible_for_shift(emp):
                    continue
                emp_max = max_weekly_points_override.get(emp, MAX_WEEKLY_POINTS)
                if weekly_points[week][emp] + cost <= emp_max:
                    candidates.append(emp)

            if not candidates:
                # Tier 1b: all regular employees are at/over target — allow gap-fill employees
                for emp in gap_fill_employees:
                    if not _eligible_for_shift(emp):
                        continue
                    emp_max = max_weekly_points_override.get(emp, MAX_WEEKLY_POINTS)
                    if weekly_points[week][emp] + cost <= emp_max:
                        candidates.append(emp)

            if not candidates:
                # Tier 2: allow +1 over target for regular employees
                for emp in regular_employees:
                    if not _eligible_for_shift(emp):
                        continue
                    emp_max = max_weekly_points_override.get(emp, MAX_WEEKLY_POINTS)
                    if weekly_points[week][emp] + cost <= emp_max + 1:
                        candidates.append(emp)

            if not candidates:
                # Tier 2b: allow gap-fill employees +1 as well
                for emp in gap_fill_employees:
                    if not _eligible_for_shift(emp):
                        continue
                    emp_max = max_weekly_points_override.get(emp, MAX_WEEKLY_POINTS)
                    if weekly_points[week][emp] + cost <= emp_max + 1:
                        candidates.append(emp)

            if not candidates:
                # Tier 3: last resort — ignore points entirely (but keep night cap)
                for emp in employees:
                    if _eligible_for_shift(emp):
                        candidates.append(emp)

            if not candidates:
                # Tier 4: absolute last resort — ignore night cap too
                for emp in employees:
                    if _can_work(emp, d, shift):
                        candidates.append(emp)

            if not candidates:
                schedule[(d, shift)] = "❌ אין זמין"
                continue

            # Sort: urgency first (fewer remaining opportunities), then points, then total
            def sort_key(e):
                emp_max = max_weekly_points_override.get(e, MAX_WEEKLY_POINTS)
                still_needed = max(0, emp_max - weekly_points[week][e])
                remaining = _remaining_slots(e, d, week)
                urgency = -(still_needed / max(remaining, 1))
                # For night shifts: prefer employees with 0 nights done this week,
                # then overflow-preference employees, then others
                night_penalty = 0
                if shift == NIGHT:
                    nights_done = weekly_nights[week].get(e, 0)
                    is_preferred = 1 if e in night_overflow_preference else 2
                    night_penalty = (nights_done, is_preferred)
                else:
                    night_penalty = (0, 0)
                return (night_penalty, urgency, weekly_points[week][e], total_points[e])

            random.shuffle(candidates)
            candidates.sort(key=sort_key)
            chosen = candidates[0]

            schedule[(d, shift)] = chosen
            weekly_points[week][chosen] += cost
            total_points[chosen] += cost

            if shift == NIGHT:
                night_assignments[d] = chosen
                weekly_nights[week][chosen] = weekly_nights[week].get(chosen, 0) + 1

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
