"""
משמרות שבוע 29.3 - 4.4.2026
כולל ערב פסח (1.4) ופסח (2.4)

אילוצים:
- יהב: מקסימום 2 משמרות בשבוע (2 נקודות)
- אורי (קבוע):
    ראשון: לומד 10-15 → לא יכול בוקר (8-16)
    שני: 12-20 לא יכול → לא בוקר (8-16) ולא ערב (14-22)
    שלישי: 15-20 → לא ערב (14-22)
- ניר: מעדיף לא לילה ברביעי (אילוץ בדוי)
- משה: לא יכול ערב פסח ופסח (1-2.4) — חופשה (אילוץ בדוי)
- עומר: מעדיף לא בוקר בשישי (אילוץ בדוי)
- עומרי: לא יכול ערב בשלישי (אילוץ בדוי)
"""
from datetime import date
from scheduler import generate_schedule, print_schedule, MORNING, EVENING, NIGHT
from excel_export import export_to_excel

# === Configuration ===
EMPLOYEES = ["ניר", "משה", "עומר", "עומרי", "אורי", "יהב"]
START_DATE = date(2026, 3, 29)  # ראשון
NUM_DAYS = 7  # שבוע אחד: 29.3 - 4.4

# === Holidays ===
HOLIDAYS = {
    date(2026, 4, 1): "ערב פסח",
    date(2026, 4, 2): "פסח",
    date(2026, 4, 3): "חול המועד פסח",
    date(2026, 4, 4): "חול המועד פסח",
}

# === Constraints ===
CONSTRAINTS = {
    # אורי — אילוצים קבועים
    "אורי": [
        # ראשון 29.3: לומד 10-15 → חופף לבוקר (8-16)
        (date(2026, 3, 29), MORNING),
        # שני 30.3: 12-20 לא יכול → חופף לבוקר (8-16) וערב (14-22)
        (date(2026, 3, 30), MORNING),
        (date(2026, 3, 30), EVENING),
        # שלישי 31.3: 15-20 → חופף לערב (14-22)
        (date(2026, 3, 31), EVENING),
    ],
    # משה — חופשה בערב פסח ופסח (בדוי)
    "משה": [
        (date(2026, 4, 1), MORNING),
        (date(2026, 4, 1), EVENING),
        (date(2026, 4, 1), NIGHT),
        (date(2026, 4, 2), MORNING),
        (date(2026, 4, 2), EVENING),
        (date(2026, 4, 2), NIGHT),
    ],
    # ניר — מעדיף לא לילה ברביעי (בדוי)
    "ניר": [
        (date(2026, 4, 1), NIGHT),  # ערב פסח — לא לילה
    ],
    # עומר — מעדיף לא בוקר בשישי (בדוי)
    "עומר": [
        (date(2026, 4, 3), MORNING),  # שישי חוה"מ — לא בוקר
    ],
    # עומרי — לא יכול ערב בשלישי (בדוי)
    "עומרי": [
        (date(2026, 3, 31), EVENING),  # שלישי — לא ערב
    ],
}

# === Per-employee weekly point limits ===
# יהב: מקסימום 1 משמרת (1 נקודה), אם אין ברירה 2
MAX_WEEKLY_OVERRIDE = {
    "יהב": 1,
}

# === Generate — Internal sheet (פנימי) ===
# With constraints, double shifts allowed (except Friday = weekday 4)
print("🔄 מייצר לוח פנימי (עם אילוצים, אפשר 2 משמרות ביום חוץ מיום שישי)...")
schedule_internal = generate_schedule(
    employees=EMPLOYEES,
    start_date=START_DATE,
    num_days=NUM_DAYS,
    constraints=CONSTRAINTS,
    max_weekly_points_override=MAX_WEEKLY_OVERRIDE,
    allow_double_shift=True,
    no_double_shift_weekday=4,  # Friday
)

# === Generate — External sheet (לשלישות) ===
# No constraints, no double shifts in same day
print("🔄 מייצר לוח לשלישות (ללא אילוצים, ללא 2 משמרות ביום)...")
schedule_external = generate_schedule(
    employees=EMPLOYEES,
    start_date=START_DATE,
    num_days=NUM_DAYS,
    constraints=CONSTRAINTS,
    max_weekly_points_override=MAX_WEEKLY_OVERRIDE,
    allow_double_shift=False,
    ignore_constraints=True,
)

print()

# === Print to console ===
print("📋 לוח פנימי:")
print_schedule(schedule_internal, START_DATE, NUM_DAYS)
print()
print("📋 לוח לשלישות:")
print_schedule(schedule_external, START_DATE, NUM_DAYS)

# === Export to Excel ===
output_file = "/root/shift-scheduler/משמרות_29.3-4.4.xlsx"
export_to_excel(
    schedule_internal=schedule_internal,
    schedule_external=schedule_external,
    employees=EMPLOYEES,
    start_date=START_DATE,
    num_days=NUM_DAYS,
    output_path=output_file,
    holidays=HOLIDAYS,
)
print(f"\n✅ קובץ אקסל נשמר: {output_file}")

# === Print summary ===
from scheduler import SHIFT_COST, SHIFTS
from datetime import timedelta

print("\n📊 סיכום נקודות (פנימי):")
for emp in EMPLOYEES:
    total = 0
    shifts_list = []
    for i in range(NUM_DAYS):
        d = START_DATE + timedelta(days=i)
        for shift in SHIFTS:
            if schedule_internal.get((d, shift)) == emp:
                total += SHIFT_COST[shift]
                shifts_list.append(f"{d.strftime('%d.%m')} {shift}")
    print(f"  {emp}: {total} נקודות — {', '.join(shifts_list)}")

# === Print constraints info ===
print("\n📌 אילוצים שהוגדרו:")
print("  יהב: מקס 1 משמרת בשבוע (2 אם אין ברירה)")
print("  אורי: ראשון בוקר ❌ | שני בוקר+ערב ❌ | שלישי ערב ❌")
print("  משה: חופשה 1-2.4 (ערב פסח + פסח)")
print("  ניר: לא לילה בערב פסח")
print("  עומר: לא בוקר בשישי חוה\"מ")
print("  עומרי: לא ערב בשלישי")
