"""
משמרות שבוע 29.3 - 4.4.2026
כולל ערב פסח (1.4) ופסח (2.4)

שינויים מהשבוע הקודם:
- יהב לא עובד השבוע
- אורי → שונה שם ל"צרפתי" (אותם אילוצים)
- אורי הלוי — עובד חדש, מצטרף מרביעי 1.4, שיבוץ קבוע: ערב+לילה ברביעי
  בגליון לשלישות: משמרת הערב של אורי הלוי נפרסת לשבת

אילוצים:
- משה: לא רביעי (1.4), לא שישי (3.4). חמישי לילה = שיבוץ קבוע.
- עומר: עדיף לא ראשון בוקר+ערב, לא שני לילה, לא שלישי כלל, עדיף לא רביעי ערב.
- צרפתי: ראשון לא בוקר, שני לא בוקר+ערב, שלישי לא ערב
- עומרי: יכול רק: ראשון ערב, שני בוקר, שלישי לילה, חמישי ערב. השאר חסום.
- ניר: הכל חוץ מרביעי ערב+לילה
- אורי הלוי: מרביעי בלבד. שיבוץ קבוע: רביעי ערב+לילה.
"""
from datetime import date
from scheduler import generate_schedule, print_schedule, MORNING, EVENING, NIGHT, SHIFTS, SHIFT_COST
from excel_export import export_to_excel
from datetime import timedelta

# === Configuration ===
EMPLOYEES = ["ניר", "משה", "עומר", "עומרי", "צרפתי", "אורי הלוי"]
START_DATE = date(2026, 3, 29)  # ראשון
NUM_DAYS = 7  # 29.3 - 4.4

# === Dates ===
SUN = date(2026, 3, 29)   # ראשון
MON = date(2026, 3, 30)   # שני
TUE = date(2026, 3, 31)   # שלישי
WED = date(2026, 4, 1)    # רביעי — ערב פסח
THU = date(2026, 4, 2)    # חמישי — פסח
FRI = date(2026, 4, 3)    # שישי — חול המועד
SAT = date(2026, 4, 4)    # שבת — חול המועד

# === Holidays ===
HOLIDAYS = {
    WED: "ערב פסח",
    THU: "פסח",
    FRI: "חול המועד פסח",
    SAT: "חול המועד פסח",
}

# === Fixed assignments (שיבוצים קבועים) ===
FIXED_ASSIGNMENTS = {
    # אורי הלוי — רביעי ערב + לילה
    (WED, EVENING): "אורי הלוי",
    (WED, NIGHT): "אורי הלוי",
    # משה — חמישי לילה
    (THU, NIGHT): "משה",
}

# === Employee start dates (עובדים חלקיים) ===
EMPLOYEE_START_DATE = {
    "אורי הלוי": WED,  # מצטרף רק מרביעי
}

# === Excluded employees ===
EXCLUDED = ["יהב"]  # לא עובד השבוע

# === Constraints ===
# עומרי — יכול רק משמרות ספציפיות, חוסמים את כל השאר
OMRI_ALLOWED = {
    (SUN, EVENING),    # ראשון ערב
    (MON, MORNING),    # שני בוקר
    (TUE, NIGHT),      # שלישי לילה
    (THU, EVENING),    # חמישי ערב
}
# בונים אילוצים לעומרי — חוסמים כל מה שלא ברשימה
omri_constraints = []
for i in range(NUM_DAYS):
    d = START_DATE + timedelta(days=i)
    for shift in SHIFTS:
        if (d, shift) not in OMRI_ALLOWED:
            omri_constraints.append((d, shift))

CONSTRAINTS = {
    # צרפתי — אילוצים קבועים (כמו אורי הישן)
    "צרפתי": [
        (SUN, MORNING),          # ראשון: לומד 10-15 → לא בוקר
        (MON, MORNING),          # שני: 12-20 → לא בוקר
        (MON, EVENING),          # שני: 12-20 → לא ערב
        (TUE, EVENING),          # שלישי: 15-20 → לא ערב
    ],
    # משה — לא רביעי, לא שישי
    "משה": [
        (WED, MORNING),
        (WED, EVENING),
        (WED, NIGHT),
        (FRI, MORNING),
        (FRI, EVENING),
        (FRI, NIGHT),
    ],
    # עומר — עדיף לא ראשון בוקר+ערב, לא שני לילה, לא שלישי כלל, עדיף לא רביעי ערב
    "עומר": [
        (SUN, MORNING),          # עדיף לא ראשון בוקר
        (SUN, EVENING),          # עדיף לא ראשון ערב
        (MON, NIGHT),            # לא שני לילה
        (TUE, MORNING),          # לא שלישי
        (TUE, EVENING),          # לא שלישי
        (TUE, NIGHT),            # לא שלישי
        (WED, EVENING),          # עדיף לא רביעי ערב (ערב חג)
    ],
    # ניר — הכל חוץ מרביעי ערב+לילה
    "ניר": [
        (WED, EVENING),
        (WED, NIGHT),
    ],
    # עומרי — כל מה שלא ברשימת המותר
    "עומרי": omri_constraints,
}

# === Generate — Internal sheet (פנימי) ===
print("🔄 מייצר לוח פנימי...")
schedule_internal = generate_schedule(
    employees=EMPLOYEES,
    start_date=START_DATE,
    num_days=NUM_DAYS,
    constraints=CONSTRAINTS,
    allow_double_shift=True,
    no_double_shift_weekday=4,  # Friday
    max_weekly_nights=1,
    night_overflow_preference=["צרפתי"],
    fixed_assignments=FIXED_ASSIGNMENTS,
    employee_start_date=EMPLOYEE_START_DATE,
    excluded_employees=EXCLUDED,
)

# === Generate — External sheet (לשלישות) ===
# מבוסס על הפנימי, מפזרים double shifts לעובדים אחרים
print("🔄 מייצר לוח לשלישות (מבוסס על פנימי, מפזרים כפילויות)...")
schedule_external = dict(schedule_internal)

# מעבר על כל יום — אם לעובד יש יותר ממשמרת אחת, מעבירים את הפחות חשובה למישהו אחר
for i in range(NUM_DAYS):
    d = START_DATE + timedelta(days=i)
    for emp in EMPLOYEES:
        if emp in EXCLUDED:
            continue
        emp_shifts = [s for s in SHIFTS if schedule_external.get((d, s)) == emp]
        if len(emp_shifts) <= 1:
            continue

        # קובעים איזה משמרת להוריד (עדיפות: משאירים לילה > ערב > בוקר)
        priority = {NIGHT: 3, EVENING: 2, MORNING: 1}
        emp_shifts.sort(key=lambda s: priority[s], reverse=True)
        shift_to_keep = emp_shifts[0]
        shifts_to_reassign = emp_shifts[1:]

        for shift in shifts_to_reassign:
            # מנסים למצוא מחליף שלא עובד באותו יום
            best = None
            for candidate in EMPLOYEES:
                if candidate in EXCLUDED:
                    continue
                if candidate == emp:
                    continue
                # בודקים שהמחליף לא עובד כבר באותו יום בלוח הלשלישות
                already_working = any(schedule_external.get((d, s2)) == candidate for s2 in SHIFTS)
                if already_working:
                    continue
                # בודקים אילוצים
                if candidate in EMPLOYEE_START_DATE and d < EMPLOYEE_START_DATE[candidate]:
                    continue
                if any(c == (d, shift) for c in CONSTRAINTS.get(candidate, [])):
                    continue
                best = candidate
                break

            if best:
                schedule_external[(d, shift)] = best
            else:
                # אם אין מחליף — משאירים ריק (מוחקים)
                del schedule_external[(d, shift)]

# === External overrides ===
# בגליון לשלישות: משמרת ערב של אורי הלוי ברביעי נפרסת לשבת
EXTERNAL_OVERRIDES = {
    (SAT, EVENING): "אורי הלוי",
}

# מחילים את ה-overrides על הלוח הלשלישות (גם להדפסה וגם לאקסל)
for (d, shift), emp in EXTERNAL_OVERRIDES.items():
    schedule_external[(d, shift)] = emp

print()

# === Print to console ===
print("📋 לוח פנימי:")
print_schedule(schedule_internal, START_DATE, NUM_DAYS)
print()
print("📋 לוח לשלישות:")
print_schedule(schedule_external, START_DATE, NUM_DAYS)

# === Export to Excel ===
output_file = "/root/shift-scheduler/משמרות_29.3-4.4_v2.xlsx"
export_to_excel(
    schedule_internal=schedule_internal,
    schedule_external=schedule_external,
    employees=EMPLOYEES,
    start_date=START_DATE,
    num_days=NUM_DAYS,
    output_path=output_file,
    holidays=HOLIDAYS,
    external_overrides=EXTERNAL_OVERRIDES,
)
print(f"\n✅ קובץ אקסל נשמר: {output_file}")

# === Print summary ===
print("\n📊 סיכום נקודות (פנימי):")
for emp in EMPLOYEES:
    if emp in EXCLUDED:
        print(f"  {emp}: לא עובד השבוע")
        continue
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
print("  יהב: לא עובד השבוע")
print("  צרפתי: ראשון בוקר ❌ | שני בוקר+ערב ❌ | שלישי ערב ❌")
print("  משה: לא רביעי, לא שישי. חמישי לילה קבוע.")
print("  עומר: עדיף לא ראשון בוקר+ערב, לא שני לילה, לא שלישי, עדיף לא רביעי ערב")
print("  עומרי: רק ראשון ערב, שני בוקר, שלישי לילה, חמישי ערב")
print("  ניר: הכל חוץ מרביעי ערב+לילה")
print("  אורי הלוי: מרביעי, קבוע ערב+לילה ברביעי. לשלישות: ערב נפרס לשבת")
