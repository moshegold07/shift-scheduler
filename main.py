"""
Shift Scheduler — Main Script
הרצה לדוגמה עם הנתונים מהצילום מסך
"""
from datetime import date
from scheduler import generate_schedule, print_schedule, MORNING, EVENING, NIGHT
from excel_export import export_to_excel

# === Configuration ===
EMPLOYEES = ["ניר", "משה", "עומר", "אורי", "יהב"]
START_DATE = date(2026, 3, 15)  # 15.3 — ראשון
NUM_DAYS = 14  # שבועיים

# === Constraints (example) ===
# Format: employee -> [(date, shift), ...]
# e.g., אורי can't do morning on 18.3
CONSTRAINTS = {
    # "אורי": [(date(2026, 3, 18), MORNING)],
}

# === Generate ===
print("🔄 מייצר לוח משמרות...")
schedule = generate_schedule(
    employees=EMPLOYEES,
    start_date=START_DATE,
    num_days=NUM_DAYS,
    constraints=CONSTRAINTS,
)

# === Print to console ===
print("\n📋 לוח משמרות:")
print_schedule(schedule, START_DATE, NUM_DAYS)

# === Export to Excel ===
output_file = "/root/shift-scheduler/משמרות.xlsx"
export_to_excel(
    schedule=schedule,
    employees=EMPLOYEES,
    start_date=START_DATE,
    num_days=NUM_DAYS,
    output_path=output_file,
)
print(f"\n✅ קובץ אקסל נשמר: {output_file}")

# === Print summary ===
from scheduler import SHIFT_COST, SHIFTS
from datetime import timedelta

print("\n📊 סיכום נקודות:")
for emp in EMPLOYEES:
    total = 0
    for i in range(NUM_DAYS):
        d = START_DATE + timedelta(days=i)
        for shift in SHIFTS:
            if schedule.get((d, shift)) == emp:
                total += SHIFT_COST[shift]
    print(f"  {emp}: {total} נקודות")
