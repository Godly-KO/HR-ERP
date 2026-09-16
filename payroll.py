"""Pure attendance and payroll calculations for the HR ERP app."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
import re
from typing import Any

import pandas as pd


LATE_MARK_LIMIT = 4
EARLY_LEAVE_LIMIT = 2
DAY_DIVISOR = 30
REGULAR_END_MINUTE = 18 * 60
FLEX_START_MINUTE = 9 * 60
LATE_MINUTE = 10 * 60
HALF_DAY_MINUTE = 11 * 60
FLEX_REQUIRED_MINUTES = 9 * 60
HALF_DAY_DURATION_MINUTES = 4 * 60
OT_BLOCK_MINUTES = 3 * 60 + 30


@dataclass(frozen=True)
class ParsedTime:
    minutes: int | None
    error: str | None = None


def parse_time(value: Any) -> ParsedTime:
    """Parse Excel-style HH.MM, text times, or native Excel time values."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ParsedTime(None, "Missing punch")
    if isinstance(value, datetime):
        return ParsedTime(value.hour * 60 + value.minute)
    if isinstance(value, time):
        return ParsedTime(value.hour * 60 + value.minute)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        numeric = float(value)
        if 0 <= numeric < 1:
            return ParsedTime(round(numeric * 24 * 60))
        hour = int(numeric)
        minute = round((numeric - hour) * 100)
        if 0 <= hour <= 23 and 0 <= minute <= 59:
            return ParsedTime(hour * 60 + minute)
        return ParsedTime(None, f"Invalid numeric time: {value}")

    text = str(value).strip()
    if not text:
        return ParsedTime(None, "Missing punch")
    text = re.sub(r"\.(\d{2})\.\d{2}$", r":\1", text)
    text = re.sub(r"^(\d{1,2})\.(\d{1,2})$", r"\1:\2", text)
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p"):
        try:
            parsed = datetime.strptime(text.upper(), fmt)
            return ParsedTime(parsed.hour * 60 + parsed.minute)
        except ValueError:
            pass
    return ParsedTime(None, f"Invalid time: {text}")


def format_time(minutes: int | None) -> str:
    if minutes is None:
        return "-"
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def parse_date(value: Any) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    return None if pd.isna(parsed) else parsed.date()


def normalize_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def find_header_row(raw: pd.DataFrame, required: set[str]) -> int | None:
    for index, row in raw.iterrows():
        values = {normalize_header(value) for value in row.tolist() if pd.notna(value)}
        if required.issubset(values):
            return int(index)
    return None


def load_workbook(uploaded_file: Any) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Find attendance and salary data from the supplied multi-sheet workbook."""
    workbook = pd.ExcelFile(uploaded_file)
    attendance = salary = None
    for sheet in workbook.sheet_names:
        raw = pd.read_excel(workbook, sheet_name=sheet, header=None)
        headers = find_header_row(raw, {"date", "employeenumber", "employeename"})
        if headers is not None and ({"intime", "outtime"} <= {normalize_header(x) for x in raw.iloc[headers].tolist()}):
            candidate = raw.iloc[headers + 1 :].copy()
            candidate.columns = [normalize_header(x) for x in raw.iloc[headers].tolist()]
            attendance = candidate
        salary_header = find_header_row(raw, {"employeenumber", "employeename", "monthlysalary"})
        if salary_header is not None:
            candidate = raw.iloc[salary_header + 1 :].copy()
            candidate.columns = [normalize_header(x) for x in raw.iloc[salary_header].tolist()]
            salary = candidate
    if attendance is None:
        raise ValueError("Could not find an attendance sheet with Date, Employee Number, In Time, and Out Time columns.")
    if salary is None:
        raise ValueError("Could not find a salary sheet with Employee Number, Employee Name, and Monthly Salary columns.")
    return attendance, salary


def calculate_daily(attendance: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for _, item in attendance.iterrows():
        employee_id = str(item.get("employeenumber", "")).strip()
        if not employee_id or employee_id.lower() == "nan":
            continue
        in_time = parse_time(item.get("intime"))
        out_time = parse_time(item.get("outtime"))
        work_date = parse_date(item.get("date"))
        errors = [message for message in (in_time.error, out_time.error) if message]
        if work_date is None:
            errors.append("Invalid date")
        duration = None
        if not errors and out_time.minutes is not None and in_time.minutes is not None:
            duration = out_time.minutes - in_time.minutes
            if duration < 0:
                errors.append("Punch-out is earlier than punch-in")
                duration = None

        valid = not errors
        half_day = valid and (in_time.minutes > HALF_DAY_MINUTE or duration <= HALF_DAY_DURATION_MINUTES)
        late_trigger = valid and (
            in_time.minutes > LATE_MINUTE
            or (FLEX_START_MINUTE < in_time.minutes <= LATE_MINUTE and duration < FLEX_REQUIRED_MINUTES)
        )
        # A half day is visible as a late/early event where relevant, but not counted as a late mark.
        late_mark = bool(late_trigger and not half_day)
        if valid:
            if in_time.minutes <= FLEX_START_MINUTE:
                early_threshold = REGULAR_END_MINUTE
            elif in_time.minutes <= LATE_MINUTE:
                early_threshold = in_time.minutes + 7 * 60
            else:
                early_threshold = 16 * 60
            early_leave = out_time.minutes < early_threshold
            overtime_minutes = max(0, out_time.minutes - REGULAR_END_MINUTE)
            overtime_days = (overtime_minutes // OT_BLOCK_MINUTES) * 0.5
        else:
            early_threshold = None
            early_leave = False
            overtime_minutes = 0
            overtime_days = 0.0

        rows.append(
            {
                "Date": work_date,
                "Employee ID": employee_id,
                "Employee Name": str(item.get("employeename", "")).strip(),
                "Punch In": format_time(in_time.minutes),
                "Punch Out": format_time(out_time.minutes),
                "Work Minutes": duration,
                "Work Hours": round(duration / 60, 2) if duration is not None else None,
                "Late Trigger": bool(late_trigger),
                "Late Mark": late_mark,
                "Half Day": bool(half_day),
                "Early Leave": bool(early_leave),
                "Early Leave Threshold": format_time(early_threshold),
                "OT Minutes": overtime_minutes,
                "OT Days": overtime_days,
                "Status": "Needs review" if errors else "Calculated",
                "Validation Notes": "; ".join(errors),
            }
        )
    return pd.DataFrame(rows)


def calculate_payroll(daily: pd.DataFrame, salaries: pd.DataFrame) -> pd.DataFrame:
    salary_rows = salaries.copy()
    salary_rows["Employee ID"] = salary_rows["employeenumber"].astype(str).str.strip()
    salary_rows["Employee Name"] = salary_rows["employeename"].astype(str).str.strip()
    salary_rows["Monthly Salary"] = pd.to_numeric(salary_rows["monthlysalary"], errors="coerce")
    salary_rows = salary_rows[["Employee ID", "Employee Name", "Monthly Salary"]].dropna(subset=["Monthly Salary"])

    if daily.empty:
        return pd.DataFrame()
    summary = daily.groupby("Employee ID", as_index=False).agg(
        Attendance_Days=("Date", "count"),
        Late_Marks=("Late Mark", "sum"),
        Half_Days=("Half Day", "sum"),
        Early_Leaves=("Early Leave", "sum"),
        OT_Days=("OT Days", "sum"),
        Review_Records=("Status", lambda values: (values == "Needs review").sum()),
    )
    result = salary_rows.merge(summary, on="Employee ID", how="left")
    for column in ("Attendance_Days", "Late_Marks", "Half_Days", "Early_Leaves", "OT_Days", "Review_Records"):
        result[column] = result[column].fillna(0)
    result["Excess Late Marks"] = (result["Late_Marks"] - LATE_MARK_LIMIT).clip(lower=0)
    result["Excess Early Leaves"] = (result["Early_Leaves"] - EARLY_LEAVE_LIMIT).clip(lower=0)
    result["Day Pay"] = result["Monthly Salary"] / DAY_DIVISOR
    result["Half-Day Deduction"] = result["Half_Days"] * 0.5 * result["Day Pay"]
    result["Late-Mark Deduction"] = result["Excess Late Marks"] * 0.5 * result["Day Pay"]
    result["OT Pay"] = result["OT_Days"] * result["Day Pay"]
    result["Total Deductions"] = result["Half-Day Deduction"] + result["Late-Mark Deduction"]
    result["Payable Salary"] = (result["Monthly Salary"] - result["Total Deductions"] + result["OT Pay"]).round(0)
    result["Payroll Status"] = result["Review_Records"].map(lambda count: "Needs review" if count else "Ready")
    return result.rename(
        columns={
            "Attendance_Days": "Attendance Days",
            "Late_Marks": "Late Marks",
            "Half_Days": "Half Days",
            "Early_Leaves": "Early Leaves",
            "OT_Days": "OT Days",
            "Review_Records": "Review Records",
        }
    ).sort_values("Employee ID")


def currency(value: float) -> str:
    return f"₹{value:,.0f}"
