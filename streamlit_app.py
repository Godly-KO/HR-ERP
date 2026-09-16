from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st

from payroll import DAY_DIVISOR, EARLY_LEAVE_LIMIT, LATE_MARK_LIMIT, calculate_daily, calculate_payroll, currency, load_workbook
from payslip import create_payslip


st.set_page_config(page_title="HR ERP | Payroll", page_icon="💼", layout="wide")


def export_payroll(summary: pd.DataFrame, daily: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Payroll Summary", index=False)
        daily.to_excel(writer, sheet_name="Daily Attendance", index=False)
    return output.getvalue()


def money_columns(frame: pd.DataFrame) -> dict[str, st.column_config.NumberColumn]:
    names = {"Monthly Salary", "Day Pay", "Half-Day Deduction", "Late-Mark Deduction", "OT Pay", "Total Deductions", "Payable Salary"}
    return {name: st.column_config.NumberColumn(name, format="₹%d") for name in names if name in frame.columns}


def build_action_queue(daily: pd.DataFrame, payroll: pd.DataFrame) -> pd.DataFrame:
    """Build a small, action-oriented queue rather than another raw exception table."""
    actions: list[dict[str, str | int]] = []
    invalid = daily[daily["Status"] == "Needs review"]
    if not invalid.empty:
        actions.append({
            "Priority": "Critical",
            "Action": "Correct missing or invalid punches",
            "Affected employees": int(invalid["Employee ID"].nunique()),
            "Details": f"{len(invalid)} attendance record(s) are excluded from automated rules.",
        })
    duplicates = daily[daily.duplicated(["Employee ID", "Date"], keep=False)]
    if not duplicates.empty:
        actions.append({
            "Priority": "Warning",
            "Action": "Resolve duplicate attendance records",
            "Affected employees": int(duplicates["Employee ID"].nunique()),
            "Details": f"{len(duplicates)} duplicate row(s) need verification.",
        })
    review_count = int((payroll["Payroll Status"] == "Needs review").sum())
    if review_count:
        actions.append({
            "Priority": "Warning",
            "Action": "Review affected payrolls",
            "Affected employees": review_count,
            "Details": "Resolve attendance exceptions before approving payment.",
        })
    ready_count = int((payroll["Payroll Status"] == "Ready").sum())
    if ready_count:
        actions.append({
            "Priority": "Ready",
            "Action": "Approve ready payrolls and issue payslips",
            "Affected employees": ready_count,
            "Details": "Salary slips can be generated from the Payslips tab after approval.",
        })
    return pd.DataFrame(actions)


st.title("HR ERP")
st.caption("Attendance-led monthly payroll and salary-slip generation")

with st.sidebar:
    st.header("Monthly upload")
    upload = st.file_uploader("Attendance and salary workbook", type=["xlsx"], help="Upload the workbook containing attendance and Salary Data sheets.")
    st.divider()
    st.header("Rules applied")
    st.markdown(
        f"""
        - Day pay: monthly salary ÷ {DAY_DIVISOR}
        - {LATE_MARK_LIMIT} late marks allowed; each later mark deducts 0.5 day
        - Check-in after 11:00 or 4 hours or less: half day
        - {EARLY_LEAVE_LIMIT} early leaves allowed; excess is flagged, not deducted
        - OT: each completed 3.5 hours after 18:00 earns 0.5 day pay
        """
    )

if not upload:
    st.info("Upload the monthly Excel workbook to calculate payroll. No employee data is stored by this app.")
    st.markdown("**Expected columns:** attendance needs Date, Employee Number, Employee Name, In Time, and Out Time. Salary data needs Employee Number, Employee Name, and Monthly Salary.")
    st.stop()

try:
    attendance, salaries = load_workbook(upload)
    daily = calculate_daily(attendance)
    payroll = calculate_payroll(daily, salaries)
except Exception as error:
    st.error(f"The workbook could not be processed: {error}")
    st.stop()

if daily.empty or payroll.empty:
    st.error("No usable attendance or salary rows were found.")
    st.stop()

valid_dates = [value for value in daily["Date"].dropna()]
pay_period = pd.Timestamp(max(valid_dates)).strftime("%B %Y") if valid_dates else "Current period"
dashboard, overview, employees, quality, payslips = st.tabs(["Dashboard", "Payroll overview", "Employee details", "Data quality", "Payslips"])

with dashboard:
    latest_date = max(valid_dates)
    latest_day = daily[daily["Date"] == latest_date].copy()
    roster_size = len(payroll)
    recorded_employee_count = latest_day["Employee ID"].nunique()
    attendance_rate = (recorded_employee_count / roster_size * 100) if roster_size else 0
    daily_late = int(latest_day["Late Mark"].sum())
    daily_half = int(latest_day["Half Day"].sum())
    daily_early = int(latest_day["Early Leave"].sum())
    daily_review = int((latest_day["Status"] == "Needs review").sum())
    daily_ot = int((latest_day["OT Days"] > 0).sum())

    st.subheader(f"Daily operations - {pd.Timestamp(latest_date).strftime('%d %b %Y')}")
    st.caption("The latest attendance date in the uploaded workbook. Attendance rate measures recorded employees against the payroll roster.")
    metrics = st.columns(6)
    metrics[0].metric("Recorded", f"{recorded_employee_count}/{roster_size}", f"{attendance_rate:.0f}%")
    metrics[1].metric("Late arrivals", daily_late)
    metrics[2].metric("Half days", daily_half)
    metrics[3].metric("Early leaves", daily_early)
    metrics[4].metric("Needs review", daily_review)
    metrics[5].metric("OT eligible", daily_ot)

    st.subheader("Daily exceptions")
    daily_exceptions = latest_day[
        (latest_day["Late Trigger"])
        | (latest_day["Half Day"])
        | (latest_day["Early Leave"])
        | (latest_day["OT Days"] > 0)
        | (latest_day["Status"] == "Needs review")
    ]
    if daily_exceptions.empty:
        st.success("No attendance exceptions for the latest uploaded day.")
    else:
        st.dataframe(
            daily_exceptions[["Employee ID", "Employee Name", "Punch In", "Punch Out", "Late Mark", "Half Day", "Early Leave", "OT Days", "Status", "Validation Notes"]],
            use_container_width=True,
            hide_index=True,
        )

    st.subheader("Monthly attendance trends")
    daily_trends = daily.groupby("Date", as_index=False).agg(
        **{
            "Recorded attendance": ("Employee ID", "nunique"),
            "Late marks": ("Late Mark", "sum"),
            "Half days": ("Half Day", "sum"),
            "Early leaves": ("Early Leave", "sum"),
            "Records needing review": ("Status", lambda values: (values == "Needs review").sum()),
        }
    ).sort_values("Date")
    trend_left, trend_right = st.columns(2)
    with trend_left:
        st.caption("Recorded attendance and exceptions by date")
        st.line_chart(daily_trends.set_index("Date")[["Recorded attendance", "Late marks", "Half days", "Early leaves"]])
    with trend_right:
        st.caption("Records requiring HR review")
        st.bar_chart(daily_trends.set_index("Date")[["Records needing review"]])

    st.subheader("Monthly payroll overview")
    payroll_metrics = st.columns(4)
    payroll_metrics[0].metric("Base payroll", currency(payroll["Monthly Salary"].sum()))
    payroll_metrics[1].metric("Deductions", currency(payroll["Total Deductions"].sum()))
    payroll_metrics[2].metric("OT cost", currency(payroll["OT Pay"].sum()))
    payroll_metrics[3].metric("Final payable", currency(payroll["Payable Salary"].sum()))
    payroll_chart = payroll[["Employee Name", "Monthly Salary", "Total Deductions", "OT Pay", "Payable Salary"]].set_index("Employee Name")
    st.bar_chart(payroll_chart)

    insight_left, insight_right = st.columns(2)
    with insight_left:
        st.subheader("Employee insights")
        top_ot = payroll[payroll["OT Days"] > 0].sort_values(["OT Pay", "OT Days"], ascending=False)
        if top_ot.empty:
            st.info("No employee earned overtime in this upload.")
        else:
            st.caption("Highest overtime earners")
            st.dataframe(top_ot[["Employee ID", "Employee Name", "OT Days", "OT Pay"]].head(10), use_container_width=True, hide_index=True, column_config=money_columns(top_ot))
        late_risk = payroll[payroll["Late Marks"] >= LATE_MARK_LIMIT - 1].sort_values("Late Marks", ascending=False)
        if not late_risk.empty:
            st.caption("Employees nearing or exceeding the late-mark allowance")
            st.dataframe(late_risk[["Employee ID", "Employee Name", "Late Marks", "Excess Late Marks"]], use_container_width=True, hide_index=True)
    with insight_right:
        st.subheader("Repeated attendance issues")
        repeat_issues = daily[daily["Status"] == "Needs review"].groupby(["Employee ID", "Employee Name"], as_index=False).size().rename(columns={"size": "Missing or invalid punches"}).sort_values("Missing or invalid punches", ascending=False)
        if repeat_issues.empty:
            st.success("No missing or invalid punches in this upload.")
        else:
            st.dataframe(repeat_issues, use_container_width=True, hide_index=True)
        st.caption("Department/team comparisons are unavailable because the uploaded workbook has no department field.")

    st.subheader("Action queue")
    actions = build_action_queue(daily, payroll)
    if actions.empty:
        st.success("No payroll actions are pending.")
    else:
        st.dataframe(actions, use_container_width=True, hide_index=True)

with overview:
    total_payable = payroll["Payable Salary"].sum()
    total_deductions = payroll["Total Deductions"].sum()
    total_ot = payroll["OT Pay"].sum()
    needs_review = (payroll["Payroll Status"] == "Needs review").sum()
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Employees", len(payroll))
    col2.metric("Payable payroll", currency(total_payable))
    col3.metric("Deductions", currency(total_deductions))
    col4.metric("Needs review", int(needs_review))
    st.subheader(f"Payroll summary - {pay_period}")
    st.dataframe(payroll, use_container_width=True, hide_index=True, column_config=money_columns(payroll))
    st.download_button("Download payroll workbook", export_payroll(payroll, daily), "payroll_results.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with employees:
    employee_options = payroll.apply(lambda row: f"{row['Employee ID']} - {row['Employee Name']}", axis=1).tolist()
    selected = st.selectbox("Employee", employee_options)
    employee_id = selected.split(" - ", 1)[0]
    employee_summary = payroll.loc[payroll["Employee ID"] == employee_id].iloc[0]
    st.subheader(f"{employee_summary['Employee Name']} - {pay_period}")
    metrics = st.columns(5)
    metrics[0].metric("Payable", currency(employee_summary["Payable Salary"]))
    metrics[1].metric("Late marks", int(employee_summary["Late Marks"]))
    metrics[2].metric("Half days", int(employee_summary["Half Days"]))
    metrics[3].metric("Early leaves", int(employee_summary["Early Leaves"]))
    metrics[4].metric("OT days", f"{employee_summary['OT Days']:.1f}")
    employee_daily = daily[daily["Employee ID"] == employee_id].sort_values("Date")
    st.dataframe(employee_daily, use_container_width=True, hide_index=True)

with quality:
    issues = daily[daily["Status"] == "Needs review"]
    duplicates = daily[daily.duplicated(["Employee ID", "Date"], keep=False)].sort_values(["Employee ID", "Date"])
    st.subheader("Records needing attention")
    if issues.empty:
        st.success("No missing or invalid punches found.")
    else:
        st.warning(f"{len(issues)} record(s) were excluded from attendance rules until corrected.")
        st.dataframe(issues, use_container_width=True, hide_index=True)
    if not duplicates.empty:
        st.warning("Duplicate employee/date records found. Verify them before approving payroll.")
        st.dataframe(duplicates, use_container_width=True, hide_index=True)

with payslips:
    st.subheader("Individual salary slip")
    selected = st.selectbox("Employee for payslip", employee_options, key="payslip_employee")
    employee_id = selected.split(" - ", 1)[0]
    employee_summary = payroll.loc[payroll["Employee ID"] == employee_id].iloc[0]
    if employee_summary["Payroll Status"] == "Needs review":
        st.warning("This employee has attendance records needing review. Verify them before issuing the slip.")
    slip = create_payslip(employee_summary, pay_period)
    safe_id = employee_id.replace("/", "-")
    st.download_button("Download PDF salary slip", slip, f"payslip_{safe_id}.pdf", "application/pdf")
