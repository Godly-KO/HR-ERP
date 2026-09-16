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
overview, employees, quality, payslips = st.tabs(["Payroll overview", "Employee details", "Data quality", "Payslips"])

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

