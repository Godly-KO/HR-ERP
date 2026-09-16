"""PDF payslip creation."""

from __future__ import annotations

from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

from payroll import currency


def create_payslip(row, pay_period: str) -> bytes:
    output = BytesIO()
    document = SimpleDocTemplate(output, pagesize=A4, rightMargin=20 * mm, leftMargin=20 * mm, topMargin=18 * mm)
    styles = getSampleStyleSheet()
    story = [Paragraph("Square One Communications", styles["Title"]), Paragraph("Salary Slip", styles["Heading2"]), Spacer(1, 8 * mm)]
    details = [
        ["Employee name", str(row["Employee Name"])],
        ["Employee ID", str(row["Employee ID"])],
        ["Pay period", pay_period],
        ["Payroll status", str(row["Payroll Status"])],
    ]
    calculation = [
        ["Monthly salary", currency(row["Monthly Salary"])],
        ["Half-day deduction", f"({currency(row['Half-Day Deduction'])})"],
        ["Late-mark deduction", f"({currency(row['Late-Mark Deduction'])})"],
        ["Overtime pay", currency(row["OT Pay"])],
        ["Payable salary", currency(row["Payable Salary"])],
    ]
    for data in (details, calculation):
        table = Table(data, colWidths=[65 * mm, 85 * mm])
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#0F172A")),
            ("TEXTCOLOR", (0, 0), (0, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("PADDING", (0, 0), (-1, -1), 8),
        ]))
        story.extend([table, Spacer(1, 7 * mm)])
    story.append(Paragraph("This payslip is generated from the uploaded attendance data. Records needing review should be verified before payment.", styles["BodyText"]))
    document.build(story)
    return output.getvalue()
