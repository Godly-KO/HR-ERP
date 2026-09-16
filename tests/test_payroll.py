import unittest

import pandas as pd

from payroll import calculate_daily, calculate_payroll, parse_time


class PayrollRulesTests(unittest.TestCase):
    def test_excel_style_time_is_not_decimal_hours(self):
        self.assertEqual(parse_time(10.31).minutes, 10 * 60 + 31)
        self.assertEqual(parse_time("19.00.00").minutes, 19 * 60)

    def test_late_mark_after_10_and_half_day_after_11(self):
        rows = pd.DataFrame([
            {"date": "01 Aug 2026", "employeenumber": "E1", "employeename": "A", "intime": 10.01, "outtime": 19.01},
            {"date": "02 Aug 2026", "employeenumber": "E1", "employeename": "A", "intime": 11.01, "outtime": 19.01},
        ])
        daily = calculate_daily(rows)
        self.assertTrue(daily.iloc[0]["Late Mark"])
        self.assertTrue(daily.iloc[1]["Half Day"])
        self.assertFalse(daily.iloc[1]["Late Mark"])

    def test_ot_is_counted_in_completed_three_and_half_hour_blocks(self):
        rows = pd.DataFrame([
            {"date": "01 Aug 2026", "employeenumber": "E1", "employeename": "A", "intime": 9.00, "outtime": 21.30},
        ])
        daily = calculate_daily(rows)
        self.assertEqual(daily.iloc[0]["OT Days"], 0.5)

    def test_fifth_late_mark_deducts_half_day(self):
        rows = pd.DataFrame([
            {"date": f"0{day} Aug 2026", "employeenumber": "E1", "employeename": "A", "intime": 10.01, "outtime": 19.01}
            for day in range(1, 6)
        ])
        salary = pd.DataFrame([{"employeenumber": "E1", "employeename": "A", "monthlysalary": 30000}])
        payroll = calculate_payroll(calculate_daily(rows), salary)
        self.assertEqual(payroll.iloc[0]["Late-Mark Deduction"], 500)
        self.assertEqual(payroll.iloc[0]["Payable Salary"], 29500)


if __name__ == "__main__":
    unittest.main()
