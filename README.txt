PAYROLL MASTER APPLICATION - MULTIPLE EMPLOYEES

Install
1. Python 3.10+ (with Tkinter - included on Windows/macOS; on Ubuntu: sudo apt install python3-tk)
2. pip install -r requirements.txt
3. Run: ./run_payroll_app.sh   (Windows: run_payroll_app.bat, or: python payroll_app.py)

Files
- payroll_app.py            GUI (fast grid, pop-ups), Excel open/save, PDF
- payroll_calc.py           all payroll maths (no GUI) - this is what the tests cover
- test_payroll_calc.py      run:  python3 -m unittest -v
- Payroll_Reference.xlsx    source workbook (salary sheets, employee master, ActiveMembers_<date>)
- Payroll_Data.xlsx         working file written by SAVE TO EXCEL
- last_session.json         remembers the last Excel file/tab you opened (created automatically; safe to delete)
- fonts/                    DejaVu Sans, bundled so the PDF can print the rupee sign
- backups/                  automatic copy of Payroll_Data.xlsx before every save (last 10 kept)

How a month works
1. On start-up a window asks WHICH EXCEL FILE and WHICH TAB (sheet) to open.
   Pick a saved payroll tab ("Payroll Data" or a month such as "Payroll October 2026"), the employee master,
   or "All employees" to build a fresh sheet from the master + salary sheets. Excel menu > "Open Excel / Tab" (Ctrl+O) reopens this window. The tab name is used as the company name (it is filled in for
   employees that have no company yet and printed on the PDF; there is no Company column on screen).
2. Check Payroll Month and the settings bar (click "More settings" for ESI / PF options). Settings are saved in the workbook.
3. Click a white cell and type. ENTER (or Tab) saves the cell and moves to the NEXT cell to the right; at the end of a row it
   jumps to the first cell of the next row. Shift+Tab goes back, Up/Down arrows move within the column, Esc cancels.
   Drop-down cells (OT Method, ESI Exempt): press Enter / click to choose, Space opens the list.
4. Working Days, Paid Holidays and Paid Days are whole numbers (a typed 24.5 is rounded to 25).
5. ADVANCE: click "[ Advance ]" on an employee row (or select the row and press the ADVANCE button). A pop-up holds, for
   General Advance and Festival Advance: previous balance, advance given, recovery, defer to next month, amount deferred
   to next month, and balance. The grid shows only the total Advance Deduction and Advance Balance.
6. Rows turn red when something needs attention (negative net salary, negative advance balance) and yellow for warnings.
   CALCULATE ALL lists every warning.
7. Excel menu > Save to Excel (Ctrl+S) updates sheet "Payroll Data" and also keeps a sheet per month ("Payroll October 2026").
8. PDF / Print menu > Generate PDF or Print: employees with no salary entered are left out; totals are added.

Calculation rules (all configurable in the settings bar)
- Pay is pro-rated: (Working Days + Paid Holidays, max Company Working Days) / Company Working Days.
  Basic, Allowances and Other Earnings are the FULL-MONTH amounts. Untick the box to pay in full.
- PF = 12% of earned basic, capped at the PF wage ceiling (15,000; set 0 for no cap).
  EPF part / EPS part (3.67% / 8.33%) are a split of that PF amount for reporting only.
- ESI = ESI rate (0.75%) x (gross - OT - allowances). Not charged if the employee is marked
  "ESI Exempt", or if wages (gross excluding OT) exceed the ESI wage ceiling (21,000).
  Allowances are excluded from the base to match the supplied salary sheets; untick to include them.
- Deductions are rounded to whole rupees (matches the sheets); untick for paise.
- OT methods: Hours x OT Rate | Hours x OT Rate x Multiplier | Basic / Company Days / Hours per Day x Hours x Multiplier.
- Net = Gross - (PF + ESI + Festival Recovery + General Recovery + Other Deductions); deferred recoveries are not deducted.
- Advance Balance (General and Festival, each) = Previous + Given - Recovery. It is shown as is (a negative value is flagged, not hidden).
- Defer = the recovery is NOT deducted this month; the amount is shown as "Deferred to next month" and the balance stays higher.
Please confirm these rates and ceilings against current statutory rules before relying on them.

Notes
- Legacy .xls import needs LibreOffice. The re-saved workbook may show formulas without values in
  some viewers until opened in Excel (openpyxl does not store calculated results).
- Payroll_Data.xlsx and Payroll_Reference.xlsx contain employee names and pay. See .gitignore.
