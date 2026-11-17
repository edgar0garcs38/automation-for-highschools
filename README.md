# Accounts Receivable Analyzer for Schools

A desktop app that helps a school's finance team track **overdue tuition and fees**: who owes what, in which program, and since which month. It produces formatted Excel reports ready to share with each department.

## Features
- **Two analysis modes** from the start menu: *pending fees* and *overdue portfolio*. Each accepts its own export format, with automatic column-name matching (accents, spacing, aliases).
- **Classification:** each charge is labeled as enrollment, tuition installment, registration or other, and each student as regular program or validation course
- **Filters** by year, program, month, fee type and modality
- **KPI summary** and pivot tables of amounts owed and number of students per program and month
- **Interactive charts** (bar, line and 3D) embedded in the app, switchable between *amount owed* and *number of students*
- **Excel export** built with openpyxl:
  - Styled summary, amounts and students sheets
  - A per-program validation workbook with one detail sheet per program (including student contact phones) and a summary sheet linked by formulas

## Tech stack
Python · Tkinter · pandas · NumPy · matplotlib · openpyxl

## Run it
```bash
pip install -r requirements.txt
python RubrosPendientesITQ.py
```
Then pick the analysis type and load the receivables Excel export from the school's system.

> **Note:** Student and financial data is confidential and is not included in this repository.
