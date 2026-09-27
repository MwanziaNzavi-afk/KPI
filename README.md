# KPI Bonus Control Centre

## What this does
- Imports the KPI structures from the supplied F&B and Service Excel workbooks.
- Uploads daily sales from Excel/CSV.
- Matches sales rows to the master employee roster.
- Stores daily KPI scores and calculates weighted KPI performance.
- Captures employee acknowledgement or correction requests after feedback is shared.
- Creates next-day coaching follow-ups, owned by a manager or rotational supervisor, with recorded outcomes.
- Accepts supervisor remarks, with a Google Forms/Google Sheets integration path.
- Exports a management-ready Excel workbook with Dashboard, KPI Detail, Sales, Remarks, employee reviews, coaching follow-ups and Charts.

## Daily operating flow
1. A supervisor enters or syncs each team member's measurable KPI scores.
2. The employee reviews the result in **Employee review**, acknowledges it or requests a correction.
3. Log any improvement commitment in **Next-day follow-up**; the assigned owner records the next shift's outcome.
4. Use **Dashboard** for an at-a-glance view and **Excel Export** for the director's report.

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Google remarks
1. Open a blank Google Sheet.
2. Extensions → Apps Script.
3. Paste `google_apps_script.gs`.
4. Run `setupKpiRemarks()` once and authorize it.
5. The script creates a Google Form and supervisor-specific prefilled links.
6. Publish/share the response sheet as needed. For the Python app, publish the response sheet as CSV and set:
   `GOOGLE_SHEET_CSV_URL=<published CSV URL>`
7. The form responses should contain: Timestamp, Work Date, Group, Team Member, Supervisor, Remark.

For strict access control, use Google Workspace accounts and restrict the form to the organization's users.
