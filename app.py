
import io, os, re, sqlite3
from datetime import date
import pandas as pd
import streamlit as st

DB = os.getenv("KPI_DB", "kpi_bonus.db")
GOOGLE_SHEET_CSV_URL = os.getenv("GOOGLE_SHEET_CSV_URL", "")
GOOGLE_KPI_SHEET_URL = os.getenv(
    "GOOGLE_KPI_SHEET_URL",
    "https://docs.google.com/spreadsheets/d/1htrldZzYaeRUKtGxLcfJ7IHt2uM-1S3meRORdksrAiQ/export?format=xlsx",
)
ROSTER_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "roster.csv")

SERVICE_KPIS = {
    "Discipline": .15, "Etiquette": .10, "Hygiene": .05,
    "Staff Responsiveness": .05, "Absentism": .15, "Communication": .10,
    "Customer Service": .20, "Handling Customer Complaints": .05,
    "Teamwork": .10, "Service Sequence": .05
}
FB_KPIS = {
    "Personal Hygiene": .10, "Customer Service": .20, "Order Accuracy": .10,
    "Product Knowledge": .10, "Upselling Rate": .05, "Teamwork/Collaboration": .10,
    "Punctuality": .10, "Sales": .15, "Display & Fridge Restock & Refill": .10
}
BAR_KPIS = {
    "Bar Cleanliness & Set Up": .15, "Guest Engagement & Service": .175,
    "Stock Control & Requisition Accuracy": .075, "Cocktail Quality & Consistency": .075,
    "Shift Attendance and Responsibility": .125, "Teamwork & Support": .20,
    "Upselling & Promotion Execution": .075, "Bar Inventory & Waste Control": .125
}

def norm(x):
    return re.sub(r"[^a-z0-9]+", "", str(x).lower()) if x is not None else ""

def db():
    con=sqlite3.connect(DB)
    con.execute("""CREATE TABLE IF NOT EXISTS roster(
        name TEXT PRIMARY KEY, department TEXT, grp TEXT, supervisor TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS sales(
        work_date TEXT, name TEXT, sales REAL, source TEXT,
        PRIMARY KEY(work_date,name))""")
    con.execute("""CREATE TABLE IF NOT EXISTS kpi_scores(
        work_date TEXT, name TEXT, kpi TEXT, score REAL,
        PRIMARY KEY(work_date,name,kpi))""")
    con.execute("""CREATE TABLE IF NOT EXISTS remarks(
        work_date TEXT, grp TEXT, name TEXT, supervisor TEXT, remark TEXT,
        PRIMARY KEY(work_date,grp,name))""")
    # Employees can acknowledge feedback or ask for a score to be reviewed.
    # Follow-ups are deliberately separate: they are actions to be checked the next day.
    con.execute("""CREATE TABLE IF NOT EXISTS feedback_reviews(
        work_date TEXT, name TEXT, status TEXT, employee_note TEXT, updated_at TEXT,
        PRIMARY KEY(work_date,name))""")
    con.execute("""CREATE TABLE IF NOT EXISTS follow_ups(
        follow_up_date TEXT, source_date TEXT, name TEXT, area TEXT, action TEXT,
        owner TEXT, status TEXT, outcome TEXT,
        PRIMARY KEY(follow_up_date,name,area))""")
    con.commit(); return con

def seed_roster_if_empty():
    """Initialize a new database from the bundled roster file."""
    con = db()
    has_roster = con.execute("SELECT 1 FROM roster LIMIT 1").fetchone()
    if has_roster or not os.path.exists(ROSTER_CSV):
        return
    roster = pd.read_csv(ROSTER_CSV).rename(columns={"group": "grp"})
    required = {"name", "department", "grp", "supervisor"}
    if not required.issubset(roster.columns):
        raise ValueError("roster.csv must contain name, department, group, and supervisor columns.")
    roster = roster[["name", "department", "grp", "supervisor"]].fillna("")
    con.executemany(
        "INSERT OR REPLACE INTO roster(name, department, grp, supervisor) VALUES(?,?,?,?)",
        roster.itertuples(index=False, name=None),
    )
    con.commit()

def load_roster():
    seed_roster_if_empty()
    con=db()
    return pd.read_sql_query("SELECT * FROM roster ORDER BY department, grp, name",con)

def kpi_map(dept, grp=""):
    g=(grp or "").upper()
    if "BAR" in g or "MIXOLOGIST" in g or "SOMMELLIER" in g or "BARTENDER" in g:
        return BAR_KPIS
    if dept.lower()=="service":
        return SERVICE_KPIS
    return FB_KPIS

def weighted_kpi_score(scores, name, department, grp):
    """Return a safe weighted KPI score when a date has partial or no entries."""
    kpis = kpi_map(department, grp)
    if scores.empty or not {"name", "kpi", "score"}.issubset(scores.columns):
        return 0.0
    employee_scores = scores[scores["name"] == name]
    return sum(
        float(employee_scores.loc[employee_scores["kpi"] == kpi, "score"].iloc[0]) * weight
        if not employee_scores.loc[employee_scores["kpi"] == kpi, "score"].empty else 0.0
        for kpi, weight in kpis.items()
    )

def daily_performance(work_date):
    """One reporting row per employee, including the feedback lifecycle."""
    roster = load_roster()
    con = db()
    scores = pd.read_sql_query("SELECT * FROM kpi_scores WHERE work_date=?", con, params=(str(work_date),))
    sales = pd.read_sql_query("SELECT * FROM sales WHERE work_date=?", con, params=(str(work_date),))
    reviews = pd.read_sql_query("SELECT name, status FROM feedback_reviews WHERE work_date=?", con, params=(str(work_date),))
    status_by_name = dict(zip(reviews["name"], reviews["status"])) if not reviews.empty else {}
    rows = []
    for _, person in roster.iterrows():
        employee_sales = sales[sales["name"] == person["name"]]
        sale = float(employee_sales["sales"].iloc[0]) if not employee_sales.empty else 0.0
        rows.append({
            "Name": person["name"], "Department": person["department"], "Group": person["grp"],
            "Supervisor": person["supervisor"], "Sales": sale,
            "KPI Score": weighted_kpi_score(scores, person["name"], person["department"], person["grp"]),
            "Feedback status": status_by_name.get(person["name"], "Awaiting review"),
        })
    return pd.DataFrame(rows)

def import_sales(upload, work_date):
    if upload.name.lower().endswith(".csv"):
        df=pd.read_csv(upload)
    else:
        xls=pd.ExcelFile(upload)
        frames=[pd.read_excel(upload,sheet_name=s) for s in xls.sheet_names]
        df=pd.concat(frames,ignore_index=True)
    df.columns=[str(c).strip() for c in df.columns]
    name_col=next((c for c in df.columns if norm(c) in {"name","employee","employeename","staff","staffname"}),None)
    sales_col=next((c for c in df.columns if norm(c) in {"sales","individualsales","amount","revenue","netsales","total_sales"}),None)
    if not name_col or not sales_col:
        raise ValueError("Could not identify Name and Sales columns. Expected columns such as Name and Sales.")
    df=df[[name_col,sales_col]].rename(columns={name_col:"name",sales_col:"sales"})
    df["sales"]=pd.to_numeric(df["sales"],errors="coerce").fillna(0)
    df["name_key"]=df["name"].map(norm)
    roster=load_roster()
    lookup={norm(n):n for n in roster["name"]}
    df["matched_name"]=df["name_key"].map(lookup)
    con=db()
    matched=df[df["matched_name"].notna()].copy()
    for _,r in matched.iterrows():
        con.execute("INSERT OR REPLACE INTO sales VALUES(?,?,?,?)",
                    (str(work_date),r["matched_name"],float(r["sales"]),upload.name))
    con.commit()
    return df, matched

def pull_google_remarks():
    if not GOOGLE_SHEET_CSV_URL:
        return pd.DataFrame(columns=["work_date","grp","name","supervisor","remark"])
    try:
        remarks = pd.read_csv(GOOGLE_SHEET_CSV_URL)
        field_map = {
            "workdate": "work_date", "date": "work_date",
            "group": "grp", "team": "grp",
            "teammember": "name", "employee": "name", "employeename": "name",
            "supervisor": "supervisor", "remark": "remark", "remarks": "remark",
        }
        remarks = remarks.rename(columns={
            column: field_map[norm(column)]
            for column in remarks.columns if norm(column) in field_map
        })
        return remarks.reindex(columns=["work_date", "grp", "name", "supervisor", "remark"])
    except Exception as e:
        st.warning(f"Google remarks feed could not be read: {e}")
        return pd.DataFrame(columns=["work_date","grp","name","supervisor","remark"])

def sync_google_kpis(work_date):
    """Import KPI scores from the shared Google Sheet into the selected date."""
    try:
        sheets = pd.read_excel(GOOGLE_KPI_SHEET_URL, sheet_name=None, header=None)
    except Exception as e:
        raise ValueError(f"Google KPI sheet could not be read: {e}") from e

    roster = load_roster()
    roster_lookup = {norm(name): name for name in roster["name"]}
    kpi_by_norm = {norm(kpi): kpi for kpi in SERVICE_KPIS | FB_KPIS | BAR_KPIS}
    rows = []
    for sheet_name, sheet in sheets.items():
        header_candidates = sheet.index[sheet.iloc[:, 0].map(norm).eq("name")]
        if len(header_candidates) == 0:
            continue
        header_row = header_candidates[0]
        headers = sheet.iloc[header_row].tolist()
        kpi_columns = {
            idx: kpi_by_norm[norm(label)]
            for idx, label in enumerate(headers)
            if norm(label) in kpi_by_norm
        }
        for _, source_row in sheet.iloc[header_row + 1:].iterrows():
            source_name = source_row.iloc[0]
            name = roster_lookup.get(norm(source_name))
            if not name:
                continue
            for column, kpi in kpi_columns.items():
                score = pd.to_numeric(source_row.iloc[column], errors="coerce")
                if pd.notna(score):
                    rows.append((str(work_date), name, kpi, float(score)))
    if not rows:
        return 0, 0
    con = db()
    con.executemany("INSERT OR REPLACE INTO kpi_scores VALUES(?,?,?,?)", rows)
    con.commit()
    return len(rows), len({row[1] for row in rows})

def export_excel(work_date):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.formatting.rule import ColorScaleRule
    con=db()
    roster=load_roster()
    sales=pd.read_sql_query("SELECT * FROM sales WHERE work_date=?", con=con, params=(str(work_date),))
    scores=pd.read_sql_query("SELECT * FROM kpi_scores WHERE work_date=?", con=con, params=(str(work_date),))
    remarks=pull_google_remarks()
    if not remarks.empty:
        remarks=remarks[remarks["work_date"].astype(str)==str(work_date)]
    wb=Workbook()
    ws=wb.active; ws.title="Dashboard"
    ws["A1"]="KPI BONUS DASHBOARD"; ws["A1"].font=Font(size=20,bold=True)
    ws["A2"]="Date"; ws["B2"]=str(work_date)
    headers=["Name","Department","Group","Supervisor","Sales","KPI Score %","KPI Bonus","Remarks"]
    for c,h in enumerate(headers,1):
        ws.cell(4,c,h).font=Font(bold=True)
    merged=roster.merge(sales,on="name",how="left").merge(
        scores.pivot_table(index="name",columns="kpi",values="score",aggfunc="mean").reset_index()
        if not scores.empty else pd.DataFrame({"name":[]} ),on="name",how="left")
    # Compute KPI score from stored score values and weight maps.
    rows=[]
    for _,r in roster.iterrows():
        total=weighted_kpi_score(scores, r["name"], r["department"], r["grp"])
        srow=sales[sales["name"]==r["name"]]
        sale=float(srow["sales"].iloc[0]) if len(srow) else 0
        rr=remarks[remarks["name"]==r["name"]] if not remarks.empty and "name" in remarks else pd.DataFrame()
        remark=str(rr["remark"].iloc[-1]) if len(rr) else ""
        rows.append([r["name"],r["department"],r["grp"],r["supervisor"],sale,total,round(sale*0,2),remark])
    for ridx,row in enumerate(rows,5):
        for c,val in enumerate(row,1): ws.cell(ridx,c,val)
    ws.freeze_panes="A5"; ws.auto_filter.ref=f"A4:H{max(4,len(rows)+4)}"
    widths=[24,14,22,20,14,14,14,45]
    for i,w in enumerate(widths,1): ws.column_dimensions[chr(64+i)].width=w
    ws.conditional_formatting.add(f"F5:F{max(5,len(rows)+4)}",
        ColorScaleRule(start_type='min',start_color='F8696B',
                       mid_type='percentile',mid_value=50,mid_color='FFEB84',
                       end_type='max',end_color='63BE7B'))
    # KPI detail
    wk=wb.create_sheet("KPI Detail")
    kh=["Date","Name","KPI","Score"]
    for c,h in enumerate(kh,1): wk.cell(1,c,h).font=Font(bold=True)
    for i,(_,r) in enumerate(scores.iterrows(),2):
        wk.cell(i,1,r["work_date"]); wk.cell(i,2,r["name"]); wk.cell(i,3,r["kpi"]); wk.cell(i,4,r["score"])
    # Sales
    ws2=wb.create_sheet("Sales")
    for c,h in enumerate(["Date","Name","Sales","Source"],1): ws2.cell(1,c,h).font=Font(bold=True)
    for i,(_,r) in enumerate(sales.iterrows(),2):
        ws2.cell(i,1,r["work_date"]); ws2.cell(i,2,r["name"]); ws2.cell(i,3,r["sales"]); ws2.cell(i,4,r["source"])
    # Remarks
    wr=wb.create_sheet("Remarks")
    for c,h in enumerate(["Date","Group","Name","Supervisor","Remark"],1): wr.cell(1,c,h).font=Font(bold=True)
    if not remarks.empty:
        for i,(_,r) in enumerate(remarks.iterrows(),2):
            for c,h in enumerate(["work_date","grp","name","supervisor","remark"],1):
                wr.cell(i,c,r.get(h,""))
    # Employee acknowledgement/correction requests and coaching follow-ups.
    review_data = pd.read_sql_query("SELECT * FROM feedback_reviews WHERE work_date=?", con, params=(str(work_date),))
    wv = wb.create_sheet("Employee Reviews")
    for c,h in enumerate(["Date", "Name", "Status", "Employee comment", "Updated"],1): wv.cell(1,c,h).font=Font(bold=True)
    for i,(_,r) in enumerate(review_data.iterrows(),2):
        for c,h in enumerate(["work_date", "name", "status", "employee_note", "updated_at"],1): wv.cell(i,c,r.get(h,""))
    follow_data = pd.read_sql_query("SELECT * FROM follow_ups WHERE source_date=? OR follow_up_date=?", con, params=(str(work_date), str(work_date)))
    wf = wb.create_sheet("Coaching Follow-ups")
    for c,h in enumerate(["Follow-up date", "Feedback date", "Name", "Area", "Action", "Owner", "Status", "Outcome"],1): wf.cell(1,c,h).font=Font(bold=True)
    for i,(_,r) in enumerate(follow_data.iterrows(),2):
        for c,h in enumerate(["follow_up_date", "source_date", "name", "area", "action", "owner", "status", "outcome"],1): wf.cell(i,c,r.get(h,""))
    # Chart sheet: KPI score ranking
    ch=wb.create_sheet("Charts")
    ch["A1"]="Employee KPI Score Ranking"; ch["A1"].font=Font(size=16,bold=True)
    ranked=sorted(rows,key=lambda x:x[5],reverse=True)
    for i,row in enumerate(ranked,3):
        ch.cell(i,1,row[0]); ch.cell(i,2,row[5])
    chart=BarChart(); chart.title="KPI Score by Employee"; chart.y_axis.title="Score"; chart.x_axis.title="Employee"
    chart.add_data(Reference(ch,min_col=2,min_row=2,max_row=2+len(ranked)),titles_from_data=True)
    chart.set_categories(Reference(ch,min_col=1,min_row=3,max_row=2+len(ranked)))
    chart.height=10; chart.width=20; ch.add_chart(chart,"D3")
    for sh in wb.worksheets:
        for row in sh.iter_rows():
            for cell in row:
                cell.alignment=Alignment(vertical="top",wrap_text=True)
    out=io.BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue()

st.set_page_config(page_title="KPI Bonus Control Centre",layout="wide")
st.title("KPI Bonus Control Centre")
st.caption("Daily sales → employee allocation → weighted KPI score → supervisor remarks → Excel dashboard")

con=db()
with st.sidebar:
    st.header("Setup")
    roster=load_roster()
    st.write(f"Roster: **{len(roster)}** employees")
    work_date=st.date_input("Work date",date.today())
    if st.button("Sync KPI scores from Google Sheet", use_container_width=True):
        try:
            score_count, employee_count = sync_google_kpis(work_date)
            if score_count:
                st.success(f"Synced {score_count} KPI scores for {employee_count} employees.")
            else:
                st.info("No Google Sheet KPI rows matched the current roster. Existing scores were not changed.")
        except ValueError as e:
            st.error(str(e))
    st.divider()
    st.info("Use the Google Apps Script supplied with this project to create supervisor-specific remark links.")

tabs=st.tabs(["Dashboard","Daily Sales","Supervisor KPI","Employee review","Next-day follow-up","Remarks","Excel Export"])

with tabs[0]:
    df = daily_performance(work_date)
    if df.empty:
        st.warning("Load the roster first using the included roster.csv or add it in the database.")
    else:
        followups = pd.read_sql_query("SELECT * FROM follow_ups WHERE follow_up_date=?", con, params=(str(work_date),))
        acknowledged = (df["Feedback status"] == "Acknowledged").sum()
        corrections = (df["Feedback status"] == "Correction requested").sum()
        c1,c2,c3,c4=st.columns(4)
        c1.metric("Employees",len(df))
        c2.metric("Sales loaded",f"{df['Sales'].sum():,.0f}")
        c3.metric("Avg KPI",f"{df['KPI Score'].mean()*100:.1f}%")
        c4.metric("Feedback actioned",f"{acknowledged + corrections}/{len(df)}")
        c5,c6=st.columns(2)
        c5.metric("Correction requests", corrections)
        c6.metric("Follow-ups due today", len(followups))
        st.dataframe(df.sort_values("KPI Score",ascending=False),use_container_width=True,hide_index=True)
        st.subheader("KPI score by employee")
        st.bar_chart(df.sort_values("KPI Score", ascending=False).set_index("Name")["KPI Score"])
        st.subheader("Performance by department")
        st.bar_chart(df.groupby("Department")["KPI Score"].mean().sort_values(ascending=False))

with tabs[1]:
    st.subheader("Upload daily sales")
    st.write("Accepted: Excel or CSV. The app matches the uploaded employee names against the master roster and stores the daily amount.")
    f=st.file_uploader("Sales file",type=["xlsx","xls","csv"])
    if f and st.button("Import sales",type="primary"):
        try:
            raw,matched=import_sales(f,work_date)
            st.success(f"Imported {len(matched)} matched employees.")
            unmatched=raw[raw["matched_name"].isna()]
            if len(unmatched): st.warning(f"{len(unmatched)} rows were not matched.")
            st.dataframe(raw,use_container_width=True,hide_index=True)
        except Exception as e:
            st.error(str(e))

with tabs[2]:
    st.subheader("Supervisor KPI entry")
    roster=load_roster()
    if roster.empty:
        st.warning("Roster is empty.")
    else:
        dept=st.selectbox("Department",sorted(roster["department"].unique()))
        grps=sorted(roster.loc[roster.department==dept,"grp"].unique())
        grp=st.selectbox("Group",grps)
        team=roster[(roster.department==dept)&(roster.grp==grp)]
        st.write(f"**{grp}** — {len(team)} team members")
        mp=kpi_map(dept,grp)
        with st.form("kpi_form"):
            selected={}
            for _,person in team.iterrows():
                st.markdown(f"**{person['name']}**")
                cols=st.columns(min(5,len(mp)))
                for idx,(k,w) in enumerate(mp.items()):
                    with cols[idx % len(cols)]:
                        selected[(person["name"],k)]=st.selectbox(k,[0,0.5,1],index=2,key=f"{person['name']}_{k}")
            if st.form_submit_button("Save KPI scores",type="primary"):
                for (name,k),v in selected.items():
                    con.execute("INSERT OR REPLACE INTO kpi_scores VALUES(?,?,?,?)",
                                (str(work_date),name,k,float(v)))
                con.commit(); st.success("KPI scores saved.")

with tabs[3]:
    st.subheader("Employee feedback review")
    st.write("Use this page with the employee after scores are published. They can acknowledge the feedback or request a correction with an explanation.")
    roster = load_roster()
    if not roster.empty:
        employee = st.selectbox("Employee", roster["name"].tolist(), key="review_employee")
        existing = con.execute(
            "SELECT status, employee_note FROM feedback_reviews WHERE work_date=? AND name=?",
            (str(work_date), employee),
        ).fetchone()
        status = st.radio("Feedback decision", ["Acknowledged", "Correction requested"],
                          index=0 if not existing or existing[0] == "Acknowledged" else 1)
        note = st.text_area("Employee comment / correction detail",
                            value=existing[1] if existing else "",
                            placeholder="State the score or feedback item to review, and why.")
        if st.button("Save employee review", type="primary"):
            con.execute("INSERT OR REPLACE INTO feedback_reviews VALUES(?,?,?,?,datetime('now'))",
                        (str(work_date), employee, status, note))
            con.commit()
            st.success("Employee feedback review saved.")
        review_rows = pd.read_sql_query(
            "SELECT name AS Employee, status AS Status, employee_note AS Comment, updated_at AS Updated FROM feedback_reviews WHERE work_date=? ORDER BY updated_at DESC",
            con, params=(str(work_date),))
        if not review_rows.empty:
            st.dataframe(review_rows, use_container_width=True, hide_index=True)

with tabs[4]:
    st.subheader("Next-day improvement follow-up")
    st.write("Log a focused action for the next shift, then record whether the employee improved. This is the coaching trail for supervisors and management.")
    roster = load_roster()
    if not roster.empty:
        source_date = st.date_input("Feedback date", value=work_date, key="followup_source")
        follow_date = st.date_input("Follow-up date", value=work_date, key="followup_date")
        employee = st.selectbox("Employee", roster["name"].tolist(), key="followup_employee")
        area = st.text_input("Improvement area", placeholder="e.g. Order accuracy")
        action = st.text_area("Agreed action", placeholder="e.g. Repeat the order back to every guest before submitting it.")
        owner = st.text_input("Follow-up owner", placeholder="Manager or rotational supervisor")
        status = st.selectbox("Status", ["Open", "Completed", "Needs further coaching"])
        outcome = st.text_area("Outcome / evidence", placeholder="Record the next-day observation once completed.")
        if st.button("Save follow-up", type="primary"):
            if not area.strip() or not action.strip():
                st.error("Enter both an improvement area and an agreed action.")
            else:
                con.execute("INSERT OR REPLACE INTO follow_ups VALUES(?,?,?,?,?,?,?,?)",
                    (str(follow_date), str(source_date), employee, area.strip(), action.strip(), owner.strip(), status, outcome.strip()))
                con.commit()
                st.success("Next-day follow-up saved.")
        due = pd.read_sql_query(
            "SELECT source_date AS 'Feedback date', name AS Employee, area AS Area, action AS Action, owner AS Owner, status AS Status, outcome AS Outcome FROM follow_ups WHERE follow_up_date=? ORDER BY status, name",
            con, params=(str(follow_date),))
        if not due.empty:
            st.dataframe(due, use_container_width=True, hide_index=True)

with tabs[5]:
    st.subheader("Supervisor remarks")
    st.write("Remarks can come from the Google Form/Sheet feed, or can be entered here for testing.")
    roster=load_roster()
    if not roster.empty:
        grp=st.selectbox("Group",sorted(roster["grp"].unique()),key="remark_group")
        team=roster[roster.grp==grp]
        person=st.selectbox("Team member",team["name"].tolist())
        supervisor=st.text_input("Supervisor")
        remark=st.text_area("Remark")
        if st.button("Save remark"):
            con.execute("INSERT OR REPLACE INTO remarks VALUES(?,?,?,?,?)",
                        (str(work_date),grp,person,supervisor,remark))
            con.commit(); st.success("Remark saved.")
    google=pull_google_remarks()
    if not google.empty:
        st.subheader("Google remarks feed")
        st.dataframe(google.tail(50),use_container_width=True,hide_index=True)
    elif GOOGLE_SHEET_CSV_URL:
        st.info("No remarks found for the selected date.")
    else:
        st.caption("Set GOOGLE_SHEET_CSV_URL to your published Google Sheet CSV endpoint to sync supervisor remarks.")

with tabs[6]:
    st.subheader("Excel dashboard export")
    st.write("The export includes the dashboard, KPI detail, sales, remarks, employee reviews, coaching follow-ups and charts.")
    if st.button("Build Excel dashboard",type="primary"):
        data=export_excel(work_date)
        st.download_button("Download Excel dashboard",data=data,
                           file_name=f"KPI_Bonus_Dashboard_{work_date}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
