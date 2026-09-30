
import io
import os
import re
import sqlite3
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st


# ============================================================
# KPI BONUS CONTROL CENTRE — DYNAMIC DEPARTMENT KPI VERSION
# ============================================================
# Design:
# 1. KPI workbooks are the source of truth for KPI names/weights/groups.
# 2. No department KPI list is hard-coded.
# 3. Daily scores are stored by date + employee + KPI.
# 4. Sales are stored separately and matched to the master roster.
# 5. Bonus pools and sales targets are configurable in the app.
# 6. Excel export produces an auditable dashboard and calculation detail.
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DB = os.getenv("KPI_DB", str(BASE_DIR / "kpi_bonus.db"))
GOOGLE_SHEET_CSV_URL = os.getenv("GOOGLE_SHEET_CSV_URL", "")

DEFAULT_SCORES = [1.0, 0.8, 0.6, 0.4, 0.0]


# -----------------------------
# General helpers
# -----------------------------

def norm(x):
    if x is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(x).lower())


def clean_name(x):
    return re.sub(r"\s+", " ", str(x or "").strip())


def to_float(x, default=None):
    try:
        if x is None or str(x).strip() == "":
            return default
        return float(str(x).replace(",", "").replace("%", "").strip())
    except Exception:
        return default


def _ensure_column(con, table, column, definition):
    """Add a missing column to an existing SQLite table.

    CREATE TABLE IF NOT EXISTS does not update an already-created table, so
    this small migration keeps older kpi_bonus.db files compatible with the
    current application.
    """
    cols = {row[1] for row in con.execute(f"PRAGMA table_info({table})").fetchall()}
    if column not in cols:
        con.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def db():
    con = sqlite3.connect(DB)
    con.execute("""
        CREATE TABLE IF NOT EXISTS roster(
            name TEXT PRIMARY KEY,
            department TEXT NOT NULL,
            grp TEXT NOT NULL DEFAULT '',
            supervisor TEXT NOT NULL DEFAULT '',
            target_type TEXT NOT NULL DEFAULT 'individual'
        )
    """)
    con.execute("""
        CREATE TABLE IF NOT EXISTS kpi_templates(
            department TEXT NOT NULL,
            grp TEXT NOT NULL,
            kpi TEXT NOT NULL,
            weight REAL NOT NULL,
            source_file TEXT,
            PRIMARY KEY(department, grp, kpi)
        )
    """)
    con.execute("""
        CREATE TABLE IF NOT EXISTS kpi_scores(
            work_date TEXT NOT NULL,
            name TEXT NOT NULL,
            department TEXT NOT NULL,
            grp TEXT NOT NULL,
            kpi TEXT NOT NULL,
            score REAL NOT NULL,
            comment TEXT NOT NULL DEFAULT '',
            PRIMARY KEY(work_date, name, kpi)
        )
    """)
    con.execute("""
        CREATE TABLE IF NOT EXISTS sales(
            work_date TEXT NOT NULL,
            name TEXT NOT NULL,
            sales REAL NOT NULL,
            source TEXT NOT NULL,
            PRIMARY KEY(work_date, name)
        )
    """)
    con.execute("""
        CREATE TABLE IF NOT EXISTS targets(
            period TEXT NOT NULL,
            name TEXT NOT NULL,
            department TEXT NOT NULL,
            grp TEXT NOT NULL,
            target REAL NOT NULL,
            target_type TEXT NOT NULL DEFAULT 'individual',
            PRIMARY KEY(period, name)
        )
    """)
    con.execute("""
        CREATE TABLE IF NOT EXISTS bonus_pools(
            period TEXT NOT NULL,
            department TEXT NOT NULL,
            grp TEXT NOT NULL,
            pool REAL NOT NULL DEFAULT 0,
            target REAL NOT NULL DEFAULT 0,
            target_type TEXT NOT NULL DEFAULT 'department',
            PRIMARY KEY(period, department, grp)
        )
    """)
    con.execute("""
        CREATE TABLE IF NOT EXISTS remarks(
            work_date TEXT NOT NULL,
            grp TEXT NOT NULL,
            name TEXT NOT NULL,
            supervisor TEXT NOT NULL DEFAULT '',
            remark TEXT NOT NULL DEFAULT '',
            PRIMARY KEY(work_date, grp, name)
        )
    """)

    # ---- Backward-compatible database migration ----
    # Older versions of the app created roster without target_type.
    # CREATE TABLE IF NOT EXISTS will not add a new column to an existing
    # table, so explicitly migrate old databases here.
    _ensure_column(con, "roster", "grp", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(con, "roster", "supervisor", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(con, "roster", "target_type", "TEXT NOT NULL DEFAULT 'individual'")

    _ensure_column(con, "kpi_scores", "department", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(con, "kpi_scores", "grp", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(con, "kpi_scores", "comment", "TEXT NOT NULL DEFAULT ''")

    _ensure_column(con, "bonus_pools", "target", "REAL NOT NULL DEFAULT 0")
    _ensure_column(con, "bonus_pools", "target_type", "TEXT NOT NULL DEFAULT 'department'")

    con.execute("UPDATE roster SET target_type = 'individual' WHERE target_type IS NULL OR TRIM(target_type) = ''")
    con.execute("UPDATE roster SET supervisor = '' WHERE supervisor IS NULL")
    con.execute("UPDATE roster SET grp = '' WHERE grp IS NULL")
    con.commit()
    return con


# -----------------------------
# KPI workbook parser
# -----------------------------

DEPARTMENT_PREFIXES = {
    "SERVICE": "Service",
    "BAR": "Bar",
    "KTN": "KTN",
    "H_K": "Housekeeping",
}

LOCAL_KPI_PATTERNS = {
    "Service": ("SERVICE",),
    "Bar": ("BAR",),
    "KTN": ("KTN",),
    "Housekeeping": ("H_K",),
}


def department_name(filename):
    u = Path(str(filename)).name.upper()
    for prefix, dept in DEPARTMENT_PREFIXES.items():
        if u.startswith(prefix):
            return dept
    return Path(str(filename)).stem


def _is_formula(v):
    return isinstance(v, str) and v.strip().startswith("=")


def _text_cells(row):
    out = {}
    for c, v in enumerate(row):
        if c == 0 or v is None:
            continue
        if isinstance(v, str):
            s = v.strip()
            if s and norm(s) not in {"totalkpi", "totalkpi%", "comments"}:
                out[c] = s
    return out


def _numeric_cells(row):
    return {
        c: to_float(v)
        for c, v in enumerate(row)
        if v is not None and to_float(v) is not None
    }


def _weights_from_labels(headers):
    result = {}
    for c, label in headers.items():
        # Only use parentheses as a weight when the number is clearly part
        # of the KPI label, e.g. "Hygiene (10)".
        m = re.search(r"\((\d+(?:\.\d+)?)\)", str(label))
        if m:
            result[c] = float(m.group(1)) / 100.0
    return result


def _weights_from_formula(formula, header_cols):
    result = {}
    if not isinstance(formula, str):
        return result

    for c in header_cols:
        n = c + 1
        col = ""
        while n:
            n, rem = divmod(n - 1, 26)
            col = chr(65 + rem) + col

        m = re.search(
            rf"{re.escape(col)}\d+\s*\*\s*(\d+(?:\.\d+)?)",
            formula,
            re.I,
        )
        if m:
            result[c] = float(m.group(1)) / 100.0
    return result


def _looks_like_header(texts):
    if len(texts) < 2:
        return False
    bad = {"staffname", "total", "totalkpi", "comments"}
    return sum(norm(v) not in bad for v in texts.values()) >= 2


def parse_kpi_workbook(source, display_name=None):
    """Parse a KPI workbook efficiently without loading the entire sheet into memory."""
    import openpyxl

    filename = display_name or getattr(source, "name", str(source))
    department = department_name(filename)

    # read_only=True is important here. The KPI workbooks are templates, and
    # loading them normally can make Streamlit appear frozen before the UI
    # finishes rendering. We only need cell values/formulas.
    wb = openpyxl.load_workbook(
        source,
        data_only=False,
        read_only=True,
    )

    templates = []
    roster = []

    try:
        for ws in wb.worksheets:
            # Only inspect the first 25 rows to locate the Staff Name row.
            preview = list(ws.iter_rows(min_row=1, max_row=25, values_only=True))
            if not preview:
                continue

            staff_row = None
            for r, row in enumerate(preview):
                vals = {norm(v) for v in row if v is not None}
                if "staffname" in vals:
                    staff_row = r
                    break

            if staff_row is None:
                continue

            # Staff row is 0-based in preview; Excel rows are 1-based.
            staff_excel_row = staff_row + 1
            staff_values = preview[staff_row]

            current_headers = _text_cells(staff_values)
            current_weights = _weights_from_labels(current_headers)
            current_group = ""
            pending_header_row = None

            # Most of these workbooks are small, but Excel formatting can make
            # max_row enormous. 5,000 rows is safely above the actual staff
            # sections while preventing accidental million-row scans.
            max_row = ws.max_row or staff_excel_row
            max_row = min(max_row, 5000)

            for row in ws.iter_rows(
                min_row=staff_excel_row + 1,
                max_row=max_row,
                values_only=True,
            ):
                first_raw = row[0] if row else None
                first = clean_name(first_raw)
                texts = _text_cells(row)
                nums = _numeric_cells(row)
                formulas = [v for v in row if _is_formula(v)]
                has_formula = bool(formulas)

                if not first and not texts and not nums:
                    continue

                if not has_formula and _looks_like_header(texts):
                    if first and "%" in first:
                        current_group = first
                        current_headers = texts
                        current_weights = _weights_from_labels(current_headers)
                        pending_header_row = None
                        continue

                    if not first:
                        current_headers = texts
                        current_weights = _weights_from_labels(current_headers)
                        pending_header_row = True
                        continue

                    if len(texts) >= 2 and first:
                        current_group = first
                        current_headers = texts
                        current_weights = _weights_from_labels(current_headers)
                        pending_header_row = None
                        continue

                header_cols = list(current_headers.keys())
                if header_cols:
                    row_weights = {
                        c: nums[c]
                        for c in header_cols
                        if c in nums and 0 <= nums[c] <= 1
                    }
                    if len(row_weights) >= max(2, len(header_cols) // 2):
                        if first:
                            current_group = first
                            pending_header_row = None
                        current_weights.update(row_weights)
                        continue

                if (
                    not has_formula
                    and first
                    and current_headers
                    and not texts
                    and not nums
                ):
                    current_group = first
                    pending_header_row = None
                    continue

                if has_formula and first:
                    if pending_header_row:
                        current_group = "UNLABELLED GROUP"
                        pending_header_row = None

                    if current_headers and len(current_weights) < len(current_headers):
                        formula = next((v for v in row if _is_formula(v)), None)
                        formula_weights = _weights_from_formula(
                            formula,
                            current_headers.keys(),
                        )
                        for col, weight in formula_weights.items():
                            current_weights.setdefault(col, weight)

                    valid = [
                        (c, label, current_weights.get(c))
                        for c, label in current_headers.items()
                        if current_weights.get(c) is not None
                    ]

                    if valid:
                        roster.append({
                            "name": first,
                            "department": department,
                            "grp": current_group,
                        })
                        for _, label, weight in valid:
                            templates.append({
                                "department": department,
                                "grp": current_group,
                                "kpi": label.strip(),
                                "weight": float(weight),
                                "source_file": filename,
                            })
    finally:
        wb.close()

    unique_roster = {}
    for item in roster:
        key = (item["department"], item["grp"], norm(item["name"]))
        if key[2]:
            unique_roster[key] = item

    unique_templates = {}
    for item in templates:
        key = (item["department"], item["grp"], norm(item["kpi"]))
        unique_templates[key] = item

    return list(unique_templates.values()), list(unique_roster.values())

def discover_local_kpi_files():
    """
    Automatically find the KPI workbooks beside app.py on Windows.

    Expected project folder:
      C:\\Users\\Admin\\Desktop\\KPI_Bonus_Control_Centre\\

    The app deliberately ignores other Excel files such as dashboard
    templates and generated exports.
    """
    base_dir = BASE_DIR
    found = []

    for path in sorted(base_dir.iterdir()):
        if not path.is_file() or path.suffix.lower() not in {".xlsx", ".xls"}:
            continue

        upper = path.name.upper()
        if upper.startswith("SERVICE"):
            found.append(("Service", path))
        elif upper.startswith("BAR"):
            found.append(("Bar", path))
        elif upper.startswith("KTN"):
            found.append(("KTN", path))
        elif upper.startswith("H_K"):
            found.append(("Housekeeping", path))

    return found


def sync_local_kpi_workbooks(force=False):
    """
    Sync local KPI workbooks into SQLite.

    The sync is based on file modification times so Streamlit reruns do not
    repeatedly parse the workbooks. A manual Sync button can force it.
    """
    files = discover_local_kpi_files()
    if not files:
        return []

    state = st.session_state.setdefault("kpi_file_mtimes", {})
    changed = force or any(
        state.get(str(path)) != path.stat().st_mtime_ns
        for _, path in files
    )

    if not changed:
        return []

    all_templates = []
    all_roster = []
    source_files = []
    errors = []

    for dept, path in files:
        try:
            templates, staff = parse_kpi_workbook(path, path.name)
            all_templates.extend(templates)
            all_roster.extend(staff)
            source_files.append(path.name)
            state[str(path)] = path.stat().st_mtime_ns
        except Exception as exc:
            errors.append(f"{path.name}: {exc}")

    if all_templates:
        save_kpi_configuration(
            all_templates,
            all_roster,
            source_files=source_files,
            departments={dept for dept, _ in files},
        )

    if errors:
        for error in errors:
            st.error(f"KPI workbook error: {error}")

    return [
        {
            "department": dept,
            "file": str(path),
            "employees": sum(
                1 for r in all_roster if r["department"] == dept
            ),
        }
        for dept, path in files
    ]

def save_kpi_configuration(templates, roster, source_files=None, departments=None):
    con = db()
    source_files = set(source_files or [])
    departments = set(departments or {r["department"] for r in roster})

    # Preserve existing supervisor/target settings before replacing the
    # roster for departments supplied by the local KPI workbooks.
    old = pd.read_sql_query(
        "SELECT name, supervisor, target_type FROM roster",
        con,
    )
    preserved = {
        norm(row["name"]): (
            str(row["supervisor"] or ""),
            str(row["target_type"] or "individual"),
        )
        for _, row in old.iterrows()
    }

    # Remove the old KPI definitions from these source files/departments.
    if source_files:
        placeholders = ",".join("?" for _ in source_files)
        con.execute(
            f"DELETE FROM kpi_templates WHERE source_file IN ({placeholders})",
            tuple(source_files),
        )

    if departments:
        placeholders = ",".join("?" for _ in departments)
        con.execute(
            f"DELETE FROM kpi_templates WHERE department IN ({placeholders})",
            tuple(departments),
        )

        # Replace only the roster departments represented by the current
        # workbooks. Existing supervisor/target settings are restored below.
        con.execute(
            f"DELETE FROM roster WHERE department IN ({placeholders})",
            tuple(departments),
        )

    for t in templates:
        con.execute(
            """
            INSERT OR REPLACE INTO kpi_templates
            (department, grp, kpi, weight, source_file)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                t["department"],
                t["grp"],
                t["kpi"],
                t["weight"],
                t["source_file"],
            ),
        )

    # Roster is unique by employee name in the current database design.
    # Exact duplicate rows inside one workbook are collapsed by the parser.
    for r in roster:
        supervisor, target_type = preserved.get(
            norm(r["name"]),
            ("", "individual"),
        )
        con.execute(
            """
            INSERT OR REPLACE INTO roster
            (name, department, grp, supervisor, target_type)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                r["name"],
                r["department"],
                r["grp"],
                supervisor,
                target_type,
            ),
        )

    con.commit()

def load_roster():
    con = db()
    return pd.read_sql_query(
        "SELECT * FROM roster ORDER BY department, grp, name", con
    )


def load_templates(department=None, grp=None):
    con = db()
    sql = "SELECT * FROM kpi_templates WHERE 1=1"
    params = []

    if department:
        sql += " AND department=?"
        params.append(department)
    if grp is not None:
        sql += " AND grp=?"
        params.append(grp)

    sql += " ORDER BY department, grp, rowid"
    return pd.read_sql_query(sql, con, params=params)


def kpi_map(department, grp):
    df = load_templates(department, grp)
    if df.empty:
        return {}
    return dict(zip(df["kpi"], df["weight"]))


# -----------------------------
# Scores and KPI calculations
# -----------------------------

def weighted_kpi_score(scores_df, department, grp, name):
    mp = kpi_map(department, grp)
    if not mp:
        return 0.0

    person = scores_df[
        (scores_df["name"] == name) &
        (scores_df["department"] == department) &
        (scores_df["grp"] == grp)
    ]

    total = 0.0
    for kpi, weight in mp.items():
        rows = person[person["kpi"] == kpi]
        score = float(rows.iloc[-1]["score"]) if not rows.empty else 0.0
        total += score * float(weight)

    # The supplied workbooks use a 95% daily cap unless justified.
    return min(0.95, total)


def validate_score_comments(team, score_values, comments):
    problems = []

    for name in team["name"]:
        for kpi, value in score_values.get(name, {}).items():
            if value == 0 and not str(comments.get((name, kpi), "")).strip():
                problems.append(f"{name}: '{kpi}' is 0.0 and needs a comment.")

    return problems


# -----------------------------
# Sales import
# -----------------------------

def import_sales(upload, work_date):
    if upload.name.lower().endswith(".csv"):
        df = pd.read_csv(upload)
    else:
        xls = pd.ExcelFile(upload)
        frames = [
            pd.read_excel(upload, sheet_name=s, header=0)
            for s in xls.sheet_names
        ]
        df = pd.concat(frames, ignore_index=True)

    df.columns = [str(c).strip() for c in df.columns]

    name_col = next(
        (
            c for c in df.columns
            if norm(c) in {
                "name", "employee", "employeename", "staff",
                "staffname", "waiter", "waitername"
            }
        ),
        None
    )

    sales_col = next(
        (
            c for c in df.columns
            if norm(c) in {
                "sales", "individualsales", "amount", "revenue",
                "netsales", "totalsales", "total_sales"
            }
        ),
        None
    )

    if not name_col or not sales_col:
        raise ValueError(
            "Could not identify Name and Sales columns. "
            "Expected columns such as Name and Sales."
        )

    df = df[[name_col, sales_col]].copy()
    df.columns = ["name", "sales"]
    df["name"] = df["name"].map(clean_name)
    df["sales"] = pd.to_numeric(df["sales"], errors="coerce").fillna(0)
    df["name_key"] = df["name"].map(norm)

    roster = load_roster()
    lookup = {norm(n): n for n in roster["name"]}

    df["matched_name"] = df["name_key"].map(lookup)

    con = db()

    # Aggregate duplicate rows before saving.
    matched = df[df["matched_name"].notna()].copy()
    matched = matched.groupby("matched_name", as_index=False)["sales"].sum()

    for _, r in matched.iterrows():
        con.execute("""
            INSERT OR REPLACE INTO sales(work_date, name, sales, source)
            VALUES (?, ?, ?, ?)
        """, (
            str(work_date),
            r["matched_name"],
            float(r["sales"]),
            upload.name
        ))

    con.commit()

    return df, matched


# -----------------------------
# Targets / bonus pools
# -----------------------------

def save_target(period, name, department, grp, target, target_type):
    con = db()
    con.execute("""
        INSERT OR REPLACE INTO targets
        (period, name, department, grp, target, target_type)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (period, name, department, grp, float(target), target_type))
    con.commit()


def save_bonus_pool(period, department, grp, pool, target, target_type):
    con = db()
    con.execute("""
        INSERT OR REPLACE INTO bonus_pools
        (period, department, grp, pool, target, target_type)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (period, department, grp, float(pool), float(target), target_type))
    con.commit()


def calculate_bonus(period, work_date):
    """
    Monthly/period bonus engine.

    Two target models are supported:

    1. individual:
       Employee sales achievement = period sales / individual target.
       Bonus = pool * capped achievement * period KPI score.

    2. department:
       Group sales achievement = period group sales / group target.
       Earned group pool = pool * capped group achievement.
       The earned pool is then distributed proportionally by each employee's
       KPI score. This keeps the department target collective while still
       differentiating individual performance.

    The department model is deliberately visible in the output so the
    Performance Lead can audit and explain every amount.
    """
    con = db()

    roster = load_roster()

    sales = pd.read_sql_query(
        """
        SELECT * FROM sales
        WHERE substr(work_date, 1, 7)=?
        """,
        con,
        params=(period,)
    )

    scores = pd.read_sql_query(
        """
        SELECT * FROM kpi_scores
        WHERE substr(work_date, 1, 7)=?
        """,
        con,
        params=(period,)
    )

    targets = pd.read_sql_query(
        "SELECT * FROM targets WHERE period=?",
        con,
        params=(period,)
    )

    pools = pd.read_sql_query(
        "SELECT * FROM bonus_pools WHERE period=?",
        con,
        params=(period,)
    )

    if not pools.empty:
        pools["target_type"] = pools["target_type"].fillna("department")

    # Period KPI = average of each employee's daily KPI score.
    kpi_by_person = {}
    for _, person in roster.iterrows():
        name = person["name"]
        dept = person["department"]
        grp = person["grp"]

        person_scores = scores[
            (scores["name"] == name) &
            (scores["department"] == dept) &
            (scores["grp"] == grp)
        ]

        if person_scores.empty:
            kpi_by_person[name] = 0.0
            continue

        dates = sorted(person_scores["work_date"].astype(str).unique())
        daily_scores = []

        for d in dates:
            day_scores = person_scores[
                person_scores["work_date"].astype(str) == str(d)
            ]
            total = 0.0
            mp = kpi_map(dept, grp)

            for kpi, weight in mp.items():
                vals = day_scores[day_scores["kpi"] == kpi]["score"]
                value = float(vals.iloc[-1]) if not vals.empty else 0.0
                total += value * float(weight)

            daily_scores.append(min(0.95, total))

        kpi_by_person[name] = (
            sum(daily_scores) / len(daily_scores)
            if daily_scores else 0.0
        )

    # Period sales by employee.
    sales_by_name = (
        sales.groupby("name", as_index=False)["sales"].sum()
        if not sales.empty
        else pd.DataFrame(columns=["name", "sales"])
    )

    # Build a working table.
    base = []

    for _, person in roster.iterrows():
        name = person["name"]
        dept = person["department"]
        grp = person["grp"]

        sr = sales_by_name[sales_by_name["name"] == name]
        actual_sales = float(sr["sales"].iloc[0]) if not sr.empty else 0.0

        tr = targets[targets["name"] == name]
        individual_target = (
            float(tr["target"].iloc[0])
            if not tr.empty else 0.0
        )

        pool_row = pools[
            (pools["department"] == dept) &
            (pools["grp"] == grp)
        ]

        pool = float(pool_row["pool"].iloc[0]) if not pool_row.empty else 0.0
        pool_target = (
            float(pool_row["target"].iloc[0])
            if not pool_row.empty else 0.0
        )
        pool_type = (
            str(pool_row["target_type"].iloc[0])
            if not pool_row.empty else "individual"
        )

        base.append({
            "Name": name,
            "Department": dept,
            "Group": grp,
            "Sales": actual_sales,
            "Individual Target": individual_target,
            "KPI Score": kpi_by_person.get(name, 0.0),
            "Bonus Pool": pool,
            "Pool Target": pool_target,
            "Pool Target Type": pool_type,
        })

    result = pd.DataFrame(base)

    if result.empty:
        return result

    # Initialize outputs.
    result["Sales Achievement"] = 0.0
    result["Earned Pool"] = 0.0
    result["Bonus"] = 0.0
    result["Calculation"] = ""

    # Individual target model.
    individual_mask = result["Pool Target Type"].str.lower().eq("individual")

    for idx in result.index[individual_mask]:
        row = result.loc[idx]

        target = float(row["Individual Target"])
        achievement = (
            min(1.0, float(row["Sales"]) / target)
            if target > 0 else 0.0
        )

        earned = float(row["Bonus Pool"]) * achievement
        bonus = earned * float(row["KPI Score"])

        result.loc[idx, "Sales Achievement"] = achievement
        result.loc[idx, "Earned Pool"] = earned
        result.loc[idx, "Bonus"] = bonus
        result.loc[idx, "Calculation"] = (
            "Individual: Pool × Sales Achievement × KPI"
        )

    # Department/group target model.
    department_mask = ~individual_mask

    for (dept, grp), idxs in result[department_mask].groupby(
        ["Department", "Group"]
    ).groups.items():

        target_values = result.loc[idxs, "Pool Target"]
        target = float(target_values.iloc[0]) if len(target_values) else 0.0

        group_sales = float(result.loc[idxs, "Sales"].sum())
        achievement = min(1.0, group_sales / target) if target > 0 else 0.0

        pool = float(result.loc[idxs, "Bonus Pool"].iloc[0])
        earned_pool = pool * achievement

        # For collective departments, first earn the group pool from
        # group sales achievement. Then allocate that earned pool according
        # to each employee's share of the group's actual sales, and finally
        # apply the employee's KPI score.
        group_sales = max(group_sales, 0.0)

        for idx in idxs:
            employee_sales = max(float(result.loc[idx, "Sales"]), 0.0)
            sales_share = (
                employee_sales / group_sales
                if group_sales > 0 else 0.0
            )
            kpi = max(float(result.loc[idx, "KPI Score"]), 0.0)
            bonus = earned_pool * sales_share * kpi

            result.loc[idx, "Sales Achievement"] = achievement
            result.loc[idx, "Earned Pool"] = earned_pool * sales_share
            result.loc[idx, "Bonus"] = bonus
            result.loc[idx, "Calculation"] = (
                "Department: Pool × Group Achievement × "
                "Employee Sales Share × KPI"
            )

    return result


# -----------------------------
# Google remarks
# -----------------------------

def pull_google_remarks():
    columns = ["work_date", "grp", "name", "supervisor", "remark"]

    if not GOOGLE_SHEET_CSV_URL:
        return pd.DataFrame(columns=columns)

    try:
        df = pd.read_csv(GOOGLE_SHEET_CSV_URL)
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        return df[columns]
    except Exception as e:
        st.warning(f"Google remarks feed could not be read: {e}")
        return pd.DataFrame(columns=columns)


# -----------------------------
# Excel export
# -----------------------------

def export_excel(work_date, period):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.chart import BarChart, Reference

    con = db()
    roster = load_roster()

    sales = pd.read_sql_query(
        "SELECT * FROM sales WHERE work_date=?",
        con, params=(str(work_date),)
    )

    scores = pd.read_sql_query(
        "SELECT * FROM kpi_scores WHERE work_date=?",
        con, params=(str(work_date),)
    )

    remarks = pull_google_remarks()
    if not remarks.empty and "work_date" in remarks.columns:
        remarks = remarks[
            remarks["work_date"].astype(str) == str(work_date)
        ]

    bonus = calculate_bonus(period, work_date)

    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"

    ws["A1"] = "KPI BONUS DASHBOARD"
    ws["A1"].font = Font(size=20, bold=True)
    ws["A2"] = "Work Date"
    ws["B2"] = str(work_date)
    ws["A3"] = "Period"
    ws["B3"] = period

    headers = [
        "Name", "Department", "Group", "Sales",
        "Sales Target", "Sales Achievement",
        "KPI Score", "Bonus Pool", "KPI-Adjusted Bonus"
    ]

    for c, h in enumerate(headers, 1):
        ws.cell(5, c, h).font = Font(bold=True)

    for ridx, row in enumerate(bonus.itertuples(index=False), 6):
        for c, value in enumerate(row, 1):
            ws.cell(ridx, c, value)

    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:I{max(5, len(bonus) + 5)}"

    widths = [25, 15, 25, 15, 15, 18, 14, 15, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w

    ws.conditional_formatting.add(
        f"G6:G{max(6, len(bonus) + 5)}",
        ColorScaleRule(
            start_type="min", start_color="F8696B",
            mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max", end_color="63BE7B"
        )
    )

    # KPI Detail
    wk = wb.create_sheet("KPI Detail")
    for c, h in enumerate(
        ["Date", "Name", "Department", "Group", "KPI", "Score", "Comment"], 1
    ):
        wk.cell(1, c, h).font = Font(bold=True)

    for i, (_, r) in enumerate(scores.iterrows(), 2):
        vals = [
            r["work_date"], r["name"], r["department"],
            r["grp"], r["kpi"], r["score"], r["comment"]
        ]
        for c, v in enumerate(vals, 1):
            wk.cell(i, c, v)

    # KPI Templates
    wt = wb.create_sheet("KPI Templates")
    templates = load_templates()
    for c, h in enumerate(
        ["Department", "Group", "KPI", "Weight", "Source File"], 1
    ):
        wt.cell(1, c, h).font = Font(bold=True)

    for i, (_, r) in enumerate(templates.iterrows(), 2):
        for c, h in enumerate(
            ["department", "grp", "kpi", "weight", "source_file"], 1
        ):
            wt.cell(i, c, r[h])

    # Sales
    ws2 = wb.create_sheet("Sales")
    for c, h in enumerate(["Date", "Name", "Sales", "Source"], 1):
        ws2.cell(1, c, h).font = Font(bold=True)

    for i, (_, r) in enumerate(sales.iterrows(), 2):
        vals = [r["work_date"], r["name"], r["sales"], r["source"]]
        for c, v in enumerate(vals, 1):
            ws2.cell(i, c, v)

    # Bonus calculations
    wbns = wb.create_sheet("Bonus Calculation")
    if not bonus.empty:
        for c, h in enumerate(bonus.columns, 1):
            wbns.cell(1, c, h).font = Font(bold=True)
        for i, row in enumerate(bonus.itertuples(index=False), 2):
            for c, v in enumerate(row, 1):
                wbns.cell(i, c, v)

    # Remarks
    wr = wb.create_sheet("Remarks")
    for c, h in enumerate(
        ["Date", "Group", "Name", "Supervisor", "Remark"], 1
    ):
        wr.cell(1, c, h).font = Font(bold=True)

    if not remarks.empty:
        for i, (_, r) in enumerate(remarks.iterrows(), 2):
            for c, h in enumerate(
                ["work_date", "grp", "name", "supervisor", "remark"], 1
            ):
                wr.cell(i, c, r.get(h, ""))

    # Charts
    ch = wb.create_sheet("Charts")
    ch["A1"] = "Employee KPI Score"
    ch["A1"].font = Font(size=16, bold=True)
    ch["A2"] = "Employee"
    ch["B2"] = "KPI Score"

    ranked = bonus.sort_values("KPI Score", ascending=False)

    for i, (_, row) in enumerate(ranked.iterrows(), 3):
        ch.cell(i, 1, row["Name"])
        ch.cell(i, 2, row["KPI Score"])

    if len(ranked):
        chart = BarChart()
        chart.title = "KPI Score by Employee"
        chart.y_axis.title = "Score"
        chart.x_axis.title = "Employee"
        chart.add_data(
            Reference(ch, min_col=2, min_row=2, max_row=2 + len(ranked)),
            titles_from_data=True
        )
        chart.set_categories(
            Reference(ch, min_col=1, min_row=3, max_row=2 + len(ranked))
        )
        chart.height = 10
        chart.width = 20
        ch.add_chart(chart, "D3")

    for sh in wb.worksheets:
        for row in sh.iter_rows():
            for cell in row:
                cell.alignment = Alignment(
                    vertical="top", wrap_text=True
                )

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# ============================================================
# STREAMLIT APP
# ============================================================

st.set_page_config(
    page_title="KPI Bonus Control Centre",
    layout="wide"
)

st.title("KPI Bonus Control Centre")
st.caption(
    "Dynamic department KPIs → daily scoring → sales achievement → "
    "bonus calculation → auditable Excel dashboard"
)

con = db()

# Automatically load the KPI workbooks that sit beside app.py.
# The status is rendered before parsing so the page never looks frozen while
# Excel files are being read.
with st.status("Loading KPI workbooks...", expanded=False) as kpi_load_status:
    local_kpi_status = sync_local_kpi_workbooks()
    if local_kpi_status:
        kpi_load_status.update(label="KPI workbooks loaded", state="complete")
    else:
        kpi_load_status.update(label="KPI workbooks ready", state="complete")

with st.sidebar:
    st.header("Control Centre")

    work_date = st.date_input("Work date", date.today())
    period = st.text_input(
        "Bonus period",
        value=work_date.strftime("%Y-%m")
    )

    st.divider()
    st.write("### Current roster")
    roster = load_roster()
    st.metric("Employees", len(roster))

    if not roster.empty:
        st.write(
            roster.groupby("department")["name"].count()
        )


tabs = st.tabs([
    "1. KPI Setup",
    "2. Dashboard",
    "3. Daily Sales",
    "4. KPI Scoring",
    "5. Targets & Bonus",
    "6. Remarks",
    "7. Excel Export",
])


# ============================================================
# TAB 1 — KPI SETUP
# ============================================================

with tabs[0]:
    st.subheader("KPI Workbook Configuration")

    st.info(
        "The app automatically reads the KPI workbooks stored in the same "
        "folder as app.py. Each department keeps its own KPI names and weights."
    )

    local_files = discover_local_kpi_files()
    if local_files:
        st.write("### Local KPI files")
        status_rows = []
        for dept_name, path in local_files:
            status_rows.append({
                "Department": dept_name,
                "Workbook": path.name,
                "Location": str(path),
                "Found": "Yes",
            })
        st.dataframe(
            pd.DataFrame(status_rows),
            use_container_width=True,
            hide_index=True,
        )

        if st.button("Re-sync local KPI workbooks", type="primary"):
            sync_local_kpi_workbooks(force=True)
            st.rerun()
    else:
        st.warning(
            "No Service, Bar, KTN or H_K KPI workbook was found beside app.py."
        )

    st.divider()
    st.write("### Optional manual import")
    st.caption(
        "Use this only when testing a workbook that is not stored in the project folder."
    )

    uploads = st.file_uploader(
        "Upload KPI workbooks",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="kpi_uploads",
    )

    if uploads and st.button("Import manual KPI workbook(s)"):
        all_templates = []
        all_roster = []
        source_files = []
        errors = []

        for upload in uploads:
            try:
                templates, staff = parse_kpi_workbook(upload)
                all_templates.extend(templates)
                all_roster.extend(staff)
                source_files.append(upload.name)
                st.success(
                    f"{upload.name}: {len(staff)} employees / "
                    f"{len(set((t['department'], t['grp'], t['kpi']) for t in templates))} "
                    "KPI definitions"
                )
            except Exception as exc:
                errors.append(f"{upload.name}: {exc}")

        if all_templates:
            save_kpi_configuration(
                all_templates,
                all_roster,
                source_files=source_files,
                departments={t["department"] for t in all_templates},
            )
            st.success("Manual KPI configuration and roster updated.")

        for error in errors:
            st.error(error)

    templates = load_templates()

    if not templates.empty:
        st.write("### Active KPI structure")

        summary = (
            templates.groupby(["department", "grp"])
            .agg(
                KPIs=("kpi", "count"),
                Weight=("weight", "sum"),
            )
            .reset_index()
        )
        summary["Weight %"] = summary["Weight"] * 100
        summary = summary.drop(columns=["Weight"])
        st.dataframe(summary, use_container_width=True, hide_index=True)

        dept = st.selectbox(
            "View department",
            sorted(templates["department"].unique()),
        )

        dept_groups = sorted(
            templates.loc[
                templates["department"] == dept, "grp"
            ].unique()
        )
        grp = st.selectbox("View group", dept_groups)

        detail = templates[
            (templates["department"] == dept) &
            (templates["grp"] == grp)
        ].copy()
        detail["Weight %"] = detail["weight"] * 100

        st.dataframe(
            detail[["kpi", "Weight %", "source_file"]],
            use_container_width=True,
            hide_index=True,
        )

        total_weight = detail["weight"].sum() * 100
        if abs(total_weight - 100) > 0.01:
            st.warning(
                f"This source KPI structure totals {total_weight:.1f}%, "
                "so review the workbook before using it for bonuses."
            )
        else:
            st.success("KPI weights total 100%.")

# ============================================================

# TAB 2 — DASHBOARD
# ============================================================

with tabs[1]:
    st.subheader("Daily performance dashboard")

    roster = load_roster()

    if roster.empty:
        st.warning("Load the KPI workbooks first.")
    else:
        sales = pd.read_sql_query(
            "SELECT * FROM sales WHERE work_date=?",
            con,
            params=(str(work_date),)
        )

        scores = pd.read_sql_query(
            "SELECT * FROM kpi_scores WHERE work_date=?",
            con,
            params=(str(work_date),)
        )

        rows = []

        for _, person in roster.iterrows():
            kpi = weighted_kpi_score(
                scores,
                person["department"],
                person["grp"],
                person["name"]
            )

            sr = sales[sales["name"] == person["name"]]
            sale = float(sr["sales"].iloc[0]) if not sr.empty else 0.0

            rows.append([
                person["name"],
                person["department"],
                person["grp"],
                sale,
                kpi
            ])

        df = pd.DataFrame(
            rows,
            columns=[
                "Name", "Department", "Group",
                "Sales", "KPI Score"
            ]
        )

        c1, c2, c3, c4 = st.columns(4)

        c1.metric("Employees", len(df))
        c2.metric("Sales loaded", f"{df['Sales'].sum():,.0f}")
        c3.metric(
            "Average KPI",
            f"{df['KPI Score'].mean() * 100:.1f}%"
        )
        c4.metric(
            "Highest KPI",
            f"{df['KPI Score'].max() * 100:.1f}%"
        )

        st.dataframe(
            df.sort_values(
                ["Department", "KPI Score"],
                ascending=[True, False]
            ),
            use_container_width=True,
            hide_index=True
        )

        st.bar_chart(
            df.sort_values("KPI Score", ascending=False)
            .set_index("Name")["KPI Score"]
            .head(20)
        )


# ============================================================
# TAB 3 — DAILY SALES
# ============================================================

with tabs[2]:
    st.subheader("Upload daily sales")

    st.write(
        "Accepted: Excel or CSV. Employee names are matched against "
        "the imported KPI roster. Duplicate sales rows are aggregated."
    )

    f = st.file_uploader(
        "Sales file",
        type=["xlsx", "xls", "csv"],
        key="sales_upload"
    )

    if f and st.button("Import Sales", type="primary"):
        try:
            raw, matched = import_sales(f, work_date)

            st.success(
                f"Imported sales for {len(matched)} matched employees."
            )

            unmatched = raw[raw["matched_name"].isna()]

            if len(unmatched):
                st.warning(
                    f"{len(unmatched)} rows were not matched to the roster."
                )
                st.dataframe(
                    unmatched,
                    use_container_width=True,
                    hide_index=True
                )

            st.write("### Import preview")
            st.dataframe(
                raw,
                use_container_width=True,
                hide_index=True
            )

        except Exception as e:
            st.error(str(e))


# ============================================================
# TAB 4 — KPI SCORING
# ============================================================

with tabs[3]:
    st.subheader("Supervisor KPI entry")

    roster = load_roster()

    if roster.empty:
        st.warning("Load the KPI workbooks first.")
    else:
        dept = st.selectbox(
            "Department",
            sorted(roster["department"].unique()),
            key="score_dept"
        )

        groups = sorted(
            roster.loc[
                roster["department"] == dept, "grp"
            ].unique()
        )

        grp = st.selectbox(
            "Group",
            groups,
            key="score_grp"
        )

        team = roster[
            (roster["department"] == dept) &
            (roster["grp"] == grp)
        ]

        mp = kpi_map(dept, grp)

        if not mp:
            st.error("No KPI definition exists for this department/group.")
        else:
            st.write(
                f"**{grp}** — {len(team)} team members"
            )

            st.caption(
                "Scoring follows the supplied KPI workbooks: "
                "1.0 / 0.8 / 0.6 / 0.4 / 0.0. "
                "The resulting daily KPI is capped at 95% unless justified."
            )

            with st.form(f"kpi_form_{dept}_{grp}_{work_date}"):
                selected = {}
                comments = {}

                for _, person in team.iterrows():
                    name = person["name"]
                    st.markdown(f"**{name}**")

                    cols = st.columns(min(5, len(mp)))

                    for idx, (kpi, weight) in enumerate(mp.items()):
                        with cols[idx % len(cols)]:
                            selected.setdefault(name, {})[kpi] = st.selectbox(
                                f"{kpi} ({weight * 100:.1f}%)",
                                DEFAULT_SCORES,
                                index=2,
                                key=f"{work_date}_{dept}_{grp}_{name}_{kpi}"
                            )

                    comments[name] = st.text_area(
                        f"Comment for {name} — required for any 0.0",
                        key=f"comment_{work_date}_{dept}_{grp}_{name}"
                    )

                if st.form_submit_button(
                    "Save KPI Scores",
                    type="primary"
                ):
                    problems = validate_score_comments(
                        team,
                        selected,
                        {
                            (name, kpi): comments.get(name, "")
                            for name in selected
                            for kpi in selected[name]
                        }
                    )

                    if problems:
                        for p in problems:
                            st.error(p)
                    else:
                        for name, kpis in selected.items():
                            for kpi, value in kpis.items():
                                con.execute("""
                                    INSERT OR REPLACE INTO kpi_scores
                                    (work_date, name, department, grp, kpi, score, comment)
                                    VALUES (?, ?, ?, ?, ?, ?, ?)
                                """, (
                                    str(work_date),
                                    name,
                                    dept,
                                    grp,
                                    kpi,
                                    float(value),
                                    comments.get(name, "")
                                ))

                        con.commit()
                        st.success("KPI scores saved successfully.")


# ============================================================
# TAB 5 — TARGETS & BONUS
# ============================================================

with tabs[4]:
    st.subheader("Targets and bonus pool")

    st.info(
        "This section deliberately keeps sales achievement and KPI scoring "
        "separate. You can see exactly how a bonus is produced rather than "
        "hiding the calculation inside the dashboard."
    )

    roster = load_roster()

    if roster.empty:
        st.warning("Load the KPI workbooks first.")
    else:
        st.write("### Employee sales targets")

        target_dept = st.selectbox(
            "Department",
            sorted(roster["department"].unique()),
            key="target_dept"
        )

        target_team = roster[
            roster["department"] == target_dept
        ].copy()

        with st.form("targets_form"):
            for _, person in target_team.iterrows():
                existing = pd.read_sql_query(
                    "SELECT target FROM targets WHERE period=? AND name=?",
                    con,
                    params=(period, person["name"])
                )

                current = (
                    float(existing["target"].iloc[0])
                    if not existing.empty else 0.0
                )

                target = st.number_input(
                    f"{person['name']} — target",
                    min_value=0.0,
                    value=current,
                    step=1000.0,
                    key=f"target_{period}_{person['name']}"
                )

                if st.session_state.get(
                    f"_target_values_{period}"
                ) is None:
                    pass

            if st.form_submit_button("Save Targets"):
                # Read the widgets from session state.
                for _, person in target_team.iterrows():
                    key = f"target_{period}_{person['name']}"
                    target = float(st.session_state.get(key, 0.0))
                    save_target(
                        period,
                        person["name"],
                        person["department"],
                        person["grp"],
                        target,
                        "individual"
                    )

                st.success("Targets saved.")

        st.divider()

        st.write("### Bonus pools")

        pool_dept = st.selectbox(
            "Bonus pool department",
            sorted(roster["department"].unique()),
            key="pool_dept"
        )

        pool_groups = sorted(
            roster.loc[
                roster["department"] == pool_dept, "grp"
            ].unique()
        )

        pool_grp = st.selectbox(
            "Bonus pool group",
            pool_groups,
            key="pool_grp"
        )

        existing_pool = pd.read_sql_query(
            """
            SELECT * FROM bonus_pools
            WHERE period=? AND department=? AND grp=?
            """,
            con,
            params=(period, pool_dept, pool_grp)
        )

        current_pool = (
            float(existing_pool["pool"].iloc[0])
            if not existing_pool.empty else 0.0
        )

        current_target = (
            float(existing_pool["target"].iloc[0])
            if not existing_pool.empty else 0.0
        )

        pool = st.number_input(
            "Bonus pool amount",
            min_value=0.0,
            value=current_pool,
            step=1000.0
        )

        dept_target = st.number_input(
            "Department/group target",
            min_value=0.0,
            value=current_target,
            step=1000.0
        )

        pool_type = st.selectbox(
            "Pool target basis",
            ["department", "individual"],
            index=0,
            help=(
                "Use individual for employees with individual sales targets. "
                "Use department for teams whose sales target is achieved collectively."
            )
        )

        if st.button("Save Bonus Pool", type="primary"):
            save_bonus_pool(
                period,
                pool_dept,
                pool_grp,
                pool,
                dept_target,
                pool_type
            )
            st.success("Bonus pool saved.")

        st.divider()

        st.write("### Current bonus calculation")

        bonus_df = calculate_bonus(period, work_date)

        if bonus_df.empty:
            st.info("No bonus data yet.")
        else:
            display = bonus_df.copy()
            display["Sales Achievement"] = (
                display["Sales Achievement"] * 100
            ).round(1)
            display["KPI Score"] = (
                display["KPI Score"] * 100
            ).round(1)

            st.dataframe(
                display,
                use_container_width=True,
                hide_index=True
            )

            st.caption(
                "Individual model: Pool × capped sales achievement × KPI. "
                "Department model: Pool × capped department achievement, "
                "then distributed by each employee's KPI share."
            )


# ============================================================
# TAB 6 — REMARKS
# ============================================================

with tabs[5]:
    st.subheader("Supervisor remarks")

    roster = load_roster()

    if not roster.empty:
        grp = st.selectbox(
            "Group",
            sorted(roster["grp"].unique()),
            key="remark_group"
        )

        team = roster[roster["grp"] == grp]

        person = st.selectbox(
            "Team member",
            team["name"].tolist(),
            key="remark_person"
        )

        supervisor = st.text_input("Supervisor", key="remark_supervisor")
        remark = st.text_area("Remark", key="remark_text")

        if st.button("Save Remark"):
            con.execute("""
                INSERT OR REPLACE INTO remarks
                (work_date, grp, name, supervisor, remark)
                VALUES (?, ?, ?, ?, ?)
            """, (
                str(work_date),
                grp,
                person,
                supervisor,
                remark
            ))
            con.commit()
            st.success("Remark saved.")

    google = pull_google_remarks()

    if not google.empty:
        st.subheader("Google remarks feed")
        st.dataframe(
            google.tail(50),
            use_container_width=True,
            hide_index=True
        )
    elif GOOGLE_SHEET_CSV_URL:
        st.info("No Google remarks found.")
    else:
        st.caption(
            "Set GOOGLE_SHEET_CSV_URL to a published Google Sheet CSV endpoint "
            "to sync supervisor remarks."
        )


# ============================================================
# TAB 7 — EXCEL EXPORT
# ============================================================

with tabs[6]:
    st.subheader("Excel dashboard export")

    st.write(
        "The export contains Dashboard, KPI Detail, KPI Templates, "
        "Sales, Bonus Calculation, Remarks and Charts."
    )

    if st.button("Build Excel Dashboard", type="primary"):
        try:
            data = export_excel(work_date, period)

            st.download_button(
                "Download Excel Dashboard",
                data=data,
                file_name=f"KPI_Bonus_Dashboard_{period}_{work_date}.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )
            )

        except Exception as e:
            st.error(f"Could not build Excel dashboard: {e}")
