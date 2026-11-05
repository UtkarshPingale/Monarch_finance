from turtle import title
import os, socket, socketserver, webbrowser, traceback, secrets
import csv, io, math, hashlib
from http.server   import SimpleHTTPRequestHandler
from urllib.parse  import urlparse, parse_qs
from datetime      import date, datetime, timedelta
import cgi, tempfile
from click import style
from sqlalchemy              import create_engine, text
from reportlab.platypus      import SimpleDocTemplate, Paragraph, Table, Spacer
from reportlab.lib.styles    import getSampleStyleSheet
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen        import canvas as rl_canvas
from reportlab.lib.colors    import Color, grey
from openpyxl                import Workbook


# ═══════════════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════════════

DB_USER = "postgres"
DB_PASS = "postgres123"
DB_HOST = "192.168.1.149"  
DB_PORT = "5432"
DB_NAME = "postgres"

SCHEMA_NAME = "monarch_db"           

ENGINE = create_engine(
    f"postgresql+psycopg2://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}",
    connect_args={"options": f"-csearch_path={SCHEMA_NAME}"},
    future=True,
    pool_pre_ping=True,
)

SESSION_COOKIE   = "monarch_session"
SESSION_LIFETIME = 60 * 60 * 8
SESSIONS: dict   = {}
TOKEN            = secrets.token_hex(8)

C = {
    "primary":      "#000e24",
    "secondary":    "#775a19",
    "bg":           "#f7f9fb",
    "surface_low":  "#f2f4f6",
    "surface_high": "#e6e8ea",
    "outline":      "#c4c6d0",
    "on_surface":   "#191c1e",
    "on_var":       "#43474e",
    "error":        "#ba1a1a",
    "sec_cont":     "#fed488",
    "dark":         "#1a2540",
    "mid":          "#2c3e6b",
}

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
LOGO_URL  = "/logo.jpg"
LOGO_PATH = r"Z:\Personal\Devloper\EPD\logo.jpg"
REF_TABLE_MAP = {
    "lob":    ("ref_lob", "name"),
    "biz":    ("ref_business_line", "name"),
    "pm":     ("ref_project_managers", "name"),
    "status": ("ref_project_status", "name"),
    "doccat": ("ref_doc_category", "name"),
    "expcat": ("ref_expense_category", "name"),
}

# ═══════════════════════════════════════════════════════════════
#  AUTH
# ═══════════════════════════════════════════════════════════════

def _hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

def verify_user(username: str, password: str):
    try:
        with ENGINE.connect() as conn:
            row = conn.execute(
                text("SELECT id,username,password_hash,display_name,role "
                     "FROM users WHERE username=:u"), {"u": username}
            ).fetchone()
        if row and row[2] == _hash_pw(password):
            return {"id": row[0], "username": row[1],
                    "display_name": row[3] or row[1], "role": row[4]}
    except Exception:
        pass
    return None

def create_session(user: dict) -> str:
    token = secrets.token_hex(24)
    SESSIONS[token] = {**user, "expires": datetime.now() + timedelta(seconds=SESSION_LIFETIME)}
    return token

def destroy_session(token: str):
    SESSIONS.pop(token, None)
def reset_user_password(username: str, new_password: str) -> bool:
    try:
        with ENGINE.begin() as conn:
            row = conn.execute(text("SELECT id FROM users WHERE username=:u"), {"u": username}).fetchone()
            if not row:
                return False
            conn.execute(text("UPDATE users SET password_hash=:ph WHERE username=:u"),
                         {"ph": _hash_pw(new_password), "u": username})
        return True
    except Exception:
        return False
def _get_session_token(handler) -> str | None:
    for part in handler.headers.get("Cookie", "").split(";"):
        p = part.strip()
        if p.startswith(SESSION_COOKIE + "="):
            return p.split("=", 1)[1].strip()
    return None

def get_session_user(handler) -> str | None:
    token = _get_session_token(handler)
    if not token: return None
    sess  = SESSIONS.get(token)
    if not sess: return None
    if sess["expires"] < datetime.now():
        SESSIONS.pop(token, None)
        return None
    return sess.get("display_name") or sess.get("username")


# ═══════════════════════════════════════════════════════════════
#  DATABASE — SCHEMA INIT
# ═══════════════════════════════════════════════════════════════
def migrate_db():
    """Add any missing columns to existing tables."""
    migrations = [
        ("projects", "variation_amt",                "ALTER TABLE projects ADD COLUMN IF NOT EXISTS variation_amt REAL DEFAULT 0"),
        ("projects", "revised_order_value",          "ALTER TABLE projects ADD COLUMN IF NOT EXISTS revised_order_value REAL DEFAULT 0"),
        ("projects", "hier_business_line",           "ALTER TABLE projects ADD COLUMN IF NOT EXISTS hier_business_line TEXT"),
        ("projects", "parent_project",               "ALTER TABLE projects ADD COLUMN IF NOT EXISTS parent_project TEXT"),
        ("projects", "child_project",                "ALTER TABLE projects ADD COLUMN IF NOT EXISTS child_project TEXT"),
        ("projects", "baseline_eac",                 "ALTER TABLE projects ADD COLUMN IF NOT EXISTS baseline_eac REAL DEFAULT 0"),
        ("projects", "baseline_margin",              "ALTER TABLE projects ADD COLUMN IF NOT EXISTS baseline_margin REAL DEFAULT 0"),
        ("projects", "rev_latest_eac",               "ALTER TABLE projects ADD COLUMN IF NOT EXISTS rev_latest_eac REAL DEFAULT 0"),
        ("projects", "rev_latest_margin",            "ALTER TABLE projects ADD COLUMN IF NOT EXISTS rev_latest_margin REAL DEFAULT 0"),
        ("projects", "physical_completion",          "ALTER TABLE projects ADD COLUMN IF NOT EXISTS physical_completion REAL DEFAULT 0"),
        ("projects", "billed_amount",                "ALTER TABLE projects ADD COLUMN IF NOT EXISTS billed_amount REAL DEFAULT 0"),
        ("projects", "last_billed_date",             "ALTER TABLE projects ADD COLUMN IF NOT EXISTS last_billed_date TEXT"),
        ("projects", "collection_till_date",         "ALTER TABLE projects ADD COLUMN IF NOT EXISTS collection_till_date REAL DEFAULT 0"),
        ("projects", "itd_cost",                     "ALTER TABLE projects ADD COLUMN IF NOT EXISTS itd_cost REAL DEFAULT 0"),
        ("projects", "bill_submission_pending_date", "ALTER TABLE projects ADD COLUMN IF NOT EXISTS bill_submission_pending_date TEXT"),
        ("projects", "project_type",                 "ALTER TABLE projects ADD COLUMN IF NOT EXISTS project_type TEXT DEFAULT 'parent'"),
        ("projects", "parent_id",                    "ALTER TABLE projects ADD COLUMN IF NOT EXISTS parent_id INTEGER"),
        ("projects", "folder_path", "ALTER TABLE projects ADD COLUMN IF NOT EXISTS folder_path TEXT"),
    ]
    with ENGINE.begin() as conn:
        for table, col, sql in migrations:
            try:
                conn.execute(text(sql))
                print(f"[migrate] Ensured column: {table}.{col}")
            except Exception as e:
                print(f"[migrate] Skipped {table}.{col}: {e}")

def init_db():
    with ENGINE.begin() as conn:

        conn.execute(text('CREATE SCHEMA IF NOT EXISTS "monarch_db"'))
        
        migrate_db()

        result = conn.execute(text("""
            SELECT schema_name
            FROM information_schema.schemata
            WHERE schema_name='monarch_db'
        """)).fetchall()

        print("Schema Check:", result)

        conn.execute(text('SET search_path TO "monarch_db"'))

        print(
            "Current Schema:",
            conn.execute(text("SELECT current_schema()")).scalar()
        )

        # CREATE TABLE statements here
    """Create all tables and seed reference data if empty."""
    with ENGINE.begin() as conn:

        conn.execute(text(f"""
            CREATE TABLE IF NOT EXISTS {SCHEMA_NAME}.users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                display_name TEXT,
                role TEXT DEFAULT 'user',
                created_at TEXT
            )
        """))
        conn.execute(text("""
    CREATE TABLE IF NOT EXISTS erp_updates (
        id          SERIAL PRIMARY KEY,
        project_id  INTEGER REFERENCES projects(id) ON DELETE CASCADE,
        category    TEXT NOT NULL,
        amount      REAL DEFAULT 0,
        created_by  TEXT,
        created_at  TEXT
    )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_lob (
                id     SERIAL PRIMARY KEY,
                name   TEXT UNIQUE NOT NULL,
                active BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_project_managers (
                id           SERIAL PRIMARY KEY,
                name TEXT UNIQUE NOT NULL,
                email        TEXT,
                active       BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_project_status (
                id     SERIAL PRIMARY KEY,
                name   TEXT UNIQUE NOT NULL,
                color  TEXT DEFAULT '#43474e',
                active BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_business_line (
                id     SERIAL PRIMARY KEY,
                name   TEXT UNIQUE NOT NULL,
                active BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_doc_category (
                id     SERIAL PRIMARY KEY,
                name   TEXT UNIQUE NOT NULL,
                active BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_expense_category (
                id       SERIAL PRIMARY KEY,
                name     TEXT UNIQUE NOT NULL,
                color_bg TEXT DEFAULT '#f5f5f5',
                color_fg TEXT DEFAULT '#374151',
                active   BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_clients (
                id     SERIAL PRIMARY KEY,
                name   TEXT UNIQUE NOT NULL,
                active BOOLEAN DEFAULT TRUE
            )"""))

        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS projects (
                id                   SERIAL PRIMARY KEY,
                co_no                TEXT,
                project_name         TEXT,
                business_line        TEXT,
                project_manager      TEXT,
                client_name          TEXT,
                order_number         TEXT,
                order_issue_date     TEXT,
                order_value          REAL DEFAULT 0,
                variation_amt        REAL DEFAULT 0,
                revised_order_value  REAL DEFAULT 0,
                hier_business_line   TEXT,
                parent_project       TEXT,
                child_project        TEXT,
                baseline_eac         REAL DEFAULT 0,
                baseline_margin      REAL DEFAULT 0,
                rev_latest_eac       REAL DEFAULT 0,
                rev_latest_margin    REAL DEFAULT 0,
                project_status       TEXT DEFAULT 'Active',
                physical_completion  REAL DEFAULT 0,
                billed_amount        REAL DEFAULT 0,
                last_billed_date     TEXT,
                collection_till_date REAL DEFAULT 0,
                itd_cost             REAL DEFAULT 0,
                bill_submission_pending_date TEXT,
                project_type         TEXT DEFAULT 'parent',
                parent_id            INTEGER,
                created_at           TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS documents (
                id           SERIAL PRIMARY KEY,
                project_id   INTEGER REFERENCES projects(id) ON DELETE CASCADE,
                doc_name     TEXT,
                doc_category TEXT,
                doc_link     TEXT,
                created_at   TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS expenditures (
                id              SERIAL PRIMARY KEY,
                project_id      INTEGER REFERENCES projects(id) ON DELETE CASCADE,
                lob             TEXT,
                data_as_of      TEXT,
                erp_update_date TEXT,
                itd_cogs        REAL DEFAULT 0,
                itd_djcs        REAL DEFAULT 0,
                itd_equip       REAL DEFAULT 0,
                itd_manpower    REAL DEFAULT 0,
                itd_vehicles    REAL DEFAULT 0,
                itd_others      REAL DEFAULT 0,
                erp_cogs        REAL DEFAULT 0,
                erp_djcs        REAL DEFAULT 0,
                erp_equip       REAL DEFAULT 0,
                erp_manpower    REAL DEFAULT 0,
                erp_vehicles    REAL DEFAULT 0,
                erp_others      REAL DEFAULT 0,
                created_at      TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS expense_entries (
                id             SERIAL PRIMARY KEY,
                expenditure_id INTEGER,
                project_id     INTEGER REFERENCES projects(id) ON DELETE CASCADE,
                category       TEXT,
                description    TEXT,
                quantity       REAL DEFAULT 0,
                unit_cost      REAL DEFAULT 0,
                line_total     REAL DEFAULT 0,
                created_at     TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS variations (
                id             SERIAL PRIMARY KEY,
                project_id     INTEGER REFERENCES projects(id) ON DELETE CASCADE,
                child_co_no    TEXT,
                child_name     TEXT,
                parent_name    TEXT,
                original_value REAL DEFAULT 0,
                variation_amt  REAL DEFAULT 0,
                revised_value  REAL DEFAULT 0,
                variation_date TEXT,
                status         TEXT DEFAULT 'Pending',
                created_at     TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS billing_collection (
                id              SERIAL PRIMARY KEY,
                project_id      INTEGER REFERENCES projects(id) ON DELETE CASCADE,
                ra_no           TEXT,
                billing_date    TEXT,
                billed_amt      REAL DEFAULT 0,
                collection_date TEXT,
                collection_amt  REAL DEFAULT 0,
                created_at      TEXT
            )"""))
        exists = conn.execute(text("SELECT id FROM users WHERE username='admin'")).fetchone()
        if not exists:
            conn.execute(text("""
                INSERT INTO users (username,password_hash,display_name,role,created_at)
                VALUES ('admin',:ph,'Admin','superadmin',:ca)
            """), {"ph": _hash_pw("monarch123"), "ca": datetime.now().isoformat()})

        _seed_ref(conn, "ref_lob", "name", [
            "Railways", "DPR-Highways", "Land Acquisition", "Locals",
            "Geotechnical", "Geospatial", "Project Management", "ELU preparation",
        ])
        _seed_ref(conn, "ref_project_managers", "name", [
            "Arjun Sharma", "Priya Nair", "Rajesh Kumar",
            "Anita Singh", "Vikram Mehta", "Sanjay Datta", "Meera Iyer",
        ])
        _seed_ref(conn, "ref_project_status", "name", [
            "Active", "Closed", "Execution Closed", "Cancelled", "Hold",
        ])
        _seed_ref(conn, "ref_business_line", "name", [
            "Construction", "Engineering", "Surveying",
        ])
        _seed_ref(conn, "ref_doc_category", "name", [
            "Contract", "Technical Drawing", "Site Photo", "Permit", "Report", "Other",
        ])
        _seed_ref_expense(conn)
        _seed_ref(conn, "ref_clients", "name", [
            "National Highways Authority", "NHAI India", "Urban Dev Authority",
            "Apex Real Estate", "Water Resources Board",
        ])


def _seed_ref(conn, table: str, col: str, values: list):
    existing = conn.execute(text(f"SELECT {col} FROM {table}")).fetchall()
    existing_vals = {r[0] for r in existing}
    for v in values:
        if v not in existing_vals:
            conn.execute(text(f"INSERT INTO {table} ({col}) VALUES (:v)"), {"v": v})


def _seed_ref_expense(conn):
    existing = conn.execute(text("SELECT name FROM ref_expense_category")).fetchall()
    if existing: return
    defaults = [
        ("COGS",      "#e8edf7", "#00234b"),
        ("DJCS",      "#fdf3e0", "#775a19"),
        ("Equipment", "#eaf4eb", "#166534"),
        ("Manpower",  "#fde8e8", "#9b1c1c"),
        ("Vehicles",  "#f0e8fd", "#5b21b6"),
        ("Others",    "#f5f5f5", "#374151"),
    ]
    for name, bg, fg in defaults:
        conn.execute(text("""
            INSERT INTO ref_expense_category (name,color_bg,color_fg)
            VALUES (:n,:bg,:fg)
        """), {"n": name, "bg": bg, "fg": fg})


# ═══════════════════════════════════════════════════════════════
#  REFERENCE DATA FETCHERS
# ═══════════════════════════════════════════════════════════════

def _fetch_list(table: str, col: str = "name", order: str = "name") -> list[str]:
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                f"SELECT {col} FROM {table} WHERE active IS NOT FALSE ORDER BY {order}"
            )).fetchall()
        return [r[0] for r in rows]
    except Exception as e:
        print(f"[_fetch_list] {table}.{col} error: {e}")
        return []

def get_lob_list()           -> list[str]: return _fetch_list("ref_lob")
def get_project_managers()   -> list[str]: return _fetch_list("ref_project_managers", "name")
def get_project_statuses()   -> list[str]: return _fetch_list("ref_project_status")
def get_business_lines()     -> list[str]: return _fetch_list("ref_business_line")
def get_doc_categories()     -> list[str]: return _fetch_list("ref_doc_category")
def get_expense_categories() -> list[str]: return _fetch_list("ref_expense_category")
def get_client_list()        -> list[str]: return _fetch_list("ref_clients")

def get_expense_category_colors() -> dict:
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                "SELECT name,color_bg,color_fg FROM ref_expense_category WHERE active=TRUE"
            )).fetchall()
        return {r[0]: (r[1], r[2]) for r in rows}
    except Exception:
        return {}

def get_parent_project_names() -> list[str]:
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                "SELECT project_name FROM projects WHERE project_type='parent' ORDER BY project_name"
            )).fetchall()
        return [r[0] for r in rows]
    except Exception:
        return []

def get_child_project_names(parent_name: str = "") -> list[str]:
    try:
        with ENGINE.connect() as conn:
            if parent_name:
                rows = conn.execute(text(
                    "SELECT project_name FROM projects WHERE parent_project=:p ORDER BY project_name"
                ), {"p": parent_name}).fetchall()
            else:
                rows = conn.execute(text(
                    "SELECT project_name FROM projects WHERE project_type='child' ORDER BY project_name"
                )).fetchall()
        return [r[0] for r in rows]
    except Exception:
        return []


def add_ref_value(table: str, col: str, value: str) -> bool:
    try:
        with ENGINE.begin() as conn:
            conn.execute(text(
                f"INSERT INTO {table} ({col}) VALUES (:v) ON CONFLICT ({col}) DO NOTHING"
            ), {"v": value})
        return True
    except Exception:
        return False

def delete_ref_value(table: str, col: str, value: str) -> bool:
    try:
        with ENGINE.begin() as conn:
            conn.execute(text(f"UPDATE {table} SET active=FALSE WHERE {col}=:v"), {"v": value})
        return True
    except Exception:
        return False


# ═══════════════════════════════════════════════════════════════
#  CORE DB OPERATIONS
# ═══════════════════════════════════════════════════════════════

def _n(v, default=0.0):
    try:    return float(v or default)
    except: return default
def _rev_ov(p: dict) -> float:
    """Revised Order Value — falls back to Order Value + Variation if not set."""
    return _n(p.get("revised_order_value")) or (_n(p.get("order_value")) + _n(p.get("variation_amt")))

def _rows_to_dicts(conn, sql: str, params: dict = None) -> list[dict]:
    result = conn.execute(text(sql), params or {})
    keys   = list(result.keys())
    return [dict(zip(keys, row)) for row in result.fetchall()]


def save_project(fields: dict) -> int:
    num = ["order_value","variation_amt","revised_order_value","baseline_eac","baseline_margin",
           "rev_latest_eac","rev_latest_margin","physical_completion","billed_amount",
           "collection_till_date","itd_cost"]
    for f in num: fields[f] = _n(fields.get(f))

    # auto-calc revised_order_value if not set
    if not fields["revised_order_value"]:
        fields["revised_order_value"] = fields["order_value"] + fields["variation_amt"]

    text_cols = ["co_no","project_name","business_line","project_manager","client_name",
                 "order_number","order_issue_date","hier_business_line",
                 "parent_project","child_project","project_status",
                 "last_billed_date","bill_submission_pending_date"]
    for f in text_cols: fields.setdefault(f, "")
 
    fields["created_at"] = datetime.now().isoformat()
    client = fields.get("client_name", "").strip()
    if client: add_ref_value("ref_clients", "name", client)
 
    parent_val     = fields.get("parent_project", "").strip()
    submitted_type = fields.get("project_type", "").strip()
 
    PLACEHOLDER_VALUES = {"", "select or type parent…", "none", "n/a"}
    if parent_val and parent_val.lower() not in PLACEHOLDER_VALUES:
        fields["project_type"] = "child"
    elif submitted_type == "child":
        fields["project_type"] = "child"
    else:
        fields["project_type"] = "parent"
        fields["parent_project"] = ""
 
    new_id = None
    with ENGINE.begin() as conn:
        row = conn.execute(text("""
            INSERT INTO projects (
                co_no, project_name, business_line, project_manager, client_name,
                order_number, order_issue_date, order_value, variation_amt, revised_order_value,
                hier_business_line, parent_project, child_project,
                baseline_eac, baseline_margin, rev_latest_eac, rev_latest_margin,
                project_status, physical_completion, billed_amount, last_billed_date,
                collection_till_date, itd_cost, bill_submission_pending_date,
                project_type, folder_path, created_at
            ) VALUES (
                :co_no, :project_name, :business_line, :project_manager, :client_name,
                :order_number, :order_issue_date, :order_value, :variation_amt, :revised_order_value,
                :hier_business_line, :parent_project, :child_project,
                :baseline_eac, :baseline_margin, :rev_latest_eac, :rev_latest_margin,
                :project_status, :physical_completion, :billed_amount, :last_billed_date,
                :collection_till_date, :itd_cost, :bill_submission_pending_date,
                :project_type, :folder_path, :created_at
            ) RETURNING id"""), fields)
        new_id = row.fetchone()[0]
 
    child_name = fields.get("project_name", "").strip()
    if fields["project_type"] == "child" and parent_val and child_name and new_id:
        with ENGINE.begin() as conn2:
            conn2.execute(text("""
                UPDATE projects
                SET child_project = CASE
                    WHEN child_project IS NULL OR child_project = ''
                        THEN :child
                    WHEN child_project NOT LIKE '%' || :child || '%'
                        THEN child_project || ', ' || :child
                    ELSE child_project
                END
                WHERE project_name = :parent
                AND   project_type = 'parent'
            """), {"child": child_name, "parent": parent_val})
 
    return new_id


def update_project(pid: int, fields: dict):
    num = ["order_value","variation_amt","revised_order_value","baseline_eac","baseline_margin",
           "rev_latest_eac","rev_latest_margin","physical_completion","billed_amount",
           "collection_till_date","itd_cost"]
    for f in num: fields[f] = _n(fields.get(f))

    if not fields["revised_order_value"]:
        fields["revised_order_value"] = fields["order_value"] + fields["variation_amt"]

    text_cols = ["co_no","project_name","business_line","project_manager","client_name",
                 "order_number","order_issue_date","hier_business_line",
                 "parent_project","child_project","project_status",
                 "last_billed_date","bill_submission_pending_date"]
    for f in text_cols: fields.setdefault(f, "")

    fields["folder_path"] = (fields.get("folder_path_check") or fields.get("folder_path") or "").strip()

    fields["id"] = pid
    client = fields.get("client_name", "").strip()
    if client: add_ref_value("ref_clients", "name", client)
 
    parent_val     = fields.get("parent_project", "").strip()
    submitted_type = fields.get("project_type", "").strip()
 
    PLACEHOLDER_VALUES = {"", "select or type parent…", "none", "n/a"}
    if parent_val and parent_val.lower() not in PLACEHOLDER_VALUES:
        fields["project_type"] = "child"
    elif submitted_type == "child":
        fields["project_type"] = "child"
    else:
        fields["project_type"] = "parent"
        fields["parent_project"] = ""
 
    with ENGINE.begin() as conn:
        conn.execute(text("""
            UPDATE projects SET
                co_no=:co_no,
                project_name=:project_name,
                project_type=:project_type,
                business_line=:business_line,
                project_manager=:project_manager,
                client_name=:client_name,
                order_number=:order_number,
                order_issue_date=:order_issue_date,
                order_value=:order_value,
                variation_amt=:variation_amt,
                revised_order_value=:revised_order_value,
                hier_business_line=:hier_business_line,
                parent_project=:parent_project,
                child_project=:child_project,
                project_status=:project_status,
                physical_completion=:physical_completion,
                billed_amount=:billed_amount,
                last_billed_date=:last_billed_date,
                collection_till_date=:collection_till_date,
                itd_cost=:itd_cost,
                bill_submission_pending_date=:bill_submission_pending_date,
                rev_latest_eac=:rev_latest_eac,
                rev_latest_margin=:rev_latest_margin,
                baseline_eac=:baseline_eac,
                baseline_margin=:baseline_margin
                folder_path=:folder_path
            WHERE id=:id"""), fields)
 
    child_name = fields.get("project_name", "").strip()
    if fields["project_type"] == "child" and parent_val and child_name:
        with ENGINE.begin() as conn2:
            conn2.execute(text("""
                UPDATE projects
                SET child_project = CASE
                    WHEN child_project IS NULL OR child_project = '' THEN :child
                    WHEN child_project NOT LIKE '%' || :child || '%' THEN child_project || ', ' || :child
                    ELSE child_project
                END
                WHERE project_name = :parent AND project_type = 'parent'
            """), {"child": child_name, "parent": parent_val})

def delete_project(pid: int):
    with ENGINE.begin() as conn:
        conn.execute(text("DELETE FROM projects WHERE id=:id"), {"id": pid})


def fetch_all_projects(lob_filter: str = "") -> list[dict]:
    try:
        with ENGINE.connect() as conn:
            if lob_filter and lob_filter != "All":
                return _rows_to_dicts(conn,
                    "SELECT * FROM projects WHERE hier_business_line=:l OR business_line=:l ORDER BY id DESC",
                    {"l": lob_filter})
            else:
                return _rows_to_dicts(conn, "SELECT * FROM projects ORDER BY id DESC")
    except Exception:
        return []


def fetch_project_by_id(pid: int) -> dict | None:
    try:
        with ENGINE.connect() as conn:
            rows = _rows_to_dicts(conn, "SELECT * FROM projects WHERE id=:id", {"id": pid})
        return rows[0] if rows else None
    except Exception:
        return None
    
def fetch_parent_with_children_totals(parent_name: str, exclude_id: int = 0) -> dict:
    try:
        with ENGINE.connect() as conn:
            parents = _rows_to_dicts(conn,
                "SELECT * FROM projects WHERE project_name=:n AND project_type='parent' LIMIT 1",
                {"n": parent_name})
            if not parents:
                return {}
            par = parents[0]
            if exclude_id:
                children = _rows_to_dicts(conn,
                    "SELECT * FROM projects WHERE parent_project=:n AND id != :ex",
                    {"n": parent_name, "ex": exclude_id})
            else:
                children = _rows_to_dicts(conn,
                    "SELECT * FROM projects WHERE parent_project=:n",
                    {"n": parent_name})
        def s(k): return sum(_n(c.get(k)) for c in children)

        par_ov   = _rev_ov(par)
        par_eac  = _n(par.get("rev_latest_eac") or par.get("baseline_eac"))
        par_bil  = _n(par.get("billed_amount"))
        par_col  = _n(par.get("collection_till_date"))
        par_itd  = _n(par.get("itd_cost"))
        par_base_eac = _n(par.get("baseline_eac"))

        used_ov   = sum(_rev_ov(c) for c in children)
        used_eac  = s("rev_latest_eac") or s("baseline_eac")
        used_bil  = s("billed_amount")
        used_col  = s("collection_till_date")
        used_itd  = s("itd_cost")
        used_base = s("baseline_eac")

        return {
            "parent_id":   par.get("id"),
            "parent_name": par.get("project_name"),

            "parent_co_no":               par.get("co_no") or "",
            "parent_business_line":       par.get("hier_business_line") or par.get("business_line") or "",
            "parent_project_manager":     par.get("project_manager") or "",
            "parent_client_name":         par.get("client_name") or "",
            "parent_order_number":        par.get("order_number") or "",
            "parent_order_issue_date":    par.get("order_issue_date") or "",
            "parent_status":              par.get("project_status") or "Active",
            "parent_physical_completion": _n(par.get("physical_completion")),

            "parent_baseline_eac":        par_base_eac,
            "parent_baseline_margin":     _n(par.get("baseline_margin")),
            "parent_rev_latest_margin":   _n(par.get("rev_latest_margin")),

            "parent_order_value": par_ov,
            "parent_eac":         par_eac,
            "parent_billed":      par_bil,
            "parent_collection":  par_col,
            "parent_itd_cost":    par_itd,
            "child_count":        len(children),

            "used_order_value":   used_ov,
            "used_eac":           used_eac,
            "used_billed":        used_bil,
            "used_collection":    used_col,
            "used_itd_cost":      used_itd,
            "used_baseline_eac":  used_base,

            "rem_order_value":    par_ov   - used_ov,
            "rem_eac":            par_eac  - used_eac,
            "rem_billed":         par_bil  - used_bil,
            "rem_collection":     par_col  - used_col,
            "rem_itd_cost":       par_itd  - used_itd,
            "rem_baseline_eac":   par_base_eac - used_base,
        }
    except Exception as e:
        print(f"[fetch_parent_with_children_totals] {e}")
        return 
 
def fetch_projects_for_select():
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
    "SELECT id,co_no,project_name,client_name,order_number,order_value,business_line,"
    "variation_amt,revised_order_value "
    "FROM projects ORDER BY id DESC")).fetchall()
        return rows
    except Exception:
        return []


# ── Documents ──────────────────────────────────────────────

def link_document(fields: dict):
    fields["created_at"] = datetime.now().isoformat()
    with ENGINE.begin() as conn:
        conn.execute(text("""
            INSERT INTO documents (project_id,doc_name,doc_category,doc_link,created_at)
            VALUES (:project_id,:doc_name,:doc_category,:doc_link,:created_at)
        """), fields)

def delete_document(doc_id: int):
    with ENGINE.begin() as conn:
        conn.execute(text("DELETE FROM documents WHERE id=:id"), {"id": doc_id})

def fetch_documents(project_id: int = None) -> list:
    try:
        with ENGINE.connect() as conn:
            if project_id:
                rows = conn.execute(text(
                    "SELECT id,doc_name,doc_category,doc_link FROM documents WHERE project_id=:p"
                ), {"p": project_id}).fetchall()
            else:
                rows = conn.execute(text(
                    "SELECT id,doc_name,doc_category,doc_link FROM documents ORDER BY id DESC"
                )).fetchall()
        return rows
    except Exception:
        return []


# ── Expenditures ───────────────────────────────────────────
def save_expenditure(fields: dict) -> int:
    pid_raw = fields.get("project_id", 0)
    pid_int = int(pid_raw) if str(pid_raw).strip().isdigit() else 0
    fields["project_id"] = pid_int

    if pid_int <= 0:
        print("[save_expenditure] Rejected: no project_id supplied")
        return 0

    # map of erp_<field> -> category name used in erp_updates / ref_expense_category
    legacy_map = {"erp_cogs":"COGS","erp_djcs":"DJCS","erp_equip":"Equipment",
                  "erp_manpower":"Manpower","erp_vehicles":"Vehicles","erp_others":"Others"}

    fixed_num = ("erp_cogs","erp_djcs","erp_equip","erp_manpower","erp_vehicles","erp_others",
                 "itd_cogs","itd_djcs","itd_equip","itd_manpower","itd_vehicles","itd_others")
    for f in fixed_num:
        fields[f] = _n(fields.get(f))

    fields.setdefault("created_at", datetime.now().isoformat())
    now_iso = datetime.now().isoformat()

    with ENGINE.begin() as conn:
        # 1) Fetch *current* cumulative totals per category from erp_updates history
        existing_totals = dict(conn.execute(text(
            "SELECT category, SUM(amount) FROM erp_updates WHERE project_id=:pid GROUP BY category"
        ), {"pid": pid_int}).fetchall())

        # 2) For every ERP field the user typed a non-zero "this update" amount into,
        #    insert a new history row (only if > 0 — don't log no-op updates)
        for erp_field, category in legacy_map.items():
            delta = _n(fields.get(erp_field))
            if delta:
                conn.execute(text("""
                    INSERT INTO erp_updates (project_id, category, amount, created_by, created_at)
                    VALUES (:pid, :cat, :amt, :by, :ts)
                """), {"pid": pid_int, "cat": category, "amt": delta,
                       "by": fields.get("created_by", "system"), "ts": now_iso})

        # 3) Recompute fresh cumulative totals (existing + this update)
        new_totals = {}
        for erp_field, category in legacy_map.items():
            prev = _n(existing_totals.get(category, 0))
            delta = _n(fields.get(erp_field))
            new_totals[erp_field] = prev + delta

        # 4) Store the *cumulative* values into expenditures.erp_* (snapshot row),
        #    not the raw delta the user typed
        snapshot_fields = {**fields, **new_totals}

        existing = conn.execute(text(
            "SELECT id FROM expenditures WHERE project_id=:pid ORDER BY id DESC LIMIT 1"
        ), {"pid": pid_int}).fetchone()

        if existing:
            exp_id = existing[0]
            conn.execute(text("""
                UPDATE expenditures SET
                    lob=:lob, data_as_of=:data_as_of, erp_update_date=:erp_update_date,
                    itd_cogs=:itd_cogs, itd_djcs=:itd_djcs, itd_equip=:itd_equip,
                    itd_manpower=:itd_manpower, itd_vehicles=:itd_vehicles, itd_others=:itd_others,
                    erp_cogs=:erp_cogs, erp_djcs=:erp_djcs, erp_equip=:erp_equip,
                    erp_manpower=:erp_manpower, erp_vehicles=:erp_vehicles, erp_others=:erp_others
                WHERE id=:id
            """), {**snapshot_fields, "id": exp_id})
        else:
            r = conn.execute(text("""
                INSERT INTO expenditures (
                    project_id, lob, data_as_of, erp_update_date,
                    itd_cogs, itd_djcs, itd_equip, itd_manpower, itd_vehicles, itd_others,
                    erp_cogs, erp_djcs, erp_equip, erp_manpower, erp_vehicles, erp_others,
                    created_at
                ) VALUES (
                    :project_id, :lob, :data_as_of, :erp_update_date,
                    :itd_cogs, :itd_djcs, :itd_equip, :itd_manpower, :itd_vehicles, :itd_others,
                    :erp_cogs, :erp_djcs, :erp_equip, :erp_manpower, :erp_vehicles, :erp_others,
                    :created_at
                ) RETURNING id"""), snapshot_fields)
            exp_id = r.fetchone()[0]

        total_erp = sum(new_totals.values())
        conn.execute(text(
            "UPDATE projects SET itd_cost=:v WHERE id=:id"
        ), {"v": total_erp, "id": pid_int})

    return exp_id


def fetch_expenditure_for_project(project_id: int) -> dict:
    try:
        with ENGINE.connect() as conn:
            rows = _rows_to_dicts(conn,
                "SELECT * FROM expenditures WHERE project_id=:p ORDER BY id DESC LIMIT 1",
                {"p": project_id})
        return rows[0] if rows else {}
    except Exception:
        return {}
def fetch_erp_totals(project_id: int) -> dict:
    """Returns {category: cumulative_total} for a project, summed from erp_updates.
    Safe to call even if the table doesn't exist yet or has no rows."""
    if not project_id:
        return {}
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                "SELECT category, SUM(amount) FROM erp_updates WHERE project_id=:pid GROUP BY category"
            ), {"pid": project_id}).fetchall()
        return {cat: (amt or 0) for cat, amt in rows}
    except Exception as e:
        print(f"[fetch_erp_totals] {e}")
        return {}

def save_expense_entry(fields: dict) -> int:
    for f in ("quantity","unit_cost"): fields[f] = _n(fields.get(f))
    fields["line_total"] = fields["quantity"] * fields["unit_cost"]
    fields["created_at"] = datetime.now().isoformat()

    cat = fields.get("category","").strip()
    if cat: add_ref_value("ref_expense_category","name",cat)

    pid = int(fields.get("project_id") or 0)

    exp_id = int(fields.get("expenditure_id") or 0)
    if exp_id <= 0 and pid > 0:
        existing = fetch_expenditure_for_project(pid)
        if existing:
            exp_id = existing["id"]
        else:
            stub = {
                "project_id": pid, "lob": "", "data_as_of": date.today().isoformat(),
                "erp_update_date": date.today().isoformat(),
                "itd_cogs":0,"itd_djcs":0,"itd_equip":0,"itd_manpower":0,"itd_vehicles":0,"itd_others":0,
                "erp_cogs":0,"erp_djcs":0,"erp_equip":0,"erp_manpower":0,"erp_vehicles":0,"erp_others":0,
            }
            exp_id = save_expenditure(stub)

    fields["expenditure_id"] = exp_id

    with ENGINE.begin() as conn:
        r = conn.execute(text("""
            INSERT INTO expense_entries
                (expenditure_id, project_id, category, description,
                 quantity, unit_cost, line_total, created_at)
            VALUES
                (:expenditure_id, :project_id, :category, :description,
                 :quantity, :unit_cost, :line_total, :created_at)
            RETURNING id"""), fields)
        return r.fetchone()[0]


def delete_expense_entry(entry_id: int):
    with ENGINE.begin() as conn:
        conn.execute(text("DELETE FROM expense_entries WHERE id=:id"), {"id": entry_id})


def fetch_expense_entries(project_id=None) -> list[dict]:
    try:
        with ENGINE.connect() as conn:
            if project_id:
                return _rows_to_dicts(conn,
                    "SELECT * FROM expense_entries WHERE project_id=:p ORDER BY id DESC",
                    {"p": project_id})
            else:
                return _rows_to_dicts(conn, "SELECT * FROM expense_entries ORDER BY id DESC")
    except Exception:
        return []


# ── Variations ─────────────────────────────────────────────

def save_variation(fields: dict) -> int:
    for f in ("original_value","variation_amt","revised_value"): fields[f] = _n(fields.get(f))
    fields["created_at"] = datetime.now().isoformat()
    with ENGINE.begin() as conn:
        r = conn.execute(text("""
            INSERT INTO variations
              (project_id, child_co_no, child_name, parent_name,
               original_value, variation_amt, revised_value, variation_date, status, created_at)
            VALUES
              (:project_id, :child_co_no, :child_name, :parent_name,
               :original_value, :variation_amt, :revised_value, :variation_date, :status, :created_at)
            RETURNING id"""), {**fields, "status": fields.get("status","Pending")})
        return r.fetchone()[0]


def delete_variation(var_id: int):
    with ENGINE.begin() as conn:
        conn.execute(text("DELETE FROM variations WHERE id=:id"), {"id": var_id})


def fetch_all_variations() -> list[dict]:
    try:
        with ENGINE.connect() as conn:
            return _rows_to_dicts(conn, "SELECT * FROM variations ORDER BY id DESC")
    except Exception:
        return []
def save_billing_entry(fields: dict) -> int:
    pid_raw = fields.get("project_id", 0)
    pid = int(pid_raw) if str(pid_raw).strip().isdigit() else 0
    if pid <= 0:
        return 0
    fields["project_id"]     = pid
    fields["billed_amt"]     = _n(fields.get("billed_amt"))
    fields["collection_amt"] = _n(fields.get("collection_amt"))
    fields.setdefault("ra_no", "")
    fields.setdefault("billing_date", "")
    fields.setdefault("collection_date", "")
    fields["created_at"] = datetime.now().isoformat()
    with ENGINE.begin() as conn:
        r = conn.execute(text("""
            INSERT INTO billing_collection
                (project_id, ra_no, billing_date, billed_amt, collection_date, collection_amt, created_at)
            VALUES
                (:project_id, :ra_no, :billing_date, :billed_amt, :collection_date, :collection_amt, :created_at)
            RETURNING id"""), fields)
        new_id = r.fetchone()[0]
    _sync_project_billing_totals(pid)
    return new_id
 
 
def update_billing_collection(entry_id: int, collection_date: str, collection_amt) -> bool:
    try:
        with ENGINE.begin() as conn:
            row = conn.execute(text(
                "SELECT project_id FROM billing_collection WHERE id=:id"
            ), {"id": entry_id}).fetchone()
            if not row:
                return False
            conn.execute(text("""
                UPDATE billing_collection
                SET collection_date=:cd, collection_amt=:ca
                WHERE id=:id
            """), {"cd": collection_date, "ca": _n(collection_amt), "id": entry_id})
        _sync_project_billing_totals(row[0])
        return True
    except Exception:
        return False
 
 
def delete_billing_entry(entry_id: int):
    row = None
    with ENGINE.begin() as conn:
        row = conn.execute(text(
            "SELECT project_id FROM billing_collection WHERE id=:id"
        ), {"id": entry_id}).fetchone()
        conn.execute(text("DELETE FROM billing_collection WHERE id=:id"), {"id": entry_id})
    if row:
        _sync_project_billing_totals(row[0])
 
 
def fetch_billing_entries(project_id: int) -> list:
    if not project_id:
        return []
    try:
        with ENGINE.connect() as conn:
            return _rows_to_dicts(conn,
                "SELECT * FROM billing_collection WHERE project_id=:p ORDER BY id DESC",
                {"p": project_id})
    except Exception:
        return []
def fetch_billing_entries_filtered(project_id=None, lob=None, pending_only=False) -> list:
    """Returns billing_collection rows joined with project name/CO.No,
    optionally narrowed by project_id and/or LOB, optionally pending-only."""
    try:
        with ENGINE.connect() as conn:
            sql = """
                SELECT bc.*, p.project_name, p.co_no, p.hier_business_line, p.business_line
                FROM billing_collection bc
                JOIN projects p ON p.id = bc.project_id
                WHERE 1=1
            """
            params = {}
            if project_id:
                sql += " AND bc.project_id = :pid"
                params["pid"] = project_id
            if lob:
                sql += " AND (p.hier_business_line = :lob OR p.business_line = :lob)"
                params["lob"] = lob
            if pending_only:
                sql += " AND (bc.collection_date IS NULL OR bc.collection_date = '')"
            sql += " ORDER BY bc.id DESC"
            return _rows_to_dicts(conn, sql, params)
    except Exception as e:
        print(f"[fetch_billing_entries_filtered] {e}")
        return [] 
 
def _sync_project_billing_totals(project_id: int):
    """Keeps projects.billed_amount / collection_till_date in sync with the
    billing_collection ledger, so Financials/KPI/Dashboard stay accurate
    automatically whenever a billing or collection entry changes."""
    try:
        with ENGINE.begin() as conn:
            row = conn.execute(text("""
                SELECT COALESCE(SUM(billed_amt),0), COALESCE(SUM(collection_amt),0)
                FROM billing_collection WHERE project_id=:p
            """), {"p": project_id}).fetchone()
            billed_total, collection_total = row[0], row[1]
            conn.execute(text("""
                UPDATE projects SET billed_amount=:b, collection_till_date=:c
                WHERE id=:id
            """), {"b": billed_total, "c": collection_total, "id": project_id})
    except Exception as e:
        print(f"[_sync_project_billing_totals] {e}")
 
 

def fetch_summary_stats() -> dict:
    projects   = fetch_all_projects()
    variations = fetch_all_variations()
    parents    = [p for p in projects if not (p.get("parent_project") or "").strip()]
    children   = [p for p in projects if (p.get("parent_project") or "").strip()]
    net_var    = sum(_n(v.get("variation_amt")) for v in variations)
    pending    = sum(1 for v in variations if (v.get("status") or "Pending") == "Pending")
    return {
        "total_parents":  len(parents)  or 0,
        "total_children": len(children) or 0,
        "net_variation":  net_var,
        "pending":        pending,
    }
import html as _htmlmod

def search_projects(query: str) -> list[dict]:
    if not query or not query.strip():
        return []
    like = f"%{query.strip()}%"
    try:
        with ENGINE.connect() as conn:
            return _rows_to_dicts(conn, """
                SELECT * FROM projects
                WHERE project_name      ILIKE :q
                   OR co_no             ILIKE :q
                   OR client_name       ILIKE :q
                   OR project_manager   ILIKE :q
                   OR order_number      ILIKE :q
                   OR business_line     ILIKE :q
                   OR hier_business_line ILIKE :q
                ORDER BY id DESC
            """, {"q": like})
    except Exception as e:
        print(f"[search_projects] {e}")
        return []


def search_results_page(query: str, username="Admin") -> str:
    safe_q  = _htmlmod.escape(query or "")
    results = search_projects(query)

    rows_html = ""
    for r in results:
        pid    = r.get("id")
        co     = _htmlmod.escape(r.get("co_no") or "—")
        name   = _htmlmod.escape(r.get("project_name") or "—")
        pm     = _htmlmod.escape(r.get("project_manager") or "—")
        client = _htmlmod.escape(r.get("client_name") or "—")
        status = _htmlmod.escape(r.get("project_status") or "—")
        lob    = _htmlmod.escape(r.get("hier_business_line") or r.get("business_line") or "—")
        rows_html += f"""
        <tr style="border-bottom:1px solid {C['outline']}">
          <td style="padding:12px 14px;font-weight:600;color:{C['dark']}">{co}</td>
          <td style="padding:12px 14px;font-weight:600">{name}</td>
          <td style="padding:12px 14px">{lob}</td>
          <td style="padding:12px 14px">{pm}</td>
          <td style="padding:12px 14px">{client}</td>
          <td style="padding:12px 14px">{status}</td>
          <td style="padding:12px 14px;text-align:right;white-space:nowrap">
            <a href="/edit_project?id={pid}" class="btn-outline" style="margin-right:6px">✏ Edit</a>
            <a href="/expenditures?project_id={pid}" class="btn-outline" style="margin-right:6px">Expenditures</a>
            <a href="/billing_collection?project_id={pid}" class="btn-outline">Billing</a>
          </td>
        </tr>"""

    if not rows_html:
        rows_html = (f'<tr><td colspan="7" style="padding:28px;text-align:center;color:#aaa">'
                     f'No projects matched &ldquo;{safe_q}&rdquo;.</td></tr>')

    html  = _html_head("Search Results") + "<body>"
    html += _sidebar("") + _topbar("SEARCH RESULTS", username, search_value=query)
    html += f"""
<div class="main"><div class="canvas">
  <h2 style="font-size:22px;font-weight:800;color:{C['dark']};margin-bottom:4px">Search Results</h2>
  <p style="font-size:13px;color:#6b7280;margin-bottom:20px">
    {len(results)} project(s) matched &ldquo;{safe_q}&rdquo;
  </p>
  <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;overflow:hidden">
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:{C['surface_low']};border-bottom:1px solid {C['outline']}">
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">CO.No</th>
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Project Name</th>
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">LOB</th>
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">PM</th>
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Client</th>
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Status</th>
            <th style="text-align:right;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Actions</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>
  </div>
</div></div>
</body></html>"""
    return html

# ═══════════════════════════════════════════════════════════════
#  SHARED HTML HELPERS
# ═══════════════════════════════════════════════════════════════

NAV_ITEMS = [
    ("dashboard",   "⊞",  "Dashboard"),
    ("data_entry",  "✎",  "Data Entry"),
    ("expenditures","💳", "Expenditures"),
    ("financials",  "📊", "Financials"),
    ("variations",  "📈", "Parent Project Details"),
    ("reports",     "📋", "Consolidated Reports"),
    ("kpi",         "🎯", "KPI Monitoring"),
    ("billing_collection", "🧾", "Billing & Collection"),
]

SHARED_CSS = f"""
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:{C['bg']};color:{C['on_surface']};display:flex;min-height:100vh}}
.sidebar{{width:200px;min-height:100vh;background:#fff;border-right:1px solid {C['outline']};display:flex;flex-direction:column;padding:20px 0 16px;position:fixed;top:0;left:0;z-index:200}}
.brand{{padding:0 16px 20px;font-size:13px;font-weight:700;line-height:1.4;color:{C['dark']};border-bottom:1px solid #eee;margin-bottom:8px}}
.nav-item{{display:flex;align-items:center;gap:9px;padding:9px 16px;text-decoration:none;font-size:13px;color:#4a5568;border-left:3px solid transparent;transition:background .15s;white-space:normal;line-height:1.3}}
.nav-item:hover{{background:{C['surface_high']}}}
.nav-item.active{{background:#f0f3fa;color:{C['dark']};font-weight:700;border-left-color:{C['dark']}}}
.nav-icon{{font-size:14px;width:18px;text-align:center;flex-shrink:0}}
.sidebar-btn{{margin:auto 14px 0;background:{C['dark']};color:#fff;border:none;border-radius:6px;padding:9px 0;font-size:12px;font-weight:700;cursor:pointer;width:calc(100% - 28px)}}
.topbar{{position:fixed;top:0;left:200px;right:0;height:52px;background:#fff;border-bottom:1px solid {C['outline']};display:flex;align-items:center;justify-content:space-between;padding:0 24px;z-index:199}}
.topbar-title{{font-size:12px;font-weight:700;color:#6b7280;letter-spacing:.04em;text-transform:uppercase}}
.topbar-right{{display:flex;align-items:center;gap:14px}}
.search-box{{background:{C['surface_high']};border:1px solid {C['outline']};border-radius:6px;padding:6px 12px;font-size:12px;color:#9ca3af;width:180px;outline:none}}
.topbar-nav a{{font-size:12px;font-weight:600;color:{C['on_var']};text-decoration:none;margin-left:18px}}
.topbar-nav a:hover{{color:{C['dark']}}}
.proj-sel-btn{{background:{C['dark']};color:#fff;border:none;border-radius:4px;padding:6px 12px;font-size:12px;font-weight:700;cursor:pointer}}
.avatar{{background:{C['dark']};color:#fff;border-radius:50%;width:28px;height:28px;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700}}
.avatar-link{{text-decoration:none;display:flex;align-items:center;gap:8px;color:{C['on_surface']}}}
.main{{margin-left:200px;padding-top:52px;flex:1;min-height:100vh;background:{C['bg']}}}
.canvas{{max-width:1280px;margin:0 auto;padding:24px}}
.btn-outline{{background:#fff;border:1px solid {C['outline']};border-radius:6px;padding:7px 14px;font-size:12px;font-weight:600;cursor:pointer;color:#374151;text-decoration:none}}
.btn-outline:hover{{background:{C['surface_high']}}}
.btn-primary{{background:{C['dark']};color:#fff;border:none;border-radius:6px;padding:7px 14px;font-size:12px;font-weight:600;cursor:pointer;text-decoration:none;display:inline-block}}
.btn-primary:hover{{background:{C['mid']}}}
.alert{{padding:11px 16px;border-radius:8px;margin-bottom:16px;font-size:13px;font-weight:600}}
.alert.ok{{background:#e8f5e9;color:#1b5e20;border:1px solid #a5d6a7}}
.alert.err{{background:#ffdad6;color:{C['error']};border:1px solid #ffb4ab}}
.del-btn{{background:none;border:none;color:{C['error']};cursor:pointer;font-size:17px;padding:2px 5px;border-radius:4px}}
.del-btn:hover{{background:#ffdad6}}
.nav-details summary::-webkit-details-marker {{ display:none }}
.nav-details summary {{ list-style:none }}
.nav-details summary::marker {{ content:"" }}
@media(max-width:800px){{.sidebar{{width:0;overflow:hidden}}.main{{margin-left:0}}.topbar{{left:0}}}}
</style>"""

def _html_head(title="Monarch Surveyors") -> str:
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="icon" href="{LOGO_URL}" type="image/jpeg">
<title>{title} | Monarch Surveyors</title>
{SHARED_CSS}
</head>"""

def _sidebar(active="dashboard", sub_section="") -> str:
    links = ""
    for key, icon, label in NAV_ITEMS:
        cls = "nav-item active" if key == active else "nav-item"
        links += f'<a href="/{key}" class="{cls}"><span class="nav-icon">{icon}</span>{label}</a>\n'

    return f"""
<aside class="sidebar">
  <div class="brand" style="display:flex;align-items:center;gap:10px;padding-bottom:14px">
    <img src="{LOGO_URL}" alt="Monarch logo" style="width:36px;height:36px;border-radius:6px;flex-shrink:0">
    <span style="line-height:1.3">Monarch<br>Surveyors &amp; Engineering<br>Consultants Ltd.</span>
  </div>
  {links}
  <button class="sidebar-btn" onclick="window.location='/reports'">⊕ New Report</button>
</aside>"""
def _topbar(subtitle="FINANCIAL MONITORING SYSTEM", username="Admin", search_value="") -> str:
    initials = "".join(w[0] for w in (username or "A").split()[:2]).upper() or "A"
    sv = (search_value or "").replace('"', "&quot;")
    return f"""
<header class="topbar">
  <span class="topbar-title">{subtitle}</span>
  <div class="topbar-right">
    <form method="get" action="/search" style="margin:0">
      <input class="search-box" name="q" placeholder="🔍  Search…" type="text" value="{sv}" autocomplete="off">
    </form>
    <nav class="topbar-nav">
      <a href="/dashboard">Dashboard</a>
      <a href="/reports">Reports</a>
      <a href="/kpi">KPI</a>
    </nav>
    <button class="proj-sel-btn" onclick="window.location='/data_entry'">Project Selector</button>
    <span style="font-size:18px;cursor:pointer">🔔</span>
    <a class="avatar-link" href="/logout" title="Sign out">
      <div class="avatar">{initials}</div>
      <span style="font-size:13px;font-weight:600">{username}</span>
    </a>
  </div>
</header>"""
def _sel(name: str, options: list[str], selected: str = "", placeholder: str = "Select…",
         cls: str = "", onchange: str = "", extra: str = "") -> str:
    opts = f'<option value="" disabled {"selected" if not selected else ""}>{placeholder}</option>'
    for o in options:
        opts += f'<option value="{o}" {"selected" if o == selected else ""}>{o}</option>'
    return (f'<select name="{name}" class="{cls}"'
            + (f' onchange="{onchange}"' if onchange else "")
            + (f' {extra}' if extra else "") + f'>{opts}</select>')

def _field_html(label: str, input_html: str) -> str:
    return f'<div class="field"><label>{label}</label>{input_html}</div>'

def _fmt_inr(n) -> str:
    try:    return f"{int(abs(float(n))):,}"
    except: return "0"


# ═══════════════════════════════════════════════════════════════
#  LOGIN PAGE
# ═══════════════════════════════════════════════════════════════

def login_page(alert_html: str = "", next_path: str = "/dashboard") -> str:
    year = datetime.now().year
    try:
        projects    = fetch_all_projects()
        lp_count    = len(projects)
        lp_ov       = sum(_rev_ov(p)                         for p in projects)
        lp_bil      = sum(_n(p.get("billed_amount"))         for p in projects)
        lp_col      = sum(_n(p.get("collection_till_date"))  for p in projects)
        lp_ub       = max(lp_bil - lp_col, 0)
        lp_pct      = f"{lp_col / lp_ov * 100:.1f}%" if lp_ov else "—"
        lp_ov_fmt   = f"₹ {lp_ov  / 1e7:.2f} Crore" if lp_ov  else "—"
        lp_col_fmt  = f"{lp_pct} (₹ {lp_col / 1e7:.1f} Cr)" if lp_col else "—"
        lp_ub_fmt   = f"₹ {lp_ub  / 1e7:.2f} Crore" if lp_ub  else "—"
        lp_active   = sum(1 for p in projects if (p.get("project_status") or "").lower() == "active")
        lp_count_fmt = f"{lp_active} Active" if lp_active else f"{lp_count} Total"
    except Exception:
        lp_count_fmt = "—"; lp_ov_fmt = "—"; lp_col_fmt = "—"; lp_ub_fmt = "—"

    return f"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="icon" href="{LOGO_URL}" type="image/jpeg">
<title>Sign In | Monarch Surveyors</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,Arial,sans-serif;background:{C['dark']};min-height:100vh;
     display:flex;align-items:center;justify-content:center;padding:20px}}
.wrap{{display:flex;width:100%;max-width:960px;min-height:580px;background:#fff;
       border-radius:18px;overflow:hidden;box-shadow:0 32px 80px rgba(0,0,0,.45)}}
.lp{{width:380px;flex-shrink:0;background:{C['dark']};padding:44px 36px;
     display:flex;flex-direction:column;justify-content:space-between}}
.lp-logo-wrap{{width:64px;height:64px;border-radius:14px;background:#fff;padding:8px;
               display:flex;align-items:center;justify-content:center;margin-bottom:20px;
               box-shadow:0 4px 14px rgba(0,0,0,.3)}}
.lp-logo-wrap img{{width:100%;height:100%;object-fit:contain;border-radius:6px}}
.lp-title{{font-size:15px;font-weight:700;color:#fff;line-height:1.4;margin-bottom:6px}}
.lp-eyebrow{{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;
             color:{C['secondary']};margin-bottom:24px}}
.lp-tag{{font-size:13px;color:rgba(255,255,255,.55);line-height:1.6;margin-bottom:32px}}
.kpi-row{{display:flex;align-items:center;gap:12px;padding:10px 14px;border-radius:8px;
           border:1px solid rgba(255,255,255,.07);background:rgba(255,255,255,.04);margin-bottom:8px}}
.kpi-icon{{width:34px;height:34px;border-radius:7px;display:flex;align-items:center;
            justify-content:center;font-size:15px;flex-shrink:0}}
.ki-gold{{background:rgba(119,90,25,.35);color:{C['secondary']}}}
.ki-blue{{background:rgba(44,62,107,.5);color:#adc7f8}}
.ki-rose{{background:rgba(186,26,26,.25);color:#f7c1c1}}
.kpi-lbl{{font-size:11px;color:rgba(255,255,255,.5);text-transform:uppercase;margin-bottom:2px}}
.kpi-val{{font-size:15px;font-weight:700;color:#fff}}
.lp-footer{{font-size:11px;color:rgba(255,255,255,.4);line-height:1.5}}
.rp{{flex:1;padding:48px 52px;display:flex;flex-direction:column;justify-content:center}}
.eyebrow{{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:{C['secondary']};margin-bottom:8px}}
.title{{font-size:30px;font-weight:800;color:{C['dark']};margin-bottom:6px}}
.sub{{font-size:14px;color:#6b7280;margin-bottom:28px}}
.alert-err{{background:#fff1f0;color:{C['error']};border:1px solid #ffb4ab;
             padding:12px 16px;border-radius:9px;font-size:13px;font-weight:600;margin-bottom:20px}}
.rf{{display:flex;flex-direction:column;gap:4px;margin-bottom:16px}}
.rf label{{font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']}}}
.rf-wrap{{position:relative}}
.rf-ico{{position:absolute;left:13px;top:50%;transform:translateY(-50%);font-size:16px;color:#9ca3af;pointer-events:none}}
.rf input{{width:100%;height:46px;border:1.5px solid {C['outline']};border-radius:9px;
            padding:0 14px 0 42px;font-size:14px;color:{C['on_surface']};background:#fff;outline:none}}
.rf input:focus{{border-color:{C['dark']};box-shadow:0 0 0 4px rgba(0,14,36,.07)}}
.submit-btn{{width:100%;height:48px;background:{C['dark']};color:#fff;border:none;border-radius:9px;
              font-size:13px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;cursor:pointer;margin-top:8px}}
.submit-btn:hover{{background:{C['mid']}}}
.hint{{margin-top:20px;padding:12px 16px;background:{C['surface_low']};border-radius:8px;
        border:1px solid {C['outline']};font-size:12px;color:#6b7280;display:flex;align-items:center;gap:8px}}
.hint code{{background:rgba(0,14,36,.08);color:{C['dark']};padding:1px 6px;border-radius:4px;font-size:11.5px;font-weight:700}}
.footer{{margin-top:20px;font-size:11px;color:#9ca3af;text-align:center}}
@media(max-width:780px){{.lp{{display:none}}.wrap{{max-width:480px}}.rp{{padding:40px 32px}}}}
</style></head><body>
<div class="wrap">

  <!-- LEFT — branding panel -->
  <section class="lp">
    <div>
      <div class="lp-logo-wrap">
        <img src="{LOGO_URL}" alt="Monarch logo">
      </div>
      <p class="lp-title">Monarch Surveyors &amp;<br>Engineering Consultants Ltd.</p>
      <p class="lp-eyebrow">Financial Monitoring System</p>
      <p class="lp-tag">Real-time visibility into project order values, expenditure, collections, and portfolio risk — all from a single live dashboard.</p>

      <div class="kpi-row">
        <div class="kpi-icon ki-blue">📁</div>
        <div>
          <div class="kpi-lbl">Project Portfolio</div>
          <div class="kpi-val">{lp_count_fmt}</div>
        </div>
      </div>
      <div class="kpi-row">
        <div class="kpi-icon ki-gold">💰</div>
        <div>
          <div class="kpi-lbl">Order Value</div>
          <div class="kpi-val">{lp_ov_fmt}</div>
        </div>
      </div>
      <div class="kpi-row">
        <div class="kpi-icon ki-rose">📊</div>
        <div>
          <div class="kpi-lbl">Collection</div>
          <div class="kpi-val">{lp_col_fmt}</div>
        </div>
      </div>
    </div>
    <div class="lp-footer">Secure login &middot; Session-based authentication</div>
  </section>

  <!-- RIGHT — sign-in form -->
  <section class="rp">
    <p class="eyebrow">Financial Monitoring System</p>
    <h1 class="title">Welcome back</h1>
    <p class="sub">Sign in to access your project dashboard and reports.</p>
    {alert_html}
    <form method="post" action="/login">
      <input type="hidden" name="next" value="{next_path}">
      <div class="rf">
        <label for="lf-user">Username</label>
        <div class="rf-wrap">
          <span class="rf-ico">👤</span>
          <input type="text" id="lf-user" name="username" placeholder="e.g. admin"
                 autocomplete="username" required>
        </div>
      </div>
      <div class="rf">
        <label for="lf-pass">Password</label>
        <div class="rf-wrap">
          <span class="rf-ico">🔒</span>
          <input type="password" id="lf-pass" name="password"
                 placeholder="••••••••" autocomplete="current-password" required>
        </div>
      </div>
      <button type="submit" class="submit-btn">Sign in to Dashboard →</button>
    </form>
    <p style="text-align:center;margin-top:16px">
      <a href="/reset_password" style="font-size:12px;font-weight:600;color:{C['secondary']};text-decoration:none">
        Forgot your password? Reset it here →
      </a>
    </p>
    <p class="footer">© {year} Monarch Surveyors &amp; Engineering Consultants Ltd.</p>
  </section>

</div>
<script>document.getElementById('lf-user').focus();</script>
</body></html>"""
def reset_password_page(alert_html: str = "", prefill_username: str = "") -> str:
    year = datetime.now().year
    return f"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="icon" href="{LOGO_URL}" type="image/jpeg">
<title>Reset Password | Monarch Surveyors</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,Arial,sans-serif;background:{C['dark']};min-height:100vh;
     display:flex;align-items:center;justify-content:center;padding:20px}}
.wrap{{width:100%;max-width:440px;background:#fff;border-radius:18px;overflow:hidden;
       box-shadow:0 32px 80px rgba(0,0,0,.45);padding:44px 40px}}
.rp-logo-wrap{{width:56px;height:56px;border-radius:12px;background:{C['dark']};padding:7px;
               display:flex;align-items:center;justify-content:center;margin-bottom:20px}}
.rp-logo-wrap img{{width:100%;height:100%;object-fit:contain;border-radius:6px}}
.eyebrow{{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:{C['secondary']};margin-bottom:8px}}
.title{{font-size:26px;font-weight:800;color:{C['dark']};margin-bottom:6px}}
.sub{{font-size:13px;color:#6b7280;margin-bottom:24px;line-height:1.5}}
.alert-err{{background:#fff1f0;color:{C['error']};border:1px solid #ffb4ab;
             padding:12px 16px;border-radius:9px;font-size:13px;font-weight:600;margin-bottom:20px}}
.alert-ok{{background:#e8f5e9;color:#1b5e20;border:1px solid #a5d6a7;
            padding:12px 16px;border-radius:9px;font-size:13px;font-weight:600;margin-bottom:20px}}
.rf{{display:flex;flex-direction:column;gap:4px;margin-bottom:16px}}
.rf label{{font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']}}}
.rf-wrap{{position:relative}}
.rf-ico{{position:absolute;left:13px;top:50%;transform:translateY(-50%);font-size:16px;color:#9ca3af;pointer-events:none}}
.rf input{{width:100%;height:46px;border:1.5px solid {C['outline']};border-radius:9px;
            padding:0 14px 0 42px;font-size:14px;color:{C['on_surface']};background:#fff;outline:none}}
.rf input:focus{{border-color:{C['dark']};box-shadow:0 0 0 4px rgba(0,14,36,.07)}}
.submit-btn{{width:100%;height:48px;background:{C['dark']};color:#fff;border:none;border-radius:9px;
              font-size:13px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;cursor:pointer;margin-top:8px}}
.submit-btn:hover{{background:{C['mid']}}}
.back-link{{display:block;text-align:center;margin-top:18px;font-size:12px;font-weight:600;color:{C['on_var']};text-decoration:none}}
.back-link:hover{{color:{C['dark']}}}
.footer{{margin-top:24px;font-size:11px;color:#9ca3af;text-align:center}}
</style></head><body>
<div class="wrap">
  <div class="rp-logo-wrap">
    <img src="{LOGO_URL}" alt="Monarch logo">
  </div>
  <p class="eyebrow">Account Recovery</p>
  <h1 class="title">Reset Password</h1>
  <p class="sub">Enter your username and choose a new password. Contact your administrator if you no longer have access to this account.</p>
  {alert_html}
  <form method="post" action="/reset_password">
    <div class="rf">
      <label for="rp-user">Username</label>
      <div class="rf-wrap">
        <span class="rf-ico">👤</span>
        <input type="text" id="rp-user" name="username" placeholder="e.g. admin"
               value="{prefill_username}" autocomplete="username" required>
      </div>
    </div>
    <div class="rf">
      <label for="rp-new">New Password</label>
      <div class="rf-wrap">
        <span class="rf-ico">🔒</span>
        <input type="password" id="rp-new" name="new_password" placeholder="At least 6 characters"
               autocomplete="new-password" required minlength="6">
      </div>
    </div>
    <div class="rf">
      <label for="rp-conf">Confirm New Password</label>
      <div class="rf-wrap">
        <span class="rf-ico">🔒</span>
        <input type="password" id="rp-conf" name="confirm_password" placeholder="Re-enter password"
               autocomplete="new-password" required minlength="6">
      </div>
    </div>
    <button type="submit" class="submit-btn">Update Password →</button>
  </form>
  <a href="/login" class="back-link">← Back to Sign In</a>
  <p class="footer">© {year} Monarch Surveyors &amp; Engineering Consultants Ltd.</p>
</div>
</body></html>"""

# ═══════════════════════════════════════════════════════════════
#  PAGE 1 — DASHBOARD
# ═══════════════════════════════════════════════════════════════
from datetime import datetime
# ── LOB palette (must stay in same order as get_lob_list()) ──────────────────
_LOB_COLORS = [
    "#1d4ed8",  # blue       — DPR
    "#f97316",  # orange     — DPR-Highways
    "#16a34a",  # green      — ELU preparation
    "#9333ea",  # purple     — Geospatial
    "#0891b2",  # cyan       — Geotechnical
    "#dc2626",  # red        — Land Acquisition
    "#ca8a04",  # yellow     — Locals
    "#be185d",  # pink       — Project Management
    "#0f766e",  # teal       — RAILWAY PROJECT
    "#65a30d",  # lime       — Railways
    "#ea580c",  # deep orange— Railway Tamping
]

from datetime import datetime

_LOB_COLORS = [
    "#1d4ed8",  # blue       — DPR
    "#f97316",  # orange     — DPR-Highways
    "#16a34a",  # green      — ELU preparation
    "#9333ea",  # purple     — Geospatial
    "#0891b2",  # cyan       — Geotechnical
    "#dc2626",  # red        — Land Acquisition
    "#ca8a04",  # yellow     — Locals
    "#be185d",  # pink       — Project Management
    "#0f766e",  # teal       — RAILWAY PROJECT
    "#65a30d",  # lime       — Railways
    "#ea580c",  # deep orange— Railway Tamping
]

def export_dashboard_pdf_bytes() -> bytes:
    projects = fetch_all_projects()
    lob_list = get_lob_list() or []
    def _lob(p): return p.get("hier_business_line") or p.get("business_line") or "Other"

    agg = {}
    for p in projects:
        lb = _lob(p)
        d = agg.setdefault(lb, {"count":0,"ov":0.0,"exp":0.0,"col":0.0})
        d["count"] += 1
        d["ov"]    += _rev_ov(p)
        d["exp"]   += _n(p.get("itd_cost"))
        d["col"]   += _n(p.get("collection_till_date"))

    ordered = lob_list or list(agg.keys())
    rows = [[lb, agg.get(lb,{}).get("count",0),
              f"{agg.get(lb,{}).get('ov',0):,.0f}",
              f"{agg.get(lb,{}).get('exp',0):,.0f}",
              f"{agg.get(lb,{}).get('col',0):,.0f}"] for lb in ordered]

    total_ov = sum(_rev_ov(p) for p in projects)
    return _build_table_pdf(
        "Executive Performance Dashboard",
        f"LOB performance summary — {len(projects)} projects",
        ["Line of Business","Projects","Order Value (₹)","Expenditure (₹)","Collection (₹)"],
        rows, footer_label="Total Order Value", footer_value=total_ov)

def dashboard(username: str = "Admin") -> str:
    projects   = fetch_all_projects()
    total_proj = len(projects)
    lob_list   = get_lob_list() or []

    # ── Aggregate totals ──────────────────────────────────────────────────────
    total_ov     = sum(_rev_ov(p)                         for p in projects)
    total_billed = sum(_n(p.get("billed_amount"))        for p in projects)
    total_col    = sum(_n(p.get("collection_till_date")) for p in projects)
    total_exp    = sum(_n(p.get("itd_cost"))             for p in projects)
    total_rev_ov = sum(_n(p.get("revised_order_value") or (_n(p.get("order_value")) + _n(p.get("variation_amt")))) for p in projects)
    unbilled     = max(total_rev_ov - total_col, 0)   # Col AF: Revised OV - Collection
    avg_comp     = (
        sum(_n(p.get("physical_completion")) for p in projects) / total_proj
        if total_proj else 0
    )
    col_pct = f"{total_col / total_ov * 100:.1f}%" if total_ov else "—"
    comp_w  = f"{avg_comp:.1f}%"
    
    # ── Per-LOB aggregates ────────────────────────────────────────────────────
    def _lob(p):
        return p.get("hier_business_line") or p.get("business_line") or "Other"

    lob_data = {}
    for p in projects:
        lb = _lob(p)
        if lb not in lob_data:
            lob_data[lb] = {"count": 0, "ov": 0.0, "exp": 0.0,
                            "col": 0.0, "billed": 0.0}
        lob_data[lb]["count"]  += 1
        lob_data[lb]["ov"]     += _rev_ov(p)
        lob_data[lb]["exp"]    += _n(p.get("itd_cost"))
        lob_data[lb]["col"]    += _n(p.get("collection_till_date"))
        lob_data[lb]["billed"] += _n(p.get("billed_amount"))

    ordered_lobs = lob_list if lob_list else list(lob_data.keys())

    # ── Formatting helpers ────────────────────────────────────────────────────
    def _fmt(v: float) -> str:
        if v >= 1e7:
            return f"&#8377; {v / 1e7:.2f} Cr"
        if v >= 1e5:
            return f"&#8377; {v / 1e5:.1f} L"
        return f"&#8377; {v:,.0f}"

    def _fmt_short(v: float) -> str:
        if v >= 1e7:
            return f"&#8377; {v / 1e7:.1f} Cr"
        return f"&#8377; {v:,.0f}"

    # ── Chart.js JS data arrays ───────────────────────────────────────────────
    def _js_arr(vals):
        return "[" + ",".join(str(round(v, 2)) for v in vals) + "]"

    lob_colors_js = "[" + ",".join(f'"{_LOB_COLORS[i % len(_LOB_COLORS)]}"'
                                   for i in range(len(ordered_lobs))) + "]"
    lob_labels_js = "[" + ",".join(f'"{lb}"' for lb in ordered_lobs) + "]"

    counts_js = _js_arr([lob_data.get(lb, {}).get("count", 0) for lb in ordered_lobs])
    ov_js     = _js_arr([lob_data.get(lb, {}).get("ov",    0) for lb in ordered_lobs])
    exp_js    = _js_arr([lob_data.get(lb, {}).get("exp",   0) for lb in ordered_lobs])
    col_js    = _js_arr([lob_data.get(lb, {}).get("col",   0) for lb in ordered_lobs])
    wip_js    = _js_arr([
        max(lob_data.get(lb, {}).get("billed", 0) - lob_data.get(lb, {}).get("col", 0), 0)
        for lb in ordered_lobs
    ])

    # ── KPI cards ─────────────────────────────────────────────────────────────
    def _kcard(label, val, sub_html):
        return f"""
        <div style="background:#fff;border-radius:10px;padding:16px 18px;
                    border:0.5px solid {C['outline']}">
          <p style="font-size:10px;font-weight:700;letter-spacing:.07em;
                    color:#9ca3af;text-transform:uppercase;margin-bottom:8px">{label}</p>
          <p style="font-size:24px;font-weight:800;color:{C['dark']};line-height:1">{val}</p>
          {sub_html}
        </div>"""

    cards_html = (
        _kcard("Total Project Inventory", str(total_proj),
               '<p style="font-size:11px;color:#16a34a;font-weight:600;margin-top:6px">'
               '&#8593; +8.4% vs Prev. Month</p>') +
        _kcard("Aggregate Collection", _fmt_short(total_col),
               f'<p style="font-size:11px;color:#6b7280;margin-top:6px">'
               f'{col_pct} Realization Rate</p>') +
        _kcard("Unbilled WIP", _fmt_short(unbilled),
               ('<p style="font-size:11px;color:#dc2626;font-weight:600;margin-top:6px">'
                '&#9888; Attention Required</p>'
                if unbilled > 0 else
                '<p style="font-size:11px;color:#16a34a;font-weight:600;margin-top:6px">'
                '&#10003; Clear</p>')) +
        _kcard("Avg. Completion", comp_w,
               f'<div style="background:#e5e7eb;border-radius:99px;height:7px;'
               f'margin-top:10px;overflow:hidden">'
               f'<div style="height:100%;width:{comp_w};background:{C["dark"]};'
               f'border-radius:99px"></div></div>')
    )

    # ── Global totals bar ─────────────────────────────────────────────────────
    glob_stats = [
        ("Projects",     f"{total_proj} Units", "#fff"),
        ("Order Value",  _fmt_short(total_ov),  "#fff"),
        ("Expenditure",  _fmt_short(total_exp), "#fff"),
        ("Collection",   _fmt_short(total_col), "#fff"),
        ("Unbilled WIP", _fmt_short(unbilled),  "#f87171"),
    ]
    glob_html = "".join(
        f'<div style="text-align:center;padding:0 16px;'
        f'{"border-right:none" if i == len(glob_stats) - 1 else "border-right:1px solid #1e3a6a"}">'
        f'<div style="font-size:9px;color:#64748b;text-transform:uppercase;'
        f'letter-spacing:.06em;margin-bottom:2px">{lbl}</div>'
        f'<div style="font-size:12px;font-weight:700;color:{clr}">{val}</div>'
        f'</div>'
        for i, (lbl, val, clr) in enumerate(glob_stats)
    )

    # ── LOB table rows ────────────────────────────────────────────────────────
    lob_rows = ""
    for lb in ordered_lobs:
        d   = lob_data.get(lb, {})
        ov  = d.get("ov",    0)
        col = d.get("col",   0)
        exp = d.get("exp",   0)
        cnt = d.get("count", 0)
        pct = f"{col / ov * 100:.1f}%" if ov else "—"
        lob_rows += f"""
        <tr style="border-bottom:0.5px solid #f1f3f7">
          <td style="padding:9px 14px;font-weight:600;color:{C['dark']}">{lb}</td>
          <td style="padding:9px 14px;text-align:right;font-family:'Courier New',monospace">{cnt}</td>
          <td style="padding:9px 14px;text-align:right;font-family:'Courier New',monospace">{_fmt(ov)}</td>
          <td style="padding:9px 14px;text-align:right;font-family:'Courier New',monospace">{_fmt(exp)}</td>
          <td style="padding:9px 14px;text-align:right;font-family:'Courier New',monospace">{_fmt(col)}</td>
          <td style="padding:9px 14px;text-align:right;font-weight:700;color:{C['dark']}">{pct}</td>
        </tr>"""

    if not lob_rows:
        lob_rows = (f'<tr><td colspan="6" style="padding:20px;text-align:center;color:#aaa">'
                    f'No projects yet. <a href="/data_entry">Add via Data Entry &#8594;</a></td></tr>')

    # ── LOB legend ────────────────────────────────────────────────────────────
    legend_html = "".join(
        f'<span style="display:flex;align-items:center;gap:5px;font-size:10px;color:#6b7280">'
        f'<span style="width:10px;height:10px;border-radius:2px;'
        f'background:{_LOB_COLORS[i % len(_LOB_COLORS)]}"></span>{lb}</span>'
        for i, lb in enumerate(ordered_lobs)
    )

    # ── Chart.js + custom fixed tooltip (never clipped by any parent) ─────────
    chart_script = f"""
<style>
#ch-tooltip {{
  position: fixed;
  background: rgba(15,27,61,0.93);
  color: #fff;
  padding: 8px 12px;
  border-radius: 7px;
  font-size: 12px;
  font-family: sans-serif;
  pointer-events: none;
  z-index: 99999;
  display: none;
  max-width: 240px;
  line-height: 1.5;
  box-shadow: 0 4px 16px rgba(0,0,0,0.35);
}}
#ch-tooltip .tt-title {{
  font-weight: 700;
  font-size: 13px;
  margin-bottom: 5px;
  padding-bottom: 5px;
  border-bottom: 1px solid rgba(255,255,255,0.18);
}}
#ch-tooltip .tt-row {{
  display: flex;
  align-items: center;
  gap: 7px;
}}
#ch-tooltip .tt-swatch {{
  width: 10px;
  height: 10px;
  border-radius: 2px;
  flex-shrink: 0;
}}
#ch-tooltip .tt-val {{
  font-weight: 600;
  font-size: 13px;
}}
</style>
<div id="ch-tooltip">
  <div class="tt-title" id="tt-title"></div>
  <div class="tt-row">
    <div class="tt-swatch" id="tt-swatch"></div>
    <div class="tt-val" id="tt-body"></div>
  </div>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script>
(function () {{
  var ttBox    = document.getElementById('ch-tooltip');
  var ttTitle  = document.getElementById('tt-title');
  var ttBody   = document.getElementById('tt-body');
  var ttSwatch = document.getElementById('tt-swatch');

  function fmtVal(v, isCount) {{
    if (isCount) return v + ' units';
    if (v >= 10000000) return '\u20b9 ' + (v / 10000000).toFixed(2) + ' Cr';
    if (v >= 100000)   return '\u20b9 ' + (v / 100000).toFixed(1) + ' L';
    return '\u20b9 ' + v.toFixed(0).replace(/\B(?=(\d{{3}})+(?!\d))/g, ',');
  }}

  function customTooltip(context, isCount) {{
    var model = context.tooltip;
    if (model.opacity === 0) {{
      ttBox.style.display = 'none';
      return;
    }}
    if (model.dataPoints && model.dataPoints.length) {{
      var dp = model.dataPoints[0];
      ttTitle.textContent          = dp.label;
      ttBody.textContent           = fmtVal(dp.raw, isCount);
      ttSwatch.style.background    = model.labelColors[0].backgroundColor;
    }}
    var rect = context.chart.canvas.getBoundingClientRect();
    var x    = rect.left + window.scrollX + model.caretX;
    var y    = rect.top  + window.scrollY + model.caretY;
    var ttW  = ttBox.offsetWidth  || 180;
    var ttH  = ttBox.offsetHeight || 60;
    var winW = window.innerWidth;
    var winH = window.innerHeight;
    var left = x + 14;
    var top  = y - ttH - 10;
    if (left + ttW > winW - 10) {{ left = x - ttW - 14; }}
    if (top < 0)                {{ top  = y + 14; }}
    ttBox.style.left    = left + 'px';
    ttBox.style.top     = top  + 'px';
    ttBox.style.display = 'block';
  }}

  var colors = {lob_colors_js};
  var labels = {lob_labels_js};
  var sets = [
    {{ id: 'ch0', data: {counts_js}, isCount: true  }},
    {{ id: 'ch1', data: {ov_js},     isCount: false }},
    {{ id: 'ch2', data: {exp_js},    isCount: false }},
    {{ id: 'ch3', data: {col_js},    isCount: false }},
    {{ id: 'ch4', data: {wip_js},    isCount: false }}
  ];

  sets.forEach(function (s) {{
    var el = document.getElementById(s.id);
    if (!el) return;
    var isCount = s.isCount;
    new Chart(el, {{
      type: 'doughnut',
      data: {{
        labels: labels,
        datasets: [{{
          data: s.data,
          backgroundColor: colors,
          borderWidth: 0,
          hoverOffset: 6
        }}]
      }},
      options: {{
        responsive: false,
        cutout: '62%',
        plugins: {{
          legend: {{ display: false }},
          tooltip: {{
            enabled: false,
            external: function (ctx) {{ customTooltip(ctx, isCount); }}
          }}
        }}
      }}
    }});
  }});

  document.addEventListener('mouseleave', function () {{
    ttBox.style.display = 'none';
  }});
}})();
</script>"""

    # ── Footer ────────────────────────────────────────────────────────────────
    now     = datetime.now().strftime("%d %b %Y, %I:%M %p IST")
    pending = fetch_summary_stats().get("pending", 0)

    # ── Assemble full page ────────────────────────────────────────────────────
    html  = _html_head("Executive Dashboard") + "<body>"
    html += _sidebar("dashboard")
    html += _topbar("FINANCIAL MONITORING SYSTEM", username)
    html += f"""
<div class="main"><div class="canvas">

  <!-- Page header -->
  <p style="font-size:11px;color:#9ca3af;margin-bottom:6px">
    REPORTS / <span style="color:#374151;font-weight:500">LOB PERFORMANCE SUMMARY</span>
  </p>
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px">
    <div>
      <h1 style="font-size:22px;font-weight:800;color:{C['dark']}">
        Executive Performance Dashboard
      </h1>
      <p style="font-size:12px;color:#6b7280;margin-top:2px">
        Live data from database &mdash; {total_proj} projects loaded
      </p>
    </div>
    <div style="display:flex;gap:8px">
      <a href="/export_pdf/dashboard" style="display:flex;align-items:center;gap:5px;
         padding:7px 13px;background:#fff;border:0.5px solid #d1d5db;
         border-radius:8px;font-size:12px;color:#374151;text-decoration:none">
        &#8595; Export PDF
      </a>
    </div>
  </div>

  <!-- KPI cards -->
  <div style="display:grid;grid-template-columns:repeat(4,1fr);
              gap:14px;margin-bottom:18px">
    {cards_html}
  </div>

  <!-- Global totals bar -->
  <div style="background:{C['dark']};border-radius:10px 10px 0 0;
              padding:13px 20px;display:flex;align-items:center;flex-wrap:wrap;
              gap:10px;overflow-x:auto">
    <span style="font-size:10px;font-weight:700;letter-spacing:.07em;
                 color:#64748b;text-transform:uppercase;margin-right:auto;
                 white-space:nowrap">
      Global Portfolio Totals
    </span>
    {glob_html}
  </div>
  <!-- LOB breakdown table -->
  <div style="background:#fff;border-radius:0 0 10px 10px;
              border:0.5px solid {C['outline']};border-top:none;margin-bottom:18px">
    <div style="padding:12px 16px;border-bottom:0.5px solid {C['outline']};
                display:flex;justify-content:space-between;align-items:center;
                background:{C['surface_low']}">
      <span style="font-size:13px;font-weight:700;color:{C['dark']}">
        LOB Performance Breakdown
      </span>
      <a href="/data_entry" style="background:{C['dark']};color:#fff;
         border-radius:8px;padding:5px 10px;font-size:11px;text-decoration:none">
        + Add Project
      </a>
    </div>
    <div style="overflow-x:auto">
      <table style="width:100%;min-width:800px;border-collapse:collapse;font-size:12px">
        <thead>
          <tr style="background:{C['surface_low']};border-bottom:0.5px solid {C['outline']}">
            <th style="padding:9px 14px;text-align:left;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Line of Business</th>
            <th style="padding:9px 14px;text-align:right;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Projects</th>
            <th style="padding:9px 14px;text-align:right;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Order Value (&#8377;)</th>
            <th style="padding:9px 14px;text-align:right;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Expenditure (&#8377;)</th>
            <th style="padding:9px 14px;text-align:right;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Collection (&#8377;)</th>
            <th style="padding:9px 14px;text-align:right;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Realization %</th>
          </tr>
        </thead>
        <tbody>{lob_rows}</tbody>
      </table>
    </div>
  </div>

  <!-- 5 doughnut charts -->
  <div style="background:#fff;border-radius:10px;border:0.5px solid {C['outline']};
              margin-bottom:18px;overflow-x:auto">
    <div style="display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));
                gap:0;padding:20px 8px;min-width:760px">
      <div style="display:flex;flex-direction:column;align-items:center;
                  border-right:0.5px solid {C['outline']};padding:0 12px">
        <p style="font-size:9px;font-weight:700;letter-spacing:.06em;
                  text-transform:uppercase;color:#9ca3af;margin-bottom:8px">Ongoing Projects</p>
        <canvas id="ch0" width="130" height="130"
                role="img" aria-label="Doughnut chart of ongoing projects by LOB"></canvas>
        <p style="font-size:12px;font-weight:800;color:{C['dark']};margin-top:8px">N = {total_proj}</p>
        <p style="font-size:9px;color:#9ca3af;margin-top:1px">Total unit count</p>
      </div>

      <div style="display:flex;flex-direction:column;align-items:center;
                  border-right:0.5px solid {C['outline']};padding:0 12px">
        <p style="font-size:9px;font-weight:700;letter-spacing:.06em;
                  text-transform:uppercase;color:#9ca3af;margin-bottom:8px">Order Value</p>
        <canvas id="ch1" width="130" height="130"
                role="img" aria-label="Doughnut chart of order value by LOB"></canvas>
        <p style="font-size:12px;font-weight:800;color:{C['dark']};margin-top:8px">{_fmt_short(total_ov)}</p>
        <p style="font-size:9px;color:#9ca3af;margin-top:1px">Projected revenue</p>
      </div>

      <div style="display:flex;flex-direction:column;align-items:center;
                  border-right:0.5px solid {C['outline']};padding:0 12px">
        <p style="font-size:9px;font-weight:700;letter-spacing:.06em;
                  text-transform:uppercase;color:#9ca3af;margin-bottom:8px">Expenditure</p>
        <canvas id="ch2" width="130" height="130"
                role="img" aria-label="Doughnut chart of expenditure by LOB"></canvas>
        <p style="font-size:12px;font-weight:800;color:{C['dark']};margin-top:8px">{_fmt_short(total_exp)}</p>
        <p style="font-size:9px;color:#9ca3af;margin-top:1px">Cost incurred</p>
      </div>

      <div style="display:flex;flex-direction:column;align-items:center;
                  border-right:0.5px solid {C['outline']};padding:0 12px">
        <p style="font-size:9px;font-weight:700;letter-spacing:.06em;
                  text-transform:uppercase;color:#9ca3af;margin-bottom:8px">Collection</p>
        <canvas id="ch3" width="130" height="130"
                role="img" aria-label="Doughnut chart of collection by LOB"></canvas>
        <p style="font-size:12px;font-weight:800;color:{C['dark']};margin-top:8px">{_fmt_short(total_col)}</p>
        <p style="font-size:9px;color:#9ca3af;margin-top:1px">Realized funds</p>
      </div>

      

    </div>
    <!-- LOB legend -->
    <div style="padding:10px 16px 14px;border-top:0.5px solid {C['outline']};
                display:flex;flex-wrap:wrap;gap:14px;justify-content:center">
      {legend_html}
    </div>
  </div>

  <!-- Footer -->
  <div style="display:flex;justify-content:space-between;align-items:center;
              font-size:10px;color:#9ca3af;padding:0 2px;margin-bottom:16px">
    <span>
      <span style="color:#22c55e;margin-right:3px">&#9679;</span>
      Data Last Synced: {now}
    </span>
    <span>&#10003; Audit Status: Verified (V.2.4) &ndash; Monarch Precision Standard</span>
    <span style="color:#dc2626;font-weight:600">
      &#9888; {pending} Variations Pending Final Approval
    </span>
  </div>

  <p style="font-size:10px;color:#d1d5db;text-align:center;margin-bottom:8px">
    &copy; 2024 Monarch Surveyors and Engineering Consultants Ltd. | Confidential
  </p>

</div></div>
{chart_script}
</body></html>"""

    return html
# ═══════════════════════════════════════════════════════════════
#  PAGE 2 — DATA ENTRY
# ═══════════════════════════════════════════════════════════════

def project_data_entry(alert_html="", prefill=None, username="Admin") -> str:
    p = prefill or {}

    ENTRY_CSS = f"""
<style>
.page-title{{font-size:26px;font-weight:800;color:{C['dark']};margin-bottom:4px}}
.page-sub{{font-size:13px;color:#6b7280;margin-bottom:20px}}
.upload-zone{{border:2px dashed {C['outline']};border-radius:12px;padding:28px;background:#fff;text-align:center;cursor:pointer;margin-bottom:24px;transition:border-color .2s}}
.upload-zone:hover{{border-color:{C['dark']};background:#f0f4ff}}
.card{{background:#fff;border:1px solid {C['outline']};border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:20px}}
.tab-bar{{display:flex;border-bottom:1px solid {C['outline']};background:{C['surface_low']};overflow-x:auto}}
.tab-btn{{flex:1;min-width:160px;padding:13px 16px;border:none;background:transparent;font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']};cursor:pointer;border-bottom:2px solid transparent;transition:color .15s}}
.tab-btn:hover,.tab-active{{color:{C['dark']}!important;border-bottom-color:{C['dark']}!important}}
.section{{padding:28px;display:none}}
.sec-title{{font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['dark']};margin-bottom:14px}}
.g2{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
.g3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}}
.g12{{display:grid;grid-template-columns:1fr 2fr;gap:16px}}
.g21{{display:grid;grid-template-columns:2fr 1fr;gap:16px}}
.mt{{margin-top:16px}}
.field{{display:flex;flex-direction:column;gap:4px}}
.field label{{font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}}}
.field input,.field select,.field textarea{{height:40px;border:1px solid {C['outline']};border-radius:6px;padding:0 12px;font-size:14px;color:{C['on_surface']};background:#fff;outline:none;transition:border-color .15s;width:100%}}
.field input:focus,.field select:focus{{border-color:{C['dark']};box-shadow:0 0 0 3px rgba(0,14,36,.07)}}
.cur-wrap{{position:relative}}
.cur-wrap .pre{{position:absolute;left:10px;top:50%;transform:translateY(-50%);font-weight:700;color:{C['on_var']};font-size:13px;pointer-events:none}}
.cur-wrap input{{padding-left:24px}}
.suf-wrap{{position:relative}}
.suf-wrap .suf{{position:absolute;right:10px;top:50%;transform:translateY(-50%);color:{C['on_var']};font-size:13px;pointer-events:none}}
.suf-wrap input{{padding-right:28px}}
.fin-card{{background:{C['surface_low']};border:1px solid {C['outline']};border-radius:8px;padding:16px}}
.info-ban{{display:flex;align-items:flex-start;gap:10px;background:#fffbf0;border:1px solid {C['sec_cont']};border-radius:8px;padding:12px 16px;font-size:13px;color:#5d4201;margin-top:14px}}
.proj-table{{width:100%;border-collapse:collapse;font-size:13px}}
.proj-table th{{padding:9px 14px;font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']};text-align:left;background:{C['surface_low']};border-bottom:1px solid {C['outline']}}}
.proj-table td{{padding:11px 14px;border-bottom:1px solid #eee}}
.proj-table tr:hover td{{background:{C['surface_low']}}}
.btn-prev{{background:transparent;border:2px solid {C['outline']};color:{C['dark']};padding:9px 24px;border-radius:8px;font-size:11px;font-weight:700;text-transform:uppercase;cursor:pointer;display:inline-flex;align-items:center;gap:6px}}
.btn-next,.btn-sub{{background:{C['dark']};color:#fff;border:none;padding:9px 24px;border-radius:8px;font-size:11px;font-weight:700;text-transform:uppercase;cursor:pointer;box-shadow:0 2px 6px rgba(0,0,0,.18);display:inline-flex;align-items:center;gap:6px}}
.exp-bar{{display:flex;gap:10px;margin-bottom:16px}}
.exp-btn{{background:{C['surface_low']};border:1px solid {C['outline']};color:{C['dark']};padding:7px 16px;border-radius:6px;font-size:12px;font-weight:700;text-decoration:none;letter-spacing:.04em;text-transform:uppercase}}
</style>"""

    TAB_JS = """<script>
var currentTab=1,totalTabs=4;
function switchTab(n){currentTab=n;updateUI();}
function nextTab(){if(currentTab<totalTabs){currentTab++;updateUI();}}
function prevTab(){if(currentTab>1){currentTab--;updateUI();}}
function updateUI(){
  for(var i=1;i<=totalTabs;i++){
    var s=document.getElementById('section-'+i);
    var t=document.getElementById('tab-'+i);
    if(s)s.style.display=(i===currentTab?'block':'none');
    if(t)t.classList.toggle('tab-active',i===currentTab);
  }
  var dt=document.getElementById('docs-table-sec');
  if(dt)dt.style.display=(currentTab===4?'block':'none');
  var pb=document.getElementById('prev-btn');
  var nb=document.getElementById('next-btn');
  var sb=document.getElementById('sub-btn');
  if(pb)pb.style.display=(currentTab===1?'none':'inline-flex');
  if(nb)nb.style.display=(currentTab===totalTabs?'none':'inline-flex');
  if(sb)sb.style.display=(currentTab===totalTabs?'inline-flex':'none');
}
function setFieldsReadOnly(readonly){
  var TRACKED=['order_value','variation_amt','revised_order_value','baseline_eac','baseline_margin','rev_latest_eac','rev_latest_margin','physical_completion','billed_amount','collection_till_date','itd_cost','bill_submission_pending_date','last_billed_date','project_status','business_line','project_manager','client_name','order_number','order_issue_date','co_no','project_name'];
  TRACKED.forEach(function(nm){
    var el=document.querySelector('[name="'+nm+'"]');
    if(!el)return;
    if(readonly){el.setAttribute('readonly','readonly');el.setAttribute('tabindex','-1');el.style.background='#f2f4f6';el.style.color='#6b7280';el.style.cursor='not-allowed';el.style.pointerEvents='none';el.style.borderColor='#e5e7eb';if(el.tagName==='SELECT')el.style.opacity='0.7';}
    else{el.removeAttribute('readonly');el.removeAttribute('tabindex');el.style.background='';el.style.color='';el.style.cursor='';el.style.pointerEvents='';el.style.borderColor='';el.style.opacity='';}
  });
  var banner=document.getElementById('readonly-banner');
  if(banner)banner.style.display=readonly?'flex':'none';
  var picker=document.getElementById('parent-picker-wrap');
  if(picker)picker.style.display=readonly?'block':'none';
  var subBtn=document.getElementById('sub-btn');
  if(subBtn)subBtn.style.display=readonly?'none':'inline-flex';
}


function setType(type){
  var rbParent=document.getElementById('rb-parent');
  var rbChild=document.getElementById('rb-child');
  var lblP=document.getElementById('lbl-parent');
  var lblC=document.getElementById('lbl-child');
  var wrap=document.getElementById('parent-field-wrap');
  if(rbParent)rbParent.checked=(type==='parent');
  if(rbChild)rbChild.checked=(type==='child');
  var dark='#1a2540';var white='#ffffff';var plain='#000e24';
  if(lblP){lblP.style.background=(type==='parent')?dark:white;lblP.style.color=(type==='parent')?white:plain;}
  if(lblC){lblC.style.background=(type==='child')?dark:white;lblC.style.color=(type==='child')?white:plain;}
  if(wrap)wrap.style.display=(type==='child')?'block':'none';
  // Fields stay editable regardless of type — "Parent" vs "Child" only
  // controls whether the parent-link field is shown, never a readonly lock.
  setFieldsReadOnly(false);
}
function onParentPickerChange(sel){
  var parentName=sel.value;
  if(!parentName)return;
  fetchAndFillParent(parentName);
}
var _parentFillCache={};
function fetchAndFillParent(parentName){
  if(!parentName)return;
  if(_parentFillCache[parentName]){fillFields(_parentFillCache[parentName]);return;}
  fetch('/get_parent_data?parent_name='+encodeURIComponent(parentName)+'&exclude_id=0')
    .then(function(r){return r.json();})
    .then(function(d){if(d&&d.parent_id){_parentFillCache[parentName]=d;fillFields(d);}})
    .catch(function(){});
}
function fillFields(d){
  function setVal(nm,val){var el=document.querySelector('[name="'+nm+'"]');if(!el||val===undefined||val===null)return;el.value=val;}
  setVal('co_no',d.parent_co_no);setVal('project_name',d.parent_name);
  setVal('business_line',d.parent_business_line);setVal('project_manager',d.parent_project_manager);
  setVal('client_name',d.parent_client_name);setVal('order_number',d.parent_order_number);
  setVal('order_issue_date',d.parent_order_issue_date);setVal('project_status',d.parent_status);
  setVal('physical_completion',d.parent_physical_completion);setVal('baseline_eac',d.parent_baseline_eac);
  setVal('baseline_margin',d.parent_baseline_margin);setVal('rev_latest_margin',d.parent_rev_latest_margin);
  setVal('order_value',d.parent_order_value);setVal('rev_latest_eac',d.parent_eac);
  setVal('billed_amount',d.parent_billed);setVal('collection_till_date',d.parent_collection);
  setVal('itd_cost',d.parent_itd_cost);
  setTimeout(calcVariation, 10);
}
function switchToChild(){
  var picker=document.getElementById('parent-picker-sel');
  var selectedParent=picker?picker.value:'';
  setType('child');
  if(selectedParent){
    var pinput=document.querySelector('input[name="parent_project"]');
    if(pinput){pinput.value=selectedParent;}
  }
}
function submitProject(){
  var form=document.getElementById('main-project-form');
  if(form){form.submit();}
}
function updatePills(){
  for(var i=1;i<=4;i++){
    var pill=document.getElementById('pill-'+i);
    if(!pill)continue;
    pill.className='step-pill'+(i<currentTab?' done':(i===currentTab?' active':''));
  }
}
window.onload=function(){
  updateUI();
  var rbChild=document.getElementById('rb-child');
  var initialType=(rbChild&&rbChild.checked)?'child':'parent';
  var isEditing=!!(document.querySelector('input[name="id"]')&&document.querySelector('input[name="id"]').value);
  if(isEditing){
    var wrap=document.getElementById('parent-field-wrap');
    if(wrap)wrap.style.display=(initialType==='child')?'block':'none';
    var rb=document.getElementById('rb-'+initialType);
    if(rb)rb.checked=true;
    var lblP=document.getElementById('lbl-parent');
    var lblC=document.getElementById('lbl-child');
    var dark='#1a2540',white='#ffffff',plain='#000e24';
    if(lblP){lblP.style.background=(initialType==='parent')?dark:white;lblP.style.color=(initialType==='parent')?white:plain;}
    if(lblC){lblC.style.background=(initialType==='child')?dark:white;lblC.style.color=(initialType==='child')?white:plain;}
  } else {
    setType('parent');
  }
  // Always ensure fields are editable on load — manual entry must never start locked.
  setFieldsReadOnly(false);
};
function calcVariation(){
  var ovEl = document.querySelector('[name="order_value"]');
  var vaEl = document.querySelector('[name="variation_amt"]');
  var rov  = document.getElementById('revised_order_value');
  if(!ovEl || !vaEl || !rov) return;
  var ov = parseFloat(ovEl.value) || 0;
  var va = parseFloat(vaEl.value) || 0;
  rov.value = (ov + va).toFixed(2);
}

document.addEventListener('input', function(e){
  if(e.target.name === 'order_value' || e.target.name === 'variation_amt'){
    calcVariation();
  }
});

document.addEventListener('change', function(e){
  if(e.target.name === 'order_value' || e.target.name === 'variation_amt'){
    calcVariation();
  }
});
function submitDocLink(){
  var link = document.querySelector('[name="doc_link"]')?.value?.trim();
  var cat  = document.querySelector('[name="doc_category"]')?.value?.trim();
  var name = document.querySelector('[name="doc_name"]')?.value?.trim();
  var pid  = document.getElementById('doc-project-id')?.value?.trim();

  if(!link){ alert('Please enter a Document Repository Link.'); return; }
  if(!cat){  alert('Please select a Document Category.'); return; }
  if(!name){ alert('Please enter a Document Name.'); return; }

  var f = document.createElement('form');
  f.method = 'post';
  f.action = '/link_doc';

  [['doc_link',link],['doc_category',cat],['doc_name',name],['project_id',pid]].forEach(function(pair){
    var i = document.createElement('input');
    i.type = 'hidden'; i.name = pair[0]; i.value = pair[1];
    f.appendChild(i);
  });

  document.body.appendChild(f);
  f.submit();

}
</script>"""

    business_lines   = get_lob_list()
    project_managers = get_project_managers()
    lob_list         = get_lob_list()
    project_statuses = get_project_statuses()
    doc_categories   = get_doc_categories()
    parent_projects  = get_parent_project_names()
    parent_opts      = "".join(f'<option value="{pp}">{pp}</option>' for pp in parent_projects)
    child_projects   = get_child_project_names()
    clients          = get_client_list()

    def val(k, d=""): return p.get(k, d)
    def _ref_manager_block(title: str, category: str, values: list) -> str:
        chips = "".join(f"""
        <span style="display:inline-flex;align-items:center;gap:6px;
                      background:{C['surface_low']};border:1px solid {C['outline']};
                      border-radius:14px;padding:4px 10px 4px 12px;font-size:12px;
                      color:{C['on_surface']};margin:0 6px 6px 0">
          {v}
          <button type="button"
            style="font-size:11px;line-height:1;padding:0;border:none;background:none;
                   cursor:pointer;color:#b91c1c"
            onclick="if(confirm('Remove &quot;{v}&quot; from {title}?')){{
              var f=document.createElement('form');f.method='post';f.action='/delete_ref_option';
              var c=document.createElement('input');c.type='hidden';c.name='category';c.value='{category}';
              var n=document.createElement('input');n.type='hidden';n.name='value';n.value='{v}';
              f.appendChild(c);f.appendChild(n);document.body.appendChild(f);f.submit();
            }}">✕</button>
        </span>""" for v in values) or '<p style="font-size:12px;color:#aaa;margin-bottom:8px">No values yet.</p>'

        return f"""
    <div style="border:1px solid {C['outline']};border-radius:8px;padding:14px;background:#fff">
      <p style="font-size:12px;font-weight:700;color:{C['dark']};margin-bottom:10px">{title}</p>
      <div style="margin-bottom:10px">{chips}</div>
      <form method="post" action="/add_ref_option" style="display:flex;gap:6px">
        <input type="hidden" name="category" value="{category}">
        <input type="text" name="value" placeholder="Add new…" required
               style="flex:1;height:32px;border:1px solid {C['outline']};border-radius:6px;
                      padding:0 10px;font-size:12px;outline:none">
        <button type="submit"
                style="height:32px;padding:0 12px;background:{C['dark']};color:#fff;border:none;
                       border-radius:6px;font-size:11px;font-weight:700;cursor:pointer">+ Add</button>
      </form>
    </div>"""
    def _fi(lbl, name, ph="", ftype="text"):
        v = val(name)
        return _field_html(lbl, f'<input type="{ftype}" name="{name}" placeholder="{ph}" value="{v}">')

    def _sl(lbl, name, opts, ph="Select…"):
        return _field_html(lbl, _sel(name, opts, val(name), ph))

    def _sl_autocomplete(lbl, name, opts, ph="Type or select…"):
        dl_id   = f"dl_{name}"
        dl_opts = "".join(f'<option value="{o}">' for o in opts)
        inp = (f'<input type="text" name="{name}" placeholder="{ph}" '
               f'value="{val(name)}" list="{dl_id}" autocomplete="off">'
               f'<datalist id="{dl_id}">{dl_opts}</datalist>')
        return _field_html(lbl, inp)

    def _money(lbl, name, ph="0.00"):
        v = val(name)
        return _field_html(lbl,
            f'<div class="cur-wrap"><span class="pre">₹</span>'
            f'<input type="number" name="{name}" placeholder="{ph}" value="{v}" step="0.01">'
            f'</div>')

    def _pct(lbl, name, dv="0"):
        v = val(name, dv)
        return _field_html(lbl,
            f'<div class="suf-wrap">'
            f'<input type="number" name="{name}" min="0" max="100" value="{v}">'
            f'<span class="suf">%</span></div>')

    doc_rows = ""
    for r in fetch_documents():
        doc_rows += f"""
        <tr>
          <td>{r[1]}</td><td>{r[2]}</td>
          <td><a href="{r[3]}" target="_blank" style="color:{C['dark']};font-weight:600;font-size:13px;text-decoration:none">🔗 View File</a></td>
          <td style="text-align:right">
            <button type="button" class="del-btn"
              onclick="if(confirm('Delete document?')){{var f=document.createElement('form');f.method='post';f.action='/delete_doc';var i=document.createElement('input');i.type='hidden';i.name='doc_id';i.value='{r[0]}';f.appendChild(i);document.body.appendChild(f);f.submit();}}">🗑</button>
          </td>
        </tr>"""
    if not doc_rows:
        doc_rows = '<tr><td colspan="4" style="text-align:center;color:#aaa;padding:18px">No documents yet.</td></tr>'

    all_projects = fetch_all_projects()
    proj_rows_html = ""
    for pr in all_projects[:10]:
        proj_rows_html += f"""
        <tr>
          <td>{pr.get('co_no','')}</td>
          <td style="font-weight:600;color:{C['dark']}">{pr.get('project_name','')}</td>
          <td>{pr.get('project_manager','')}</td>
          <td>{pr.get('project_status','')}</td>
          <td style="text-align:right">
            <a href="/edit_project?id={pr['id']}" style="color:{C['dark']};font-size:12px;font-weight:700;text-decoration:none">✏ Edit</a>
            &nbsp;
            <button class="del-btn" type="button"
              onclick="if(confirm('Delete this project?')){{var f=document.createElement('form');f.method='post';f.action='/delete_project';var i=document.createElement('input');i.type='hidden';i.name='project_id';i.value='{pr['id']}';f.appendChild(i);document.body.appendChild(f);f.submit();}}">🗑</button>
          </td>
        </tr>"""
    if not proj_rows_html:
        proj_rows_html = '<tr><td colspan="5" style="text-align:center;color:#aaa;padding:18px">No projects yet.</td></tr>'

    project_id_hidden = f'<input type="hidden" name="id" value="{val("id")}">' if val("id") else ""

    extra_css = f"""
    <style>
    .tab-nav-bar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 16px 28px;
      border-top: 1px solid {C['outline']};
      background: {C['surface_low']};
      border-radius: 0 0 12px 12px;
    }}
    .nav-left  {{ display: flex; gap: 10px; min-width: 140px; }}
    .nav-right {{ display: flex; gap: 10px; min-width: 140px; justify-content: flex-end; }}
    .nav-center {{
      flex: 1;
      display: flex;
      justify-content: center;
      align-items: center;
      gap: 10px;
    }}
    .step-pill {{
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 5px 12px;
      border-radius: 20px;
      font-size: 10px;
      font-weight: 700;
      letter-spacing: .05em;
      text-transform: uppercase;
      background: {C['surface_high']};
      color: {C['on_var']};
      border: 1px solid {C['outline']};
      transition: all .2s;
      cursor: pointer;
    }}
    .step-pill.active {{
      background: {C['dark']};
      color: #fff;
      border-color: {C['dark']};
    }}
    .step-pill.done {{
      background: #e8f5e9;
      color: #166534;
      border-color: #a5d6a7;
    }}
    .step-num {{
      width: 18px; height: 18px;
      border-radius: 50%;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 10px;
      font-weight: 800;
      background: rgba(255,255,255,.25);
    }}
    .step-pill.active .step-num {{ background: rgba(255,255,255,.2); }}
    .step-pill.done  .step-num {{ background: #a5d6a7; color: #166534; }}
    @keyframes shake{{
      0%,100%{{transform:translateX(0)}}
      20%{{transform:translateX(-6px)}}
      40%{{transform:translateX(6px)}}
      60%{{transform:translateX(-4px)}}
      80%{{transform:translateX(4px)}}
    }}
    </style>"""

    html  = _html_head("Project Data Entry") + ENTRY_CSS + extra_css + "<body>"
    html += _sidebar("data_entry") + _topbar("PROJECT DATA ENTRY", username)
    html += f"""
<div class="main"><div class="canvas">
  <h2 class="page-title">Project Data Entry</h2>
  <p class="page-sub">Log professional surveyor project data with precision and accuracy.</p>
  {alert_html}

  <!-- Upload zone -->
  <form method="post" action="/upload" enctype="multipart/form-data">
  <div class="upload-zone" onclick="document.getElementById('fi').click()">
    <div style="width:52px;height:52px;background:{C['dark']};border-radius:50%;display:flex;
                align-items:center;justify-content:center;margin:0 auto 12px;font-size:24px;color:#fff">📤</div>
    <h3 style="font-size:17px;font-weight:700;color:{C['dark']};margin-bottom:6px">Excel / CSV Upload</h3>
    <p style="font-size:13px;color:#6b7280;margin-bottom:12px">Drag and drop or
       <strong style="color:{C['dark']};text-decoration:underline">browse files</strong></p>
    <div style="display:inline-flex;gap:12px;background:{C['surface_high']};padding:5px 14px;
                border-radius:20px;font-size:11px;font-weight:700;letter-spacing:.07em;color:{C['on_var']}">
      XLSX &nbsp;/&nbsp; CSV &nbsp;/&nbsp; MAX 25MB
    </div>
    <input id="fi" type="file" name="project_file" accept=".xlsx,.csv"
           style="display:none" onchange="this.form.submit()">
  </div>
  </form>

  <!-- Export bar -->
  <div class="exp-bar">
    <a class="exp-btn" href="/export_excel">⬇ Export Excel</a>
    <a class="exp-btn" href="/export_pdf/data_entry">⬇ Export PDF</a>
    <a class="exp-btn" href="/export_csv">⬇ Export CSV</a>
  </div>

  <!-- MAIN CARD -->
  <div class="card">

    <div class="tab-bar">
      <button class="tab-btn" id="tab-1" onclick="switchTab(1)">1. PROJECT DETAILS</button>
      <button class="tab-btn" id="tab-2" onclick="switchTab(2)">2. COST PLAN</button>
      <button class="tab-btn" id="tab-3" onclick="switchTab(3)">3. STATUS</button>
      <button class="tab-btn" id="tab-4" onclick="switchTab(4)">4. DOCUMENTS</button>
    </div>

    <form method="post" action="/submit" id="main-project-form">
      {project_id_hidden}

      <!-- TAB 1 -->
      <div class="section" id="section-1">

        <div id="readonly-banner"
             style="display:none;align-items:center;gap:10px;background:#fffbf0;
                    border:1px solid #fed488;border-radius:8px;padding:10px 16px;
                    margin-bottom:16px;font-size:13px;color:#775a19;font-weight:600">
          <span style="font-size:18px">🔒</span>
          <span>Viewing parent project — all fields are read-only.</span>
        </div>

        <!-- BLOCK 1: PROJECT HIERARCHY -->
        <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:22px;margin-bottom:20px">
          <div class="sec-title" style="margin-bottom:16px">Project Hierarchy</div>
          <div id="parent-picker-wrap" style="display:none;margin-bottom:16px">
            <div style="background:#f0f4ff;border:1px solid #c8d8f8;border-radius:8px;padding:14px 16px">
              <p style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#1e3a8a;margin-bottom:8px">SELECT PARENT PROJECT TO VIEW</p>
              <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
                <select id="parent-picker-sel" onchange="onParentPickerChange(this)"
                        style="flex:1;min-width:260px;padding:9px 12px;border:1.5px solid #93c5fd;
                               border-radius:6px;font-size:13px;font-weight:600;color:#1e3a8a;
                               background:#fff;outline:none;cursor:pointer">
                  <option value="">— Select a parent project —</option>
                  {parent_opts}
                </select>
                <button type="button" onclick="switchToChild()"
                        style="padding:9px 20px;background:#1a2540;color:#fff;border:none;
                               border-radius:6px;font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap">
                  ↪ Switch to Child Entry
                </button>
              </div>
            </div>
          </div>
          <div class="g3">
            {_sl("Line of Business (Hierarchy)","hier_business_line", lob_list, "Select LOB…")}
            <div id="parent-field-wrap">
              {_sl_autocomplete("Parent Project","parent_project", parent_projects, "Select or type parent…")}
            </div>
            <div id="child-field-wrap" style="display:none">
              {_sl_autocomplete("Child Project","child_project", child_projects, "Select or type child…")}
            </div>
          </div>
          <div id="child-value-warning"
               style="display:none;margin-top:10px;background:#fef2f2;
                      border:1.5px solid #fca5a5;border-radius:8px;
                      padding:12px 16px;font-size:12px;color:#991b1b;font-weight:600"></div>
        </div>

        <!-- BLOCK 2: CORE IDENTIFICATION -->
        <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:22px;margin-bottom:20px">
          <div class="sec-title" style="margin-bottom:16px">Core Identification</div>
          <div class="g12">
            {_fi("CO. NO.","co_no","e.g. CO-8872")}
            {_fi("Project Name","project_name","e.g. Western Railway Bridge Inspection")}
          </div>
          <div class="g3 mt">
            {_sl("Line of Business","business_line", business_lines, "Select Business Line…")}
            {_sl("Project Manager","project_manager", project_managers, "Select Manager…")}
            {_sl_autocomplete("Client Name","client_name", clients, "e.g. National Highways Authority")}
          </div>
          <div class="g3 mt">
            {_fi("Order Number","order_number","e.g. ORD-8872")}
            {_fi("Order Issue Date","order_issue_date","",ftype="date")}
            {_money("Order Value (₹)","order_value")}
          </div>
          <div class="g2 mt">
            <div class="fin-card">
              <p style="font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['secondary']};margin-bottom:14px">VARIATION</p>
              {_money("Variation Amount (₹)","variation_amt","0.00")}
            </div>
            <div class="fin-card">
              <p style="font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['dark']};margin-bottom:14px">REVISED ORDER VALUE</p>
              <div class="field">
                <label>Revised Order Value (₹)</label>
                <div class="cur-wrap">
                  <span class="pre">₹</span>
                  <input type="number" name="revised_order_value" id="revised_order_value"
                         value="{val('revised_order_value')}" step="0.01" readonly
                         style="padding-left:24px;background:#e8f0fe;color:#1e3a8a;
                                font-weight:700;cursor:not-allowed;border-color:#93c5fd">
                </div>
                <p style="font-size:11px;color:#3b5bdb;margin-top:8px;font-weight:600">
                  = Order Value + Variation Amount
                </p>
              </div>
            </div>
        </div>
         </div>
        <!-- BLOCK 3: PROJECT TYPE -->
        <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:22px;margin-bottom:20px">
          <div class="sec-title" style="margin-bottom:16px">Project Type &amp; Classification</div>
          <div style="margin-bottom:16px">
            <label style="font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;
                          color:{C['on_var']};display:block;margin-bottom:8px">PROJECT TYPE</label>
            <div style="display:flex;gap:0;border:1px solid {C['outline']};border-radius:8px;overflow:hidden;width:fit-content">
              <label id="lbl-parent"
                     style="display:flex;align-items:center;gap:8px;padding:10px 24px;cursor:pointer;font-size:13px;font-weight:700;transition:all .2s"
                     onclick="setType('parent')">
                <input type="radio" name="project_type" value="parent" id="rb-parent" style="width:15px;height:15px">
                Parent Project
              </label>
              <label id="lbl-child"
                     style="display:flex;align-items:center;gap:8px;padding:10px 24px;cursor:pointer;font-size:13px;font-weight:700;transition:all .2s"
                     onclick="setType('child')">
                <input type="radio" name="project_type" value="child" id="rb-child" style="width:15px;height:15px">
                Child Project
              </label>
            </div>
          </div>
        </div>

       <!-- ALL PROJECTS TABLE — now sits below the step navigation -->
    <div style="padding:0 28px 28px">
      <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;overflow:hidden">
        <div style="padding:12px 18px;border-bottom:1px solid {C['outline']};
                    background:{C['surface_low']};display:flex;align-items:center;justify-content:space-between">
          <span style="font-size:13px;font-weight:700;color:{C['dark']}">All Projects ({len(all_projects)})</span>
          <a href="/data_entry" style="background:{C['dark']};color:#fff;border-radius:6px;padding:5px 12px;font-size:11px;text-decoration:none;font-weight:700">+ New Project</a>
        </div>
        <div style="overflow-x:auto">
          <table class="proj-table">
            <thead><tr>
              <th>CO. No.</th>
              <th>Project Name</th>
              <th>Manager</th>
              <th>Status</th>
              <th style="text-align:right">Actions</th>
            </tr></thead>
            <tbody>{proj_rows_html}</tbody>
          </table>
        </div>
      </div>
    </div>
  </div><!-- end .card -->
       

      </div><!-- end section-1 -->

      <!-- TAB 2: COST PLAN -->
      <div class="section" id="section-2">
        <div class="sec-title">Financial Baseline &amp; Revision</div>
        <div class="g2">
          <div class="fin-card">
            <p style="font-size:11px;font-weight:700;letter-spacing:.06em;color:{C['secondary']};text-transform:uppercase;margin-bottom:14px">BASELINE DATA</p>
            {_money("Baseline EAC (₹)","baseline_eac")}
            <div class="mt">{_pct("Baseline Margin (%)","baseline_margin")}</div>
          </div>
          <div class="fin-card">
            <p style="font-size:11px;font-weight:700;letter-spacing:.06em;color:{C['dark']};text-transform:uppercase;margin-bottom:14px">LATEST REVISION</p>
            {_money("Rev. Latest EAC (₹)","rev_latest_eac","Enter updated EAC")}
            <div class="mt">{_pct("Rev. Latest Margin (%)","rev_latest_margin")}</div>
          </div>
        </div>
      </div><!-- end section-2 -->

      <!-- TAB 3: STATUS -->
      <div class="section" id="section-3">
        <div class="sec-title">Current Execution Status</div>
        <div class="g3">
          {_sl("Project Status","project_status", project_statuses, "Select Status…")}
          {_pct("Physical % Completion","physical_completion")}
          {_money("Up to date Billed Amount (₹)","billed_amount")}
        </div>
        <div class="g12 mt">
          {_money("Collection till Date (₹)","collection_till_date")}
          <div></div>
        </div>
        <div class="info-ban">
          <span style="font-size:18px">ℹ️</span>
          <span>Physical completion should align with the Billed Amount to avoid revenue leakage.</span>
        </div>
      </div><!-- end section-3 -->

 <!-- ─── TAB 4: DOCUMENTS ─── -->
      <div class="section" id="section-4">
        <div class="sec-title">Project Documents</div>
          <input type="hidden" id="doc-project-id" value="{val('id')}">
          <div class="g21">
            {_fi("Document Repository Link","doc_link","e.g. SharePoint or Google Drive URL")}
            {_sl("Document Category","doc_category", doc_categories, "Select Category…")}
          </div>
          <div class="g21 mt" style="align-items:flex-end">
            {_fi("Document Name","doc_name","e.g. Bridge Inspection Report v1.0")}
            <div class="field">
              <label>&nbsp;</label>
              <button type="button" onclick="submitDocLink()"
                      style="height:40px;background:{C['secondary']};color:#fff;border:none;
                             border-radius:6px;font-size:11px;font-weight:700;letter-spacing:.06em;
                             text-transform:uppercase;cursor:pointer;width:100%">
                LINK DOCUMENT
              </button>
            </div>
          </div>


      <!-- Docs table (tab 4 only) -->
      <div id="docs-table-sec" style="display:none;padding:0 28px 20px">
        <p class="sec-title" style="margin-top:20px">Linked Documents</p>
        <div style="border:1px solid {C['outline']};border-radius:8px;overflow:hidden">
          <table class="proj-table">
            <thead><tr>
              <th>Document Name</th><th>Category</th><th>Link</th>
              <th style="text-align:right">Action</th>
            </tr></thead>
            <tbody>{doc_rows}</tbody>
          </table>
        </div>
      </div>
      <!-- NAVIGATION BAR -->
      <div class="tab-nav-bar">
        <div class="nav-left">
          <button id="prev-btn" type="button" onclick="prevTab()" class="btn-prev" style="display:none">
            ◀ Previous
          </button>
        </div>
        <div class="nav-center">
          <button type="button" class="step-pill" id="pill-1" onclick="switchTab(1)">
            <span class="step-num">1</span> Details
          </button>
          <button type="button" class="step-pill" id="pill-2" onclick="switchTab(2)">
            <span class="step-num">2</span> Cost Plan
          </button>
          <button type="button" class="step-pill" id="pill-3" onclick="switchTab(3)">
            <span class="step-num">3</span> Status
          </button>
          <button type="button" class="step-pill" id="pill-4" onclick="switchTab(4)">
            <span class="step-num">4</span> Documents
          </button>
        </div>
        <div class="nav-right">
          <button id="next-btn" type="button" onclick="nextTab()" class="btn-next">
            Next Step ▶
          </button>
          <button id="sub-btn" type="button" onclick="submitProject()" class="btn-sub" style="display:none">
            ✔ Submit Project
          </button>
        </div>
      </div>

    </form>
  </div><!-- end .card -->
  </form>

  <!-- Manage Dropdown Options -->
  <div class="card" style="margin-top:20px">
    <div style="padding:14px 20px;border-bottom:1px solid {C['outline']};background:{C['surface_low']}">
      <span style="font-size:13px;font-weight:700;color:{C['dark']}">⚙ Manage Dropdown Options</span>
      <span style="font-size:12px;color:{C['on_var']};margin-left:8px">Add / remove values from all dropdowns</span>
    </div>
    <div style="padding:20px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px">
    {_ref_manager_block("Line of Business (LOB)",  "lob",    lob_list)}
    {_ref_manager_block("Business Lines",          "biz",    business_lines)}
    {_ref_manager_block("Project Managers",        "pm",     project_managers)}
    {_ref_manager_block("Project Statuses",        "status", project_statuses)}
    {_ref_manager_block("Document Categories",     "doccat", doc_categories)}
    {_ref_manager_block("Expense Categories",      "expcat", get_expense_categories())}
    </div>
  </div>

</div></div>"""

    html += TAB_JS
    return html
# ═══════════════════════════════════════════════════════════════
#  PAGE 3 — EXPENDITURES
# ═══════════════════════════════════════════════════════════════
def _process_expenditure_upload(tmp_path: str) -> str:
    import openpyxl, math

    try:
        wb = openpyxl.load_workbook(tmp_path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
    except Exception as e:
        return f'<div class="alert err">❌ Could not open file: {e}</div>'

    if not rows:
        return '<div class="alert err">❌ File is empty.</div>'

    # Row 0 is the header
    headers = [str(c).strip().lower() if c is not None else "" for c in rows[0]]

    def _col(name):
        try:
            return headers.index(name.lower())
        except ValueError:
            return None

    cm = {
        "project_id":      _col("project_id"),
        "lob":             _col("lob"),
        "data_as_of":      _col("data_as_of"),
        "erp_update_date": _col("erp_update_date"),
        "itd_cogs":        _col("itd_cogs"),
        "itd_djcs":        _col("itd_djcs"),
        "itd_equip":       _col("itd_equip"),
        "itd_manpower":    _col("itd_manpower"),
        "itd_vehicles":    _col("itd_vehicles"),
        "itd_others":      _col("itd_others"),
        "erp_cogs":        _col("erp_cogs"),
        "erp_djcs":        _col("erp_djcs"),
        "erp_equip":       _col("erp_equip"),
        "erp_manpower":    _col("erp_manpower"),
        "erp_vehicles":    _col("erp_vehicles"),
        "erp_others":      _col("erp_others"),
    }

    def _num(row, key):
        idx = cm.get(key)
        if idx is None or idx >= len(row):
            return 0.0
        v = row[idx]
        if v is None or isinstance(v, bool):
            return 0.0
        if isinstance(v, (int, float)):
            f = float(v)
            return 0.0 if math.isnan(f) else f
        s = str(v).strip()
        if not s or s.startswith("="):
            return 0.0
        try:
            return float(s.replace(",", ""))
        except ValueError:
            return 0.0

    def _txt(row, key):
        idx = cm.get(key)
        if idx is None or idx >= len(row):
            return ""
        v = row[idx]
        if v is None or isinstance(v, bool):
            return ""
        s = str(v).strip()
        return "" if s.startswith("=") else s

    def _date(row, key):
        idx = cm.get(key)
        if idx is None or idx >= len(row):
            return datetime.now().date().isoformat()
        v = row[idx]
        if v is None:
            return datetime.now().date().isoformat()
        if hasattr(v, "strftime"):
            return v.strftime("%Y-%m-%d")
        s = str(v).strip()
        if not s or s.startswith("="):
            return datetime.now().date().isoformat()
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                from datetime import datetime as dt
                return dt.strptime(s[:10], fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return s[:10]

    # Maps erp field key -> category name used in erp_updates table
    # Must match exactly what fetch_erp_totals() groups by
    ERP_CATEGORY_MAP = {
        "erp_cogs":     "COGS",
        "erp_djcs":     "DJCS",
        "erp_equip":    "Equipment",
        "erp_manpower": "Manpower",
        "erp_vehicles": "Vehicles",
        "erp_others":   "Others",
    }

    updated = inserted = skipped = 0
    errors = []
    today = datetime.now().date().isoformat()
    now_iso = datetime.now().isoformat()

    for row_num, row in enumerate(rows[1:], start=2):
        if all(c is None for c in row):
            continue

        raw_pid = _txt(row, "project_id")
        if not raw_pid:
            skipped += 1
            continue

        # Look up project — try numeric ID first, then co_no string
        proj = None
        try:
            with ENGINE.connect() as conn:
                try:
                    pid_int = int(float(raw_pid))
                    rows_db = _rows_to_dicts(conn,
                        "SELECT * FROM projects WHERE id = :id LIMIT 1",
                        {"id": pid_int})
                    if rows_db:
                        proj = rows_db[0]
                except (ValueError, TypeError):
                    pass

                if not proj:
                    rows_db = _rows_to_dicts(conn,
                        "SELECT * FROM projects WHERE co_no = :c LIMIT 1",
                        {"c": raw_pid})
                    if rows_db:
                        proj = rows_db[0]
        except Exception as e:
            errors.append(f"Row {row_num} (project_id={raw_pid!r}): DB lookup failed — {e}")
            continue

        if not proj:
            errors.append(f"Row {row_num}: No project found for project_id={raw_pid!r} — skipped.")
            skipped += 1
            continue

        pid = proj["id"]

        lob_val         = _txt(row, "lob") or proj.get("hier_business_line") or proj.get("business_line") or ""
        data_as_of      = _date(row, "data_as_of")      or today
        erp_update_date = _date(row, "erp_update_date") or today

        itd_cogs     = _num(row, "itd_cogs")
        itd_djcs     = _num(row, "itd_djcs")
        itd_equip    = _num(row, "itd_equip")
        itd_manpower = _num(row, "itd_manpower")
        itd_vehicles = _num(row, "itd_vehicles")
        itd_others   = _num(row, "itd_others")

        erp_cogs     = _num(row, "erp_cogs")
        erp_djcs     = _num(row, "erp_djcs")
        erp_equip    = _num(row, "erp_equip")
        erp_manpower = _num(row, "erp_manpower")
        erp_vehicles = _num(row, "erp_vehicles")
        erp_others   = _num(row, "erp_others")

        # ── KEY FIX: if erp_ cols are zero, use itd_ cols as the erp values ──
        # This is what gets shown in the ERP Reconciliation section
        erp_sum = erp_cogs + erp_djcs + erp_equip + erp_manpower + erp_vehicles + erp_others
        if erp_sum == 0:
            erp_cogs     = itd_cogs
            erp_djcs     = itd_djcs
            erp_equip    = itd_equip
            erp_manpower = itd_manpower
            erp_vehicles = itd_vehicles
            erp_others   = itd_others

        erp_vals = {
            "erp_cogs":     erp_cogs,
            "erp_djcs":     erp_djcs,
            "erp_equip":    erp_equip,
            "erp_manpower": erp_manpower,
            "erp_vehicles": erp_vehicles,
            "erp_others":   erp_others,
        }

        exp_fields = {
            "project_id":      pid,
            "lob":             lob_val,
            "data_as_of":      data_as_of,
            "erp_update_date": erp_update_date,
            "itd_cogs":        itd_cogs,
            "itd_djcs":        itd_djcs,
            "itd_equip":       itd_equip,
            "itd_manpower":    itd_manpower,
            "itd_vehicles":    itd_vehicles,
            "itd_others":      itd_others,
            "erp_cogs":        erp_cogs,
            "erp_djcs":        erp_djcs,
            "erp_equip":       erp_equip,
            "erp_manpower":    erp_manpower,
            "erp_vehicles":    erp_vehicles,
            "erp_others":      erp_others,
        }

        try:
            with ENGINE.begin() as conn:

                # ── STEP 1: Delete old erp_updates rows for this project
                #    and replace with fresh values from the upload.
                #    This keeps fetch_erp_totals() accurate — no double-counting.
                conn.execute(text(
                    "DELETE FROM erp_updates WHERE project_id = :pid"
                ), {"pid": pid})

                # ── STEP 2: Insert one erp_updates row per category
                #    fetch_erp_totals() does SUM(amount) GROUP BY category,
                #    so one clean row per category per project is correct.
                for erp_key, category in ERP_CATEGORY_MAP.items():
                    amt = erp_vals[erp_key]
                    if amt:   # skip zero rows to keep the table clean
                        conn.execute(text("""
                            INSERT INTO erp_updates
                                (project_id, category, amount, created_by, created_at)
                            VALUES
                                (:pid, :cat, :amt, :by, :ts)
                        """), {
                            "pid": pid,
                            "cat": category,
                            "amt": amt,
                            "by":  "excel_upload",
                            "ts":  now_iso,
                        })

                # ── STEP 3: Upsert expenditures row ─────────────────────────
                existing = conn.execute(text(
                    "SELECT id FROM expenditures WHERE project_id=:p ORDER BY id DESC LIMIT 1"
                ), {"p": pid}).fetchone()

                if existing:
                    conn.execute(text("""
                        UPDATE expenditures SET
                            lob=:lob,
                            data_as_of=:data_as_of,
                            erp_update_date=:erp_update_date,
                            itd_cogs=:itd_cogs,
                            itd_djcs=:itd_djcs,
                            itd_equip=:itd_equip,
                            itd_manpower=:itd_manpower,
                            itd_vehicles=:itd_vehicles,
                            itd_others=:itd_others,
                            erp_cogs=:erp_cogs,
                            erp_djcs=:erp_djcs,
                            erp_equip=:erp_equip,
                            erp_manpower=:erp_manpower,
                            erp_vehicles=:erp_vehicles,
                            erp_others=:erp_others
                        WHERE project_id=:project_id
                    """), exp_fields)
                    updated += 1
                else:
                    exp_fields["created_at"] = now_iso
                    conn.execute(text("""
                        INSERT INTO expenditures (
                            project_id, lob, data_as_of, erp_update_date,
                            itd_cogs, itd_djcs, itd_equip, itd_manpower,
                            itd_vehicles, itd_others,
                            erp_cogs, erp_djcs, erp_equip, erp_manpower,
                            erp_vehicles, erp_others, created_at
                        ) VALUES (
                            :project_id, :lob, :data_as_of, :erp_update_date,
                            :itd_cogs, :itd_djcs, :itd_equip, :itd_manpower,
                            :itd_vehicles, :itd_others,
                            :erp_cogs, :erp_djcs, :erp_equip, :erp_manpower,
                            :erp_vehicles, :erp_others, :created_at
                        )
                    """), exp_fields)
                    inserted += 1

                # ── STEP 4: Sync itd_cost on the project row ─────────────────
                total_itd = (itd_cogs + itd_djcs + itd_equip +
                             itd_manpower + itd_vehicles + itd_others)
                if total_itd:
                    conn.execute(text(
                        "UPDATE projects SET itd_cost=:v WHERE id=:id"
                    ), {"v": total_itd, "id": pid})

        except Exception as e:
            errors.append(f"Row {row_num} (project_id={raw_pid!r}): {e}")

    msg = (f"✅ Expenditure upload complete — "
           f"{inserted} new records inserted, {updated} existing records updated"
           + (f", {skipped} rows skipped" if skipped else "") + ".")

    if errors:
        err_detail = "<br>".join(errors[:10])
        return (f'<div class="alert ok">{msg}</div>'
                f'<div class="alert err">⚠ Some rows had issues:<br>{err_detail}</div>')
    return f'<div class="alert ok">{msg}</div>'
def expenditures(alert_html="", prefill=None, username="Admin") -> str:
    p     = prefill or {}
    today = date.today().isoformat()

    lob_list   = get_lob_list()
    exp_cats   = get_expense_categories()
    cat_colors = get_expense_category_colors()
    proj_rows  = fetch_projects_for_select()

    sel_pid  = p.get("project_id","")
    sel_lob  = p.get("lob","")

    # ── Filter projects by LOB if selected ──────────────────────────────────
    if sel_lob:
        filtered_rows = [r for r in proj_rows
                         if str(r[6] or "").strip().lower() == sel_lob.strip().lower()]
    else:
        filtered_rows = proj_rows

    proj_opts = "".join(
        f'<option value="{r[0]}" {"selected" if str(sel_pid)==str(r[0]) else ""}>'
        f'{r[1] or ""} | {r[2] or ""}</option>'
        for r in filtered_rows
    ) or '<option value="">— No projects for this LOB —</option>'

    sel_proj = next((r for r in proj_rows if str(r[0])==str(sel_pid)), None)

    meta_client = sel_proj[3] if sel_proj else "—"
    meta_ordno  = sel_proj[4] if sel_proj else "—"
    sel_rev_ov  = ((_n(sel_proj[8]) or (_n(sel_proj[5]) + _n(sel_proj[7]))) if sel_proj else 0)
    meta_ordval = f"₹ {sel_rev_ov:,.0f}" if sel_rev_ov else "—"
    meta_lob    = sel_proj[6] if sel_proj else "—"

    lob_opts = "".join(
        f'<option value="{lob}" {"selected" if sel_lob == lob else ""}>{lob}</option>'
        for lob in lob_list
    )

    existing_exp = {}
    if sel_pid:
        existing_exp = fetch_expenditure_for_project(int(sel_pid))

    erp_totals = fetch_erp_totals(int(sel_pid)) if sel_pid else {}

    itd_defaults = {
        "COGS":      "erp_cogs",
        "DJCS":      "erp_djcs",
        "Equipment": "erp_equip",
        "Manpower":  "erp_manpower",
        "Vehicles":  "erp_vehicles",
        "Others":    "erp_others",
    }

    # ── ERP Reconciliation columns ───────────────────────────────────────────
    # cumulative_val  = total from erp_updates  (shown in the blue card)
    # stored_erp_val  = snapshot in expenditures (shown in input as readonly)
    erp_cols_html = ""
    for cat in exp_cats:
        col_key           = f"erp_{cat.lower().replace(' ','_').replace('-','_')}"
        legacy            = itd_defaults.get(cat, col_key)
        stored_erp_key    = legacy                          # e.g. "erp_cogs"
        cumulative_val    = erp_totals.get(cat, 0)         # from erp_updates SUM
        stored_erp_val    = _n(existing_exp.get(stored_erp_key, 0))  # from expenditures

        # The display card shows the cumulative total from erp_updates
        # The input shows the stored snapshot — pre-filled and readonly after upload
        erp_cols_html += f"""
<div style="min-width:0">
  <!-- Cumulative total card -->
  <div style="background:{C['surface_low']};border:1px solid {C['outline']};border-radius:8px;
              padding:12px 14px;margin-bottom:10px">
    <div style="font-size:9px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;
                color:{C['on_var']};margin-bottom:6px">{cat.upper()}</div>
    <div style="font-size:16px;font-weight:800;color:{C['dark']};letter-spacing:-.02em">
      ₹ {_n(cumulative_val):,.0f}
    </div>
    <div style="font-size:10px;color:{C['on_var']};margin-top:3px;opacity:.7">Cumulative ITD</div>
  </div>
  <!-- ERP input — pre-filled with stored snapshot value -->
  <div style="font-size:9px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;
              color:{C['on_var']};margin-bottom:6px">ERP {cat.upper()}</div>
  <div style="position:relative">
    <span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);
                 font-size:12px;color:{C['on_var']};pointer-events:none">₹</span>
    <input type="number" id="{legacy}" name="{legacy}"
           value="{stored_erp_val if stored_erp_val else 0}"
           step="0.01"
           data-existing="{_n(cumulative_val)}"
           oninput="calcErpTotal()"
           style="width:100%;padding:9px 9px 9px 24px;border:1px solid {C['outline']};
                  border-radius:6px;font-size:13px;font-weight:600;background:#fff;
                  outline:none;box-sizing:border-box">
  </div>
</div>"""

    # ── Donut chart — use erp_totals (cumulative) not existing_exp snapshot ──
    donut_segs = []
    palette    = ["#003f88","#d00000","#ffba08","#3a86ff","#06a77d","#7209b7","#fb5607","#1a1a2e"]

    erp_chart_vals = {cat: _n(erp_totals.get(cat, 0)) for cat in exp_cats}
    total_itd_chart = sum(erp_chart_vals.values())

    # Fallback to expenditures snapshot if erp_updates is empty
    if total_itd_chart == 0:
        legacy_map_donut = {
            "COGS":      "erp_cogs",
            "DJCS":      "erp_djcs",
            "Equipment": "erp_equip",
            "Manpower":  "erp_manpower",
            "Vehicles":  "erp_vehicles",
            "Others":    "erp_others",
        }
        erp_chart_vals = {
            cat: _n(existing_exp.get(
                legacy_map_donut.get(cat, f"erp_{cat.lower().replace(' ','_')}"), 0
            ))
            for cat in exp_cats
        }
        total_itd_chart = sum(erp_chart_vals.values())

    total_itd = total_itd_chart or 1
    for i, cat in enumerate(exp_cats):
        bg        = palette[i % len(palette)]
        val       = erp_chart_vals[cat]
        pct_exact = (val / total_itd * 100) if total_itd else (100 / max(len(exp_cats), 1))
        pct_disp  = round(pct_exact)
        donut_segs.append((cat, bg, pct_exact, pct_disp, val))

    if total_itd_chart >= 10_00_000:
        total_itd_label = f"₹{total_itd_chart/10_00_000:.1f}M"
    elif total_itd_chart >= 1_00_000:
        total_itd_label = f"₹{total_itd_chart/1_00_000:.1f}L"
    else:
        total_itd_label = f"₹{total_itd_chart:,.0f}"

    r_val = 15.9155
    sw    = 14/140*36
    cx, cy = 18, 18
    import math

    def _arc_point(pct):
        ang = math.radians(-90 + (pct / 100.0) * 360)
        return cx + r_val * math.cos(ang), cy + r_val * math.sin(ang)

    arc_parts = [f'<circle cx="{cx}" cy="{cy}" r="{r_val}" fill="none" stroke="#f2f4f6" stroke-width="{sw:.2f}"/>']
    cur = 0.0
    for label, colour, pct_exact, pct_disp, valamt in donut_segs:
        if pct_exact <= 0.0001:
            continue
        start_pct, end_pct = cur, cur + pct_exact
        x1, y1 = _arc_point(start_pct)
        x2, y2 = _arc_point(min(end_pct, 99.999))
        large_arc = 1 if pct_exact > 50 else 0
        arc_parts.append(
            f'<path d="M {x1:.4f} {y1:.4f} A {r_val} {r_val} 0 {large_arc} 1 {x2:.4f} {y2:.4f}" '
            f'fill="none" stroke="{colour}" stroke-width="{sw:.2f}" stroke-linecap="butt" '
            f'style="cursor:pointer" '
            f'onmousemove="showDonutTip(event,\'{label}\',\'{_n(valamt):,.0f}\',\'{pct_disp}\')" '
            f'onmouseleave="hideDonutTip()"/>'
        )
        cur = end_pct
    donut_svg = (
        f'<svg width="160" height="160" viewBox="0 0 36 36">'
        + "".join(arc_parts)
        + '</svg>'
    )

    legend_items = "".join(
        f'<div style="display:flex;align-items:center;justify-content:space-between;'
        f'padding:7px 0;border-bottom:1px solid {C["outline"]}">'
        f'<div style="display:flex;align-items:center;gap:8px">'
        f'<span style="width:10px;height:10px;border-radius:50%;background:{col};'
        f'flex-shrink:0;display:inline-block"></span>'
        f'<span style="font-size:13px;color:{C["on_var"]}">{lbl}</span>'
        f'</div>'
        f'<span style="font-size:13px;font-weight:700;color:{C["dark"]}">{pct_disp}%</span>'
        f'</div>'
        for lbl, col, pct_exact, pct_disp, valamt in donut_segs
    )

    data_as_of = existing_exp.get("data_as_of") or p.get("data_as_of", today)
    erp_update = existing_exp.get("erp_update_date") or p.get("erp_update_date", today)
    cat_opts   = "".join(f'<option value="{c}">{c}</option>' for c in exp_cats)
    legacy_map = {
        "COGS":      "erp_cogs",
        "DJCS":      "erp_djcs",
        "Equipment": "erp_equip",
        "Manpower":  "erp_manpower",
        "Vehicles":  "erp_vehicles",
        "Others":    "erp_others",
    }
    erp_field_ids = [legacy_map.get(c, f"erp_{c.lower().replace(' ','_')}") for c in exp_cats]
    erp_ids_js    = "['" + "','".join(erp_field_ids) + "']"

    css = f"""<style>
*{{box-sizing:border-box}}
.exp-title{{font-size:24px;font-weight:800;color:{C['dark']};letter-spacing:-.02em}}

/* ── Filter bar — no button, dropdowns side by side ── */
.sel-bar{{
  background:#fff;border:1px solid {C['outline']};border-radius:10px;
  padding:16px 20px;display:flex;align-items:flex-end;gap:16px;
  margin-bottom:20px;
}}
.sel-bar .field{{flex:1;min-width:0}}
.sel-bar .field label{{
  font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
  color:{C['on_var']};display:block;margin-bottom:6px
}}
.sel-bar .field select{{
  width:100%;padding:9px 12px;border:1px solid {C['outline']};border-radius:6px;
  font-size:13px;background:#fff;outline:none;color:{C['dark']};cursor:pointer;
  transition:border-color .15s;
}}
.sel-bar .field select:focus{{border-color:{C['dark']}}}

.meta-strip{{display:grid;grid-template-columns:repeat(4,1fr);background:#fff;
             border:1px solid {C['outline']};border-radius:10px;overflow:hidden}}
.meta-cell{{padding:16px 20px;border-right:1px solid {C['outline']}}}
.meta-cell:last-child{{border-right:none}}
.mlbl{{font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
       color:{C['on_var']};margin-bottom:5px}}
.mval{{font-size:15px;font-weight:700;color:{C['dark']}}}
.side-panel{{position:fixed;inset-y:0;right:0;width:400px;background:#fff;
             border-left:1px solid {C['outline']};box-shadow:-4px 0 32px rgba(0,0,0,.10);
             z-index:300;transform:translateX(100%);transition:transform .3s ease;
             display:flex;flex-direction:column}}
.side-panel.open{{transform:translateX(0)}}
.sp-hdr{{padding:20px 24px;border-bottom:1px solid {C['outline']};
         background:{C['surface_low']};display:flex;align-items:center;justify-content:space-between}}
.sp-body{{flex:1;padding:24px;overflow-y:auto}}
.sp-fld{{margin-bottom:16px}}
.sp-fld label{{font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
               color:{C['on_var']};display:block;margin-bottom:6px}}
.sp-fld input,.sp-fld select{{width:100%;padding:9px 12px;border:1px solid {C['outline']};
                               border-radius:6px;font-size:13px;background:#fff;
                               outline:none;color:{C['dark']}}}
.sp-g2{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
.sp-foot{{padding:16px 24px;border-top:1px solid {C['outline']};
          background:{C['surface_low']};display:flex;gap:10px}}
.overlay{{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:299;display:none}}
.overlay.show{{display:block}}
.erp-card{{background:#fff;border:1px solid {C['outline']};border-radius:10px;overflow:hidden}}
.erp-hdr{{background:{C['surface_low']};padding:18px 28px;border-bottom:1px solid {C['outline']};
          display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:16px}}
.data-tbl{{width:100%;border-collapse:collapse;font-size:13px}}
.data-tbl thead tr{{background:{C['surface_low']};border-bottom:1px solid {C['outline']}}}
.data-tbl th{{padding:10px 16px;font-size:9px;font-weight:700;text-transform:uppercase;
              letter-spacing:.07em;color:{C['on_var']};text-align:left}}
.data-tbl tbody tr{{border-bottom:1px solid {C['outline']}}}
.data-tbl tbody tr:last-child{{border-bottom:none}}
.data-tbl tbody tr:hover{{background:{C['surface_low']}}}
.exp-upload-zone{{
  border:2px dashed {C['outline']};border-radius:12px;padding:22px 28px;
  background:#fff;text-align:center;cursor:pointer;margin-bottom:20px;
  transition:border-color .2s,background .2s;display:flex;align-items:center;
  gap:20px;flex-wrap:wrap;justify-content:center;
}}
.exp-upload-zone:hover{{border-color:{C['dark']};background:#f0f4ff}}
.exp-upload-zone .uz-icon{{
  width:48px;height:48px;background:{C['dark']};border-radius:50%;
  display:flex;align-items:center;justify-content:center;
  font-size:22px;color:#fff;flex-shrink:0;
}}
.exp-upload-zone .uz-text{{text-align:left}}
.exp-upload-zone .uz-text h4{{font-size:15px;font-weight:700;color:{C['dark']};margin-bottom:4px}}
.exp-upload-zone .uz-text p{{font-size:12px;color:#6b7280}}
.exp-upload-zone .uz-badge{{
  display:inline-flex;gap:10px;background:{C['surface_high']};
  padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;
  letter-spacing:.07em;color:{C['on_var']};white-space:nowrap;
}}
</style>"""

    js = f"""<script>
/* ── Auto-submit on LOB change: repopulate project list ── */
function onLobChange(sel) {{
  document.getElementById('exp-filter-form').submit();
}}

/* ── Auto-submit on project change: load that project's data ── */
function onProjectChange(sel) {{
  if (sel.value) {{
    document.getElementById('exp-filter-form').submit();
  }}
}}

function openPanel(){{
  document.getElementById('sp-cat').value='{exp_cats[0] if exp_cats else ""}';
  document.getElementById('sp-desc').value='';
  document.getElementById('sp-qty').value='';
  document.getElementById('sp-ucost').value='';
  document.getElementById('sp-total').textContent='₹ 0.00';
  document.getElementById('side-panel').classList.add('open');
  document.getElementById('overlay').classList.add('show');
}}
function closePanel(){{
  document.getElementById('side-panel').classList.remove('open');
  document.getElementById('overlay').classList.remove('show');
}}
document.addEventListener('keydown',function(e){{if(e.key==='Escape')closePanel();}});
function calcTotal(){{
  var q=parseFloat(document.getElementById('sp-qty').value)||0;
  var u=parseFloat(document.getElementById('sp-ucost').value)||0;
  document.getElementById('sp-total').textContent='₹ '+(q*u).toLocaleString('en-IN',{{minimumFractionDigits:2}});
}}
function calcErpTotal(){{
  var ids={erp_ids_js};
  var total=0;
  ids.forEach(function(id){{
    var el=document.getElementById(id);
    if(el) total += (parseFloat(el.dataset.existing)||0) + (parseFloat(el.value)||0);
  }});
  document.getElementById('erp-total').value=total.toLocaleString('en-IN',{{minimumFractionDigits:2}});
}}
function showDonutTip(e, label, value, pct){{
  var tip  = document.getElementById('donut-tooltip');
  var wrap = document.getElementById('donut-wrap');
  var rect = wrap.getBoundingClientRect();
  tip.innerHTML = '<div style="font-weight:700;margin-bottom:2px">'+label+'</div>'
                + '<div>₹ '+value+' ('+pct+'%)</div>';
  tip.style.left = (e.clientX - rect.left) + 'px';
  tip.style.top  = (e.clientY - rect.top - 8) + 'px';
  tip.style.opacity = '1';
}}
function hideDonutTip(){{
  document.getElementById('donut-tooltip').style.opacity = '0';
}}
window.addEventListener('load', calcErpTotal);
</script>"""

    html  = _html_head("Expenditure Updates") + css + "<body>"
    html += _sidebar("expenditures") + _topbar("EXPENDITURE & ITD COST BREAKDOWN", username)
    html += f"""
<div class="main"><div class="canvas" style="max-width:1200px">

  <!-- Page header -->
  <div style="border-bottom:1px solid {C['outline']};padding-bottom:20px;margin-bottom:24px">
    <h2 class="exp-title">Expenditure Updates</h2>
    <p style="font-size:13px;color:#6b7280;margin-top:4px">
      Comprehensive analysis and ERP reconciliation for current projects.
    </p>
  </div>

  {alert_html}

  <!-- ── EXCEL UPLOAD ZONE ── -->
  <form method="post" action="/upload_expenditure" enctype="multipart/form-data">
    <div class="exp-upload-zone" onclick="document.getElementById('exp-fi').click()">
      <div class="uz-icon">📤</div>
      <div class="uz-text">
        <h4>Bulk Upload via Excel</h4>
        <p>Upload your expenditure template — columns will be matched by project_id / CO.No
           and saved directly to the database.</p>
      </div>
      <div>
        <div class="uz-badge">XLSX &nbsp;/&nbsp; MAX 25MB</div>
        <p style="font-size:11px;color:#9ca3af;margin-top:6px;text-align:center">
          Click or drag to upload
        </p>
      </div>
      <input id="exp-fi" type="file" name="expenditure_file" accept=".xlsx"
             style="display:none" onchange="this.form.submit()">
    </div>
  </form>

  <!-- ── FILTER BAR — no button, auto-submits on change ── -->
  <form method="post" action="/expenditures_retrieve" id="exp-filter-form">
    <div class="sel-bar">
      <div class="field">
        <label>Line of Business (LOB)</label>
        <select name="lob" onchange="onLobChange(this)">
          <option value="">All LOB</option>
          {lob_opts}
        </select>
      </div>
      <div class="field">
        <label>Active Project</label>
        <select name="project_id" onchange="onProjectChange(this)">
          <option value="">— Select a project —</option>
          {proj_opts}
        </select>
      </div>
    </div>
  </form>

  <!-- Project meta strip — only shown when a project is selected -->
  {'<div class="meta-strip" style="margin-bottom:28px">' +
    f'<div class="meta-cell"><div class="mlbl">Client Name</div><div class="mval">{meta_client}</div></div>' +
    f'<div class="meta-cell"><div class="mlbl">Order Number</div><div class="mval">{meta_ordno}</div></div>' +
    f'<div class="meta-cell"><div class="mlbl">Order Value</div><div class="mval">{meta_ordval}</div></div>' +
    f'<div class="meta-cell"><div class="mlbl">LOB</div><div class="mval">{meta_lob}</div></div>' +
   '</div>'
   if sel_pid else
   f'<div style="background:{C["surface_low"]};border:1px solid {C["outline"]};border-radius:10px;'
   f'padding:24px;text-align:center;color:{C["on_var"]};font-size:13px;margin-bottom:28px">'
   f'← Select a project above to view its expenditure data</div>'
  }

  <!-- Donut chart + legend -->
  <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;
              padding:28px 32px;display:flex;align-items:center;gap:40px;margin-bottom:28px">
    <div style="position:relative;width:160px;height:160px;flex-shrink:0" id="donut-wrap">
      {donut_svg}
      <div style="position:absolute;inset:0;display:flex;flex-direction:column;
                  align-items:center;justify-content:center;text-align:center;
                  gap:2px;pointer-events:none">
        <span style="font-size:18px;font-weight:800;color:{C['dark']};
                     letter-spacing:-.03em">{total_itd_label}</span>
        <span style="font-size:9px;font-weight:700;text-transform:uppercase;
                     color:{C['on_var']};letter-spacing:.06em">Total ITD</span>
      </div>
      <div id="donut-tooltip"
           style="position:absolute;pointer-events:none;opacity:0;transition:opacity .12s;
                  background:{C['dark']};color:#fff;font-size:11px;padding:6px 10px;
                  border-radius:6px;white-space:nowrap;z-index:50;transform:translateY(-100%)">
      </div>
    </div>
    <div style="flex:1;display:grid;grid-template-columns:1fr 1fr;gap:0 40px">
      {legend_items}
    </div>
  </div>

  <!-- ERP Cost Reconciliation -->
  <div class="erp-card" style="margin-bottom:28px">
    <div class="erp-hdr">
      <div style="display:flex;align-items:center;gap:10px">
        <span style="font-size:22px">⚖️</span>
        <span style="font-size:18px;font-weight:700;color:{C['dark']}">ERP Cost Reconciliation</span>
      </div>
      <div style="display:flex;gap:32px;align-items:flex-start">
        <div>
          <div style="font-size:9px;font-weight:700;text-transform:uppercase;
                      letter-spacing:.08em;color:{C['on_var']};margin-bottom:4px">
            Data Available as of
          </div>
          <div style="font-size:13px;font-weight:700;color:{C['dark']}">{data_as_of}</div>
        </div>
        <div>
          <div style="font-size:9px;font-weight:700;text-transform:uppercase;
                      letter-spacing:.08em;color:{C['on_var']};margin-bottom:4px">
            Last ERP Update Date
          </div>
          <div style="font-size:13px;font-weight:700;color:{C['dark']}">{erp_update}</div>
        </div>
      </div>
    </div>

    <form method="post" action="/save_expenditure">
      <input type="hidden" name="project_id"      value="{sel_pid}">
      <input type="hidden" name="lob"             value="{sel_lob}">
      <input type="hidden" name="data_as_of"      value="{data_as_of}">
      <input type="hidden" name="erp_update_date" value="{erp_update}">

      <div style="padding:24px 28px">
        <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:20px">
          {erp_cols_html}
        </div>
      </div>

      <!-- Total footer bar -->
      <div style="background:{C['dark']};color:#fff;padding:20px 28px;display:flex;
                  align-items:center;justify-content:space-between;gap:24px;flex-wrap:wrap">
        <div>
          <div style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;
                      opacity:.75;margin-bottom:3px">Calculated Total Cumulative ITD Cost (A/c)</div>
          <div style="font-size:11px;opacity:.5;font-style:italic">
            Aggregated value for system-wide verification &amp; auditing
          </div>
        </div>
        <div style="display:flex;align-items:center;gap:16px">
          <div style="position:relative;width:210px">
            <span style="position:absolute;left:14px;top:50%;transform:translateY(-50%);
                         font-size:18px;font-weight:700;color:{C['dark']};pointer-events:none">₹</span>
            <input type="text" id="erp-total" readonly value="0.00"
                   style="width:100%;padding:11px 11px 11px 34px;background:#fff;
                          color:{C['dark']};border:none;border-radius:7px;font-size:18px;
                          font-weight:800;outline:none;letter-spacing:-.02em">
          </div>
          <button type="submit"
                  style="background:rgba(255,255,255,.15);color:#fff;
                         border:1.5px solid rgba(255,255,255,.4);border-radius:7px;
                         padding:11px 28px;font-size:11px;font-weight:700;
                         letter-spacing:.07em;text-transform:uppercase;cursor:pointer;
                         backdrop-filter:blur(4px)"
                  onmouseover="this.style.background='rgba(255,255,255,.25)'"
                  onmouseout="this.style.background='rgba(255,255,255,.15)'">
            UPDATE DATA
          </button>
        </div>
      </div>
    </form>
  </div>

  <!-- Add Entry side panel -->
  <div class="side-panel" id="side-panel">
    <div class="sp-hdr">
      <span style="font-size:16px;font-weight:700;color:{C['dark']}">Add Expense Entry</span>
      <button onclick="closePanel()" type="button"
              style="background:none;border:none;font-size:20px;cursor:pointer;
                     color:{C['on_var']};line-height:1">✕</button>
    </div>
    <div class="sp-body">
      <form method="post" action="/save_entry" id="entry-form">
        <input type="hidden" name="project_id"     value="{sel_pid}">
        <input type="hidden" name="expenditure_id" value="0">
        <div class="sp-fld">
          <label>Expense Category</label>
          <select name="category" id="sp-cat">{cat_opts}</select>
        </div>
        <div class="sp-fld">
          <label>Description</label>
          <input type="text" name="description" id="sp-desc"
                 placeholder="e.g. Site Infrastructure Supplies">
        </div>
        <div class="sp-g2">
          <div class="sp-fld">
            <label>Quantity</label>
            <input type="number" name="quantity" id="sp-qty"
                   placeholder="0" step="1" min="0" oninput="calcTotal()">
          </div>
          <div class="sp-fld">
            <label>Unit Cost (₹)</label>
            <input type="number" name="unit_cost" id="sp-ucost"
                   placeholder="0.00" step="0.01" min="0" oninput="calcTotal()">
          </div>
        </div>
        <div style="background:#f0f4ff;border:1px solid #c8d8f8;border-radius:8px;
                    padding:16px 18px;margin-top:8px">
          <div style="font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
                      color:{C['on_var']};margin-bottom:6px">Line Total</div>
          <div style="font-size:24px;font-weight:800;color:{C['dark']};
                      letter-spacing:-.02em" id="sp-total">₹ 0.00</div>
        </div>
      </form>
    </div>
    <div class="sp-foot">
      <button onclick="closePanel()" type="button"
              style="flex:1;padding:11px;border:1px solid {C['outline']};border-radius:6px;
                     background:#fff;color:{C['on_var']};font-size:11px;font-weight:700;
                     text-transform:uppercase;letter-spacing:.05em;cursor:pointer">Cancel</button>
      <button onclick="document.getElementById('entry-form').submit()" type="button"
              style="flex:1;padding:11px;background:{C['dark']};color:#fff;border:none;
                     border-radius:6px;font-size:11px;font-weight:700;text-transform:uppercase;
                     letter-spacing:.05em;cursor:pointer">Save Entry</button>
    </div>
  </div>
  <div class="overlay" id="overlay" onclick="closePanel()"></div>

{js}
</body></html>"""
    return html
# ═══════════════════════════════════════════════════════════════
#  PAGE 4 — FINANCIALS  (UPDATED with child budget/expense rows)
# ═══════════════════════════════════════════════════════════════
def export_financials_pdf_bytes(lob_filter: str = "") -> bytes:
    projects = fetch_all_projects(lob_filter=lob_filter)
    parents  = [p for p in projects if p.get("project_type","parent") != "child"]
    rows, total_ov = [], 0.0
    for p in parents:
        ov  = _rev_ov(p); itd = _n(p.get("itd_cost")); col = _n(p.get("collection_till_date"))
        ebf = round(1.35*itd, 2)
        total_ov += ov
        rows.append([p.get("co_no","—"), p.get("project_name","—"), p.get("project_manager","—"),
                     f"{ov:,.0f}", f"{itd:,.0f}", f"{col:,.0f}", f"{ebf:,.0f}"])
    return _build_table_pdf(
        "Financial Performance Report",
        f"Project performance analysis — {lob_filter or 'All LOB'}",
        ["CO.No","Project","PM","Order Value","ITD","Collection","EBF (1.35x)"],
        rows, footer_label="Total Order Value", footer_value=total_ov)

def financials(alert_html="", prefill=None, username="Admin") -> str:
    p       = prefill or {}
    lob_all = ["All Lines of Business"] + get_lob_list()
    sel_lob = p.get("lob", "All Lines of Business")
    lob_opts = "".join(
        f'<option {"selected" if lo == sel_lob else ""}>{lo}</option>'
        for lo in lob_all
    )

    lob_filt = "" if sel_lob == "All Lines of Business" else sel_lob
    projects  = fetch_all_projects(lob_filter=lob_filt)

    # ── Only parent projects in the main table ──
    parent_list = [pr for pr in projects if pr.get("project_type", "parent") != "child"]
    all_projects_count = len(fetch_all_projects())  # total unfiltered for "Showing X of Y"

    # ── Totals ──
    total_ov  = sum(_n(pr.get("revised_order_value") or (_n(pr.get("order_value")) + _n(pr.get("variation_amt")))) for pr in parent_list)
    total_itd = sum(_n(pr.get("itd_cost"))             for pr in parent_list)   # Col Z
    total_col = sum(_n(pr.get("collection_till_date")) for pr in parent_list)
    total_ebf = round(1.35 * total_itd, 2)                                       # Col AA

    # ── FIXED: avg_perf = Collection as % of Order Value (display KPI) ──
    avg_perf = (total_col / total_ov * 100) if total_ov else 0

    # ── FIXED: ratio = Collection / EBF  (meaningful: ≥1.0 means collection covers EBF) ──
    # 1.35 is the EBF multiplier; if Collection >= EBF it means we've recovered costs+margin
    ratio     = total_col / total_ebf if total_ebf else 0
    ratio_pct = ratio * 100  # for display purposes

    # ── Threshold: ratio >= 1.0 means collection has covered the 1.35x cost target ──
    ratio_str     = f"{ratio:.2f}x"
    ratio_lbl     = "Profit — Collection ≥ EBF" if ratio >= 1.0 else "Loss — Collection < EBF"
    ratio_badge_cls = "background:#f0fdf4;color:#166534" if ratio >= 1.0 else "background:#fef2f2;color:#991b1b"

    # Sub-label showing how far above/below threshold
    if ratio >= 1.0:
        surplus = total_col - total_ebf
        ratio_sub = f"Surplus of ₹{surplus/1e5:,.1f}L above EBF target"
    else:
        deficit = total_ebf - total_col
        ratio_sub = f"Deficit of ₹{deficit/1e5:,.1f}L below EBF target"

    n_rows = len(parent_list)

    def to_lakhs(v):
        """Convert raw ₹ value to Lakhs, formatted to 2 decimal places."""
        return round(_n(v) / 1e5, 2)

    # ── Build table rows ──
    proj_rows_html = ""
    for pr in parent_list:
        co   = pr.get("co_no")           or "—"
        name = pr.get("project_name")    or "—"
        pm   = pr.get("project_manager") or "—"
        st   = pr.get("project_status")  or "Active"
        pid  = pr.get("id", "")

        rev_ov = _n(pr.get("revised_order_value")) or (_n(pr.get("order_value")) + _n(pr.get("variation_amt")))
        ov_l  = to_lakhs(rev_ov)                          # Col I — Revised Order Value
        itd_l = to_lakhs(pr.get("itd_cost"))              # Col Z — Total ITD Cost
        col_l = to_lakhs(pr.get("collection_till_date"))  # Col S
        ebf_l = round(1.35 * itd_l, 2)                    # Col AA = 1.35 × ITD

        # ── FIXED: Collection Performance = (Collection - EBF) / EBF × 100
        # Positive = collection exceeded the 1.35x billing floor (PROFIT)
        # Negative = collection is short of the 1.35x billing floor (LOSS)
        if ebf_l:
            perf = (col_l - ebf_l) / ebf_l * 100
        else:
            perf = 0.0

        # ── FIXED: EBF badge color based on col_l vs ebf_l (not trivial ebf/itd=1.35) ──
        if col_l >= ebf_l:
            # Collection has met or exceeded the EBF target → PROFIT
            ebf_cls = "color:#166534;background:#f0fdf4;border:1px solid #bbf7d0"
        elif col_l >= ebf_l * 0.80:
            # Collection is within 80–99% of EBF → WARNING
            ebf_cls = "color:#92400e;background:#fffbeb;border:1px solid #fde68a"
        else:
            # Collection is below 80% of EBF → LOSS
            ebf_cls = "color:#991b1b;background:#fef2f2;border:1px solid #fecaca"

        # Coll. Perf sign + color
        if perf >= 0:
            perf_str   = f"+{perf:.1f}%"
            perf_color = "#166534"
            perf_bg    = "#f0fdf4"
        else:
            perf_str   = f"{perf:.1f}%"
            perf_color = "#991b1b"
            perf_bg    = "#fef2f2"

        # Status badge
        if st in ("Cancelled", "Hold"):
            st_cls = "background:#ffdad6;color:#ba1a1a"
        elif st in ("Closed", "Execution Closed"):
            st_cls = f"background:{C['surface_high']};color:{C['on_var']}"
        else:
            st_cls = "background:#fde68a;color:#92400e"

        fp = (pr.get('folder_path') or '').replace("'", "\\'")
        folder_link = (
            f'&nbsp;<a href="javascript:void(0)" onclick="openProjectFolder(\'{fp}\')" '
            f'style="color:{C["secondary"]};font-size:12px;font-weight:700;text-decoration:none" '
            f'title="{fp}">📂 Folder</a>'
        ) if fp else ""
        proj_rows_html += f"""
        <tr>
          <td>{pr.get('co_no','')}</td>
          <td style="font-weight:600;color:{C['dark']}">{pr.get('project_name','')}</td>
          <td>{pr.get('project_manager','')}</td>
          <td>{pr.get('project_status','')}</td>
          <td style="text-align:right">
            <a href="/edit_project?id={pr['id']}" style="color:{C['dark']};font-size:12px;font-weight:700;text-decoration:none">✏ Edit</a>
            {folder_link}
            &nbsp;
            <button class="del-btn" type="button"
              onclick="if(confirm('Delete this project?')){{var f=document.createElement('form');f.method='post';f.action='/delete_project';var i=document.createElement('input');i.type='hidden';i.name='project_id';i.value='{pr['id']}';f.appendChild(i);document.body.appendChild(f);f.submit();}}">🗑</button>
          </td>
        </tr>"""

    if not proj_rows_html:
        proj_rows_html = (
            f'<tr><td colspan="9" style="padding:28px;text-align:center;color:#aaa">'
            f'No projects found. <a href="/data_entry" style="color:{C["dark"]};font-weight:700">Add projects →</a>'
            f'</td></tr>'
        )

    # ── Totals in Lakhs ──
    total_ov_l  = round(total_ov  / 1e5, 2)
    total_itd_l = round(total_itd / 1e5, 2)
    total_col_l = round(total_col / 1e5, 2)
    total_ebf_l = round(1.35 * total_itd_l, 2)

    # Total footer Coll. Perf = (total_col - total_ebf) / total_ebf * 100
    total_perf = (total_col - total_ebf) / total_ebf * 100 if total_ebf else 0
    if total_perf >= 0:
        total_perf_str   = f"+{total_perf:.1f}%"
        total_perf_color = "#166534"
    else:
        total_perf_str   = f"{total_perf:.1f}%"
        total_perf_color = "#991b1b"

    css = f"""<style>
.fin-wrap{{background:#fff;border:1px solid {C['outline']};border-radius:12px;overflow:hidden;margin-bottom:20px}}
.fin-thead th{{padding:10px 14px;font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;
               color:{C['on_var']};background:{C['surface_low']};border-bottom:2px solid {C['outline']};white-space:nowrap}}
.fin-tfoot td{{padding:12px 14px;font-size:11px;font-weight:700;font-family:'Courier New',monospace;
               background:{C['surface_high']};border-top:2px solid {C['outline']}}}
</style>"""

    html  = _html_head("Financial Performance") + css + "<body>"
    html += _sidebar("financials") + _topbar("FINANCIAL PERFORMANCE", username)
    html += f"""
<div class="main"><div class="canvas" style="max-width:1400px">
  {alert_html}

  <!-- Page header -->
  <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:24px">
    <div>
      <h2 style="font-size:28px;font-weight:700;color:{C['dark']}">Financial Performance Report</h2>
      <p style="font-size:14px;color:#6b7280;margin-top:4px">Project-level cost, collection and EBF deviation analysis.</p>
    </div>
    <div style="display:flex;gap:10px">
      <a href="/export_pdf/financials?lob={sel_lob if sel_lob != 'All Lines of Business' else ''}" class="btn-outline">⬇ Export PDF</a>
    </div>
  </div>

  <!-- LOB filter -->
  <form method="post" action="/financials_filter">
    <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:12px 20px;
                display:flex;align-items:center;gap:16px;margin-bottom:24px">
      <span style="font-size:18px">≡</span>
      <span style="font-size:14px;font-weight:600;color:{C['dark']}">Line of Business</span>
      <div style="height:20px;width:1px;background:{C['outline']}"></div>
      <select name="lob" onchange="this.form.submit()"
              style="border:none;background:transparent;font-size:14px;color:{C['dark']};
                     outline:none;cursor:pointer;min-width:240px;font-weight:500">
        {lob_opts}
      </select>
      <span style="margin-left:auto;font-size:18px;color:{C['on_var']}">⌄</span>
    </div>
  </form>

  <!-- Main table card -->
  <div class="fin-wrap">
    <!-- Table header bar -->
    <div style="padding:16px 20px;border-bottom:1px solid {C['outline']};
                display:flex;align-items:center;justify-content:space-between;
                background:{C['surface_low']}">
      <span style="font-size:16px;font-weight:700;color:{C['dark']}">Project Performance Analysis</span>
      <span style="font-size:11px;font-weight:700;color:{C['on_var']};
                   background:{C['surface_high']};border:1px solid {C['outline']};
                   padding:4px 12px;border-radius:20px">
        Showing {n_rows} of {all_projects_count} Projects
      </span>
    </div>

    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse">
        <thead class="fin-thead">
          <tr>
            <th style="text-align:left;padding-left:14px">CO.<br>NO.</th>
            <th style="text-align:left">PROJECT NAME</th>
            <th style="text-align:left">PM<br>NAME</th>
            <th style="text-align:right">ORDER VALUE<br>(₹ LAKHS)</th>
            <th style="text-align:right">ITD<br>(₹ LAKHS)</th>
            <th style="text-align:right">COLLECTION<br>(₹ LAKHS)</th>
            <th style="text-align:center">EBF = 1.35×ITD<br>(₹ LAKHS)</th>
            <th style="text-align:center">COLL. PERF.<br>(COL−EBF)/EBF</th>
            <th style="text-align:center">ACTION</th>
          </tr>
        </thead>
        <tbody>
          {proj_rows_html}
        </tbody>
        <tfoot class="fin-tfoot">
          <tr>
            <td colspan="3" style="font-size:11px;font-weight:700;text-transform:uppercase;
                                   color:{C['dark']};letter-spacing:.05em;font-family:inherit">
              Totals ({sel_lob if sel_lob != "All Lines of Business" else "All LOB"})
            </td>
            <td style="text-align:right;color:{C['dark']}">{total_ov_l:,.2f}</td>
            <td style="text-align:right;color:{C['dark']}">{total_itd_l:,.2f}</td>
            <td style="text-align:right;color:{C['dark']}">{total_col_l:,.2f}</td>
            <td style="text-align:center;color:{C['dark']}">{total_ebf_l:,.2f}</td>
            <td style="text-align:center">
              <span style="font-weight:700;color:{total_perf_color}">{total_perf_str}</span>
              <br><span style="font-size:9px;font-weight:600;color:{C['on_var']}">PORTFOLIO</span>
            </td>
            <td></td>
          </tr>
        </tfoot>
      </table>
    </div>
  </div>

  <!-- EBF Ratio card — FIXED: shows Collection/EBF ratio -->
  <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;
              padding:20px 28px;margin-bottom:20px">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:16px">

      <!-- Left: explanation -->
      <div style="flex:1;min-width:260px">
        <div style="font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;
                    color:{C['on_var']};margin-bottom:6px">
          Collection vs EBF Target (Profit/Loss Indicator)
        </div>
        <div style="font-size:13px;color:#6b7280;line-height:1.6">
          EBF = 1.35 × ITD Cost &nbsp;·&nbsp; Threshold: Collection / EBF ≥ 1.0<br>
          <span style="color:#166534;font-weight:600">≥ 1.0x → Profit</span> &nbsp;|&nbsp;
          <span style="color:#991b1b;font-weight:600">&lt; 1.0x → Loss</span>
          &nbsp;·&nbsp; {ratio_sub}
        </div>

        <!-- Mini breakdown bar -->
        <div style="margin-top:14px">
          <div style="display:flex;justify-content:space-between;margin-bottom:4px">
            <span style="font-size:10px;font-weight:700;text-transform:uppercase;
                         letter-spacing:.05em;color:{C['on_var']}">Collection Coverage of EBF</span>
            <span style="font-size:11px;font-weight:700;color:{'#166534' if ratio>=1 else '#991b1b'}">{min(ratio*100,100):.1f}%</span>
          </div>
          <div style="background:#e5e7eb;border-radius:99px;height:10px;overflow:hidden">
            <div style="height:100%;width:{min(ratio*100,100):.1f}%;
                        background:{'#166534' if ratio>=1 else '#dc2626'};
                        border-radius:99px;transition:width .4s ease"></div>
          </div>
          <div style="display:flex;justify-content:space-between;margin-top:4px">
            <span style="font-size:10px;color:#9ca3af">₹{total_col_l:,.1f}L collected</span>
            <span style="font-size:10px;color:#9ca3af">EBF target: ₹{total_ebf_l:,.1f}L</span>
          </div>
        </div>
      </div>

      <!-- Right: big ratio number -->
      <div style="text-align:center;min-width:160px">
        <div style="font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;
                    color:{C['on_var']};margin-bottom:6px">Collection / EBF</div>
        <div style="font-size:52px;font-weight:800;color:{C['dark']};line-height:1">{ratio_str}</div>
        <span style="display:inline-block;padding:5px 18px;border-radius:6px;font-size:12px;
                     font-weight:700;margin-top:8px;{ratio_badge_cls}">{ratio_lbl}</span>
        <div style="font-size:10px;color:#9ca3af;margin-top:6px">
          Avg Collection Rate: {avg_perf:.1f}% of Order Value
        </div>
      </div>

    </div>
  </div>

  <!-- Legend card -->
  <div style="background:{C['surface_low']};border:1px solid {C['outline']};border-radius:8px;
              padding:12px 20px;margin-bottom:20px;display:flex;gap:20px;flex-wrap:wrap;align-items:center">
    <span style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:{C['on_var']}">
      EBF Badge Key:
    </span>
    <span style="font-size:11px;font-weight:700;color:#166534;background:#f0fdf4;
                 border:1px solid #bbf7d0;padding:3px 10px;border-radius:4px">
      ✓ Collection ≥ EBF — Profit
    </span>
    <span style="font-size:11px;font-weight:700;color:#92400e;background:#fffbeb;
                 border:1px solid #fde68a;padding:3px 10px;border-radius:4px">
      ⚠ Collection 80–99% of EBF — Warning
    </span>
    <span style="font-size:11px;font-weight:700;color:#991b1b;background:#fef2f2;
                 border:1px solid #fecaca;padding:3px 10px;border-radius:4px">
      ✗ Collection &lt; 80% of EBF — Loss
    </span>
    <span style="font-size:10px;color:#9ca3af;margin-left:auto">
      Coll. Perf. = (Collection − EBF) ÷ EBF × 100
    </span>
  </div>

</div></div>

<!-- FAB -->
<button onclick="window.location='/data_entry'" title="Create Project"
        style="position:fixed;bottom:28px;right:28px;width:52px;height:52px;
               background:{C['dark']};color:#fff;border:none;border-radius:50%;
               font-size:26px;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.24)">＋</button>

</body></html>"""
    return html


# ═══════════════════════════════════════════════════════════════
#  PAGE 5 — VARIATIONS
# ═══════════════════════════════════════════════════════════════
def export_variations_pdf_bytes() -> bytes:
    db_vars = fetch_all_variations()
    rows, total_var = [], 0.0
    for v in db_vars:
        amt = _n(v.get("variation_amt")); total_var += amt
        rows.append([v.get("child_co_no",""), v.get("child_name",""), v.get("parent_name",""),
                     f"{_n(v.get('original_value')):,.0f}", f"{amt:,.0f}",
                     f"{_n(v.get('revised_value')):,.0f}", v.get("variation_date","")])
    return _build_table_pdf(
        "Parent Project Details",
        f"{len(db_vars)} variation records",
        ["Child CO.No","Child Name","Parent","Original (₹)","Variation (₹)","Revised (₹)","Date"],
        rows, footer_label="Net Variation", footer_value=total_var)


def variations(alert_html="", prefill=None, username="Admin") -> str:
    stats       = fetch_summary_stats()
    db_projects = fetch_all_projects()
    db_vars     = fetch_all_variations()
    proj_rows   = fetch_projects_for_select()

    net_v   = stats["net_variation"]
    net_fmt = f"₹{net_v/1e6:.1f}M" if net_v >= 1e6 else f"₹{net_v:,.0f}"

    cutoff_iso     = (datetime.now() - timedelta(days=90)).isoformat()
    parents_last_q = sum(
        1 for p in db_projects
        if p.get("project_type") != "child"
        and (p.get("created_at") or "") >= cutoff_iso
    )
    total_ov_db = sum(_n(p.get("order_value")) for p in db_projects)
    avg_children = f"{stats['total_children'] / max(stats['total_parents'], 1):.1f} avg/parent"
    net_var_pct  = (f"{abs(net_v) / max(total_ov_db, 1) * 100:.1f}% of Order Value"
                    if total_ov_db else "—")

    # ── SWAPPED: now uses rich financials-style expand with KPI cards ──
    parent_list_v = [p for p in db_projects if p.get("project_type", "parent") != "child"]
    child_list_v  = [p for p in db_projects if p.get("project_type", "") == "child"]

    hier_rows = ""
    for par in parent_list_v:
        co   = par.get("co_no")           or "—"
        name = par.get("project_name")    or "—"
        pm   = par.get("project_manager") or "—"
        rev_ov = _rev_ov(par)
        ov     = rev_ov
        itd  = _n(par.get("itd_cost"))               # Col Z — Total ITD Cost
        col  = _n(par.get("collection_till_date"))
        ebf  = round(1.35 * itd, 2)                  # Col AA
        perf = (col / rev_ov * 100) if rev_ov else 0 # Col AB
        pid  = par.get("id", "")
        st   = par.get("project_status") or "Active"

        ratio_ebf = ebf / itd if itd else 0
        if ratio_ebf >= 1.30:   ebf_cls = "background:#f0fdf4;color:#166534;border:1px solid #bbf7d0"
        elif ratio_ebf >= 1.10: ebf_cls = "background:#fffbeb;color:#92400e;border:1px solid #fde68a"
        else:                   ebf_cls = "background:#fef2f2;color:#991b1b;border:1px solid #fecaca"

        if st in ("Cancelled","Hold"):            st_cls = "background:#ffdad6;color:#ba1a1a"
        elif st in ("Closed","Execution Closed"): st_cls = f"background:{C['surface_high']};color:{C['on_var']}"
        else:                                     st_cls = "background:#fde68a;color:#92400e"

        kids     = [c for c in child_list_v
                    if (c.get("parent_project") or "").strip().lower() == name.strip().lower()]
        has_kids = len(kids) > 0

        hier_rows += f"""
        <tr class="parent-fin-row" data-pid="{pid}" data-open="0"
            style="cursor:{'pointer' if has_kids else 'default'};"
            {'onclick="openChildModal(\''+str(pid)+'\')"' if has_kids else ''}>
          <td style="padding:14px 16px">
            <div style="display:flex;align-items:center;gap:10px">
              <span class="expand-icon" style="font-size:18px;color:{C['on_var']};
                    transition:transform .2s;display:inline-block;min-width:20px;text-align:center">
                {'›' if has_kids else '—'}
              </span>
              <div>
                <p style="font-weight:700;color:{C['dark']};font-size:14px">{co}: {name}</p>
                <div style="display:flex;align-items:center;gap:8px;margin-top:4px">
                  <span style="padding:2px 8px;border-radius:3px;font-size:10px;
                               font-weight:700;text-transform:uppercase;{st_cls}">{st}</span>
                  <span style="font-size:11px;color:{C['on_var']}">{pm}</span>
                  {f'<span style="font-size:10px;background:{C["dark"]};color:#fff;padding:1px 7px;border-radius:10px;font-weight:700">{len(kids)} sub-project{"s" if len(kids)>1 else ""}</span>' if has_kids else ''}
                </div>
              </div>
            </div>
          </td>
          <td style="text-align:right;font-family:'Courier New',monospace;font-size:13px;padding:14px 16px;font-weight:600">{ov:,.0f}</td>
          <td style="text-align:right;font-family:'Courier New',monospace;font-size:13px;padding:14px 16px;font-weight:600">{itd:,.0f}</td>
          <td style="text-align:right;font-family:'Courier New',monospace;font-size:13px;padding:14px 16px;font-weight:600">{col:,.0f}</td>
          <td style="text-align:center;padding:14px 16px">
            <span style="display:inline-block;padding:3px 10px;border-radius:4px;font-size:12px;font-weight:700;{ebf_cls}">{ebf:,.0f}</span>
          </td>
          <td style="text-align:right;font-weight:700;color:{C['dark']};padding:14px 16px">{perf:.1f}%</td>
        </tr>"""

        if has_kids:
            for ch in kids:
                ch_co   = ch.get("co_no",          "—")
                ch_name = ch.get("project_name",   "—")
                ch_pm   = ch.get("project_manager","—")
                ch_st   = ch.get("project_status", "Active")
                ch_ov   = _rev_ov(ch)
                ch_itd  = _n(ch.get("billed_amount"))
                ch_col  = _n(ch.get("collection_till_date"))
                ch_ebf  = round(1.35 * ch_itd, 2)
                ch_perf = (ch_col / ch_ov * 100) if ch_ov else 0

                ch_budget    = ch_ov
                ch_expense   = _n(ch.get("itd_cost") or ch.get("billed_amount"))
                ch_remaining = ch_budget - ch_expense
                ch_exp_pct   = (ch_expense / ch_budget * 100) if ch_budget else 0

                if ch_exp_pct >= 100:
                    warn_badge = (
                        f'<span style="display:inline-flex;align-items:center;gap:4px;'
                        f'background:#fef2f2;color:#991b1b;border:1px solid #fca5a5;'
                        f'padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700">🚨 CRITICAL — Over Budget</span>'
                    )
                    exp_pct_style   = "color:#991b1b;font-weight:800"
                    bar_color       = "#dc2626"
                    remaining_style = "color:#991b1b;font-weight:700"
                elif ch_exp_pct >= 80:
                    warn_badge = (
                        f'<span style="display:inline-flex;align-items:center;gap:4px;'
                        f'background:#fff7ed;color:#c2410c;border:1px solid #fdba74;'
                        f'padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700">⚠ WARNING — High Spend</span>'
                    )
                    exp_pct_style   = "color:#c2410c;font-weight:700"
                    bar_color       = "#f97316"
                    remaining_style = "color:#c2410c;font-weight:600"
                else:
                    warn_badge      = ""
                    exp_pct_style   = f"color:{C['dark']};font-weight:600"
                    bar_color       = "#166534"
                    remaining_style = "color:#166534;font-weight:600"

                bar_width = min(ch_exp_pct, 100)

                if ch_st in ("Cancelled","Hold"):            ch_st_cls = "background:#ffdad6;color:#ba1a1a"
                elif ch_st in ("Closed","Execution Closed"): ch_st_cls = f"background:{C['surface_high']};color:{C['on_var']}"
                else:                                         ch_st_cls = "background:#e0f2fe;color:#0369a1"

                hier_rows += f"""
                <tr class="child-fin-row child-of-{pid}" style="display:none;background:#f8faff;">
                  <td colspan="6" style="padding:0 0 0 56px;">
                    <div style="border-left:3px solid {C['mid']};margin:6px 16px 6px 0;padding:14px 16px;
                                background:#fff;border-radius:0 8px 8px 0;
                                border:1px solid {C['outline']};border-left:3px solid {C['mid']};">
                      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;">
                        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
                          <span style="font-size:13px;font-weight:700;color:{C['dark']}">{ch_co}: {ch_name}</span>
                          <span style="padding:2px 8px;border-radius:3px;font-size:10px;font-weight:700;text-transform:uppercase;{ch_st_cls}">{ch_st}</span>
                          <span style="font-size:11px;color:{C['on_var']}">PM: {ch_pm}</span>
                          {warn_badge}
                        </div>
                      </div>
                      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:12px;">
                        <div style="background:{C['surface_low']};border-radius:6px;padding:10px 12px;border:1px solid {C['outline']};">
                          <div style="font-size:9px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']};margin-bottom:4px;">Budget (Order Val)</div>
                          <div style="font-size:15px;font-weight:700;color:{C['dark']};">₹{ch_budget:,.0f}</div>
                        </div>
                        <div style="background:{C['surface_low']};border-radius:6px;padding:10px 12px;border:1px solid {C['outline']};">
                          <div style="font-size:9px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']};margin-bottom:4px;">Expense (ITD Cost)</div>
                          <div style="font-size:15px;{exp_pct_style};">₹{ch_expense:,.0f}</div>
                        </div>
                        <div style="background:{C['surface_low']};border-radius:6px;padding:10px 12px;border:1px solid {C['outline']};">
                          <div style="font-size:9px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']};margin-bottom:4px;">Remaining Budget</div>
                          <div style="font-size:15px;{remaining_style};">₹{ch_remaining:,.0f}</div>
                        </div>
                        <div style="background:{C['surface_low']};border-radius:6px;padding:10px 12px;border:1px solid {C['outline']};">
                          <div style="font-size:9px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']};margin-bottom:4px;">Expense %</div>
                          <div style="font-size:15px;{exp_pct_style};">{ch_exp_pct:.1f}%</div>
                        </div>
                      </div>
                      <div>
                        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                          <span style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:{C['on_var']};">Budget Utilization</span>
                          <span style="font-size:11px;font-weight:700;{exp_pct_style};">{ch_exp_pct:.1f}% used</span>
                        </div>
                        <div style="background:#e5e7eb;border-radius:99px;height:8px;overflow:hidden;">
                          <div style="height:100%;width:{bar_width:.1f}%;background:{bar_color};border-radius:99px;transition:width .4s ease;"></div>
                        </div>
                      </div>
                      <div style="display:flex;gap:24px;margin-top:12px;padding-top:10px;border-top:1px solid {C['outline']};flex-wrap:wrap;">
                        <div><span style="font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']};">ITD Billed</span>
                          <span style="font-size:12px;font-weight:600;color:{C['dark']};margin-left:6px;font-family:'Courier New',monospace;">₹{ch_itd:,.0f}</span></div>
                        <div><span style="font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']};">Collection</span>
                          <span style="font-size:12px;font-weight:600;color:{C['dark']};margin-left:6px;font-family:'Courier New',monospace;">₹{ch_col:,.0f}</span></div>
                        <div><span style="font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']};">EBF (1.35×)</span>
                          <span style="font-size:12px;font-weight:600;color:{C['dark']};margin-left:6px;font-family:'Courier New',monospace;">₹{ch_ebf:,.0f}</span></div>
                        <div><span style="font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']};">Coll. Perf.</span>
                          <span style="font-size:12px;font-weight:600;color:{C['dark']};margin-left:6px;">{ch_perf:.1f}%</span></div>
                      </div>
                    </div>
                  </td>
                </tr>"""

    if not hier_rows:
        hier_rows = f'<tr><td colspan="6" style="padding:20px;text-align:center;color:#aaa">No projects. <a href="/data_entry">Add via Data Entry →</a></td></tr>'

    var_rows = ""
    for v in db_vars:
        amt     = _n(v.get("variation_amt"))
        orig    = _n(v.get("original_value"))
        revised = _n(v.get("revised_value"))
        if amt > 0:   amt_style = f"color:{C['error']};font-weight:700"; amt_str = f"+{_fmt_inr(amt)}"
        elif amt < 0: amt_style = "color:#166534;font-weight:700";        amt_str = _fmt_inr(amt)
        else:         amt_style = f"color:{C['on_var']}";                 amt_str = "0"
        var_rows += f"""
        <tr>
          <td>
            <p style="font-weight:700;color:{C['dark']};font-size:13px">{v.get('child_co_no','')}: {v.get('child_name','')}</p>
            <p style="font-size:11px;color:{C['on_var']}">Parent: {v.get('parent_name','')}</p>
          </td>
          <td style="text-align:right;font-family:'Courier New',monospace">{_fmt_inr(orig)}</td>
          <td style="text-align:right;font-family:'Courier New',monospace;{amt_style}">{amt_str}</td>
          <td style="text-align:right;font-family:'Courier New',monospace;font-weight:700">{_fmt_inr(revised)}</td>
          <td style="font-family:'Courier New',monospace;font-size:13px;color:{C['on_var']}">{v.get('variation_date','')}</td>
          <td style="text-align:right">
            <form method="post" action="/delete_variation" style="display:inline">
              <input type="hidden" name="var_id" value="{v['id']}">
              <button class="del-btn">🗑</button>
            </form>
          </td>
        </tr>"""
    if not var_rows:
        var_rows = '<tr><td colspan="6" style="padding:20px;text-align:center;color:#aaa">No variations. Use "+ New Variation".</td></tr>'

    proj_sel_opts = "".join(
        f'<option value="{r[0]}">{r[1] or ""} | {r[2] or ""}</option>'
        for r in proj_rows
    ) or '<option value="">— No projects —</option>'
    # Build JSON for modal
    import json
    modal_data = {}
    for par in parent_list_v:
        pid_key = str(par.get("id", ""))
        kids    = [c for c in child_list_v
                   if (c.get("parent_project") or "").strip().lower()
                   == (par.get("project_name") or "").strip().lower()]
        modal_data[pid_key] = {
            "parent_name": par.get("project_name", "—"),
            "parent_co":   par.get("co_no", "—"),
            "children": [
                {
                    "co_no":        c.get("co_no", "—"),
                    "name":         c.get("project_name", "—"),
                    "pm":           c.get("project_manager", "—"),
                    "status":       c.get("project_status", "—"),
                    "order_value":  _n(c.get("order_value")),
                    "revised_ov":   _n(c.get("revised_order_value")) or (_n(c.get("order_value")) + _n(c.get("variation_amt"))),
                    "variation":    _n(c.get("variation_amt")),
                    "itd_cost":     _n(c.get("itd_cost") or c.get("billed_amount")),
                    "billed":       _n(c.get("billed_amount")),
                    "collection":   _n(c.get("collection_till_date")),
                    "phys_pct":     _n(c.get("physical_completion")),
                    "ebf":          round(1.35 * _n(c.get("billed_amount")), 2),
                    "baseline_eac": _n(c.get("baseline_eac")),
                    "rev_eac":      _n(c.get("rev_latest_eac")),
                    "margin":       _n(c.get("rev_latest_margin")),
                }
                for c in kids
            ]
        }
    modal_data_json = json.dumps(modal_data) 
    html  = _html_head("Parent Project Details") + "<body>"
    html += _sidebar("variations") + _topbar("Parent Project Details", username)
    html += f"""
<style>
.parent-fin-row:hover td{{background:rgba(0,14,36,.03)}}
.var-panel{{position:fixed;inset-y:0;right:0;width:420px;background:#fff;border-left:1px solid {C['outline']};box-shadow:-4px 0 24px rgba(0,0,0,.12);z-index:400;transform:translateX(100%);transition:transform .3s ease;display:flex;flex-direction:column}}
.var-panel.open{{transform:translateX(0)}}
.overlay{{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:399;display:none;backdrop-filter:blur(1px)}}
.overlay.show{{display:block}}
@keyframes fadeSlide{{from{{opacity:0;transform:translateY(-4px)}}to{{opacity:1;transform:translateY(0)}}}}
</style>
<div class="main"><div class="canvas" style="max-width:1360px">
  {alert_html}
  <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:24px">
    <div>
      <h2 style="font-size:28px;font-weight:700;color:{C['dark']}">Parent Project Details</h2>
      <p style="font-size:14px;color:#6b7280;margin-top:4px">Live hierarchy and financial deviation analysis.</p>
    </div>
    <div style="display:flex;gap:10px">
      <a href="/export_pdf/variations" class="btn-outline">⬇ Export PDF</a>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:28px">
    {" ".join(f'<div style="background:{C["surface_low"]};border:1px solid {C["outline"]};border-radius:8px;padding:14px 16px"><div style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C["on_var"]};margin-bottom:8px">{lbl}</div><div style="display:flex;align-items:flex-end;justify-content:space-between"><span style="font-size:26px;font-weight:800;color:{vc}">{vl}</span><span style="font-size:11px;font-weight:600;color:{sc}">{sub}</span></div></div>'
      for lbl,vl,vc,sub,sc in [
        ("Total Parent Projects",  str(stats['total_parents']),  C['dark'], f"+{parents_last_q} Last 90 Days", "#92400e"),
        ("Total Child Projects",   str(stats['total_children']), C['dark'], avg_children,                      C['on_var']),
        ("Net Variation Impact",   net_fmt,                      C['error'], net_var_pct,                      C['error']),
        ("Pending Variations",     f"{stats['pending']:02d}",    C['dark'], "Awaiting Approval",               C['secondary']),
      ])}
  </div>

  <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;overflow:hidden;margin-bottom:24px">
    <div style="padding:12px 18px;border-bottom:1px solid {C['outline']};background:{C['surface_low']}">
      <span style="font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['dark']}">Project Hierarchy &amp; Financial Roll-up</span>
    </div>
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead><tr style="background:{C['surface_low']};border-bottom:1px solid {C['outline']}">
          <th style="text-align:left;padding-left:16px">CO.NO. &amp; Project Name</th>
          <th style="text-align:right">Order Value (₹)</th>
          <th style="text-align:right">ITD Billed (₹)</th>
          <th style="text-align:right">Collection (₹)</th>
          <th style="text-align:center">EBF = 1.35 × ITD</th>
          <th style="text-align:right">Coll. Perf. (%)</th>
        </tr></thead>
        <tbody>{hier_rows}</tbody>
      </table>
    </div>
  </div>

  <!-- Legend for warning colors -->
  <div style="display:flex;gap:12px;align-items:center;margin-top:14px;margin-bottom:4px;flex-wrap:wrap;">
    <span style="font-size:11px;font-weight:700;text-transform:uppercase;color:{C['on_var']};letter-spacing:.04em;">Child Project Status:</span>
    <span style="display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:600;color:#166534;background:#f0fdf4;border:1px solid #bbf7d0;padding:2px 10px;border-radius:4px;">✓ &lt;80% — On Track</span>
    <span style="display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:600;color:#c2410c;background:#fff7ed;border:1px solid #fdba74;padding:2px 10px;border-radius:4px;">⚠ 80–99% — High Spend</span>
    <span style="display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:600;color:#991b1b;background:#fef2f2;border:1px solid #fca5a5;padding:2px 10px;border-radius:4px;">🚨 ≥100% — Over Budget</span>
  </div>
  <div style="text-align:right;font-size:11px;font-style:italic;color:#9ca3af;margin-top:6px">
    Figures in ₹ · Click › on a parent row to expand child project details
  </div>
</div></div>

<div class="var-panel" id="var-panel">
  <div style="padding:18px 20px;border-bottom:1px solid {C['outline']};background:{C['surface_low']};display:flex;align-items:center;justify-content:space-between">
    <span style="font-size:16px;font-weight:700;color:{C['dark']}">New Variation Entry</span>
    <button onclick="closeVarPanel()" style="background:none;border:none;font-size:20px;cursor:pointer;color:{C['on_var']}">✕</button>
  </div>
  <div style="flex:1;padding:22px 20px;overflow-y:auto">
    <form method="post" action="/save_variation" id="var-form">
      {" ".join([
        _field_html("Child CO. No.",   '<input type="text" name="child_co_no" placeholder="e.g. C-24-03">'),
        _field_html("Child Project Name", '<input type="text" name="child_name" placeholder="e.g. Foundation Works">'),
        _field_html("Parent Project",  f'<select name="project_id">{proj_sel_opts}</select>'),
        _field_html("Parent Name (display)", '<input type="text" name="parent_name" placeholder="e.g. Skyway Expressway Phase I">'),
      ])}
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:14px">
        {_field_html("Original Value (₹)", '<input type="number" name="original_value" id="vp-orig" placeholder="0" step="1" oninput="calcRevised()">')}
        {_field_html("Variation Amount (₹)", '<input type="number" name="variation_amt" id="vp-vamt" placeholder="0 or negative" step="1" oninput="calcRevised()">')}
      </div>
      <div style="background:#f0f4ff;border:1px solid #c8d8f8;border-radius:8px;padding:12px 16px;margin:14px 0">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']};margin-bottom:5px">Revised Value (auto)</div>
        <div style="font-size:20px;font-weight:800;color:{C['dark']}" id="vp-revised-val">₹ 0.00</div>
      </div>
      {_field_html("Date of Variation", f'<input type="date" name="variation_date" value="{date.today().isoformat()}">')}
    </form>
  </div>
  <div style="padding:16px 20px;border-top:1px solid {C['outline']};background:{C['surface_low']};display:flex;gap:10px">
    <button onclick="closeVarPanel()" style="flex:1;padding:10px;border:1px solid {C['outline']};border-radius:6px;background:#fff;color:{C['on_var']};font-size:11px;font-weight:700;text-transform:uppercase;cursor:pointer">CANCEL</button>
    <button onclick="document.getElementById('var-form').submit()" style="flex:1;padding:10px;background:{C['dark']};color:#fff;border:none;border-radius:6px;font-size:11px;font-weight:700;text-transform:uppercase;cursor:pointer">SAVE VARIATION</button>
  </div>
</div>
<!-- ── CHILD DETAIL MODAL ── -->
<div id="child-modal-overlay"
     onclick="closeChildModal()"
     style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);
            z-index:500;backdrop-filter:blur(2px);align-items:center;justify-content:center"></div>

<div id="child-modal"
     style="display:none;position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);
            width:min(92vw,900px);max-height:85vh;background:#fff;border-radius:14px;
            box-shadow:0 24px 80px rgba(0,0,0,.22);z-index:501;flex-direction:column;overflow:hidden">

  <!-- header -->
  <div style="padding:18px 24px;border-bottom:1px solid {C['outline']};
              background:{C['surface_low']};display:flex;align-items:center;justify-content:space-between;flex-shrink:0">
    <div>
      <p style="font-size:11px;font-weight:700;text-transform:uppercase;
                letter-spacing:.07em;color:{C['on_var']};margin-bottom:2px">Parent Project</p>
      <h3 id="modal-parent-title"
          style="font-size:18px;font-weight:800;color:{C['dark']}">—</h3>
    </div>
    <button onclick="closeChildModal()"
            style="width:32px;height:32px;border-radius:50%;border:1px solid {C['outline']};
                   background:#fff;font-size:18px;cursor:pointer;color:{C['on_var']};
                   display:flex;align-items:center;justify-content:center">✕</button>
  </div>

  <!-- summary bar -->
  <div id="modal-summary-bar"
       style="display:flex;gap:0;border-bottom:1px solid {C['outline']};flex-shrink:0;overflow-x:auto"></div>

  <!-- child cards scroll area -->
  <div id="modal-body"
       style="overflow-y:auto;padding:20px 24px;flex:1;display:flex;flex-direction:column;gap:16px"></div>

  <!-- footer -->
  <div style="padding:12px 24px;border-top:1px solid {C['outline']};background:{C['surface_low']};
              font-size:11px;color:{C['on_var']};flex-shrink:0">
    All figures in ₹ &nbsp;·&nbsp; Revised Order Value = Original + Variation &nbsp;·&nbsp; EBF = 1.35 × ITD Billed
  </div>
</div>
<div class="overlay" id="var-overlay" onclick="closeVarPanel()"></div>
<script>
function openVarPanel(){{document.getElementById('var-panel').classList.add('open');document.getElementById('var-overlay').classList.add('show');}}
function closeVarPanel(){{document.getElementById('var-panel').classList.remove('open');document.getElementById('var-overlay').classList.remove('show');}}
document.addEventListener('keydown',function(e){{if(e.key==='Escape')closeVarPanel();}});
function calcRevised(){{
  var o=parseFloat(document.getElementById('vp-orig').value)||0;
  var v=parseFloat(document.getElementById('vp-vamt').value)||0;
  document.getElementById('vp-revised-val').textContent='₹ '+(o+v).toLocaleString('en-IN',{{minimumFractionDigits:2}});
}}
var CHILD_DATA = {modal_data_json};

function openChildModal(pid) {{
  var d = CHILD_DATA[pid];
  if (!d || !d.children.length) return;

  document.getElementById('modal-parent-title').textContent = d.parent_co + ': ' + d.parent_name;

  // ── summary bar ──
  var totalOV  = d.children.reduce(function(s,c){{return s+c.revised_ov;}}, 0);
  var totalITD = d.children.reduce(function(s,c){{return s+c.itd_cost;}}, 0);
  var totalCol = d.children.reduce(function(s,c){{return s+c.collection;}}, 0);
  var totalVar = d.children.reduce(function(s,c){{return s+c.variation;}}, 0);
  var bar = document.getElementById('modal-summary-bar');
  bar.innerHTML = [
    ['Sub-projects',   d.children.length,              '#1a2540', '#fff'],
    ['Total Rev. OV',  '₹'+fmt(totalOV),               '#1e3a8a', '#e8f0fe'],
    ['Net Variation',  (totalVar>=0?'+':'')+'₹'+fmt(Math.abs(totalVar)), totalVar<0?'#166534':'#991b1b', totalVar<0?'#f0fdf4':'#fef2f2'],
    ['ITD Billed',     '₹'+fmt(totalITD),              '#1a2540', '#f8faff'],
    ['Collection',     '₹'+fmt(totalCol),              '#1a2540', '#f8faff'],
  ].map(function(item){{
    return '<div style="flex:1;min-width:140px;padding:12px 18px;border-right:1px solid #e5e7eb;background:'+item[3]+'">'
      + '<div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#6b7280;margin-bottom:4px">'+item[0]+'</div>'
      + '<div style="font-size:18px;font-weight:800;color:'+item[2]+'">'+item[1]+'</div>'
      + '</div>';
  }}).join('');

  // ── child cards ──
  var body = document.getElementById('modal-body');
  body.innerHTML = d.children.map(function(c) {{
    var expPct   = c.revised_ov ? (c.itd_cost / c.revised_ov * 100) : 0;
    var barClr   = expPct >= 100 ? '#dc2626' : expPct >= 80 ? '#f97316' : '#166534';
    var stClr    = (c.status==='Cancelled'||c.status==='Hold')
                   ? 'background:#ffdad6;color:#ba1a1a'
                   : (c.status==='Closed'||c.status==='Execution Closed')
                   ? 'background:#e5e7eb;color:#374151'
                   : 'background:#e0f2fe;color:#0369a1';
    var varClr   = c.variation < 0 ? '#166534' : c.variation > 0 ? '#991b1b' : '#6b7280';
    var varSign  = c.variation > 0 ? '+' : '';
    var warnHTML = expPct >= 100
      ? '<span style="background:#fef2f2;color:#991b1b;border:1px solid #fca5a5;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700">🚨 Over Budget</span>'
      : expPct >= 80
      ? '<span style="background:#fff7ed;color:#c2410c;border:1px solid #fdba74;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700">⚠ High Spend</span>'
      : '<span style="background:#f0fdf4;color:#166534;border:1px solid #bbf7d0;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700">✓ On Track</span>';

    return '<div style="border:1px solid #e5e7eb;border-radius:10px;overflow:hidden">'
      // card header
      + '<div style="padding:12px 16px;background:#f8faff;border-bottom:1px solid #e5e7eb;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px">'
      +   '<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">'
      +     '<span style="font-size:14px;font-weight:800;color:#1a2540">'+c.co_no+': '+c.name+'</span>'
      +     '<span style="padding:2px 8px;border-radius:3px;font-size:10px;font-weight:700;text-transform:uppercase;'+stClr+'">'+c.status+'</span>'
      +     '<span style="font-size:11px;color:#6b7280">PM: '+c.pm+'</span>'
      +     warnHTML
      +   '</div>'
      + '</div>'
      // kpi grid
      + '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:0;border-bottom:1px solid #e5e7eb">'
      + [
          ['Order Value',    '₹'+fmt(c.order_value),  '#1a2540'],
          ['Variation',      varSign+'₹'+fmt(Math.abs(c.variation)), varClr],
          ['Revised OV',     '₹'+fmt(c.revised_ov),   '#1e3a8a'],
          ['Baseline EAC',   '₹'+fmt(c.baseline_eac), '#1a2540'],
          ['Rev. EAC',       '₹'+fmt(c.rev_eac),      '#1a2540'],
          ['Margin %',       c.margin.toFixed(1)+'%',  '#166534'],
          ['ITD Cost',       '₹'+fmt(c.itd_cost),     '#1a2540'],
          ['ITD Billed',     '₹'+fmt(c.billed),       '#1a2540'],
          ['Collection',     '₹'+fmt(c.collection),   '#1a2540'],
          ['EBF (1.35×)',    '₹'+fmt(c.ebf),          '#1a2540'],
          ['Phys. Complete', c.phys_pct.toFixed(1)+'%','#1e3a8a'],
        ].map(function(kpi){{
          return '<div style="padding:10px 14px;border-right:1px solid #f0f0f0">'
            + '<div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#6b7280;margin-bottom:3px">'+kpi[0]+'</div>'
            + '<div style="font-size:14px;font-weight:700;color:'+kpi[2]+';font-family:Courier New,monospace">'+kpi[1]+'</div>'
            + '</div>';
        }}).join('')
      + '</div>'
      // progress bar
      + '<div style="padding:10px 16px">'
      +   '<div style="display:flex;justify-content:space-between;margin-bottom:4px">'
      +     '<span style="font-size:10px;font-weight:700;text-transform:uppercase;color:#6b7280">Budget Utilization (ITD vs Revised OV)</span>'
      +     '<span style="font-size:11px;font-weight:700;color:'+barClr+'">'+expPct.toFixed(1)+'%</span>'
      +   '</div>'
      +   '<div style="background:#e5e7eb;border-radius:99px;height:7px;overflow:hidden">'
      +     '<div style="height:100%;width:'+Math.min(expPct,100).toFixed(1)+'%;background:'+barClr+';border-radius:99px;transition:width .4s"></div>'
      +   '</div>'
      + '</div>'
      + '</div>';
  }}).join('');

  // show modal
  document.getElementById('child-modal').style.display        = 'flex';
  document.getElementById('child-modal-overlay').style.display = 'flex';
  document.body.style.overflow = 'hidden';
}}

function closeChildModal() {{
  document.getElementById('child-modal').style.display        = 'none';
  document.getElementById('child-modal-overlay').style.display = 'none';
  document.body.style.overflow = '';
}}

document.addEventListener('keydown', function(e) {{
  if (e.key === 'Escape') {{ closeChildModal(); closeVarPanel(); }}
}});

function fmt(n) {{
  return (n||0).toLocaleString('en-IN', {{minimumFractionDigits:0, maximumFractionDigits:0}});
}}

function toggleFin(row) {{
  // kept as no-op — expand is now handled by modal
}}
</script>
</body></html>"""
    return html


# ═══════════════════════════════════════════════════════════════
#  PAGE 6 — KPI MONITORING
# ═══════════════════════════════════════════════════════════════
import json

KPI_DEFS = [
    ("RF1","EAC < ITD Cost",                  "EAC is lower than Inception-to-Date Cost"),
    ("RF2","ITD Cost > Revenue",               "ITD Cost exceeds up-to-date Revenue Collection"),
    ("RF3","ITD Cost > Order Val",             "Total ITD Cost has surpassed Order Value"),
    ("RF4","OA > Collection",                  "Actual OA Amount exceeds Revenue Collection"),
    ("RF5","GM < EAC Margin",                  "Actual Gross Margin below projected EAC Margin"),
    ("RF6","Phys Comp < Threshold",            "Physical Completion below Minimum Threshold"),
    ("RF7","Billing Coverage Ratio < 1.0",     "Billed Amount is less than Physically-Earned Value (BCR = Billed ÷ (Phys% × Revised OV))"),
    ("RF8","Coll % < EBF",                     "Collection % trailing behind EBF"),
]


def _fmt_amt(v: float) -> str:
    v = _n(v)
    if abs(v) >= 1e7:
        return f"₹{v/1e7:,.2f}Cr"
    return f"₹{v:,.0f}"


def _eval_kpis(proj: dict):
    """Returns (flags, reasons) — flags is a list[bool], reasons is a dict
    mapping RF-code -> human readable explanation, but ONLY for codes that
    are flagged (True). This powers the click-to-explain popup."""
    ov      = _n(proj.get("order_value"))
    rev_ov  = _n(proj.get("revised_order_value")) or (ov + _n(proj.get("variation_amt")))
    itd     = _n(proj.get("itd_cost"))                          # Col Z — Total ITD Cost
    col     = _n(proj.get("collection_till_date"))              # Col S
    bil     = _n(proj.get("billed_amount"))                     # Col Q — Uptodate Bill Amt
    eac     = _n(proj.get("rev_latest_eac") or proj.get("baseline_eac"))   # Col N
    mgn     = _n(proj.get("rev_latest_margin") or proj.get("baseline_margin"))  # Col O (stored as %)
    phys    = _n(proj.get("physical_completion"))               # Col AE (e.g. 0.60 fraction or 60.0 pct)

    # Normalise physical completion to 0-1 fraction
    phys_frac = phys / 100.0 if phys > 1 else phys
    # Keep a % version for RF6 comparisons
    phys_pct  = phys_frac * 100.0

    # Col AA — EBF = ITD Cost × 1.35
    ebf = 1.35 * itd

    # Col AC — Actual Gross Margin % = (Revised OV - ITD Cost) / Revised OV × 100
    actual_gm = ((rev_ov - itd) / rev_ov * 100) if rev_ov else 0

    # Col AD — Financial Completion %
    fin_completion_pct = (itd / eac * 100) if eac else 0

    # RF6 threshold: physical completion should be >= 90% of financial completion
    phys_thresh = fin_completion_pct * 0.9

    # ── RF7: Billing Coverage Ratio = Billed ÷ (Phys% × Revised OV) ─────────
    # Earned Value = the portion of Revised OV that physical work has 'earned'
    earned_value = phys_frac * rev_ov
    bcr = (bil / earned_value) if earned_value else 0.0
    # ─────────────────────────────────────────────────────────────────────────

    flags = [
        bool(eac > 0 and eac < itd),            # RF1: EAC < ITD Cost
        bool(itd > col),                          # RF2: ITD Cost > Collection
        bool(itd > rev_ov),                       # RF3: ITD Cost > Revised OV
        bool(bil > col),                          # RF4: Billed > Collection (OA > Collection)
        bool(actual_gm < mgn),                   # RF5: Actual GM% < EAC Margin%
        bool(phys_pct < phys_thresh),             # RF6: Phys% < 90% of Financial%
        bool(bcr < 1.0),                          # RF7: Billing Coverage Ratio < 1.0
        bool(col < ebf * 0.8),                   # RF8: Collection < 80% of EBF
    ]

    reasons = {}
    if flags[0]:
        reasons["RF1"] = (f"EAC of {_fmt_amt(eac)} is lower than the Inception-to-Date Cost of "
                           f"{_fmt_amt(itd)} — the project is forecast to cost more than its own "
                           f"current estimate at completion.")
    if flags[1]:
        reasons["RF2"] = (f"ITD Cost of {_fmt_amt(itd)} exceeds Revenue Collection of {_fmt_amt(col)} "
                           f"— money spent on the project has outpaced money actually collected.")
    if flags[2]:
        reasons["RF3"] = (f"ITD Cost of {_fmt_amt(itd)} has surpassed the Revised Order Value of "
                           f"{_fmt_amt(rev_ov)} — cumulative cost is now larger than the contract value.")
    if flags[3]:
        reasons["RF4"] = (f"Billed Amount of {_fmt_amt(bil)} exceeds Revenue Collection of {_fmt_amt(col)} "
                           f"— invoices raised are running ahead of cash actually received.")
    if flags[4]:
        reasons["RF5"] = (f"Actual Gross Margin of {actual_gm:.1f}% is below the EAC Margin of {mgn:.1f}% "
                           f"— real profitability is trailing the projected margin.")
    if flags[5]:
        reasons["RF6"] = (f"Physical Completion of {phys_pct:.1f}% is below the minimum threshold of "
                           f"{phys_thresh:.1f}% (90% of Financial Completion, which stands at "
                           f"{fin_completion_pct:.1f}%) — work on the ground is lagging the money spent.")
    if flags[6]:
        reasons["RF7"] = (
            f"BCR is {bcr:.2f} — only {_fmt_amt(bil)} billed against "
            f"{_fmt_amt(earned_value)} earned ({phys_pct:.0f}% physical completion). "
            f"Billing is short by {_fmt_amt(earned_value - bil)}."
        )
    if flags[7]:
        reasons["RF8"] = (f"Collection of {_fmt_amt(col)} is below 80% of EBF "
                           f"({_fmt_amt(ebf * 0.8)}, where EBF = ITD Cost × 1.35 = {_fmt_amt(ebf)}) "
                           f"— collections are falling behind the expected billing-fee benchmark.")

    return flags, reasons


def kpi(alert_html="", username="Admin") -> str:
    projects  = fetch_all_projects()
    evaluated = [(p, *_eval_kpis(p)) for p in projects]   # (proj, flags, reasons)
    total     = len(evaluated) or 1
    red_count = sum(1 for _, f, _ in evaluated if any(f))
    risk_pct  = round(red_count / total * 100)
    unbilled  = sum(max(_n(p.get("billed_amount")) - _n(p.get("collection_till_date")), 0) for p in projects)
    ub_fmt    = f"₹{unbilled/1e7:.2f}Cr" if unbilled >= 1e7 else f"₹{unbilled:,.0f}"
    gauge_col = C['error'] if risk_pct >= 60 else C['secondary'] if risk_pct >= 30 else "#166534"
    gauge_off = round(125.66 * (1 - risk_pct / 100), 2)

    rf_heads = "".join(
        f'<th style="text-align:center;min-width:52px;cursor:help;position:relative">'
        f'{code}'
        f'<div style="display:none;position:absolute;top:calc(100%+4px);left:50%;transform:translateX(-50%);background:#2d3133;color:#eff1f3;padding:6px 10px;border-radius:5px;font-size:10px;width:140px;z-index:40;white-space:normal;line-height:1.4;box-shadow:0 2px 8px rgba(0,0,0,.3)" class="rf-tip">{tip}</div>'
        f'</th>'
        for code, _, tip in KPI_DEFS
    )

    tbody = ""
    reasons_map = {}   # pid -> {code: reason}  — fed into JS for the popup
    for idx, (proj, flags, reasons) in enumerate(evaluated):
        pid  = str(idx)
        co   = proj.get("co_no","—"); name = proj.get("project_name","—")
        lob  = proj.get("business_line") or proj.get("hier_business_line") or "—"
        if reasons:
            reasons_map[pid] = {"name": name, "co": co, "items": reasons}

        cells = "".join(
            (
                f'<td style="text-align:center">'
                f'<span onclick="showFlag(\'{pid}\',\'{KPI_DEFS[i][0]}\')" '
                f'style="font-size:18px;color:{C["error"]};opacity:1;cursor:pointer" '
                f'title="Click to see why">⚑</span></td>'
                if f else
                f'<td style="text-align:center">'
                f'<span style="font-size:16px;color:#16a34a;opacity:.85">✓</span></td>'
            )
            for i, f in enumerate(flags)
        )
        tbody += f"""
        <tr>
          <td style="position:sticky;left:0;background:#fff;padding:12px 14px 12px 20px;min-width:220px;border-bottom:1px solid #f0f2f4">
            <div style="font-weight:700;color:{C['dark']};font-size:14px">{name}</div>
            <div style="font-size:12px;color:{C['on_var']}">{lob}</div>
          </td>
          <td style="font-family:'Courier New',monospace;font-size:13px;color:{C['on_var']};padding:12px 10px;border-bottom:1px solid #f0f2f4">{co}</td>
          {cells}
          <td style="text-align:right;padding:12px 16px;border-bottom:1px solid #f0f2f4">
            <span style="font-size:18px;cursor:pointer;color:{C['on_var']}">⊞</span>
          </td>
        </tr>"""

    if not tbody:
        tbody = f'<tr><td colspan="11" style="padding:24px;text-align:center;color:#aaa">No projects yet. <a href="/data_entry">Add projects →</a></td></tr>'

    half       = len(KPI_DEFS) // 2
    meth_left  = "".join(
        f'<div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:10px">'
        f'<span style="background:{C["dark"]};color:#fff;font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px;flex-shrink:0;font-family:monospace">{code}</span>'
        f'<p style="font-size:13px;color:{C["on_surface"]};line-height:1.5">{desc}</p></div>'
        for code, _, desc in KPI_DEFS[:half]
    )
    meth_right = "".join(
        f'<div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:10px">'
        f'<span style="background:{C["dark"]};color:#fff;font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px;flex-shrink:0;font-family:monospace">{code}</span>'
        f'<p style="font-size:13px;color:{C["on_surface"]};line-height:1.5">{desc}</p></div>'
        for code, _, desc in KPI_DEFS[half:]
    )

    # Lookup of RF code -> short label, used by the popup to title itself
    rf_labels = {code: tip for code, _, tip in KPI_DEFS}

    html  = _html_head("KPI Monitor") + "<body>"
    html += _sidebar("kpi") + _topbar("KPI MONITORING", username)
    html += f"""
<style>
.kpi-bento{{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:24px}}
.kpi-card{{background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:20px 22px}}
.kpi-tbl{{width:100%;border-collapse:collapse;font-size:13px}}
.kpi-tbl thead tr{{background:{C['surface_low']};border-bottom:1px solid {C['outline']}}}
.kpi-tbl th{{padding:11px 10px;font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']};white-space:nowrap}}
th:hover .rf-tip{{display:block!important}}
#flag-modal-backdrop{{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:99;display:none;align-items:center;justify-content:center}}
#flag-modal{{background:#fff;border-radius:12px;max-width:420px;width:90%;padding:22px 24px;box-shadow:0 12px 40px rgba(0,0,0,.25)}}
#flag-modal .fm-code{{display:inline-block;background:{C['error']};color:#fff;font-size:11px;font-weight:700;padding:3px 8px;border-radius:5px;font-family:monospace;margin-bottom:10px}}
#flag-modal h4{{font-size:16px;font-weight:700;color:{C['dark']};margin-bottom:2px}}
#flag-modal .fm-sub{{font-size:12px;color:{C['on_var']};margin-bottom:14px}}
#flag-modal p.fm-reason{{font-size:14px;line-height:1.6;color:{C['on_surface']};background:{C['surface_low']};padding:12px 14px;border-radius:8px;border-left:3px solid {C['error']}}}
#flag-modal .fm-close{{margin-top:18px;width:100%;padding:9px;border-radius:7px;border:1px solid {C['outline']};background:#fff;font-size:13px;font-weight:600;cursor:pointer;color:{C['dark']}}}
</style>
<div class="main"><div class="canvas" style="max-width:1360px">
  {alert_html}
  <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:24px">
    <div>
      <h2 style="font-size:28px;font-weight:700;color:{C['dark']}">KPI Monitor</h2>
      <p style="font-size:14px;color:#6b7280;margin-top:4px">Red flag tracking across {total} live projects from DB</p>
    </div>
    <input style="padding:7px 12px;border:1px solid {C['outline']};border-radius:6px;font-size:13px;width:220px;outline:none;background:{C['surface_low']}"
           placeholder="🔍 Search projects…" type="text" oninput="filterTable(this.value)">
  </div>

  <div class="kpi-bento">
    <div class="kpi-card">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin-bottom:10px">Total Red Flag Projects</p>
      <div style="display:flex;align-items:flex-end;gap:10px">
        <span style="font-size:32px;font-weight:700;line-height:1;color:{C['error']}">{red_count}</span>
        <span style="font-size:13px;color:{C['on_var']};margin-bottom:4px">of {total} projects</span>
      </div>
      <div style="margin-top:14px;padding-top:12px;border-top:1px solid {C['outline']};font-size:13px;font-weight:700;color:{C['error']}">↑ {risk_pct}% flagged</div>
    </div>
    <div class="kpi-card" style="display:flex;flex-direction:column;align-items:center">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin-bottom:10px;align-self:flex-start">Critical Risk Level</p>
      <div style="position:relative">
        <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg" style="width:180px;height:90px">
          <path fill="none" stroke="{C['surface_high']}" stroke-width="12" stroke-linecap="round" d="M 10 50 A 40 40 0 0 1 90 50"/>
          <path fill="none" stroke="{gauge_col}" stroke-width="12" stroke-linecap="round"
                stroke-dasharray="125.66" stroke-dashoffset="125.66"
                d="M 10 50 A 40 40 0 0 1 90 50"
                style="transition:stroke-dashoffset 1s ease-in-out" id="gauge-arc"/>
        </svg>
        <div style="position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;padding-bottom:4px">
          <span style="font-size:24px;font-weight:700;color:{gauge_col};line-height:1">{risk_pct}%</span>
          <span style="font-size:11px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">{("High Risk" if risk_pct>=60 else "Medium" if risk_pct>=30 else "Low Risk")}</span>
        </div>
      </div>
    </div>
    <div class="kpi-card">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin-bottom:10px">Unbilled Exposure</p>
      <div style="display:flex;align-items:flex-end;gap:10px">
        <span style="font-size:32px;font-weight:700;line-height:1;color:{C['dark']}">{ub_fmt}</span>
        <span style="font-size:13px;color:{C['on_var']};margin-bottom:4px">Current Outstandings</span>
      </div>
    </div>
  </div>

  <div style="background:#fff;border:1px solid {C['outline']};border-radius:8px;overflow:hidden;margin-bottom:24px">
    <div style="padding:14px 20px;border-bottom:1px solid {C['outline']};display:flex;align-items:center;justify-content:space-between;background:{C['surface_low']}">
      <h3 style="font-size:20px;font-weight:700;color:{C['dark']}">KPI Red Flag Monitor</h3>
      <a href="/export_pdf/kpi" class="btn-primary" style="font-size:11px;padding:7px 14px">⬇ Export Report</a>
    </div>
    <div style="overflow-x:auto">
      <table class="kpi-tbl">
        <thead><tr>
          <th style="position:sticky;left:0;background:{C['surface_low']};z-index:20;min-width:240px;text-align:left;padding-left:20px">Project Name</th>
          <th style="min-width:90px;text-align:left;padding-left:14px">CO.NO.</th>
          {rf_heads}
          <th style="text-align:right;padding-right:16px">Actions</th>
        </tr></thead>
        <tbody id="kpi-tbody">{tbody}</tbody>
      </table>
    </div>
    <div style="padding:12px 20px;background:{C['surface_low']};border-top:1px solid {C['outline']};display:flex;justify-content:space-between;align-items:center">
      <span style="font-size:12px;color:{C['on_var']}">Showing {total} projects · {red_count} red flags</span>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
    <div style="background:{C['surface_low']};border:1px solid {C['outline']};border-radius:10px;padding:18px">
      <h4 style="font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['dark']};margin-bottom:14px">Red Flag Methodology</h4>
      {meth_left}
    </div>
    <div style="background:{C['surface_low']};border:1px solid {C['outline']};border-radius:10px;padding:18px">
      <h4 style="font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['dark']};margin-bottom:14px;visibility:hidden">Methodology Continued</h4>
      {meth_right}
    </div>
  </div>
</div></div>

<div id="flag-modal-backdrop" onclick="if(event.target===this)closeFlagModal()">
  <div id="flag-modal">
    <span class="fm-code" id="fm-code">RF1</span>
    <h4 id="fm-title">Project Name</h4>
    <div class="fm-sub" id="fm-sub">CO.NO. —</div>
    <p class="fm-reason" id="fm-reason">Reason goes here.</p>
    <button class="fm-close" onclick="closeFlagModal()">Close</button>
  </div>
</div>

<script>
window.addEventListener('load',function(){{
  var arc=document.getElementById('gauge-arc');
  if(arc){{arc.style.strokeDashoffset='125.66';setTimeout(function(){{arc.style.strokeDashoffset='{gauge_off}';}},400);}}
}});
function filterTable(q){{
  q=q.toLowerCase();
  document.querySelectorAll('#kpi-tbody tr').forEach(function(r){{
    r.style.display=r.textContent.toLowerCase().includes(q)?'':'none';
  }});
}}

const KPI_REASONS = {json.dumps(reasons_map, ensure_ascii=False)};
const RF_LABELS = {json.dumps(rf_labels, ensure_ascii=False)};

function showFlag(pid, code){{
  const proj = KPI_REASONS[pid];
  if(!proj || !proj.items[code]) return;
  document.getElementById('fm-code').textContent = code;
  document.getElementById('fm-title').textContent = proj.name;
  document.getElementById('fm-sub').textContent = 'CO.NO. ' + proj.co + ' · ' + (RF_LABELS[code] || '');
  document.getElementById('fm-reason').textContent = proj.items[code];
  document.getElementById('flag-modal-backdrop').style.display = 'flex';
}}
function closeFlagModal(){{
  document.getElementById('flag-modal-backdrop').style.display = 'none';
}}
document.addEventListener('keydown', function(e){{
  if(e.key === 'Escape') closeFlagModal();
}});
</script>
</body></html>"""
    return html
# ═══════════════════════════════════════════════════════════════
""" page 7 REPORTS"""
# ═══════════════════════════════════════════════════════════════

import io
import json as _json
from openpyxl import Workbook as _WB
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Excel export (replaces PDF export) ───────────────────────────────────────

def export_reports_excel_bytes(section: str = "") -> bytes:
    """Export the selected report section as a formatted .xlsx file."""
    projects = fetch_all_projects()

    def _oi_amt(p): return _n(p.get("revised_order_value")) or _n(p.get("order_value"))
    def _lob(p):    return (p.get("hier_business_line") or p.get("business_line") or "—").strip()

    SECTIONS = {
        "oi":  ("Order Intake Report",    "Client",     lambda p: (p.get("project_name","—"), _lob(p), p.get("client_name","—"),                               _oi_amt(p))),
        "col": ("Collection Report",      "Date",       lambda p: (p.get("project_name","—"), _lob(p), (p.get("created_at") or "")[:10] or "—",               _n(p.get("collection_till_date")))),
        "itd": ("ITD Cost Report",        "Order Date", lambda p: (p.get("project_name","—"), _lob(p), (p.get("order_issue_date") or "")[:10] or "—",          _n(p.get("billed_amount")))),
        "ub":  ("Unbilled Amount Report", "Date",       lambda p: (p.get("project_name","—"), _lob(p), (p.get("created_at") or "")[:10] or "—",               max(_n(p.get("order_value")) - _n(p.get("billed_amount")), 0))),
    }
    sec_key = section if section in SECTIONS else "oi"
    title, col3_head, fn = SECTIONS[sec_key]

    wb = _WB()
    ws = wb.active
    ws.title = title[:31]

    # Styles
    hdr_fill   = PatternFill("solid", fgColor="1F2937")
    hdr_font   = Font(bold=True, color="FFFFFF", size=10)
    total_fill = PatternFill("solid", fgColor="F3F4F6")
    total_font = Font(bold=True, size=10)
    amt_font   = Font(name="Courier New", size=10)
    thin       = Side(style="thin", color="E5E7EB")
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)
    center     = Alignment(horizontal="center", vertical="center")
    right      = Alignment(horizontal="right",  vertical="center")
    left       = Alignment(horizontal="left",   vertical="center")

    # Title row
    ws.merge_cells("A1:D1")
    ws["A1"] = title
    ws["A1"].font      = Font(bold=True, size=14, color="1F2937")
    ws["A1"].alignment = left
    ws.row_dimensions[1].height = 28

    # Header row
    headers = ["Project Name", "LOB", col3_head, "Amount (₹)"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = right if c == 4 else left
        cell.border    = border
    ws.row_dimensions[2].height = 20

    # Data rows
    total = 0.0
    for r_idx, p in enumerate(projects, start=3):
        name, lob, detail, amt = fn(p)
        total += amt
        row_data = [name, lob, detail, amt]
        fill = PatternFill("solid", fgColor="FAFAFA") if r_idx % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        for c, val in enumerate(row_data, 1):
            cell = ws.cell(row=r_idx, column=c, value=val)
            cell.fill      = fill
            cell.border    = border
            cell.alignment = right if c == 4 else left
            if c == 4:
                cell.font       = amt_font
                cell.number_format = '₹#,##0.00'
            else:
                cell.font = Font(size=10)

    # Total footer row
    footer_row = len(projects) + 3
    ws.merge_cells(f"A{footer_row}:C{footer_row}")
    ws[f"A{footer_row}"]       = "TOTAL"
    ws[f"A{footer_row}"].font  = total_font
    ws[f"A{footer_row}"].fill  = total_fill
    ws[f"A{footer_row}"].alignment = left
    ws[f"A{footer_row}"].border    = border
    ws[f"D{footer_row}"]       = total
    ws[f"D{footer_row}"].font  = Font(bold=True, name="Courier New", size=11)
    ws[f"D{footer_row}"].fill  = total_fill
    ws[f"D{footer_row}"].alignment = right
    ws[f"D{footer_row}"].border    = border
    ws[f"D{footer_row}"].number_format = '₹#,##0.00'

    # Column widths
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── Main reports page ─────────────────────────────────────────────────────────

def reports(alert_html="", username="Admin", active_section="") -> str:
    """
    Consolidated Reports page — single sidebar entry, no sub-items.
    Filter bar: LOB picker + 4 report-type tabs.
    Export button downloads Excel.
    """
    lob_list = get_lob_list()
    projects = fetch_all_projects()

    def _oi_amt(p):
        rev = _n(p.get("revised_order_value"))
        return rev if rev else _n(p.get("order_value"))

    def _lob(p):
        return (p.get("hier_business_line") or p.get("business_line") or "").strip()

    now        = datetime.now().strftime("%d %b %Y, %I:%M %p IST")
    active_tab = active_section if active_section in ("oi", "col", "itd", "ub") else "oi"

    # Build all project rows as JSON for client-side filtering
    rows_data = []
    for p in projects:
        rows_data.append({
            "name":   p.get("project_name") or "—",
            "lob":    _lob(p),
            "client": p.get("client_name") or "—",
            "date":   (p.get("created_at") or p.get("order_issue_date") or "")[:10] or "—",
            "oi":     _oi_amt(p),
            "col":    _n(p.get("collection_till_date")),
            "itd":    _n(p.get("billed_amount")),
            "ub":     max(_n(p.get("order_value")) - _n(p.get("billed_amount")), 0),
        })

    rows_json     = _json.dumps(rows_data, ensure_ascii=False)
    lob_opts_json = _json.dumps(lob_list or [], ensure_ascii=False)

    html  = _html_head("Consolidated Reports") + "<body>"
    # Pass no active_section sub-key — sidebar shows only "Consolidated Reports"
    html += _sidebar("reports") + _topbar("CONSOLIDATED REPORTS", username)
    html += f"""
<style>
.rep-bar {{
    display:flex; align-items:center; gap:12px; flex-wrap:wrap;
    background:#fff; border:1px solid {C['outline']}; border-radius:10px;
    padding:12px 18px; margin-bottom:20px;
}}
.rep-lob-label {{
    font-size:11px; font-weight:700; letter-spacing:.06em;
    text-transform:uppercase; color:{C['on_var']}; white-space:nowrap;
}}
.rep-lob-select {{
    padding:7px 12px; border:1px solid {C['outline']}; border-radius:6px;
    font-size:12px; background:{C['surface_low']}; color:{C['dark']};
    outline:none; min-width:160px; cursor:pointer;
}}
.rep-divider {{ width:1px; height:28px; background:{C['outline']}; flex-shrink:0; }}
.rep-tabs    {{ display:flex; gap:6px; flex-wrap:wrap; }}
.rep-tab {{
    padding:7px 18px; border-radius:6px; border:1px solid {C['outline']};
    font-size:12px; font-weight:600; cursor:pointer;
    background:{C['surface_low']}; color:{C['on_var']};
    transition:all .15s; white-space:nowrap;
}}
.rep-tab:hover  {{ background:{C['surface_high']}; color:{C['dark']}; }}
.rep-tab.active {{ background:{C['dark']}; color:#fff; border-color:{C['dark']}; }}
.rep-card {{
    background:#fff; border:1px solid {C['outline']};
    border-radius:10px; overflow:hidden; margin-bottom:20px;
}}
.rep-card-head {{
    background:{C['dark']}; padding:12px 18px;
    display:flex; justify-content:space-between; align-items:center;
}}
.rep-tbl {{ width:100%; border-collapse:collapse; font-size:13px; }}
.rep-tbl thead tr {{ background:{C['surface_low']}; border-bottom:1px solid {C['outline']}; }}
.rep-tbl th {{
    padding:9px 14px; font-size:10px; font-weight:700;
    letter-spacing:.06em; text-transform:uppercase; color:{C['on_var']}; text-align:left;
}}
.rep-tbl tbody tr:nth-child(even) {{ background:{C['surface_low']}; }}
.rep-tbl td {{ padding:11px 14px; border-bottom:1px solid #f5f5f5; color:{C['dark']}; }}
.rep-footer {{
    background:{C['surface_high']}; border-top:1px solid {C['outline']};
    padding:11px 18px; display:flex; justify-content:space-between; align-items:center;
}}
#no-rows {{
    display:none; padding:24px; text-align:center; color:#aaa; font-size:13px;
}}
.export-xl {{
    display:inline-flex; align-items:center; gap:6px;
    padding:8px 16px; border-radius:7px; font-size:12px; font-weight:600;
    background:#166534; color:#fff; text-decoration:none; border:none; cursor:pointer;
    transition:background .15s;
}}
.export-xl:hover {{ background:#14532d; }}
</style>

<div class="main"><div class="canvas" style="max-width:1100px">
  {alert_html}

  <!-- Page header -->
  <div style="display:flex;justify-content:space-between;align-items:flex-end;
              margin-bottom:20px;flex-wrap:wrap;gap:12px">
    <div>
      <p style="font-size:11px;font-weight:600;letter-spacing:.07em;text-transform:uppercase;
                color:{C['secondary']};margin:0 0 4px">Corporate Performance Suite</p>
      <h2 style="font-size:26px;font-weight:700;color:{C['dark']};margin:0"
          id="page-title">Consolidated Reports</h2>
    </div>
    <a href="/export_excel/reports?section={active_tab}" class="export-xl" id="export-btn">
      ⬇ Export Excel
    </a>
  </div>

  <!-- Filter bar: LOB + 4 tabs -->
  <div class="rep-bar">
    <span class="rep-lob-label">LOB</span>
    <select class="rep-lob-select" id="lob-filter" onchange="applyFilters()">
      <option value="">All LOB</option>
    </select>
    <div class="rep-divider"></div>
    <div class="rep-tabs">
      <button class="rep-tab" data-tab="oi"  onclick="switchTab('oi')">Order Intake</button>
      <button class="rep-tab" data-tab="col" onclick="switchTab('col')">Collection</button>
      <button class="rep-tab" data-tab="itd" onclick="switchTab('itd')">ITD Cost</button>
      <button class="rep-tab" data-tab="ub"  onclick="switchTab('ub')">Unbilled Amount</button>
    </div>
  </div>

  <!-- Report card -->
  <div class="rep-card">
    <div class="rep-card-head">
      <span style="font-size:12px;font-weight:700;letter-spacing:.05em;
                   text-transform:uppercase;color:#fff" id="card-title">Order Intake Report</span>
      <span style="font-size:12px;color:#adc7f8" id="card-tag">Revised Order Value</span>
    </div>
    <div style="max-height:480px;overflow-y:auto">
      <table class="rep-tbl">
        <thead>
          <tr>
            <th style="min-width:260px">Project Name</th>
            <th style="min-width:140px">LOB</th>
            <th id="col3-head" style="min-width:120px">Client</th>
            <th style="text-align:right;min-width:140px">Amount (₹)</th>
          </tr>
        </thead>
        <tbody id="rep-tbody"></tbody>
      </table>
      <div id="no-rows">No matching projects.</div>
    </div>
    <div class="rep-footer">
      <span style="font-size:11px;font-weight:700;text-transform:uppercase;
                   color:{C['dark']}" id="footer-label">TOTAL ORDER VALUE</span>
      <span style="font-size:14px;font-weight:800;font-family:'Courier New',monospace;
                   color:{C['dark']}" id="footer-val">₹ 0</span>
    </div>
  </div>

  <div style="margin-top:12px;padding:10px 0;border-top:1px solid {C['outline']};
              display:flex;justify-content:space-between;font-size:11px;color:{C['on_var']}">
    <span>© 2025 Monarch Surveyors &amp; Engineering Consultants Ltd.</span>
    <span>🟢 Live DB &nbsp;|&nbsp; {now}</span>
  </div>
</div></div>

<script>
const ALL_ROWS = {rows_json};
const LOB_LIST = {lob_opts_json};

const TAB_META = {{
  oi:  {{ title:"Order Intake Report",    tag:"Revised Order Value", col3:"Client",     footer:"TOTAL ORDER VALUE", key:"oi"  }},
  col: {{ title:"Collection Report",      tag:"Realized Revenue",    col3:"Date",        footer:"TOTAL COLLECTIONS", key:"col" }},
  itd: {{ title:"ITD Cost Report",        tag:"Inception-to-Date",   col3:"Order Date",  footer:"TOTAL ITD COSTS",   key:"itd" }},
  ub:  {{ title:"Unbilled Amount Report", tag:"Outstanding WIP",     col3:"Date",        footer:"TOTAL UNBILLED",    key:"ub"  }},
}};

let currentTab = "{active_tab}";

// Populate LOB dropdown from DB values
const lobSel = document.getElementById('lob-filter');
LOB_LIST.forEach(function(lb) {{
  if (!lb) return;
  const o = document.createElement('option');
  o.value = lb; o.textContent = lb;
  lobSel.appendChild(o);
}});

function fmt(v) {{
  if (Math.abs(v) >= 1e7) return '₹\u200a' + (v/1e7).toFixed(2) + '\u00a0Cr';
  if (Math.abs(v) >= 1e5) return '₹\u200a' + (v/1e5).toFixed(2) + '\u00a0L';
  return '₹\u200a' + v.toLocaleString('en-IN', {{maximumFractionDigits:0}});
}}

function switchTab(tab) {{
  currentTab = tab;
  document.querySelectorAll('.rep-tab').forEach(function(b) {{
    b.classList.toggle('active', b.dataset.tab === tab);
  }});
  const m = TAB_META[tab];
  document.getElementById('card-title').textContent   = m.title;
  document.getElementById('card-tag').textContent     = m.tag;
  document.getElementById('col3-head').textContent    = m.col3;
  document.getElementById('footer-label').textContent = m.footer;
  document.getElementById('page-title').textContent   = m.title;
  document.getElementById('export-btn').href          = '/export_excel/reports?section=' + tab;
  applyFilters();
}}

function applyFilters() {{
  const lob   = lobSel.value;
  const key   = TAB_META[currentTab].key;
  const tbody = document.getElementById('rep-tbody');
  tbody.innerHTML = '';
  let total = 0, count = 0;

  ALL_ROWS.forEach(function(r, i) {{
    if (lob && r.lob !== lob) return;
    const amt  = r[key] || 0;
    const col3 = (currentTab === 'oi') ? r.client : r.date;
    total += amt;
    count++;
    const bg   = (count % 2 === 0) ? '#f9fafb' : '#fff';
    const tr   = document.createElement('tr');
    tr.style.background = bg;
    tr.innerHTML =
      '<td style="padding:11px 14px;font-weight:600;color:#111827">'         + r.name             + '</td>' +
      '<td style="padding:11px 14px;font-size:12px;color:#6b7280">'          + (r.lob  || '—')    + '</td>' +
      '<td style="padding:11px 14px;color:#6b7280">'                          + (col3  || '—')    + '</td>' +
      '<td style="padding:11px 14px;text-align:right;font-family:Courier New,monospace;font-weight:700;color:#111827">' + fmt(amt) + '</td>';
    tbody.appendChild(tr);
  }});

  document.getElementById('no-rows').style.display  = count ? 'none' : 'block';
  document.getElementById('footer-val').textContent = fmt(total);
}}

// Initialise on load
switchTab(currentTab);
</script>
</body></html>"""
    return html

# ═══════════════════════════════════════════════════════════════
#  PAGE 8— BILLING & COLLECTIONS
# ═══════════════════════════════════════════════════════════════
def billing_collection(alert_html="", prefill=None, username="Admin") -> str:
    p     = prefill or {}
    today = date.today().isoformat()

    lob_list  = get_lob_list()
    proj_rows = fetch_projects_for_select()

    sel_pid        = p.get("project_id", "")
    sel_lob        = p.get("lob", "")
    pending_filter = p.get("pending_filter", "all")

    proj_opts = '<option value="">All Projects</option>' + "".join(
        f'<option value="{r[0]}" {"selected" if str(sel_pid)==str(r[0]) else ""}>'
        f'{r[1] or ""} | {r[2] or ""}</option>'
        for r in proj_rows
    )

    def _lob_label(x):
        if isinstance(x, (list, tuple)):
            return str(x[0])
        if isinstance(x, dict):
            return str(x.get("name") or x.get("lob") or next(iter(x.values()), ""))
        return str(x)

    lob_opts = '<option value="">All LOB</option>' + "".join(
        f'<option value="{_lob_label(lob)}" {"selected" if sel_lob==_lob_label(lob) else ""}>{_lob_label(lob)}</option>'
        for lob in lob_list
    )

    # ── Entries table: ALL entries by default, narrowed by filters ──
    entries = fetch_billing_entries_filtered(
        project_id=sel_pid or None,
        lob=sel_lob or None,
        pending_only=(pending_filter == "pending")
    )

    # ── Summary cards reflect ONLY the project(s) actually present in the entries above ──
    entry_pids = {str(e.get("project_id")) for e in entries if e.get("project_id") not in (None, "")}
    all_projects = fetch_all_projects(lob_filter=sel_lob) if sel_lob else fetch_all_projects()
    filt_projects = [pr for pr in all_projects if str(pr.get("id")) in entry_pids]
    if sel_pid:
        filt_projects = [pr for pr in filt_projects if str(pr.get("id")) == str(sel_pid)]

    rev_ov_sum = sum(_rev_ov(pr) for pr in filt_projects)
    itd_sum    = sum(_n(pr.get("itd_cost")) for pr in filt_projects)
    phys_avg   = (sum(_n(pr.get("physical_completion")) for pr in filt_projects) / len(filt_projects)) if filt_projects else 0
    eac_sum    = sum(_n(pr.get("rev_latest_eac") or pr.get("baseline_eac")) for pr in filt_projects)
    fin_pct    = (itd_sum / eac_sum * 100) if eac_sum else 0

    total_billed     = sum(_n(e.get("billed_amt"))     for e in entries)
    total_collection = sum(_n(e.get("collection_amt")) for e in entries)
    pending_count    = sum(
        1 for e in entries
        if not ((e.get("collection_date") or "").strip() and _n(e.get("collection_amt")) > 0)
    )

    rows_html = ""
    for e in entries:
        has_date = bool((e.get("collection_date") or "").strip())
        has_amt  = _n(e.get("collection_amt")) > 0
        is_pending = not (has_date and has_amt)
        proj_label = f'{e.get("co_no","") or "—"}: {e.get("project_name","") or "—"}'

        if is_pending:
            coll_cell = (
                f'<td colspan="2" style="padding:10px 16px">'
                f'<form method="post" action="/record_collection" style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">'
                f'<input type="hidden" name="entry_id" value="{e["id"]}">'
                f'<input type="hidden" name="project_id" value="{sel_pid}">'
                f'<input type="hidden" name="lob" value="{sel_lob}">'
                f'<input type="hidden" name="pending_filter" value="{pending_filter}">'
                f'<input type="date" name="collection_date" value="{e.get("collection_date","") or ""}" required class="bc-inline-inp" style="width:128px">'
                f'<input type="number" name="collection_amt" value="{(_n(e.get("collection_amt")) or "") if _n(e.get("collection_amt")) > 0 else ""}" placeholder="Amount" step="0.01" required class="bc-inline-inp" style="width:100px">'
                f'<button type="submit" class="bc-record-btn">Record</button>'
                f'</form></td>'
            )
        else:
            coll_cell = (
                f'<td style="padding:12px 16px;color:{C["on_var"]}">{e.get("collection_date","")}</td>'
                f'<td style="padding:12px 16px;text-align:right;font-variant-numeric:tabular-nums;font-weight:700;color:{C["dark"]}">'
                f'₹ {_n(e.get("collection_amt")):,.0f}</td>'
            )

        row_cls = "is-pending" if is_pending else ""
        rows_html += f"""
        <tr class="{row_cls}">
          <td style="padding:12px 16px;font-weight:600;color:{C['dark']}">
            {proj_label}
            {'<span class="bc-badge bc-badge-pending">Pending</span>' if is_pending else ''}
          </td>
          <td style="padding:12px 16px;color:{C['on_var']}">{e.get('ra_no','') or '—'}</td>
          <td style="padding:12px 16px;color:{C['on_var']}">{e.get('billing_date','')}</td>
          <td style="padding:12px 16px;text-align:right;font-variant-numeric:tabular-nums;font-weight:700;color:{C['dark']}">₹ {_n(e.get('billed_amt')):,.0f}</td>
          {coll_cell}
          <td style="padding:12px 16px;text-align:right">
            {f'''<form method="post" action="/delete_billing_entry" style="display:inline">
              <input type="hidden" name="entry_id" value="{e['id']}">
              <input type="hidden" name="project_id" value="{sel_pid}">
              <input type="hidden" name="lob" value="{sel_lob}">
              <input type="hidden" name="pending_filter" value="{pending_filter}">
              <button class="del-btn" onclick="return confirm('Delete this entry?')" title="Delete entry">
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 7h16M9 7V4h6v3M6 7l1 14h10l1-14"/></svg>
              </button>
            </form>''' if is_pending else f'''<span class="bc-badge bc-badge-locked">
                <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="5" y="11" width="14" height="9" rx="1.5"/><path d="M8 11V7a4 4 0 1 1 8 0v4"/></svg>
                Locked
              </span>'''}
          </td>
        </tr>"""
    if not rows_html:
        rows_html = (f'<tr><td colspan="6" style="padding:40px;text-align:center;color:#aaa">'
                     f'No billing entries match this filter. Use "+ Add Entry" to log one.</td></tr>')

    proj_opts_panel = "".join(
        f'<option value="{r[0]}" {"selected" if str(sel_pid)==str(r[0]) else ""}>'
        f'{r[1] or ""} | {r[2] or ""}</option>'
        for r in proj_rows
    ) or '<option value="">— No projects yet —</option>'

    css = f"""<style>
.bc-sel-bar{{background:#fff;border:1px solid {C['outline']};border-radius:12px;padding:18px 22px;display:flex;align-items:flex-end;gap:16px;flex-wrap:wrap;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.bc-sel-bar .field{{flex:1;min-width:180px}}
.bc-sel-bar .field label{{font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:{C['on_var']};display:block;margin-bottom:6px}}
.bc-sel-bar .field select{{width:100%;padding:9px 12px;border:1px solid {C['outline']};border-radius:7px;font-size:13px;background:#fff;outline:none;color:{C['dark']};transition:border-color .15s}}
.bc-sel-bar .field select:focus{{border-color:{C['dark']}}}

.bc-stats{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px}}
.bc-stat{{background:#fff;border:1px solid {C['outline']};border-left:3px solid {C['dark']};border-radius:10px;padding:14px 16px;box-shadow:0 1px 2px rgba(0,0,0,.03);transition:box-shadow .15s,transform .15s}}
.bc-stat:hover{{box-shadow:0 6px 16px rgba(0,0,0,.07);transform:translateY(-1px)}}
.bc-stat .lbl{{font-size:9px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin-bottom:8px}}
.bc-stat .val{{font-size:19px;font-weight:800;color:{C['dark']};font-variant-numeric:tabular-nums}}

.bc-card{{background:#fff;border:1px solid {C['outline']};border-radius:12px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.bc-card-hdr{{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;border-bottom:1px solid {C['outline']}}}
.bc-card-hdr h3{{font-size:14px;font-weight:700;color:{C['dark']}}}
.bc-card-hdr .count{{font-size:11px;color:{C['on_var']}}}
.bc-card-hdr .count b{{color:{C['dark']}}}

/* ── Table: single clean header row, no vertical dividers, quiet zebra ── */
.bc-tbl{{width:100%;border-collapse:collapse;font-size:13px;table-layout:fixed}}
.bc-tbl td{{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.bc-tbl td:first-child{{white-space:normal}}
.bc-tbl thead th{{padding:11px 16px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;font-size:10.5px;color:{C['on_var']};text-align:left;background:{C['surface_low']};border-bottom:1.5px solid {C['outline']}}}
.bc-tbl thead th.text-right{{text-align:right}}
.bc-tbl tbody tr{{border-bottom:1px solid {C['outline']};transition:background .12s}}
.bc-tbl tbody tr:last-child{{border-bottom:none}}
.bc-tbl tbody tr:nth-child(even):not(.is-pending){{background:#fbfbfb}}
.bc-tbl tbody tr:hover{{background:{C['surface_low']}}}
.bc-tbl tbody tr.is-pending{{background:#fafaf8;box-shadow:inset 3px 0 0 {C['dark']}}}
.bc-tbl tbody tr.is-pending:hover{{background:#f3f3f0}}
.bc-tbl tfoot tr{{background:{C['surface_high']};border-top:2px solid {C['outline']}}}
.bc-tbl tfoot td{{padding:12px 16px;font-weight:700;color:{C['dark']}}}

.bc-badge{{display:inline-flex;align-items:center;gap:4px;margin-left:8px;padding:3px 9px;font-size:9px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;border-radius:20px;vertical-align:middle}}
.bc-badge-pending{{color:{C['dark']};background:#ececec;border:1px solid {C['outline']}}}
.bc-badge-locked{{color:{C['on_var']};background:#f1f1f1}}
.bc-inline-inp{{height:32px;border:1px solid {C['outline']};border-radius:6px;padding:0 9px;font-size:12px;color:{C['dark']}}}
.bc-record-btn{{height:32px;padding:0 14px;background:{C['dark']};color:#fff;border:none;border-radius:6px;font-size:11px;font-weight:700;cursor:pointer;transition:opacity .15s}}
.bc-record-btn:hover{{opacity:.85}}
.del-btn{{display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;background:none;border:none;border-radius:6px;cursor:pointer;color:{C['on_var']};opacity:.7;transition:opacity .15s,background .15s,color .15s}}
.del-btn:hover{{opacity:1;background:#f1f1f1;color:{C['dark']}}}

.side-panel{{position:fixed;inset-y:0;right:0;width:400px;background:#fff;border-left:1px solid {C['outline']};box-shadow:-4px 0 32px rgba(0,0,0,.10);z-index:300;transform:translateX(100%);transition:transform .3s ease;display:flex;flex-direction:column}}
.side-panel.open{{transform:translateX(0)}}
.sp-hdr{{padding:20px 24px;border-bottom:1px solid {C['outline']};background:{C['surface_low']};display:flex;align-items:center;justify-content:space-between}}
.sp-body{{flex:1;padding:24px;overflow-y:auto}}
.sp-section-lbl{{font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin:20px 0 10px;padding-top:14px;border-top:1px solid {C['outline']}}}
.sp-section-lbl:first-child{{margin-top:0;padding-top:0;border-top:none}}
.sp-row{{display:flex;gap:10px}}
.sp-row .sp-fld{{flex:1}}
.sp-fld{{margin-bottom:14px}}
.sp-fld label{{font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:{C['on_var']};display:block;margin-bottom:6px}}
.sp-fld input,.sp-fld select{{width:100%;padding:9px 12px;border:1px solid {C['outline']};border-radius:7px;font-size:13px;background:#fff;outline:none;color:{C['dark']};transition:border-color .15s}}
.sp-fld input:focus,.sp-fld select:focus{{border-color:{C['dark']}}}
.sp-foot{{padding:16px 24px;border-top:1px solid {C['outline']};background:{C['surface_low']};display:flex;gap:10px}}
.overlay{{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:299;display:none}}
.overlay.show{{display:block}}
</style>"""

    js = """<script>
function openBCPanel(){
  document.getElementById('bc-side-panel').classList.add('open');
  document.getElementById('bc-overlay').classList.add('show');
}
function closeBCPanel(){
  document.getElementById('bc-side-panel').classList.remove('open');
  document.getElementById('bc-overlay').classList.remove('show');
}
document.addEventListener('keydown',function(e){if(e.key==='Escape')closeBCPanel();});
</script>"""

    html  = _html_head("Billing & Collection") + css + "<body>"
    html += _sidebar("billing_collection") + _topbar("BILLING & COLLECTION DETAILS", username)
    html += f"""
<div class="main"><div class="canvas" style="max-width:1280px">

  <div style="margin-bottom:20px">
    <h2 style="font-size:24px;font-weight:800;color:{C['dark']};letter-spacing:-.01em">Billing &amp; Collection Details</h2>
    <p style="font-size:13px;color:#6b7280;margin-top:4px">Track RA-wise billing and the collections received across all projects.</p>
    <div style="width:42px;height:3px;background:{C['dark']};border-radius:2px;margin-top:12px;opacity:.85"></div>
  </div>

  {alert_html}

  <!-- Filter bar -->
  <form method="post" action="/billing_collection_retrieve" style="margin-bottom:20px" id="bc-filter-form">
    <div class="bc-sel-bar">
      <div class="field">
        <label>Project Name</label>
        <select name="project_id" onchange="this.form.submit()">{proj_opts}</select>
      </div>
      <div class="field">
        <label>LOB Selection Filter</label>
        <select name="lob" onchange="this.form.submit()">{lob_opts}</select>
      </div>
      <div class="field">
        <label>Pending Filter</label>
        <select name="pending_filter" onchange="this.form.submit()">
          <option value="all" {"selected" if pending_filter=="all" else ""}>All</option>
          <option value="pending" {"selected" if pending_filter=="pending" else ""}>Pending only</option>
        </select>
      </div>
    </div>
  </form>

  <!-- Summary strip -->
  <div class="bc-stats">
    <div class="bc-stat"><div class="lbl">Order Value</div><div class="val">₹ {rev_ov_sum:,.0f}</div></div>
    <div class="bc-stat"><div class="lbl">ITD</div><div class="val">₹ {itd_sum:,.0f}</div></div>
    <div class="bc-stat"><div class="lbl">Phy %</div><div class="val">{phys_avg:.1f}%</div></div>
    <div class="bc-stat"><div class="lbl">Financial %</div><div class="val">{fin_pct:.1f}%</div></div>
  </div>

  <!-- Entries table -->
  <div class="bc-card">
    <div class="bc-card-hdr">
      <h3>Billing Entries</h3>
      <div style="display:flex;align-items:center;gap:16px">
        <span class="count">{len(entries)} entries{f' · <b>{pending_count} pending</b>' if pending_count else ''}</span>
        <button type="button" class="btn-primary" onclick="openBCPanel()">+ Add Entry</button>
      </div>
    </div>
    <div style="overflow-x:auto">
      <table class="bc-tbl">
        <colgroup>
          <col style="width:25%">
          <col style="width:9%">
          <col style="width:14%">
          <col style="width:13%">
          <col style="width:15%">
          <col style="width:14%">
          <col style="width:10%">
        </colgroup>
        <thead>
          <tr>
            <th>Project</th>
            <th>RA No.</th>
            <th>Date of Billing</th>
            <th class="text-right">Billed Amt</th>
            <th>Date of Collection</th>
            <th class="text-right">Collection Amt</th>
            <th class="text-right">Action</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
        <tfoot>
          <tr>
            <td colspan="3">Total</td>
            <td style="text-align:right;font-variant-numeric:tabular-nums">₹ {total_billed:,.0f}</td>
            <td></td>
            <td style="text-align:right;font-variant-numeric:tabular-nums">₹ {total_collection:,.0f}</td>
            <td></td>
          </tr>
        </tfoot>
      </table>
    </div>
  </div>

</div></div>

<!-- Add Entry side panel -->
<div class="side-panel" id="bc-side-panel">
  <div class="sp-hdr">
    <span style="font-size:16px;font-weight:700;color:{C['dark']}">Add Billing Entry</span>
    <button onclick="closeBCPanel()" type="button"
            style="background:none;border:none;font-size:20px;cursor:pointer;color:{C['on_var']}">✕</button>
  </div>
  <div class="sp-body">
    <form method="post" action="/save_billing_entry" id="bc-entry-form">
      <input type="hidden" name="lob" value="{sel_lob}">
      <input type="hidden" name="pending_filter" value="{pending_filter}">

      <div class="sp-fld">
        <label>Project</label>
        <select name="project_id" required>
          <option value="">— Select project —</option>
          {proj_opts_panel}
        </select>
      </div>

      <div class="sp-section-lbl">Billing</div>
      <div class="sp-fld"><label>RA No.</label><input type="text" name="ra_no" placeholder="e.g. RA-05"></div>
      <div class="sp-row">
        <div class="sp-fld"><label>Date of Billing</label><input type="date" name="billing_date" value="{today}"></div>
        <div class="sp-fld"><label>Billed Amount (₹)</label><input type="number" name="billed_amt" step="0.01" placeholder="0.00"></div>
      </div>

      <div class="sp-section-lbl">Collection (optional)</div>
      <div class="sp-row">
        <div class="sp-fld"><label>Date of Collection</label><input type="date" name="collection_date"></div>
        <div class="sp-fld"><label>Collection Amount (₹)</label><input type="number" name="collection_amt" step="0.01" placeholder="0.00"></div>
      </div>
    </form>
  </div>
  <div class="sp-foot">
    <button onclick="closeBCPanel()" type="button"
            style="flex:1;padding:11px;border:1px solid {C['outline']};border-radius:7px;background:#fff;color:{C['on_var']};font-size:11px;font-weight:700;text-transform:uppercase;cursor:pointer">Cancel</button>
    <button onclick="document.getElementById('bc-entry-form').submit()" type="button"
            style="flex:1;padding:11px;background:{C['dark']};color:#fff;border:none;border-radius:7px;font-size:11px;font-weight:700;text-transform:uppercase;cursor:pointer">Save Entry</button>
  </div>
</div>
<div class="overlay" id="bc-overlay" onclick="closeBCPanel()"></div>

{js}
</body></html>"""
    return html
# ═══════════════════════════════════════════════════════════════
#  HTTP HANDLER
# ═══════════════════════════════════════════════════════════════

def _parse_body(handler) -> dict:
    n   = int(handler.headers.get("Content-Length", 0))
    raw = handler.rfile.read(n).decode("utf-8", errors="replace")
    return {k: v[0] for k, v in parse_qs(raw, keep_blank_values=True).items()}

def _drain_body(handler):
    try:
        n = int(handler.headers.get("Content-Length", 0))
        if n > 0: handler.rfile.read(n)
    except Exception:
        pass

def _html_resp(handler, body: str, code=200):
    b = body.encode("utf-8")
    handler.send_response(code)
    handler.send_header("Content-Type",   "text/html; charset=utf-8")
    handler.send_header("Content-Length", str(len(b)))
    handler.end_headers()
    handler.wfile.write(b)

def _redirect(handler, loc="/"):
    handler.send_response(302)
    handler.send_header("Location", loc)
    handler.end_headers()

def _file_resp(handler, data: bytes, ctype: str, fname: str):
    handler.send_response(200)
    handler.send_header("Content-Type",        ctype)
    handler.send_header("Content-Disposition", f'attachment; filename="{fname}"')
    handler.send_header("Content-Length",      str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)

def _process_upload(tmp_path: str, suffix: str) -> str:
    """Parse the Power BI template Excel (header on row 3) and bulk-insert projects."""
    try:
        import openpyxl, math

        wb = openpyxl.load_workbook(tmp_path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))

        # Find header row (first row containing 'Co.No')
        header_row_idx = None
        for i, row in enumerate(rows):
            if any(str(c).strip().lower() in ("co.no", "co no", "co.no.")
                   for c in row if c is not None):
                header_row_idx = i
                break

        if header_row_idx is None:
            return '<div class="alert err">❌ Could not find header row in uploaded file.</div>'

        headers = [str(c).strip() if c is not None else "" for c in rows[header_row_idx]]

        def _col(*candidates):
            """Return column index for first matching header name (case-insensitive)."""
            for name in candidates:
                for i, h in enumerate(headers):
                    if h.lower() == name.lower():
                        return i
            return None

        col_map = {
            "co_no":                        _col("Co.No", "Co No", "CO.NO"),
            "project_name":                 _col("Project Name"),
            "hier_business_line":           _col("LOB"),
            "project_manager":              _col("PM"),
            "client_name":                  _col("Client Name"),
            "order_number":                 _col("Order Number"),
            "order_value":                  _col("Order Value/OA Amt", "Order Value"),
            "variation_amt":                _col("Variation"),
            "revised_order_value":          _col("Revised Order Value"),
            "parent_project":               _col("Parent Project"),
            "baseline_eac":                 _col("Baseline EAC"),
            "baseline_margin":              _col("Baseline Margin"),
            "rev_latest_eac":               _col("Latest EAC"),
            "rev_latest_margin":            _col("Latest Margin"),
            "project_status":               _col("Project Status"),
            "billed_amount":                _col("Uptodate Bill Amt"),
            "last_billed_date":             _col("Last Billed Date"),
            "collection_till_date":         _col("Collection till date"),
            "itd_cost":                     _col("Total ITD"),
            "physical_completion":          _col("Physical % of Completion"),
            "bill_submission_pending_date": _col("Bill submission date but not yet collected"),
        }

        exp_col_map = {
            "erp_cogs":     _col("Total COGS"),
            "erp_djcs":     _col("Total DJC"),
            "erp_equip":    _col("Total Equipment"),
            "erp_vehicles": _col("Total Vehicle"),
            "erp_manpower": _col("Total Man Power"),
            "erp_others":   _col("Total Others"),
        }

        def _safe_num(v):
            """Convert any cell value to float. Handles bool, None, formula strings."""
            if v is None:
                return 0.0
            if isinstance(v, bool):       # Excel TRUE/FALSE cached results
                return 1.0 if v else 0.0
            if isinstance(v, (int, float)):
                return 0.0 if math.isnan(float(v)) else float(v)
            s = str(v).strip()
            if s.startswith("=") or not s:  # formula string with no cached value
                return 0.0
            try:
                return float(s.replace(",", ""))
            except ValueError:
                return 0.0

        def _safe_text(v):
            """Convert any cell value to clean string. Never returns a bool."""
            if v is None:
                return ""
            if isinstance(v, bool):
                return ""
            s = str(v).strip()
            if s.startswith("="):          # formula with no cached value
                return ""
            return s

        def _safe_date(v):
            """Convert date/datetime/string to ISO date string."""
            if v is None:
                return ""
            if isinstance(v, bool):
                return ""
            if hasattr(v, "strftime"):
                return v.strftime("%Y-%m-%d")
            s = str(v).strip()
            return "" if s.startswith("=") else s

        def _get(row, key):
            idx = col_map.get(key)
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        inserted = 0
        skipped  = 0
        errors   = []

        for row_num, row in enumerate(rows[header_row_idx + 1:], start=header_row_idx + 2):
            # Skip fully empty rows
            if all(c is None for c in row):
                continue

            co_no = _safe_text(_get(row, "co_no"))
            if not co_no:
                skipped += 1
                continue

            ov      = _safe_num(_get(row, "order_value"))
            var_amt = _safe_num(_get(row, "variation_amt"))

            # Revised order value: use Excel value if numeric, else calculate
            rev_raw = _get(row, "revised_order_value")
            rev_ov  = _safe_num(rev_raw) if (rev_raw is not None and not isinstance(rev_raw, bool) and not str(rev_raw).startswith("=")) else ov + var_amt

            # Physical completion: Excel stores as decimal (0.6 = 60%)
            phys_raw = _safe_num(_get(row, "physical_completion"))
            phys = phys_raw * 100.0 if 0 < phys_raw <= 1.0 else phys_raw

            # Parent project: integer 0 means "no parent" → treat as empty
            parent_raw = _get(row, "parent_project")
            if isinstance(parent_raw, bool):
                parent_val = ""
            elif isinstance(parent_raw, (int, float)) and not isinstance(parent_raw, bool):
                parent_val = "" if parent_raw == 0 else str(int(parent_raw))
            else:
                parent_val = _safe_text(parent_raw)

            # Baseline/latest margin: stored as decimal in Excel (0.30 = 30%)
            b_margin  = _safe_num(_get(row, "baseline_margin"))
            rl_margin = _safe_num(_get(row, "rev_latest_margin"))
            # Convert to percentage if stored as decimal
            if 0 < abs(b_margin)  <= 1.0: b_margin  *= 100.0
            if 0 < abs(rl_margin) <= 1.0: rl_margin *= 100.0

            fields = {
                "co_no":                        co_no,
                "project_name":                 _safe_text(_get(row, "project_name")),
                "hier_business_line":           _safe_text(_get(row, "hier_business_line")),
                "business_line":                _safe_text(_get(row, "hier_business_line")),
                "project_manager":              _safe_text(_get(row, "project_manager")),
                "client_name":                  _safe_text(_get(row, "client_name")),
                "order_number":                 _safe_text(_get(row, "order_number")),
                "order_issue_date":             "",
                "order_value":                  ov,
                "variation_amt":                var_amt,
                "revised_order_value":          rev_ov,
                "parent_project":               parent_val,
                "child_project":                "",
                "baseline_eac":                 _safe_num(_get(row, "baseline_eac")),
                "baseline_margin":              b_margin,
                "rev_latest_eac":               _safe_num(_get(row, "rev_latest_eac")),
                "rev_latest_margin":            rl_margin,
                "project_status":               _safe_text(_get(row, "project_status")) or "Active",
                "billed_amount":                _safe_num(_get(row, "billed_amount")),
                "last_billed_date":             _safe_date(_get(row, "last_billed_date")),
                "collection_till_date":         _safe_num(_get(row, "collection_till_date")),
                "itd_cost":                     _safe_num(_get(row, "itd_cost")),
                "physical_completion":          phys,
                "bill_submission_pending_date": _safe_date(_get(row, "bill_submission_pending_date")),
                "project_type":                 "parent",  # will be corrected by save_project logic
            }

            # Register new ref values
            for ref_table, ref_col, ref_val in [
                ("ref_project_managers", "name", fields["project_manager"]),
                ("ref_clients",          "name", fields["client_name"]),
                ("ref_lob",              "name", fields["hier_business_line"]),
            ]:
                if ref_val:
                    add_ref_value(ref_table, ref_col, ref_val)

            try:
              new_pid = save_project(fields)

              # ── DIRECT FIX: push Total ITD from Excel into projects.itd_cost ──
              itd_direct = _safe_num(_get(row, "itd_cost"))   # reads "Total ITD" column
              if itd_direct and new_pid:
                  with ENGINE.begin() as conn_fix:
                      conn_fix.execute(
                          text("UPDATE projects SET itd_cost = :v WHERE id = :id"),
                          {"v": itd_direct, "id": new_pid}
                      )

              # Build expenditure record from ITD cost breakdown columns
              exp_fields = {
                  "project_id":      new_pid,
                  "lob":             fields["hier_business_line"],
                  "data_as_of":      "",
                  "erp_update_date": "",
                  "itd_cogs":    0.0, "itd_djcs":  0.0, "itd_equip":    0.0,
                  "itd_manpower":0.0, "itd_vehicles":0.0,"itd_others":  0.0,
                  "erp_cogs":    0.0, "erp_djcs":  0.0, "erp_equip":    0.0,
                  "erp_manpower":0.0, "erp_vehicles":0.0,"erp_others":  0.0,
              }
              for erp_key, idx in exp_col_map.items():
                  val = _safe_num(row[idx] if idx is not None and idx < len(row) else None)
                  itd_key = erp_key.replace("erp_", "itd_")
                  exp_fields[erp_key] = val
                  exp_fields[itd_key] = val

              # ── If all breakdown cols are zero, put Total ITD into erp_others ──
              total_breakdown = sum(exp_fields[k] for k in exp_fields if k.startswith("erp_"))
              if total_breakdown == 0 and itd_direct > 0:
                  exp_fields["erp_others"] = itd_direct
                  exp_fields["itd_others"] = itd_direct

              save_expenditure(exp_fields)
              inserted += 1

            except Exception as e:
                import traceback
                errors.append(f"Row {row_num} ({co_no}): {e}")

        msg = f"✅ Upload complete — {inserted} projects imported, {skipped} blank rows skipped."
        if errors:
            err_detail = "<br>".join(errors[:8])
            return (f'<div class="alert ok">{msg}</div>'
                    f'<div class="alert err">⚠ Some rows had errors:<br>{err_detail}</div>')
        return f'<div class="alert ok">{msg}</div>'

    except Exception as e:
        import traceback
        return (f'<div class="alert err">❌ Upload failed: {e}<br>'
                f'<pre style="font-size:11px">{traceback.format_exc()}</pre></div>')


class MonarchHandler(SimpleHTTPRequestHandler):
    def log_message(self, fmt, *args): pass

    def do_GET(self):
        path = urlparse(self.path).path.rstrip("/") or "/"
        try:
            if path == "/login":
                _html_resp(self, login_page()); return
            if path == "/reset_password":
                _html_resp(self, reset_password_page()); return
            if path == "/logout":
                token = _get_session_token(self)
                if token: destroy_session(token)
                self.send_response(302)
                self.send_header("Set-Cookie",
                    f"{SESSION_COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax")
                self.send_header("Location", "/login")
                self.end_headers(); return
            if path == "/logo.jpg":
                try:
                    with open(LOGO_PATH, "rb") as f:
                        data = f.read()
                    self.send_response(200)
                    self.send_header("Content-Type", "image/jpeg")
                    self.send_header("Content-Length", str(len(data)))
                    self.send_header("Cache-Control", "public, max-age=86400")
                    self.end_headers()
                    self.wfile.write(data)
                except FileNotFoundError:
                    self.send_response(404); self.end_headers()
                return

            user = get_session_user(self)
            if not user: _redirect(self, "/login"); return

            if   path in ("/","/dashboard","/index"): _html_resp(self, dashboard(user))
            elif path == "/data_entry":   _html_resp(self, project_data_entry(username=user))
            elif path == "/search":
                q = parse_qs(urlparse(self.path).query).get("q", [""])[0]
                _html_resp(self, search_results_page(q, username=user))
            elif path == "/expenditures":
                qs = parse_qs(urlparse(self.path).query)
                prefill = {k: v[0] for k, v in qs.items()} if qs else None
                _html_resp(self, expenditures(prefill=prefill, username=user))
            elif path == "/financials":   _html_resp(self, financials(username=user))
            elif path == "/variations":   _html_resp(self, variations(username=user))
            elif path == "/reports":
                section = parse_qs(urlparse(self.path).query).get("section", [""])[0]
                _html_resp(self, reports(username=user, active_section=section))
            elif path == "/kpi":          _html_resp(self, kpi(username=user))
            elif path == "/billing_collection":
                qs = parse_qs(urlparse(self.path).query)
                prefill = {k: v[0] for k, v in qs.items()} if qs else None
                _html_resp(self, billing_collection(prefill=prefill, username=user))
            elif path == "/edit_project":
                pid  = int(parse_qs(urlparse(self.path).query).get("id",["0"])[0])
                proj = fetch_project_by_id(pid)
                if proj:
                    _html_resp(self, project_data_entry(
                        alert_html=f'<div class="alert ok">✏ Editing project #{pid} — resubmit to save changes.</div>',
                        prefill=proj, username=user))
                else:
                    _redirect(self, "/data_entry")

            elif path == "/export_excel": _file_resp(self, export_excel_bytes(),  XLSX_MIME, "monarch_projects.xlsx")
            elif path == "/export_csv":   _file_resp(self, export_csv_bytes(),    "text/csv","monarch_projects.csv")
            elif path == "/export_pdf":   _file_resp(self, export_pdf_bytes(),    "application/pdf","monarch_projects.pdf")
            elif path == "/export/csv":   _file_resp(self, export_lob_csv(),      "text/csv","lob_summary.csv")
            elif path == "/export/xlsx":  _file_resp(self, export_lob_xlsx(),     XLSX_MIME, "lob_summary.xlsx")
            elif path == "/export/pdf":   _file_resp(self, export_lob_pdf(),      "application/pdf","lob_summary.pdf")
            elif path == "/export_pdf/kpi":         _file_resp(self, export_pdf_bytes(),        "application/pdf", "kpi_report.pdf")
            elif path == "/export_pdf/dashboard":   _file_resp(self, export_dashboard_pdf_bytes(), "application/pdf", "dashboard_report.pdf")
            elif path == "/export_pdf/data_entry":  _file_resp(self, export_data_entry_pdf_bytes(), "application/pdf", "projects_report.pdf")
            elif path == "/export_pdf/financials":
                lob = parse_qs(urlparse(self.path).query).get("lob", [""])[0]
                _file_resp(self, export_financials_pdf_bytes(lob), "application/pdf", "financials_report.pdf")
            elif path == "/export_pdf/variations":  _file_resp(self, export_variations_pdf_bytes(), "application/pdf", "variations_report.pdf")
            elif path == "/export_excel/reports":
                sec = parse_qs(urlparse(self.path).query).get("section", ["oi"])[0]
                fname = {"oi":"order_intake_report.xlsx","col":"collection_report.xlsx","itd":"itd_cost_report.xlsx","ub":"unbilled_amount_report.xlsx"}.get(sec, "consolidated_report.xlsx")
                _file_resp(self, export_reports_excel_bytes(sec), XLSX_MIME, fname)
            elif path == "/get_parent_data":
                import json as _json
                qs      = parse_qs(urlparse(self.path).query)
                pname   = qs.get("parent_name", [""])[0]
                excl_id = int(qs.get("exclude_id", ["0"])[0])
                data    = fetch_parent_with_children_totals(pname, excl_id) if pname else {}
                payload = _json.dumps(data).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type",   "application/json")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            
            else:
              _html_resp(self,
                    _html_head("404") + "<body>" + _sidebar() + _topbar(username=user) +
                    f'<div class="main"><div class="canvas" style="padding-top:60px">'
                    f'<p style="font-size:20px;color:{C["error"]}">404 — Not found: <code>{path}</code></p>'
                    f'<p><a href="/">← Dashboard</a></p></div></div></body></html>', code=404)

        except Exception:
            _html_resp(self,
                f'<html><body><pre style="padding:32px;color:red">{traceback.format_exc()}</pre></body></html>',
                code=500)

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            if path == "/login":
                fields   = _parse_body(self)
                username = fields.get("username","").strip()
                password = fields.get("password","")
                next_url = fields.get("next","/dashboard")
                user_obj = verify_user(username, password)
                if not user_obj:
                    _html_resp(self, login_page(
                        alert_html='<div class="alert err">⚠ Incorrect username or password.</div>',
                        next_path=next_url))
                    return
                token = create_session(user_obj)
                self.send_response(302)
                self.send_header("Set-Cookie",
                    f"{SESSION_COOKIE}={token}; Path=/; HttpOnly; SameSite=Lax")
                self.send_header("Location", next_url or "/dashboard")
                self.end_headers(); return

            if path == "/reset_password":
                fields   = _parse_body(self)
                username = fields.get("username","").strip()
                new_pw   = fields.get("new_password","")
                conf_pw  = fields.get("confirm_password","")

                if not username or not new_pw:
                    _html_resp(self, reset_password_page(
                        alert_html='<div class="alert-err">⚠ Please enter your username and a new password.</div>',
                        prefill_username=username))
                    return
                if new_pw != conf_pw:
                    _html_resp(self, reset_password_page(
                        alert_html='<div class="alert-err">⚠ Passwords do not match. Try again.</div>',
                        prefill_username=username))
                    return
                if len(new_pw) < 6:
                    _html_resp(self, reset_password_page(
                        alert_html='<div class="alert-err">⚠ Password must be at least 6 characters.</div>',
                        prefill_username=username))
                    return

                ok = reset_user_password(username, new_pw)
                if not ok:
                    _html_resp(self, reset_password_page(
                        alert_html='<div class="alert-err">⚠ No account found with that username.</div>',
                        prefill_username=username))
                    return

                _html_resp(self, login_page(
                    alert_html='<div class="alert-ok">✅ Password updated successfully. Please sign in with your new password.</div>'))
                return

            user = get_session_user(self)
            if not user:
                _redirect(self, "/login"); return

            # ── Multipart routes (read body themselves — must come before _parse_body) ──

            if path == "/upload":
                content_type = self.headers.get("Content-Type", "")
                if "multipart/form-data" in content_type:
                    try:
                        import email, email.policy, tempfile, os

                        length   = int(self.headers.get("Content-Length", 0))
                        raw_body = self.rfile.read(length)

                        full_msg_bytes = (
                            f"Content-Type: {content_type}\r\n\r\n".encode() + raw_body
                        )
                        msg = email.message_from_bytes(
                            full_msg_bytes, policy=email.policy.compat32
                        )

                        file_data = None
                        filename  = None

                        for part in msg.walk():
                            cd = part.get("Content-Disposition", "")
                            if 'name="project_file"' in cd or "name='project_file'" in cd:
                                for segment in cd.split(";"):
                                    segment = segment.strip()
                                    if segment.startswith("filename="):
                                        filename = segment.split("=", 1)[1].strip().strip('"\'')
                                file_data = part.get_payload(decode=True)
                                break

                        if file_data and filename:
                            suffix   = ".xlsx" if filename.lower().endswith(".xlsx") else ".csv"
                            tmp_path = None
                            try:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                                    tmp.write(file_data)
                                    tmp_path = tmp.name
                                alert = _process_upload(tmp_path, suffix)
                            finally:
                                if tmp_path and os.path.exists(tmp_path):
                                    try:    os.unlink(tmp_path)
                                    except OSError: pass
                            _html_resp(self, project_data_entry(alert_html=alert, username=user))
                            return

                    except Exception as exc:
                        import traceback
                        _html_resp(self, project_data_entry(
                            alert_html=f'<div class="alert err">Upload error: {exc}<br>'
                                      f'<pre style="font-size:11px">{traceback.format_exc()}</pre></div>',
                            username=user))
                        return
                _redirect(self, "/data_entry")
                return
            elif path == "/check_folder_access":
                import os, json

                try:
                    length = int(self.headers.get("Content-Length", 0))
                    raw_body = self.rfile.read(length) if length > 0 else b"{}"
                    data = json.loads(raw_body.decode("utf-8") or "{}")
                except Exception:
                    data = {}

                folder_path = (data.get("folder_path") or "").strip()

                if not folder_path:
                    result = {"accessible": False, "message": "No path provided."}
                    status_code = 400
                else:
                    try:
                        if os.path.isdir(folder_path) and os.access(folder_path, os.R_OK):
                            result = {"accessible": True}
                        else:
                            result = {"accessible": False, "message": "Path not found or insufficient permissions."}
                    except Exception:
                        result = {"accessible": False, "message": "Error while checking this path."}
                    status_code = 200

                body = json.dumps(result).encode("utf-8")
                self.send_response(status_code)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            elif path == "/upload_expenditure":
                content_type = self.headers.get("Content-Type", "")
                if "multipart/form-data" in content_type:
                    try:
                        import email, email.policy, tempfile, os

                        length   = int(self.headers.get("Content-Length", 0))
                        raw_body = self.rfile.read(length)

                        full_msg_bytes = (
                            f"Content-Type: {content_type}\r\n\r\n".encode() + raw_body
                        )
                        msg = email.message_from_bytes(
                            full_msg_bytes, policy=email.policy.compat32
                        )

                        file_data = None
                        filename  = None

                        for part in msg.walk():
                            cd = part.get("Content-Disposition", "")
                            if 'name="expenditure_file"' in cd or "name='expenditure_file'" in cd:
                                for segment in cd.split(";"):
                                    segment = segment.strip()
                                    if segment.startswith("filename="):
                                        filename = segment.split("=", 1)[1].strip().strip('"\'')
                                file_data = part.get_payload(decode=True)
                                break

                        if file_data and filename:
                            tmp_path = None
                            try:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                                    tmp.write(file_data)
                                    tmp_path = tmp.name
                                alert = _process_expenditure_upload(tmp_path)
                            finally:
                                if tmp_path and os.path.exists(tmp_path):
                                    try:    os.unlink(tmp_path)
                                    except OSError: pass
                            _html_resp(self, expenditures(alert_html=alert, username=user))
                            return

                    except Exception as exc:
                        import traceback
                        _html_resp(self, expenditures(
                            alert_html=f'<div class="alert err">Upload error: {exc}<br>'
                                      f'<pre style="font-size:11px">{traceback.format_exc()}</pre></div>',
                            username=user))
                        return
                _redirect(self, "/expenditures")
                return

            # ── All other routes: parse form body first ──
            fields = _parse_body(self)

            if path == "/submit":
                pid = fields.get("id","").strip()
                if pid:
                    update_project(int(pid), fields)
                else:
                    save_project(fields)
                _redirect(self, "/data_entry")

            elif path == "/delete_project":
                pid = int(fields.get("project_id", 0))
                if pid: delete_project(pid)
                _redirect(self, "/data_entry")

            elif path == "/link_doc":
                link_document({
                    "project_id":   fields.get("project_id", None),
                    "doc_name":     fields.get("doc_name",""),
                    "doc_category": fields.get("doc_category",""),
                    "doc_link":     fields.get("doc_link","#"),
                })
                _redirect(self, "/data_entry")

            elif path == "/delete_doc":
                did = int(fields.get("doc_id", 0))
                if did: delete_document(did)
                _redirect(self, "/data_entry")

            elif path == "/expenditures_retrieve":
                pid = fields.get("project_id","")
                lob = fields.get("lob","")
                _redirect(self, f"/expenditures?project_id={pid}&lob={lob}")

            elif path == "/billing_collection_retrieve":
                pid = fields.get("project_id","")
                lob = fields.get("lob","")
                pf  = fields.get("pending_filter","all")
                _redirect(self, f"/billing_collection?project_id={pid}&lob={lob}&pending_filter={pf}")

            elif path == "/save_billing_entry":
                pid = fields.get("project_id","")
                lob = fields.get("lob","")
                save_billing_entry(fields)
                _redirect(self, f"/billing_collection?project_id={pid}&lob={lob}")

            elif path == "/record_collection":
                eid = int(fields.get("entry_id", 0))
                pid = fields.get("project_id","")
                lob = fields.get("lob","")
                if eid:
                    update_billing_collection(eid, fields.get("collection_date",""), fields.get("collection_amt",0))
                _redirect(self, f"/billing_collection?project_id={pid}&lob={lob}")

            elif path == "/delete_billing_entry":
                eid = int(fields.get("entry_id", 0))
                pid = fields.get("project_id","")
                lob = fields.get("lob","")
                if eid: delete_billing_entry(eid)
                _redirect(self, f"/billing_collection?project_id={pid}&lob={lob}")

            elif path == "/save_expenditure":
                result = save_expenditure(fields)
                pid = fields.get("project_id","")
                lob = fields.get("lob","")
                if not result:
                    _html_resp(self, expenditures(
                        alert_html='<div class="alert err">⚠ Please select and retrieve a project before updating ERP data.</div>',
                        prefill=fields, username=user))
                else:
                    _redirect(self, f"/expenditures?project_id={pid}&lob={lob}")

            elif path == "/save_entry":
                save_expense_entry(fields)
                pid = fields.get("project_id","")
                _redirect(self, f"/expenditures?project_id={pid}")

            elif path == "/delete_entry":
                eid = int(fields.get("entry_id", 0))
                if eid: delete_expense_entry(eid)
                _redirect(self, "/expenditures")

            elif path == "/financials_filter":
                _html_resp(self, financials(prefill=fields, username=user))

            elif path == "/save_variation":
                try:
                    orig = _n(fields.get("original_value"))
                    vamt = _n(fields.get("variation_amt"))
                    fields["revised_value"] = orig + vamt
                except: pass
                save_variation(fields)
                _redirect(self, "/variations")

            elif path == "/delete_variation":
                vid = int(fields.get("var_id", 0))
                if vid: delete_variation(vid)
                _redirect(self, "/variations")

            elif path == "/add_ref":
                key = fields.get("ref_key","")
                val = fields.get("ref_val","").strip()
                if key in REF_TABLE_MAP and val:
                    table, col = REF_TABLE_MAP[key]
                    add_ref_value(table, col, val)
                _redirect(self, "/data_entry")

            elif path == "/delete_ref":
                key = fields.get("ref_key","")
                val = fields.get("ref_val","").strip()
                if key in REF_TABLE_MAP and val:
                    table, col = REF_TABLE_MAP[key]
                    delete_ref_value(table, col, val)
                _redirect(self, "/data_entry")

            elif path == "/add_ref_option":
                cat = fields.get("category","")
                val = fields.get("value","").strip()
                if cat in REF_TABLE_MAP and val:
                    table, col = REF_TABLE_MAP[cat]
                    add_ref_value(table, col, val)
                _redirect(self, "/data_entry")

            elif path == "/delete_ref_option":
                cat = fields.get("category","")
                val = fields.get("value","").strip()
                if cat in REF_TABLE_MAP and val:
                    table, col = REF_TABLE_MAP[cat]
                    delete_ref_value(table, col, val)
                _redirect(self, "/data_entry")

            else:
                _redirect(self, "/")

        except Exception:
            _html_resp(self,
                f'<html><body><pre style="padding:32px;color:red">{traceback.format_exc()}</pre></body></html>',
                code=500)
# ═══════════════════════════════════════════════════════════════
#  SERVER LAUNCHER
# ═══════════════════════════════════════════════════════════════

def find_free_port(start=8080, end=9000) -> int:
    for port in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:    s.bind(("", port)); return port
            except: continue
    raise RuntimeError("No free port found.")


def run_server():
    init_db()
    port = find_free_port()
    url  = f"http://localhost:{port}"
    print(f"\n{'='*62}")
    print("  Monarch Surveyors — Financial Monitoring System  v4.1")
    print(f"{'='*62}")
    print(f"  Login        : {url}/login")
    print(f"  Dashboard    : {url}/dashboard")
    print(f"  Credentials  : admin / monarch123")
    print(f"  DB           : {DB_USER}@{DB_HOST}:{DB_PORT}/{DB_NAME}")
    print(f"  Press Ctrl+C to stop.")
    print(f"{'='*62}\n")
    with socketserver.TCPServer(("", port), MonarchHandler) as httpd:
        httpd.allow_reuse_address = True
        webbrowser.open(url + "/login")
        try:    httpd.serve_forever()
        except KeyboardInterrupt: print("\nServer stopped.")


if __name__ == "__main__":
    run_server()