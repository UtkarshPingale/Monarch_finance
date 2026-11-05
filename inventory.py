#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inventory_app.py — Monarch Surveyors Inventory Management System
=================================================================
ONE file contains everything:
  • Backend  — http.server (Python standard library, no pip installs)
  • Database — PostgreSQL via SQLAlchemy (no hardcoded data; all data
               entered through the application UI)
  • Frontend — server-rendered HTML + CSS + tiny inline JS

Run:
    python inventory_app.py
    → open http://localhost:9881   (login: admin / monarch123)

Screens (matching the Monarch design set):
    /dashboard     KPIs, consumption chart, low-stock alerts
    /items         Stock Directory — opening / purchased / distributed /
                   closing, rate, value, alert level, CSV export
    /requisitions  LOB demands; PENDING → IN TRANSIT → DELIVERED / REJECTED.
                   Delivering posts a stock-OUT and is blocked if stock < qty
    /procurement   Purchase orders; "Mark Received" posts a stock-IN.
                   Open orders older than 30 days flag as OVERDUE
    /reports       Material × LOB distribution matrix, CSV export
    /movements     Raw in/out ledger with manual entries

Stock is always derived, never stored twice:
    closing = opening + purchased(IN movements) − distributed(OUT movements)
"""

import csv
import email
import hashlib
import html
import io
import os
import re
import secrets
import socket
import socketserver
import sys
import threading
import traceback
import webbrowser
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse
from sqlalchemy import create_engine, text

# ─────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────

# ── PostgreSQL connection (mirror of monarch_app.py pattern) ──
DB_USER   = os.getenv("INV_DB_USER", "postgres")
DB_PASS   = os.getenv("INV_DB_PASS", "Monarch%40123")
DB_HOST   = os.getenv("INV_DB_HOST", "192.168.1.149")
DB_PORT   = os.getenv("INV_DB_PORT", "5432")
DB_NAME   = os.getenv("INV_DB_NAME", "postgres")
SCHEMA    = os.getenv("INV_DB_SCHEMA", "inventory_db")

ENGINE = create_engine(
    f"postgresql+psycopg2://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}",
    connect_args={"options": f"-csearch_path={SCHEMA}"},
    future=True,
    pool_pre_ping=True,
)

HOST = "0.0.0.0"
PORT = 9881

COOKIE_NAME   = "monarch_session"
SESSION_TTL   = 60 * 60 * 8          # 8 hours
SESSIONS: dict = {}                  # token -> session dict

LOGO_MARK = "MS"  # local text mark

LOBS = [
    "Railways", "DPR-Highways", "Land Acquisition", "Locals",
    "Geotechnical", "Geospatial", "Project Management", "ELU Preparation",
]

# ─────────────────────────────────────────────────────────────
#  CSS  (plain string — no f-string, no brace escaping needed)
# ─────────────────────────────────────────────────────────────

CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Inter,Arial,sans-serif;background:#f7f9fb;color:#191c1e;min-height:100vh}
a{text-decoration:none;color:inherit}
.layout{display:flex;min-height:100vh}
.sidebar{width:232px;background:#fff;border-right:1px solid #e0e3e5;position:fixed;top:0;left:0;bottom:0;display:flex;flex-direction:column;z-index:50}
.brand{display:flex;align-items:center;gap:10px;padding:18px 16px 14px;border-bottom:1px solid #eceef0;margin-bottom:6px}
.brand-mark{width:38px;height:38px;border-radius:8px;background:#000e24;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:15px;flex-shrink:0;overflow:hidden}
.brand-mark img{width:100%;height:100%;object-fit:contain;background:#fff}
.brand-name{font-size:13px;font-weight:800;color:#000e24;line-height:1.25}
.brand-sub{font-size:10px;color:#43474e;line-height:1.3}
.nav{flex:1;overflow-y:auto;padding:6px 10px}
.nav-item{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:8px;color:#43474e;font-size:13px;font-weight:600;margin-bottom:2px;transition:background .15s}
.nav-item:hover{background:#eef1f6}
.nav-item.active{background:#000e24;color:#fff}
.nav-item .ico{width:20px;text-align:center;flex-shrink:0}
.sidebar-foot{border-top:1px solid #eceef0;padding:12px 14px}
.user-row{display:flex;align-items:center;gap:9px}
.avatar{width:30px;height:30px;border-radius:50%;background:#775a19;color:#fff;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:800;flex-shrink:0}
.user-meta{font-size:12px;font-weight:700;line-height:1.2}
.user-role{font-size:10px;color:#43474e}
.logout{display:block;margin-top:10px;font-size:11px;font-weight:700;color:#ba1a1a}
.main{flex:1;margin-left:232px;min-width:0}
.topbar{height:54px;background:#fff;border-bottom:1px solid #e0e3e5;display:flex;align-items:center;justify-content:space-between;padding:0 24px;position:sticky;top:0;z-index:40}
.topbar-title{font-size:11px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#6b7280}
.topbar-user{display:flex;align-items:center;gap:8px;font-size:12px;font-weight:700}
.canvas{max-width:1440px;margin:0 auto;padding:22px 24px 40px}
.page-head{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;margin-bottom:20px;flex-wrap:wrap}
.page-head h1{font-size:26px;font-weight:800;color:#000e24;letter-spacing:-.01em}
.page-head p{font-size:14px;color:#43474e;margin-top:3px}
.head-actions{display:flex;gap:10px;flex-wrap:wrap}
.btn{display:inline-flex;align-items:center;gap:6px;border:none;border-radius:8px;padding:9px 16px;font-size:12.5px;font-weight:700;cursor:pointer;transition:all .15s;font-family:inherit}
.btn:hover{filter:brightness(1.08)}
.btn:active{transform:scale(.97)}
.btn-primary{background:#000e24;color:#fff}
.btn-outline{background:#fff;border:1px solid #000e24;color:#000e24}
.btn-sm{padding:5px 10px;font-size:11px;border-radius:6px}
.btn-navy{background:#000e24;color:#fff}
.btn-green{background:#16a34a;color:#fff}
.btn-red{background:#fff;border:1px solid #ba1a1a;color:#ba1a1a}
.btn-amber{background:#fed488;color:#5d4201}
.btn-gold{background:#775a19;color:#fff}
.alert{padding:11px 16px;border-radius:8px;margin-bottom:16px;font-size:13px;font-weight:700}
.alert.ok{background:#e8f5e9;color:#1b5e20;border:1px solid #a5d6a7}
.alert.err{background:#ffdad6;color:#93000a;border:1px solid #ffb4ab}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:16px;margin-bottom:22px}
.kpi{background:#fff;border:1px solid #e0e3e5;border-radius:10px;padding:16px 18px;box-shadow:0 1px 3px rgba(0,0,0,.04)}
.kpi .k-label{font-size:11px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;color:#43474e;display:flex;justify-content:space-between;align-items:center}
.kpi .k-val{font-size:26px;font-weight:800;color:#000e24;margin-top:6px;letter-spacing:-.01em}
.kpi .k-sub{font-size:12px;color:#6b7280;margin-top:4px}
.kpi.accent{border-left:4px solid #000e24}
.kpi.warn{border-left:4px solid #775a19}
.kpi.danger{border-left:4px solid #ba1a1a}
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px}
.card{background:#fff;border:1px solid #e0e3e5;border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.04)}
.card-head{padding:14px 18px;border-bottom:1px solid #eceef0;display:flex;align-items:center;justify-content:space-between;gap:10px}
.card-head h3{font-size:14px;font-weight:800;color:#000e24}
.card-body{padding:6px 18px 18px}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;padding:10px 12px;font-size:11px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;color:#43474e;background:#f2f4f6;border-bottom:1px solid #e0e3e5;white-space:nowrap}
td{padding:10px 12px;border-bottom:1px solid #f0f2f4;vertical-align:middle}
tr:hover td{background:#fafbfc}
.num{text-align:right;font-variant-numeric:tabular-nums}
.mono{font-family:Consolas,Menlo,monospace;font-size:12.5px}
.badge{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:800}
.badge .dot{width:6px;height:6px;border-radius:50%;background:currentColor}
.badge-pending{background:#fff3cd;color:#856404}
.badge-transit{background:#d6e3ff;color:#00234b}
.badge-delivered{background:#d1fae5;color:#065f46}
.badge-rejected{background:#ffdad6;color:#93000a}
.badge-received{background:#d1fae5;color:#065f46}
.badge-overdue{background:#ffdad6;color:#93000a}
.badge-ok{background:#d1fae5;color:#065f46}
.badge-low{background:#fff3cd;color:#856404}
.empty{color:#9aa0a6;font-style:italic;text-align:center;padding:26px 0}
.form-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-top:6px}
.field label{display:block;font-size:11px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;color:#43474e;margin-bottom:5px}
.field input,.field select,.field textarea{width:100%;border:1px solid #c4c6d0;border-radius:8px;padding:9px 12px;font-size:13px;font-family:inherit;background:#fff;outline:none}
.field input:focus,.field select:focus,.field textarea:focus{border-color:#000e24;box-shadow:0 0 0 2px rgba(0,14,36,.12)}
.filter-bar{background:#fff;border:1px solid #e0e3e5;border-radius:10px;padding:14px 16px;margin-bottom:16px;display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap}
.filter-bar .field{flex:1;min-width:150px}
.inline-form{display:inline}
.table-wrap{overflow-x:auto}
.chart{display:flex;align-items:flex-end;gap:14px;height:200px;padding:10px 6px 0;border-bottom:1px solid #e0e3e5}
.bar-col{flex:1;display:flex;flex-direction:column;align-items:center;gap:6px;min-width:0}
.bar-track{width:100%;max-width:46px;height:160px;display:flex;align-items:flex-end;background:#f2f4f6;border-radius:6px 6px 0 0;overflow:hidden}
.bar{width:100%;background:#00234b;border-radius:6px 6px 0 0;transition:height .3s}
.bar-label{font-size:10px;font-weight:800;color:#43474e;text-transform:uppercase;white-space:nowrap}
.bar-val{font-size:10px;color:#6b7280}
.pagination{display:flex;justify-content:space-between;align-items:center;padding:12px 18px;border-top:1px solid #eceef0;font-size:12px;color:#43474e}
.pill{display:inline-block;background:#f2f4f6;color:#43474e;border-radius:6px;padding:2px 8px;font-size:11px;font-weight:700}
.matrix th,.matrix td{white-space:nowrap}
.tfoot td{background:#000e24;color:#fff;font-weight:800}
@media(max-width:900px){.grid-2{grid-template-columns:1fr}.sidebar{transform:translateX(-100%);transition:transform .2s}.sidebar.open{transform:translateX(0)}.main{margin-left:0}}
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(15,23,42,.5);z-index:200;align-items:center;justify-content:center;padding:16px}
.modal-overlay.open{display:flex}
.modal-box{background:#fff;border-radius:12px;width:100%;max-width:340px;padding:22px;box-shadow:0 20px 50px rgba(0,0,0,.25)}
.modal-box h3{font-size:15px;font-weight:800;color:#000e24;margin-bottom:4px}
.modal-box p{font-size:12px;color:#6b7280;margin-bottom:16px}
.modal-actions{display:flex;gap:10px;margin-top:18px}
.modal-actions .btn{flex:1;justify-content:center}
"""

# ─────────────────────────────────────────────────────────────
#  SCHEMA
# ─────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────
#  SCHEMA  (PostgreSQL DDL)
# ─────────────────────────────────────────────────────────────

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS users (
    id          SERIAL PRIMARY KEY,
    username    TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    display_name  TEXT,
    role        TEXT DEFAULT 'superadmin',
    created_at  TEXT
);
CREATE TABLE IF NOT EXISTS items (
    id          SERIAL PRIMARY KEY,
    code        TEXT UNIQUE NOT NULL,
    name        TEXT NOT NULL,
    category    TEXT,
    unit        TEXT,
    opening_qty REAL DEFAULT 0,
    rate        REAL DEFAULT 0,
    alert_qty   REAL DEFAULT 0,
    created_at  TEXT
);
CREATE TABLE IF NOT EXISTS stock_movements (
    id          SERIAL PRIMARY KEY,
    item_id     INTEGER NOT NULL REFERENCES items(id) ON DELETE CASCADE,
    qty         REAL NOT NULL,
    direction   TEXT NOT NULL CHECK(direction IN ('IN','OUT')),
    ref_type    TEXT,
    ref_no      TEXT,
    m_date      TEXT,
    note        TEXT,
    created_by  TEXT,
    created_at  TEXT
);
CREATE TABLE IF NOT EXISTS requisitions (
    id            SERIAL PRIMARY KEY,
    req_no        TEXT UNIQUE NOT NULL,
    item_id       INTEGER NOT NULL REFERENCES items(id) ON DELETE CASCADE,
    lob           TEXT NOT NULL,
    qty           REAL NOT NULL,
    req_date      TEXT,
    status        TEXT DEFAULT 'PENDING',
    note          TEXT,
    delivery_date TEXT,
    created_by    TEXT,
    created_at    TEXT
);
CREATE TABLE IF NOT EXISTS procurement_orders (
    id           SERIAL PRIMARY KEY,
    po_no        TEXT UNIQUE NOT NULL,
    item_id      INTEGER NOT NULL REFERENCES items(id) ON DELETE CASCADE,
    qty          REAL NOT NULL,
    rate         REAL DEFAULT 0,
    supplier     TEXT,
    order_date   TEXT,
    receipt_date TEXT,
    status       TEXT DEFAULT 'PENDING',
    created_by   TEXT,
    created_at   TEXT
);
"""

# ─────────────────────────────────────────────────────────────
#  DB HELPERS  (SQLAlchemy — same pattern as monarch_app.py)
# ─────────────────────────────────────────────────────────────

class _ConnCtx:
    """Thin wrapper returned by _conn() so existing call-sites that do
    conn.execute(sql, params) / conn.commit() / conn.close() keep working
    without modification.  Internally it uses SQLAlchemy's connection."""

    def __init__(self):
        self._ctx = ENGINE.begin()       # auto-commits on __exit__
        self._conn = self._ctx.__enter__()
        self._closed = False

    # ---- sqlite3-compatible execute ----
    def execute(self, sql, params=()):
        """Accept both positional-? (sqlite3 style) and :name params."""
        # Convert positional ? placeholders to SQLAlchemy :p0, :p1 …
        import re
        named_sql = sql
        named_params: dict = {}
        if params and '?' in sql:
            parts = re.split(r'(\?)', sql)
            idx = 0
            out_parts = []
            for part in parts:
                if part == '?':
                    key = f'p{idx}'
                    named_params[key] = params[idx]
                    out_parts.append(f':{key}')
                    idx += 1
                else:
                    out_parts.append(part)
            named_sql = ''.join(out_parts)
        elif params and isinstance(params, dict):
            named_params = params
        return self._conn.execute(text(named_sql), named_params)

    def commit(self):
        """No-op: ENGINE.begin() auto-commits on close."""
        pass

    def close(self):
        if not self._closed:
            self._closed = True
            self._ctx.__exit__(None, None, None)

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self._ctx.__exit__(*args)


def _conn():
    """Return a DB connection context compatible with the existing call-sites."""
    return _ConnCtx()


def _q(conn, sql, params=()):
    """Run a SELECT and return list-of-dicts."""
    cur = conn.execute(sql, params)
    cols = list(cur.keys()) if hasattr(cur, 'keys') else []
    return [dict(zip(cols, r)) for r in cur.fetchall()]


def _one(conn, sql, params=()):
    """Run a SELECT and return the first row as dict, or None."""
    cur = conn.execute(sql, params)
    cols = list(cur.keys()) if hasattr(cur, 'keys') else []
    row = cur.fetchone()
    return dict(zip(cols, row)) if row else None


def _esc(v):
    return html.escape(str(v if v is not None else ""), quote=True)


def _group(ip):
    if len(ip) <= 3:
        return ip
    head, tail = ip[:-3], ip[-3:]
    parts = []
    while len(head) > 2:
        parts.insert(0, head[-2:])
        head = head[:-2]
    if head:
        parts.insert(0, head)
    return ",".join(parts) + "," + tail


def _money(v, dec=2):
    try:
        f = float(v or 0)
    except (TypeError, ValueError):
        f = 0.0
    neg = "-" if f < 0 else ""
    f = abs(f)
    s = "{:.{}f}".format(f, dec)
    ip, _, fp = s.partition(".")
    out = neg + _group(ip)
    if dec:
        out += "." + fp
    return out


def _inr(v):
    return _money(v, 0)


def _num_raw(v):
    """Plain numeric string (no thousands grouping) — safe for URL query
    params and numeric <input> values. Trims a trailing .0."""
    try:
        f = float(v or 0)
    except (TypeError, ValueError):
        f = 0.0
    if f == int(f):
        return str(int(f))
    return "{:g}".format(f)


def _today():
    return datetime.now().strftime("%Y-%m-%d")


def find_column_map(headers):
    """Smart detector mapping spreadsheet header titles to internal item attributes."""
    mapping = {}
    normalized = [str(h).strip().lower() if h is not None else "" for h in headers]
    patterns = {
        "code": [
            r"^(material\s*code|item\s*code|mat\s*code|product\s*code|sku)$",
            r"^(sr\.?\s*no\.?|srno|serial\s*no\.?|s\.?\s*no\.?)$",
            r"^(code|item_id|id)$"
        ],
        "name": [
            r"^(material\s*name|item\s*name|product\s*name)$",
            r"^(material|item|particulars|description|item\s*desc)$",
            r"^(name|title)$"
        ],
        "qty": [
            r"^(total\s*no\.?|total\s*nos\.?|total\s*qty|total\s*count)$",
            r"^(opening\s*qty|opening\s*stock|opening\s*balance|opening)$",
            r"^(qty|quantity|stock|balance|total|count)$"
        ],
        "unit": [
            r"^(unit\s*of\s*measure|uom|unit)$"
        ],
        "rate": [
            r"^(unit\s*rate|unit\s*price|purchase\s*rate|cost\s*price)$",
            r"^(rate|rating|price|cost|val)$"
        ],
        "category": [
            r"^(material\s*category|item\s*category|category|group|department|dept|type)$"
        ],
        "alert_qty": [
            r"^(alert\s*qty|alert\s*level|min\s*qty|min\s*stock|reorder\s*level|minimum)$"
        ]
    }
    used_indices = set()
    for field, pat_list in patterns.items():
        found = False
        for pat in pat_list:
            for idx, h in enumerate(normalized):
                if idx in used_indices or not h:
                    continue
                if re.search(pat, h):
                    mapping[field] = idx
                    used_indices.add(idx)
                    found = True
                    break
            if found:
                break
    return mapping


def parse_qty_unit(raw):
    """Extract numeric quantity and unit from raw cell values (e.g. '4 Nos', '1 Box', 'Both 2')."""
    if raw is None:
        return 0.0, "Nos"
    if isinstance(raw, (int, float)):
        return float(raw), "Nos"
    s = str(raw).strip()
    if not s:
        return 0.0, "Nos"

    m_both = re.match(r"^both\s+(\d+(?:\.\d+)?)$", s, re.IGNORECASE)
    if m_both:
        return float(m_both.group(1)), "Set"

    m_frac = re.match(r"^(\d+)\s+1/2\s*([a-zA-Z]+)?$", s, re.IGNORECASE)
    if m_frac:
        val = float(m_frac.group(1)) + 0.5
        u = m_frac.group(2)
        unit = u.capitalize() if u else "Nos"
        return val, unit

    m = re.match(r"^(\d+(?:\.\d+)?)\s*([a-zA-Z]+)?$", s, re.IGNORECASE)
    if m:
        val = float(m.group(1))
        u = m.group(2)
        if u:
            u_clean = u.capitalize()
            norm = {
                "no": "Nos", "nos": "Nos",
                "pkt": "Pkt", "pkts": "Pkt",
                "box": "Box", "boxes": "Box",
                "roll": "Roll", "rolls": "Roll",
                "pc": "Pcs", "pcs": "Pcs",
                "book": "Book", "books": "Book",
                "file": "Files", "files": "Files",
                "pen": "Nos", "pens": "Nos",
                "bundle": "Bundle", "bundles": "Bundle"
            }
            unit = norm.get(u_clean.lower(), u_clean)
        else:
            unit = "Nos"
        return val, unit

    return 0.0, "Nos"


def determine_category(name):
    """Categorize item by name keywords if category is not provided."""
    c = name.lower()
    if any(re.search(r"\b" + re.escape(k) + r"\b", c) for k in [
        "canon", "epson", "hp 72", "hp 730", "konica", "ink", "toner",
        "refil", "refill", "cartridge", "maintaince", "maintain box"
    ]):
        return "Printing & Ink"

    if any(re.search(r"\b" + re.escape(k) + r"\b", c) for k in [
        "laptop bag", "pen drive", "pendrive", "mouse pad", "calculator",
        "shreder", "extension board", "gps tracker", "battery", "duracell"
    ]):
        return "Electronics & IT"

    if any(re.search(r"\b" + re.escape(k) + r"\b", c) for k in ["safety", "jacket"]):
        return "Safety"

    if any(re.search(r"\b" + re.escape(k) + r"\b", c) for k in [
        "rat killer", "good night", "odomos", "freshner", "whitener", "liquid", "micro fiber cloth"
    ]):
        return "Cleaning"

    return "Stationary"


def _stock_map(conn):
    """Derived stock per item: opening, purchased(IN), distributed(OUT), closing."""
    rows = _q(conn, """
        SELECT i.id,
               COALESCE(i.opening_qty,0) AS opening,
               COALESCE((SELECT SUM(qty) FROM stock_movements m
                         WHERE m.item_id = i.id AND m.direction = 'IN'),0) AS purchased,
               COALESCE((SELECT SUM(qty) FROM stock_movements m
                         WHERE m.item_id = i.id AND m.direction = 'OUT'),0) AS distributed
        FROM items i
    """)
    out = {}
    for r in rows:
        closing = r["opening"] + r["purchased"] - r["distributed"]
        out[r["id"]] = {**r, "closing": closing}
    return out


def _requisition_shortfalls(conn):
    """Items where total open (PENDING / IN TRANSIT) requisition demand
    exceeds current closing stock — these need a purchase order."""
    stock = _stock_map(conn)
    items = {it["id"]: it for it in _q(conn, "SELECT * FROM items")}
    demand_rows = _q(conn, """
        SELECT item_id, SUM(qty) AS demand
        FROM requisitions
        WHERE status IN ('PENDING','IN TRANSIT')
        GROUP BY item_id
    """)
    out = []
    for r in demand_rows:
        closing = stock.get(r["item_id"], {}).get("closing", 0)
        if r["demand"] > closing:
            it = items.get(r["item_id"], {})
            out.append({
                "item_id": r["item_id"],
                "code": it.get("code") or "—",
                "name": it.get("name") or "—",
                "unit": it.get("unit") or "",
                "demand": r["demand"],
                "closing": closing,
                "shortfall": r["demand"] - closing,
            })
    out.sort(key=lambda x: x["shortfall"], reverse=True)
    return out


# ─────────────────────────────────────────────────────────────
#  INIT  (schema + admin user only — no hardcoded seed data)
# ─────────────────────────────────────────────────────────────

def init_db():
    """Create the schema (if it doesn't exist) and ensure the admin user
    is present.  No sample / seed data is inserted — all data enters the
    system through the application UI."""

    # Ensure the PostgreSQL schema (namespace) exists
    with ENGINE.begin() as conn:
        conn.execute(text(f"CREATE SCHEMA IF NOT EXISTS {SCHEMA}"))

    # Create tables (DDL must run outside the schema-creation transaction
    # because SQLAlchemy's begin() may not support mixing DDL + DML in
    # some PostgreSQL setups; use separate begin blocks for safety)
    for stmt in SCHEMA_SQL.strip().split(';'):
        stmt = stmt.strip()
        if not stmt:
            continue
        try:
            with ENGINE.begin() as conn:
                conn.execute(text(stmt))
        except Exception as e:
            print(f"[init_db] DDL warning (likely already exists): {e}")

    # Ensure the built-in admin account exists
    with ENGINE.begin() as conn:
        row = conn.execute(
            text("SELECT id FROM users WHERE username = :u"), {"u": "admin"}
        ).fetchone()
        if not row:
            conn.execute(text("""
                INSERT INTO users (username, password_hash, display_name, role, created_at)
                VALUES (:u, :ph, :dn, :r, :ca)
            """), {
                "u":  "admin",
                "ph": hashlib.sha256(b"monarch123").hexdigest(),
                "dn": "Admin",
                "r":  "superadmin",
                "ca": _today(),
            })
    print("[init_db] Database ready.")


def _stock_map_period(conn, start, end):
    """Derived stock per item for a specific date window [start, end] (YYYY-MM-DD):
    opening = balance just before the window, purchased/distributed = movements
    inside the window, closing = balance at the end of the window."""
    rows = _q(conn, """
        SELECT i.id,
               COALESCE(i.opening_qty,0)
                 + COALESCE((SELECT SUM(qty) FROM stock_movements m
                             WHERE m.item_id=i.id AND m.direction='IN' AND m.m_date < :s1),0)
                 - COALESCE((SELECT SUM(qty) FROM stock_movements m
                             WHERE m.item_id=i.id AND m.direction='OUT' AND m.m_date < :s2),0)
                 AS opening,
               COALESCE((SELECT SUM(qty) FROM stock_movements m
                         WHERE m.item_id=i.id AND m.direction='IN'
                               AND m.m_date >= :s3 AND m.m_date <= :e1),0) AS purchased,
               COALESCE((SELECT SUM(qty) FROM stock_movements m
                         WHERE m.item_id=i.id AND m.direction='OUT'
                               AND m.m_date >= :s4 AND m.m_date <= :e2),0) AS distributed
        FROM items i
    """, {"s1": start, "s2": start, "s3": start, "e1": end, "s4": start, "e2": end})
    out = {}
    for r in rows:
        closing = r["opening"] + r["purchased"] - r["distributed"]
        out[r["id"]] = {**r, "closing": closing}
    return out


# _seed() removed — no hardcoded data is inserted.
# All inventory data is entered through the application UI.


# ─────────────────────────────────────────────────────────────
#  HTML LAYOUT
# ─────────────────────────────────────────────────────────────

NAV = [
    ("/dashboard",    "dashboard",    "▦", "Dashboard"),
    ("/items",        "items",        "▤", "Stock Directory"),
    ("/requisitions", "requisitions", "≣", "Requisitions"),
    ("/procurement",  "procurement",  "☰", "Procurement"),
    ("/reports",      "reports",      "▥", "Reports"),
    ("/movements",    "movements",    "⇄", "Movements"),
    ("/users",        "users",        "☺", "User Accounts"),
]


def _nav_html(active):
    out = ""
    for href, key, ico, label in NAV:
        cls = "nav-item active" if key == active else "nav-item"
        out += '<a class="{}" href="{}"><span class="ico">{}</span>{}</a>'.format(cls, href, ico, label)
    return out


def _layout(title, active, user, content):
    initials = "".join(w[0] for w in (user.get("display_name") or "Admin").split()[:2]).upper() or "A"
    nav = _nav_html(active)
    return """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title} | Monarch Inventory</title>
<style>{css}</style>
</head>
<body>
<div class="layout">
  <aside class="sidebar">
    <div class="brand">
      <div class="brand-mark">{logo}</div>
      <div>
        <div class="brand-name">Monarch Surveyors</div>
        <div class="brand-sub">&amp; Engineering Consultants Ltd.</div>
      </div>
    </div>
    <nav class="nav">{nav}</nav>
    <div class="sidebar-foot">
      <div class="user-row">
        <div class="avatar">{initials}</div>
        <div>
          <div class="user-meta">{uname}</div>
          <div class="user-role">{role}</div>
        </div>
      </div>
      <a class="logout" href="/logout">Sign out</a>
    </div>
  </aside>
  <main class="main">
    <header class="topbar">
      <span class="topbar-title">Inventory Management System</span>
      <span class="topbar-user"><span class="avatar" style="width:26px;height:26px;font-size:10px">{initials}</span>{uname}</span>
    </header>
    <div class="canvas">{content}</div>
  </main>
</div>
</body>
</html>""".format(
        title=_esc(title), css=CSS, nav=nav, logo=_esc(LOGO_MARK),
        initials=_esc(initials), uname=_esc(user.get("display_name") or user.get("username") or "Admin"),
        role=_esc(user.get("role") or ""), content=content,
    )


def _page_head(title, sub, actions=""):
    return '<div class="page-head"><div><h1>{}</h1><p>{}</p></div><div class="head-actions">{}</div></div>'.format(
        _esc(title), _esc(sub), actions,
    )


def _flash(qs):
    ok = (qs.get("ok") or [""])[0]
    err = (qs.get("err") or [""])[0]
    out = ""
    if ok:
        out += '<div class="alert ok">&#10003; {}</div>'.format(_esc(ok))
    if err:
        out += '<div class="alert err">&#9888; {}</div>'.format(_esc(err))
    return out


def _status_badge(status):
    m = {
        "PENDING": ("badge-pending", "Pending"),
        "IN TRANSIT": ("badge-transit", "In Transit"),
        "DELIVERED": ("badge-delivered", "Delivered"),
        "REJECTED": ("badge-rejected", "Rejected"),
        "RECEIVED": ("badge-received", "Received"),
        "OVERDUE": ("badge-overdue", "Overdue"),
    }
    cls, label = m.get(status, ("badge-pending", status or "—"))
    return '<span class="badge {}"><span class="dot"></span>{}</span>'.format(cls, _esc(label))


def _item_options(conn, selected_id=None):
    opts = ""
    for it in _q(conn, "SELECT id, code, name, unit FROM items ORDER BY code"):
        sel = " selected" if it["id"] == selected_id else ""
        opts += '<option value="{}"{}>{} — {} ({})</option>'.format(
            it["id"], sel, _esc(it["code"]), _esc(it["name"]), _esc(it["unit"] or "unit"))
    return opts


def _lob_options(selected=None):
    opts = ""
    for lob in LOBS:
        sel = " selected" if lob == selected else ""
        opts += '<option value="{}"{}>{}</option>'.format(_esc(lob), sel, _esc(lob))
    return opts


def _select(name, options_html, all_label):
    return '<select name="{}"><option value="">{}</option>{}</select>'.format(
        name, _esc(all_label), options_html)


# ─────────────────────────────────────────────────────────────
#  PAGES
# ─────────────────────────────────────────────────────────────

def page_login(err=""):
    err_html = '<div class="alert err">&#9888; {}</div>'.format(_esc(err)) if err else ""
    return """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Sign in | Monarch Inventory</title>
<style>{css}</style>
</head>
<body style="display:flex;align-items:center;justify-content:center;min-height:100vh">
<div style="width:100%;max-width:400px;padding:24px">
  <div style="background:#fff;border:1px solid #e0e3e5;border-radius:14px;padding:32px;box-shadow:0 10px 30px rgba(0,0,0,.06)">
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:22px">
      <div class="brand-mark" style="width:44px;height:44px;font-size:17px">{logo}</div>
      <div>
        <div style="font-size:15px;font-weight:800;color:#000e24">Monarch Surveyors</div>
        <div style="font-size:11px;color:#43474e">&amp; Engineering Consultants Ltd.</div>
      </div>
    </div>
    <h1 style="font-size:20px;font-weight:800;color:#000e24;margin-bottom:4px">Inventory Management</h1>
    <p style="font-size:13px;color:#6b7280;margin-bottom:20px">Sign in to manage stock, requisitions and procurement.</p>
    {err_html}
    <form method="post" action="/login">
      <div class="field" style="margin-bottom:14px">
        <label>Username</label>
        <input type="text" name="username" autocomplete="username" required>
      </div>
      <div class="field" style="margin-bottom:20px">
        <label>Password</label>
        <input type="password" name="password" autocomplete="current-password" required>
      </div>
      <button class="btn btn-primary" style="width:100%;justify-content:center;padding:11px">Sign In</button>
    </form>
    <p style="margin-top:16px;font-size:11px;color:#9aa0a6;text-align:center">Default login: admin / monarch123</p>
  </div>
</div>
</body>
</html>""".format(css=CSS, err_html=err_html, logo=_esc(LOGO_MARK))


def page_dashboard(conn, qs, user):
    items = _q(conn, "SELECT * FROM items ORDER BY code")
    stock = _stock_map(conn)

    total_value = 0.0
    low = []
    for it in items:
        closing = stock[it["id"]]["closing"]
        total_value += closing * (it["rate"] or 0)
        if it["alert_qty"] and closing <= it["alert_qty"]:
            low.append({**it, "closing": closing})

    pending_reqs = _q(conn, "SELECT * FROM requisitions WHERE status IN ('PENDING','IN TRANSIT') ORDER BY req_date DESC LIMIT 8")
    pending_count = len(_q(conn, "SELECT id FROM requisitions WHERE status IN ('PENDING','IN TRANSIT')"))
    po_count = len(_q(conn, "SELECT id FROM procurement_orders"))

    # consumption by category (from OUT movements)
    cats = _q(conn, """
        SELECT COALESCE(i.category,'Other') AS cat, SUM(m.qty) AS total
        FROM stock_movements m JOIN items i ON i.id = m.item_id
        WHERE m.direction='OUT'
        GROUP BY cat ORDER BY total DESC
    """)
    bars = ""
    if cats:
        mx = max(c["total"] for c in cats)
        for c in cats:
            h = int(round(c["total"] / mx * 160)) if mx else 0
            bars += ('<div class="bar-col">'
                     '<div class="bar-track"><div class="bar" style="height:{}px"></div></div>'
                     '<div class="bar-label">{}</div><div class="bar-val">{}</div></div>').format(
                max(h, 4), _esc(c["cat"]), _inr(c["total"]))

    # low stock rows
    low_rows = ""
    for it in low:
        low_rows += ('<tr><td class="mono">{}</td><td style="font-weight:700">{}</td>'
                     '<td>{}</td><td class="num" style="color:#ba1a1a;font-weight:800">{}</td>'
                     '<td class="num">{}</td><td class="num"><a class="btn btn-primary btn-sm" '
                     'href="/procurement?new=1&amp;item_id={}">Reorder</a></td></tr>').format(
            _esc(it["code"]), _esc(it["name"]), _esc(it["category"] or "—"),
            _inr(it["closing"]), _inr(it["alert_qty"]), it["id"])
    if not low_rows:
        low_rows = '<tr><td colspan="6" class="empty">No items at or below alert level.</td></tr>'

    # pending requisitions rows
    req_rows = ""
    for r in pending_reqs:
        it = next((x for x in items if x["id"] == r["item_id"]), {})
        req_rows += ('<tr><td class="mono">{}</td><td style="font-weight:700">{}</td>'
                     '<td>{}</td><td class="num">{}</td><td>{}</td><td>{}</td></tr>').format(
            _esc(r["req_no"]), _esc(it.get("name") or "—"), _esc(r["lob"]),
            _inr(r["qty"]), _esc(r["req_date"]), _status_badge(r["status"]))
    if not req_rows:
        req_rows = '<tr><td colspan="6" class="empty">No pending requisitions.</td></tr>'

    content = _flash(qs) + _page_head(
        "Inventory Dashboard",
        "Real-time oversight of stock, demands and procurement cycles.",
        '<a class="btn btn-outline" href="/items">&#128196; Stock Directory</a>'
        '<a class="btn btn-primary" href="/procurement?new=1">&#43; New Procurement</a>',
    )

    content += """
    <div class="kpi-grid">
      <div class="kpi accent"><div class="k-label">Total Stock Value</div>
        <div class="k-val">&#8377;{val}</div><div class="k-sub">closing &times; rate, all materials</div></div>
      <div class="kpi danger"><div class="k-label">Low Stock Alerts</div>
        <div class="k-val">{low}</div><div class="k-sub">requires immediate procurement</div></div>
      <div class="kpi warn"><div class="k-label">Pending Requisitions</div>
        <div class="k-val">{reqs}</div><div class="k-sub">awaiting approval / transit</div></div>
      <div class="kpi accent"><div class="k-label">Total PO Raised</div>
        <div class="k-val">{pos}</div><div class="k-sub">purchase orders on record</div></div>
    </div>
    <div class="grid-2">
      <div class="card">
        <div class="card-head"><h3>Material Consumption Trends</h3>
          <span class="pill">by category &middot; distributed qty</span></div>
        <div class="card-body"><div class="chart">{bars}</div></div>
      </div>
      <div class="card">
        <div class="card-head"><h3>Low Stock Alerts</h3>
          <a href="/items" style="font-size:12px;font-weight:700;color:#000e24">View all &rarr;</a></div>
        <div class="table-wrap"><table>
          <thead><tr><th>Code</th><th>Material</th><th>Category</th><th class="num">Current</th><th class="num">Alert</th><th></th></tr></thead>
          <tbody>{low_rows}</tbody>
        </table></div>
      </div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Pending Requisitions</h3>
        <a href="/requisitions" style="font-size:12px;font-weight:700;color:#000e24">View all &rarr;</a></div>
      <div class="table-wrap"><table>
        <thead><tr><th>Req No</th><th>Material</th><th>LOB</th><th class="num">Qty</th><th>Req Date</th><th>Status</th></tr></thead>
        <tbody>{req_rows}</tbody>
      </table></div>
    </div>
    """.format(
        val=_inr(total_value), low=len(low), reqs=pending_count, pos=po_count,
        bars=bars, low_rows=low_rows, req_rows=req_rows,
    )
    return _layout("Dashboard", "dashboard", user, content)


MONTHS = [
    ("01", "January"), ("02", "February"), ("03", "March"), ("04", "April"),
    ("05", "May"), ("06", "June"), ("07", "July"), ("08", "August"),
    ("09", "September"), ("10", "October"), ("11", "November"), ("12", "December"),
]


def page_items(conn, qs, user):
    import calendar

    q = (qs.get("q") or [""])[0].strip().lower()
    cat = (qs.get("cat") or [""])[0]
    al = (qs.get("al") or [""])[0]
    month = (qs.get("month") or [""])[0].strip()
    year = (qs.get("year") or [""])[0].strip()

    items = _q(conn, "SELECT * FROM items ORDER BY code")
    period_label = ""
    if month and year:
        try:
            y, m = int(year), int(month)
            last_day = calendar.monthrange(y, m)[1]
            start = "{:04d}-{:02d}-01".format(y, m)
            end = "{:04d}-{:02d}-{:02d}".format(y, m, last_day)
            stock = _stock_map_period(conn, start, end)
            period_label = " — {} {}".format(dict(MONTHS).get(month, month), year)
        except ValueError:
            stock = _stock_map(conn)
    else:
        stock = _stock_map(conn)
    cats = sorted({it["category"] for it in items if it["category"]})

    rows = []
    for it in items:
        s = stock[it["id"]]
        closing = s["closing"]
        if q and q not in (it["code"] + " " + it["name"]).lower():
            continue
        if cat and it["category"] != cat:
            continue
        if al == "low" and not (it["alert_qty"] and closing <= it["alert_qty"]):
            continue
        if al == "ok" and it["alert_qty"] and closing <= it["alert_qty"]:
            continue
        rows.append({**it, **s})

    total_value = sum(r["closing"] * (r["rate"] or 0) for r in rows)
    low_count = sum(1 for r in rows if r["alert_qty"] and r["closing"] <= r["alert_qty"])

    # alert badge
    def alert_badge(it):
        if it["alert_qty"] and it["closing"] <= it["alert_qty"]:
            return '<span class="badge badge-low"><span class="dot"></span>Low</span>'
        return '<span class="badge badge-ok"><span class="dot"></span>OK</span>'

    body_rows = ""
    for r in rows:
        closing = r["closing"]
        val = closing * (r["rate"] or 0)
        body_rows += (
            '<tr>'
            '<td class="mono">{code}</td>'
            '<td style="font-weight:700">{name}</td>'
            '<td>{unit}</td>'
            '<td class="num">{opening}</td>'
            '<td class="num" style="color:#065f46">{purchased}</td>'
            '<td class="num" style="color:#93000a">{distributed}</td>'
            '<td class="num" style="font-weight:800">{closing}</td>'
            '<td class="num">{rate}</td>'
            '<td class="num" style="font-weight:800">&#8377;{val}</td>'
            '<td>{badge}</td>'
            '<td class="num" style="white-space:nowrap">'
            '<a class="btn btn-outline btn-sm" href="/items?edit={id}">Edit</a> '
            '<form class="inline-form" method="post" action="/items/delete" '
            'onsubmit="return confirm(&#39;Delete this material?&#39;)">'
            '<input type="hidden" name="id" value="{id}">'
            '<button class="btn btn-red btn-sm">Delete</button></form>'
            '</td></tr>'
        ).format(
            code=_esc(r["code"]), name=_esc(r["name"]), unit=_esc(r["unit"] or "—"),
            opening=_inr(r["opening"]), purchased=_inr(r["purchased"]),
            distributed=_inr(r["distributed"]), closing=_inr(closing),
            rate=_money(r["rate"]), val=_inr(val), badge=alert_badge(r), id=r["id"],
        )
    if not body_rows:
        body_rows = '<tr><td colspan="11" class="empty">No materials match the filters.</td></tr>'

    # import excel form
    import_html = ""
    if (qs.get("import") or [""])[0] == "1":
        import_html = """
        <div class="card" style="margin-bottom:18px; border:2px dashed #000e24; background:#fafbfc">
          <div class="card-head" style="display:flex; justify-content:space-between; align-items:center;">
            <h3>📥 Import Stock Directory from Excel</h3>
            <a href="/items" class="btn btn-outline btn-sm">✕ Close</a>
          </div>
          <div class="card-body">
            <form method="post" action="/items/import" enctype="multipart/form-data">
              <div style="background:#e8f0fe; border-left:4px solid #1a73e8; border-radius:4px; padding:12px 16px; margin-bottom:16px; font-size:13px; color:#1e3a8a; line-height:1.5;">
                <b>⚡ Smart Column Mapping:</b> Automatically detects columns for <b>Material Name</b>, <b>Code / Sr.No.</b>, <b>Quantity</b> (handles mixed text like <i>"4 Nos"</i>, <i>"1 Box"</i>, <i>"4.5 Pkt"</i>, <i>"Both 2"</i>), <b>Unit</b>, <b>Rate</b> (defaults to 0.00 if omitted so admin can add rating later), and <b>Category</b>.
              </div>
              <div class="form-grid">
                <div class="field" style="grid-column:span 2">
                  <label><b>Select Spreadsheet (.xlsx, .xls, .csv)</b></label>
                  <input type="file" name="excel_file" accept=".xlsx, .xls, .csv" required style="padding:10px; background:#fff; border:1px solid #cbd5e1; border-radius:6px; width:100%;">
                </div>
                <div class="field" style="grid-column:span 2">
                  <label><b>Import Mode</b></label>
                  <div style="display:flex; gap:24px; margin-top:6px; flex-wrap:wrap;">
                    <label style="display:flex; align-items:center; gap:8px; cursor:pointer; font-size:13px;">
                      <input type="radio" name="mode" value="upsert" checked>
                      <span><b>Update & Append</b> (Matches existing materials by code or name to update stock, preserves entered rates, and adds any new materials)</span>
                    </label>
                    <label style="display:flex; align-items:center; gap:8px; cursor:pointer; font-size:13px;">
                      <input type="radio" name="mode" value="replace">
                      <span style="color:#b91c1c;"><b>Full Replace</b> (Wipes stock directory and loads fresh from this spreadsheet)</span>
                    </label>
                  </div>
                </div>
              </div>
              <div style="margin-top:18px; display:flex; gap:10px;">
                <button type="submit" class="btn btn-primary">📤 Upload & Map Spreadsheet</button>
                <a class="btn btn-outline" href="/items">Cancel</a>
              </div>
            </form>
          </div>
        </div>
        """

    # add / edit form
    form_html = ""
    if (qs.get("new") or [""])[0] == "1" or qs.get("edit"):
        edit_id = (qs.get("edit") or [""])[0]
        it = {}
        if edit_id:
            it = _one(conn, "SELECT * FROM items WHERE id=?", (edit_id,)) or {}
        action = "/items/save"
        title = "Edit Material" if edit_id else "Add New Material"
        cat_opts = ""
        for c in cats:
            sel = " selected" if c == it.get("category") else ""
            cat_opts += '<option value="{}"{}>{}</option>'.format(_esc(c), sel, _esc(c))
        form_html = """
        <div class="card" style="margin-bottom:18px">
          <div class="card-head"><h3>{title}</h3></div>
          <div class="card-body">
            <form method="post" action="{action}">
              <input type="hidden" name="id" value="{iid}">
              <div class="form-grid">
                <div class="field"><label>Material Code</label>
                  <input type="text" name="code" value="{code}" required></div>
                <div class="field"><label>Material Name</label>
                  <input type="text" name="name" value="{name}" required></div>
                <div class="field"><label>Category</label>
                  <select name="category">{cat_opts}</select></div>
                <div class="field"><label>Unit</label>
                  <input type="text" name="unit" value="{unit}" placeholder="e.g. Bags, Tons"></div>
                <div class="field"><label>Opening Qty</label>
                  <input type="number" step="any" min="0" name="opening_qty" value="{opening}"></div>
                <div class="field"><label>Rate (&#8377;)</label>
                  <input type="number" step="any" min="0" name="rate" value="{rate}"></div>
                <div class="field"><label>Alert Qty</label>
                  <input type="number" step="any" min="0" name="alert_qty" value="{alert}"></div>
              </div>
              <div style="margin-top:16px;display:flex;gap:10px">
                <button class="btn btn-primary">Save Material</button>
                <a class="btn btn-outline" href="/items">Cancel</a>
              </div>
            </form>
          </div>
        </div>
        """.format(
            title=_esc(title), action=action, iid=it.get("id") or "",
            code=_esc(it.get("code") or ""), name=_esc(it.get("name") or ""),
            unit=_esc(it.get("unit") or ""), cat_opts=cat_opts,
            opening=_esc(it.get("opening_qty") or ""), rate=_esc(it.get("rate") or ""),
            alert=_esc(it.get("alert_qty") or ""),
        )

    # filter bar
    cat_opts = "".join(
        '<option value="{}"{}>{}</option>'.format(_esc(c), " selected" if c == cat else "", _esc(c))
        for c in cats)
    this_year = datetime.now().year
    month_opts = "".join(
        '<option value="{}"{}>{}</option>'.format(mv, " selected" if mv == month else "", ml)
        for mv, ml in MONTHS)
    year_opts = "".join(
        '<option value="{}"{}>{}</option>'.format(y, " selected" if str(y) == year else "", y)
        for y in range(this_year, this_year - 6, -1))
    filter_html = """
    <form class="filter-bar" method="get" action="/items">
      <div class="field"><label>Material</label>
        <input type="text" name="q" value="{q}" placeholder="Search code or name..."></div>
      <div class="field"><label>Category</label>
        <select name="cat"><option value="">All Categories</option>{cat_opts}</select></div>
      <div class="field"><label>Alert Level</label>
        <select name="al">
          <option value="">All Levels</option>
          <option value="low"{sel_low}>Low Stock</option>
          <option value="ok"{sel_ok}>OK</option>
        </select></div>
      <div class="field"><label>Month</label>
        <select name="month"><option value="">Any Month</option>{month_opts}</select></div>
      <div class="field"><label>Year</label>
        <select name="year"><option value="">Any Year</option>{year_opts}</select></div>
      <button class="btn btn-primary">Filter</button>
      <a class="btn btn-outline" href="/items">Clear</a>
    </form>
    """.format(q=_esc(q), cat_opts=cat_opts, sel_low=" selected" if al == "low" else "",
               sel_ok=" selected" if al == "ok" else "", month_opts=month_opts, year_opts=year_opts)

    content = _flash(qs) + _page_head(
        "Stock Directory" + period_label,
        "Master list of materials — opening, purchased, distributed, closing and value.",
        '<a class="btn btn-outline" href="/items?import=1">&#8679; Import Excel</a>'
        '<a class="btn btn-outline" href="/export/items">&#8595; Export Excel</a>'
        '<a class="btn btn-primary" href="/items?new=1">&#43; Add New Material</a>',
    ) + import_html + form_html + filter_html + """
    <div class="card">
      <div class="table-wrap"><table>
        <thead><tr>
          <th>Material Code</th><th>Material Name</th><th>Unit</th>
          <th class="num">Opening</th><th class="num">Purchased</th><th class="num">Distributed</th>
          <th class="num">Closing</th><th class="num">Rate (&#8377;)</th><th class="num">Value (&#8377;)</th>
          <th>Alert</th><th>Actions</th>
        </tr></thead>
        <tbody>{body_rows}</tbody>
      </table></div>
      <div class="pagination"><span>Showing {n} of {n} entries</span>
        <span>Total value: <b>&#8377;{total}</b> &middot; Low stock: <b>{low}</b></span></div>
    </div>
    """.format(body_rows=body_rows, n=len(rows), total=_inr(total_value), low=low_count)

    return _layout("Stock Directory", "items", user, content)


def page_requisitions(conn, qs, user):
    lob = (qs.get("lob") or [""])[0]
    items = {it["id"]: it for it in _q(conn, "SELECT * FROM items")}
    stock = _stock_map(conn)

    sql = "SELECT * FROM requisitions"
    params = ()
    if lob:
        sql += " WHERE lob=?"
        params = (lob,)
    sql += " ORDER BY req_date DESC, id DESC"
    reqs = _q(conn, sql, params)

    delivered = [r for r in reqs if r["status"] == "DELIVERED"]
    pending = [r for r in reqs if r["status"] in ("PENDING", "IN TRANSIT")]
    lobs_requested = {r["lob"] for r in reqs}
    lobs_supplied = {r["lob"] for r in delivered}

    rows = ""
    for r in reqs:
        it = items.get(r["item_id"], {})
        closing = stock.get(r["item_id"], {}).get("closing", 0)
        is_shortfall = r["status"] in ("PENDING", "IN TRANSIT") and r["qty"] > closing

        actions = ""
        if is_shortfall:
            actions = ('<span class="badge badge-rejected"><span class="dot"></span>Qty High</span>')
        else:
            if r["status"] == "PENDING":
                actions += ('<form class="inline-form" method="post" action="/requisitions/action">'
                            '<input type="hidden" name="id" value="{id}"><input type="hidden" name="act" value="transit">'
                            '<button class="btn btn-navy btn-sm">In Transit</button></form> ').format(id=r["id"])
            if r["status"] in ("PENDING", "IN TRANSIT"):
                actions += ('<button type="button" class="btn btn-green btn-sm js-accept-delivery" '
                            'data-id="{id}" data-reqno="{reqno}">Accept Delivery</button> '
                            '<form class="inline-form" method="post" action="/requisitions/action">'
                            '<input type="hidden" name="id" value="{id}"><input type="hidden" name="act" value="reject">'
                            '<button class="btn btn-red btn-sm">Reject</button></form> ').format(
                    id=r["id"], reqno=_esc(r["req_no"]))
            actions += ('<a class="btn btn-outline btn-sm" href="/requisitions?edit={id}">Edit</a> '
                        '<form class="inline-form" method="post" action="/requisitions/delete" '
                        'onsubmit="return confirm(&#39;Delete this requisition?&#39;)">'
                        '<input type="hidden" name="id" value="{id}">'
                        '<button class="btn btn-red btn-sm">Del</button></form>').format(id=r["id"])

        status_html = _status_badge(r["status"])

        rows += ('<tr><td class="mono">{no}</td>'
                 '<td><div style="font-weight:700">{name}</div>'
                 '<div class="mono" style="color:#6b7280;font-size:11px">{code}</div></td>'
                 '<td>{unit}</td><td class="num">{existing}</td>'
                 '<td class="num" style="font-weight:800">{qty}</td>'
                 '<td>{date}</td><td>{status}</td><td>{delivered_on}</td>'
                 '<td style="white-space:nowrap">{actions}</td></tr>').format(
            no=_esc(r["req_no"]), name=_esc(it.get("name") or "—"), code=_esc(it.get("code") or ""),
            unit=_esc(it.get("unit") or "—"), existing=_inr(closing), qty=_inr(r["qty"]),
            date=_esc(r["req_date"]), status=status_html,
            delivered_on=_esc(r["delivery_date"]) if r["delivery_date"] else "—", actions=actions)
    if not rows:
        rows = '<tr><td colspan="9" class="empty">No requisitions found.</td></tr>'

    # add / edit form
    form_html = ""
    if (qs.get("new") or [""])[0] == "1" or qs.get("edit"):
        edit_id = (qs.get("edit") or [""])[0]
        r = {}
        if edit_id:
            r = _one(conn, "SELECT * FROM requisitions WHERE id=?", (edit_id,)) or {}
        item_opts = _item_options(conn, r.get("item_id"))
        lob_opts = _lob_options(r.get("lob"))
        form_html = """
        <div class="card" style="margin-bottom:18px">
          <div class="card-head"><h3>{title}</h3></div>
          <div class="card-body">
            <form method="post" action="/requisitions/save">
              <input type="hidden" name="id" value="{iid}">
              <div class="form-grid">
                <div class="field"><label>Material</label>
                  <select name="item_id" required>{item_opts}</select></div>
                <div class="field"><label>LOB</label>
                  <select name="lob" required>{lob_opts}</select></div>
                <div class="field"><label>Required Qty</label>
                  <input type="number" step="any" min="0" name="qty" value="{qty}" required></div>
                <div class="field"><label>Req Date</label>
                  <input type="date" name="req_date" value="{date}"></div>
                <div class="field"><label>Note</label>
                  <input type="text" name="note" value="{note}"></div>
              </div>
              <div style="margin-top:16px;display:flex;gap:10px">
                <button class="btn btn-primary">Save Requisition</button>
                <a class="btn btn-outline" href="/requisitions">Cancel</a>
              </div>
            </form>
          </div>
        </div>
        """.format(
            title="Edit Requisition" if edit_id else "Add New Requisition",
            iid=r.get("id") or "", item_opts=item_opts, lob_opts=lob_opts,
            qty=_esc(r.get("qty") or ""), date=_esc(r.get("req_date") or _today()),
            note=_esc(r.get("note") or ""),
        )

    lob_opts_all = _lob_options(lob)
    filter_html = """
    <form class="filter-bar" method="get" action="/requisitions">
      <div class="field" style="flex:0 0 260px"><label>Filter by LOB</label>
        <select name="lob"><option value="">All LOBs</option>{lob_opts}</select></div>
      <button class="btn btn-primary">Filter</button>
      <a class="btn btn-outline" href="/requisitions">Clear</a>
      <span style="margin-left:auto;align-self:center;font-size:12px;color:#6b7280">Workflow: Pending &rarr; In Transit &rarr; Delivered / Rejected</span>
    </form>
    """.format(lob_opts=lob_opts_all)

    content = _flash(qs) + _page_head(
        "Requisitions",
        "Received demands from various Lines of Business (LOBs).",
        '<a class="btn btn-outline" href="/export/requisitions">&#8595; Export</a>'
        '<a class="btn btn-primary" href="/requisitions?new=1">&#43; Add New Requisition</a>',
    ) + """
    <div class="kpi-grid">
      <div class="kpi accent"><div class="k-label">Up to Date Requisitions</div>
        <div class="k-val">{delivered}</div><div class="k-sub">delivered</div></div>
      <div class="kpi warn"><div class="k-label">Pending</div>
        <div class="k-val">{pending}</div><div class="k-sub">awaiting approval / transit</div></div>
      <div class="kpi"><div class="k-label">Total LOBs Who Raised</div>
        <div class="k-val">{lobs_req}</div><div class="k-sub">distinct lines of business</div></div>
      <div class="kpi"><div class="k-label">LOBs Supplied</div>
        <div class="k-val">{lobs_sup}</div><div class="k-sub">materials delivered</div></div>
    </div>
    """.format(delivered=len(delivered), pending=len(pending),
               lobs_req=len(lobs_requested), lobs_sup=len(lobs_supplied))

    content += form_html + filter_html + """
    <div class="card">
      <div class="table-wrap"><table>
        <thead><tr>
          <th>Req No</th><th>Material</th><th>Unit</th>
          <th class="num">Existing Qty</th><th class="num">Required Qty</th>
          <th>Req Date</th><th>Status</th><th>Delivered On</th><th>Actions</th>
        </tr></thead>
        <tbody>{rows}</tbody>
      </table></div>
      <div class="pagination"><span>Showing {n} entries</span></div>
    </div>

    <div class="modal-overlay" id="deliverModal">
      <div class="modal-box">
        <h3>Accept Delivery</h3>
        <p id="deliverModalReq">Confirm the delivery date to post the stock-OUT.</p>
        <form method="post" action="/requisitions/action">
          <input type="hidden" name="id" id="deliverModalId" value="">
          <input type="hidden" name="act" value="deliver">
          <div class="field"><label>Delivery Date</label>
            <input type="date" name="delivery_date" id="deliverModalDate" value="{today}" required></div>
          <div class="modal-actions">
            <button type="button" class="btn btn-outline" onclick="closeDeliverModal()">Cancel</button>
            <button type="submit" class="btn btn-green">Accept Delivery</button>
          </div>
        </form>
      </div>
    </div>
    <script>
    function openDeliverModal(id, reqNo) {{
      document.getElementById('deliverModalId').value = id;
      document.getElementById('deliverModalReq').textContent =
        'Requisition ' + reqNo + ' — confirm the delivery date to post the stock-OUT.';
      document.getElementById('deliverModalDate').value = '{today}';
      document.getElementById('deliverModal').classList.add('open');
    }}
    function closeDeliverModal() {{
      document.getElementById('deliverModal').classList.remove('open');
    }}
    document.getElementById('deliverModal').addEventListener('click', function(e) {{
      if (e.target === this) closeDeliverModal();
    }});
    document.querySelectorAll('.js-accept-delivery').forEach(function(btn) {{
      btn.addEventListener('click', function() {{
        openDeliverModal(btn.getAttribute('data-id'), btn.getAttribute('data-reqno'));
      }});
    }});
    </script>
    """.format(rows=rows, n=len(reqs), today=_esc(_today()))

    return _layout("Requisitions", "requisitions", user, content)


def page_procurement(conn, qs, user):
    items = {it["id"]: it for it in _q(conn, "SELECT * FROM items")}
    stock = _stock_map(conn)
    shortfalls = _requisition_shortfalls(conn)

    mat = (qs.get("mat") or [""])[0]
    st = (qs.get("st") or [""])[0]
    today = _today()
    cutoff = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")

    pos = _q(conn, "SELECT * FROM procurement_orders ORDER BY order_date DESC, id DESC")
    rows_data = []
    for po in pos:
        effective = po["status"]
        if effective == "PENDING" and po["order_date"] and po["order_date"] < cutoff:
            effective = "OVERDUE"
        rows_data.append({**po, "effective": effective})

    if mat:
        try:
            mat_id = int(mat)
        except (TypeError, ValueError):
            mat_id = 0
        rows_data = [po for po in rows_data if po["item_id"] == mat_id]
    if st == "pending":
        rows_data = [po for po in rows_data if po["effective"] in ("PENDING", "OVERDUE")]
    elif st == "received":
        rows_data = [po for po in rows_data if po["effective"] == "RECEIVED"]
    elif st == "overdue":
        rows_data = [po for po in rows_data if po["effective"] == "OVERDUE"]

    total_order_value = sum(po["qty"] * po["rate"] for po in rows_data)
    pending_orders = [po for po in rows_data if po["effective"] in ("PENDING", "OVERDUE")]
    overdue = [po for po in rows_data if po["effective"] == "OVERDUE"]
    received = [po for po in rows_data if po["effective"] == "RECEIVED"]
    pct_delivered = round(len(received) / len(rows_data) * 100) if rows_data else 0

    rows = ""
    for po in rows_data:
        it = items.get(po["item_id"], {})
        closing = stock.get(po["item_id"], {}).get("closing", 0)
        actions = ""
        if po["effective"] in ("PENDING", "OVERDUE"):
            actions += ('<form class="inline-form" method="post" action="/procurement/receive" '
                        'onsubmit="return confirm(&#39;Mark received and post the stock-IN?&#39;)">'
                        '<input type="hidden" name="id" value="{id}">'
                        '<button class="btn btn-green btn-sm">Mark Received</button></form> ').format(id=po["id"])
        actions += ('<a class="btn btn-outline btn-sm" href="/procurement?edit={id}">Edit</a> '
                    '<form class="inline-form" method="post" action="/procurement/delete" '
                    'onsubmit="return confirm(&#39;Delete this purchase order?&#39;)">'
                    '<input type="hidden" name="id" value="{id}">'
                    '<button class="btn btn-red btn-sm">Del</button></form>').format(id=po["id"])

        rows += ('<tr><td class="mono">{code}</td>'
                 '<td><div style="font-weight:700">{name}</div>'
                 '<div class="mono" style="color:#6b7280;font-size:11px">{po}</div></td>'
                 '<td>{unit}</td><td class="num">{existing}</td>'
                 '<td class="num" style="font-weight:800">{qty}</td>'
                 '<td class="num" style="font-weight:700">&#8377;{value}</td>'
                 '<td>{order_date}</td><td>{receipt}</td><td>{status}</td>'
                 '<td style="white-space:nowrap">{actions}</td></tr>').format(
            code=_esc(it.get("code") or "—"), name=_esc(it.get("name") or "—"),
            po=_esc(po["po_no"]), unit=_esc(it.get("unit") or "—"),
            existing=_inr(closing), qty=_inr(po["qty"]),
            value=_inr(po["qty"] * (po["rate"] or 0)),
            order_date=_esc(po["order_date"]), receipt=_esc(po["receipt_date"] or "—"),
            status=_status_badge(po["effective"]), actions=actions)
    if not rows:
        rows = '<tr><td colspan="10" class="empty">No purchase orders found.</td></tr>'

    # add / edit form
    form_html = ""
    if (qs.get("new") or [""])[0] == "1" or qs.get("edit"):
        edit_id = (qs.get("edit") or [""])[0]
        pre = (qs.get("item_id") or [""])[0]
        pre_qty = (qs.get("qty") or [""])[0]
        po = {}
        if edit_id:
            po = _one(conn, "SELECT * FROM procurement_orders WHERE id=?", (edit_id,)) or {}
        sel_item = po.get("item_id")
        if not sel_item and pre:
            try:
                sel_item = int(pre)
            except (TypeError, ValueError):
                sel_item = None
        sel_qty = po.get("qty")
        if not sel_qty and pre_qty:
            sel_qty = pre_qty
        item_opts = _item_options(conn, sel_item)
        form_html = """
        <div class="card" style="margin-bottom:18px">
          <div class="card-head"><h3>{title}</h3></div>
          <div class="card-body">
            <form method="post" action="/procurement/save">
              <input type="hidden" name="id" value="{iid}">
              <div class="form-grid">
                <div class="field"><label>Material</label>
                  <select name="item_id" required>{item_opts}</select></div>
                <div class="field"><label>Order Qty</label>
                  <input type="number" step="any" min="0" name="qty" value="{qty}" required></div>
                <div class="field"><label>Rate (&#8377;)</label>
                  <input type="number" step="any" min="0" name="rate" value="{rate}"></div>
                <div class="field"><label>Supplier</label>
                  <input type="text" name="supplier" value="{supplier}"></div>
                <div class="field"><label>Order Date</label>
                  <input type="date" name="order_date" value="{order_date}"></div>
              </div>
              <div style="margin-top:16px;display:flex;gap:10px">
                <button class="btn btn-primary">Save Purchase Order</button>
                <a class="btn btn-outline" href="/procurement">Cancel</a>
              </div>
            </form>
          </div>
        </div>
        """.format(
            title="Edit Purchase Order" if edit_id else "New Procurement Entry",
            iid=po.get("id") or "", item_opts=item_opts,
            qty=_esc(sel_qty or ""), rate=_esc(po.get("rate") or ""),
            supplier=_esc(po.get("supplier") or ""),
            order_date=_esc(po.get("order_date") or _today()),
        )

    mat_opts = "".join(
        '<option value="{}"{}>{}</option>'.format(
            it["id"], " selected" if mat == str(it["id"]) else "",
            _esc(it["code"] + " — " + it["name"]))
        for it in sorted(items.values(), key=lambda x: x["code"]))
    filter_html = """
    <form class="filter-bar" method="get" action="/procurement">
      <div class="field"><label>Material</label>
        <select name="mat"><option value="">All Materials</option>{mat_opts}</select></div>
      <div class="field"><label>Status</label>
        <select name="st">
          <option value="">All Statuses</option>
          <option value="pending"{s_pending}>Pending / Overdue</option>
          <option value="received"{s_received}>Received</option>
          <option value="overdue"{s_overdue}>Overdue</option>
        </select></div>
      <button class="btn btn-primary">Filter</button>
      <a class="btn btn-outline" href="/procurement">Clear</a>
    </form>
    """.format(mat_opts=mat_opts,
               s_pending=" selected" if st == "pending" else "",
               s_received=" selected" if st == "received" else "",
               s_overdue=" selected" if st == "overdue" else "")

    shortfall_html = ""
    if shortfalls:
        sf_rows = "".join(
            ('<tr><td class="mono">{code}</td>'
             '<td style="font-weight:700">{name}</td>'
             '<td class="num">{demand}</td>'
             '<td class="num">{closing}</td>'
             '<td class="num" style="font-weight:800;color:#93000a">{short}</td>'
             '<td><a class="btn btn-amber btn-sm" href="/procurement?new=1&item_id={iid}&qty={qty}">'
             '&#43; Raise PO</a></td></tr>').format(
                code=_esc(s["code"]), name=_esc(s["name"]),
                demand=_inr(s["demand"]) + " " + _esc(s["unit"]),
                closing=_inr(s["closing"]) + " " + _esc(s["unit"]),
                short=_inr(s["shortfall"]) + " " + _esc(s["unit"]),
                iid=s["item_id"], qty=_num_raw(s["demand"]))
            for s in shortfalls
        )
        shortfall_html = """
        <div class="card" style="margin-bottom:18px;border-color:#ffb4ab">
          <div class="card-head"><h3 style="color:#93000a">&#9888; Quantity High — Requisition Demand Exceeds Stock</h3></div>
          <div class="card-body">
            <p style="font-size:12.5px;color:#6b7280;margin:6px 0 10px">
              These materials have open requisitions asking for more than what's currently in stock.
              Raise a purchase order to cover the shortfall.</p>
            <div class="table-wrap"><table>
              <thead><tr>
                <th>Material Code</th><th>Material</th>
                <th class="num">Requested (Open Req.)</th>
                <th class="num">Available Stock</th>
                <th class="num">Shortfall</th><th></th>
              </tr></thead>
              <tbody>{rows}</tbody>
            </table></div>
          </div>
        </div>
        """.format(rows=sf_rows)

    content = _flash(qs) + _page_head(
        "Material Procurement Tracking",
        "Purchase orders, receipts and delivery status.",
        '<a class="btn btn-primary" href="/procurement?new=1">&#43; New Procurement Entry</a>',
    ) + shortfall_html + """
    <div class="kpi-grid">
      <div class="kpi accent"><div class="k-label">Total Order Value</div>
        <div class="k-val">&#8377;{val}</div><div class="k-sub">for visible orders</div></div>
      <div class="kpi warn"><div class="k-label">Pending Delivery</div>
        <div class="k-val">{pending}</div><div class="k-sub">not yet received</div></div>
      <div class="kpi danger"><div class="k-label">Delayed Materials</div>
        <div class="k-val">{overdue}</div><div class="k-sub">over 30 days open</div></div>
      <div class="kpi"><div class="k-label">Orders Delivered</div>
        <div class="k-val">{pct}%</div><div class="k-sub">{received} of {total} received</div></div>
    </div>
    """.format(val=_inr(total_order_value), pending=len(pending_orders),
               overdue=len(overdue), pct=pct_delivered,
               received=len(received), total=len(rows_data))

    content += form_html + filter_html + """
    <div class="card">
      <div class="table-wrap"><table>
        <thead><tr>
          <th>Material Code</th><th>Material</th><th>Unit</th>
          <th class="num">Existing Qty</th><th class="num">Order Qty</th>
          <th class="num">Order Value</th><th>Order Date</th><th>Receipt Date</th>
          <th>Status</th><th>Actions</th>
        </tr></thead>
        <tbody>{rows}</tbody>
      </table></div>
      <div class="pagination"><span>Showing {n} entries</span></div>
    </div>
    """.format(rows=rows, n=len(rows_data))

    return _layout("Procurement", "procurement", user, content)


def page_reports(conn, qs, user):
    frm = (qs.get("from") or [""])[0] or "2000-01-01"
    to = (qs.get("to") or [""])[0] or "2100-12-31"

    rows = _q(conn, """
        SELECT r.lob, r.qty, i.code, i.name, i.unit, i.id AS item_id
        FROM requisitions r JOIN items i ON i.id = r.item_id
        WHERE r.status IN ('DELIVERED','IN TRANSIT')
          AND r.req_date >= ? AND r.req_date <= ?
    """, (frm, to))

    matrix = {}          # item_id -> {lob: qty}
    item_meta = {}       # item_id -> {code, name, unit}
    lob_totals = {lob: 0.0 for lob in LOBS}
    for r in rows:
        matrix.setdefault(r["item_id"], {})
        matrix[r["item_id"]][r["lob"]] = matrix[r["item_id"]].get(r["lob"], 0) + r["qty"]
        lob_totals[r["lob"]] = lob_totals.get(r["lob"], 0) + r["qty"]
        item_meta[r["item_id"]] = {"code": r["code"], "name": r["name"], "unit": r["unit"]}

    body = ""
    grand = 0.0
    for iid in matrix:
        row = ""
        row_total = 0.0
        for lob in LOBS:
            v = matrix[iid].get(lob, 0)
            row_total += v
            grand += v
            row += '<td class="num">{}</td>'.format(_inr(v) if v else "—")
        m = item_meta[iid]
        body += ('<tr><td class="mono">{code}</td><td style="font-weight:700">{name}</td>'
                 '<td>{unit}</td>{cells}'
                 '<td class="num" style="font-weight:800">{row_total}</td></tr>').format(
            code=_esc(m["code"]), name=_esc(m["name"]), unit=_esc(m["unit"] or "—"),
            cells=row, row_total=_inr(row_total))
    if not body:
        body = '<tr><td colspan="{}" class="empty">No distributions in this period.</td></tr>'.format(
            len(LOBS) + 4)

    foot = "".join('<td class="num">{}</td>'.format(_inr(lob_totals.get(lob, 0))) for lob in LOBS)
    head_cells = "".join('<th class="num">{}</th>'.format(_esc(lob)) for lob in LOBS)

    # delivery log — requisitions whose accepted delivery date falls in range
    delivered_rows = _q(conn, """
        SELECT r.req_no, r.lob, r.qty, r.delivery_date, i.code, i.name, i.unit
        FROM requisitions r JOIN items i ON i.id = r.item_id
        WHERE r.status = 'DELIVERED' AND r.delivery_date IS NOT NULL
          AND r.delivery_date >= ? AND r.delivery_date <= ?
        ORDER BY r.delivery_date DESC, r.id DESC
    """, (frm, to))
    delivery_rows_html = "".join(
        ('<tr><td class="mono">{no}</td><td><div style="font-weight:700">{name}</div>'
         '<div class="mono" style="color:#6b7280;font-size:11px">{code}</div></td>'
         '<td>{lob}</td><td class="num">{qty}</td><td>{unit}</td><td>{ddate}</td></tr>').format(
            no=_esc(dr["req_no"]), name=_esc(dr["name"]), code=_esc(dr["code"]),
            lob=_esc(dr["lob"]), qty=_inr(dr["qty"]), unit=_esc(dr["unit"] or "—"),
            ddate=_esc(dr["delivery_date"]))
        for dr in delivered_rows
    ) or '<tr><td colspan="6" class="empty">No accepted deliveries in this period.</td></tr>'

    content = _flash(qs) + _page_head(
        "Material Quantity Distribution",
        "Material utilization across Lines of Business (LOB) — from delivered & in-transit requisitions.",
        '<a class="btn btn-outline" href="/export/reports?from={}&amp;to={}">&#8595; Export Excel</a>'.format(
            _esc(frm), _esc(to)),
    ) + """
    <form class="filter-bar" method="get" action="/reports">
      <div class="field" style="flex:0 0 190px"><label>From Date</label>
        <input type="date" name="from" value="{frm}"></div>
      <div class="field" style="flex:0 0 190px"><label>To Date</label>
        <input type="date" name="to" value="{to}"></div>
      <button class="btn btn-primary">Apply Filters</button>
      <a class="btn btn-outline" href="/reports">Reset</a>
    </form>
    <div class="card">
      <div class="table-wrap"><table class="matrix">
        <thead><tr>
          <th>Material Code</th><th>Material</th><th>Unit</th>
          {head_cells}<th class="num">Total</th>
        </tr></thead>
        <tbody>{body}</tbody>
        <tfoot class="tfoot">
          <tr><td colspan="3" style="text-transform:uppercase;letter-spacing:.05em">Total Consumption</td>
          {foot}<td class="num">{grand}</td></tr>
        </tfoot>
      </table></div>
      <div class="pagination"><span>{n} materials with distribution in period</span></div>
    </div>

    <div class="card" style="margin-top:18px">
      <div class="card-head"><h3>Delivery Log</h3>
        <span style="font-size:12px;color:#6b7280">Requisitions with an accepted delivery date</span></div>
      <div class="table-wrap"><table>
        <thead><tr>
          <th>Req No</th><th>Material</th><th>LOB</th>
          <th class="num">Qty</th><th>Unit</th><th>Delivery Date</th>
        </tr></thead>
        <tbody>{delivery_rows}</tbody>
      </table></div>
      <div class="pagination"><span>{dn} deliveries in period</span></div>
    </div>
    """.format(frm=_esc(frm), to=_esc(to), head_cells=head_cells,
               body=body, foot=foot, grand=_inr(grand), n=len(matrix),
               delivery_rows=delivery_rows_html, dn=len(delivered_rows))

    return _layout("Reports", "reports", user, content)


def page_movements(conn, qs, user):
    items = {it["id"]: it for it in _q(conn, "SELECT id, code, name FROM items")}

    moves = _q(conn, """
        SELECT m.* FROM stock_movements m
        ORDER BY m.id DESC LIMIT 200
    """)

    rows = ""
    for m in moves:
        it = items.get(m["item_id"], {})
        arrow = '<span style="color:#065f46;font-weight:800">&#8593; IN</span>' if m["direction"] == "IN" else \
                '<span style="color:#93000a;font-weight:800">&#8595; OUT</span>'
        rows += ('<tr><td class="mono">{id}</td><td>{date}</td>'
                 '<td><span class="mono">{code}</span> {name}</td>'
                 '<td>{arrow}</td><td class="num" style="font-weight:800">{qty}</td>'
                 '<td class="mono">{ref}</td><td>{note}</td><td>{by}</td></tr>').format(
            id=m["id"], date=_esc(m["m_date"]), code=_esc(it.get("code") or "—"),
            name=_esc(it.get("name") or ""), arrow=arrow, qty=_inr(m["qty"]),
            ref=_esc((m["ref_type"] or "") + " " + (m["ref_no"] or "")),
            note=_esc(m["note"] or "—"), by=_esc(m["created_by"] or "—"))
    if not rows:
        rows = '<tr><td colspan="8" class="empty">No stock movements recorded.</td></tr>'

    item_opts = _item_options(conn)
    content = _flash(qs) + _page_head(
        "Stock Movements",
        "Every stock-IN and stock-OUT with its source document.",
        '<a class="btn btn-outline" href="/export/movements">&#8595; Export</a>',
    ) + """
    <div class="card" style="margin-bottom:18px">
      <div class="card-head"><h3>Add Manual Movement</h3></div>
      <div class="card-body">
        <form method="post" action="/movements/save">
          <div class="form-grid">
            <div class="field"><label>Material</label>
              <select name="item_id" required>{item_opts}</select></div>
            <div class="field"><label>Direction</label>
              <select name="direction" required>
                <option value="IN">&#8593; Stock IN (purchase / return)</option>
                <option value="OUT">&#8595; Stock OUT (distribution / loss)</option>
              </select></div>
            <div class="field"><label>Quantity</label>
              <input type="number" step="any" min="0" name="qty" required></div>
            <div class="field"><label>Date</label>
              <input type="date" name="m_date" value="{today}"></div>
            <div class="field"><label>Note</label>
              <input type="text" name="note" placeholder="e.g. Site transfer, damaged stock"></div>
          </div>
          <div style="margin-top:16px"><button class="btn btn-primary">Post Movement</button></div>
        </form>
      </div>
    </div>
    <div class="card">
      <div class="table-wrap"><table>
        <thead><tr><th>#</th><th>Date</th><th>Material</th><th>Direction</th>
          <th class="num">Qty</th><th>Reference</th><th>Note</th><th>By</th></tr></thead>
        <tbody>{rows}</tbody>
      </table></div>
      <div class="pagination"><span>Latest 200 movements</span></div>
    </div>
    """.format(item_opts=item_opts, today=_today(), rows=rows)

    return _layout("Movements", "movements", user, content)


def page_users(conn, qs, user):
    is_admin = (user.get("role") or "").lower() == "superadmin"
    users = _q(conn, "SELECT * FROM users ORDER BY id")

    body_rows = ""
    for u in users:
        del_btn = ""
        if is_admin and u["username"] != user.get("username"):
            del_btn = (
                '<form class="inline-form" method="post" action="/users/delete" '
                'onsubmit="return confirm(&#39;Delete this user?&#39;)">'
                '<input type="hidden" name="id" value="{id}">'
                '<button class="btn btn-red btn-sm">Delete</button></form>'
            ).format(id=u["id"])
        body_rows += (
            '<tr>'
            '<td class="mono">{uname}</td>'
            '<td style="font-weight:700">{dname}</td>'
            '<td><span class="pill">{role}</span></td>'
            '<td>{created}</td>'
            '<td class="num">{actions}</td>'
            '</tr>'
        ).format(
            uname=_esc(u["username"]), dname=_esc(u["display_name"] or ""),
            role=_esc((u["role"] or "user").capitalize()),
            created=_esc(u["created_at"] or ""), actions=del_btn,
        )
    if not body_rows:
        body_rows = '<tr><td colspan="5" class="empty">No user accounts yet.</td></tr>'

    form_html = ""
    if is_admin:
        form_html = """
        <div class="card" style="margin-bottom:18px">
          <div class="card-head"><h3>Add User</h3></div>
          <div class="card-body">
            <form method="post" action="/users/save">
              <div class="form-grid">
                <div class="field"><label>Username</label>
                  <input type="text" name="username" required></div>
                <div class="field"><label>Display Name</label>
                  <input type="text" name="display_name" placeholder="Optional"></div>
                <div class="field"><label>Password</label>
                  <input type="password" name="password" required></div>
                <div class="field"><label>Role</label>
                  <select name="role">
                    <option value="user">User</option>
                    <option value="admin">Admin</option>
                    <option value="superadmin">Superadmin</option>
                  </select></div>
              </div>
              <div style="margin-top:16px;display:flex;gap:10px">
                <button class="btn btn-primary">Add User</button>
              </div>
            </form>
          </div>
        </div>
        """
    else:
        form_html = '<div class="alert err">&#9888; Only superadmin users can add or remove accounts.</div>'

    content = _flash(qs) + _page_head(
        "User Accounts",
        "Manage who can sign in — username, password and role.",
    ) + form_html + """
    <div class="card">
      <div class="table-wrap"><table>
        <thead><tr><th>Username</th><th>Display Name</th><th>Role</th><th>Created</th><th>Actions</th></tr></thead>
        <tbody>{body_rows}</tbody>
      </table></div>
    </div>
    """.format(body_rows=body_rows)

    return _layout("User Accounts", "users", user, content)


# ─────────────────────────────────────────────────────────────
#  EXPORTS (CSV — opens in Excel)
# ─────────────────────────────────────────────────────────────

def _csv_bytes(header, rows):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(header)
    for r in rows:
        w.writerow(r)
    return ("\ufeff" + buf.getvalue()).encode("utf-8")


def export_items(handler):
    conn = _conn()
    items = _q(conn, "SELECT * FROM items ORDER BY code")
    stock = _stock_map(conn)
    rows = []
    for it in items:
        closing = stock[it["id"]]["closing"]
        rows.append([
            it["code"], it["name"], it["category"] or "", it["unit"] or "",
            it["opening_qty"], stock[it["id"]]["purchased"], stock[it["id"]]["distributed"],
            closing, it["rate"], round(closing * (it["rate"] or 0), 2), it["alert_qty"],
        ])
    conn.close()
    handler.send_csv("stock_directory.csv",
                     ["Material Code", "Material Name", "Category", "Unit",
                      "Opening", "Purchased", "Distributed", "Closing",
                      "Rate", "Value", "Alert Qty"], rows)


def export_requisitions(handler):
    conn = _conn()
    rows = _q(conn, """
        SELECT r.*, i.code, i.name, i.unit
        FROM requisitions r JOIN items i ON i.id = r.item_id
        ORDER BY r.id DESC
    """)
    conn.close()
    out = [[r["req_no"], r["code"], r["name"], r["unit"], r["lob"], r["qty"],
            r["req_date"], r["status"], r["delivery_date"] or "", r["note"] or "",
            r["created_by"] or ""] for r in rows]
    handler.send_csv("requisitions.csv",
                     ["Req No", "Material Code", "Material", "Unit", "LOB",
                      "Qty", "Req Date", "Status", "Delivery Date", "Note", "Created By"], out)


def export_reports(handler, qs):
    frm = (qs.get("from") or [""])[0] or "2000-01-01"
    to = (qs.get("to") or [""])[0] or "2100-12-31"
    conn = _conn()
    rows = _q(conn, """
        SELECT r.lob, r.qty, i.code, i.name, i.unit, i.id AS item_id
        FROM requisitions r JOIN items i ON i.id = r.item_id
        WHERE r.status IN ('DELIVERED','IN TRANSIT')
          AND r.req_date >= ? AND r.req_date <= ?
    """, (frm, to))
    conn.close()

    matrix = {}
    item_meta = {}
    for r in rows:
        matrix.setdefault(r["item_id"], {})
        matrix[r["item_id"]][r["lob"]] = matrix[r["item_id"]].get(r["lob"], 0) + r["qty"]
        item_meta[r["item_id"]] = {"code": r["code"], "name": r["name"], "unit": r["unit"]}

    header = ["Material Code", "Material", "Unit"] + LOBS + ["Total"]
    out = []
    for iid in matrix:
        m = item_meta[iid]
        line = [m["code"], m["name"], m["unit"]]
        total = 0.0
        for lob in LOBS:
            v = matrix[iid].get(lob, 0)
            total += v
            line.append(v)
        line.append(total)
        out.append(line)
    handler.send_csv("distribution_report.csv", header, out)


def export_movements(handler):
    conn = _conn()
    rows = _q(conn, """
        SELECT m.*, i.code, i.name FROM stock_movements m
        JOIN items i ON i.id = m.item_id ORDER BY m.id DESC
    """)
    conn.close()
    out = [[m["id"], m["m_date"], m["code"], m["name"], m["direction"], m["qty"],
            (m["ref_type"] or "") + " " + (m["ref_no"] or ""), m["note"] or "",
            m["created_by"] or ""] for m in rows]
    handler.send_csv("stock_movements.csv",
                     ["ID", "Date", "Code", "Material", "Direction", "Qty",
                      "Reference", "Note", "By"], out)


# ─────────────────────────────────────────────────────────────
#  HTTP HANDLER
# ─────────────────────────────────────────────────────────────

class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def setup(self):
        BaseHTTPRequestHandler.setup(self)
        try:
            self.connection.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
        except OSError:
            pass

    def log_message(self, fmt, *args):
        pass

    # ---- response helpers ----

    def send_html(self, body, code=200):
        b = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(b)
        except (ConnectionResetError, BrokenPipeError, OSError):
            pass

    def send_csv(self, fname, header, rows):
        b = _csv_bytes(header, rows)
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self.send_header("Content-Disposition",
                         "attachment; filename=\"" + fname + "\"")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        try:
            self.wfile.write(b)
        except (ConnectionResetError, BrokenPipeError, OSError):
            pass

    def redirect(self, loc):
        self.send_response(302)
        self.send_header("Location", loc)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def json_out(self, obj, code=200):
        b = json.dumps(obj).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        try:
            self.wfile.write(b)
        except (ConnectionResetError, BrokenPipeError, OSError):
            pass

    # ---- session ----

    def session_user(self):
        cookie = self.headers.get("Cookie", "")
        for part in cookie.split(";"):
            p = part.strip()
            if p.startswith(COOKIE_NAME + "="):
                tok = p.split("=", 1)[1].strip()
                s = SESSIONS.get(tok)
                if s and s["expires"] > datetime.now():
                    return s
                SESSIONS.pop(tok, None)
                break
        return None

    def set_session_cookie(self, token):
        self.send_header("Set-Cookie", "{}={}; Path=/; HttpOnly; Max-Age={}".format(
            COOKIE_NAME, token, SESSION_TTL))

    # ---- routing ----

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        if path == "/login":
            if self.session_user():
                self.redirect("/dashboard")
            else:
                self.send_html(page_login())
            return

        if path == "/logout":
            cookie = self.headers.get("Cookie", "")
            for part in cookie.split(";"):
                p = part.strip()
                if p.startswith(COOKIE_NAME + "="):
                    SESSIONS.pop(p.split("=", 1)[1].strip(), None)
            self.send_response(302)
            self.send_header("Location", "/login")
            self.send_header("Set-Cookie", COOKIE_NAME + "=; Path=/; Max-Age=0")
            self.send_header("Content-Length", "0")
            self.end_headers()
            return

        user = self.session_user()
        if not user:
            self.redirect("/login")
            return

        conn = _conn()
        try:
            if path in ("/", "/dashboard"):
                self.send_html(page_dashboard(conn, qs, user))
            elif path == "/items":
                self.send_html(page_items(conn, qs, user))
            elif path == "/requisitions":
                self.send_html(page_requisitions(conn, qs, user))
            elif path == "/procurement":
                self.send_html(page_procurement(conn, qs, user))
            elif path == "/reports":
                self.send_html(page_reports(conn, qs, user))
            elif path == "/movements":
                self.send_html(page_movements(conn, qs, user))
            elif path == "/users":
                self.send_html(page_users(conn, qs, user))
            elif path == "/export/items":
                export_items(self)
            elif path == "/export/requisitions":
                export_requisitions(self)
            elif path == "/export/reports":
                export_reports(self, qs)
            elif path == "/export/movements":
                export_movements(self)
            elif path == "/health":
                self.json_out({"status": "ok"})
            else:
                self.send_html("<h1>404 — Not Found</h1>", 404)
        finally:
            conn.close()

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/login":
            body = self._read_form()
            username = (body.get("username") or "").strip()
            password = body.get("password") or ""
            conn = _conn()
            row = _one(conn, "SELECT * FROM users WHERE username=?", (username,))
            conn.close()
            if row and row["password_hash"] == hashlib.sha256(password.encode()).hexdigest():
                token = secrets.token_hex(24)
                SESSIONS[token] = {
                    "username": row["username"],
                    "display_name": row["display_name"] or row["username"],
                    "role": row["role"] or "user",
                    "expires": datetime.now() + timedelta(seconds=SESSION_TTL),
                }
                self.send_response(302)
                self.send_header("Location", "/dashboard")
                self.set_session_cookie(token)
                self.send_header("Content-Length", "0")
                self.end_headers()
            else:
                self.send_html(page_login("Invalid username or password."))
            return

        user = self.session_user()
        if not user:
            self.redirect("/login")
            return

        body = self._read_form()
        conn = _conn()
        try:
            if path == "/items/save":
                self._save_item(conn, body)
            elif path == "/items/delete":
                self._delete_item(conn, body)
            elif path == "/items/import":
                self._import_items_excel(conn, body)
            elif path == "/requisitions/save":
                self._save_requisition(conn, body)
            elif path == "/requisitions/action":
                self._req_action(conn, body)
            elif path == "/requisitions/delete":
                self._delete_requisition(conn, body)
            elif path == "/procurement/save":
                self._save_procurement(conn, body)
            elif path == "/procurement/receive":
                self._receive_procurement(conn, body)
            elif path == "/procurement/delete":
                self._delete_procurement(conn, body)
            elif path == "/movements/save":
                self._save_movement(conn, body)
            elif path == "/users/save":
                self._save_user(conn, body, user)
            elif path == "/users/delete":
                self._delete_user(conn, body, user)
            else:
                self.send_html("<h1>404 — Not Found</h1>", 404)
        finally:
            conn.close()

    def _read_form(self):
        self.uploaded_files = {}
        try:
            n = int(self.headers.get("Content-Length") or 0)
        except (TypeError, ValueError):
            n = 0
        raw = self.rfile.read(n) if n else b""
        ctype = self.headers.get("Content-Type") or ""
        if "multipart/form-data" in ctype:
            import email
            msg = email.message_from_bytes(f"Content-Type: {ctype}\r\n\r\n".encode("utf-8") + raw)
            fields = {}
            for part in msg.get_payload():
                name = part.get_param("name", header="content-disposition")
                if not name:
                    continue
                filename = part.get_filename()
                if filename:
                    self.uploaded_files[name] = {
                        "filename": filename,
                        "data": part.get_payload(decode=True) or b"",
                        "content_type": part.get_content_type()
                    }
                else:
                    fields[name] = (part.get_payload(decode=True) or b"").decode("utf-8", errors="replace")
            return fields
        return {k: v[0] for k, v in parse_qs(raw.decode("utf-8", errors="replace"),
                                             keep_blank_values=True).items()}

    def _num(self, v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    # ---- item actions ----

    def _save_item(self, conn, b):
        iid = b.get("id") or None
        code = (b.get("code") or "").strip()
        name = (b.get("name") or "").strip()
        if not code or not name:
            self.redirect("/items?new=1&err=" + _urlq("Code and name are required."))
            return
        dup = _one(conn, "SELECT id FROM items WHERE code=? AND (? IS NULL OR id != ?)",
                   (code, iid, iid))
        if dup:
            self.redirect("/items?err=" + _urlq("Material code already exists."))
            return
        if iid:
            conn.execute(
                "UPDATE items SET code=?, name=?, category=?, unit=?, opening_qty=?, rate=?, alert_qty=? WHERE id=?",
                (code, name, (b.get("category") or "").strip(), (b.get("unit") or "").strip(),
                 self._num(b.get("opening_qty")), self._num(b.get("rate")),
                 self._num(b.get("alert_qty")), iid),
            )
        else:
            conn.execute(
                "INSERT INTO items (code,name,category,unit,opening_qty,rate,alert_qty,created_at) VALUES (?,?,?,?,?,?,?,?)",
                (code, name, (b.get("category") or "").strip(), (b.get("unit") or "").strip(),
                 self._num(b.get("opening_qty")), self._num(b.get("rate")),
                 self._num(b.get("alert_qty")), _today()),
            )
        conn.commit()
        self.redirect("/items?ok=" + _urlq("Material saved."))

    def _delete_item(self, conn, b):
        iid = b.get("id")
        conn.execute("DELETE FROM items WHERE id=?", (iid,))
        conn.commit()
        self.redirect("/items?ok=" + _urlq("Material deleted."))

    def _import_items_excel(self, conn, b):
        f = getattr(self, "uploaded_files", {}).get("excel_file")
        if not f or not f.get("data"):
            self.redirect("/items?import=1&err=" + _urlq("Please select a valid Excel or CSV file to upload."))
            return

        filename = f.get("filename", "").lower()
        file_bytes = f["data"]
        mode = (b.get("mode") or "upsert").strip()

        rows = []
        if filename.endswith(".csv"):
            try:
                text_stream = io.StringIO(file_bytes.decode("utf-8", errors="replace"))
                reader = csv.reader(text_stream)
                rows = [row for row in reader if any(str(cell).strip() for cell in row)]
            except Exception as e:
                self.redirect("/items?import=1&err=" + _urlq(f"Error reading CSV file: {e}"))
                return
        else:
            try:
                import openpyxl
                wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
                sheet = wb.active
                for r in range(1, sheet.max_row + 1):
                    row_vals = [sheet.cell(r, c).value for c in range(1, sheet.max_column + 1)]
                    if any(v is not None and str(v).strip() != "" for v in row_vals):
                        rows.append(row_vals)
            except Exception as e:
                self.redirect("/items?import=1&err=" + _urlq(f"Error reading Excel file: {e}"))
                return

        if not rows:
            self.redirect("/items?import=1&err=" + _urlq("The uploaded spreadsheet is empty."))
            return

        # Scan the first 10 rows to detect the header row
        header_row_idx = None
        col_map = None
        for idx in range(min(10, len(rows))):
            cm = find_column_map(rows[idx])
            if "name" in cm:
                header_row_idx = idx
                col_map = cm
                break

        if col_map is None:
            self.redirect("/items?import=1&err=" + _urlq("Could not detect column headers. Please make sure the sheet has a 'Material Name' or 'Item Name' column."))
            return

        data_rows = rows[header_row_idx + 1:]
        if not data_rows:
            self.redirect("/items?import=1&err=" + _urlq("No data rows found below the header row."))
            return

        # Fetch existing items for duplicate checks
        existing_items = _q(conn, "SELECT * FROM items")
        existing_by_code = {it["code"].strip().lower(): it for it in existing_items if it.get("code")}
        existing_by_name = {it["name"].strip().lower(): it for it in existing_items if it.get("name")}

        max_inv_num = 0
        for it in existing_items:
            m = re.match(r"^INV-(\d+)$", it.get("code") or "", re.IGNORECASE)
            if m:
                max_inv_num = max(max_inv_num, int(m.group(1)))

        parsed_items = []
        for r_idx, row in enumerate(data_rows, start=header_row_idx + 2):
            # Name
            name_idx = col_map.get("name")
            raw_name = row[name_idx] if name_idx is not None and len(row) > name_idx else None
            name = str(raw_name).strip() if raw_name is not None else ""
            if not name:
                continue

            # Code
            code = ""
            if "code" in col_map and len(row) > col_map["code"]:
                raw_code = row[col_map["code"]]
                if raw_code is not None:
                    c_str = str(raw_code).strip()
                    if c_str.isdigit():
                        code = f"INV-{int(c_str):03d}"
                    elif c_str:
                        code = c_str

            # Quantity and Unit
            raw_qty = row[col_map["qty"]] if "qty" in col_map and len(row) > col_map["qty"] else None
            qty, parsed_unit = parse_qty_unit(raw_qty)

            unit = ""
            if "unit" in col_map and len(row) > col_map["unit"]:
                raw_unit = row[col_map["unit"]]
                if raw_unit is not None and str(raw_unit).strip():
                    unit = str(raw_unit).strip().capitalize()
            if not unit:
                unit = parsed_unit or "Nos"

            # Rate
            rate = 0.0
            if "rate" in col_map and len(row) > col_map["rate"]:
                raw_rate = row[col_map["rate"]]
                if raw_rate is not None and str(raw_rate).strip() != "":
                    rate = self._num(raw_rate)

            # Category
            category = ""
            if "category" in col_map and len(row) > col_map["category"]:
                raw_cat = row[col_map["category"]]
                if raw_cat is not None and str(raw_cat).strip():
                    category = str(raw_cat).strip()
            if not category:
                category = determine_category(name)

            # Alert Qty
            alert_qty = 0.0
            if "alert_qty" in col_map and len(row) > col_map["alert_qty"]:
                raw_alert = row[col_map["alert_qty"]]
                if raw_alert is not None and str(raw_alert).strip() != "":
                    alert_qty = self._num(raw_alert)

            parsed_items.append({
                "code": code,
                "name": name,
                "category": category,
                "unit": unit,
                "opening_qty": qty,
                "rate": rate,
                "alert_qty": alert_qty
            })

        if not parsed_items:
            self.redirect("/items?import=1&err=" + _urlq("No valid material rows found in the spreadsheet."))
            return

        added_count = 0
        updated_count = 0
        today_str = _today()

        if mode == "replace":
            conn.execute("DELETE FROM procurement_orders;")
            conn.execute("DELETE FROM requisitions;")
            conn.execute("DELETE FROM stock_movements;")
            conn.execute("DELETE FROM items;")
            try:
                conn.execute("ALTER SEQUENCE items_id_seq RESTART WITH 1;")
                conn.execute("ALTER SEQUENCE stock_movements_id_seq RESTART WITH 1;")
                conn.execute("ALTER SEQUENCE requisitions_id_seq RESTART WITH 1;")
                conn.execute("ALTER SEQUENCE procurement_orders_id_seq RESTART WITH 1;")
            except Exception:
                pass

            used_codes = set()
            for idx, it in enumerate(parsed_items, start=1):
                code = it["code"]
                if not code or code in used_codes:
                    code = f"INV-{idx:03d}"
                used_codes.add(code)
                conn.execute(
                    "INSERT INTO items (code, name, category, unit, opening_qty, rate, alert_qty, created_at) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (code, it["name"], it["category"], it["unit"], it["opening_qty"], it["rate"], it["alert_qty"], today_str)
                )
                added_count += 1
            conn.commit()
            self.redirect("/items?ok=" + _urlq(f"Stock Directory replaced! {added_count} materials successfully imported."))
            return

        # Default mode: upsert
        used_codes = set(existing_by_code.keys())
        for it in parsed_items:
            matched = None
            if it["code"] and it["code"].lower() in existing_by_code:
                matched = existing_by_code[it["code"].lower()]
            elif it["name"].strip().lower() in existing_by_name:
                matched = existing_by_name[it["name"].strip().lower()]

            if matched:
                target_id = matched["id"]
                final_rate = it["rate"] if it["rate"] > 0 else (matched.get("rate") or 0.0)
                final_code = it["code"] if it["code"] else matched["code"]
                final_cat = it["category"] if it["category"] else (matched.get("category") or "Stationary")
                final_unit = it["unit"] if it["unit"] else (matched.get("unit") or "Nos")

                conn.execute(
                    "UPDATE items SET code=?, name=?, category=?, unit=?, opening_qty=?, rate=?, alert_qty=? WHERE id=?",
                    (final_code, it["name"], final_cat, final_unit, it["opening_qty"], final_rate, it["alert_qty"], target_id)
                )
                updated_count += 1
            else:
                code = it["code"]
                if not code or code.lower() in used_codes:
                    max_inv_num += 1
                    code = f"INV-{max_inv_num:03d}"
                used_codes.add(code.lower())

                conn.execute(
                    "INSERT INTO items (code, name, category, unit, opening_qty, rate, alert_qty, created_at) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (code, it["name"], it["category"], it["unit"], it["opening_qty"], it["rate"], it["alert_qty"], today_str)
                )
                added_count += 1

        conn.commit()
        self.redirect("/items?ok=" + _urlq(f"Excel import complete: {added_count} new materials added, {updated_count} existing updated."))

    # ---- requisition actions ----

    def _save_requisition(self, conn, b):
        rid = b.get("id") or None
        item_id = b.get("item_id")
        lob = (b.get("lob") or "").strip()
        qty = self._num(b.get("qty"))
        rdate = (b.get("req_date") or _today()).strip()
        note = (b.get("note") or "").strip()
        if not item_id or not lob or qty <= 0:
            self.redirect("/requisitions?err=" + _urlq("Material, LOB and a positive quantity are required."))
            return
        if rid:
            conn.execute(
                "UPDATE requisitions SET item_id=?, lob=?, qty=?, req_date=?, note=? WHERE id=?",
                (item_id, lob, qty, rdate, note, rid),
            )
        else:
            nxt = _one(conn, "SELECT COALESCE(MAX(id),0)+1 AS n FROM requisitions")["n"]
            req_no = "REQ-{}-{:04d}".format(datetime.now().year, 1000 + nxt)
            conn.execute(
                "INSERT INTO requisitions (req_no,item_id,lob,qty,req_date,status,note,created_by,created_at) VALUES (?,?,?,?,?,?,?,?,?)",
                (req_no, item_id, lob, qty, rdate, "PENDING", note,
                 user_display(self), _today()),
            )
        conn.commit()

        item = _one(conn, "SELECT * FROM items WHERE id=?", (item_id,)) or {}
        closing = _stock_map(conn).get(int(item_id), {}).get("closing", 0)
        if qty > closing:
            self.redirect("/requisitions?ok=" + _urlq(
                "Requisition saved. Note: requested {} {} exceeds available stock of {} {} — "
                "flagged on the Procurement page for purchase.".format(
                    _inr(qty), item.get("unit") or "units",
                    _inr(closing), item.get("unit") or "units")))
            return
        self.redirect("/requisitions?ok=" + _urlq("Requisition saved."))

    def _req_action(self, conn, b):
        rid = b.get("id")
        act = b.get("act") or ""
        r = _one(conn, "SELECT * FROM requisitions WHERE id=?", (rid,))
        if not r:
            self.redirect("/requisitions?err=" + _urlq("Requisition not found."))
            return
        item = _one(conn, "SELECT * FROM items WHERE id=?", (r["item_id"],))
        stock = _stock_map(conn)

        if act == "transit" and r["status"] == "PENDING":
            conn.execute("UPDATE requisitions SET status='IN TRANSIT' WHERE id=?", (rid,))
            conn.commit()
            self.redirect("/requisitions?ok=" + _urlq("Requisition moved to IN TRANSIT."))
        elif act == "deliver" and r["status"] in ("PENDING", "IN TRANSIT"):
            delivery_date = (b.get("delivery_date") or _today()).strip() or _today()
            closing = stock.get(r["item_id"], {}).get("closing", 0)
            if closing < r["qty"]:
                self.redirect("/requisitions?err=" + _urlq(
                    "Insufficient stock: only {} {} available, {} required.".format(
                        _inr(closing), item["unit"] or "units", _inr(r["qty"]))))
                return
            conn.execute(
                "INSERT INTO stock_movements (item_id,qty,direction,ref_type,ref_no,m_date,note,created_by,created_at) VALUES (?,?,?,?,?,?,?,?,?)",
                (r["item_id"], r["qty"], "OUT", "REQ", r["req_no"], delivery_date,
                 "Delivered to " + r["lob"], user_display(self), _today()),
            )
            conn.execute("UPDATE requisitions SET status='DELIVERED', delivery_date=? WHERE id=?",
                         (delivery_date, rid))
            conn.commit()
            self.redirect("/requisitions?ok=" + _urlq(
                "Delivered {} {} on {} — stock-OUT posted.".format(
                    _inr(r["qty"]), item["unit"] or "units", delivery_date)))
        elif act == "reject" and r["status"] in ("PENDING", "IN TRANSIT"):
            conn.execute("UPDATE requisitions SET status='REJECTED' WHERE id=?", (rid,))
            conn.commit()
            self.redirect("/requisitions?ok=" + _urlq("Requisition rejected."))
        else:
            self.redirect("/requisitions?err=" + _urlq("Action not allowed for this status."))

    def _delete_requisition(self, conn, b):
        rid = b.get("id")
        conn.execute("DELETE FROM requisitions WHERE id=?", (rid,))
        conn.commit()
        self.redirect("/requisitions?ok=" + _urlq("Requisition deleted."))

    # ---- procurement actions ----

    def _save_procurement(self, conn, b):
        pid = b.get("id") or None
        item_id = b.get("item_id")
        qty = self._num(b.get("qty"))
        rate = self._num(b.get("rate"))
        supplier = (b.get("supplier") or "").strip()
        odate = (b.get("order_date") or _today()).strip()
        if not item_id or qty <= 0:
            self.redirect("/procurement?err=" + _urlq("Material and a positive quantity are required."))
            return
        if pid:
            conn.execute(
                "UPDATE procurement_orders SET item_id=?, qty=?, rate=?, supplier=?, order_date=? WHERE id=?",
                (item_id, qty, rate, supplier, odate, pid),
            )
        else:
            nxt = _one(conn, "SELECT COALESCE(MAX(id),0)+1 AS n FROM procurement_orders")["n"]
            po_no = "PO-{}-{:04d}".format(datetime.now().year, 1000 + nxt)
            conn.execute(
                "INSERT INTO procurement_orders (po_no,item_id,qty,rate,supplier,order_date,status,created_by,created_at) VALUES (?,?,?,?,?,?,?,?,?)",
                (po_no, item_id, qty, rate, supplier, odate, "PENDING",
                 user_display(self), _today()),
            )
        conn.commit()
        self.redirect("/procurement?ok=" + _urlq("Purchase order saved."))

    def _receive_procurement(self, conn, b):
        pid = b.get("id")
        po = _one(conn, "SELECT * FROM procurement_orders WHERE id=?", (pid,))
        if not po:
            self.redirect("/procurement?err=" + _urlq("Purchase order not found."))
            return
        if po["status"] == "RECEIVED":
            self.redirect("/procurement?err=" + _urlq("This order was already received."))
            return
        item = _one(conn, "SELECT * FROM items WHERE id=?", (po["item_id"],))
        conn.execute(
            "INSERT INTO stock_movements (item_id,qty,direction,ref_type,ref_no,m_date,note,created_by,created_at) VALUES (?,?,?,?,?,?,?,?,?)",
            (po["item_id"], po["qty"], "IN", "PO", po["po_no"], _today(),
             "PO received from " + (po["supplier"] or "vendor"), user_display(self), _today()),
        )
        conn.execute("UPDATE procurement_orders SET status='RECEIVED', receipt_date=? WHERE id=?",
                     (_today(), pid))
        conn.commit()
        self.redirect("/procurement?ok=" + _urlq(
            "Received {} {} — stock-IN posted.".format(_inr(po["qty"]), item["unit"] or "units")))

    def _delete_procurement(self, conn, b):
        pid = b.get("id")
        conn.execute("DELETE FROM procurement_orders WHERE id=?", (pid,))
        conn.commit()
        self.redirect("/procurement?ok=" + _urlq("Purchase order deleted."))

    # ---- movement actions ----

    def _save_movement(self, conn, b):
        item_id = b.get("item_id")
        direction = (b.get("direction") or "").upper()
        qty = self._num(b.get("qty"))
        mdate = (b.get("m_date") or _today()).strip()
        note = (b.get("note") or "").strip()
        if not item_id or direction not in ("IN", "OUT") or qty <= 0:
            self.redirect("/movements?err=" + _urlq("Material, direction and a positive quantity are required."))
            return
        conn.execute(
            "INSERT INTO stock_movements (item_id,qty,direction,ref_type,ref_no,m_date,note,created_by,created_at) VALUES (?,?,?,?,?,?,?,?,?)",
            (item_id, qty, direction, "MANUAL", "", mdate, note, user_display(self), _today()),
        )
        conn.commit()
        self.redirect("/movements?ok=" + _urlq("Movement posted — stock updated."))

    # ---- user account actions ----

    def _save_user(self, conn, b, actor):
        if (actor.get("role") or "").lower() != "superadmin":
            self.redirect("/users?err=" + _urlq("Only superadmin users can add accounts."))
            return
        username = (b.get("username") or "").strip()
        password = b.get("password") or ""
        display_name = (b.get("display_name") or "").strip() or username
        role = (b.get("role") or "user").strip().lower()
        if role not in ("user", "admin", "superadmin"):
            role = "user"
        if not username or not password:
            self.redirect("/users?err=" + _urlq("Username and password are required."))
            return
        if _one(conn, "SELECT id FROM users WHERE username=?", (username,)):
            self.redirect("/users?err=" + _urlq("That username already exists."))
            return
        conn.execute(
            "INSERT INTO users (username,password_hash,display_name,role,created_at) VALUES (?,?,?,?,?)",
            (username, hashlib.sha256(password.encode()).hexdigest(), display_name, role, _today()),
        )
        conn.commit()
        self.redirect("/users?ok=" + _urlq("User \"{}\" added.".format(username)))

    def _delete_user(self, conn, b, actor):
        if (actor.get("role") or "").lower() != "superadmin":
            self.redirect("/users?err=" + _urlq("Only superadmin users can remove accounts."))
            return
        uid = b.get("id")
        row = _one(conn, "SELECT * FROM users WHERE id=?", (uid,))
        if not row:
            self.redirect("/users?err=" + _urlq("User not found."))
            return
        if row["username"] == actor.get("username"):
            self.redirect("/users?err=" + _urlq("You can't delete your own account."))
            return
        conn.execute("DELETE FROM users WHERE id=?", (uid,))
        conn.commit()
        self.redirect("/users?ok=" + _urlq("User \"{}\" removed.".format(row["username"])))


def user_display(handler):
    s = handler.session_user()
    return s.get("display_name") if s else "Admin"


def _urlq(s):
    from urllib.parse import quote
    return quote(str(s))


# ─────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────

def main():
    init_db()
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print("=" * 60)
    print("  Monarch Surveyors — Inventory Management System")
    print("  " + "=" * 56)
    print("  URL      : http://127.0.0.1:{}".format(PORT))
    print("  LAN      : http://<your-ip>:{}".format(PORT))
    print("  Login    : admin / monarch123")
    print("  Database : {}@{}/{} (schema: {})".format(DB_USER, DB_HOST, DB_NAME, SCHEMA))
    print("  Ctrl+C to stop")
    print("=" * 60)

    def _open():
        try:
            webbrowser.open("http://127.0.0.1:{}".format(PORT))
        except Exception:
            pass
    threading.Timer(0.8, _open).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    finally:
        server.server_close()


# ═══════════════════════════════════════════════════════════════
# SERVER BOOT — LAN ONLY
# ═══════════════════════════════════════════════════════════════
class ThreadedServer(socketserver.ThreadingTCPServer):
    allow_reuse_address = True

def get_lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def run_server(port=9882):

    try:
        init_db()
    except Exception:
        traceback.print_exc()
        return

    ip = get_lan_ip()
    lan_url = f"http://{ip}:{port}"

    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(base_dir)
        os.makedirs("exports", exist_ok=True)
    except Exception:
        pass

    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    print("\n" + "=" * 70)
    print(">> MONARCH SURVEYORS PMS v4.1")
    print("=" * 70)
    print(f"Login URL : {lan_url}/login")
    print(f"Dashboard : {lan_url}/dashboard")
    print(f"Database  : {DB_USER}@{DB_HOST}:{DB_PORT}/{DB_NAME}")
    print(f"Exports   : {os.path.abspath('exports')}")
    print("=" * 70)
    print("Access from any PC on the same network")
    print("Press CTRL+C to stop server")
    print("=" * 70 + "\n")

    try:
        with ThreadedServer(("0.0.0.0", port), Handler) as httpd:

            try:
                webbrowser.open(lan_url + "/login")
            except Exception:
                pass

            httpd.serve_forever()

    except KeyboardInterrupt:
        print("\nServer stopped.")

    except Exception:
        traceback.print_exc()


if __name__ == "__main__":
    run_server(9881)