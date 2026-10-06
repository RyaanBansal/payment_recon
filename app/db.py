"""SQLite helpers + schema. All data (sheets, match results, agent/remarks) lives in data/recon.db."""
import os
import sqlite3
import json
from datetime import datetime

DB_PATH = os.environ.get(
    "RECON_DB", os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "recon.db")
)

SCHEMA = """
CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT);
-- Agent names list: empty for now. Add rows here (or POST /api/agents) and the dropdown fills automatically.
CREATE TABLE IF NOT EXISTS agents(name TEXT PRIMARY KEY);
-- Manual inputs, keyed by Zoho/Razorpay payment id so they survive re-uploads and re-matching.
CREATE TABLE IF NOT EXISTS annotations(zoho_id TEXT PRIMARY KEY, agent TEXT, remarks TEXT);
CREATE TABLE IF NOT EXISTS match_results(
    zoho_id TEXT PRIMARY KEY, result TEXT, matched_by TEXT,
    payment_amount REAL, difference REAL, basis TEXT, invoice_number TEXT, exchange_rate REAL, pay_currency TEXT
);
-- Match results of the other monthly sheets (Kotak, Cash Free ...): one row per sheet row, keyed by (sheet, row key).
-- Agent/remarks for these rows live in `annotations`, with zoho_id = '<sheet>:<row key>'.
CREATE TABLE IF NOT EXISTS src_match_results(
    source TEXT, row_key TEXT, result TEXT, matched_by TEXT,
    payment_amount REAL, difference REAL, basis TEXT, invoice_number TEXT,
    PRIMARY KEY (source, row_key)
);
"""


def q(name: str) -> str:
    """Quote an identifier."""
    return '"' + name.replace('"', '""') + '"'


def conn() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def init_db() -> None:
    with conn() as c:
        c.executescript(SCHEMA)
        # migration for databases created before invoice_number existed
        cols = {r["name"] for r in c.execute("PRAGMA table_info(match_results)")}
        if "invoice_number" not in cols:
            c.execute("ALTER TABLE match_results ADD COLUMN invoice_number TEXT")
        if "exchange_rate" not in cols:  # migration: exchange rate of the matched customer payment(s)
            c.execute("ALTER TABLE match_results ADD COLUMN exchange_rate REAL")
        if "pay_currency" not in cols:  # migration: currency of the matched customer payment (Zoho sheet)
            c.execute("ALTER TABLE match_results ADD COLUMN pay_currency TEXT")


def now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def get_meta(c, key, default=None):
    r = c.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return json.loads(r["value"]) if r else default


def set_meta(c, key, value) -> None:
    c.execute("INSERT OR REPLACE INTO meta(key,value) VALUES(?,?)", (key, json.dumps(value)))


def table_exists(c, name: str) -> bool:
    return c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)).fetchone() is not None