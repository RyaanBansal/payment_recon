"""FastAPI app: serves the single-page UI + JSON API."""
import os
import re
from io import BytesIO
from datetime import datetime
from typing import List, Optional

from fastapi import FastAPI, File, HTTPException, Query, Request, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from openpyxl import Workbook
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import ingest, ingest_sheets, match_cashfree, match_kotak
from .db import conn, get_meta, init_db, now, q, set_meta, table_exists
from .matching import run_match

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
app = FastAPI(title="Payment - Zoho Reconciliation")
app.mount("/static", StaticFiles(directory=os.path.join(BASE, "static")), name="static")
init_db()

PAY_DATE_FMT = ("yyyy-mm-dd", "%Y-%m-%d")   # same as the Zoho sheet
ZOHO_DATE_FMT = ("dd/mm/yyyy", "%d/%m/%Y")  # same as the Razorpay sheet


@app.get("/")
def index():
    return FileResponse(os.path.join(BASE, "static", "index.html"))


# ---------- helpers ----------
def iso_date(value: Optional[str], fmt) -> Optional[str]:
    if not value:
        return None
    try:
        return datetime.strptime(value.strip(), fmt[1]).strftime("%Y-%m-%d")
    except ValueError:
        raise HTTPException(400, f"Date '{value}' is not in {fmt[0]} format.")


def page_bounds(page: int, size: int):
    return size, (max(page, 1) - 1) * size


def public(cols):
    return [c for c in cols if not c.startswith("_")]


def with_rate_column(cols):
    """Razorpay sheet columns with 'Currency (Paid In)' (currency of the matched Zoho payment) and 'Exchange Rate'
    placed right after the currency column. Returns (columns, position of the 'Currency (Paid In)' column)."""
    cols = list(cols)
    anchor = next((i for i, c in enumerate(cols) if c.lower() == "currency"), None)
    if anchor is None:  # no currency column in the sheet: fall back to after amount, else at the end
        anchor = next((i for i, c in enumerate(cols) if c.lower() == "amount"), len(cols) - 1)
    pos = anchor + 1
    return cols[:pos] + ["Currency (Paid In)", "Exchange Rate"] + cols[pos:], pos


def effective_rate(rate, currency):
    """Rate of the matched customer payment (from the Zoho sheet). An INR payment with no match (Missing) is
    trivially 1; a foreign-currency payment with no match has no rate to use, so it stays blank."""
    if rate is not None:
        return rate
    return 1.0 if (currency or "").strip().upper() == "INR" else None


def effective_pay_currency(pay_cur, razorpay_cur):
    """Currency of the Zoho payment. With no Zoho match (Missing) an INR Razorpay payment is INR, in line with the
    rate of 1 shown for it; a foreign-currency one stays blank."""
    if pay_cur:
        return pay_cur
    return "INR" if (razorpay_cur or "").strip().upper() == "INR" else None


def amount_inr(amount, rate, currency=None):
    """Razorpay 'amount' x exchange rate. (NOT the Payment Amount column, which comes from the Zoho sheet.)
    A Razorpay payment already in INR is not converted (its amount is rupees; the rate belongs to the Zoho payment)."""
    if rate is None:
        return None
    try:
        value = float(str(amount).replace(",", ""))
        if (currency or "").strip().upper() == "INR":
            return round(value, 2)
        return round(value * rate, 2)
    except (TypeError, ValueError):
        return None


# ---------- status ----------
@app.get("/api/status")
def status():
    with conn() as c:
        n = lambda t: c.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] if table_exists(c, t) else 0  # noqa: E731
        upd, mat = get_meta(c, "data_updated_at"), get_meta(c, "matched_at")
        return {
            "payments_rows": n("payments"), "zoho_rows": n("zoho"), "kotak_rows": n("kotak"), "cashfree_rows": n("cashfree"),
            "payments_loaded_at": get_meta(c, "payments_loaded_at"), "zoho_loaded_at": get_meta(c, "zoho_loaded_at"),
            "kotak_loaded_at": get_meta(c, "kotak_loaded_at"), "cashfree_loaded_at": get_meta(c, "cashfree_loaded_at"),
            "matched_at": mat, "fee_adjust": get_meta(c, "fee_adjust", True),
            "stale": bool(mat and upd and upd > mat),
        }


# ---------- clear data ----------
SHEET_KINDS = ("payments", "zoho", "kotak", "cashfree")      # uploaded sheets (table name = kind)
MATCH_SOURCES = ("razorpay", "kotak", "cashfree")            # match tables (Razorpay's is built from the zoho sheet)


def wipe_matches(c, source: str, remarks: bool):
    """Delete the match results of one sheet; with remarks=True also the agent/remarks typed on those rows."""
    if source == "razorpay":
        if remarks:
            c.execute("DELETE FROM annotations WHERE zoho_id IN (SELECT zoho_id FROM match_results)")
        c.execute("DELETE FROM match_results")
    else:
        if remarks:
            c.execute("DELETE FROM annotations WHERE zoho_id LIKE ?", (source + ":%",))
        c.execute("DELETE FROM src_match_results WHERE source = ?", (source,))
        c.execute("DELETE FROM meta WHERE key = ?", (f"{source}_matched_at",))


def tidy_after_clear(c):
    """Nothing matched any more -> forget when we matched, so the 'data changed since last match' banner can't linger."""
    left = c.execute("SELECT (SELECT COUNT(*) FROM match_results) + (SELECT COUNT(*) FROM src_match_results)").fetchone()[0]
    if not left:
        c.execute("DELETE FROM meta WHERE key = 'matched_at'")


@app.post("/api/clear/sheet/{kind}")
def clear_sheet(kind: str):
    """Remove an uploaded sheet from the database (rows, columns, load time) together with what was built on it.
    Clearing the Zoho sheet (customer payments) also empties every match table, as they were computed against it;
    remarks typed on the other sheets' rows are kept for those sheets, they only go when that sheet is cleared."""
    if kind not in SHEET_KINDS:
        raise HTTPException(404, "Unknown sheet")
    with conn() as c:
        if kind == "payments":
            for source in MATCH_SOURCES:
                wipe_matches(c, source, remarks=False)
        else:
            wipe_matches(c, "razorpay" if kind == "zoho" else kind, remarks=True)
        c.execute(f"DROP TABLE IF EXISTS {q(kind)}")
        c.execute("DELETE FROM meta WHERE key IN (?, ?)", (f"{kind}_cols", f"{kind}_loaded_at"))
        tidy_after_clear(c)
    return {"ok": True}


@app.post("/api/clear/matches/{source}")
def clear_matches(source: str):
    """Empty one match table (and the agent/remarks on it). The uploaded sheet stays; click Match to rebuild."""
    if source not in MATCH_SOURCES:
        raise HTTPException(404, "Unknown sheet")
    with conn() as c:
        wipe_matches(c, source, remarks=True)
        tidy_after_clear(c)
    return {"ok": True}


# ---------- upload ----------
XLSX = (".xlsx", ".xlsm")
# monthly sheets: kind -> (name used in messages, accepted file types, loader(filename, bytes))
MONTHLY = {
    "zoho": ("Razorpay", XLSX, lambda name, data: ingest.load_zoho(data)),
    "kotak": ("Kotak", (".csv",) + XLSX, lambda name, data: ingest_sheets.load_kotak(data, name)),
    "cashfree": ("Cash Free", XLSX, lambda name, data: ingest_sheets.load_cashfree(data, name)),
}


@app.post("/api/upload/{kind}")
async def upload(kind: str, files: List[UploadFile] = File(...)):
    """Upload one or more sheets at once.

    payments (customer payments = the Zoho sheet, full list): all files are combined and replace the previous
    data; if any file is invalid nothing is changed. zoho (Razorpay) / kotak / cashfree (monthly): each file is
    loaded in turn (upsert, so re-uploading a month is safe) and a bad file doesn't stop the others - the
    per-file outcome is in `results`.
    """
    if kind != "payments" and kind not in MONTHLY:
        raise HTTPException(404, "Unknown sheet type")
    if not files:
        raise HTTPException(400, "No files selected.")
    allowed = XLSX if kind == "payments" else MONTHLY[kind][1]
    for f in files:
        if not (f.filename or "").lower().endswith(allowed):
            raise HTTPException(400, f"{f.filename or 'A file'}: unsupported file type - please upload {', '.join(allowed)}.")
    blobs = [(f.filename, await f.read()) for f in files]

    if kind == "payments":
        try:
            return ingest.load_payments(blobs)
        except ingest.SheetError as e:
            raise HTTPException(400, str(e))

    label, _, loader = MONTHLY[kind]
    results, warnings, total_rows = [], [], 0
    for name, data in blobs:
        try:
            r = loader(name, data)
            total_rows += r["rows"]
            results.append({"filename": name, "ok": True, "message": r["message"], "warnings": r["warnings"]})
            warnings += [f"{name}: {w}" for w in r["warnings"]]
        except ingest.SheetError as e:
            results.append({"filename": name, "ok": False, "message": str(e), "warnings": []})
    done = sum(r["ok"] for r in results)
    if not done:
        raise HTTPException(400, "\n".join(f"{r['filename']}: {r['message']}" for r in results))
    msg = f"Loaded {done} of {len(results)} {label} sheet(s) - {total_rows:,} rows in total."
    return {"rows": total_rows, "message": msg, "warnings": warnings, "results": results}


# ---------- match ----------
class MatchReq(BaseModel):
    fee_adjust: bool = True


# one matching file per sheet (see app/match_*.py); Razorpay's lives in app/matching.py
MATCHERS = {"kotak": match_kotak.run, "cashfree": match_cashfree.run}


@app.post("/api/match")
def match(req: MatchReq):
    """Match every monthly sheet that has been uploaded against the customer payments (Zoho sheet)."""
    with conn() as c:
        if not table_exists(c, "payments"):
            raise HTTPException(400, "Upload the Zoho sheet before matching.")
        have = [t for t in ("zoho", "kotak", "cashfree") if table_exists(c, t)]
        if not have:
            raise HTTPException(400, "Upload at least one monthly sheet (Razorpay, Kotak or Cash Free) before matching.")
    summary = {}
    if "zoho" in have:
        summary["razorpay"] = run_match(req.fee_adjust)
    with conn() as c:
        for source, run in MATCHERS.items():
            if source in have:
                summary[source] = run(c)
        set_meta(c, "matched_at", now())
        set_meta(c, "fee_adjust", req.fee_adjust)
    return {**summary.get("razorpay", {}), "sources": summary}


@app.get("/api/matches")
def matches(
    result: str = "All", status: str = "captured", currency: str = "", q_: str = Query("", alias="q"),
    page: int = 1, page_size: int = 100,
):
    with conn() as c:
        if not table_exists(c, "zoho"):
            return {"columns": [], "rows": [], "total": 0, "counts": {}, "statuses": [], "currencies": []}
        cols = public(get_meta(c, "zoho_cols", []))
        has_cur = "currency" in {x.lower() for x in cols}
        where, args = ["1=1"], []
        if status and status != "All":
            where.append('z."status" = ?'); args.append(status)
        if currency and has_cur:
            where.append('z."currency" = ?'); args.append(currency)
        if q_.strip():
            where.append('(z."id" LIKE ? OR z."notes" LIKE ? OR z."description" LIKE ?)')
            args += [f"%{q_.strip()}%"] * 3
        base = f'FROM zoho z JOIN match_results m ON m.zoho_id = z.id LEFT JOIN annotations a ON a.zoho_id = z.id WHERE {" AND ".join(where)}'
        counts = {r["result"]: r["n"] for r in c.execute(f"SELECT m.result, COUNT(*) n {base} GROUP BY m.result", args)}
        counts["All"] = sum(counts.values())
        w2, a2 = base, list(args)
        if result != "All":
            w2 += " AND m.result = ?"; a2.append(result)
        total = c.execute(f"SELECT COUNT(*) {w2}", a2).fetchone()[0]
        lim, off = page_bounds(page, page_size)
        sel = ", ".join(f"z.{q(x)}" for x in cols)
        rs = c.execute(
            f"SELECT m.result, m.matched_by, m.payment_amount, m.difference, m.basis, m.invoice_number, m.exchange_rate, m.pay_currency, z._fee AS _fee, "
            f"a.agent, a.remarks, z.id AS _id, {sel} "
            f"{w2} ORDER BY z._created, z.id LIMIT ? OFFSET ?", a2 + [lim, off]).fetchall()
        out_cols, pos = with_rate_column(cols)
        rows = []
        for r in rs:
            cells = [r[x] for x in cols]
            cur = r[next(x for x in cols if x.lower() == "currency")] if has_cur else None
            rate = effective_rate(r["exchange_rate"], cur)
            amt = r[next(x for x in cols if x.lower() == "amount")] if any(x.lower() == "amount" for x in cols) else None
            cells[pos:pos] = [effective_pay_currency(r["pay_currency"], cur), rate]
            rows.append({
                "id": r["_id"], "result": r["result"], "matched_by": r["matched_by"], "payment_amount": r["payment_amount"],
                "fee": r["_fee"], "amount_inr": amount_inr(amt, rate, cur),
                "difference": r["difference"], "basis": r["basis"], "invoice_number": r["invoice_number"] or "",
                "agent": r["agent"], "remarks": r["remarks"] or "",
                "cells": cells,
            })
        currencies = [r[0] for r in c.execute('SELECT DISTINCT "currency" FROM zoho WHERE "currency" IS NOT NULL ORDER BY 1')] if has_cur else []
        statuses = [r[0] for r in c.execute('SELECT DISTINCT "status" FROM zoho WHERE "status" IS NOT NULL ORDER BY 1')]
        return {"columns": out_cols, "paid_col": pos, "rate_col": pos + 1, "rows": rows, "total": total, "counts": counts, "statuses": statuses,
                "currencies": currencies,
                "agents": [r[0] for r in c.execute("SELECT name FROM agents ORDER BY name")]}


MATCH_EXTRA = ["Result", "Matched By", "Agent", "Remarks", "Payment Amount", "Fees", "Amount (INR)", "Difference", "Amount Basis"]


@app.get("/api/matches/export")
def export_matches(result: str = "All", status: str = "captured", currency: str = ""):
    """Excel download of the match tables: one sheet each for Matched / Mismatched / Missing (or just one)."""
    if result not in ("All", "Matched", "Mismatched", "Missing"):
        raise HTTPException(400, "Unknown result type")
    names = ["Matched", "Mismatched", "Missing"] if result == "All" else [result]
    wb = Workbook(write_only=True)
    with conn() as c:
        cols = public(get_meta(c, "zoho_cols", []))
        lower = [x.lower() for x in cols]
        has_cur = "currency" in lower
        cur_i = lower.index("currency") if has_cur else None
        amt_i = lower.index("amount") if "amount" in lower else None
        out_cols, pos = with_rate_column(cols)
        sel = ", ".join(f"z.{q(x)}" for x in cols)
        n = len(cols)
        for name in names:
            ws = wb.create_sheet(name)
            ws.append(["Invoice Number"] + out_cols + MATCH_EXTRA)
            where, args = "m.result = ?", [name]
            if status and status != "All":
                where += ' AND z."status" = ?'; args.append(status)
            if currency and has_cur:
                where += ' AND z."currency" = ?'; args.append(currency)
            # row layout: 0 invoice | 1..n sheet cols | n+1 result, matched_by, agent, remarks, payment_amount, difference, basis | exchange_rate
            for r in c.execute(
                f"SELECT m.invoice_number, {sel}, m.result, m.matched_by, a.agent, a.remarks, "
                f"m.payment_amount, m.difference, m.basis, m.exchange_rate, z.\"_fee\", m.pay_currency "
                f"FROM zoho z JOIN match_results m ON m.zoho_id = z.id LEFT JOIN annotations a ON a.zoho_id = z.id "
                f'WHERE {where} ORDER BY z."_created", z."id"', args):
                v = list(r)
                cells = v[1:1 + n]
                res, matched_by, agent, remarks, pay_amt, diff, basis, rate, fee, pay_cur = v[1 + n:]
                if res != "Matched" and not agent:
                    agent = "Unassigned"
                rate = effective_rate(rate, cells[cur_i] if has_cur else None)
                a_inr = amount_inr(cells[amt_i] if amt_i is not None else None, rate, cells[cur_i] if has_cur else None)
                cells[pos:pos] = [effective_pay_currency(pay_cur, cells[cur_i] if has_cur else None), rate]
                ws.append([v[0]] + cells + [res, matched_by, agent, remarks, pay_amt, fee, a_inr, diff, basis])
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    fname = f"reconciliation_{'all' if result == 'All' else result.lower()}.xlsx"
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f'attachment; filename="{fname}"'})


class Annotation(BaseModel):
    zoho_id: str
    agent: Optional[str] = None
    remarks: Optional[str] = None


@app.post("/api/annotation")
def annotate(a: Annotation):
    """Set agent and/or remarks for a payment. Only the fields sent are changed."""
    fields = a.model_dump(exclude_unset=True)
    fields.pop("zoho_id")
    if not fields:
        return {"ok": True}
    if "agent" in fields and fields["agent"] in ("", "Unassigned"):
        fields["agent"] = None
    with conn() as c:
        c.execute("INSERT OR IGNORE INTO annotations(zoho_id) VALUES (?)", (a.zoho_id,))
        for k, v in fields.items():
            c.execute(f"UPDATE annotations SET {k} = ? WHERE zoho_id = ?", (v, a.zoho_id))
    return {"ok": True}


class Agent(BaseModel):
    name: str


@app.get("/api/agents")
def agents():
    with conn() as c:
        return [r[0] for r in c.execute("SELECT name FROM agents ORDER BY name")]


@app.post("/api/agents")
def add_agent(a: Agent):
    if not a.name.strip():
        raise HTTPException(400, "Name required")
    with conn() as c:
        c.execute("INSERT OR IGNORE INTO agents(name) VALUES (?)", (a.name.strip(),))
    return agents()


# ---------- raw data tabs ----------
@app.get("/api/payments")
def payments(mode: str = "", deposit_to: str = "", currency: str = "", date_from: str = "", date_to: str = "", page: int = 1, page_size: int = 100):
    d1, d2 = iso_date(date_from, PAY_DATE_FMT), iso_date(date_to, PAY_DATE_FMT)
    with conn() as c:
        if not table_exists(c, "payments"):
            return {"columns": [], "rows": [], "total": 0, "modes": [], "deposit_tos": [], "currencies": []}
        cols = public(get_meta(c, "payments_cols", []))
        cur_col = next((x for x in cols if x in ("Currency Code", "Currency")), None)
        where, args = ["1=1"], []
        if mode: where.append('"Mode" = ?'); args.append(mode)
        if deposit_to: where.append('"Deposit To" = ?'); args.append(deposit_to)
        if currency and cur_col: where.append(f'{q(cur_col)} = ?'); args.append(currency)
        if d1: where.append('"_date" >= ?'); args.append(d1)
        if d2: where.append('"_date" <= ?'); args.append(d2)
        w = " AND ".join(where)
        total = c.execute(f"SELECT COUNT(*) FROM payments WHERE {w}", args).fetchone()[0]
        lim, off = page_bounds(page, page_size)
        rows = c.execute(f'SELECT {", ".join(q(x) for x in cols)} FROM payments WHERE {w} ORDER BY rowid LIMIT ? OFFSET ?', args + [lim, off]).fetchall()
        distinct = lambda col: [r[0] for r in c.execute(f'SELECT DISTINCT {q(col)} FROM payments WHERE {q(col)} IS NOT NULL ORDER BY 1')]  # noqa: E731
        currencies = distinct(cur_col) if cur_col else []
        return {"columns": cols, "rows": [list(r) for r in rows], "total": total,
                "modes": distinct("Mode"), "deposit_tos": distinct("Deposit To"), "currencies": currencies}


@app.get("/api/zoho")
def zoho(method: str = "", currency: str = "", date_from: str = "", date_to: str = "", page: int = 1, page_size: int = 100):
    d1, d2 = iso_date(date_from, ZOHO_DATE_FMT), iso_date(date_to, ZOHO_DATE_FMT)
    with conn() as c:
        if not table_exists(c, "zoho"):
            return {"columns": [], "rows": [], "total": 0, "methods": [], "currencies": []}
        cols = public(get_meta(c, "zoho_cols", []))
        has_cur = "currency" in {x.lower() for x in cols}
        where, args = ["1=1"], []
        if method: where.append('"method" = ?'); args.append(method)
        if currency and has_cur: where.append('"currency" = ?'); args.append(currency)
        if d1: where.append('substr("_created",1,10) >= ?'); args.append(d1)
        if d2: where.append('substr("_created",1,10) <= ?'); args.append(d2)
        w = " AND ".join(where)
        total = c.execute(f"SELECT COUNT(*) FROM zoho WHERE {w}", args).fetchone()[0]
        lim, off = page_bounds(page, page_size)
        rows = c.execute(f'SELECT {", ".join(q(x) for x in cols)} FROM zoho WHERE {w} ORDER BY "_created", "id" LIMIT ? OFFSET ?', args + [lim, off]).fetchall()
        methods = [r[0] for r in c.execute('SELECT DISTINCT "method" FROM zoho WHERE "method" IS NOT NULL ORDER BY 1')]
        currencies = [r[0] for r in c.execute('SELECT DISTINCT "currency" FROM zoho WHERE "currency" IS NOT NULL ORDER BY 1')] if has_cur else []
        return {"columns": cols, "rows": [list(r) for r in rows], "total": total, "methods": methods, "currencies": currencies}

# ======================================================================================================
# Other monthly sheets (Kotak, Cash Free): one generic set of endpoints, driven by this table.
#   selects: dropdown filters; `col` is a stored column, `canon` finds the sheet's own column by canonical name.
#   match_default: value the Match tab starts on (when that value exists in the data).
# ======================================================================================================
SHEET_DATE_FMT = ("dd-mm-yyyy", "%d-%m-%Y")  # both sheets show dates this way
SOURCES = {
    "kotak": {
        "label": "Kotak", "table": "kotak",
        "selects": [{"param": "drcr", "label": "Dr / Cr", "col": "_drcr", "match_default": "CR"}],
        "search": ["description", "chqrefno"],
    },
    "cashfree": {
        "label": "Cash Free", "table": "cashfree",
        "selects": [{"param": "status", "label": "Transaction status", "canon": "transactionstatus", "match_default": "SUCCESS"},
                    {"param": "mode", "label": "Payment mode", "canon": "paymentmode"}],
        "search": ["orderid", "cashfreeorderid", "referenceid"],
    },
}
SOURCE_EXTRA = ["Result", "Matched By", "Agent", "Remarks", "Payment Amount", "Difference", "Amount Basis"]
BLANK = "(blank)"


def source_cfg(source: str):
    if source not in SOURCES:
        raise HTTPException(404, "Unknown sheet")
    return SOURCES[source]


def find_col(cols, canon_name):
    return next((x for x in cols if ingest_sheets.canon(x) == canon_name), None)


def source_filters(c, cfg, cols, params, use_defaults=False, prefix=""):
    """WHERE parts for the dropdown filters. Returns (where, args, descriptors for the UI incl. options + chosen value)."""
    where, args, shown = [], [], []
    for sel in cfg["selects"]:
        col = sel.get("col") or find_col(cols, sel["canon"])
        if col is None:
            continue
        opts = [r[0] for r in c.execute(f'SELECT DISTINCT {q(col)} FROM {q(cfg["table"])} ORDER BY 1')]
        has_blank = any(o in (None, "") for o in opts)
        opts = [o for o in opts if o not in (None, "")] + ([BLANK] if has_blank else [])
        val = params.get(sel["param"])
        if val is None and use_defaults:
            val = sel.get("match_default") if sel.get("match_default") in opts else "All"
        val = val or "All"
        if val == BLANK:
            where.append(f"({prefix}{q(col)} IS NULL OR {prefix}{q(col)} = '')")
        elif val != "All":
            where.append(f"{prefix}{q(col)} = ?")
            args.append(val)
        shown.append({"param": sel["param"], "label": sel["label"], "options": opts, "value": val})
    return where, args, shown


def source_search(cfg, cols, text, prefix=""):
    names = [find_col(cols, n) for n in cfg["search"]]
    names = [n for n in names if n]
    if not text.strip() or not names:
        return [], []
    return [f"({' OR '.join(f'{prefix}{q(n)} LIKE ?' for n in names)})"], [f"%{text.strip()}%"] * len(names)


def date_where(d1, d2):
    where, args = [], []
    if d1: where.append('substr("_dt",1,10) >= ?'); args.append(d1)
    if d2: where.append('substr("_dt",1,10) <= ?'); args.append(d2)
    return where, args


def match_join(source, table):
    return (f'FROM {q(table)} z JOIN src_match_results m ON m.source = \'{source}\' AND m.row_key = z."_key" '
            f'LEFT JOIN annotations a ON a.zoho_id = \'{source}:\' || z."_key"')


@app.get("/api/recon/{source}/matches")
def source_matches(source: str, request: Request, result: str = "All", q_: str = Query("", alias="q"),
                   page: int = 1, page_size: int = 100, defaults: int = 0):
    cfg = source_cfg(source)
    with conn() as c:
        if not (table_exists(c, cfg["table"]) and c.execute("SELECT 1 FROM src_match_results WHERE source = ? LIMIT 1", (source,)).fetchone()):
            has_sheet = table_exists(c, cfg["table"])
            return {"columns": [], "rows": [], "total": 0, "counts": {}, "filters": [], "agents": [], "has_sheet": has_sheet}
        cols = public(get_meta(c, f"{source}_cols", []))
        where, args, shown = source_filters(c, cfg, cols, dict(request.query_params), bool(defaults), "z.")
        sw, sa = source_search(cfg, cols, q_, "z.")
        where, args = where + sw, args + sa
        base = f'{match_join(source, cfg["table"])} WHERE {" AND ".join(where or ["1=1"])}'
        counts = {r["result"]: r["n"] for r in c.execute(f"SELECT m.result, COUNT(*) n {base} GROUP BY m.result", args)}
        counts["All"] = sum(counts.values())
        w2, a2 = base, list(args)
        if result != "All":
            w2 += " AND m.result = ?"; a2.append(result)
        total = c.execute(f"SELECT COUNT(*) {w2}", a2).fetchone()[0]
        lim, off = page_bounds(page, page_size)
        rs = c.execute(
            f"SELECT m.result, m.matched_by, m.payment_amount, m.difference, m.basis, m.invoice_number, a.agent, a.remarks, z.\"_key\", "
            f'{", ".join("z." + q(x) for x in cols)} {w2} ORDER BY z."_dt", z."_key" LIMIT ? OFFSET ?', a2 + [lim, off]).fetchall()
        rows = [{"id": f"{source}:{r[8]}", "result": r[0], "matched_by": r[1], "payment_amount": r[2], "difference": r[3],
                 "basis": r[4], "invoice_number": r[5] or "", "agent": r[6], "remarks": r[7] or "", "cells": list(r[9:])} for r in rs]
        return {"columns": cols, "rows": rows, "total": total, "counts": counts, "filters": shown, "has_sheet": True,
                "agents": [r[0] for r in c.execute("SELECT name FROM agents ORDER BY name")]}


@app.get("/api/recon/{source}/export")
def source_export(source: str, request: Request, result: str = "All"):
    """Excel download of this sheet's match tables: Matched / Mismatched / Missing sheets (or just one)."""
    cfg = source_cfg(source)
    if result not in ("All", "Matched", "Mismatched", "Missing"):
        raise HTTPException(400, "Unknown result type")
    names = ["Matched", "Mismatched", "Missing"] if result == "All" else [result]
    wb = Workbook(write_only=True)
    with conn() as c:
        if not table_exists(c, cfg["table"]):
            raise HTTPException(400, f"Upload the {cfg['label']} sheet first.")
        cols = public(get_meta(c, f"{source}_cols", []))
        fw, fa, _ = source_filters(c, cfg, cols, dict(request.query_params), False, "z.")
        sw, sa = source_search(cfg, cols, request.query_params.get("q", ""), "z.")
        fw, fa = fw + sw, fa + sa
        for name in names:
            ws = wb.create_sheet(name)
            ws.append(["Invoice Number"] + cols + SOURCE_EXTRA)
            where = " AND ".join(["m.result = ?"] + fw)
            for r in c.execute(
                    f'SELECT m.invoice_number, {", ".join("z." + q(x) for x in cols)}, m.result, m.matched_by, a.agent, a.remarks, '
                    f'm.payment_amount, m.difference, m.basis {match_join(source, cfg["table"])} '
                    f'WHERE {where} ORDER BY z."_dt", z."_key"', [name] + fa):
                v = list(r)
                n = len(cols)
                res, matched_by, agent, remarks, pay_amt, diff, basis = v[1 + n:]
                if res != "Matched" and not agent:
                    agent = "Unassigned"
                ws.append([v[0]] + v[1:1 + n] + [res, matched_by, agent, remarks, pay_amt, diff, basis])
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    fname = f"reconciliation_{source}_{'all' if result == 'All' else result.lower()}.xlsx"
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f'attachment; filename="{fname}"'})


@app.get("/api/sheet/{source}")
def source_sheet(source: str, request: Request, date_from: str = "", date_to: str = "", page: int = 1, page_size: int = 100):
    """The sheet as uploaded (original columns), with its dropdown filters and a date range."""
    cfg = source_cfg(source)
    d1, d2 = iso_date(date_from, SHEET_DATE_FMT), iso_date(date_to, SHEET_DATE_FMT)
    with conn() as c:
        if not table_exists(c, cfg["table"]):
            return {"columns": [], "rows": [], "total": 0, "filters": []}
        cols = public(get_meta(c, f"{source}_cols", []))
        where, args, shown = source_filters(c, cfg, cols, dict(request.query_params))
        dw, da = date_where(d1, d2)
        w = " AND ".join(where + dw or ["1=1"])
        total = c.execute(f'SELECT COUNT(*) FROM {q(cfg["table"])} WHERE {w}', args + da).fetchone()[0]
        lim, off = page_bounds(page, page_size)
        rows = c.execute(f'SELECT {", ".join(q(x) for x in cols)} FROM {q(cfg["table"])} WHERE {w} ORDER BY "_dt", "_key" LIMIT ? OFFSET ?',
                         args + da + [lim, off]).fetchall()
        return {"columns": cols, "rows": [list(r) for r in rows], "total": total, "filters": shown}