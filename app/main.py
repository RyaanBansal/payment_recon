"""FastAPI app: serves the single-page UI + JSON API."""
import os
import re
from io import BytesIO
from datetime import datetime
from typing import Optional

from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from openpyxl import Workbook
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import ingest
from .db import conn, get_meta, init_db, q, table_exists
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


# ---------- status ----------
@app.get("/api/status")
def status():
    with conn() as c:
        n = lambda t: c.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] if table_exists(c, t) else 0  # noqa: E731
        upd, mat = get_meta(c, "data_updated_at"), get_meta(c, "matched_at")
        return {
            "payments_rows": n("payments"), "zoho_rows": n("zoho"),
            "payments_loaded_at": get_meta(c, "payments_loaded_at"), "zoho_loaded_at": get_meta(c, "zoho_loaded_at"),
            "matched_at": mat, "fee_adjust": get_meta(c, "fee_adjust", True),
            "stale": bool(mat and upd and upd > mat),
        }


# ---------- upload ----------
@app.post("/api/upload/{kind}")
async def upload(kind: str, file: UploadFile = File(...)):
    if kind not in ("payments", "zoho"):
        raise HTTPException(404, "Unknown sheet type")
    if not (file.filename or "").lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Please upload an Excel (.xlsx) file.")
    data = await file.read()
    try:
        return (ingest.load_payments if kind == "payments" else ingest.load_zoho)(data)
    except ingest.SheetError as e:
        raise HTTPException(400, str(e))


# ---------- match ----------
class MatchReq(BaseModel):
    fee_adjust: bool = True


@app.post("/api/match")
def match(req: MatchReq):
    with conn() as c:
        if not (table_exists(c, "payments") and table_exists(c, "zoho")):
            raise HTTPException(400, "Upload both sheets before matching.")
    return run_match(req.fee_adjust)


@app.get("/api/matches")
def matches(
    result: str = "All", status: str = "captured", q_: str = Query("", alias="q"),
    page: int = 1, page_size: int = 100,
):
    with conn() as c:
        if not table_exists(c, "zoho"):
            return {"columns": [], "rows": [], "total": 0, "counts": {}, "statuses": []}
        cols = public(get_meta(c, "zoho_cols", []))
        where, args = ["1=1"], []
        if status and status != "All":
            where.append('z."status" = ?'); args.append(status)
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
            f"SELECT m.result, m.matched_by, m.payment_amount, m.difference, m.basis, m.invoice_number, "
            f"a.agent, a.remarks, z.id AS _id, {sel} "
            f"{w2} ORDER BY z._created, z.id LIMIT ? OFFSET ?", a2 + [lim, off]).fetchall()
        rows = [{
            "id": r["_id"], "result": r["result"], "matched_by": r["matched_by"], "payment_amount": r["payment_amount"],
            "difference": r["difference"], "basis": r["basis"], "invoice_number": r["invoice_number"] or "",
            "agent": r["agent"], "remarks": r["remarks"] or "",
            "cells": [r[x] for x in cols],
        } for r in rs]
        statuses = [r[0] for r in c.execute('SELECT DISTINCT "status" FROM zoho WHERE "status" IS NOT NULL ORDER BY 1')]
        return {"columns": cols, "rows": rows, "total": total, "counts": counts, "statuses": statuses,
                "agents": [r[0] for r in c.execute("SELECT name FROM agents ORDER BY name")]}


MATCH_EXTRA = ["Result", "Matched By", "Agent", "Remarks", "Payment Amount", "Difference", "Amount Basis"]


@app.get("/api/matches/export")
def export_matches(result: str = "All", status: str = "captured"):
    """Excel download of the match tables: one sheet each for Matched / Mismatched / Missing (or just one)."""
    if result not in ("All", "Matched", "Mismatched", "Missing"):
        raise HTTPException(400, "Unknown result type")
    names = ["Matched", "Mismatched", "Missing"] if result == "All" else [result]
    wb = Workbook(write_only=True)
    with conn() as c:
        cols = public(get_meta(c, "zoho_cols", []))
        sel = ", ".join(f"z.{q(x)}" for x in cols)
        for name in names:
            ws = wb.create_sheet(name)
            ordered_cols = list(cols)
            if "Currency" in ordered_cols and "Exchange Rate" in ordered_cols:
                ex = ordered_cols.pop(ordered_cols.index("Exchange Rate"))
                ordered_cols.insert(ordered_cols.index("Currency") + 1, ex)
            ws.append(["Invoice Number"] + ordered_cols + MATCH_EXTRA)
            where, args = "m.result = ?", [name]
            if status and status != "All":
                where += ' AND z."status" = ?'; args.append(status)
            for r in c.execute(
                f"SELECT m.invoice_number, {sel}, m.result, m.matched_by, a.agent, a.remarks, "
                f"m.payment_amount, m.difference, m.basis "
                f"FROM zoho z JOIN match_results m ON m.zoho_id = z.id LEFT JOIN annotations a ON a.zoho_id = z.id "
                f'WHERE {where} ORDER BY z."_created", z."id"', args):
                v = list(r)
                n = len(cols) + 1  # "Invoice Number" occupies index 0, then the sheet columns
                if v[n] != "Matched" and not v[n + 2]:
                    v[n + 2] = "Unassigned"
                amount_idx = cols.index("Amount") if "Amount" in cols else -1
                exchange_idx = cols.index("Exchange Rate") if "Exchange Rate" in cols else -1
                if amount_idx >= 0 and exchange_idx >= 0:
                    amount_val = v[1 + amount_idx]
                    rate_val = v[1 + exchange_idx]
                    try:
                        amount_inr = float(str(amount_val).replace(',', '')) * float(str(rate_val).replace(',', ''))
                    except (TypeError, ValueError):
                        amount_inr = ""
                    v.insert(n + 5, amount_inr)
                else:
                    v.insert(n + 5, "")
                ws.append(v)
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
        where, args = ["1=1"], []
        if mode: where.append('"Mode" = ?'); args.append(mode)
        if deposit_to: where.append('"Deposit To" = ?'); args.append(deposit_to)
        if currency: where.append('"Currency" = ?'); args.append(currency)
        if d1: where.append('"_date" >= ?'); args.append(d1)
        if d2: where.append('"_date" <= ?'); args.append(d2)
        w = " AND ".join(where)
        total = c.execute(f"SELECT COUNT(*) FROM payments WHERE {w}", args).fetchone()[0]
        lim, off = page_bounds(page, page_size)
        rows = c.execute(f'SELECT {", ".join(q(x) for x in cols)} FROM payments WHERE {w} ORDER BY rowid LIMIT ? OFFSET ?', args + [lim, off]).fetchall()
        distinct = lambda col: [r[0] for r in c.execute(f'SELECT DISTINCT {q(col)} FROM payments WHERE {q(col)} IS NOT NULL ORDER BY 1')]  # noqa: E731
        currencies = distinct("Currency") if "Currency" in cols else []
        return {"columns": cols, "rows": [list(r) for r in rows], "total": total,
                "modes": distinct("Mode"), "deposit_tos": distinct("Deposit To"), "currencies": currencies}


@app.get("/api/zoho")
def zoho(method: str = "", currency: str = "", date_from: str = "", date_to: str = "", page: int = 1, page_size: int = 100):
    d1, d2 = iso_date(date_from, ZOHO_DATE_FMT), iso_date(date_to, ZOHO_DATE_FMT)
    with conn() as c:
        if not table_exists(c, "zoho"):
            return {"columns": [], "rows": [], "total": 0, "methods": [], "currencies": []}
        cols = public(get_meta(c, "zoho_cols", []))
        where, args = ["1=1"], []
        if method: where.append('"method" = ?'); args.append(method)
        if currency: where.append('"Currency" = ?'); args.append(currency)
        if d1: where.append('substr("_created",1,10) >= ?'); args.append(d1)
        if d2: where.append('substr("_created",1,10) <= ?'); args.append(d2)
        w = " AND ".join(where)
        total = c.execute(f"SELECT COUNT(*) FROM zoho WHERE {w}", args).fetchone()[0]
        lim, off = page_bounds(page, page_size)
        rows = c.execute(f'SELECT {", ".join(q(x) for x in cols)} FROM zoho WHERE {w} ORDER BY "_created", "id" LIMIT ? OFFSET ?', args + [lim, off]).fetchall()
        methods = [r[0] for r in c.execute('SELECT DISTINCT "method" FROM zoho WHERE "method" IS NOT NULL ORDER BY 1')]
        currencies = [r[0] for r in c.execute('SELECT DISTINCT "Currency" FROM zoho WHERE "Currency" IS NOT NULL ORDER BY 1')] if "Currency" in cols else []
        return {"columns": cols, "rows": [list(r) for r in rows], "total": total, "methods": methods, "currencies": currencies}