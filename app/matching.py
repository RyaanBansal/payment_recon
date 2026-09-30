"""Reconciliation: Razorpay sheet vs Zoho sheet."""
import json
import re
from collections import defaultdict

from .db import conn, get_meta, set_meta, now

EXACT_TOL = 0.005
FEE_TOL = 1.00  # Rs. tolerance when comparing net of the Razorpay fee (fee/GST rounding differs by a few paise)


def notes_invoice_id(notes):
    """notes looks like {"invoice_id_0":"2192622000009235076"}; other shapes ([], {"comment":""}, ...) -> None."""
    if not notes:
        return None
    try:
        d = json.loads(notes)
        if isinstance(d, dict):
            v = d.get("invoice_id_0")
            return str(v).strip() if v not in (None, "") else None
        return None
    except (ValueError, TypeError):
        m = re.search(r'invoice_id_0"?\s*[:=]\s*"?(\d+)', notes)
        return m.group(1) if m else None


def _index(rows, key, has_invoice):
    """key -> {CustomerPayment ID: (Amount, Invoice Number)}. Summing distinct customer payments (not rows)
    avoids double counting a payment that is split across several invoices."""
    idx = defaultdict(dict)
    for r in rows:
        k = r[key]
        if k:
            inv_no = r["Invoice Number"] if has_invoice else None
            idx[k][r["CustomerPayment ID"]] = (r["_amount"] or 0.0, inv_no)
    return idx


def _classify(zamt, pamt, fee, fee_adjust):
    """Compare a Zoho amount to a (possibly summed) payment amount. Returns (difference, basis|None)."""
    diff = round(zamt - pamt, 2)
    if abs(diff) < EXACT_TOL:
        return diff, "Exact"
    if fee_adjust and fee is not None and abs(diff - fee) <= FEE_TOL:
        return diff, "Net of Razorpay fee"
    return diff, None


def run_match(fee_adjust: bool = True) -> dict:
    with conn() as c:
        pay_cols = get_meta(c, "payments_cols", [])
        has_invoice = "Invoice Number" in pay_cols  # older Zoho sheet exports may not include this column
        inv_sel = ', "Invoice Number"' if has_invoice else ""
        pay = c.execute(
            f'SELECT "CustomerPayment ID","InvoicePayment ID","Reference Number","_amount"{inv_sel} FROM payments'
        ).fetchall()
        by_ipid = _index(pay, "InvoicePayment ID", has_invoice)
        by_ref = _index(pay, "Reference Number", has_invoice)
        out, counts = [], defaultdict(int)
        for z in c.execute('SELECT "id","notes","_amount","_fee" FROM zoho').fetchall():
            zamt = z["_amount"] or 0.0
            inv = notes_invoice_id(z["notes"])
            primary = by_ipid.get(inv) if inv else None
            ref = by_ref.get(z["id"])

            if primary:
                cand, by = primary, "InvoicePayment ID"
                pamt = round(sum(v[0] for v in primary.values()), 2)
                diff, basis = _classify(zamt, pamt, z["_fee"], fee_adjust)
                # The notes field only ever names one invoice, so a mismatch here might just mean this
                # Razorpay payment actually covers several invoices sharing the same Reference Number.
                # If the sum across that whole group matches, use the group instead.
                if basis is None and ref and len(ref) > 1:
                    ref_pamt = round(sum(v[0] for v in ref.values()), 2)
                    ref_diff, ref_basis = _classify(zamt, ref_pamt, z["_fee"], fee_adjust)
                    if ref_basis is not None:
                        cand, by, pamt, diff, basis = ref, "Reference Number", ref_pamt, ref_diff, ref_basis
            elif ref:
                cand, by = ref, "Reference Number"
                pamt = round(sum(v[0] for v in ref.values()), 2)
                diff, basis = _classify(zamt, pamt, z["_fee"], fee_adjust)
            else:
                cand = by = pamt = diff = basis = None

            if cand is None:
                # Missing: no counterpart in the Zoho sheet at all, so there is no invoice number to show.
                res, inv_no = "Missing", None
            else:
                res = "Matched" if basis else "Mismatched"
                # Invoice number(s) attached to the matching Zoho-sheet row(s); several if the payment
                # covers multiple invoices. Set for Matched and Mismatched only, never for Missing.
                nums = sorted({v[1] for v in cand.values() if v[1]})
                inv_no = ", ".join(nums) if nums else None
            counts[res] += 1
            out.append((z["id"], res, by, pamt, diff, basis, inv_no))
        c.execute("DELETE FROM match_results")
        c.executemany("INSERT INTO match_results VALUES (?,?,?,?,?,?,?)", out)
        set_meta(c, "matched_at", now())
        set_meta(c, "fee_adjust", fee_adjust)
    return {"total": len(out), **{k: counts.get(k, 0) for k in ("Matched", "Mismatched", "Missing")}}