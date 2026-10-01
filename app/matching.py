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


def _to_rate(v):
    try:
        f = float(str(v).replace(",", ""))
        return f if f > 0 else None
    except (TypeError, ValueError):
        return None


def _index(rows, key, has_invoice, has_rate, has_cur):
    """key -> {CustomerPayment ID: (Amount, Invoice Number, Exchange Rate, Currency (Zoho sheet))}. Summing distinct customer payments
    (not rows) avoids double counting a payment that is split across several invoices."""
    idx = defaultdict(dict)
    for r in rows:
        k = r[key]
        if k:
            inv_no = r["Invoice Number"] if has_invoice else None
            rate = _to_rate(r["Exchange Rate"]) if has_rate else None
            cur = (r[CUR_COL] or "").strip().upper() or None if has_cur else None
            idx[k][r["CustomerPayment ID"]] = (r["_amount"] or 0.0, inv_no, rate, cur)
    return idx


CUR_COL = "Currency Code"  # currency column of the Zoho (customer payment) sheet


def _rate_cur_of(cand):
    """(exchange rate, currency) of the matched customer payment(s), both from the SAME Zoho row so they always
    agree (first row by CustomerPayment ID if there are several)."""
    items = sorted(cand.items())
    for _, v in items:
        if v[2] is not None:
            return v[2], v[3]
    return None, next((v[3] for _, v in items if v[3]), None)


def _rate_of(cand):
    """Exchange rate of the matched customer payment(s). If they somehow carry different rates, use the first
    one (sorted by CustomerPayment ID) so the result is stable."""
    for _, v in sorted(cand.items()):
        if v[2] is not None:
            return v[2]
    return None


def _inr(zamt, cand, currency):
    """Razorpay amount in INR = amount x exchange rate of the matched customer payment(s) (same value the UI shows
    as Amount (INR)). A Razorpay payment that is already in INR is NOT multiplied: the matched Zoho payment may be
    recorded in a foreign currency (rate ~90), but Razorpay's amount is rupees already. With no rate available the
    amount is used as is."""
    if (currency or "").strip().upper() == "INR":
        return zamt
    rate = _rate_of(cand)
    return round(zamt * rate, 2) if rate else zamt


def _classify(zamt_inr, pamt, fee, fee_adjust):
    """Compare the Razorpay amount in INR (Amount (INR)) with the payment amount from the Zoho sheet (INR).
    Razorpay's amount includes its gateway fee, the Zoho payment does not, so with fee_adjust the fee is subtracted
    in the difference: difference = Amount (INR) - Payment Amount - Fees. The Payment Amount itself is left as is.
    The fee is already in INR (also for USD payments). FEE_TOL absorbs fee/GST rounding.
    Returns (payment_amount, difference, basis|None)."""
    fee = fee or 0.0
    if fee_adjust and fee:
        diff = round(zamt_inr - pamt - fee, 2)
        if abs(diff) <= FEE_TOL:
            return pamt, diff, "Net of Razorpay fee"
    raw = round(zamt_inr - pamt, 2)
    if abs(raw) < EXACT_TOL:  # Razorpay amount equals the payment amount without any fee
        return pamt, raw, "Exact"
    if fee_adjust and fee:
        return pamt, round(zamt_inr - pamt - fee, 2), None
    return pamt, raw, None


def run_match(fee_adjust: bool = True) -> dict:
    with conn() as c:
        pay_cols = get_meta(c, "payments_cols", [])
        has_invoice = "Invoice Number" in pay_cols  # older Zoho sheet exports may not include this column
        has_rate = "Exchange Rate" in pay_cols
        has_cur = CUR_COL in pay_cols
        inv_sel = ((', "Invoice Number"' if has_invoice else "") + (', "Exchange Rate"' if has_rate else "")
                   + (f', "{CUR_COL}"' if has_cur else ""))
        pay = c.execute(
            f'SELECT "CustomerPayment ID","InvoicePayment ID","Reference Number","_amount"{inv_sel} FROM payments'
        ).fetchall()
        by_ipid = _index(pay, "InvoicePayment ID", has_invoice, has_rate, has_cur)
        by_ref = _index(pay, "Reference Number", has_invoice, has_rate, has_cur)
        out, counts = [], defaultdict(int)
        zoho_has_cur = "currency" in {x.lower() for x in get_meta(c, "zoho_cols", [])}
        cur_sel = ', "currency"' if zoho_has_cur else ", NULL AS currency"
        for z in c.execute(f'SELECT "id","notes","_amount","_fee"{cur_sel} FROM zoho').fetchall():
            zamt = z["_amount"] or 0.0
            zcur = z["currency"]
            inv = notes_invoice_id(z["notes"])
            primary = by_ipid.get(inv) if inv else None
            ref = by_ref.get(z["id"])

            if primary:
                cand, by = primary, "InvoicePayment ID"
                pamt = round(sum(v[0] for v in primary.values()), 2)
                pamt, diff, basis = _classify(_inr(zamt, primary, zcur), pamt, z["_fee"], fee_adjust)
                # The notes field only ever names one invoice, so a mismatch here might just mean this
                # Razorpay payment actually covers several invoices sharing the same Reference Number.
                # If the sum across that whole group matches, use the group instead.
                if basis is None and ref and len(ref) > 1:
                    ref_pamt = round(sum(v[0] for v in ref.values()), 2)
                    ref_used, ref_diff, ref_basis = _classify(_inr(zamt, ref, zcur), ref_pamt, z["_fee"], fee_adjust)
                    if ref_basis is not None:
                        cand, by, pamt, diff, basis = ref, "Reference Number", ref_used, ref_diff, ref_basis
            elif ref:
                cand, by = ref, "Reference Number"
                pamt = round(sum(v[0] for v in ref.values()), 2)
                pamt, diff, basis = _classify(_inr(zamt, ref, zcur), pamt, z["_fee"], fee_adjust)
            else:
                cand = by = pamt = diff = basis = None

            if cand is None:
                # Missing: no counterpart in the Zoho sheet at all, so there is no invoice number to show.
                res, inv_no, rate, pcur = "Missing", None, None, None
            else:
                res = "Matched" if basis else "Mismatched"
                # Invoice number(s) attached to the matching Zoho-sheet row(s); several if the payment
                # covers multiple invoices. Set for Matched and Mismatched only, never for Missing.
                nums = sorted({v[1] for v in cand.values() if v[1]})
                inv_no = ", ".join(nums) if nums else None
                rate, pcur = _rate_cur_of(cand)
            counts[res] += 1
            out.append((z["id"], res, by, pamt, diff, basis, inv_no, rate, pcur))
        c.execute("DELETE FROM match_results")
        c.executemany("INSERT INTO match_results VALUES (?,?,?,?,?,?,?,?,?)", out)
        set_meta(c, "matched_at", now())
        set_meta(c, "fee_adjust", fee_adjust)
    return {"total": len(out), **{k: counts.get(k, 0) for k in ("Matched", "Mismatched", "Missing")}}