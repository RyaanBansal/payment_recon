"""Pieces shared by the per-sheet matching files (match_kotak.py, match_cashfree.py, ...).

Each of those files holds the matching logic for ONE sheet; what they have in common lives here:
looking up the customer payments (Zoho sheet) by Reference Number, comparing amounts, saving the results.
"""
import re
from collections import defaultdict

from .db import get_meta, set_meta, now

EXACT_TOL = 0.005


# ---------- reference normalisation ----------
def norm_ref(value) -> str:
    """Comparable form of a reference / id: trimmed, upper-case. A purely numeric id that Excel re-formatted
    as a number ('551,781,154.00') is turned back into its digits ('551781154')."""
    if value is None:
        return ""
    s = str(value).strip()
    if re.fullmatch(r"\d{1,3}(,\d{3})+(\.0+)?|\d+\.0+", s):
        s = s.replace(",", "").split(".")[0]
    return s.upper()


def loose_ref(ref: str) -> str:
    """Bank references carry a channel prefix ('NEFTINW-1701994307', 'UPI-624409102382'). The payment may have
    been recorded with or without it, so this drops the prefix (anything before the first '-') when what is
    left still looks like a real id."""
    if "-" in ref:
        tail = ref.split("-", 1)[1]
        if len(tail) >= 6:
            return tail
    return ref


# ---------- the customer payments (Zoho sheet) ----------
def build_payment_index(c):
    """Index the customer payments (Zoho sheet) by Reference Number.

    Returns (exact, loose): each maps reference -> {CustomerPayment ID: (amount in INR, Invoice Number)}.
    Keyed by distinct CustomerPayment ID (not by row) so a payment split across several invoices is
    counted once per payment, and its amounts add up to the whole receipt.
    """
    cols = get_meta(c, "payments_cols", [])
    has_inv = "Invoice Number" in cols
    sel = 'SELECT "CustomerPayment ID","Reference Number","_amount"' + (', "Invoice Number"' if has_inv else "") + " FROM payments"
    exact, loose = defaultdict(dict), defaultdict(dict)
    for n, r in enumerate(c.execute(sel)):
        ref = norm_ref(r["Reference Number"])
        if not ref:
            continue
        cp = r["CustomerPayment ID"] or f"row{n}"
        val = (r["_amount"] or 0.0, r["Invoice Number"] if has_inv else None)
        exact[ref][cp] = val
        loose[loose_ref(ref)][cp] = val
    return exact, loose


def total_amount(cand) -> float:
    return round(sum(v[0] for v in cand.values()), 2)


def invoice_numbers(cand):
    nums = sorted({v[1] for v in cand.values() if v[1]})
    return ", ".join(nums) if nums else None


def amount_check(amount, payment_amount, alt_amount=None, alt_label=None):
    """Compare a sheet amount with the customer payment amount.
    Returns (difference, basis): basis is None when the amounts do not agree (-> Mismatched)."""
    diff = round(amount - payment_amount, 2)
    if abs(diff) < EXACT_TOL:
        return diff, "Exact"
    if alt_amount is not None:
        alt_diff = round(alt_amount - payment_amount, 2)
        if abs(alt_diff) < EXACT_TOL:
            return alt_diff, alt_label
    return diff, None


# ---------- saving ----------
def save_results(c, source: str, rows) -> dict:
    """rows: (row_key, result, matched_by, payment_amount, difference, basis, invoice_number).
    Replaces the earlier results of this sheet and returns the counts."""
    counts = defaultdict(int)
    for r in rows:
        counts[r[1]] += 1
    c.execute("DELETE FROM src_match_results WHERE source = ?", (source,))
    c.executemany("INSERT INTO src_match_results VALUES (?,?,?,?,?,?,?,?)", [(source,) + tuple(r) for r in rows])
    set_meta(c, f"{source}_matched_at", now())
    return {"total": len(rows), **{k: counts.get(k, 0) for k in ("Matched", "Mismatched", "Missing")}}
