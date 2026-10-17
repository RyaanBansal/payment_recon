"""Matching logic for the HDFC bank statement.

HDFC's `Chq./Ref.No.` is the 16-digit zero-padded UPI reference ('0000581598905825' = UPI ref 581598905825).
The Zoho sheet (customer payments) holds it in one of two forms, so there are two steps:

1. Full reference: the Chq./Ref.No. without its leading zeros = Zoho `Reference Number` (also without leading
   zeros). Shown as "Chq./Ref.No." in Matched By.
2. Masked reference: recent HDFC receipts were recorded as 'xxxxxxxxxxxx5825' - only the last 4 digits are
   known, and 4 digits are not unique. So the payment must also be a customer payment deposited to HDFC, dated
   within 7 days of the bank date, and the one with the same amount wins (see match_common.pick_masked).
   Shown as "Chq./Ref.No. (last digits)" in Matched By.
3. Nothing found -> Missing.
4. Hit -> compare the transaction amount (Deposit Amt for credits, Withdrawal Amt for debits) with the customer
   payment amount (INR): equal = Matched, else Mismatched.

The Match tab starts on credits. Debits (card autopay, tax, own-account transfers) have no customer payment and
show as Missing.
"""
import re

from .match_common import (amount_check, build_masked_index, build_payment_index, invoice_numbers, pick_masked,
                           save_results, total_amount)

SOURCE = "hdfc"
BY_REF = "Chq./Ref.No."
BY_REF_MASKED = "Chq./Ref.No. (last digits)"
WINDOW_DAYS = 7


def _no_zeros(ref: str) -> str:
    return re.sub(r"^0+", "", ref)


def run(c) -> dict:
    exact, _ = build_payment_index(c)
    full = {}
    for ref, cands in exact.items():  # Zoho references compared without leading zeros, like the bank's
        full.setdefault(_no_zeros(ref), {}).update(cands)
    masked = build_masked_index(c, "hdfc")
    out = []
    for r in c.execute('SELECT "_key","_ref","_amount","_dt" FROM hdfc'):
        ref, amount = r["_ref"], r["_amount"] or 0.0
        cand, by = None, None
        if ref:
            if ref in full:
                cand, by = full[ref], BY_REF
            elif ref.isdigit() and len(ref) >= 10:  # UPI references only (12 digits); NEFT/RTGS ids are never masked
                cand = pick_masked(masked, ref, amount, r["_dt"], WINDOW_DAYS)
                by = BY_REF_MASKED
        if cand is None:
            out.append((r["_key"], "Missing", None, None, None, None, None))
            continue
        pay_amount = total_amount(cand)
        diff, basis = amount_check(amount, pay_amount)
        out.append((r["_key"], "Matched" if basis else "Mismatched", by, pay_amount, diff, basis, invoice_numbers(cand)))
    return save_results(c, SOURCE, out)
