"""Matching logic for the CASH FREE transactions sheet.

Matching id:  Cash Free `Order Id`  ->  Zoho sheet (customer payments) `Reference Number`.

For every Cash Free row:
1. Look for its Order Id in the customer payments' Reference Number (case-insensitive, whitespace ignored).
2. No hit (or no Order Id) -> Missing.
3. Hit -> compare amounts with the customer payment amount (INR):
   - Order Amount equal            -> Matched ("Order Amount")
   - else Transaction Amount equal -> Matched ("Transaction Amount"): the customer paid the order amount plus
     gateway / convenience charges, and the payment was recorded with the charges included
   - neither                       -> Mismatched (Difference = Order Amount - Payment Amount)

Only SUCCESS transactions are meaningful for reconciliation; the Match tab shows those by default.
"""
from .match_common import (amount_check, build_payment_index, invoice_numbers, save_results, total_amount)

SOURCE = "cashfree"
BY_ORDER_ID = "Order Id"


def run(c) -> dict:
    exact, _ = build_payment_index(c)
    out = []
    for r in c.execute('SELECT "_key","_ref","_amount","_txn_amount" FROM cashfree'):
        cand = exact.get(r["_ref"]) if r["_ref"] else None
        if cand is None:
            out.append((r["_key"], "Missing", None, None, None, None, None))
            continue
        pay_amount = total_amount(cand)
        diff, basis = amount_check(r["_amount"] or 0.0, pay_amount, r["_txn_amount"], "Transaction Amount")
        if basis == "Exact":
            basis = "Order Amount"
        out.append((r["_key"], "Matched" if basis else "Mismatched", BY_ORDER_ID, pay_amount, diff, basis, invoice_numbers(cand)))
    return save_results(c, SOURCE, out)
