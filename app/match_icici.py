"""Matching logic for the ICICI bank statement.

Matching id:  ICICI `Tran. Id`  ->  Zoho sheet (customer payments) `Reference Number`
(case-insensitive, whitespace ignored). The Zoho sheet records ICICI receipts with the bank's transaction id
(S12345678), not with the NEFT/IMPS reference inside the remarks.

For every ICICI row:
1. Look for its Tran. Id in the customer payments' Reference Number. No hit -> Missing.
2. Hit -> compare the transaction amount (Deposit Amt for credits, Withdrawal Amt for debits) with the customer
   payment amount (INR): equal = Matched, else Mismatched (Difference = bank amount - Payment Amount).

The Match tab starts on credits (customer payments). Settlement credits from Cash Free and transfers between
own accounts have no customer payment with that id, so they show as Missing.
"""
from .match_common import amount_check, build_payment_index, invoice_numbers, save_results, total_amount

SOURCE = "icici"
BY_TRAN_ID = "Tran. Id"


def run(c) -> dict:
    exact, _ = build_payment_index(c)
    out = []
    for r in c.execute('SELECT "_key","_ref","_amount" FROM icici'):
        cand = exact.get(r["_ref"]) if r["_ref"] else None
        if cand is None:
            out.append((r["_key"], "Missing", None, None, None, None, None))
            continue
        pay_amount = total_amount(cand)
        diff, basis = amount_check(r["_amount"] or 0.0, pay_amount)
        out.append((r["_key"], "Matched" if basis else "Mismatched", BY_TRAN_ID, pay_amount, diff, basis, invoice_numbers(cand)))
    return save_results(c, SOURCE, out)
