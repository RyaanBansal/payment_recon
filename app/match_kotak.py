"""Matching logic for the KOTAK bank statement.

Matching id:  Kotak `Chq /Ref No.`  ->  Zoho sheet (customer payments) `Reference Number`.

For every Kotak row:
1. Look for its Chq/Ref No. in the customer payments' Reference Number (case-insensitive, whitespace ignored).
2. No hit -> try again ignoring the channel prefix (NEFTINW-1701994307 also finds 1701994307), since the
   payment may have been recorded without it. Shown in "Matched By" so it is visible when this was needed.
3. No hit at all, or the Kotak row has no Chq/Ref No. -> Missing.
4. Hit -> compare the Kotak Amount with the customer payment amount (INR): equal = Matched, else Mismatched.

Debit rows are matched like any other row; the Match tab shows credits by default (customer payments are credits).
"""
from .match_common import (amount_check, build_payment_index, invoice_numbers, loose_ref, save_results,
                           total_amount)

SOURCE = "kotak"
BY_REF = "Chq/Ref No."
BY_REF_LOOSE = "Chq/Ref No. (prefix ignored)"


def run(c, table: str = "kotak", source: str = SOURCE) -> dict:
    """`table` / `source` let the Kotak matcher run against the same rules on a different table if needed."""
    exact, loose = build_payment_index(c)
    out = []
    for r in c.execute(f'SELECT "_key","_ref","_amount" FROM "{table}"'):
        ref, amount = r["_ref"], r["_amount"] or 0.0
        cand, by = None, None
        if ref:
            if ref in exact:
                cand, by = exact[ref], BY_REF
            else:
                key = loose_ref(ref)
                if key in loose:
                    cand, by = loose[key], BY_REF_LOOSE
        if cand is None:
            out.append((r["_key"], "Missing", None, None, None, None, None))
            continue
        pay_amount = total_amount(cand)
        diff, basis = amount_check(amount, pay_amount)
        out.append((r["_key"], "Matched" if basis else "Mismatched", by, pay_amount, diff, basis, invoice_numbers(cand)))
    return save_results(c, source, out)