# Payment - Zoho Reconciliation

FastAPI + plain HTML/JS, Python backend, local SQLite (`data/recon.db`, created on first run).

## Run
```bash
pip install -r requirements.txt
uvicorn app.main:app --reload        # then open http://127.0.0.1:8000
```
(`xlrd` is needed to read HDFC's old `.xls` statement format.)

## How it works
1. Upload the sheets (several files at once is fine), then click **Upload** in that box (loads it into SQLite). A line under each box shows what is in the database.
   - **Zoho sheet** box (customer payments) = full list -> each upload **replaces** the previous one.
   - **Monthly payment sheets** box: pick the sheet type in the dropdown (Razorpay, Kotak, Cash Free, ICICI, HDFC, Kotak statement) -> uploads **accumulate**; re-uploading a month / period is safe.
     You only need to upload the ones you want to reconcile.
2. Click **Match**. Every monthly sheet that has been uploaded is matched against the Zoho sheet. Results appear in the
   **Results** tab, one sub tab per sheet (Razorpay / Kotak / Cash Free / ICICI / HDFC / Kotak Statement), each with Matched / Mismatched / Missing chips,
   its own filters, Agent / Remarks and Excel download.
3. Original tables, one main tab each: **Zoho Sheet**, **Razorpay Sheet**, **Kotak Sheet**, **Cash Free Sheet**, **ICICI Sheet**, **HDFC Sheet**, **Kotak Statement Sheet**
   (Zoho: dates `yyyy-mm-dd`; Razorpay: `dd/mm/yyyy`; all the others: `dd-mm-yyyy`).

## Clearing data
Every tab has a red **Clear data** button (asks for confirmation; cannot be undone):
- **Zoho Sheet / Razorpay Sheet / Kotak / Cash Free / ICICI / HDFC / Kotak Statement Sheet** tabs: delete that uploaded sheet from the database, plus its Match / Mismatch table and the agent / remarks typed on it.
  Clearing the **Zoho** sheet also empties all Match / Mismatch tables (they were computed against it); remarks on the other sheets are kept until those sheets are cleared.
- **Match / Mismatch sub tabs**: delete only that sheet's match table and its agent / remarks. The uploaded sheet stays - click **Match** to rebuild.

## Files
| File | Role |
|---|---|
| `app/ingest.py` | Zoho (customer payments) + Razorpay loaders |
| `app/ingest_sheets.py` | Kotak (`.csv`/`.xlsx`), Kotak statement, ICICI (`.xlsx`), HDFC (`.xls`/`.xlsx`) and Cash Free (`.xlsx`, all month tabs) loaders |
| `app/matching.py` | Razorpay matching logic |
| `app/match_kotak.py` | Kotak matching logic |
| `app/match_kotak_stmt.py` | Kotak statement matching logic (same rules as Kotak, own table) |
| `app/match_icici.py` | ICICI matching logic |
| `app/match_hdfc.py` | HDFC matching logic |
| `app/match_cashfree.py` | Cash Free matching logic |
| `app/match_common.py` | shared by the per-sheet matchers (payment lookup, masked-reference lookup, amount check, saving results) |

To add another sheet: write `ingest` + `match_<sheet>.py` (a `run(c)` that returns the counts), register it in `MONTHLY`,
`MATCHERS`, `SOURCES` (and `SHEET_KINDS` / `MATCH_SOURCES`) in `app/main.py`, and add it to `MONTHLY_UPLOADS` (upload dropdown), `SRC`, the tab / sub tab / panel
HTML and the `activate` / `loadSub` maps in `static/index.html`.

## Bank statements (ICICI, HDFC)
These are read as downloaded from the bank: the account details above the table and the notes / totals below it are ignored (the header row is found
automatically). Dates are shown as `dd-mm-yyyy`. Each row gets a `Dr / Cr` value (ICICI and HDFC: a deposit is CR, a withdrawal DR) and the Match tab
starts on `CR`, because customer payments are credits; pick All / DR to see the rest. Lump-sum credits (Razorpay / Cash Free settlements, transfers between
own accounts) and all debits have no customer payment, so they show as **Missing** - that is expected.
For every bank sheet the payment amount is the sum of the distinct customer payments (CustomerPayment ID) with that reference, in INR (amount x exchange rate),
compared with the bank amount: same = **Matched**, else **Mismatched** (Difference = bank amount - Payment Amount).

## Matching rules - Kotak (app/match_kotak.py)
Matching id: Kotak **Chq /Ref No.** = Zoho sheet **Reference Number** (case-insensitive).
1. No exact hit -> try again without the channel prefix (`NEFTINW-1701994307` also finds `1701994307`); shown as *Chq/Ref No. (prefix ignored)* in Matched By.
2. No hit, or the Kotak row has no Chq/Ref No. -> **Missing**. Hit -> Kotak **Amount** vs the payment amount: same = **Matched**, else **Mismatched**.
3. Credits are what customers pay, so the Match tab starts on `Dr / Cr = CR`; pick All / DR / (blank) to see the rest.
- Lines without an amount (the "Closing Balance" footer) are skipped on upload. Refs Excel turned into numbers (`551,781,154.00`) are restored to digits for matching; the table keeps what the sheet shows.
- The matched bank statement sources remain the standard monthly bank sheets used in the app; the separate Kotak statement sheet has been removed.

## Matching rules - ICICI (app/match_icici.py)
Matching id: ICICI **Tran. Id** (`S12345678`) = Zoho sheet **Reference Number** (case-insensitive). The Zoho sheet records ICICI receipts with this id, not with the NEFT/IMPS reference in the remarks.
No hit -> **Missing**. Hit -> bank amount (Deposit Amt for credits, Withdrawal Amt for debits) vs the payment amount.

## Matching rules - HDFC (app/match_hdfc.py)
HDFC **Chq./Ref.No.** is the 16-digit zero-padded UPI reference (`0000581598905825` = `581598905825`).
1. Without its leading zeros it equals a Zoho **Reference Number** -> *Chq./Ref.No.* in Matched By.
2. Recent HDFC receipts are recorded in Zoho as `xxxxxxxxxxxx5825` (only the last 4 digits). Last digits are not unique, so the candidate must also be a customer payment
   deposited to **HDFC Bank**, dated within **7 days** of the bank date; the one with the same amount wins (nearest date if several). Only UPI-style numeric references are tried this way.
   Shown as *Chq./Ref.No. (last digits)* in Matched By. If nothing has the same amount, a single candidate is still shown as **Mismatched**; several candidates with different amounts are ambiguous -> **Missing**.
3. No hit -> **Missing**. Hit -> amount comparison as above.

## Matching rules - Cash Free (app/match_cashfree.py)
Matching id: Cash Free **Order Id** = Zoho sheet **Reference Number** (case-insensitive). No hit -> **Missing**. Hit -> compare with the payment amount:
**Order Amount** equal = Matched; else **Transaction Amount** equal (customer paid gateway charges on top) = Matched, basis *Transaction Amount*; else **Mismatched**
(Difference = Order Amount - Payment Amount). The Match tab starts on `Transaction status = SUCCESS`.
- All month tabs of the workbook are read; their columns are merged (tabs don't all have the same columns). Dates Excel swapped (day/month) are corrected, as for Razorpay.

## Matching rules - Razorpay (app/matching.py)
For every Zoho row:
1. Take `invoice_id_0` from `notes` and look for it in the payments sheet's **InvoicePayment ID**.
2. If there is no invoice id, or it finds nothing -> fall back to Zoho **id** = payments **Reference Number**.
3. No hit in either -> **Missing**. Hit -> compare amounts: same = **Matched**, different = **Mismatched**.

Amount comparison: Zoho `amount` vs the customer payment's `Amount`. With the "net of Razorpay fee" box ticked
(default), `amount - fee` equal to the payment amount (within Rs 1) also counts as Matched, because Razorpay's
amount includes the gateway fee. Untick for strict equality.

Zoho **Agent** (dropdown, default *Unassigned*, only on Mismatched/Missing) and **Remarks** are saved in the
`annotations` table by payment id, so they survive re-uploads and re-matching.

## Adding the agent names later
`POST /api/agents {"name": "Asha"}` or `sqlite3 data/recon.db "INSERT INTO agents VALUES ('Asha')"`.
The dropdown picks them up automatically.

## API
`/api/status`, `POST /api/upload/{payments|zoho|kotak|cashfree|icici|hdfc|kotakstmt}`, `POST /api/match`, `/api/matches` (Razorpay),
`/api/recon/{kotak|cashfree|icici|hdfc|kotakstmt}/matches`, `/api/recon/{...}/export`, `/api/sheet/{...}`,
`POST /api/clear/sheet/{payments|zoho|kotak|cashfree|icici|hdfc|kotakstmt}`, `POST /api/clear/matches/{razorpay|kotak|cashfree|icici|hdfc|kotakstmt}`, `POST /api/annotation`, `/api/agents`, `/api/payments`, `/api/zoho`  (interactive docs at `/docs`).