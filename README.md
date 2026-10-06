# Payment - Zoho Reconciliation

FastAPI + plain HTML/JS, Python backend, local SQLite (`data/recon.db`, created on first run).

## Run
```bash
pip install -r requirements.txt
uvicorn app.main:app --reload        # then open http://127.0.0.1:8000
```

## How it works
1. Upload the sheets (several files at once is fine), then click **Upload** in that box (loads it into SQLite). A line under each box shows what is in the database.
   - **Zoho sheet** box (customer payments) = full list -> each upload **replaces** the previous one.
   - **Monthly payment sheets** box: pick the sheet type in the dropdown (Razorpay, Kotak, Cash Free) -> uploads **accumulate**; re-uploading a month is safe.
     You only need to upload the ones you want to reconcile.
2. Click **Match**. Every monthly sheet that has been uploaded is matched against the Zoho sheet. Results appear in the
   **Match / Mismatch** tab, one sub tab per sheet (Razorpay / Kotak / Cash Free), each with Matched / Mismatched / Missing chips,
   its own filters, Agent / Remarks and Excel download.
3. Original tables, one main tab each: **Zoho Sheet**, **Razorpay Sheet**, **Kotak Sheet**, **Cash Free Sheet**
   (Zoho: dates `yyyy-mm-dd`; Razorpay: `dd/mm/yyyy`; Kotak and Cash Free: `dd-mm-yyyy` - always as in the sheet).

## Clearing data
Every tab has a red **Clear data** button (asks for confirmation; cannot be undone):
- **Zoho Sheet / Razorpay Sheet / Kotak Sheet / Cash Free Sheet** tabs: delete that uploaded sheet from the database, plus its Match / Mismatch table and the agent / remarks typed on it.
  Clearing the **Zoho** sheet also empties all Match / Mismatch tables (they were computed against it); remarks on the other sheets are kept until those sheets are cleared.
- **Match / Mismatch sub tabs** (Razorpay / Kotak / Cash Free): delete only that sheet's match table and its agent / remarks. The uploaded sheet stays - click **Match** to rebuild.

## Files
| File | Role |
|---|---|
| `app/ingest.py` | Zoho (customer payments) + Razorpay loaders |
| `app/ingest_sheets.py` | Kotak (`.csv`/`.xlsx`) + Cash Free (`.xlsx`, all month tabs) loaders |
| `app/matching.py` | Razorpay matching logic |
| `app/match_kotak.py` | Kotak matching logic |
| `app/match_cashfree.py` | Cash Free matching logic |
| `app/match_common.py` | shared by the Kotak / Cash Free matchers (payment lookup, amount check, saving results) |

To add another sheet: write `ingest` + `match_<sheet>.py` (a `run(c)` that returns the counts), register it in `MONTHLY`,
`MATCHERS` and `SOURCES` in `app/main.py`, and add it to `MONTHLY_UPLOADS` (upload dropdown) and `SRC` in `static/index.html`.

## Matching rules - Kotak (app/match_kotak.py)
Matching id: Kotak **Chq /Ref No.** = Zoho sheet **Reference Number** (case-insensitive).
1. No exact hit -> try again without the channel prefix (`NEFTINW-1701994307` also finds `1701994307`); shown as *Chq/Ref No. (prefix ignored)* in Matched By.
2. No hit, or the Kotak row has no Chq/Ref No. -> **Missing**. Hit -> Kotak **Amount** vs the payment amount: same = **Matched**, else **Mismatched**.
3. Credits are what customers pay, so the Match tab starts on `Dr / Cr = CR`; pick All / DR / (blank) to see the rest.
- Lines without an amount (the "Closing Balance" footer) are skipped on upload. Refs Excel turned into numbers (`551,781,154.00`) are restored to digits for matching; the table keeps what the sheet shows.

## Matching rules - Cash Free (app/match_cashfree.py)
Matching id: Cash Free **Order Id** = Zoho sheet **Reference Number** (case-insensitive). No hit -> **Missing**. Hit -> compare with the payment amount:
**Order Amount** equal = Matched; else **Transaction Amount** equal (customer paid gateway charges on top) = Matched, basis *Transaction Amount*; else **Mismatched**
(Difference = Order Amount - Payment Amount). The Match tab starts on `Transaction status = SUCCESS`.
- All month tabs of the workbook are read; their columns are merged (tabs don't all have the same columns). Dates Excel swapped (day/month) are corrected, as for Razorpay.

For both: the payment amount is the sum of the distinct customer payments (CustomerPayment ID) with that reference, in INR (amount x exchange rate).

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
`/api/status`, `POST /api/upload/{payments|zoho|kotak|cashfree}`, `POST /api/match`, `/api/matches` (Razorpay),
`/api/recon/{kotak|cashfree}/matches`, `/api/recon/{kotak|cashfree}/export`, `/api/sheet/{kotak|cashfree}`,
`POST /api/clear/sheet/{payments|zoho|kotak|cashfree}`, `POST /api/clear/matches/{razorpay|kotak|cashfree}`, `POST /api/annotation`, `/api/agents`, `/api/payments`, `/api/zoho`  (interactive docs at `/docs`).