# Payment - Zoho Reconciliation

FastAPI + plain HTML/JS, Python backend, local SQLite (`data/recon.db`, created on first run).

## Run
```bash
pip install -r requirements.txt
uvicorn app.main:app --reload        # then open http://127.0.0.1:8000
```

## How it works
1. Choose each sheet in its upload box and click its **Upload** button (loads it into SQLite).
   - Customer Payments sheet = full list -> each upload **replaces** the previous one.
   - Zoho/Razorpay sheet = monthly -> uploads **accumulate** (rows upserted on `id`, so re-uploading a month is safe).
2. Click **Match**. Results appear in the **Match / Mismatch** tab (Matched / Mismatched / Missing chips with counts).
3. **Payments Sheet** tab: filter by Mode, Deposit To, date (`yyyy-mm-dd`, as in the sheet).
   **Zoho Sheet** tab: filter by method, date (`dd/mm/yyyy`, as in the sheet).

## Matching rules (app/matching.py)
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
`/api/status`, `POST /api/upload/{payments|zoho}`, `POST /api/match`, `/api/matches`, `POST /api/annotation`,
`/api/agents`, `/api/payments`, `/api/zoho`  (interactive docs at `/docs`).
