# QBO ↔ TicketVault Reconciliation

Web app: upload the QuickBooks **Profit and Loss Detail** and the matching **TicketVault**
export (`.xlsx`, `.xlsm`, or `.csv`) for each entity. Pick the company for each TicketVault
file from a drop-down (pre-filled from the filename when possible) so the two are paired
exactly. Upload as many pairs as you like in one go. Download a lean workbook that flags any
**sales** (by marketplace × day) or **cost** (by day) that don't agree, over a tolerance
($0.01 by default).

Built on the same Flask + single-page-UI + Railway pattern as the Ledger Reconciliation app.

## The two inputs

1. **QuickBooks — Profit and Loss Detail** (per entity). The account tree is read by
   indentation:
   - **Income → Sales** — one *Invoice* line per marketplace per day, `Name` like
     `Stubhub (C)`, `Vivid Seats (C)`. This is the QBO sales figure.
   - **Cost of Goods Sold** — one *Journal Entry* per day (a daily aggregate; QBO has **no
     per-marketplace cost**). This is the QBO cost figure.
   - **Foreign Exchange Conversion** — **excluded entirely** (its lines are never read).
   - The entity name is taken from the report's title row.
2. **TicketVault export.** Repeating daily blocks: a day row (`-> 07/01/2026`, the day
   total) followed by per-client rows (TicketsNow, Vivid Seats, StubHub, …). The Sales
   **Net** column (Sold − Cancelled) is the sales figure; the Cost **Net** column is the
   cost figure. Client cost is summed per day to line up with QBO's daily COGS.

## What it does

For each paired entity, **neither book is treated as the source of truth** — both values
are shown side by side and any divergence over the tolerance is flagged.

1. **Sales — marketplace × day.** QBO's per-marketplace Invoice amount is matched to
   TicketVault's Sales-Net for the same marketplace and day. A marketplace present in only
   one book (non-zero) is flagged as `QBO only` / `TicketVault only`; a real difference is
   `MISMATCH`. Marketplace names are matched case-insensitively with the `(C)` / `(CAD)`
   markers stripped (`Stubhub (C)` ↔ `StubHub`).
2. **Cost — day.** QBO's daily COGS journal is matched to TicketVault's total Cost-Net for
   the day. Because QBO records cost only as a daily aggregate, cost is checked per day,
   not per marketplace.

### Pairing (batch)

Each TicketVault file is assigned to a company by a **drop-down** in the upload list. The
drop-down shows a **fixed company roster** — the same names and order as the Season Ticket
Buy-In Review app (`Y&S, Grossman, Sternbuch, Pollak, Levine, Levovitz, Chase, Asher, Katz,
GK, TL, Waxler, TTG, YourTix`), served by `/options` so it can be edited in one place. Each
file's pick is **pre-filled from its filename** when that matches a roster name. Whatever is
selected is **authoritative**.

Internally the app resolves each uploaded QBO file's entity to a roster company using the
**master mapping list** (`Master_Mapping_List.xlsx`, bundled next to `app.py`), which maps
each QBO company name (and TicketVault spelling) to its short name — e.g. *YS Levine LLC* →
**Levine**, *YS TL Tickets* → **TL** (not Y&S, even though both contain "YS"). It then pairs
the TicketVault file to the QBO with the same company. Matching is on the distinctive core
of the name, so `LLC`/`Tickets`/etc. differences don't matter; replace the bundled file to
update the mapping. The Reconcile button stays disabled until every TicketVault file has a
company chosen. If a file is left unassigned (e.g. via the API), the app falls back to
reading the company from the file, then the filename, then period/total proximity.

### Output tabs

- **Sales Discrepancies** — only the flagged entity × marketplace × day rows: QBO,
  TicketVault, `Diff (QBO−TV)`, and a `Status`, with a live `=SUM` total.
- **Cost Discrepancies** — only the flagged entity × day rows, same shape.

If a tab has nothing to show, it says so instead of listing rows. (There is no Summary tab,
and the page shows no notes — just the download button.)

## Tuning the match

Edit the config block at the top of `app.py`:

- `TOLERANCE` — dollar threshold for "equal" (default `0.01`).
- `COMPANIES` — the roster shown in the TicketVault drop-down (names **and order**). Served
  to the page by `/options`.
- `COMPANY_ALIASES` — force a QBO entity onto a specific roster company when the automatic
  match is wrong or missing (rarely needed — the master mapping list covers the known set).
- `Master_Mapping_List.xlsx` — the authoritative QBO-company → short-name table, read at
  startup. Update company names by editing this file and redeploying; `COMPANY_MAP_RAW` in
  `app.py` is a baked-in fallback used if the file is absent.
- `MARKETPLACE_ALIASES` — force a marketplace spelled differently across the two systems
  to be treated as one (the common names already match once `(C)` is stripped).
- `NON_MARKETPLACE_CLIENTS` — TicketVault client rows that aren't QBO sales marketplaces
  (e.g. *Baseball Transfers*, *Expired*, *Transfers - Yankees*). Their **cost** still
  counts toward the day total; they're just not flagged as missing on the sales side.
- `ENTITY_ALIASES` — map a QBO entity name or a TicketVault filename to a canonical entity
  so the two pair up. Add an entry whenever the Summary tab shows a file as unpaired.

## Input format

The app reads the **raw** exports and auto-detects the header row (extra title / `As of`
rows above the table are fine). Amounts may be plain numbers, `$1,234.56`, or `(123)`
negatives; `MM/DD/YYYY` text and real date cells are both normalized.

## Run locally

```bash
pip install -r requirements.txt
python app.py        # http://localhost:5000
```

## Deploy: GitHub → Railway

1. Push this folder to a GitHub repo.
2. Railway → New Project → Deploy from GitHub repo → pick it.
3. Railway auto-detects Python (Nixpacks) and uses the start command in `railway.json`.
   No env vars needed; `$PORT` is provided automatically.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask backend — parsing, pairing, reconciliation, workbook builder |
| `Master_Mapping_List.xlsx` | QBO-company → short-name mapping (read at startup) |
| `index.html` | Single-page upload UI (two multi-file slots) |
| `requirements.txt` | Python dependencies |
| `Procfile` / `railway.json` | Start command for Railway |
