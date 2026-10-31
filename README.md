# BusinessOP

A single-page **point-of-sale (POS) + inventory manager** for small shops.
Built as a **Google Apps Script web app** with a **Google Sheets** database.

Money is stored in **paise (integers)** on both client and server to avoid
floating-point rounding errors.

## Features

- Products with cost, price, stock and low-stock alerts
- Customers & vendors (10-digit phone validation, duplicate-name guard)
- Invoices — full / part / pay-later modes
- Record payments, view balances per customer
- Receive stock from vendors → purchase log
- KPI dashboard + 7-day sales bar chart + top sellers
- Mobile-first layout (safe-area aware), desktop sidebar
- Google Sheets as the database (1 spreadsheet = 1 company)

## Tech stack

| Layer     | Tech                                            |
| --------- | ----------------------------------------------- |
| Frontend  | Vanilla JS, single HTML file, no build step     |
| Backend   | Google Apps Script (V8)                         |
| Database  | Google Sheets                                   |
| Bridge    | `google.script.run` + `HtmlService` templating  |

## First-time setup (5 minutes)

### 1. Create the Apps Script project

1. Go to <https://script.google.com> → New project
2. Rename it to **BusinessOP**
3. Create two files inside the project:
   - `Code.gs` → paste the contents of `Code.gs`
   - File → New → HTML file, name it **`Index`** (capital I) → paste the contents of `Index.html`

### 2. Run `setup()`

- Select `setup` from the function dropdown → Run
- Approve permissions
- Check Executions / Logs for the new Google Sheet URL

### 3. (Optional) Seed demo data

- Run `seedDemoData()` once to load sample products, customers, vendors

### 4. Deploy as a Web App

- Deploy → New deployment → Web app
  - Execute as: **Me**
  - Who has access: **Only myself** (or *Anyone* for a public link)
- Copy the Web app URL

### 5. Open the URL

That's it. The app pulls the bootstrap JSON and works offline for search/cart.

## Local development with `clasp` (optional)

```bash
npm install -g @google/clasp
clasp login
clasp clone <YOUR_SCRIPT_ID>
clasp push
