# Elave Ecommerce TV Dashboard

Minimal web dashboard for office TV display with:

- KPI cards: Total Sales, Orders, AOV, ROAS
- Metric chart with switchable chart style (line, bar, area, doughnut)
- 15-minute refresh (configurable)
- Simple browser passcode screen
- Google Apps Script integration (no complex backend)

## 1) Local preview

1. Keep `config.js` as-is (empty `appsScriptUrl`).
2. Open `index.html` in a browser.
3. Enter passcode from `config.js`.

This uses `sample-data.json`.

## 2) Connect your Google Sheet (Apps Script)

1. Open your Google Sheet.
2. Go to `Extensions` -> `Apps Script`.
3. Paste the code from `google-apps-script/Code.gs`.
4. Save and deploy:
1. Click `Deploy` -> `New deployment`.
2. Type: `Web app`.
3. Execute as: `Me`.
4. Who has access: your org users or anyone with link (internal use only).
5. Copy the `/exec` URL.
5. Put the URL into `config.js` as `appsScriptUrl`.

## 3) Configure dashboard settings

Edit `config.js`:

- `appsScriptUrl`: your Apps Script web app URL
- `sheetName`: tab name in the sheet
- `passcode`: browser lock code
- `currencySymbol`: keep as `EUR` (dashboard uses EUR only)
- `refreshIntervalMs`: set your preferred auto refresh (default in this repo is 15 minutes)
- `salesTargetsByMonth`: optional monthly targets map, e.g. `"2026-02": 120000`

You can also set/clear the current month target directly from the **Target Progress** card using the **Set Target** and **Clear** buttons.

## 4) Deploy (low cost)

Easiest internal hosting options:

1. Netlify (drag-and-drop site folder, free tier).
2. Vercel (connect folder/repo, free tier).
3. Shared office machine local browser tab in fullscreen.

For TV display:

- Open dashboard URL
- Press `F11` fullscreen
- Leave browser tab running

## 5) Supabase + Shopify mode (recommended next step)

This repo now includes a small backend that:

- pulls `sales/orders` from Shopify
- pulls `roas/ad_spend` from Apps Script (`mode=clean`)
- merges hourly rows into Supabase table `hourly_metrics`
- serves `/api/clean` for the dashboard

### A) Create Supabase table

1. Open Supabase SQL editor.
2. Run `supabase/schema.sql`.

### B) Configure backend env

1. Copy `backend/.env.local.example` to `backend/.env.local`.
2. Fill in:
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - `APPS_SCRIPT_URL`
   - `SHOPIFY_STORE_DOMAIN`
   - `SHOPIFY_ACCESS_TOKEN`
   - `REPORTING_TIMEZONE` (optional, example: `Europe/Amsterdam`; default `UTC`)

### C) Run first sync

```bash
npm run sync:supabase
```

### D) Start API server

```bash
npm run serve:api
```

This serves:

- `GET /api/health`: service health
- `GET /api/endpoints`: list of available API routes
- `GET /api/clean?days=120`: normalized hourly rows (primary dashboard feed)
- `GET /api/latest`: most recent row
- `GET /api/summary`: MTD summary + previous period comparison
- `GET /api/kpis`: alias of summary (for KPI cards)
- `GET /api/cells`: combined payload for all dashboard cells
- `GET /api/trend/hourly?days=30`: hourly trend series
- `GET /api/trend/daily?days=90`: daily aggregated trend
- `GET /api/quality?days=30`: data quality checks (missing hours, gaps, coverage)
- `GET /api/sources?days=30`: source coverage breakdown (Shopify vs Apps Script)
- `GET /api/products/top-units?limit=10`: Top Products (MTD by Units)
- `GET /api/products/top-revenue?limit=10`: Top Products (MTD by Revenue)
- `GET /api/products/momentum?metric=revenue&limit=10`: Product Momentum
- `GET /api/pace`: Daily Sales Pace vs Target
- `GET /api/projection`: MTD Progress + Projection
- `GET /api/finance/gross-net-returns`: Gross Sales / Net Sales / Returns
- `GET /api/aov`: MTD AOV + trend vs previous period
- `GET /api/sessions/mtd`: Website Sessions MTD + previous month MTD (same period)
- `GET /api/customers/new-vs-returning`: New vs Returning Customer Revenue
- `GET /api/channels`: Channel Split
- `GET /api/discount-impact`: Discount Impact
- `GET /api/heatmap/today`: Hourly Heatmap (Today)
- `GET /api/refund-watchlist?limit=10`: Refund Watchlist

### E) Point dashboard to backend API

In `config.js` set:

- `backendApiUrl: "http://localhost:8787"`

Keep `appsScriptUrl` set as fallback.

## 6) Production Setup (No Local Laptop Needed)

Use this for your office TV so it runs 24/7 in the cloud.

### A) Deploy backend API (Railway/Render/Fly)

1. Create a new service from this repo.
2. Start command: `npm run start`
3. Set environment variables:
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - `SUPABASE_TABLE` (usually `hourly_metrics`)
   - `SUPABASE_ORDERS_TABLE` (usually `shopify_orders`)
   - `SUPABASE_ORDER_LINES_TABLE` (usually `shopify_order_lines`)
   - `APPS_SCRIPT_URL`
   - `SHOPIFY_STORE_DOMAIN`
   - `SHOPIFY_ACCESS_TOKEN`
   - `SHOPIFY_API_VERSION` (example: `2024-10`)
   - `REPORTING_TIMEZONE` (optional, example: `Europe/Amsterdam`; default `UTC`)
   - `PORT` (platform usually sets this automatically)
4. Confirm health endpoint:
   - `https://<your-backend-domain>/api/health`

### B) Enable hourly sync in cloud (GitHub Actions)

This repo includes `.github/workflows/sync-supabase-hourly.yml`.

1. Push this repo to GitHub.
2. Add these GitHub Repository Secrets:
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - `SUPABASE_TABLE`
   - `SUPABASE_ORDERS_TABLE`
   - `SUPABASE_ORDER_LINES_TABLE`
   - `APPS_SCRIPT_URL`
   - `SHOPIFY_STORE_DOMAIN`
   - `SHOPIFY_ACCESS_TOKEN`
   - `SHOPIFY_API_VERSION`
   - `REPORTING_TIMEZONE` (optional, example: `Europe/Amsterdam`; default `UTC`)
   - `SYNC_DAYS`
3. The workflow runs hourly and can also be run manually from Actions.

### C) Deploy frontend (Vercel/Netlify)

1. Create a new static site from this repo.
2. Build command: `npm run build`
3. Publish directory: `dist`
4. Set frontend build env vars (used by `scripts/sync-config.mjs`):
   - `DASH_BACKEND_API_URL=https://<your-backend-domain>`
   - `DASH_APPS_SCRIPT_URL=<your-app-script-exec-url>`
   - `DASH_PASSCODE=<your-tv-passcode>`
   - `DASH_REFRESH_INTERVAL_MS=900000`
   - `DASH_SHEET_NAME=Triple Whale Hourly`
5. Redeploy and open the frontend URL on the office TV.

## Notes

- Passcode is a simple browser lock, not strong security.
- For stricter security, limit Apps Script deployment access to your Google Workspace users.
- MTD windows use `REPORTING_TIMEZONE` (default `UTC`) and Shopify sync excludes voided/cancelled/test orders.

## 7) Google Calendar Upcoming widget (optional)

The dashboard now includes a fixed widget: **Google Calendar Upcoming**.

### A) Backend env variables

Add these to `.env.local` (or your deployed backend env):

- `GOOGLE_OAUTH_CLIENT_ID`
- `GOOGLE_OAUTH_CLIENT_SECRET`
- `GOOGLE_OAUTH_REDIRECT_URI` (optional override)
- `GOOGLE_CALENDAR_ID` (optional, default `primary`)
- `GOOGLE_CALENDAR_REFRESH_TOKEN` (optional persistent token)

### B) OAuth redirect setup in Google Cloud

For local dev with API on port `8787`, your OAuth client should include:

- Authorized redirect URI: `http://localhost:8787/api/google/oauth/callback`

If your API runs on a different host/port, use:

- `https://<your-api-domain>/api/google/oauth/callback`

### C) Connect once

1. Start backend: `npm run serve:api`
2. Open: `http://localhost:8787/api/google/oauth/start`
3. Approve access
4. Refresh dashboard and add **Google Calendar Upcoming** from **Edit Cells**
