# Dividend Portfolio Tracker

Automated dividend ETF portfolio tracker with a live dashboard.

## Structure

```
dividend-tracker/
├── dashboard/
│   └── index.html          # Interactive portfolio dashboard
├── data/
│   └── portfolio.json      # Portfolio positions and dividend history
├── apps-script/
│   └── Code.gs             # Google Apps Script — runs Wed & Thu automatically
└── README.md
```

## How it works

1. **Google Sheet** holds your portfolio data (positions, cost basis, dividend history)
2. **Apps Script** (`Code.gs`) runs automatically every Wednesday and Thursday morning:
   - Pulls latest YieldMax dividend declarations from GlobeNewsWire
   - Updates dividend totals in the sheet
   - Recalculates all metrics (YTD avg, forecast, payback %, etc.)
   - Emails a weekly summary to nickt8r@gmail.com
3. **Dashboard** (`index.html`) reads from the sheet via a published JSON endpoint and displays live data

## Setup

See `apps-script/SETUP.md` for one-time setup instructions.

## Accounts
- Portfolio: INDIV + IRA (Fidelity)
- Email notifications: nickt8r@gmail.com
- Data source: YieldMax GlobeNewsWire announcements + Market Chameleon
