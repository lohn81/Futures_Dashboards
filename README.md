# Profit DDE Day Trade Dashboards

Local dashboards for monitoring WINFUT and WDOFUT using live data exported from Profit/Excel DDE.

## What is included

- Combined WINFUT/WDOFUT dashboard: `outputs/winfut_bovespa_dashboard.html`
- WINFUT dashboard: `outputs/winfut_dashboard.html`
- WDOFUT dashboard: `outputs/wdofut_dashboard.html`
- Local Excel bridge/API: `outputs/run_profit_excel_bridge.ps1`
- Quick launcher: `outputs/start_profit_excel_bridge.bat`

## Data source

The bridge reads:

```text
C:\Work Hunting\Quotes.xlsx
```

and serves local endpoints at:

```text
http://127.0.0.1:8765/
http://127.0.0.1:8765/api/quotes
```

## Run

Double-click:

```text
outputs/start_profit_excel_bridge.bat
```

Then open:

```text
http://127.0.0.1:8765/winfut_bovespa_dashboard.html
```

## Notes

This is a decision-support tool for day trade analysis. It does not place orders automatically and its scores are not guarantees of trade outcome.
