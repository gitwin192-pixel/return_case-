Shopee Refund Sheet Watcher

- Watches a Google Sheet for Shopee order numbers (column A), fetches refund/return details via logged-in Chrome sessions, and writes results into columns B–Q.
- Uses Chrome remote debugging to piggyback on existing logged-in profiles; can also launch headless Chrome.
- Logs to stdout with structured messages; optional dry-run mode to skip writes.

Tech Stack
- Python 3.11+ (tested with CPython 3.13)
- gspread + google-auth (Google Sheets API via service account)
- pyppeteer (connect to Chrome remote debugging)
- Google Sheets API, Shopee Seller Center API (through browser context)
- Google Chrome (installed locally)

Sheet Layout (Columns)
- A: order_sn (input)
- B: nama_produk
- C: sku_produk
- D: nama_pembeli
- E: qty
- F: solusi_request
- G: alasan_request
- H: status_text
- I: refund_amount_display
- J: forward_carrier
- K: forward_resi
- L: reverse_carrier
- M: reverse_resi
- N: reverse_status
- O: reverse_hint
- P: store_code
- Q: store_name

Files
- runner.py — main watcher logic, Chrome/session handling, Shopee fetch, sheet writes.
- service_account.json — Google service account credentials (keep private).
- test_runner_parse.py — unit test for parsing logic.
- config.json (optional) — overrides for runtime settings.

Config (config.json)
- sheet_id: Google Sheet ID (default matches current sheet).
- sheet_name: tab name (default: prototype).
- chrome_path: path to Chrome executable.
- stores: list of stores with code, name, profile_dir, debug_port.
- poll_seconds: delay between sheet polls.
- max_retries / retry_backoff_seconds: retries for Chrome/fetch.
- dry_run: if true, skip sheet writes.
- headless: if true, launch Chrome headless.
- headless_port_offset: integer added to each store’s debug_port (use to avoid attaching to visible Chrome).

Example config.json
```json
{
  "sheet_id": "1VB-BGNADGpIszIHEGGvrPpnMzONY439Sg7bKNC9IJBU",
  "sheet_name": "prototype",
  "poll_seconds": 3,
  "dry_run": false,
  "headless": true,
  "headless_port_offset": 10000,
  "stores": [
    {
      "code": "s2c",
      "name": "szcmotor",
      "profile_dir": "C:\\\\Users\\\\chess\\\\shopee_auto\\\\chrome_profile_seller1",
      "debug_port": 9222
    },
    {
      "code": "lion",
      "name": "lionkingparts",
      "profile_dir": "C:\\\\Users\\\\chess\\\\shopee_auto\\\\chrome_profile",
      "debug_port": 9223
    }
  ]
}
```

Setup
- Install deps: `pip install gspread google-auth pyppeteer`.
- Ensure Chrome is installed and `chrome_path` points to the executable.
- Place `service_account.json` in `C:\shopee_bot_keys` (or adjust `SERVICE_ACCOUNT_FILE` in runner.py if you move it).
- Make sure the service account has access to the target Google Sheet.

Requirements
- Install all deps with `pip install -r C:\shopee_bot_keys\requirements.txt`.

Running
- From `C:\shopee_bot_keys`: `python runner.py`.
- Headless: set `"headless": true` (and optionally `headless_port_offset` to avoid attaching to visible Chrome), then rerun.
- The watcher processes rows where column B is empty. To reprocess a row, clear its column B cell (e.g., remove “TIDAK KETEMU”) and it will be retried on the next poll.
- Logs show Shopee fetch attempts and sheet batch updates. If no empty rows, you’ll see “Tidak ada baris kosong untuk diupdate”.

Testing
- `python test_runner_parse.py` (basic parse_refund_summary test).

Troubleshooting
- Chrome not starting / attaches to visible Chrome: close existing Chrome using the same profile/port, or use `headless_port_offset` to shift ports for headless runs.
- error=10002: usually params/auth; ensure Chrome is logged in to Seller Center.
- No sheet updates: ensure column B is empty for rows to process, dry_run is false, and watch for batch update errors in logs.
