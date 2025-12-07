Shopee Refund Sheet Watcher (Excel/Sheet bot)

What it does
- Watches a Google Sheet (like an online Excel) for Shopee order numbers in column A.
- Uses your logged-in Chrome profiles (or headless Chrome) to fetch refund/return details.
- Fills columns B–Q with product, buyer, status, logistics, and store info automatically.
- Skips rows where column B is already filled; clear column B to retry a row.

Tech stack
- Python 3.11+ (tested with CPython 3.13)
- gspread + google-auth (Google Sheets API via service account)
- pyppeteer (connect to Chrome remote debugging)
- Google Sheets API, Shopee Seller Center API (via browser context)
- Google Chrome

Sheet layout (columns)
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
- runner.py — main watcher (Chrome/session handling, Shopee fetch, sheet writes)
- service_account.json — Google service account credentials (keep private, don’t commit)
- test_runner_parse.py — unit test for parsing logic
- config.json (optional) — overrides for runtime settings
- requirements.txt — Python dependencies
- .gitignore — ignores local secrets/config

Quick start (new machine)
1) Install Python 3.11+ and Google Chrome.
2) Clone: `git clone https://github.com/gitwin192-pixel/return_case-.git` and `cd return_case-`.
3) Install deps: `pip install -r requirements.txt`.
4) Add secrets locally: put `service_account.json` in the project folder; share the target Google Sheet with that service account email so it can edit.
5) Optional: create `config.json` to override defaults (see below).
6) Run: `python runner.py`. It fills rows where column B is empty. To retry a row, clear its column B cell (e.g., remove “TIDAK KETEMU”) and it will be reprocessed on the next poll.

Config (config.json)
- sheet_id / sheet_name: which sheet/tab to watch.
- chrome_path: path to Chrome.
- stores: list of stores with code, name, profile_dir, debug_port.
- poll_seconds: delay between polls.
- max_retries / retry_backoff_seconds: retries for Chrome connect and Shopee fetch.
- dry_run: true = log only, skip sheet writes.
- headless: true = launch Chrome headless.
- headless_port_offset: add this to each debug_port to avoid attaching to a visible Chrome session.

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

Running
- `python runner.py`
- Headless: set `"headless": true` (and optionally `headless_port_offset` to avoid attaching to visible Chrome), then rerun.
- Reprocess: clear column B for any row you want retried.
- Logs show Shopee fetch attempts and batch updates; if no empty rows, you’ll see “Tidak ada baris kosong untuk diupdate”.

Customizing
- Change sheet/tab: update `sheet_id`/`sheet_name` in config.
- Switch stores/profiles/ports: edit the `stores` list.
- Headless vs GUI: toggle `headless`; shift ports with `headless_port_offset` to stay separate from normal Chrome.
- Polling speed: adjust `poll_seconds`.
- Safety: set `dry_run` to true to test without writing to the sheet.

Testing
- `python test_runner_parse.py` (basic parse_refund_summary test).

Troubleshooting
- Chrome attaches to your visible session: close that Chrome or use `headless_port_offset` with headless=true.
- error=10002: usually params/auth; re-login to Shopee Seller Center in that profile.
- No sheet updates: ensure column B is empty for rows to process, `dry_run` is false, and watch logs for batch update errors.
