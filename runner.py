import asyncio
import json
import logging
from pathlib import Path
import subprocess
from typing import Any, Dict, Optional


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("shopee_bot")

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.json"

DEFAULT_CONFIG = {
    "sheet_id": "1VB-BGNADGpIszIHEGGvrPpnMzONY439Sg7bKNC9IJBU",
    "sheet_name": "prototype",
    "chrome_path": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "stores": [
        {
            "code": "s2c",
            "name": "szcmotor",
            "profile_dir": r"C:\Users\chess\shopee_auto\chrome_profile_seller1",
            "debug_port": 9222,
        },
        {
            "code": "lion",
            "name": "lionkingparts",
            "profile_dir": r"C:\Users\chess\shopee_auto\chrome_profile",
            "debug_port": 9223,
        },
    ],
    "poll_seconds": 3,
    "max_retries": 3,
    "retry_backoff_seconds": 1,
    "dry_run": False,
    "headless": False,
    "headless_port_offset": 0,
}


def load_config() -> dict:
    cfg = DEFAULT_CONFIG.copy()
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                user_cfg = json.load(f)
            cfg.update(user_cfg or {})
            log.info(f"Config loaded from {CONFIG_PATH}")
        except Exception:
            log.warning("Gagal load config.json, pakai default.", exc_info=True)
    return cfg


CONFIG = load_config()

POLL_SECONDS = CONFIG.get("poll_seconds", 3)
MAX_RETRIES = CONFIG.get("max_retries", 3)
RETRY_BACKOFF_SECONDS = CONFIG.get("retry_backoff_seconds", 1)
DRY_RUN = CONFIG.get("dry_run", False)
HEADLESS = CONFIG.get("headless", False)
HEADLESS_PORT_OFFSET = CONFIG.get("headless_port_offset", 0)

import gspread
from google.oauth2.service_account import Credentials
from pyppeteer import connect

# ============================================================
#  GOOGLE SHEET CONFIG
# ============================================================

SHEET_ID = CONFIG["sheet_id"]  # ID dari URL
SHEET_NAME = CONFIG["sheet_name"]  # nama tab (pojok kiri bawah di Google Sheet)

# Kolom:
#   A = no_sn (order_sn Shopee)
#   B = nama_produk
#   C = sku_produk
#   D = nama_pembeli
#   E = qty
#   F = solusi_request
#   G = alasan_request
#   H = status_text
#   I = refund_amount_display
#   J = forward_carrier
#   K = forward_resi
#   L = reverse_carrier
#   M = reverse_resi
#   N = reverse_status
#   O = reverse_hint
#   P = store_code
#   Q = store_name

# ============================================================
#  CHROME / STORE CONFIG
# ============================================================

CHROME_PATH = CONFIG["chrome_path"]

STORES = CONFIG["stores"]

REFUND_PAGE_URL = "https://seller.shopee.co.id/portal/sale/returnrefundcancel"

# ============================================================
#  GOOGLE API INIT
# ============================================================

SERVICE_ACCOUNT_FILE = r"c:\shopee_bot_keys\service_account.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def get_gspread_client():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    client = gspread.authorize(creds)
    return client


def get_worksheet():
    client = get_gspread_client()
    sh = client.open_by_key(SHEET_ID)
    ws = sh.worksheet(SHEET_NAME)
    return ws


# ============================================================
#  CHROME HELPERS (PAKAI SESSION YG SUDAH LOGIN)
# ============================================================

# Cache browser per store biar nggak connect ulang terus
_BROWSER_CACHE = {}


async def ensure_browser(store):
    """
    Pastikan ada Chrome di debug_port store ini.
    - Coba connect ke 127.0.0.1:port
    - Kalau gagal, spawn Chrome dengan user-data-dir yang dikasih
    """
    global _BROWSER_CACHE
    code = store["code"]
    port = store["debug_port"] + HEADLESS_PORT_OFFSET

    # Sudah ada di cache dan masih hidup?
    browser = _BROWSER_CACHE.get(code)
    if browser is not None:
        try:
            _ = await browser.pages()
            return browser
        except Exception:
            # koneksi mati, buang dari cache
            _BROWSER_CACHE.pop(code, None)

    url = f"http://127.0.0.1:{port}"

    # Coba connect dulu
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            browser = await connect(browserURL=url)
            _BROWSER_CACHE[code] = browser
            if attempt > 1:
                log.info(f"[{code}] Tersambung ke Chrome setelah retry {attempt}")
            return browser
        except Exception:
            if attempt == 1:
                log.info(f"[{code}] Belum ada Chrome di port {port}, coba nyalain...")
            await asyncio.sleep(RETRY_BACKOFF_SECONDS * attempt)

    # Start Chrome baru
    profile_dir = store["profile_dir"]
    cmd = [
        CHROME_PATH,
        f"--user-data-dir={profile_dir}",
        f"--remote-debugging-port={port}",
        "--no-first-run",
        "--no-default-browser-check",
    ]
    if HEADLESS:
        cmd += ["--headless=new", "--disable-gpu", "--window-size=1366,768"]
        log.info(f"[{code}] Start Chrome headless di port {port}")

    try:
        subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        log.error(f"[{code}] GAGAL start Chrome, cek CHROME_PATH / profile_dir.", exc_info=True)
        return None

    # Tunggu Chrome boot
    await asyncio.sleep(5)

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            browser = await connect(browserURL=url)
            _BROWSER_CACHE[code] = browser
            return browser
        except Exception:
            log.warning(f"[{code}] Gagal connect ke Chrome attempt {attempt} di {url}", exc_info=True)
            await asyncio.sleep(RETRY_BACKOFF_SECONDS * attempt)

    log.error(f"[{code}] Tetep nggak bisa connect ke Chrome di {url}")
    return None


async def get_refund_page(browser, store_code: str):
    """
    Ambil tab yang sudah di halaman Pengembalian/Pembatalan.
    Kalau belum ada, buka tab baru dan goto URL.
    """
    try:
        pages = await browser.pages()
    except Exception:
        log.error(f"[{store_code}] ERROR baca pages()", exc_info=True)
        return None

    page = None
    for p in pages:
        try:
            u = p.url or ""
        except Exception:
            continue
        if "portal/sale/returnrefundcancel" in u:
            page = p
            break

    if page is None:
        try:
            page = await browser.newPage()
            await page.goto(REFUND_PAGE_URL, {"waitUntil": "networkidle2"})
        except Exception:
            log.error(f"[{store_code}] ERROR buka halaman refund", exc_info=True)
            return None

    return page


# ============================================================
#  SHOPEE REFUND API VIA fetch DI DALAM TAB
# ============================================================

JS_FETCH_REFUND = """
async (orderSn) => {
  const spcMatch = document.cookie.match(/SPC_CDS=([^;]+)/);
  const spcCds = spcMatch ? spcMatch[1] : "";
  const url = spcCds
    ? `https://seller.shopee.co.id/api/v4/seller_center/return/return_list/get_exceptional_case_list?SPC_CDS=${spcCds}&SPC_CDS_VER=2`
    : "https://seller.shopee.co.id/api/v4/seller_center/return/return_list/get_exceptional_case_list";
  const payload = {
    language: "id",
    is_reverse_sorting_order: false,
    page_number: 1,
    page_size: 40,
    keyword: orderSn,
    pending_action: null,
    request_solution: null,
    forward_logistics_statuses: [],
    reverse_logistics_statuses: [],
    return_reasons: [],
    create_time_range: { lower_value: null, upper_value: null },
    compensation_amount_option: null,
    advanced_fulfilment_option: null,
    seller_request_statuses: [],
    validation_type_option: null,
    request_adjusted: null,
    refund_amount_range: { lower_value: null, upper_value: null },
    flow_tab: 1,
    case_tab: 0,
    sorting_field: 1,
    key_action_due_time_range: { lower_value: null, upper_value: null },
    platform_type: "sc",
  };

  const res = await fetch(url, {
    method: "POST",
    headers: {
      "content-type": "application/json",
    },
    credentials: "include",
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    return { _http_error: res.status };
  }

  const data = await res.json();
  return data;
}
"""


def parse_refund_summary(
    raw: Dict[str, Any], store_code: str, store_name: str
) -> Optional[Dict[str, Any]]:
    """
    Flatten the Shopee refund payload into a simpler dict we can write to the sheet.
    """
    try:
        if raw.get("error") != 0:
            return None

        cases = raw.get("data", {}).get("exceptional_case_list", [])
        if not cases:
            return None

        data = cases[0]

        order_sn = data.get("order_sn", "")
        buyer_name = data.get("buyer", {}).get("name", "")

        item_strings = []
        sku_list = []
        total_qty = 0
        for item in data.get("product_items", []):
            p = item.get("product", {})
            m = item.get("model", {})

            name = p.get("name", "")
            sku = p.get("sku", "")
            model_name = m.get("name", "")
            qty = item.get("amount", 0) or 0

            if isinstance(qty, (int, float)):
                total_qty += qty

            if sku:
                sku_list.append(sku)

            parts = [name]
            if model_name:
                parts.append(f"({model_name})")
            if sku:
                parts.append(f"[{sku}]")
            if qty:
                parts.append(f"x{qty}")

            item_strings.append(" ".join(parts))

        product_name_joined = " | ".join(item_strings) if item_strings else ""
        product_sku_joined = " | ".join(sku_list) if sku_list else ""

        request_solution = data.get("request_solution_text", "")
        request_reason = data.get("request_reason_text", "")
        status_text = data.get("header", {}).get("status_text", "")
        refund_amount_display = data.get("display_refund_amount", "")
        region = data.get("region", "")
        payment_method = data.get("payment_method", "")

        fwd = data.get("forward_logistics_info", {})
        rev = data.get("reverse_logistics_info", {})

        forward_carrier = fwd.get("shipping_carrier", "")
        forward_resi_list = fwd.get("tracking_numbers", []) or []
        forward_resi = forward_resi_list[0] if forward_resi_list else ""

        reverse_carrier = rev.get("shipping_carrier", "")
        reverse_resi_list = rev.get("tracking_numbers", []) or []
        reverse_resi = reverse_resi_list[0] if reverse_resi_list else ""
        reverse_status = rev.get("aggregated_logistics_status_text", "")
        reverse_hint = rev.get("hint_text", "")

        return {
            "order_sn": order_sn,
            "buyer_name": buyer_name,
            "product_name": product_name_joined,
            "product_sku": product_sku_joined,
            "qty": total_qty,
            "request_solution": request_solution,
            "request_reason": request_reason,
            "status_text": status_text,
            "refund_amount_display": refund_amount_display,
            "forward_carrier": forward_carrier,
            "forward_resi": forward_resi,
            "reverse_carrier": reverse_carrier,
            "reverse_resi": reverse_resi,
            "reverse_status": reverse_status,
            "reverse_hint": reverse_hint,
            "region": region,
            "payment_method": payment_method,
            "store_code": store_code,
            "store_name": store_name,
        }
    except Exception:
        traceback.print_exc()
        return None


async def fetch_refund_raw(order_sn: str, store) -> dict | None:
    """
    Panggil fetch() di dalam tab Chrome yang sudah login.
    """
    code = store["code"]
    browser = await ensure_browser(store)
    if browser is None:
        log.error(f"[{code}] Tidak ada browser.")
        return None

    page = await get_refund_page(browser, code)
    if page is None:
        return None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            raw = await page.evaluate(JS_FETCH_REFUND, order_sn)
        except Exception:
            log.warning(f"[{code}] ERROR waktu evaluate JS fetch_refund attempt {attempt}", exc_info=True)
            await asyncio.sleep(RETRY_BACKOFF_SECONDS * attempt)
            continue

        if not isinstance(raw, dict):
            log.warning(f"[{code}] Response JS bukan dict: {type(raw)}")
            return None

        if "_http_error" in raw:
            log.warning(f"[{code}] HTTP error {raw['_http_error']} buat {order_sn}")
            await asyncio.sleep(RETRY_BACKOFF_SECONDS * attempt)
            continue

        return raw

    log.error(f"[{code}] Gagal fetch {order_sn} setelah {MAX_RETRIES} attempt")
    return None


async def fetch_refund_summary(order_sn: str) -> dict | None:
    """
    Coba 1 order_sn ke SEMUA store.
    Begitu ada yang ketemu, langsung return summary.
    Kalau semua None -> return None.
    """
    for store in STORES:
        code = store["code"]
        log.info(f"[{code}] Coba cari {order_sn}...")

        raw = await fetch_refund_raw(order_sn, store)
        if raw is None:
            continue
        if raw.get("error") == 10002:
            log.warning(f"[{code}] error=10002 (params/auth?) untuk {order_sn}. Cek login Chrome.")
            continue

        cases = raw.get("data", {}).get("exceptional_case_list", [])
        if not cases:
            log.info(f"[{code}] No cases returned for {order_sn} (error={raw.get('error')})")
            continue

        summary = parse_refund_summary(raw, code, store.get("name", ""))
        if summary is not None:
            return summary

    log.info(f"  {order_sn} tidak ketemu di semua toko.")
    return None


# ============================================================
#  LOOP WATCHER GOOGLE SHEET (ASYNC)
# ============================================================

async def main_loop():
    log.info("Watcher Google Sheet jalan. Ctrl+C buat stop.")
    ws = get_worksheet()
    log.info(f"Sheet: {SHEET_ID} / tab: {SHEET_NAME}")

    while True:
        try:
            values = ws.get_all_values()
            if len(values) <= 1:
                await asyncio.sleep(POLL_SECONDS)
                continue

            seen_filled = 0
            updates = []
            for idx, row in enumerate(values[1:], start=2):
                order_sn = (row[0] if len(row) >= 1 else "").strip()
                if not order_sn:
                    continue

                current_name = (row[1] if len(row) >= 2 else "").strip()
                if current_name:
                    seen_filled += 1
                    continue

                log.info(f"[ROW {idx}] Cari {order_sn}")
                summary = await fetch_refund_summary(order_sn)

                if summary:
                    row_values = [
                        summary["product_name"],          # B
                        summary["product_sku"],           # C
                        summary["buyer_name"],            # D
                        summary["qty"],                   # E
                        summary["request_solution"],      # F
                        summary["request_reason"],        # G
                        summary["status_text"],           # H
                        summary["refund_amount_display"], # I
                        summary["forward_carrier"],       # J
                        summary["forward_resi"],          # K
                        summary["reverse_carrier"],       # L
                        summary["reverse_resi"],          # M
                        summary["reverse_status"],        # N
                        summary["reverse_hint"],          # O
                        summary["store_code"],            # P
                        summary["store_name"],            # Q
                    ]
                    log.info(f"  -> OK: {summary['product_name']} ({summary['store_code']})")
                else:
                    row_values = ["TIDAK KETEMU"] + [""] * 15
                    log.info("  -> TIDAK KETEMU")

                updates.append({"range": f"'{SHEET_NAME}'!B{idx}:Q{idx}", "values": [row_values]})

            if updates:
                log.info(f"Siap update {len(updates)} baris ke sheet...")
                if DRY_RUN:
                    log.info(f"[DRY RUN] {len(updates)} update siap, skip tulis ke sheet.")
                else:
                    # batch to avoid oversized requests
                    try:
                        for i in range(0, len(updates), 30):
                            chunk = updates[i : i + 30]
                            ws.spreadsheet.values_batch_update(
                                {"valueInputOption": "RAW", "data": chunk}
                            )
                        log.info(f"Batch update selesai untuk {len(updates)} baris")
                    except Exception:
                        log.error("Gagal batch update ke Google Sheet", exc_info=True)
            else:
                log.info(f"Tidak ada baris kosong untuk diupdate (terisi: {seen_filled})")

            await asyncio.sleep(POLL_SECONDS)

        except KeyboardInterrupt:
            log.info("Stop oleh user (Ctrl+C).")
            break
        except Exception:
            log.error("ERROR di main loop:", exc_info=True)
            await asyncio.sleep(5)


if __name__ == "__main__":
    try:
        asyncio.run(main_loop())
    except KeyboardInterrupt:
        log.info("Stop oleh user (Ctrl+C) [top-level].")
    except asyncio.CancelledError:
        log.info("Task dibatalkan.")
