import asyncio
import argparse
import re
from pathlib import Path
from datetime import datetime
import pandas as pd
from playwright.async_api import async_playwright, Error as PWError, Page, BrowserContext

# ====== KONFIGURASI DEFAULT ======
CDP_ENDPOINT = "http://localhost:9222"  # Jalankan Chrome dengan: chrome.exe --remote-debugging-port=9222
DEFAULT_EXCEL_PATH = r"C:\kuliah\SBR\Daftar Profiling SBR Kepala Madan.xlsx"
SHEET_NAME = 0

VERBOSE = True
def vlog(msg):
    if VERBOSE:
        print(msg)

PAUSE_AFTER_EDIT_CLICK_MS = 800
MAX_WAIT_MS = 8000
SLOW_MODE = True
STEP_DELAY_MS = 500
async def step_pause(page: Page, ms: int | None = None):
    if SLOW_MODE:
        await page.wait_for_timeout(ms or STEP_DELAY_MS)

LOG_CSV = "log_sbr_cancel.csv"
SCREENSHOT_DIR = Path("screenshots_cancel")
SCREENSHOT_DIR.mkdir(exist_ok=True)

def ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def normspace(s) -> str:
    import re as _re
    return _re.sub(r"\s+", " ", str(s or "")).strip()

async def safe_screenshot(page: Page, label: str):
    import re as _re
    fname = SCREENSHOT_DIR / f"{ts()}_{_re.sub(r'[^a-zA-Z0-9_-]+','-',label)[:60]}.png"
    try:
        await page.screenshot(path=str(fname), full_page=True)
        return str(fname)
    except Exception:
        return ""

async def ensure_click(locator, name="element"):
    await locator.wait_for(state="visible", timeout=MAX_WAIT_MS)
    await locator.scroll_into_view_if_needed()
    await locator.click()

async def get_active_directory_page(ctx: BrowserContext) -> Page:
    pages = ctx.pages
    if not pages:
        raise RuntimeError("Tidak ada tab terbuka. Pastikan Chrome sudah pada halaman Direktori Usaha.")
    return pages[-1]

# ---------- Klik tombol Edit di tabel ----------
async def click_edit_by_index(page, index0: int) -> bool:
    table = page.locator("#table_direktori_usaha")
    await table.wait_for(state="visible", timeout=MAX_WAIT_MS)

    rows = table.locator("tbody > tr")
    total = await rows.count()
    if index0 >= total:
        return False

    row = rows.nth(index0)

    # tombol oranye Edit (kolom aksi)
    btn = row.locator("css=td >> div.d-flex.align-items-center.col-actions >> a.btn-edit-perusahaan").first
    if await btn.count() > 0:
        await ensure_click(btn, name=f"Edit row {index0+1}")
        return True

    # fallback xpath absolut
    xpath = f'//*[@id="table_direktori_usaha"]/tbody/tr[{index0+1}]/td[10]/div/a[1]'
    btn2 = row.locator(f"xpath={xpath}")
    if await btn2.count() > 0:
        await ensure_click(btn2, name=f"Edit row {index0+1} (xpath)")
        return True

    return False

async def click_edit_by_text(page, text: str) -> bool:
    text = normspace(text)
    if not text:
        return False

    table = page.locator("#table_direktori_usaha")
    await table.wait_for(state="visible", timeout=MAX_WAIT_MS)

    row = table.locator("tbody tr").filter(has_text=re.compile(re.escape(text), re.I)).first
    try:
        await row.wait_for(state="visible", timeout=MAX_WAIT_MS)
    except:
        return False

    btn = row.locator("css=td >> div.d-flex.align-items-center.col-actions >> a.btn-edit-perusahaan").first
    if await btn.count() > 0:
        await ensure_click(btn, name="Edit by text")
        return True

    btn2 = row.locator("xpath=.//td[div[contains(@class,'col-actions')]]//a[1]")
    if await btn2.count() > 0:
        await ensure_click(btn2, name="Edit by text (fallback)")
        return True

    return False

# ---------- Alur Cancel Submit di tab form ----------
async def do_cancel_submit(new_page: Page) -> str:
    print("  Membuka tab form..."); await step_pause(new_page, 300)

    # 1) Klik tombol "Cancel Submit"
    try:
        # pakai xpath yang kamu berikan
        btn = new_page.locator("xpath=//*[@id='cancel-submit-final']/span")
        if await btn.count() == 0:
            # fallback berdasarkan teks
            btn = new_page.locator("button:has-text('Cancel Submit'), a:has-text('Cancel Submit')").first
        await ensure_click(btn, "Cancel Submit")
        print("    Klik: Cancel Submit")
    except Exception as e:
        print(f"    Gagal klik Cancel Submit: {e}")
        return "ERROR"
    await step_pause(new_page)

    # 2) Dialog konfirmasi → "Ya, batalkan!"
    try:
        modal = new_page.locator("div.modal.show, div[role='dialog']").filter(has_text=re.compile("Konfirmasi|Konfirmasi", re.I)).first
        await modal.wait_for(timeout=4000)
        ya_btn = modal.locator("button:has-text('Ya, batalkan!'), a:has-text('Ya, batalkan!')").first
        await ya_btn.click(force=True)
        print("    Konfirmasi: Ya, batalkan!")
    except Exception as e:
        print(f"    Gagal klik 'Ya, batalkan!': {e}")
        return "ERROR"
    await step_pause(new_page)

    # 3) Dialog Success → "OK"
    try:
        for _ in range(20):  # ~5 detik
            ok_btn = new_page.locator("button:has-text('OK')").first
            if await ok_btn.count() > 0 and await ok_btn.is_visible():
                await ok_btn.click(force=True)
                print("    Success: OK ditekan")
                await step_pause(new_page, 300)
                return "OK"
            await new_page.wait_for_timeout(250)
        print("    Tidak menemukan dialog Success; diasumsikan OK")
        return "OK"
    except Exception as e:
        print(f"    Gagal menutup dialog success: {e}")
        return "ERROR"

# ---------- Main runner ----------
async def run(args):
    # Baca Excel (dipakai untuk iterasi & match_by)
    df = pd.read_excel(args.excel, sheet_name=SHEET_NAME)

    # Validasi kolom untuk match_by
    if args.match_by == "idsbr" and "IDSBR" not in df.columns:
        raise RuntimeError("Match by 'idsbr' dipilih tapi kolom 'IDSBR' tidak ada di Excel")
    if args.match_by == "name" and "Nama" not in df.columns:
        raise RuntimeError("Match by 'name' dipilih tapi kolom 'Nama' tidak ada di Excel")

    start_idx = 0 if args.start is None else max(args.start - 1, 0)
    end_idx = len(df) if args.end is None else min(args.end, len(df))

    logs = []

    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp(CDP_ENDPOINT)
        context = browser.contexts[0]
        page = await get_active_directory_page(context)

        for i in range(start_idx, end_idx):
            row = df.iloc[i]
            print(f"\n=== Baris {i+1} ===")

            # 0) Klik Edit di tabel
            try:
                clicked = False
                if args.match_by == "index":
                    clicked = await click_edit_by_index(page, i - start_idx)
                elif args.match_by == "idsbr":
                    clicked = await click_edit_by_text(page, normspace(row.get("IDSBR")))
                elif args.match_by == "name":
                    clicked = await click_edit_by_text(page, normspace(row.get("Nama")))

                if not clicked:
                    shot = await safe_screenshot(page, f"gagal_klik_edit_baris_{i+1}")
                    print(f"  Tidak bisa klik Edit (lihat {shot})")
                    logs.append({"row_index": i+1, "result": "ERROR", "note": "Gagal klik Edit", "screenshot": shot})
                    break
                print("  Klik Edit berhasil")
            except Exception as e:
                shot = await safe_screenshot(page, f"exception_click_edit_baris_{i+1}")
                logs.append({"row_index": i+1, "result": "ERROR", "note": f"Exception klik Edit: {e}", "screenshot": shot})
                break

            # 0a) Popup "Ya, edit!"
            try:
                ya_edit = page.get_by_role("button", name=re.compile(r"Ya,\s*edit!?$", re.I))
                if await ya_edit.count() > 0:
                    await ensure_click(ya_edit, "Ya, edit!")
                    print("  Konfirmasi awal: Ya, edit!")
            except PWError:
                pass
            await page.wait_for_timeout(PAUSE_AFTER_EDIT_CLICK_MS)

            # 1) Ambil tab baru (form)
            try:
                new_page = await context.wait_for_event("page", timeout=MAX_WAIT_MS)
            except PWError as e:
                shot = await safe_screenshot(page, f"no_new_tab_baris_{i+1}")
                logs.append({"row_index": i+1, "result": "ERROR", "note": f"Tidak ada tab form: {e}", "screenshot": shot})
                break

            await new_page.bring_to_front()

            # 2) Jalankan alur Cancel Submit
            result = await do_cancel_submit(new_page)

            # 3) Tutup tab form & kembali
            try:
                await new_page.close()
            except PWError:
                pass
            await page.bring_to_front()
            print("  Tab form ditutup, kembali ke Direktori.")

            logs.append({"row_index": i+1, "result": result, "note": "", "screenshot": ""})
            if result != "OK":
                break

    # Simpan log
    pd.DataFrame(logs).to_csv(LOG_CSV, index=False)
    print(f"\nSelesai. Log tersimpan di: {LOG_CSV}")

def parse_args():
    ap = argparse.ArgumentParser(description="SBR Cancel Submit (attach via CDP)")
    ap.add_argument("--excel", default=DEFAULT_EXCEL_PATH, help="Path ke file Excel")
    ap.add_argument("--start", type=int, default=None, help="Mulai dari baris ke- (1-indexed)")
    ap.add_argument("--end", type=int, default=None, help="Sampai baris ke- (inklusif; default = semua)")
    ap.add_argument("--match-by", choices=["index", "idsbr", "name"], default="index",
                   help="Cara memilih tombol Edit: index (default), idsbr, atau name")
    return ap.parse_args()

if __name__ == "__main__":
    args = parse_args()
    asyncio.run(run(args))