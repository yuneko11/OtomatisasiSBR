import asyncio
import argparse
import re
from pathlib import Path
from datetime import datetime
import pandas as pd
from playwright.async_api import async_playwright, Error as PWError, Page, BrowserContext

# ====== KONFIGURASI DEFAULT ======

# Chrome dibuka dengan --remote-debugging-port=9222
CDP_ENDPOINT = "http://localhost:9222"
DEFAULT_EXCEL_PATH = r"C:\kuliah\SBR\Daftar Profiling SBR Kepala Madan.xlsx"
SHEET_NAME = 0
PAUSE_AFTER_EDIT_CLICK_MS = 1000
PAUSE_AFTER_SUBMIT_CLICK_MS = 300
MAX_WAIT_MS = 5000
LOG_CSV = "log_sbr_autofill.csv"
SCREENSHOT_DIR = Path("screenshots")
SCREENSHOT_DIR.mkdir(exist_ok=True)
SLOW_MODE = True
STEP_DELAY_MS = 700
VERBOSE = True


def vlog(msg: str) -> None:
    if VERBOSE:
        print(msg)


async def slow_pause(page: Page, ms: int | None = None):
    """Berhenti sejenak untuk memberi waktu observasi di layar."""
    if SLOW_MODE:
        await page.wait_for_timeout(ms or STEP_DELAY_MS)


def ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def normspace(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


async def safe_screenshot(page: Page, label: str) -> str:
    try:
        safe_label = re.sub(r"[^a-zA-Z0-9*-]+", "-", label)[:50]
        fname = SCREENSHOT_DIR / f"{ts()}*{safe_label}.png"
        await page.screenshot(path=str(fname), full_page=True)
        return str(fname)
    except Exception:
        return ""


async def ensure_click(locator, name: str = "element"):
    await locator.wait_for(state="visible", timeout=MAX_WAIT_MS)
    await locator.click()


async def get_active_directory_page(ctx: BrowserContext) -> Page:
    pages = ctx.pages
    if not pages:
        raise RuntimeError("Tidak ada tab terbuka. Pastikan Chrome sudah membuka halaman Direktori Usaha.")
    return pages[-1]


async def click_edit_by_index(page: Page, index0: int) -> bool:
    """
    Klik tombol Edit berdasarkan urutan (0-based) pada tabel #table_direktori_usaha.
    Selector utama: a.btn-edit-perusahaan di kolom aksi.
    Fallback: X-Path absolut yang kamu berikan.
    """
    # Pastikan tabel ada
    table = page.locator("#table_direktori_usaha")
    await table.wait_for(state="visible", timeout=MAX_WAIT_MS)

    # Ambil baris ke-index0
    rows = table.locator("tbody > tr")
    total = await rows.count()
    if index0 >= total:
        return False

    row = rows.nth(index0)

    # Tombol Edit di kolom aksi (div.col-actions > a.btn-edit-perusahaan)
    btn = row.locator("css=td >> div.d-flex.align-items-center.col-actions >> a.btn-edit-perusahaan").first
    if await btn.count() > 0:
        await btn.scroll_into_view_if_needed()
        await ensure_click(btn, name=f"Edit row {index0 + 1}")
        return True

    # Fallback: X-Path absolut (baris ke-n)
    xpath = f'//*[@id="table_direktori_usaha"]/tbody/tr[{index0 + 1}]/td[10]/div/a[1]'
    btn2 = row.locator(f"xpath={xpath}")
    if await btn2.count() > 0:
        await btn2.scroll_into_view_if_needed()
        await ensure_click(btn2, name=f"Edit row {index0 + 1} (xpath)")
        return True

    return False


async def click_edit_by_text(page: Page, text: str) -> bool:
    """
    Klik tombol Edit pada baris yang mengandung teks 'text' (IDSBR/Nama)
    di salah satu sel <td>. Cocok saat kamu pakai --match-by idsbr / name.
    """
    text = normspace(text)
    if not text:
        return False

    table = page.locator("#table_direktori_usaha")
    await table.wait_for(state="visible", timeout=MAX_WAIT_MS)

    # Cari <tr> yang punya <td> berisi teks tsb (case-insensitive)
    row = table.locator("tbody tr").filter(has_text=re.compile(re.escape(text), re.I)).first

    try:
        await row.wait_for(state="visible", timeout=MAX_WAIT_MS)
    except Exception:
        return False

    # Tombol Edit di kolom aksi
    btn = row.locator("css=td >> div.d-flex.align-items-center.col-actions >> a.btn-edit-perusahaan").first
    if await btn.count() > 0:
        await btn.scroll_into_view_if_needed()
        await ensure_click(btn, name="Edit by text")
        return True

    # Fallback: tombol <a> pertama di kolom aksi
    btn2 = row.locator("xpath=.//td[div[contains(@class,'col-actions')]]//a[1]")
    if await btn2.count() > 0:
        await btn2.scroll_into_view_if_needed()
        await ensure_click(btn2, name="Edit by text (fallback xpath)")
        return True

    return False


async def fill_form(new_page: Page, status_label: str, sumber: str, catatan: str):
    print("  Mulai mengisi form...")

    # 1. Keberadaan usaha/perusahaan
    if status_label:
        try:
            await new_page.get_by_label(re.compile(f"^{re.escape(status_label)}$", re.I)).check()
            print(f"    Keberadaan usaha diatur ke: {status_label}")
        except PWError:
            await ensure_click(
                new_page.locator("label,span").filter(has_text=re.compile(f"^{re.escape(status_label)}$", re.I)).first
            )
            print(f"    Keberadaan usaha diklik manual: {status_label}")
        await slow_pause(new_page)

    # 2) Email: uncheck toggle
    try:
        # pastikan section "IDENTITAS USAHA/PERUSAHAAN" terlihat & terbuka
        ident_section = new_page.locator(
            "xpath=//*[self::h4 or self::h5][contains(., 'IDENTITAS USAHA/PERUSAHAAN')]/ancestor::*[contains(@class,'card') or contains(@class,'section')][1]"
        )
        if await ident_section.count() > 0:
            await ident_section.scroll_into_view_if_needed()
            # kalau ada area collapse, buka
            try:
                collapse = ident_section.locator(".collapse").first
                if await collapse.count() > 0:
                    cls = await collapse.get_attribute("class") or ""
                    if "show" not in cls:
                        # cari tombol pembuka (header/title yang bisa di-klik)
                        toggler = ident_section.locator("[data-bs-toggle='collapse'], [aria-controls]").first
                        if await toggler.count() > 0:
                            await toggler.click(force=True)
                            await new_page.wait_for_timeout(250)
            except Exception:
                pass

        # cari checkbox dengan beberapa cara
        cb = new_page.locator("#check-email, input[type='checkbox']#check-email").first
        if await cb.count() == 0:
            cb = new_page.locator("xpath=//*[@id='check-email']").first

        await cb.wait_for(state="attached", timeout=3000)
        await cb.scroll_into_view_if_needed()

        if await cb.is_checked():
            try:
                await cb.uncheck(force=True)
                print("    Toggle email dinonaktifkan.")
            except Exception:
                # fallback JS kalau styled switch/overlay
                await new_page.evaluate("""
                    () => {
                        const el = document.querySelector('#check-email');
                        if (!el) return;
                        el.checked = false;
                        el.dispatchEvent(new Event('input', {bubbles:true}));
                        el.dispatchEvent(new Event('change', {bubbles:true}));
                    }
                """)
                print("    Toggle email dinonaktifkan (via JS).")
        else:
            print("    Toggle email sudah tidak aktif.")
    except Exception as e:
        # fallback terakhir: klik label for='check-email'
        try:
            await new_page.locator("label[for='check-email']").click(force=True)
            print("    Toggle email dinonaktifkan lewat label.")
        except Exception as e2:
            print(f"    Toggle email tidak ditemukan: {e2}")

    # 3. Isi sumber profiling
    if sumber:
        try:
            await new_page.get_by_placeholder(re.compile("Sumber Profiling", re.I)).fill(sumber)
            print(f"    Sumber Profiling diisi: {sumber}")
        except Exception:
            print("    Field Sumber Profiling tidak ditemukan.")
        await slow_pause(new_page)

    # 4. Isi catatan profiling
    if catatan:
        try:
            await new_page.wait_for_selector("#catatan_profiling", state="visible", timeout=3000)
            await new_page.fill("#catatan_profiling", catatan)
            await new_page.evaluate("""
                () => {
                    const el = document.querySelector('#catatan_profiling');
                    if (el) {
                        el.dispatchEvent(new Event('input', {bubbles:true}));
                        el.dispatchEvent(new Event('change', {bubbles:true}));
                    }
                }
            """)
            print(f"    Catatan diisi: {catatan}")
        except Exception as e:
            print(f"    Gagal mengisi catatan: {e}")
        await slow_pause(new_page)

    print("  Form selesai diisi.")

async def submit_and_handle(new_page: Page) -> str:
    # Klik Submit Final
    try:
        await ensure_click(new_page.get_by_role("button", name=re.compile("Submit Final", re.I)), "Submit Final")
    except PWError:
        await ensure_click(new_page.locator("text=Submit Final").first, "Submit Final (fallback)")

    await new_page.wait_for_timeout(PAUSE_AFTER_SUBMIT_CLICK_MS)

    # (1) Galat pengisian
    try:
        modal_error = new_page.get_by_text(re.compile("Masih terdapat isian yang harus diperbaiki", re.I))
        try:
            await modal_error.wait_for(state="visible", timeout=2000)
            await ensure_click(new_page.get_by_role("button", name=re.compile("^OK$", re.I)), "OK (galat)")
            return "ERROR"
        except Exception:
            pass
    except PWError:
        pass

    # (2) Cek Konsistensi - pilih Ignore bila ada
    try:
        modal_konsistensi = new_page.get_by_text(re.compile("Cek Konsistensi", re.I))
        try:
            await modal_konsistensi.wait_for(state="visible", timeout=1500)
            ignore_btn = new_page.get_by_role("button", name=re.compile("^Ignore$", re.I))
            if await ignore_btn.is_visible():
                await ensure_click(ignore_btn, "Ignore (cek konsistensi)")
                await new_page.wait_for_timeout(800)
        except Exception:
            pass
    except PWError:
        pass

    # (3) Konfirmasi - Ya, Submit!
    clicked_confirm = False
    for _ in range(16):
        btn = (
            new_page.locator("div.modal.show, div[role='dialog']")
            .locator(
                "button:has-text('Ya, Submit'), "
                "a:has-text('Ya, Submit'), "
            )
            .first
        )

        if await btn.count() > 0 and await btn.is_visible():
            await btn.scroll_into_view_if_needed()
            try:
                await btn.click(force=True)
            except Exception:
                # klik via JS jika click biasa gagal
                await new_page.evaluate(
                    """
                    () => {
                        const modal = document.querySelector('.modal.show,[role="dialog"]');
                        if (!modal) return;
                        const cand = [...modal.querySelectorAll('button,a')].find(el =>
                            /ya\\s*,?\\s*submit!?/i.test((el.textContent||'').trim())
                        );
                        if (cand) cand.click();
                    }
                    """
                )
            clicked_confirm = True
            break
        await new_page.wait_for_timeout(100)

    if not clicked_confirm:
        vlog("  > Tidak menemukan tombol 'Ya, Submit!' dengan cepat; lanjutkan alur (anggap tidak perlu konfirmasi).")

    # (4) Success - OK
    try:
        success_modal = new_page.get_by_text(re.compile("Success|Berhasil submit data final", re.I))
        try:
            await success_modal.wait_for(state="visible", timeout=150)
            ok_btn = new_page.get_by_role("button", name=re.compile("^OK$", re.I))
            if await ok_btn.is_visible():
                await ensure_click(ok_btn, "OK (success)")
                await new_page.wait_for_timeout(100)
            return "OK"
        except Exception:
            pass
    except PWError:
        pass

    # Jika tidak ada modal apa pun, anggap sukses
    return "OK"


async def run(args):
    # Baca Excel
    df = pd.read_excel(args.excel, sheet_name=SHEET_NAME)

    # Kolom wajib
    for c in ["Status", "Email", "Sumber", "Catatan"]:
        if c not in df.columns:
            raise RuntimeError(f"Kolom '{c}' tidak ditemukan di Excel")

    # Kolom untuk match by
    if args.match_by == "idsbr" and "IDSBR" not in df.columns:
        raise RuntimeError("Match by 'idsbr' dipilih tapi kolom 'IDSBR' tidak ada di Excel")
    if args.match_by == "name" and "Nama" not in df.columns:
        raise RuntimeError("Match by 'name' dipilih tapi kolom 'Nama' tidak ada di Excel")

    # Tentukan rentang baris (1-indexed input → 0-based index)
    start_idx = 0 if args.start is None else max(args.start - 1, 0)
    end_idx = len(df) if args.end is None else min(args.end, len(df))

    logs = []

    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp(CDP_ENDPOINT)
        context = browser.contexts[0]
        page = await get_active_directory_page(context)

        for i in range(start_idx, end_idx):
            row = df.iloc[i]
            status_web = normspace(row.get("Status"))
            sumber_val = normspace(row.get("Sumber"))
            catatan_val = normspace(row.get("Catatan"))

            print(f"\n=== Baris {i + 1} :: Status='{status_web}' ===")

            # --- Klik Edit ---
            try:
                clicked = False
                if args.match_by == "index":
                    # indeks relatif di halaman
                    clicked = await click_edit_by_index(page, i - start_idx)
                elif args.match_by == "idsbr":
                    clicked = await click_edit_by_text(page, normspace(row.get("IDSBR")))
                elif args.match_by == "name":
                    clicked = await click_edit_by_text(page, normspace(row.get("Nama")))

                if not clicked:
                    shot = await safe_screenshot(page, f"gagal_klik_edit_baris_{i + 1}")
                    logs.append(
                        {"row_index": i + 1, "result": "ERROR", "note": "Gagal klik Edit", "screenshot": shot}
                    )
                    print(f"  !! Gagal klik Edit (lihat screenshot: {shot})")
                    break
            except Exception as e:
                shot = await safe_screenshot(page, f"exception_click_edit_baris_{i + 1}")
                logs.append(
                    {
                        "row_index": i + 1,
                        "result": "ERROR",
                        "note": f"Exception klik Edit: {e}",
                        "screenshot": shot,
                    }
                )
                break

            # --- Popup 'Ya, edit!' ---
            try:
                ya_edit = page.get_by_role("button", name=re.compile(r"Ya,\s*edit!?$", re.I))
                await ensure_click(ya_edit, "Ya, edit!")
            except PWError:
                pass

            await page.wait_for_timeout(PAUSE_AFTER_EDIT_CLICK_MS)

            # --- Ambil tab baru ---
            try:
                new_page = await context.wait_for_event("page", timeout=MAX_WAIT_MS)
            except PWError as e:
                shot = await safe_screenshot(page, f"no_new_tab_baris_{i + 1}")
                logs.append(
                    {
                        "row_index": i + 1,
                        "result": "ERROR",
                        "note": f"Tidak ada tab form: {e}",
                        "screenshot": shot,
                    }
                )
                break

            await new_page.bring_to_front()

            # --- Isi form ---
            try:
                await fill_form(new_page, status_web, sumber_val, catatan_val)
            except Exception as e:
                shot = await safe_screenshot(new_page, f"exception_fill_form_baris_{i + 1}")
                logs.append(
                    {
                        "row_index": i + 1,
                        "result": "ERROR",
                        "note": f"Exception isi form: {e}",
                        "screenshot": shot,
                    }
                )
                await new_page.close()
                break

            # --- Submit & handle ---
            try:
                result = await submit_and_handle(new_page)
                if result == "ERROR":
                    shot = await safe_screenshot(new_page, f"galat_submit_baris_{i + 1}")
                    logs.append(
                        {
                            "row_index": i + 1,
                            "result": "ERROR",
                            "note": "Isian harus diperbaiki",
                            "screenshot": shot,
                        }
                    )
                    await new_page.close()
                    break
            except Exception as e:
                shot = await safe_screenshot(new_page, f"exception_submit_baris_{i + 1}")
                logs.append(
                    {
                        "row_index": i + 1,
                        "result": "ERROR",
                        "note": f"Exception submit: {e}",
                        "screenshot": shot,
                    }
                )
                await new_page.close()
                break

            # Tutup tab dan kembali ke direktori
            try:
                await new_page.close()
            except PWError:
                pass
            await page.bring_to_front()
            await page.wait_for_timeout(800)

            logs.append({"row_index": i + 1, "result": "OK", "note": "", "screenshot": ""})

    # Simpan log
    pd.DataFrame(logs).to_csv(LOG_CSV, index=False)
    print(f"\nSelesai. Log tersimpan di: {LOG_CSV}")


def parse_args():
    ap = argparse.ArgumentParser(description="SBR Autofill (Chrome attach via CDP)")
    ap.add_argument("--excel", default=DEFAULT_EXCEL_PATH, help="Path ke file Excel")
    ap.add_argument("--start", type=int, default=None, help="Mulai dari baris ke- (1-indexed)")
    ap.add_argument("--end", type=int, default=None, help="Sampai baris ke- (inklusif; default = semua)")
    ap.add_argument(
        "--match-by",
        choices=["index", "idsbr", "name"],
        default="index",
        help="Cara memilih tombol Edit: index (default), idsbr, atau name",
    )
    return ap.parse_args()

if __name__ == "__main__":
    args = parse_args()
    asyncio.run(run(args))