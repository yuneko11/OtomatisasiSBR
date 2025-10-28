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


def log_event(logs, row_idx: int, level: str, stage: str, note: str, screenshot: str = ""):
    entry = {
        "ts": ts(),
        "row_index": row_idx,
        "level": level,        # "OK" | "WARN" | "ERROR"
        "stage": stage,        # e.g. CLICK_EDIT / OPEN_TAB / FILL / SUBMIT / CONFIRM_SUBMIT
        "note": note,
        "screenshot": screenshot,
    }
    logs.append(entry)
    tag = "!" if level != "OK" else "-"
    print(f"  {tag} [{level}] {stage}: {note}" + (f" (ss: {screenshot})" if screenshot else ""))


async def ensure_click(locator, name: str = "element"):
    await locator.wait_for(state="visible", timeout=MAX_WAIT_MS)
    await locator.click()


async def get_active_directory_page(ctx: BrowserContext) -> Page:
    pages = ctx.pages
    if not pages:
        raise RuntimeError("Tidak ada tab terbuka. Pastikan Chrome sudah membuka halaman Direktori Usaha.")
    return pages[-1]


async def click_edit_by_index(page: Page, index0: int) -> bool:
    table = page.locator("#table_direktori_usaha")
    await table.wait_for(state="visible", timeout=MAX_WAIT_MS)
    rows = table.locator("tbody > tr")
    if index0 >= await rows.count():
        return False
    row = rows.nth(index0)
    btn = row.locator("css=td >> div.d-flex.align-items-center.col-actions >> a.btn-edit-perusahaan").first
    if await btn.count() == 0:
        btn = row.locator(f"xpath=//*[@id='table_direktori_usaha']/tbody/tr[{index0+1}]/td[10]/div/a[1]")

    for _ in range(3):
        try:
            await btn.scroll_into_view_if_needed()
            await btn.click()
            return True
        except Exception:
            await page.evaluate("() => document.querySelectorAll('.tooltip,.modal-backdrop').forEach(e=>e.remove())")
            await page.wait_for_timeout(150)
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


async def try_click(locator, visible_ms=800):
    try:
        if await locator.is_visible(timeout=visible_ms):
            await locator.click()
            return True
    except Exception:
        pass
    return False

async def submit_and_handle(new_page: Page) -> str:
    btn_role = new_page.get_by_role("button", name=re.compile("Submit Final", re.I))
    btn_text = new_page.locator("text=Submit Final").first

    if not (await try_click(btn_role) or await try_click(btn_text)):
        return "NO_SUBMIT_BUTTON"

    await new_page.wait_for_timeout(PAUSE_AFTER_SUBMIT_CLICK_MS)

    # galat pengisian
    try:
        err = new_page.get_by_text(re.compile("Masih terdapat isian yang harus diperbaiki", re.I))
        await err.wait_for(state="visible", timeout=1000)
        ok = new_page.get_by_role("button", name=re.compile("^OK$", re.I))
        if await ok.is_visible():
            await ok.click()
        return "ERROR_FILL"
    except Exception:
        pass

    # cek konsistensi → Ignore
    try:
        kons = new_page.get_by_text(re.compile("Cek Konsistensi", re.I))
        await kons.wait_for(state="visible", timeout=800)
        ign = new_page.get_by_role("button", name=re.compile("^Ignore$", re.I))
        if await ign.is_visible():
            await ign.click(force=True)
            await new_page.wait_for_timeout(250)
    except Exception:
        pass

    # konfirmasi "Ya, Submit!"
    clicked_confirm = False
    for _ in range(10):
        ya = new_page.locator("div.modal.show, div[role='dialog']").locator(
            "button:has-text('Ya, Submit'), a:has-text('Ya, Submit'), button:has-text('Ya, Submit!'), a:has-text('Ya, Submit!')"
        ).first
        if await ya.count() > 0 and await ya.is_visible():
            try:
                await ya.click(force=True)
            except Exception:
                await new_page.evaluate("""
                    () => {
                        const m = document.querySelector('.modal.show,[role="dialog"]');
                        if (!m) return;
                        const c = [...m.querySelectorAll('button,a')].find(el => /ya\\s*,?\\s*submit!?/i.test((el.textContent||'').trim()));
                        if (c) c.click();
                    }
                """)
            clicked_confirm = True
            break
        await new_page.wait_for_timeout(250)

    # sinyal sukses alternatif selain modal
    async def submit_still_visible():
        try:
            if await btn_role.is_visible(timeout=200): return True
        except Exception:
            pass
        try:
            if await btn_text.is_visible(timeout=200): return True
        except Exception:
            pass
        return False

    success_seen = False
    for _ in range(16):
        try:
            sm = new_page.get_by_text(re.compile("Success|Berhasil submit data final", re.I))
            if await sm.is_visible(timeout=120):
                okb = new_page.get_by_role("button", name=re.compile("^OK$", re.I))
                if await okb.is_visible():
                    await okb.click(force=True)
                    await new_page.wait_for_timeout(150)
                success_seen = True
                break
        except Exception:
            pass

        toast = new_page.locator(".toast, .alert-success, .swal2-popup").first
        try:
            if await toast.is_visible(timeout=120):
                success_seen = True
                break
        except Exception:
            pass

        if not await submit_still_visible():
            success_seen = True
            break

        await new_page.wait_for_timeout(200)

    if success_seen:
        return "OK"
    return "NO_SUCCESS_SIGNAL" if clicked_confirm else "NO_CONFIRM"


async def run(args):
    ok_count = 0
    err_count = 0

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
                    clicked = await click_edit_by_index(page, i - start_idx)
                elif args.match_by == "idsbr":
                    clicked = await click_edit_by_text(page, normspace(row.get("IDSBR")))
                elif args.match_by == "name":
                    clicked = await click_edit_by_text(page, normspace(row.get("Nama")))

                if not clicked:
                    shot = await safe_screenshot(page, f"gagal_klik_edit_baris_{i+1}")
                    print("  ! [ERROR] CLICK_EDIT: Tombol Edit tidak ditemukan / tidak bisa diklik")
                    logs.append({"row_index": i+1, "result": "ERROR", "note": "CLICK_EDIT", "screenshot": shot})
                    break  
            except Exception as e:
                shot = await safe_screenshot(page, f"exception_click_edit_baris_{i+1}")
                print(f"  ! [ERROR] CLICK_EDIT_EXCEPTION: {e}")
                logs.append({"row_index": i+1, "result": "ERROR", "note": f"CLICK_EDIT_EXCEPTION: {e}", "screenshot": shot})
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
                shot = await safe_screenshot(page, f"no_new_tab_baris_{i+1}")
                log_event(logs, i+1, "ERROR", "OPEN_TAB", f"Tidak ada tab form: {e}", shot)
                if args.stop_on_error:
                    break

            await new_page.bring_to_front()

            # --- Isi form ---
            try:
                await fill_form(new_page, status_web, sumber_val, catatan_val)
                log_event(logs, i+1, "OK", "FILL", "Form terisi")
            except Exception as e:
                shot = await safe_screenshot(new_page, f"exception_fill_form_baris_{i+1}")
                log_event(logs, i+1, "ERROR", "FILL", f"Exception isi form: {e}", shot)
                try:
                    await new_page.close()
                except:
                    pass
                if args.stop_on_error:
                    break

            # --- Submit & handle ---
            try:
                result = await submit_and_handle(new_page)
                if result != "OK":
                    shot = await safe_screenshot(new_page, f"submit_issue_baris_{i+1}_{result}")
                    log_event(logs, i+1, "ERROR", "SUBMIT", result, shot)
                    try:
                        await new_page.close()
                    except:
                        pass
                    if args.stop_on_error:
                        break
                    else:
                        continue
                else:
                    log_event(logs, i+1, "OK", "SUBMIT", "Submit final sukses")
            except Exception as e:
                shot = await safe_screenshot(new_page, f"exception_submit_baris_{i+1}")
                log_event(logs, i+1, "ERROR", "SUBMIT", f"EXCEPTION:{e}", shot)
                try:
                    await new_page.close()
                except:
                    pass
                if args.stop_on_error:
                    break
                else:
                    continue

            # Tutup tab dan kembali ke direktori
            try:
                await new_page.close()
            except PWError:
                pass
            await page.bring_to_front()
            await page.wait_for_timeout(800)
            log_event(logs, i+1, "OK", "ROW_DONE", "Baris selesai diproses")
            ok_count += 1

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
    ap.add_argument("--stop-on-error", action="store_true",
                help="Berhenti di error pertama (default lanjut ke baris berikutnya).")
    return ap.parse_args()

if __name__ == "__main__":
    args = parse_args()
    asyncio.run(run(args))