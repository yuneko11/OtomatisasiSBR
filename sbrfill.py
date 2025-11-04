import asyncio
import argparse
import re
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass
import pandas as pd
from playwright.async_api import async_playwright, Error as PWError, Page, BrowserContext

# ====== KONFIGURASI DEFAULT ======

# Chrome dibuka dengan --remote-debugging-port=9222
CDP_ENDPOINT = "http://localhost:9222"
SHEET_NAME = 0
REQUIRED_COLUMNS_AUTOFILL = ("Status", "Email", "Sumber", "Catatan")
PAUSE_AFTER_EDIT_CLICK_MS = 1000
PAUSE_AFTER_SUBMIT_CLICK_MS = 300
MAX_WAIT_MS = 5000
LOG_CSV = "log_sbr_autofill.csv"
SCREENSHOT_DIR = Path("screenshots")
SCREENSHOT_DIR.mkdir(exist_ok=True)
SLOW_MODE = True
STEP_DELAY_MS = 700
VERBOSE = True
STATUS_ID_MAP = {
    "Aktif": "kondisi_aktif",
    "Tutup Sementara": "kondisi_tutup_sementara",
    "Belum Beroperasi/Berproduksi": "kondisi_belum_beroperasi_berproduksi",
    "Tutup": "kondisi_tutup",
    "Alih Usaha": "kondisi_alih_usaha",
    "Tidak Ditemukan": "kondisi_tidak_ditemukan",
    "Aktif Pindah": "kondisi_aktif_pindah",
    "Aktif Nonrespon": "kondisi_aktif_nonrespon",
    "Duplikat": "kondisi_duplikat",
    "Salah Kode Wilayah": "kondisi_salah_kode_wilayah",
}

@dataclass
class ExcelSelection:
    path: Path
    sheet_index: int = 0


def _format_candidates(paths):
    return ", ".join(str(p) for p in paths)


def resolve_excel(path_arg: str | None, search_dir: Path, sheet_index: int) -> ExcelSelection:
    """
    Jika --excel diberikan -> pakai itu.
    Jika tidak -> cari *.xlsx di search_dir dan search_dir/data (harus 1 file).
    """
    if path_arg:
        p = Path(path_arg).expanduser().resolve()
        if not p.is_file():
            raise FileNotFoundError(f"File Excel tidak ditemukan: {p}")
        return ExcelSelection(path=p, sheet_index=sheet_index)

    locations = [search_dir]
    seen, candidates = set(), []
    for loc in locations:
        if not loc.exists():
            continue
        for c in sorted(loc.glob("*.xlsx")):
            r = c.resolve()
            if r not in seen:
                seen.add(r)
                candidates.append(r)

    if not candidates:
        raise FileNotFoundError(
            "Gunakan argumen --excel untuk memilih file secara eksplisit."
        )
    if len(candidates) > 1:
        raise RuntimeError(
            "Ditemukan lebih dari satu file Excel. Pilih salah satu dengan --excel. Kandidat: "
            f"{_format_candidates(candidates)}"
        )
    return ExcelSelection(path=candidates[0], sheet_index=sheet_index)


def load_dataframe(selection: ExcelSelection, dtype: dict | str | None = str) -> pd.DataFrame:
    return pd.read_excel(selection.path, sheet_name=selection.sheet_index, dtype=dtype)


def ensure_required_columns(df: pd.DataFrame, required=REQUIRED_COLUMNS_AUTOFILL) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise RuntimeError(f"Kolom wajib belum ada di Excel: {', '.join(missing)}")


def slice_rows(df: pd.DataFrame, start: int | None, end: int | None) -> tuple[int, int]:
    start_idx = 0 if start is None else max(start - 1, 0)
    end_idx = len(df) if end is None else min(end, len(df))
    return start_idx, end_idx


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
    if s is None or (isinstance(s, float) and pd.isna(s)) or pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def norm_phone_str(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
        return ""
    return "".join(re.findall(r"\d", str(v)))  # hanya digit


def normfloat_str(s: str) -> str:
    s = normspace(s)
    if not s:
        return ""
    s = s.replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return m.group(0) if m else ""


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


async def is_edit_locked_page(p: Page) -> bool:
    try:
        await p.wait_for_load_state("domcontentloaded", timeout=2500)
    except Exception:
        pass

    checks = [
        # teks utama yang muncul pada UI
        re.compile(r"sedang\s+diedit\s+oleh\s+user\s+lain", re.I),
        re.compile(r"tidak\s+bisa\s+melakukan\s+edit", re.I),
        re.compile(r"Not\s+Authorized", re.I),
        re.compile(r"Profiling\s+Info", re.I),  # judul halaman
        re.compile(r"Back\s+to\s+Home", re.I),
    ]

    # 1) cek title cepat
    try:
        title = (await p.title()) or ""
        if any(r.search(title) for r in checks):
            return True
    except Exception:
        pass

    # 2) cek konten halaman
    try:
        locator = p.get_by_text(re.compile(r"sedang\s+edit|Not\s+Authorized|Back\s+to\s+Home|Profiling\s+Info", re.I))
        if await locator.count() > 0:
            try:
                await locator.first.wait_for(state="visible", timeout=600)
            except Exception:
                pass
            return True
    except Exception:
        pass

    # 3) fallback: lihat URL pattern yang biasa muncul
    try:
        url = p.url or ""
        if re.search(r"not-?authorized", url, re.I):
            return True
    except Exception:
        pass

    return False


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


async def fill_form(
    new_page: Page,
    status_label: str,
    phone_val: str,
    email_val: str,
    lat_val: str,
    lon_val: str,
    sumber: str,
    catatan: str
):

    print("  Mulai mengisi form...")

    # 1. Keberadaan usaha/perusahaan — versi simple & pasti
    if status_label:
        label_clean = normspace(status_label)
        radio_id = STATUS_ID_MAP.get(label_clean)

        try:
            if radio_id:
                radio = new_page.locator(f"#{radio_id}")
                await radio.wait_for(state="attached", timeout=2000)
                try:
                    await radio.check()  # cara paling benar utk input[type=radio]
                except Exception:
                    await radio.click(force=True)  # fallback kecil
                print(f"    Keberadaan usaha diatur ke: {label_clean}")
            else:
                # fallback generik: cari label lalu gunakan atribut 'for'
                lbl = new_page.locator("label").filter(
                    has_text=re.compile(re.escape(label_clean), re.I)
                ).first
                await lbl.wait_for(state="visible", timeout=2000)
                for_id = await lbl.get_attribute("for")
                if for_id:
                    await new_page.locator(f"#{for_id}").check()
                else:
                    await lbl.click(force=True)
                print(f"    Keberadaan usaha diatur (fallback) ke: {label_clean}")
        except Exception as e:
            print(f"    Gagal set status '{label_clean}': {e}")

        await slow_pause(new_page)

     # 2) No. Telp + Email + Latitude + Longitude
    try:
        ident_section = new_page.locator(
            "xpath=//*[self::h4 or self::h5][contains(., 'IDENTITAS USAHA/PERUSAHAAN')]/ancestor::*[contains(@class,'card') or contains(@class,'section')][1]"
        )
        if await ident_section.count() > 0:
            await ident_section.scroll_into_view_if_needed()

        # Nomor Telepon
        phone_clean = norm_phone_str(phone_val)
        tel_input = (
            new_page.get_by_placeholder(re.compile(r"^Nomor\s*Telepon$", re.I))
            .or_(new_page.locator("input#nomor_telepon, input[name='nomor_telepon'], input[name='no_telp'], input[name='telepon']"))
        ).first
        await tel_input.wait_for(state="visible", timeout=1500)

        if phone_clean:
            await tel_input.fill("")
            await tel_input.fill(phone_clean)
            print(f"    Nomor Telepon diisi: {phone_clean}")
        else:
            print("    Nomor Telepon dilewati (Excel kosong/tidak valid).")

        # --- Toggle & input Email (logika: hanya uncheck bila web & Excel kosong) ---
        cb_email = new_page.locator("#check-email").first
        await cb_email.wait_for(state="attached", timeout=500)

        email_input = (
            new_page.locator("input#email, input[name='email'], input[type='email']")
            .or_(new_page.get_by_placeholder(re.compile(r"^email$", re.I)))
        ).first

        # Baca nilai email yang sudah ada di web
        web_state = await new_page.evaluate("""
            () => {
                const inp = document.querySelector('input#email, input[name="email"], input[type="email"]');
                return { value: inp ? (inp.value || '').trim() : '' };
            }
        """)

        web_value = (web_state.get("value") or "").strip()
        excel_value = (email_val or "").strip()

        # 1. Jika Excel punya email → isi ulang (toggle dibiarkan menyala)
        if excel_value:
            try:
                await email_input.wait_for(state="visible", timeout=400)
                await email_input.fill("")
                await email_input.fill(excel_value)
                print(f"    Email diisi: {excel_value}")
            except Exception as e:
                print(f"    Gagal mengisi email: {e}")

        # 2. Jika web sudah berisi email dan Excel kosong → biarkan toggle menyala
        elif web_value:
            print(f"    Email di web sudah ada, toggle dibiarkan aktif: {web_value}")

        # 3. Jika keduanya kosong → matikan toggle dan kosongkan input
        else:
            try:
                await new_page.evaluate("""
                    () => {
                        const cb = document.querySelector('#check-email');
                        const inp = document.querySelector('input#email, input[name="email"], input[type="email"]');
                        if (cb) {
                            cb.checked = false;
                            cb.dispatchEvent(new Event('input', {bubbles:true}));
                            cb.dispatchEvent(new Event('change', {bubbles:true}));
                        }
                        if (inp) {
                            inp.value = '';
                            inp.dispatchEvent(new Event('input', {bubbles:true}));
                            inp.dispatchEvent(new Event('change', {bubbles:true}));
                        }
                    }
                """)
                print("    Toggle email dinonaktifkan (web & Excel kosong).")
            except Exception as e:
                print(f"    Gagal menonaktifkan toggle email: {e}")

        # Latitude & Longitude
        lat_clean = normfloat_str(lat_val)
        lon_clean = normfloat_str(lon_val)

        # Latitude
        if lat_clean:
            try:
                lat_input = (
                    new_page.locator("input#latitude, input[name='latitude']").first
                    .or_(new_page.get_by_placeholder(re.compile(r"^latitude", re.I)))
                )
                await lat_input.wait_for(state="visible", timeout=1500)
                await lat_input.fill("")          # bersihkan dulu
                await lat_input.fill(lat_clean)
                print(f"    Latitude diisi: {lat_clean}")
            except Exception as e:
                print(f"    Gagal isi Latitude: {e}")
        else:
            print("    Latitude dilewati (Excel kosong/tidak valid).")

        # Longitude
        if lon_clean:
            try:
                lon_input = (
                    new_page.locator("input#longitude, input[name='longitude']").first
                    .or_(new_page.get_by_placeholder(re.compile(r"^longitude", re.I)))
                )
                await lon_input.wait_for(state="visible", timeout=1500)
                await lon_input.fill("")          # bersihkan dulu
                await lon_input.fill(lon_clean)
                print(f"    Longitude diisi: {lon_clean}")
            except Exception as e:
                print(f"    Gagal isi Longitude: {e}")
        else:
            print("    Longitude dilewati (Excel kosong/tidak valid).")

    except Exception as e:
        print(f"    Pengisian telp/email/lat/lon bermasalah: {e}")

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

    # Tentukan lokasi pencarian: folder file script
    base_dir = Path(__file__).resolve().parent

    # Pilih file Excel otomatis (atau sesuai --excel) + load dataframe sebagai string
    selection = resolve_excel(args.excel, search_dir=base_dir, sheet_index=args.sheet)
    df = load_dataframe(selection, dtype=str)

    # Validasi kolom wajib
    ensure_required_columns(df, REQUIRED_COLUMNS_AUTOFILL)

    # Validasi kolom untuk match-by
    if args.match_by == "idsbr" and "IDSBR" not in df.columns:
        raise RuntimeError("Match by 'idsbr' dipilih tapi kolom 'IDSBR' tidak ada di Excel")
    if args.match_by == "name" and "Nama" not in df.columns:
        raise RuntimeError("Match by 'name' dipilih tapi kolom 'Nama' tidak ada di Excel")

    # Rentang baris (1-indexed → 0-based)
    start_idx, end_idx = slice_rows(df, args.start, args.end)

    logs = []

    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp(CDP_ENDPOINT)
        context = browser.contexts[0]
        page = await get_active_directory_page(context)

        for i in range(start_idx, end_idx):
            row = df.iloc[i]
            nama_val = normspace (row.get("Nama"))
            status_web = normspace(row.get("Status"))
            phone_val = normspace(row.get("Nomor Telepon"))
            email_val = normspace(row.get("Email"))
            lat_val = normspace(row.get("Latitude"))
            lon_val = normspace(row.get("Longitude"))
            sumber_val = normspace(row.get("Sumber"))
            catatan_val = normspace(row.get("Catatan"))

            print(f"\n=== Baris {i + 1} :: {nama_val} :: Status = {status_web} ===")

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

            # Jika ternyata form sedang diedit profiler lain
            try:
                if await is_edit_locked_page(new_page):
                    shot = await safe_screenshot(new_page, f"edit_locked_baris_{i+1}")
                    log_event(logs, i+1, "WARN", "EDIT_LOCKED",
                            "Form sedang dikunci/diedit oleh user lain. Melewati baris ini.", shot)
                    try:
                        await new_page.close()
                    except Exception:
                        pass

                    await page.bring_to_front()
                    await page.wait_for_timeout(300)
                    continue
            except Exception:
                pass

            # --- Isi form ---
            try:
                await fill_form(
                    new_page,
                    status_web,
                    phone_val,
                    email_val,
                    lat_val,
                    lon_val,
                    sumber_val,
                    catatan_val)
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

                    level = "ERROR" if result != "ERROR_FILL" else "ERROR"
                    log_event(logs, i+1, level, "SUBMIT", result, shot)

                    if result == "ERROR_FILL":
                        print("    ERROR_FILL terdeteksi: tab form dibiarkan terbuka untuk diperiksa.")
                        await new_page.bring_to_front()

                        if args.stop_on_error:
                            print("    --stop-on-error aktif: menghentikan proses.")
                            break
                        else:
                            await page.bring_to_front()
                            await page.wait_for_timeout(300)
                            continue
                    else:
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
    ap.add_argument("--excel", default=None, help="Path ke file Excel (opsional; bila kosong akan dicari otomatis)")
    ap.add_argument("--sheet", type=int, default=SHEET_NAME, help="Index sheet Excel (default 0)")
    ap.add_argument("--start", type=int, default=None, help="Mulai dari baris ke- (1-indexed)")
    ap.add_argument("--end", type=int, default=None, help="Sampai baris ke- (inklusif; default = semua)")
    ap.add_argument("--match-by", choices=["index", "idsbr", "name"], default="index",
                    help="Cara memilih tombol Edit: index (default), idsbr, atau name")
    ap.add_argument("--stop-on-error", action="store_true",
                    help="Berhenti di error pertama (default lanjut ke baris berikutnya).")
    return ap.parse_args()

if __name__ == "__main__":
    import sys, traceback
    try:
        args = parse_args()
        print(f"[INFO] start sbrfill.py  | match_by={args.match_by} | start={args.start} | end={args.end}")
        asyncio.run(run(args))
    except SystemExit:
        raise
    except Exception as e:
        print("\n[ERROR] Uncaught exception:", e)
        traceback.print_exc()
        sys.exit(1)