>alooo, welcome to vibe coding dari yunee. Kindly ask bisa ke @yuli_lssy yaw

# Profiling SBR Autofill CLI

Profiling SBR Autofill CLI adalah script Python berbasis Playwright yang dibuat untuk **mengotomatisasi proses pengisian Profiling SBR (Statistik Badan Usaha)** secara cepat dan konsisten langsung dari data Excel ke web MatchaPro. Program didesain agar fleksibel, mudah digunakan, dan tetap aman untuk dioperasikan bersama browser Chrome yang telah login serta dilengkapi dengan monitoring via log dan screenshot.

---

### 1. Fitur Utama

- **Pengisian Otomatis Form SBR**

  Program secara otomatis membuka form, mengisi data berdasarkan file Excel, dan men-submit hasilnya langsung melalui browser Chrome. Program akan mengisi kolom Status Usaha, Nomor Telepon, Email, Latitude, Longitude, Sumber Profiling, dan Catatan Profiling secara otomatis berdasarkan data Excel.
   >Jika ada tambahan kolom lain di MatchaPro yang ingin diisi lagi, boleh menghubungi kontak di atas.

- **Smart Email Toggle**

  Program mendeteksi isi email baik dari Excel maupun web:
  - Tidak mematikan toggle jika salah satu berisi,
  - Hanya menonaktifkan jika keduanya kosong.

- **Pembatalan Submit Otomatis**

  Program secara otomatis membuka form dan membatalkan submit untuk semua ataupun pada baris-baris tertentu yang ingin dibatalkan submitnya

- **Dukungan Multi-Mode Pencocokan**

  Pilih metode pencarian baris:
  - ``--match-by index`` → berdasarkan urutan baris tabel,
  - ``--match-by idsbr`` → berdasarkan kolom IDSBR,
  - ``--match-by name`` → berdasarkan kolom Nama Usaha.

- **Logging & Screenshot Otomatis**

  Semua hasil proses tersimpan dalam log_sbr_autofill.csv dan setiap error otomatis diambil screenshot-nya.

---

### 2. Struktur Folder

```
.
OtomatisasiSBR/
│
├─ Daftar Profiling SBR.xlsx
├─ README.md
├─ sbrfill.py
├─ sbrcancel.py
├─ screenshots           (otomatis dibuat)
├─ screenshots_cancel    (otomatis dibuat)
├─ log_sbr_autofill.csv  (otomatis dibuat)
└─ log_sbr_cancel.csv    (otomatis dibuat)

```

- **`sbr_fill.py`** → membuka dan mengisi form Profiling SBR sesuai Excel.
- **`sbr_cancel.py`** → membuka dan menekan tombol *Cancel Submit* di form.
- **`Daftar Profiling SBR.xlsx`** → format excel untuk pengisian.
- Semua log dan screenshot otomatis tersimpan

---

### 3. Persiapan/Instalasi

- Memiliki browser Google Chrome
- Mengexport excel yang ingin diotomatisasikan pengisiannya di MATCHAPRO (harus user yang memiliki akses export, ketua tim ipds biasanya) *opsional
- Mengisi excel sesuai format yang ada di repo
- Menginstall Python 3.10+
- Menginstall dependencies berikut di Powershell atau CMD:
  
   ```powershell
   pip install playwright pandas openpyxl
   playwright install chromium
   ```
  
---

### 4. Panduan Penggunaan Program:
1. Clone atau download repo ke folder pc lokal
2. Buka win+R dan jalankan:
   ```
   chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\ChromeProfileSBR"
   ```
3. Chrome akan otomatis terbuka dan silahkan login ke MATCHAPRO *(jangan lupa sudah connect VPN)*
4. Di MATCHAPRO, buka Direktori Usaha dan pilih provinsi serta kab/kot ataupun desa yang ingin diotomatisasi
5. Klik kanan open terminal di folder yang berisi clone repo ini
6. Jalankan program:

   **Program Pengisian Profiling**
     
   ```powershell
   python sbrfill.py --match-by idsbr
   ```
   → program mengisi seluruh baris yang tertampil di browser dan mengisi sesuai dengan data pada excel dengan kode IDSBR yang selaras

   Program juga dapat dijalankan dengan perintah tambahan di antaranya:
   | Perintah                                          | Fungsi                                                                                   |
   | -------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
   | `--excel" C:\path\file.xlsx"`                   | Memilih file Excel tertentu. Jika kosong, skrip mencari satu-satunya `.xlsx` di folder |
   | `--match-by idsbr`                             | Cara mencari tombol**Edit** di tabel: `idsbr`, `name`, atau berdasarkan indeks tabel (`index`) |
   | `--start` / `--end`                            | Menentukan rentang baris yang ingin diisi |
   | `--stop-on-error`                              | Hentikan proses di error pertama. Tanpa perintah ini makan program akan lanjut mengisi ke baris berikutnya walaupun ada pengisian baris yang error|
   | `--no-slow-mode`                               | Mempercepat langkah (hampir tanpa jeda). Cocok jika sudah yakin proses berjalan stabil |

   Contoh menjalankan program dengan perintah tambahan
   ```powershell
   python sbrfill.py --match-by idsbr --start 5 --stop-on-error
   ```

   → program mengisi seluruh baris yang tertampil di browser dimulai langsung dari baris ke 5 dan mengisi sesuai dengan data pada excel dengan kode IDSBR yang selaras serta berhenti saat terjadi error pada pengisian

   **Program Batal Submit**

   ```powershell
   python sbrcancel.py --match-by idsbr
   ```

   → program akan membuka form seluruh baris tertampil dan meng-klik tombol "Cancel Submit"

---

## Pengembang

Dikembangkan oleh:
**Islamiati Yulia M. Lessy – Tim IPDS BPS Kabupaten Buru Selatan**

Program ini dirancang sebagai inisiatif peningkatan efisiensi Profiling SBR di lingkungan BPS.
Distribusi internal diperbolehkan untuk tujuan non-komersial dengan menyertakan kredit pengembang.

---
