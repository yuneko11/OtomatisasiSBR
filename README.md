alooo, welcome to vibe coding dari yunee. Kindly ask bisa ke @yuli_lssy yaw

Persiapan:
1. Memiliki browser Google Chrome
2. Akun Profiling SBR di MATCHAPRO
3. Mengexport excel yang ingin diotomatisasikan pengisiannya di MATCHAPRO (harus user yang memiliki akses export, ketua tim ipds biasanya)
4. Mengisi excel sesuai format yang ada di repo
   Note: Di satkerku yang kami isi cuma kolom requirement aja (status, uncheck email, sumber dan catatan profiling),
   sebenarnya programnya fleksible aja tinggal dikoding lagi kalau ada tambahan kolom yang mau diisi

Tata cara penggunaan:
1. Clone atau download repo ke folder pc lokal
2. Buka pythonnya dan sesuaikan dulu codenya, seperti direktori excel path atau mungkin ada tambahan kolom yang diingikan (bisa edit di function fillform)
3. Buka win+R dan jalankan: chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\ChromeProfileSBR" (bisa juga jalanin di cmd pake port yang sama)
4. Chrome akan otomatis terbuka dan silahkan login ke MATCHAPRO (jangan lupa sudah connect VPN)
5. Di MATCHAPRO, buka Direktori Usaha dan pilih provinsi serta kab/kot yang ingin diotomatisasi
6. Klik kanan open terminal di folder yang berisi clone repo ini
7. Jalankan: python sbrfill.py --match-by id sbr
   (Jika ingin melanjutkan pengisian dari baris 5 misalnya bisa jalankan: python sbrfill.py --match-by id sbr --start 5)
