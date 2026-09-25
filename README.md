# Dashboard ITT CGPP

- `/` — dashboard publik (`public/index.html`)
- `/admin` — upload Excel ITT untuk meng-update dashboard (`public/admin.html`)

Alur: admin upload `.xlsx` → `api/upload.js` membaca Excel (`lib/parse.js`), mencocokkan hasil klasifikasi dengan sheet **Summary**, lalu menyimpan JSON ke Vercel Blob → `api/data.js` menyajikan JSON itu ke dashboard. Kalau belum pernah ada upload, dashboard memakai `public/data.json`.

## Deploy ke Vercel (sekali saja)

1. Push repo ini ke GitHub, lalu di Vercel: **Add New → Project → Import** repo tersebut (Framework: *Other*, tanpa build command).
2. Di project Vercel: **Storage → Create → Blob** → hubungkan ke project (otomatis menambah env `BLOB_READ_WRITE_TOKEN`).
3. **Settings → Environment Variables**: tambahkan `ADMIN_PASSWORD` = password admin yang kuat.
4. **Redeploy**. Buka `https://<project>.vercel.app/admin`.

## Cek file Excel secara lokal

```bash
npm install
npm run check -- "path/ke/file.xlsx"
```

Memvalidasi file dan menulis ulang `public/data.json` (data bawaan).

## Aturan file Excel

Struktur harus sama dengan template ITT: sheet `ITT_CGPP Indonesia`, 3 sheet `ITT_Provinsi …`, 7 sheet kabupaten/kota, dan `Summary` (D3 = bulan berjalan). Posisi kolom dibaca dari header, jadi kolom boleh bergeser, tapi urutan baris indikator harus sama di semua sheet. Simpan file di Excel (bukan hasil ekspor tanpa nilai formula) supaya nilai hasil formula ikut tersimpan.
