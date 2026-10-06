# 📺 Dashboard — 1 File HTML di Apps Script

Tanpa Vercel / Supabase / CORS. Satu project Apps Script menyajikan semuanya.

## File

- `Dashboard.js` → paste ke `Code.gs`
- `Index.html` → tambah file HTML baru di editor
- `Admin.html` → tambah file HTML baru di editor

## Deploy

1. https://script.google.com → New project (misal `Dashboard CP3`)
2. `Code.gs` → hapus isi → paste `Dashboard.js`
3. **+ (plus) → HTML** → nama `Index` → paste `Index.html`
4. **+ → HTML** → nama `Admin` → paste `Admin.html`
5. Deploy → New deployment → Web app → Execute as **Me**, akses **Anyone**
6. URL hasil:
   - Dashboard TV: `.../exec`
   - Input Top 5 + foto: `.../exec?page=admin&key=cp3admin123`
7. Edit berikutnya: Manage deployments → New version (URL tetap)

## Catatan

- Ganti `ADMIN_KEY` di `Dashboard.js` dengan PIN sendiri.
- Foto diupload ke folder Drive `Dashboard5R_Foto` (dibuat otomatis),
  dishare publik-view, URL-nya ditulis ke kolom IMG sheet Leaderboard 5R.
- Nama Top 5 & foreman ditulis balik ke sheet — sheet tetap sumber utama.
- `produksi/` tidak tersentuh dan tetap jalan terpisah.
