# 🔌 Backend — Web App JSON untuk Display

Pengganti Supabase: display Astro fetch langsung ke Apps Script Web App ini.

## Deploy (sekali saja)

1. https://script.google.com → **New project** → nama misal `Backend JSON`
2. Paste seluruh `Backend.js` → Save
3. **Deploy → New deployment → Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Copy URL `.../exec` → pasang sebagai `PUBLIC_BACKEND_URL` di display (file `.env` lokal + env Vercel)
5. Kalau edit kode: **Deploy → Manage deployments → pensil → New version** (URL tidak berubah)

## Endpoint

- `GET ?action=all` → `{ dashboard, leaderboard }`
- `GET ?action=dashboard` / `?action=leaderboard`
- `&refresh=1` → bypass cache 10 menit (debug)

## Catatan

- Akun pendeploy harus punya akses (min. Viewer) ke spreadsheet sumber.
- Kalau sheet Leaderboard tidak ketemu, respons error berisi **daftar nama sheet** yang ada — cocokkan `sheetName` di config.
- `produksi/` (penarik) tidak tersentuh dan tetap jalan terpisah.
