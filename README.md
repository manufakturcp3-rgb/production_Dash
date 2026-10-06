# 🏭 Central Production 3 Majalengka — Dashboard System

Sistem dashboard produksi one-repo, **tanpa Supabase** — display baca langsung
dari Apps Script Web App:

1. **`produksi/`** — PENARIK: Apps Script narik 2 spreadsheet sumber jadi 1 sheet target
2. **`backend/`** — PENYEDIA: Apps Script Web App (`doGet` JSON) untuk display
3. **`display/`** — Astro dashboard (deploy di Vercel), fetch ke URL Web App

```
production_Dash/
├── produksi/   ← Apps Script: sync 2 sumber → 1 sheet (trigger 15 mnt)
├── backend/    ← Apps Script Web App: sheet → JSON (cache 10 mnt)
└── display/    ← Astro app: baca PUBLIC_BACKEND_URL (poll 3 mnt)
```

## 🚀 Deploy

### 1. Produksi (penarik) — sudah jalan, jangan diubah
Project lama `syncAllData` + trigger 15 menit. Tidak tersentuh.

### 2. Backend (penyedia JSON)
1. https://script.google.com → New project → paste `backend/Backend.js`
2. Deploy → New deployment → Web app → Execute as **Me**, akses **Anyone**
3. Copy URL `.../exec`. Tes: `URL?action=leaderboard`
4. Edit berikutnya: Manage deployments → New version (URL tetap)

### 3. Display (Vercel)
1. Import repo GitHub ini → **Root Directory = `display`**
2. Env var: `PUBLIC_BACKEND_URL` = URL Web app di atas
3. Deploy (auto-redeploy tiap push ke `main`)

Lokal: isi `display/.env` dengan URL yang sama, lalu `npm run dev`.

## 🔄 Alur Data

```
Spreadsheet sumber ──→ [produksi/syncAllData] ──→ Sheet target ──┐
                                                                 ├─→ [backend/doGet] ─→ [display] ─→ Vercel
DASHBOARD KHUSUS + Leaderboard 5R ──────────────────────────────┘
```
