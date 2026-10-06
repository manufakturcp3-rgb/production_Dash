# 🏭 Central Plant 3 Majalengka — Dashboard System

Sistem dashboard produksi one-repo yang terdiri dari:
1. **`produksi/`** — Sync Spreadsheet → Spreadsheet target + Firebase
2. **`featch/`** — Sync Spreadsheet → **Supabase** (hanya data yang berubah)
3. **`display/`** — Dashboard Astro yang terhubung ke Supabase, deploy di Vercel

---

## 🗂️ Struktur Folder

```
appscript/
├── produksi/        ← Apps Script: Sync ke Spreadsheet target & Firebase
├── featch/          ← Apps Script: Sync ke Supabase (smart delta sync)
├── display/         ← Astro app: Dashboard UI (deploy di Vercel)
└── eng/             ← Project Apps Script lama
```

---

## 🚀 Setup Lengkap

### Langkah 1: Setup Supabase

1. Buka [Supabase Dashboard](https://supabase.com/dashboard/project/awnjjzyyyxflvhpsnwby)
2. Buka **SQL Editor**
3. Jalankan file [`featch/SUPABASE_SCHEMA.sql`](featch/SUPABASE_SCHEMA.sql)
4. Aktifkan Realtime untuk tabel `sync_log` di **Database > Replication**

### Langkah 2: Aktifkan Apps Script API

1. Buka https://script.google.com/home/usersettings
2. Aktifkan **"Google Apps Script API"**

### Langkah 3: Push kode `featch` ke Apps Script

```bash
cd featch
clasp push --force
```

4. Buka [Apps Script Editor](https://script.google.com/u/2/home/projects/13mqVmSQ51RihG2ywGCYx-QSg0o7N7L5kncbIsjxXcs5FApyZAnJDlE9u/edit)
5. Jalankan `forceFullSync()` untuk sync pertama kali
6. Jalankan `setupTrigger15Min()` untuk trigger otomatis

### Langkah 4: Deploy Astro ke Vercel

```bash
cd display
# Push ke GitHub dulu (jika belum)
git init && git add . && git commit -m "Initial dashboard"
git remote add origin https://github.com/username/plant3-dashboard.git
git push -u origin main
```

Lalu di Vercel:
1. Import project dari GitHub
2. Framework: **Astro**
3. Environment Variables:
   - `PUBLIC_SUPABASE_URL` = `https://awnjjzyyyxflvhpsnwby.supabase.co`
   - `PUBLIC_SUPABASE_ANON_KEY` = `sb_publishable_F2QptrbgTXqFo8pZvtNgZg_mAfhuUSS`
4. Deploy!

---

## 🔄 Alur Data

```
Spreadsheet 1 (WebData_Produk, filter CP3)  ─┐
                                              ├─→ [featch/Featch.js] ─→ Supabase ─→ [display/] ─→ Vercel
Spreadsheet 2 (DASHBOARD KHUSUS, baris 6+) ─┘
```

- **featch** berjalan otomatis setiap 15 menit (via Apps Script trigger)
- **display** subscribe ke Supabase Realtime → update dashboard otomatis tanpa refresh

---

## 📊 Fitur Dashboard

- 🏷️ **Filter Produk & Shift** — klik pill untuk ganti, atau gunakan Mode TV
- 📈 **KPI Cards** — Yield, Productivity, GMP, Plan PPIC, Waste, Output vs Std, 5R, Process Score, Product Score
- 🏆 **Leaderboard** — Ranking area berdasarkan Process Score (dengan trend naik/turun)
- 📡 **Realtime** — Update otomatis via Supabase Realtime subscription
- 📺 **Mode TV** — Otomatis cycling produk & shift untuk display board

---

## ⚡ Development

```bash
# Jalankan dashboard lokal
cd display
npm install
npm run dev    # → http://localhost:4321

# Build untuk production
npm run build

# Push Apps Script
cd featch
clasp push --force
```
