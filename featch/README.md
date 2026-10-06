# 📥 Featch — Sinkronisasi Spreadsheet → Supabase

Script Apps Script yang mengambil data dari 2 spreadsheet sumber dan mengirimnya ke Supabase secara otomatis.

## Fitur
- ✅ Hanya mengirim data yang **berubah** (menggunakan hash cache)
- ✅ Filter data Source 1 hanya untuk **Plant CP3** (kolom C)
- ✅ Data Source 2 mulai dari **baris 6**, dengan nama kolom gabungan dari header multi-baris (baris 2–5)
- ✅ Trigger otomatis setiap 15/30 menit atau 1 jam
- ✅ Sync log disimpan di tabel `sync_log` Supabase

## Sumber Data
| # | Spreadsheet | Sheet | Keterangan |
|---|-------------|-------|------------|
| 1 | `1P34UU9Q1pN4afM9L8Mn23LD90QdnbQ7Y` | WebData_Produk | Filter Plant CP3 |
| 2 | `1OrmtFMggqx0j5uW_X5Nxfo23ty61F7m6ppKWEPvPu0s` | DASHBOARD KHUSUS | Mulai baris 6 |

## Tabel Supabase yang Dibutuhkan

Buat tabel berikut di Supabase Dashboard (SQL Editor):

```sql
-- Tabel 1: WebData Produk (Source 1)
CREATE TABLE IF NOT EXISTS webdata_produk (
  id BIGSERIAL PRIMARY KEY,
  row_index INTEGER,
  -- Kolom lain akan dibuat otomatis sesuai header spreadsheet
  -- Jalankan forceFullSync() dulu untuk melihat kolom apa saja yang ada
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Tabel 2: Dashboard Khusus (Source 2)
CREATE TABLE IF NOT EXISTS dashboard_khusus (
  id BIGSERIAL PRIMARY KEY,
  row_index INTEGER,
  isoweek TEXT,
  tanggal TEXT,
  -- Kolom produk__kategori__shift akan dibuat dari header multi-baris
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Tabel 3: Log sinkronisasi terakhir
CREATE TABLE IF NOT EXISTS sync_log (
  id TEXT PRIMARY KEY DEFAULT 'last_sync',
  timestamp TIMESTAMPTZ,
  source1_status TEXT,
  source1_rows INTEGER,
  source2_status TEXT,
  source2_rows INTEGER
);
```

## Cara Pakai

### 1. Push ke Apps Script
```bash
cd featch
clasp push
```

### 2. Jalankan pertama kali
Di Apps Script Editor, jalankan fungsi: `forceFullSync()`

### 3. Atur trigger otomatis
Jalankan salah satu:
- `setupTrigger15Min()` — setiap 15 menit
- `setupTrigger30Min()` — setiap 30 menit  
- `setupTrigger1Hour()` — setiap jam

## Fungsi Utama
| Fungsi | Keterangan |
|--------|------------|
| `runFeatch()` | Jalankan sync (hanya kirim jika ada perubahan) |
| `forceFullSync()` | Reset cache dan sync paksa semua data |
| `setupTrigger15Min()` | Atur trigger otomatis 15 menit |
| `removeAllFeatchTriggers()` | Hapus semua trigger |
