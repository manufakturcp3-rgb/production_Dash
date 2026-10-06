-- ============================================================
-- SUPABASE SCHEMA — Central Plant 3 Dashboard
-- Jalankan di Supabase Dashboard > SQL Editor
-- ============================================================

-- Aktifkan Realtime untuk tabel-tabel ini
-- (Supabase Dashboard > Database > Replication > pilih tabel)

-- 1. Tabel WebData Produk (Source 1 - Filter CP3)
CREATE TABLE IF NOT EXISTS public.webdata_produk (
  id          BIGSERIAL    PRIMARY KEY,
  row_index   INTEGER      NOT NULL,
  -- Kolom-kolom berikut dibuat dari header Spreadsheet.
  -- Nama kolom bergantung pada header di spreadsheet.
  -- Kolom dibuat otomatis saat pertama kali data diinsert.
  created_at  TIMESTAMPTZ  DEFAULT NOW(),
  updated_at  TIMESTAMPTZ  DEFAULT NOW()
);

-- 2. Tabel Dashboard Khusus (Source 2 - mulai baris 6)
CREATE TABLE IF NOT EXISTS public.dashboard_khusus (
  id          BIGSERIAL    PRIMARY KEY,
  row_index   INTEGER      NOT NULL,
  isoweek     TEXT,
  tanggal     TEXT,
  -- Kolom lain berformat: produk__kategori__shift
  -- Contoh: siomay__actual_fg__shift_1
  created_at  TIMESTAMPTZ  DEFAULT NOW(),
  updated_at  TIMESTAMPTZ  DEFAULT NOW()
);

-- 3. Tabel log sinkronisasi
CREATE TABLE IF NOT EXISTS public.sync_log (
  id              TEXT         PRIMARY KEY DEFAULT 'last_sync',
  timestamp       TIMESTAMPTZ,
  source1_status  TEXT,
  source1_rows    INTEGER,
  source2_status  TEXT,
  source2_rows    INTEGER,
  source3_status  TEXT,
  source3_rows    INTEGER
);

-- Kolom source3 untuk database yang sudah dibuat sebelum update ini
ALTER TABLE public.sync_log ADD COLUMN IF NOT EXISTS source3_status TEXT;
ALTER TABLE public.sync_log ADD COLUMN IF NOT EXISTS source3_rows INTEGER;

-- 4. Leaderboard 5R — Top Late & Top Overtime (Source 3a)
-- Sheet 'Leaderboard 5R', kolom A-C: RANK | TOP_LATE | TOP_OVERTIME
CREATE TABLE IF NOT EXISTS public.leaderboard_5r_ranking (
  id            BIGSERIAL    PRIMARY KEY,
  row_index     INTEGER      NOT NULL,
  rank          INTEGER,
  top_late      TEXT,
  top_overtime  TEXT,
  created_at    TIMESTAMPTZ  DEFAULT NOW(),
  updated_at    TIMESTAMPTZ  DEFAULT NOW()
);

-- 5. Leaderboard 5R — Foreman per Produk + foto (Source 3b)
-- Sheet 'Leaderboard 5R', kolom G-M: PROD | FOREMAN_1 | IMG | FOREMAN_2 | IMG | FOREMAN_3 | IMG
-- IMG disimpan sebagai URL teks saja; kalau link di spreadsheet berubah,
-- sync berikutnya akan meng-update baris ini di Supabase.
CREATE TABLE IF NOT EXISTS public.leaderboard_5r_foreman (
  id          BIGSERIAL    PRIMARY KEY,
  row_index   INTEGER      NOT NULL,
  prod        TEXT,
  foreman_1   TEXT,
  img_1       TEXT,
  foreman_2   TEXT,
  img_2       TEXT,
  foreman_3   TEXT,
  img_3       TEXT,
  created_at  TIMESTAMPTZ  DEFAULT NOW(),
  updated_at  TIMESTAMPTZ  DEFAULT NOW()
);

-- ============================================================
-- Row Level Security (RLS) — Baca Publik, Tulis via Service Key
-- ============================================================

ALTER TABLE public.webdata_produk  ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.dashboard_khusus ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.sync_log         ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.leaderboard_5r_ranking ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.leaderboard_5r_foreman ENABLE ROW LEVEL SECURITY;

-- Policy: semua user bisa baca (SELECT)
CREATE POLICY "Allow public read on webdata_produk"
  ON public.webdata_produk FOR SELECT USING (true);

CREATE POLICY "Allow public read on dashboard_khusus"
  ON public.dashboard_khusus FOR SELECT USING (true);

CREATE POLICY "Allow public read on sync_log"
  ON public.sync_log FOR SELECT USING (true);

-- Policy: hanya service role yang bisa tulis
-- (Apps Script menggunakan anon key → perlu allow insert/update/delete untuk anon)
-- Jika ingin lebih aman, gunakan service key di Apps Script.
CREATE POLICY "Allow anon insert webdata_produk"
  ON public.webdata_produk FOR INSERT WITH CHECK (true);

CREATE POLICY "Allow anon delete webdata_produk"
  ON public.webdata_produk FOR DELETE USING (true);

CREATE POLICY "Allow anon insert dashboard_khusus"
  ON public.dashboard_khusus FOR INSERT WITH CHECK (true);

CREATE POLICY "Allow anon delete dashboard_khusus"
  ON public.dashboard_khusus FOR DELETE USING (true);

CREATE POLICY "Allow anon upsert sync_log"
  ON public.sync_log FOR ALL USING (true) WITH CHECK (true);

CREATE POLICY "Allow public read on leaderboard_5r_ranking"
  ON public.leaderboard_5r_ranking FOR SELECT USING (true);

CREATE POLICY "Allow anon write leaderboard_5r_ranking"
  ON public.leaderboard_5r_ranking FOR ALL USING (true) WITH CHECK (true);

CREATE POLICY "Allow public read on leaderboard_5r_foreman"
  ON public.leaderboard_5r_foreman FOR SELECT USING (true);

CREATE POLICY "Allow anon write leaderboard_5r_foreman"
  ON public.leaderboard_5r_foreman FOR ALL USING (true) WITH CHECK (true);

-- ============================================================
-- Aktifkan Realtime (jalankan di SQL Editor Supabase)
-- ============================================================
-- Atau aktifkan dari Dashboard > Database > Replication

BEGIN;
  ALTER PUBLICATION supabase_realtime ADD TABLE public.sync_log;
  ALTER PUBLICATION supabase_realtime ADD TABLE public.webdata_produk;
  ALTER PUBLICATION supabase_realtime ADD TABLE public.dashboard_khusus;
  ALTER PUBLICATION supabase_realtime ADD TABLE public.leaderboard_5r_ranking;
  ALTER PUBLICATION supabase_realtime ADD TABLE public.leaderboard_5r_foreman;
COMMIT;
