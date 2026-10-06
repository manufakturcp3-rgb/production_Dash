-- ============================================================
-- MIGRATION 002 — RPC featch_sync: auto-tambah kolom + replace isi
-- Jalankan sekali di Supabase Dashboard > SQL Editor
-- ============================================================
-- Kenapa perlu: header spreadsheet itu dinamis (kolom 'periode',
-- 'adonan_pangsit_achivement_shift_1_kg', dst). Supabase TIDAK
-- otomatis membuat kolom baru, sehingga upsert langsung via REST
-- gagal dengan error 400 PGRST204 (column not in schema cache).
--
-- RPC ini (SECURITY DEFINER) akan:
--   1. Menambah kolom TEXT untuk setiap key yang belum ada
--      (khusus 'row_index' memakai INTEGER)
--   2. Menghapus seluruh isi tabel lama
--   3. Mengisi ulang dari data JSON yang dikirim Apps Script
-- Hanya 4 tabel dashboard yang diizinkan (allowlist di bawah).
-- ============================================================

CREATE OR REPLACE FUNCTION public.featch_sync(p_table TEXT, p_rows JSONB)
RETURNS JSONB
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $func$
DECLARE
  allowed TEXT[] := ARRAY['webdata_produk', 'dashboard_khusus', 'leaderboard_5r_ranking', 'leaderboard_5r_foreman'];
  cols TEXT[];
  col_list TEXT;
  sel_list TEXT;
  c TEXT;
BEGIN
  IF NOT (p_table = ANY (allowed)) THEN
    RAISE EXCEPTION 'featch_sync: table not allowed: %', p_table;
  END IF;

  IF p_rows IS NULL OR jsonb_typeof(p_rows) != 'array' OR jsonb_array_length(p_rows) = 0 THEN
    RETURN jsonb_build_object('status', 'empty');
  END IF;

  -- Kumpulkan semua keys dari semua baris (kecuali kolom sistem)
  SELECT array_agg(DISTINCT k) INTO cols
  FROM jsonb_array_elements(p_rows) AS elem,
       jsonb_object_keys(elem) AS k
  WHERE k NOT IN ('id', 'created_at', 'updated_at');

  -- Tambah kolom yang belum ada (row_index INTEGER, lainnya TEXT)
  FOREACH c IN ARRAY cols LOOP
    IF c = 'row_index' THEN
      EXECUTE 'ALTER TABLE ' || quote_ident(p_table) || ' ADD COLUMN IF NOT EXISTS ' || quote_ident(c) || ' INTEGER';
    ELSE
      EXECUTE 'ALTER TABLE ' || quote_ident(p_table) || ' ADD COLUMN IF NOT EXISTS ' || quote_ident(c) || ' TEXT';
    END IF;
  END LOOP;

  -- Ganti seluruh isi tabel lama
  EXECUTE 'DELETE FROM ' || quote_ident(p_table);

  -- Susun INSERT dinamis dari keys yang ada
  SELECT string_agg(quote_ident(x), ', ') INTO col_list FROM unnest(cols) AS x;
  SELECT string_agg(
    CASE WHEN x = 'row_index'
      THEN 'NULLIF(value->>' || quote_literal(x) || ','''')::INTEGER'
      ELSE 'value->>' || quote_literal(x)
    END, ', ') INTO sel_list FROM unnest(cols) AS x;

  EXECUTE 'INSERT INTO ' || quote_ident(p_table) || ' (' || col_list || ') ' ||
          'SELECT ' || sel_list || ' FROM jsonb_array_elements($1) AS value'
  USING p_rows;

  RETURN jsonb_build_object('status', 'ok', 'columns', coalesce(array_length(cols, 1), 0));
END;
$func$;

-- Izinkan anon key (yang dipakai Apps Script) memanggil RPC ini
GRANT EXECUTE ON FUNCTION public.featch_sync(TEXT, JSONB) TO anon, authenticated;
