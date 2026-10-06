/**
 * ============================================================
 * FEATCH.JS — Sinkronisasi Spreadsheet → Supabase
 * ============================================================
 * Mengambil data dari 3 sumber dan mengirim ke Supabase
 * hanya jika ada perubahan (smart delta sync).
 *
 * Sumber 1: WebData_Produk (filter Plant CP3, kolom C)
 * Sumber 2: DASHBOARD KHUSUS (mulai baris 6)
 * Sumber 3: Leaderboard 5R — Top Late/Overtime + Foreman per Produk
 * Target  : Supabase REST API
 * ============================================================
 */

const FEATCH_CONFIG = {
  SOURCE_1: {
    name: 'WebData_Produk',
    spreadsheetId: '1P34UU9Q1pN4afM9L8Mn23LD90QdnbQ7Y',
    sheetName: 'WebData_Produk',
    targetPlant: 'cp3',
    supabaseTable: 'webdata_produk'
  },
  SOURCE_2: {
    name: 'DASHBOARD_KHUSUS',
    spreadsheetId: '1OrmtFMggqx0j5uW_X5Nxfo23ty61F7m6ppKWEPvPu0s',
    sheetName: 'DASHBOARD KHUSUS',
    startRow: 6,
    supabaseTable: 'dashboard_khusus'
  },
  SOURCE_3: {
    name: 'Leaderboard_5R',
    // Spreadsheet target (yang ada sheet Leaderboard 5R)
    spreadsheetId: '146f5qPWBsDEyIn1e6WpjN-bfkGZXwGQfdqlybhsXdS4',
    sheetName: 'Leaderboard 5R',
    supabaseTable_ranking: 'leaderboard_5r_ranking',   // Top Late & Overtime
    supabaseTable_foreman: 'leaderboard_5r_foreman'    // Foreman per Produk + foto
  },
  SUPABASE: {
    URL: 'https://awnjjzyyyxflvhpsnwby.supabase.co',
    ANON_KEY: 'sb_publishable_F2QptrbgTXqFo8pZvtNgZg_mAfhuUSS',
    ENABLED: true
  },
  // Cache di PropertiesService untuk deteksi perubahan
  CACHE_KEY_1: 'featch_hash_source1',
  CACHE_KEY_2: 'featch_hash_source2',
  CACHE_KEY_3: 'featch_hash_source3'
};

// ============================================================
// ENTRY POINT — dipanggil manual atau oleh trigger
// ============================================================
function runFeatch() {
  Logger.log('=== [FEATCH] Memulai sinkronisasi... ===');
  const results = { source1: null, source2: null, source3: null };

  try {
    results.source1 = featchSource1();
  } catch (e) {
    Logger.log('[FEATCH] ERROR Source1: ' + e.message);
  }

  try {
    results.source2 = featchSource2();
  } catch (e) {
    Logger.log('[FEATCH] ERROR Source2: ' + e.message);
  }

  try {
    results.source3 = featchSource3();
  } catch (e) {
    Logger.log('[FEATCH] ERROR Source3 (Leaderboard 5R): ' + e.message);
  }

  // Update timestamp sync terakhir di Supabase
  if (FEATCH_CONFIG.SUPABASE.ENABLED) {
    upsertSupabase('sync_log', [{
      id: 'last_sync',
      timestamp: new Date().toISOString(),
      source1_status: results.source1 ? results.source1.status : 'ERROR',
      source1_rows: results.source1 ? results.source1.rows : 0,
      source2_status: results.source2 ? results.source2.status : 'ERROR',
      source2_rows: results.source2 ? results.source2.rows : 0,
      source3_status: results.source3 ? results.source3.status : 'ERROR',
      source3_rows: results.source3 ? results.source3.rows : 0
    }]);
  }

  Logger.log('=== [FEATCH] Selesai. ===');
  return results;
}

// ============================================================
// SOURCE 1: WebData_Produk — filter Plant CP3 (kolom C)
// ============================================================
function featchSource1() {
  const cfg = FEATCH_CONFIG.SOURCE_1;
  Logger.log('[FEATCH] Membaca Source 1: ' + cfg.name);

  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + cfg.sheetName);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { status: 'SKIP', rows: 0, reason: 'Sheet kosong' };
  }

  const raw = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  if (raw.length === 0) return { status: 'SKIP', rows: 0 };

  const header = raw[0];
  const target = String(cfg.targetPlant || 'cp3').toLowerCase().trim();

  // Filter baris yang memiliki Plant = CP3 (kolom C = index 2)
  const dataRows = raw.slice(1).filter(row => {
    const plant = String(row[2] || '').toLowerCase().trim().replace(/\s+/g, '');
    return plant === target || plant === target.replace(/\s+/g, '');
  });

  // Deteksi perubahan via hash sederhana
  const newHash = simpleHash(JSON.stringify(dataRows));
  const oldHash = PropertiesService.getScriptProperties().getProperty(FEATCH_CONFIG.CACHE_KEY_1);

  if (newHash === oldHash) {
    Logger.log('[FEATCH] Source 1: Tidak ada perubahan, skip upload.');
    return { status: 'UNCHANGED', rows: dataRows.length };
  }

  // Convert ke array of objects
  const records = dataRows.map((row, idx) => {
    const obj = { row_index: idx + 2 }; // +2 karena baris 1 adalah header
    header.forEach((h, c) => {
      const key = sanitizeKey(String(h || 'col_' + (c + 1)));
      const val = row[c];
      obj[key] = val instanceof Date ? val.toISOString() : (val === '' ? null : val);
    });
    return obj;
  });

  if (FEATCH_CONFIG.SUPABASE.ENABLED && records.length > 0) {
    // Hapus data lama lalu insert baru
    deleteSupabaseTable(cfg.supabaseTable);
    const batchSize = 100;
    for (let i = 0; i < records.length; i += batchSize) {
      upsertSupabase(cfg.supabaseTable, records.slice(i, i + batchSize));
    }
  }

  // Simpan hash baru
  PropertiesService.getScriptProperties().setProperty(FEATCH_CONFIG.CACHE_KEY_1, newHash);
  Logger.log('[FEATCH] Source 1 berhasil: ' + records.length + ' baris dikirim.');

  return { status: 'UPDATED', rows: records.length };
}

// ============================================================
// SOURCE 2: DASHBOARD KHUSUS — mulai baris 6, expand header
// ============================================================
function featchSource2() {
  const cfg = FEATCH_CONFIG.SOURCE_2;
  Logger.log('[FEATCH] Membaca Source 2: ' + cfg.name);

  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + cfg.sheetName);

  const startRow = cfg.startRow || 6;
  const headerStartRow = 2; // Baris header mulai dari baris 2 (baris 1 = judul)
  const numHeaderRows = startRow - headerStartRow; // = 4 baris header (2,3,4,5)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < startRow || lastCol < 1) {
    return { status: 'SKIP', rows: 0, reason: 'Tidak ada data di bawah baris ' + startRow };
  }

  // Ambil baris header (baris 2–5) untuk menyusun nama kolom
  const headerMatrix = sheet.getRange(headerStartRow, 1, numHeaderRows, lastCol).getValues();
  // Isi forward-fill untuk merged cells di header
  for (let r = 0; r < headerMatrix.length; r++) {
    let lastVal = '';
    for (let c = 2; c < headerMatrix[r].length; c++) {
      const v = String(headerMatrix[r][c] || '').trim();
      if (v !== '') { lastVal = v; }
      else if (lastVal !== '') { headerMatrix[r][c] = lastVal; }
    }
    // Kolom A dan B selalu fix
    if (!headerMatrix[r][0]) headerMatrix[r][0] = 'ISOWEEK';
    if (!headerMatrix[r][1]) headerMatrix[r][1] = 'TANGGAL';
  }

  // Buat nama kolom gabungan dari semua baris header
  const colNames = [];
  for (let c = 0; c < lastCol; c++) {
    if (c === 0) { colNames.push('isoweek'); continue; }
    if (c === 1) { colNames.push('tanggal'); continue; }
    const parts = headerMatrix
      .map(r => String(r[c] || '').trim())
      .filter(p => p !== '');
    const unique = [...new Set(parts)];
    colNames.push(sanitizeKey(unique.join('__') || 'col_' + (c + 1)));
  }

  // Ambil baris data (mulai baris startRow)
  const numDataRows = lastRow - startRow + 1;
  const dataMatrix = sheet.getRange(startRow, 1, numDataRows, lastCol).getValues();

  // Deteksi perubahan
  const newHash = simpleHash(JSON.stringify(dataMatrix));
  const oldHash = PropertiesService.getScriptProperties().getProperty(FEATCH_CONFIG.CACHE_KEY_2);

  if (newHash === oldHash) {
    Logger.log('[FEATCH] Source 2: Tidak ada perubahan, skip upload.');
    return { status: 'UNCHANGED', rows: dataMatrix.length };
  }

  // Filter baris kosong dan convert ke objects
  const records = [];
  dataMatrix.forEach((row, idx) => {
    const hasData = row.some(cell => cell !== '' && cell !== null && cell !== undefined);
    if (!hasData) return;

    const obj = { row_index: startRow + idx };
    colNames.forEach((key, c) => {
      const val = row[c];
      obj[key] = val instanceof Date ? val.toISOString() : (val === '' ? null : val);
    });
    records.push(obj);
  });

  if (FEATCH_CONFIG.SUPABASE.ENABLED && records.length > 0) {
    deleteSupabaseTable(cfg.supabaseTable);
    const batchSize = 100;
    for (let i = 0; i < records.length; i += batchSize) {
      upsertSupabase(cfg.supabaseTable, records.slice(i, i + batchSize));
    }
  }

  PropertiesService.getScriptProperties().setProperty(FEATCH_CONFIG.CACHE_KEY_2, newHash);
  Logger.log('[FEATCH] Source 2 berhasil: ' + records.length + ' baris dikirim.');

  return { status: 'UPDATED', rows: records.length };
}

// ============================================================
// SOURCE 3: Leaderboard 5R — Top Late/Overtime + Foreman per Produk
// Sheet: 'Leaderboard 5R' di spreadsheet 146f5qPWBsDEyIn1e6WpjN-bfkGZXwGQfdqlybhsXdS4
// Layout (baris 1 = header):
//   A=RANK | B=TOP_LATE | C=TOP_OVERTIME | ... | G=PROD |
//   H=FOREMAN_1 | I=IMG | J=FOREMAN_2 | K=IMG | L=FOREMAN_3 | M=IMG
// Konsep: img disimpan sebagai URL teks saja. Kalau link di
// spreadsheet berubah, hash berubah -> Supabase ikut ter-update.
// ============================================================
function featchSource3() {
  const cfg = FEATCH_CONFIG.SOURCE_3;
  Logger.log('[FEATCH] Membaca Source 3: ' + cfg.name);

  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + cfg.sheetName);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { status: 'SKIP', rows: 0, reason: 'Sheet kosong' };
  }

  const raw = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  if (raw.length < 2) return { status: 'SKIP', rows: 0 };

  // --- Deteksi posisi kolom dari header (robust terhadap geser kolom) ---
  const norm = (v) => String(v || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const header = raw[0].map(norm);

  const findCol = (names, fallback) => {
    for (let i = 0; i < header.length; i++) {
      if (names.indexOf(header[i]) !== -1) return i;
    }
    return fallback;
  };

  const cRank = findCol(['rank', 'no', 'ranking'], 0);
  const cTopLate = findCol(['toplate', 'topterlambat', 'late'], 1);
  const cTopOt = findCol(['topovertime', 'overtime', 'toplembur'], 2);
  const cProd = findCol(['prod', 'produk', 'product'], 6);

  // FOREMAN_1/2/3: cari semua kolom foreman berurutan
  const foremanCols = [];
  for (let i = 0; i < header.length; i++) {
    if (/^foreman\d*$/.test(header[i]) || header[i] === 'foreman') foremanCols.push(i);
  }
  // IMG: semua kolom img berurutan
  const imgCols = [];
  for (let i = 0; i < header.length; i++) {
    if (/^(img|image|foto|photo|linkimg|urlimg)$/.test(header[i])) imgCols.push(i);
  }

  // Pasangan foreman -> img: img terdekat di kanan setiap foreman.
  // Fallback ke posisi default H,I,J,K,L,M (index 7..12) sesuai screenshot.
  let pairs = [];
  if (foremanCols.length > 0) {
    foremanCols.slice(0, 3).forEach((fc, n) => {
      let ic = -1;
      for (let k = 0; k < imgCols.length; k++) {
        if (imgCols[k] > fc && (k === imgCols.length - 1 || imgCols[k + 1] > (foremanCols[n + 1] || 999))) { ic = imgCols[k]; break; }
      }
      // fallback sederhana: img ke-n
      if (ic === -1 && imgCols[n] !== undefined) ic = imgCols[n];
      pairs.push({ f: fc, img: ic });
    });
  } else {
    pairs = [{ f: 7, img: 8 }, { f: 9, img: 10 }, { f: 11, img: 12 }];
  }
  while (pairs.length < 3) pairs.push({ f: 7 + pairs.length * 2, img: 8 + pairs.length * 2 });

  const cellStr = (row, c) => {
    if (c === -1 || c === undefined || c >= row.length) return null;
    const v = row[c];
    if (v === '' || v === null || v === undefined) return null;
    return String(v).trim() || null;
  };

  // --- Tabel 1: ranking Top Late / Top Overtime (kolom A-C) ---
  const ranking = [];
  for (let r = 1; r < raw.length; r++) {
    const row = raw[r];
    const topLate = cellStr(row, cTopLate);
    const topOt = cellStr(row, cTopOt);
    const rankRaw = cellStr(row, cRank);
    if (!topLate && !topOt) continue; // lewati baris kosong
    const rankNum = parseInt(rankRaw, 10);
    ranking.push({
      row_index: r + 1,
      rank: isNaN(rankNum) ? (ranking.length + 1) : rankNum,
      top_late: topLate,
      top_overtime: topOt
    });
    if (ranking.length >= 20) break; // batas aman
  }

  // --- Tabel 2: foreman per produk (kolom G-M) ---
  const foremen = [];
  for (let r = 1; r < raw.length; r++) {
    const row = raw[r];
    const prod = cellStr(row, cProd);
    if (!prod) continue;
    foremen.push({
      row_index: r + 1,
      prod: prod.toUpperCase(),
      foreman_1: cellStr(row, pairs[0].f),
      img_1: cellStr(row, pairs[0].img),
      foreman_2: cellStr(row, pairs[1].f),
      img_2: cellStr(row, pairs[1].img),
      foreman_3: cellStr(row, pairs[2].f),
      img_3: cellStr(row, pairs[2].img)
    });
    if (foremen.length >= 50) break;
  }

  // --- Deteksi perubahan (satu hash untuk kedua tabel) ---
  const newHash = simpleHash(JSON.stringify({ ranking: ranking, foremen: foremen }));
  const oldHash = PropertiesService.getScriptProperties().getProperty(FEATCH_CONFIG.CACHE_KEY_3);
  if (newHash === oldHash) {
    Logger.log('[FEATCH] Source 3: Tidak ada perubahan, skip upload.');
    return { status: 'UNCHANGED', rows: ranking.length + foremen.length };
  }

  if (FEATCH_CONFIG.SUPABASE.ENABLED) {
    if (ranking.length > 0) {
      deleteSupabaseTable(cfg.supabaseTable_ranking);
      upsertSupabase(cfg.supabaseTable_ranking, ranking);
    }
    if (foremen.length > 0) {
      deleteSupabaseTable(cfg.supabaseTable_foreman);
      upsertSupabase(cfg.supabaseTable_foreman, foremen);
    }
  }

  PropertiesService.getScriptProperties().setProperty(FEATCH_CONFIG.CACHE_KEY_3, newHash);
  Logger.log('[FEATCH] Source 3 berhasil: ' + ranking.length + ' ranking, ' + foremen.length + ' produk foreman.');

  return { status: 'UPDATED', rows: ranking.length + foremen.length };
}

// ============================================================
// SUPABASE REST API HELPERS
// ============================================================

/**
 * Upsert (insert or update) ke Supabase table.
 * Menggunakan POST dengan Prefer: resolution=merge-duplicates
 */
function upsertSupabase(tableName, records) {
  if (!records || records.length === 0) return;
  const url = `${FEATCH_CONFIG.SUPABASE.URL}/rest/v1/${tableName}`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'apikey': FEATCH_CONFIG.SUPABASE.ANON_KEY,
      'Authorization': 'Bearer ' + FEATCH_CONFIG.SUPABASE.ANON_KEY,
      'Prefer': 'return=minimal,resolution=merge-duplicates'
    },
    payload: JSON.stringify(records),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  if (code >= 200 && code < 300) {
    Logger.log(`[SUPABASE] Upsert ${tableName}: ${records.length} records OK (${code})`);
  } else {
    Logger.log(`[SUPABASE] ERROR upsert ${tableName}: HTTP ${code} — ${res.getContentText().substring(0, 300)}`);
  }
}

/**
 * Hapus semua data di table Supabase sebelum insert ulang.
 * Menggunakan filter neq pada row_index >= 0 (semua baris).
 */
function deleteSupabaseTable(tableName) {
  const url = `${FEATCH_CONFIG.SUPABASE.URL}/rest/v1/${tableName}?row_index=gte.0`;
  const options = {
    method: 'delete',
    headers: {
      'apikey': FEATCH_CONFIG.SUPABASE.ANON_KEY,
      'Authorization': 'Bearer ' + FEATCH_CONFIG.SUPABASE.ANON_KEY,
      'Prefer': 'return=minimal'
    },
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  Logger.log(`[SUPABASE] Delete ${tableName}: HTTP ${res.getResponseCode()}`);
}

// ============================================================
// UTILITY FUNCTIONS
// ============================================================

/** Hash string sederhana untuk deteksi perubahan data */
function simpleHash(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return String(hash);
}

/** Bersihkan string jadi key yang valid untuk Supabase (huruf kecil, tanpa karakter aneh) */
function sanitizeKey(str) {
  return String(str)
    .toLowerCase()
    .replace(/[^a-z0-9_]/g, '_')
    .replace(/_{2,}/g, '_')
    .replace(/^_|_$/g, '')
    .substring(0, 63) // Supabase max column name length
    || 'col';
}

// ============================================================
// TRIGGER MANAGEMENT
// ============================================================

function setupTrigger15Min() {
  removeAllFeatchTriggers();
  ScriptApp.newTrigger('runFeatch')
    .timeBased()
    .everyMinutes(15)
    .create();
  Logger.log('[FEATCH] Trigger 15 menit berhasil dibuat.');
}

function setupTrigger30Min() {
  removeAllFeatchTriggers();
  ScriptApp.newTrigger('runFeatch')
    .timeBased()
    .everyMinutes(30)
    .create();
  Logger.log('[FEATCH] Trigger 30 menit berhasil dibuat.');
}

function setupTrigger1Hour() {
  removeAllFeatchTriggers();
  ScriptApp.newTrigger('runFeatch')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('[FEATCH] Trigger 1 jam berhasil dibuat.');
}

function removeAllFeatchTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runFeatch')
    .forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('[FEATCH] Semua trigger dihapus.');
}

/** Paksa reset cache dan jalankan sinkronisasi ulang penuh */
function forceFullSync() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(FEATCH_CONFIG.CACHE_KEY_1);
  props.deleteProperty(FEATCH_CONFIG.CACHE_KEY_2);
  props.deleteProperty(FEATCH_CONFIG.CACHE_KEY_3);
  Logger.log('[FEATCH] Cache direset, menjalankan sync penuh...');
  return runFeatch();
}
