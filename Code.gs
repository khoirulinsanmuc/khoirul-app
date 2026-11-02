/**
 * =====================================================================
 *  KHOIRUL INSAN APP — Backend (Google Apps Script)
 *  Aplikasi web masjid: Beranda, Al-Qur'an, Tasbih, Akun,
 *  Infaq, Keuangan, Kajian, Artikel, Qurban, Petugas, Pengumuman.
 *
 *  Sumber data Al-Qur'an & Jadwal Sholat: API equran.id v2
 *  (data jadwal sholat bersumber dari Bimas Islam Kementerian Agama RI)
 * =====================================================================
 */

// ---------------------------------------------------------------------
// 0. KONFIGURASI UMUM
// ---------------------------------------------------------------------
const APP_TITLE = 'Khoirul Insan App';
const QURAN_API_BASE = 'https://equran.id/api/v2';
const CACHE_TTL_QURAN = 21600;   // 6 jam (data surat statis)
const CACHE_TTL_JADWAL = 3600;   // 1 jam (jadwal sholat)

// Nama-nama tab (sheet) yang dipakai aplikasi ini
const SHEETS = {
  CONFIG: 'Config',
  KAJIAN: 'Kajian',
  ARTIKEL: 'Artikel',
  NASIHAT: 'Nasihat',
  INFAQ: 'Infaq',
  KEUANGAN: 'Keuangan',
  QURBAN: 'Qurban',
  PETUGAS: 'Petugas',
  PENGUMUMAN: 'Pengumuman',
  CAMPAIGN: 'Campaign',
  DONASI: 'DonasiTransaksi'
};

// ---------------------------------------------------------------------
// 1. ENTRY POINT WEB APP
// ---------------------------------------------------------------------
function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.appTitle = APP_TITLE;
  return tpl.evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://em-content.zobj.net/source/apple/391/mosque_1f54c.png');
}

// Dipakai oleh <?!= include('NamaFile'); ?> di dalam file HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------
// 2. SETUP AWAL — jalankan fungsi ini SEKALI dari editor Apps Script
//    untuk membuat semua sheet + header + contoh data.
// ---------------------------------------------------------------------
function setupAplikasi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createSheetIfMissing_(ss, SHEETS.CONFIG, ['Key', 'Value'], [
    ['NamaMasjid', 'Khoirul Insan'],
    ['Tagline', 'Jamaah Khoirul Insan'],
    ['Alamat', 'Tanjung Barat, Jakarta Selatan'],
    ['Provinsi', 'DKI Jakarta'],
    ['KabKota', 'Kota Jakarta Selatan'],
    ['NoWhatsApp', '628123456789'],
    ['NoRekeningInfaq', 'BSI 7001234567 a.n. DKM Khoirul Insan'],
    ['LogoUrl', 'https://em-content.zobj.net/source/apple/391/mosque_1f54c.png']
  ]);

  createSheetIfMissing_(ss, SHEETS.KAJIAN,
    ['id', 'tanggal', 'judul', 'jam', 'pemateri', 'status'], [
    [1, '2026-09-16', 'Kajian Rutin Bulan September: Overthinking dalam Perspektif Islam', '12:30 (Ba\'da Dzuhur)', 'Ustadzah Ronny Setiawan M.Pd', 'Selesai'],
    [2, '2026-10-14', 'Kajian Rutin Bulan Oktober', '12:30 (Ba\'da Dzuhur)', 'Segera Diumumkan', 'Terjadwal'],
    [3, '2026-11-18', 'Kajian Rutin Bulan November', '12:30 (Ba\'da Dzuhur)', 'Segera Diumumkan', 'Terjadwal']
  ]);

  createSheetIfMissing_(ss, SHEETS.ARTIKEL,
    ['id', 'judul', 'kategori', 'gambar_url', 'ringkasan', 'konten', 'tanggal', 'penulis'], [
    [1, 'Keutamaan Sholat Berjamaah 27 Derajat', 'Fiqih Ibadah',
      'https://images.unsplash.com/photo-1564769662533-4f00a87b4056?w=800',
      'Sholat berjamaah memiliki keutamaan yang sangat besar dibanding sholat sendirian.',
      'Rasulullah SAW bersabda bahwa sholat berjamaah lebih utama 27 derajat dibandingkan sholat sendirian. Keutamaan ini mendorong setiap muslim untuk senantiasa menjaga sholat berjamaah di masjid, terutama bagi kaum laki-laki. Selain pahala yang berlipat, sholat berjamaah juga mempererat ukhuwah islamiyah antar jamaah.',
      '2026-08-01', 'Admin DKM'],
    [2, 'Sejarah Pembangunan Masjid Nabawi', 'Sejarah',
      'https://images.unsplash.com/photo-1591123120675-6f7f1aae0e5b?w=800',
      'Menelusuri jejak pembangunan salah satu masjid termulia dalam Islam.',
      'Masjid Nabawi dibangun pertama kali oleh Rasulullah SAW segera setelah hijrah ke Madinah. Masjid ini terus mengalami perluasan dari masa ke masa, mulai dari zaman Khulafaur Rasyidin hingga era modern, menjadikannya salah satu masjid terbesar di dunia saat ini.',
      '2026-07-20', 'Admin DKM']
  ]);

  createSheetIfMissing_(ss, SHEETS.NASIHAT, ['id', 'teks', 'sumber'], [
    [1, 'Barangsiapa yang menempuh suatu jalan untuk menuntut ilmu, maka Allah SWT akan memudahkan baginya jalan menuju surga.', 'HR. Muslim'],
    [2, 'Sebaik-baik manusia adalah yang paling bermanfaat bagi manusia lainnya.', 'HR. Ahmad'],
    [3, 'Jagalah Allah, niscaya Allah akan menjagamu.', 'HR. Tirmidzi']
  ]);

  createSheetIfMissing_(ss, SHEETS.INFAQ,
    ['id', 'timestamp', 'nama', 'email', 'kategori', 'jumlah', 'metode', 'keterangan', 'status'], []);

  createSheetIfMissing_(ss, SHEETS.KEUANGAN,
    ['id', 'tanggal', 'jenis', 'kategori', 'jumlah', 'keterangan', 'sumber'], [
    [1, '2026-08-01', 'Masuk', 'Infaq Jumat', 3250000, 'Infaq kotak Jumat 1 Agustus', 'Kotak Infaq'],
    [2, '2026-08-05', 'Keluar', 'Operasional', 850000, 'Listrik & air bulan Juli', 'Kas Masjid'],
    [3, '2026-08-10', 'Masuk', 'Infaq Online', 1500000, 'Transfer infaq via app', 'Transfer']
  ]);

  createSheetIfMissing_(ss, SHEETS.QURBAN,
    ['id', 'timestamp', 'nama', 'email', 'jenis_hewan', 'jumlah', 'status', 'keterangan'], []);

  createSheetIfMissing_(ss, SHEETS.PETUGAS,
    ['id', 'tanggal', 'khatib', 'imam', 'muadzin'], [
    [1, '2026-08-21', 'Ust. Achmarul Hadi', 'Ust. Achmarul Hadi', 'Sdr. Yopy A W'],
    [2, '2026-08-28', 'Ust. Fulan', 'Ust. Fulan', 'Sdr. Fulan']
  ]);

  createSheetIfMissing_(ss, SHEETS.PENGUMUMAN,
    ['id', 'tanggal', 'judul', 'isi', 'penting'], [
    [1, '2026-08-20', 'Renovasi Tempat Wudhu', 'Mulai 25 Agustus, area wudhu ikhwan akan direnovasi selama 2 minggu. Mohon gunakan area sementara di sisi utara masjid.', true]
  ]);

  createSheetIfMissing_(ss, SHEETS.CAMPAIGN,
    ['id', 'judul', 'gambar_url', 'deskripsi', 'target', 'terkumpul', 'deadline', 'status'], [
    [1, 'Renovasi Tempat Wudhu', 'https://images.unsplash.com/photo-1564769662533-4f00a87b4056?w=800',
      'Bantu kami menyelesaikan renovasi area wudhu ikhwan yang sudah mulai rusak dan bocor.',
      15000000, 4250000, '2026-10-31', 'Aktif']
  ]);

  createSheetIfMissing_(ss, SHEETS.DONASI,
    ['id', 'timestamp', 'campaign_id', 'campaign_judul', 'nama', 'email', 'jumlah', 'keterangan', 'status'], []);

  // Hapus sheet default "Sheet1" jika masih kosong dan belum dipakai
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) {
    try { ss.deleteSheet(def); } catch (err) {}
  }

  SpreadsheetApp.flush();
  return 'Setup selesai! Semua sheet & contoh data berhasil dibuat.';
}

function createSheetIfMissing_(ss, name, headers, sampleRows) {
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  sheet = ss.insertSheet(name);
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f9d58').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  if (sampleRows && sampleRows.length) {
    sheet.getRange(2, 1, sampleRows.length, headers.length).setValues(sampleRows);
  }
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

// Jalankan SEKALI dari editor Apps Script untuk menambah baris QrisUrl ke sheet Config
function tambahConfigQris() {
  const sheet = getSheet_(SHEETS.CONFIG);
  const cfg = configMap_();
  if (cfg.QrisUrl) { Logger.log('QrisUrl sudah ada: ' + cfg.QrisUrl); return; }
  sheet.appendRow(['QrisUrl', 'https://i.imgur.com/Va4vMG6.jpeg']);
  Logger.log('Baris QrisUrl ditambahkan. Silakan edit value-nya manual di sheet Config.');
}

// Jalankan SEKALI dari editor Apps Script untuk menambah harga acuan zakat ke Config
function tambahConfigZakat() {
  const sheet = getSheet_(SHEETS.CONFIG);
  const cfg = configMap_();
  const tambahan = [];
  if (!cfg.HargaEmasPerGram) tambahan.push(['HargaEmasPerGram', 1500000]);
  if (!cfg.HargaPerakPerGram) tambahan.push(['HargaPerakPerGram', 18000]);
  if (!cfg.HargaBerasPerKg) tambahan.push(['HargaBerasPerKg', 15000]);
  if (!tambahan.length) { Logger.log('Semua harga acuan zakat sudah ada di Config.'); return; }
  tambahan.forEach(row => sheet.appendRow(row));
  Logger.log('Ditambahkan: ' + JSON.stringify(tambahan) + '. Silakan sesuaikan angkanya manual di sheet Config.');
}

// ---------------------------------------------------------------------
// 3. HELPER SHEET → JSON
// ---------------------------------------------------------------------
function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" tidak ditemukan. Jalankan setupAplikasi() dulu.');
  return sheet;
}

function sheetToObjects_(name) {
  const sheet = getSheet_(name);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .filter(r => r.join('') !== '') // buang baris kosong
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        obj[h] = v;
      });
      return obj;
    });
}

function configMap_() {
  const rows = sheetToObjects_(SHEETS.CONFIG);
  const map = {};
  rows.forEach(r => map[r.Key] = r.Value);
  return map;
}

function nextId_(name) {
  const rows = sheetToObjects_(name);
  if (!rows.length) return 1;
  const ids = rows.map(r => Number(r.id) || 0);
  return Math.max.apply(null, ids) + 1;
}

// ---------------------------------------------------------------------
// 4. AL-QUR'AN (via equran.id API v2)
// ---------------------------------------------------------------------
function getDaftarSurat() {
  const cache = CacheService.getScriptCache();
  const key = 'daftar_surat_v2';
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { cache.remove(key); }
  }

  const res = UrlFetchApp.fetch(QURAN_API_BASE + '/surat', { muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('API surat gagal, HTTP ' + code + ': ' + res.getContentText().substring(0, 200));
  }

  let json;
  try {
    json = JSON.parse(res.getContentText());
  } catch (e) {
    throw new Error('Response API bukan JSON valid: ' + res.getContentText().substring(0, 200));
  }

  const data = json.data || [];

  // Cache bisa gagal kalau datanya > 100KB (limit CacheService).
  // Jangan biarkan kegagalan cache menggagalkan seluruh fungsi.
  try {
    cache.put(key, JSON.stringify(data), CACHE_TTL_QURAN);
  } catch (e) {
    console.warn('Gagal cache daftar surat (kemungkinan > 100KB): ' + e.message);
  }

  return data;
}
function getDetailSurat(nomor) {
  nomor = Number(nomor);
  const cache = CacheService.getScriptCache();
  const key = 'surat_' + nomor + '_v2';
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { cache.remove(key); }
  }

  const res = UrlFetchApp.fetch(QURAN_API_BASE + '/surat/' + nomor, { muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('API detail surat gagal, HTTP ' + code + ': ' + res.getContentText().substring(0, 200));
  }

  let json;
  try {
    json = JSON.parse(res.getContentText());
  } catch (e) {
    throw new Error('Response API detail surat bukan JSON valid: ' + res.getContentText().substring(0, 200));
  }

  const data = json.data || null;

  if (data) {
    // Surat panjang (Al-Baqarah, Ali 'Imran, dll) bisa > 100KB, limit CacheService.
    // Jangan biarkan kegagalan cache menggagalkan seluruh fungsi.
    try {
      cache.put(key, JSON.stringify(data), CACHE_TTL_QURAN);
    } catch (e) {
      console.warn('Gagal cache surat ' + nomor + ' (kemungkinan > 100KB): ' + e.message);
    }
  }

  return data;
}

function getTafsirSurat(nomor) {
  nomor = Number(nomor);
  const cache = CacheService.getScriptCache();
  const key = 'tafsir_' + nomor + '_v2';
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { cache.remove(key); }
  }

  const res = UrlFetchApp.fetch(QURAN_API_BASE + '/tafsir/' + nomor, { muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('API tafsir gagal, HTTP ' + code + ': ' + res.getContentText().substring(0, 200));
  }

  let json;
  try {
    json = JSON.parse(res.getContentText());
  } catch (e) {
    throw new Error('Response API tafsir bukan JSON valid: ' + res.getContentText().substring(0, 200));
  }

  const data = json.data || null;

  if (data) {
    // Surat panjang bisa > 100KB, sama seperti kasus getDetailSurat sebelumnya
    try {
      cache.put(key, JSON.stringify(data), CACHE_TTL_QURAN);
    } catch (e) {
      console.warn('Gagal cache tafsir surat ' + nomor + ' (kemungkinan > 100KB): ' + e.message);
    }
  }

  return data;
}

function cariSurat(keyword) {
  const semua = getDaftarSurat();
  keyword = (keyword || '').toLowerCase().trim();
  if (!keyword) return semua;
  return semua.filter(s =>
    (s.namaLatin || '').toLowerCase().indexOf(keyword) !== -1 ||
    (s.arti || '').toLowerCase().indexOf(keyword) !== -1 ||
    String(s.nomor) === keyword
  );
}

// Terakhir dibaca — disimpan per akun Google (UserProperties)
function simpanTerakhirDibaca(nomorSurat, namaSurat, nomorAyat) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('terakhir_dibaca', JSON.stringify({
    nomorSurat: nomorSurat, namaSurat: namaSurat, nomorAyat: nomorAyat,
    waktu: new Date().toISOString()
  }));
  return true;
}

function getTerakhirDibaca() {
  const props = PropertiesService.getUserProperties();
  const v = props.getProperty('terakhir_dibaca');
  return v ? JSON.parse(v) : null;
}

// ---------------------------------------------------------------------
// 5. JADWAL SHOLAT (via equran.id API v2 — sumber Kemenag RI)
// ---------------------------------------------------------------------

// Ambil & cache daftar provinsi dari API (24 jam)
function getDaftarProvinsiShalat_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('shalat_provinsi_list');
  if (cached) return JSON.parse(cached);

  const res = UrlFetchApp.fetch(QURAN_API_BASE + '/shalat/provinsi', { muteHttpExceptions: true });
  const json = JSON.parse(res.getContentText());
  const data = json.data || [];
  if (data.length) cache.put('shalat_provinsi_list', JSON.stringify(data), 86400);
  return data;
}

// Ambil & cache daftar kab/kota untuk sebuah provinsi (24 jam)
function getDaftarKabkotaShalat_(provinsiApi) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'shalat_kabkota_' + provinsiApi;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const res = UrlFetchApp.fetch(QURAN_API_BASE + '/shalat/kabkota', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ provinsi: provinsiApi }),
    muteHttpExceptions: true
  });
  const json = JSON.parse(res.getContentText());
  const data = json.data || [];
  if (data.length) cache.put(cacheKey, JSON.stringify(data), 86400);
  return data;
}

// Cari kecocokan nama sedekat mungkin (case-insensitive, toleran spasi/awalan "Kota "/"Kab. ")
function cocokkanNama_(target, daftar) {
  if (!target || !daftar || !daftar.length) return null;
  const norm = s => String(s).toLowerCase().trim()
    .replace(/^kota\s+/, '').replace(/^kab\.?\s+/, '').replace(/^kabupaten\s+/, '');
  const targetNorm = norm(target);

  // 1. cocok persis
  let hit = daftar.find(d => d.toLowerCase().trim() === target.toLowerCase().trim());
  if (hit) return hit;

  // 2. cocok setelah normalisasi awalan Kota/Kab
  hit = daftar.find(d => norm(d) === targetNorm);
  if (hit) return hit;

  // 3. cocok sebagian (contains kedua arah)
  hit = daftar.find(d => norm(d).indexOf(targetNorm) !== -1 || targetNorm.indexOf(norm(d)) !== -1);
  if (hit) return hit;

  return null;
}

// Resolve nama Provinsi & KabKota dari Config ke nama persis yang dikenali API.
// Hasil resolusi di-cache supaya tidak perlu fetch daftar setiap saat.
function resolveLokasiShalat_() {
  const cfg = configMap_();
  const cache = CacheService.getScriptCache();
  const cacheKey = 'lokasi_resolved_' + cfg.Provinsi + '_' + cfg.KabKota;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const provinsiList = getDaftarProvinsiShalat_();
  const provinsiMatch = cocokkanNama_(cfg.Provinsi, provinsiList);
  if (!provinsiMatch) {
    console.error('Provinsi "' + cfg.Provinsi + '" tidak ditemukan di API. Daftar tersedia: ' + JSON.stringify(provinsiList));
    return null;
  }

  const kabkotaList = getDaftarKabkotaShalat_(provinsiMatch);
  const kabkotaMatch = cocokkanNama_(cfg.KabKota, kabkotaList);
  if (!kabkotaMatch) {
    console.error('KabKota "' + cfg.KabKota + '" tidak ditemukan di provinsi "' + provinsiMatch + '". Daftar tersedia: ' + JSON.stringify(kabkotaList));
    return null;
  }

  const resolved = { provinsi: provinsiMatch, kabkota: kabkotaMatch };
  cache.put(cacheKey, JSON.stringify(resolved), 86400);
  return resolved;
}

function fetchJadwalBulanan_(bulan, tahun) {
  const lokasi = resolveLokasiShalat_();
  if (!lokasi) return null;

  const cache = CacheService.getScriptCache();
  const key = 'jadwal_' + lokasi.kabkota + '_' + bulan + '_' + tahun;
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const payload = {
    provinsi: lokasi.provinsi,
    kabkota: lokasi.kabkota,
    bulan: Number(bulan),
    tahun: Number(tahun)
  };
  const res = UrlFetchApp.fetch(QURAN_API_BASE + '/shalat', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const json = JSON.parse(res.getContentText());
  const data = json.data || null;
  if (!data) {
    console.error('Gagal ambil jadwal shalat. Response: ' + res.getContentText());
    return null;
  }
  cache.put(key, JSON.stringify(data), CACHE_TTL_JADWAL);
  return data;
}

// Jalankan fungsi ini manual dari editor Apps Script (Run) untuk mendiagnosa
// kenapa jadwal sholat tidak muncul — hasilnya cek di Execution log.
function debugJadwalSholat() {
  const cfg = configMap_();
  Logger.log('Config Provinsi: ' + cfg.Provinsi);
  Logger.log('Config KabKota: ' + cfg.KabKota);

  const lokasi = resolveLokasiShalat_();
  if (!lokasi) {
    Logger.log('❌ Tidak berhasil mencocokkan lokasi. Lihat log error di atas untuk daftar nama yang valid.');
    return;
  }
  Logger.log('✅ Cocok dengan: ' + JSON.stringify(lokasi));

  const jadwal = getJadwalHariIni();
  Logger.log('Hasil jadwal hari ini: ' + JSON.stringify(jadwal));
}

function getJadwalHariIni() {
  const now = new Date();
  const bulan = now.getMonth() + 1;
  const tahun = now.getFullYear();
  const tgl = now.getDate();

  const bulanData = fetchJadwalBulanan_(bulan, tahun);
  if (!bulanData) return null;

  const hariIni = bulanData.jadwal.find(j => j.tanggal === tgl);
  return {
    provinsi: bulanData.provinsi,
    kabkota: bulanData.kabkota,
    tanggal_lengkap: hariIni ? hariIni.tanggal_lengkap : null,
    hari: hariIni ? hariIni.hari : null,
    jadwal: hariIni || null
  };
}

function getJadwalBulanIni(bulan, tahun) {
  const now = new Date();
  return fetchJadwalBulanan_(bulan || (now.getMonth() + 1), tahun || now.getFullYear());
}

// ---------------------------------------------------------------------
// 5b. DOA & DZIKIR (via equran.id API — /api/doa)
// ---------------------------------------------------------------------
const DOA_API_BASE = 'https://equran.id/api';

function getDaftarDoa() {
  const cache = CacheService.getScriptCache();
  const key = 'daftar_doa_v1';
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { cache.remove(key); }
  }

  const res = UrlFetchApp.fetch(DOA_API_BASE + '/doa', { muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('API doa gagal, HTTP ' + code + ': ' + res.getContentText().substring(0, 200));
  }

  let json;
  try {
    json = JSON.parse(res.getContentText());
  } catch (e) {
    throw new Error('Response API doa bukan JSON valid: ' + res.getContentText().substring(0, 200));
  }

  const data = json.data || [];

  try {
    cache.put(key, JSON.stringify(data), CACHE_TTL_QURAN);
  } catch (e) {
    console.warn('Gagal cache daftar doa (kemungkinan > 100KB): ' + e.message);
  }

  return data;
}

// Jalankan manual dari editor Apps Script untuk cek struktur field asli API doa
function debugDaftarDoa() {
  const data = getDaftarDoa();
  Logger.log('Jumlah doa: ' + data.length);
  Logger.log('Contoh item pertama: ' + JSON.stringify(data[0], null, 2));
}

// ---------------------------------------------------------------------
// 6. BERANDA — agregasi semua data untuk halaman utama
// ---------------------------------------------------------------------
function getBerandaData() {
  const cfg = configMap_();
  const jadwal = getJadwalHariIni();
  const kajian = sheetToObjects_(SHEETS.KAJIAN)
    .sort((a, b) => new Date(a.tanggal) - new Date(b.tanggal))
    .filter(k => new Date(k.tanggal) >= new Date(new Date().toDateString()))
    .slice(0, 3);
  const artikel = sheetToObjects_(SHEETS.ARTIKEL)
    .sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal))
    .slice(0, 6);
  const nasihatList = sheetToObjects_(SHEETS.NASIHAT);
  const nasihat = nasihatList.length ? nasihatList[Math.floor(Math.random() * nasihatList.length)] : null;
  const petugasList = sheetToObjects_(SHEETS.PETUGAS)
    .sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
  const petugasTerdekat = petugasList.length ? petugasList[0] : null;
  const pengumuman = sheetToObjects_(SHEETS.PENGUMUMAN)
    .sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal))
    .slice(0, 5);

  let user = null;
  try {
    const email = Session.getActiveUser().getEmail();
    user = email || null;
  } catch (e) { user = null; }

  return {
    config: cfg,
    jadwal: jadwal,
    kajian: kajian,
    artikel: artikel,
    nasihat: nasihat,
    petugas: petugasTerdekat,
    pengumuman: pengumuman,
    user: user
  };
}

// ---------------------------------------------------------------------
// 7. KAJIAN
// ---------------------------------------------------------------------
function getSemuaKajian() {
  return sheetToObjects_(SHEETS.KAJIAN).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
}

// ---------------------------------------------------------------------
// 8. ARTIKEL
// ---------------------------------------------------------------------
function getSemuaArtikel() {
  return sheetToObjects_(SHEETS.ARTIKEL).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
}

function getArtikelById(id) {
  const all = sheetToObjects_(SHEETS.ARTIKEL);
  return all.find(a => Number(a.id) === Number(id)) || null;
}

// ---------------------------------------------------------------------
// 9. INFAQ
// ---------------------------------------------------------------------
const KATEGORI_INFAQ = ['Infaq Umum', 'Infaq Jumat Berkah', 'Pembangunan Masjid', 'Operasional Masjid', 'Santunan Yatim & Dhuafa'];

function getKategoriInfaq() {
  return KATEGORI_INFAQ;
}

function submitInfaq(data) {
  const sheet = getSheet_(SHEETS.INFAQ);
  const id = nextId_(SHEETS.INFAQ);
  let email = null;
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}

  sheet.appendRow([
    id,
    new Date(),
    data.nama || 'Hamba Allah',
    email || data.email || '',
    data.kategori,
    Number(data.jumlah),
    data.metode || 'Transfer',
    data.keterangan || '',
    'Menunggu Konfirmasi'
  ]);

  // Catat juga sebagai pemasukan di sheet Keuangan (belum terverifikasi)
  const keu = getSheet_(SHEETS.KEUANGAN);
  keu.appendRow([
    nextId_(SHEETS.KEUANGAN),
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    'Masuk',
    data.kategori,
    Number(data.jumlah),
    'Infaq via app oleh ' + (data.nama || 'Hamba Allah') + ' (menunggu konfirmasi)',
    'App'
  ]);

  const cfg = configMap_();
  return {
    id: id,
    qrisUrl: cfg.QrisUrl,
    waAdmin: cfg.NoWhatsApp
  };
}

function getRiwayatInfaqSaya() {
  let email = null;
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!email) return [];
  return sheetToObjects_(SHEETS.INFAQ)
    .filter(i => i.email === email)
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
}

// ---------------------------------------------------------------------
// 10. KEUANGAN
// ---------------------------------------------------------------------
function getLaporanKeuangan() {
  const rows = sheetToObjects_(SHEETS.KEUANGAN).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
  let totalMasuk = 0, totalKeluar = 0;
  rows.forEach(r => {
    if (r.jenis === 'Masuk') totalMasuk += Number(r.jumlah) || 0;
    else totalKeluar += Number(r.jumlah) || 0;
  });
  return {
    saldo: totalMasuk - totalKeluar,
    totalMasuk: totalMasuk,
    totalKeluar: totalKeluar,
    transaksi: rows.slice(0, 30)
  };
}

// ---------------------------------------------------------------------
// 11. QURBAN
// ---------------------------------------------------------------------
function daftarQurban(data) {
  const sheet = getSheet_(SHEETS.QURBAN);
  const id = nextId_(SHEETS.QURBAN);
  let email = null;
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}

  sheet.appendRow([
    id, new Date(), data.nama, email || data.email || '',
    data.jenisHewan, Number(data.jumlah) || 1, 'Menunggu Konfirmasi', data.keterangan || ''
  ]);
  return { id: id };
}

function getDaftarQurban() {
  return sheetToObjects_(SHEETS.QURBAN).sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
}

// ---------------------------------------------------------------------
// 11b. DONASI (Campaign dari Spreadsheet + Pembayaran QRIS Tunggal)
// ---------------------------------------------------------------------
function getSemuaCampaign() {
  return sheetToObjects_(SHEETS.CAMPAIGN)
    .filter(c => String(c.status).toLowerCase() === 'aktif')
    .sort((a, b) => Number(b.id) - Number(a.id));
}

function getCampaignById(id) {
  const all = sheetToObjects_(SHEETS.CAMPAIGN);
  return all.find(c => Number(c.id) === Number(id)) || null;
}

function debugCampaign() {
  const data = sheetToObjects_(SHEETS.CAMPAIGN);
  Logger.log('Jumlah campaign: ' + data.length);
  data.forEach(c => {
    Logger.log('id: [' + c.id + '] (tipe: ' + typeof c.id + ') — judul: ' + c.judul + ' — status: ' + c.status);
  });
}

function getInfoPembayaran() {
  const cfg = configMap_();
  return {
    qrisUrl: cfg.QrisUrl || '',
    noRekening: cfg.NoRekeningInfaq || '',
    waAdmin: cfg.NoWhatsApp || ''
  };
}

function getInfoZakat() {
  const cfg = configMap_();
  return {
    hargaEmas: Number(cfg.HargaEmasPerGram) || 1500000,
    hargaPerak: Number(cfg.HargaPerakPerGram) || 18000,
    hargaBeras: Number(cfg.HargaBerasPerKg) || 15000
  };
}

function submitDonasi(data) {
  const sheet = getSheet_(SHEETS.DONASI);
  const id = nextId_(SHEETS.DONASI);
  let email = null;
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}

  sheet.appendRow([
    id,
    new Date(),
    data.campaignId,
    data.campaignJudul || '',
    data.nama || 'Hamba Allah',
    email || data.email || '',
    Number(data.jumlah),
    data.keterangan || '',
    'Menunggu Konfirmasi'
  ]);

  return { id: id };
}

// ---------------------------------------------------------------------
// 12. PETUGAS JUMAT
// ---------------------------------------------------------------------
function getSemuaPetugas() {
  return sheetToObjects_(SHEETS.PETUGAS).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
}

// ---------------------------------------------------------------------
// 13. PENGUMUMAN
// ---------------------------------------------------------------------
function getSemuaPengumuman() {
  return sheetToObjects_(SHEETS.PENGUMUMAN).sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
}

// ---------------------------------------------------------------------
// 14. TASBIH DIGITAL — tersimpan per akun Google (UserProperties)
// ---------------------------------------------------------------------
function getTasbihData() {
  const props = PropertiesService.getUserProperties();
  const v = props.getProperty('tasbih_data');
  return v ? JSON.parse(v) : { count: 0, target: 33, dzikir: 'Subhanallah' };
}

function saveTasbihData(count, target, dzikir) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('tasbih_data', JSON.stringify({ count: count, target: target, dzikir: dzikir }));
  return true;
}

// ---------------------------------------------------------------------
// 15. AKUN
// ---------------------------------------------------------------------
function getAkunSaya() {
  let email = null;
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  const riwayatInfaq = email ? getRiwayatInfaqSaya() : [];
  return {
    email: email || null,
    totalInfaq: riwayatInfaq.reduce((s, r) => s + (Number(r.jumlah) || 0), 0),
    jumlahTransaksi: riwayatInfaq.length
  };
}
