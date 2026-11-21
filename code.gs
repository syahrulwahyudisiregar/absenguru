// ===== ABSENSI GURU FACEPRINT - Code.gs (Ver. 007) =====
var TZ = 'Asia/Jakarta';

var HEAD = {
  'Pengaturan': ['Kunci', 'Nilai'],
  'Hak Akses': ['Menu', 'Kepsek', 'Wakil'],
  'Jadwal': ['Hari', 'Pagi Buka', 'Pagi Tutup', 'Batas Terlambat', 'Siang Buka', 'Siang Tutup'],
  'Jadwal Khusus': ['Tanggal', 'Pagi Buka', 'Pagi Tutup', 'Batas Terlambat', 'Siang Buka', 'Siang Tutup', 'Batas Piket'],
  'Guru': ['ID', 'Nama', 'NIP/NUPTK', 'Mapel', 'Aktif', 'Status Wajah', 'Sapaan', 'Libur', 'PIN', 'Piket'],
  'Wajah': ['GuruID', 'Descriptor', 'Waktu'],
  'Absensi': ['No', 'Tanggal', 'GuruID', 'Nama', 'Jam Pagi', 'Status Pagi', 'Jam Siang', 'Keterangan Manual', 'Foto Pagi', 'Foto Siang'],
  'Status Manual': ['No', 'GuruID', 'Nama', 'Status', 'Dari', 'Sampai', 'Keterangan'],
  'Hari Libur': ['Tanggal', 'Keterangan'],
  'Log Admin': ['Waktu', 'Peran', 'Aksi'],
  'Log Gagal': ['Waktu', 'Device', 'Jarak']
};

var MENU = ['dash','guru','wajah','absen','status','rekap','libur','set','hak','pin','tablet','jadwal','log','backup','saya','piket','lokasi','link','reset','hariini'];

var AKSES_DEF = {
  kepsek: { dash:1, absen:1, rekap:1, log:1, reset:2 },
  wakil:  { dash:1, absen:2, status:2, rekap:1, reset:2 }
};

var DEF_CFG = {
  sekolah: 'SMP NEGERI 7 PADANGSIDIMPUAN',
  fotoHari: '7',
  threshold: '0.45',
  adminGuruId: '',
  emailKepsek: '', emailWakil: '', notifMin: '3',
  pin_admin: '123456', pin_super: '999999', pin_kepsek: '111111', pin_wakil: '222222', pin_daftar: '333333'
};

var MAX_GAGAL = 5;      // percobaan PIN salah sebelum terkunci (per mode, semua perangkat)
var KUNCI_DETIK = 300;  // lama terkunci

// ---------- Setup awal (pasang baru, jalankan SEKALI dari editor) ----------
function setupAwal() {
  Object.keys(HEAD).forEach(function (n) { sh(n); });
  var c = cfgAll();
  Object.keys(DEF_CFG).forEach(function (k) {
    if (!c[k] && DEF_CFG[k]) cfgSet(k, DEF_CFG[k]);
  });
  pasangTrigger();
  linkKode();
  logAdmin('sistem', 'Setup awal dijalankan');
  Logger.log('Setup selesai. PIN bawaan: super 999999, admin 123456, kepsek 111111, wakil 222222, daftar 333333. SEGERA GANTI. Isi tab Jadwal.');
}

// ---------- Web app ----------
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Absensi Guru Faceprint')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var fn = req.fn, args = req.args || [];
    var ALLOWED = ['getWaktu','getConfig','getDataWajah','catatAbsen','catatAbsenBatch',
      'getRekapHari','logGagal','cekPerangkat','cekPinDaftar','login','getGuruList',
      'simpanWajah','hapusWajah','getDashboard','getGuruAdmin','simpanGuru','imporGuru',
      'aktifGuru','getAbsenHari','editAbsen','getStatusList','simpanStatus','hapusStatus',
      'getRekap','getLibur','tambahLibur','hapusLibur','getSettings','simpanSettings',
      'daftarTablet','hapusTablet','getLog','getBackups','buatBackup','getMapel','getFoto',
      'tutupAbsen','bukaAbsen','simpanTugasSekolah','simpanIzinPribadi','simpanGuruMassal','getGuruRingkas',
      'loginGuru','getDaftarNamaGuru','getInfoAbsen','keluarSesi','ubahPinGuru','getPinGuru','resetPinGuru',
      'getRekapSaya','getPiketLibur','simpanPiketLibur','getLokasi','simpanLokasi','getLinkGuru','buatUlangLink',
      'simpanJadwalKhusus','hapusJadwalKhusus','getJadwalHariIni','simpanJadwalHariIni','hapusJadwalHariIni',
      'getResetList','getFotoReset','resetAbsen','getResetHariIni'];
    if (ALLOWED.indexOf(fn) < 0) throw new Error('Fungsi tidak dikenal');
    var data = this[fn].apply(null, args);
    return ContentService.createTextOutput(JSON.stringify({ data: data }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: String(err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
// ---------- Helper sheet ----------
function sh(nama) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(nama);
  if (!s) {
    if (nama === 'Pengaturan') return siapkanPengaturan();
    if (nama === 'Hak Akses') return siapkanHakAkses();
    if (nama === 'Jadwal') return siapkanJadwal();
    s = ss.insertSheet(nama);
    var h = HEAD[nama];
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.getRange(1, 1, s.getMaxRows(), h.length).setNumberFormat('@');
    hias(function () { rapikanTab(nama, s); });
  }
  return s;
}

// ---------- Helper waktu (WIB) ----------
function tglStr(d) { return Utilities.formatDate(d || new Date(), TZ, 'yyyy-MM-dd'); }
function jamStr(d) { return Utilities.formatDate(d || new Date(), TZ, 'HH:mm'); }
function menitDari(hhmm) {
  var p = String(hhmm).split(':');
  return parseInt(p[0], 10) * 60 + parseInt(p[1] || '0', 10);
}
function getWaktu() { return new Date().getTime(); }
// 1=Senin ... 7=Minggu untuk tanggal 'yyyy-MM-dd'
function hariKe(tgl) {
  var p = String(tgl).split('-');
  return parseInt(Utilities.formatDate(new Date(Date.UTC(+p[0], +p[1] - 1, +p[2], 12)), TZ, 'u'), 10);
}

function logAdmin(peran, aksi) {
  sh('Log Admin').appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), peran, aksi]);
}

// ---------- Jadwal harian (tab Jadwal, Senin-Sabtu) ----------
var HARI_JDW = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
var JDW_KEYS = ['pagiBuka', 'pagiTutup', 'telat', 'siangBuka', 'siangTutup'];

function siapkanJadwal() {
  var s = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Jadwal');
  s.getRange(1, 1, 1, 6).setValues([HEAD['Jadwal']]).setFontWeight('bold');
  s.getRange(2, 1, 6, 1).setValues(HARI_JDW.map(function (h) { return [h]; }));
  s.getRange(2, 2, 6, 5).setNumberFormat('@');
  pasangRumus(s.getRange(2, 2, 6, 5), '=OR(B2=""§REGEXMATCH(B2§"^([01][0-9]|2[0-3]):[0-5][0-9]$"))',
    'Format HH:MM, contoh 07:30. Kosongkan semua jam = tidak ada absen hari itu.');
  hias(function () { s.getRange(1, 1, 7, 6).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); });
  hias(function () { s.getRange(1, 1, 7, 6).setBorder(true, true, true, true, true, true); });
  s.getRange(8, 1, 1, 2).setNumberFormat('@').setValues([['Batas Terlambat Piket', '07:15']]);
  s.getRange(9, 1).setValue('Isi semua jam tiap hari (format HH:MM). Hari yang dikosongkan = tidak ada absen. Hari Minggu otomatis libur. Baris 8: batas terlambat untuk guru yang piket.');
  s.setFrozenRows(1);
  s.autoResizeColumns(1, 6);
  return s;
}

// {1..6: {pagiBuka,pagiTutup,telat,siangBuka,siangTutup,lengkap,telatPiket}, k: {tanggal: jadwal khusus}}
function jadwalAll() {
  var s = sh('Jadwal'), o = {};
  var v = s.getRange(2, 1, 6, 6).getValues();
  var tp = teksSel(s.getRange(8, 2).getValue()); // batas terlambat piket (baris paling bawah tab Jadwal)
  if (!RE_JAM.test(tp)) tp = '';
  v.forEach(function (r, i) {
    var h = { lengkap: true, telatPiket: tp };
    JDW_KEYS.forEach(function (k, j) {
      var t = teksSel(r[1 + j]);
      if (!RE_JAM.test(t)) { t = ''; h.lengkap = false; }
      h[k] = t;
    });
    o[i + 1] = h;
  });
  o.k = khususAll(tp);
  return o;
}

// jadwal khusus per tanggal (tab Jadwal Khusus): didahulukan dari jadwal mingguan
function khususAll(tp) {
  var o = {};
  baris('Jadwal Khusus').forEach(function (r) {
    var tgl = normTgl(r[0]);
    if (!RE_TGL.test(tgl)) return;
    var h = { lengkap: true };
    JDW_KEYS.forEach(function (k, j) {
      var t = teksSel(r[1 + j]);
      if (!RE_JAM.test(t)) { t = ''; h.lengkap = false; }
      h[k] = t;
    });
    var p = teksSel(r[6]);
    h.telatPiket = RE_JAM.test(p) ? p : (tp || '');
    o[tgl] = h;
  });
  return o;
}

// jadwal lengkap untuk tanggal tgl (khusus dulu, lalu mingguan), atau null (Minggu / belum diisi)
function jamHari(tgl, jd) {
  jd = jd || jadwalAll();
  var kh = jd.k && jd.k[tgl];
  if (kh) return kh.lengkap ? kh : null;
  var h = jd[hariKe(tgl)];
  return (h && h.lengkap) ? h : null;
}

// ---------- Config & perangkat ----------
function getConfig() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('cfg_pub'); // dibagi ke semua perangkat, diperbarui tiap 30 detik / saat pengaturan diubah
  if (hit) return JSON.parse(hit);
  var c = cfgAll(), jd = jadwalAll(), j = {};
  for (var w = 1; w <= 6; w++) {
    var h = jd[w];
    j[w] = h.lengkap ? { pagiBuka: h.pagiBuka, pagiTutup: h.pagiTutup, telat: h.telat, siangBuka: h.siangBuka, siangTutup: h.siangTutup } : null;
  }
  var hi = tglStr(), kh = jd.k[hi];
  if (kh && kh.lengkap) { // jadwal khusus hari ini menggantikan jadwal mingguan untuk hari ini
    j[hariKe(hi)] = { pagiBuka: kh.pagiBuka, pagiTutup: kh.pagiTutup, telat: kh.telat, siangBuka: kh.siangBuka, siangTutup: kh.siangTutup };
  }
  var out = {
    sekolah: c.sekolah,
    jadwal: j,
    threshold: parseFloat(c.threshold) || 0.45,
    libur: baris('Hari Libur').map(function (r) { return normTgl(r[0]); }),
    ucapan: ucapanAll()
  };
  try { cache.put('cfg_pub', JSON.stringify(out), 30); } catch (e) {}
  return out;
}

function cekPerangkat(devId) {
  var c = cfgAll();
  return { ok: !!devId && c.hasOwnProperty('tablet:' + devId) };
}

// ---------- Sesi & hak akses ----------
function buatToken(data, ttl) {
  var t = Utilities.getUuid();
  CacheService.getScriptCache().put('tk_' + t, JSON.stringify(data), ttl || 21600);
  return t;
}

function sesi(token) {
  if (!token) return null;
  var v = CacheService.getScriptCache().get('tk_' + token);
  return v ? JSON.parse(v) : null;
}

var TERLARANG_KW = ['set', 'log', 'backup', 'guru', 'piket', 'lokasi', 'link']; // menu yang tidak pernah boleh dibuka Kepsek & Wakil
function aksesUntuk(mode, c) {
  if (!c && (mode === 'super' || mode === 'daftar' || mode === 'guru')) c = {}; // tidak perlu baca pengaturan
  c = c || cfgAll();
  var a = {};
  MENU.forEach(function (m) {
    if (mode === 'super') { a[m] = 2; return; }
    if (mode === 'admin') { a[m] = (m === 'saya') ? (c.guruAdmin ? 2 : 0) : 2; return; }
    if (mode === 'daftar') { a[m] = (m === 'wajah') ? 2 : 0; return; }
    if (mode === 'guru') { a[m] = (m === 'wajah') ? 2 : 0; return; } // guru hanya boleh daftar wajah dirinya sendiri
    if (m === 'saya') { a[m] = ((mode === 'kepsek' && c.guruKepsek) || (mode === 'wakil' && c.guruWakil)) ? 2 : 0; return; }
    if (m === 'wajah') { a[m] = 0; return; }
    if (TERLARANG_KW.indexOf(m) >= 0) { a[m] = 0; return; } // Kepsek & Wakil tidak boleh: Pengaturan, Log, Backup, Data Guru, Piket&Libur, Lokasi, Link
    var v = c['akses:' + mode + ':' + m];
    a[m] = (v !== undefined) ? parseInt(v, 10) : ((AKSES_DEF[mode] || {})[m] || 0);
  });
  return a;
}

// cek level menu di server; lempar error bila kurang
function butuh(token, menu, lvl) {
  var s = sesi(token);
  if (!s) throw new Error('Sesi habis, silakan login ulang');
  if ((aksesUntuk(s.mode)[menu] || 0) < lvl) throw new Error('Tidak punya akses');
  return s;
}

// lolos bila salah satu menu dalam daftar punya level minimal
function butuhSalahSatu(token, menus, lvl) {
  var s = sesi(token);
  if (!s) throw new Error('Sesi habis, silakan login ulang');
  var a = aksesUntuk(s.mode);
  for (var i = 0; i < menus.length; i++) {
    if ((a[menus[i]] || 0) >= lvl) return s;
  }
  throw new Error('Tidak punya akses');
}

// ---------- Login (kunci percobaan PIN per mode, berlaku untuk semua perangkat) ----------
function cekKunci(kk) {
  return parseInt(CacheService.getScriptCache().get(kk) || '0', 10);
}

function login(devId, mode, pin) {
  if (['super', 'admin', 'kepsek', 'wakil'].indexOf(mode) < 0) return { ok: false, pesan: 'Mode tidak dikenal' };
  var cache = CacheService.getScriptCache();
  var kk = 'gagal_' + mode;
  var n = cekKunci(kk);
  if (n >= MAX_GAGAL) return { ok: false, pesan: 'Terkunci sementara. Coba lagi 5 menit.' };
  var c = cfgAll();
  if (!pin || String(pin) !== String(c['pin_' + mode] || '')) {
    cache.put(kk, String(n + 1), KUNCI_DETIK);
    return { ok: false, pesan: (n + 1 >= MAX_GAGAL) ? 'Terkunci sementara. Coba lagi 5 menit.' : 'PIN salah' };
  }
  cache.remove(kk);
  var guruId = { 'super': c.adminGuruId, admin: c.guruAdmin, kepsek: c.guruKepsek, wakil: c.guruWakil }[mode] || '';
  var gi = guruId ? infoGuru(guruId) : { nama: '', sapaan: '' };
  var token = buatToken({ mode: mode, devId: devId || '', guruId: guruId });
  return { ok: true, token: token, nama: gi.nama, sapaan: gi.sapaan, akses: aksesUntuk(mode, c), guruId: guruId, pesan: '' };
}

function cekPinDaftar(devId, pin) {
  if (!cekPerangkat(devId).ok) return '';
  var cache = CacheService.getScriptCache();
  var kk = 'gagal_daftar';
  var n = cekKunci(kk);
  if (n >= MAX_GAGAL) return '';
  var c = cfgAll();
  if (!pin || String(pin) !== String(c.pin_daftar || '')) {
    cache.put(kk, String(n + 1), KUNCI_DETIK);
    return '';
  }
  cache.remove(kk);
  return buatToken({ mode: 'daftar', devId: devId });
}
// ===== BAGIAN 2: WAJAH & ABSEN =====

// ---------- Helper tambahan ----------
function appendTeks(nama, arr) {
  var s = sh(nama), n = s.getLastRow() + 1;
  var r = s.getRange(n, 1, 1, arr.length);
  r.setNumberFormat('@');
  r.setValues([arr]);
}

function cariBaris(nama, kolom, nilai) {
  var s = sh(nama), n = s.getLastRow();
  if (n < 2) return 0;
  var v = s.getRange(2, kolom, n - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]) === String(nilai)) return i + 2;
  }
  return 0;
}

function pasangSapaan(s) { // kolom G 'Sapaan' + dropdown Bapak/Ibu
  if (String(s.getRange(1, 7).getValue()) !== 'Sapaan') s.getRange(1, 7).setValue('Sapaan').setFontWeight('bold');
  if (s.getMaxRows() < 1000) s.insertRowsAfter(s.getMaxRows(), 1000 - s.getMaxRows());
  if (s.getRange(2, 7).getDataValidation() === null) {
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Bapak', 'Ibu'], true).setAllowInvalid(false).build();
    s.getRange(2, 7, s.getMaxRows() - 1, 1).setDataValidation(rule);
  }
}

// ---------- Libur mingguan guru (kolom H 'Libur': Selasa-Sabtu) ----------
// Saat sesi absen dibuka, guru yang libur hari itu otomatis diisi penanda 'LIBUR' di kolom Jam (dihitung Hadir).
// Bila guru tetap datang dan absen, penanda diganti jam & foto asli.
var PENANDA_LIBUR = 'LIBUR';
var HARI_LIBUR = { 'Selasa': 2, 'Rabu': 3, 'Kamis': 4, 'Jumat': 5, 'Sabtu': 6 };
function normLibur(v) { v = String(v || '').trim(); return HARI_LIBUR.hasOwnProperty(v) ? v : ''; }
function pasangLibur(s) {
  if (String(s.getRange(1, 8).getValue()) !== 'Libur') s.getRange(1, 8).setValue('Libur').setFontWeight('bold');
  if (s.getRange(2, 8).getDataValidation() === null) {
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'], true).setAllowInvalid(false).build();
    s.getRange(2, 8, Math.max(1, s.getMaxRows() - 1), 1).setDataValidation(rule);
  }
}

// Isi penanda LIBUR untuk guru yang libur hari ini, per sesi, begitu sesi dibuka. Murah: dicek paling sering tiap 2 menit.
function isiLiburHariIni() {
  var now = new Date(), tgl = tglStr(now);
  var cache = CacheService.getScriptCache();
  if (cache.get('lbk_' + tgl)) return;
  if (isLibur(tgl, now)) return;
  var J = jamHari(tgl);
  if (!J) return;
  var m = menitDari(jamStr(now)), w = hariKe(tgl);
  var bukaPagi = m >= menitDari(J.pagiBuka), bukaSiang = m >= menitDari(J.siangBuka);
  if (!bukaPagi && !bukaSiang) return;
  var libur = guruAktif().filter(function (g) { return HARI_LIBUR[g.libur] === w; });
  if (libur.length) {
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var sm = statusMap(tgl); // Izin/Sakit/Dinas Luar/Cuti punya prioritas
      libur.forEach(function (g) {
        if (sm[g.id]) return;
        var ada = jamPunya(tgl, g.id);
        if (bukaPagi && !ada.pagi && !ada.liburPagi) {
          tulisAbsen(tgl, g.id, 3, PENANDA_LIBUR);
          tulisAbsen(tgl, g.id, 4, 'libur');
        }
        if (bukaSiang && !ada.siang && !ada.liburSiang) tulisAbsen(tgl, g.id, 5, PENANDA_LIBUR);
      });
    } finally {
      lock.releaseLock();
    }
  }
  cache.put('lbk_' + tgl, '1', 120);
}
function isiLiburAman() { try { isiLiburHariIni(); } catch (e) { Logger.log('Isi libur dilewati: ' + e); } }

// Hapus penanda LIBUR hari ini milik satu guru (dipanggil saat hari libur guru diubah)
function hapusPenandaLibur(tgl, id) {
  if (!cariAbsen(tgl, id)) return;
  var ada = jamPunya(tgl, id);
  if (ada.liburPagi) { tulisAbsen(tgl, id, 3, ''); tulisAbsen(tgl, id, 4, ''); }
  if (ada.liburSiang) tulisAbsen(tgl, id, 5, '');
}

// Rapikan sheet Guru: isi ID kosong & Aktif kosong (mis. data ditempel manual)
function rapikanGuru() {
  var s = sh('Guru'), n = s.getLastRow();
  pasangSapaan(s);
  pasangLibur(s);
  pasangPinPiket(s);
  if (n < 2) return;
  var rg = s.getRange(2, 1, n - 1, 6);
  var v = rg.getValues(), ubah = false, maks = 0;
  v.forEach(function (r) {
    var m = parseInt(String(r[0]).replace(/\D/g, ''), 10) || 0;
    if (m > maks) maks = m;
  });
  v.forEach(function (r) {
    if (!String(r[1]).trim()) return; // baris tanpa nama dilewati
    if (!String(r[0]).trim()) { maks++; r[0] = 'G' + ('000' + maks).slice(-3); ubah = true; }
    var a = String(r[4]).trim().toLowerCase();
    if (a === 'ya' || a === 'true' || a === '') { if (r[4] !== 'Ya') { r[4] = 'Ya'; ubah = true; } }
    else if (r[4] !== 'Tidak') { r[4] = 'Tidak'; ubah = true; }
    if (!String(r[5]).trim()) { r[5] = 'Belum'; ubah = true; }
  });
  if (ubah) { rg.setNumberFormat('@'); rg.setValues(v); }
}

function guruAktif() {
  rapikanGuru(); // [{id,nama,nip,mapel,sapaan}]
  return baris('Guru').filter(function (r) { return String(r[4]) === 'Ya'; })
    .map(function (r) { return { id: String(r[0]), nama: r[1], nip: r[2], mapel: r[3], sapaan: String(r[6] || ''), libur: normLibur(r[7]), piket: normPiket(r[9]) }; });
}

function isLibur(tgl, d) {
  if (Utilities.formatDate(d, TZ, 'u') === '7') return true; // Minggu
  return baris('Hari Libur').some(function (r) { return normTgl(r[0]) === tgl; });
}

function fmtJam(h) { return String(h).replace(':', '.'); }

function pesanLuar(m, J) {
  if (m < menitDari(J.pagiBuka)) return 'Absen pagi dibuka jam ' + fmtJam(J.pagiBuka);
  if (m >= menitDari(J.pagiTutup) && m < menitDari(J.siangBuka)) return 'Absen siang dibuka jam ' + fmtJam(J.siangBuka);
  if (m >= menitDari(J.siangTutup)) return 'Absen tutup';
  return null;
}

// ---------- Data wajah untuk pengenalan ----------
// saja=true: hanya wajah guru yang sedang login (HP guru & Absen Saya).
// Tablet terdaftar serta Super Admin / Admin / Daftar Wajah menerima semua (untuk pencocokan & cek wajah mirip).
function getDataWajah(devId, token, saja) {
  var s = sesi(token);
  if (!cekPerangkat(devId).ok && !s) throw new Error('Perangkat tidak terdaftar');
  var hanya = '';
  if (s && (saja || ['super', 'admin', 'daftar'].indexOf(s.mode) < 0)) hanya = String(s.guruId || '__tidak_ada__');
  var peta = {};
  baris('Wajah').forEach(function (r) {
    try { peta[String(r[0])] = JSON.parse(r[1]); } catch (e) {}
  });
  return guruAktif().filter(function (g) { return peta[g.id] && peta[g.id].length && (!hanya || g.id === hanya); })
    .map(function (g) {
      return { id: g.id, nama: g.nama, nip: g.nip, mapel: g.mapel, sapaan: g.sapaan, desc: peta[g.id] };
    });
}

function getRekapHari() {
  isiLiburAman();
  var aktif = guruAktif(), ids = {};
  aktif.forEach(function (g) { ids[g.id] = 1; });
  var tgl = tglStr(), sudah = 0, lihat = {};
  baris('Absensi').forEach(function (r) {
    var id = String(r[1]);
    if (normTgl(r[0]) === tgl && ids[id] && !lihat[id] && (r[2] !== '' || r[4] !== '')) {
      lihat[id] = 1; sudah++;
    }
  });
  return { sudah: sudah, total: aktif.length };
}

// ---------- Catat absen ----------
// opsi (HP guru / Absen Saya): {manual, lat, lng, acc}
function catatAbsenInti(devId, guruId, sesiAbsen, token, opsi) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    guruId = String(guruId);
    opsi = opsi || {};
    var s = sesi(token);
    if (s && MODE_MANDIRI.indexOf(s.mode) >= 0) {
      // absen sendiri dari HP: hanya ID guru yang tertaut ke akun ini, tanpa kode perangkat, wajib di area sekolah
      if (!s.guruId || s.guruId !== guruId) return { ok: false, pesan: 'Tidak punya akses' };
      var lk = cekLokasi(opsi);
      if (!lk.ok) return lk;
    } else if (!cekPerangkat(devId).ok) {
      return { ok: false, pesan: 'Perangkat tidak terdaftar' };
    }
    if (sesiAbsen !== 'pagi' && sesiAbsen !== 'siang') return { ok: false, pesan: 'Sesi tidak dikenal' };
    if (!guruAktif().some(function (g) { return g.id === guruId; })) return { ok: false, pesan: 'Guru tidak aktif' };

    var now = new Date(), tgl = tglStr(now), jam = jamStr(now), m = menitDari(jam);
    if (isLibur(tgl, now)) return { ok: false, pesan: 'Hari ini libur' };
    var J = jamHari(tgl);
    if (!J) return { ok: false, pesan: 'Jadwal hari ini belum diatur' };

    var buka = menitDari(sesiAbsen === 'pagi' ? J.pagiBuka : J.siangBuka);
    var tutup = menitDari(sesiAbsen === 'pagi' ? J.pagiTutup : J.siangTutup);
    if (m < buka || m >= tutup) {
      return { ok: false, pesan: pesanLuar(m, J) || 'Sesi ini sedang tidak dibuka' };
    }

    var ada = jamPunya(tgl, guruId);
    if (sesiAbsen === 'pagi') {
      if (ada.pagi) {
        // Tugas Sekolah pagi sudah dicatat admin: wajah hanya menambah foto, jam & status tidak berubah
        if (ketJenis(ada.ket).tsPagi && !ada.fotoPagi) return { ok: true, jam: ada.pagi, status: 'tepat', fotoSaja: true };
        return { ok: false, sudah: true, jam: ada.pagi, pesan: 'Sudah absen pagi' };
      }
      // guru yang piket hari ini memakai batas terlambat piket (bila lebih awal dari batas biasa)
      var st = (m > menitDari(batasTelat(J, piketHari(guruId, tgl)))) ? 'terlambat' : 'tepat';
      tulisAbsen(tgl, guruId, 3, jam);
      tulisAbsen(tgl, guruId, 4, st);
      if (opsi.manual) tulisAbsen(tgl, guruId, 6, ketTambah(ada.ket, KET_MNP));
      return { ok: true, jam: jam, status: st };
    }
    if (ada.siang) return { ok: false, sudah: true, jam: ada.siang, pesan: 'Sudah absen siang' };
    tulisAbsen(tgl, guruId, 5, jam);
    if (opsi.manual) tulisAbsen(tgl, guruId, 6, ketTambah(ada.ket, KET_MNS));
    var res = { ok: true, jam: jam, status: 'siang' };
    if (!ada.pagi && !ada.liburPagi) res.tanpaPagi = true;
    return res;
  } finally {
    lock.releaseLock();
  }
}

// sinkron antrean offline: [{id,tgl,jam,sesi,foto}]
function catatAbsenBatchInti(devId, arr, fq) {
  if (!cekPerangkat(devId).ok) throw new Error('Perangkat tidak terdaftar');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var jd = jadwalAll(), baru = 0, lewat = 0, pk = piketMap();
    var aktif = {};
    guruAktif().forEach(function (g) { aktif[g.id] = 1; });
    (arr || []).forEach(function (a) {
      var id = String(a.id), tgl = String(a.tgl || ''), jam = String(a.jam || '').slice(0, 5);
      if (!aktif[id] || !RE_TGL.test(tgl) || !RE_JAM.test(jam)) { lewat++; return; }
      var ada = jamPunya(tgl, id);
      var J = jamHari(tgl, jd);
      if (a.sesi === 'pagi') {
        if (ada.pagi) {
          if (a.foto && ketJenis(ada.ket).tsPagi && !ada.fotoPagi) { fq.push({ tgl: tgl, id: id, sesi: 'pagi', jam: ada.pagi, foto: a.foto }); baru++; }
          else lewat++;
          return;
        }
        tulisAbsen(tgl, id, 3, jam);
        tulisAbsen(tgl, id, 4, (J && menitDari(jam) > menitDari(batasTelat(J, pk[id] === hariKe(tgl)))) ? 'terlambat' : 'tepat');
        if (a.foto) fq.push({ tgl: tgl, id: id, sesi: 'pagi', jam: jam, foto: a.foto });
        baru++;
      } else if (a.sesi === 'siang') {
        if (ada.siang) { lewat++; return; }
        tulisAbsen(tgl, id, 5, jam);
        if (a.foto) fq.push({ tgl: tgl, id: id, sesi: 'siang', jam: jam, foto: a.foto });
        baru++;
      } else { lewat++; }
    });
    return { pesan: baru + ' data disinkron, ' + lewat + ' dilewati' };
  } finally {
    lock.releaseLock();
  }
}

function logGagal(devId, jarak) {
  if (!cekPerangkat(devId).ok) return;
  appendTeks('Log Gagal', [Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    String(devId), String(jarak)]);
}

// ---------- Daftar wajah ----------
function getGuruList(token) {
  var sg = butuh(token, 'wajah', 2);
  rapikanGuru();
  return baris('Guru').filter(function (r) { return String(r[4]) === 'Ya' && (sg.mode !== 'guru' || String(r[0]) === String(sg.guruId)); })
    .map(function (r) {
      return { id: String(r[0]), nama: r[1], statusWajah: r[5] || 'Belum' };
    });
}

function simpanWajah(token, guruId, descArr) {
  var s = butuh(token, 'wajah', 2);
  guruId = String(s.mode === 'guru' ? s.guruId : guruId); // guru hanya boleh mendaftarkan wajahnya sendiri
  var rg = cariBaris('Guru', 1, guruId);
  if (!rg) return { ok: false, pesan: 'Guru tidak ditemukan' };
  if (!descArr || descArr.length < 4 || descArr.length > 10) {
    return { ok: false, pesan: 'Butuh 4-10 sampel wajah' };
  }
  var bersih = [];
  for (var i = 0; i < descArr.length; i++) {
    var d = descArr[i];
    if (!d || d.length !== 128) return { ok: false, pesan: 'Data wajah tidak valid' };
    bersih.push(d.map(function (x) { return Math.round(Number(x) * 10000) / 10000; }));
  }
  var json = JSON.stringify(bersih);
  var waktu = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  var r = cariBaris('Wajah', 1, guruId);
  if (r) {
    var c = sh('Wajah').getRange(r, 2, 1, 2);
    c.setNumberFormat('@');
    c.setValues([[json, waktu]]);
  } else {
    appendTeks('Wajah', [guruId, json, waktu]);
  }
  var cs = sh('Guru').getRange(rg, 6);
  cs.setNumberFormat('@');
  cs.setValue('Terdaftar');
  var nama = sh('Guru').getRange(rg, 2).getValue();
  logAdmin(s.mode, 'Daftar wajah: ' + nama);
  return { ok: true, pesan: 'Wajah ' + nama + ' tersimpan' };
}

function hapusWajah(token, guruId) {
  var s = butuh(token, 'wajah', 2);
  guruId = String(s.mode === 'guru' ? s.guruId : guruId);
  var r = cariBaris('Wajah', 1, guruId);
  if (r) sh('Wajah').deleteRow(r);
  var rg = cariBaris('Guru', 1, guruId);
  if (rg) {
    var cs = sh('Guru').getRange(rg, 6);
    cs.setNumberFormat('@');
    cs.setValue('Belum');
    logAdmin(s.mode, 'Hapus wajah: ' + sh('Guru').getRange(rg, 2).getValue());
  }
  return { ok: true, pesan: 'Data wajah dihapus' };
}
// ===== BAGIAN 3: GURU, DASHBOARD, ABSENSI HARIAN, STATUS, LIBUR =====

var STATUS_OK = ['Izin', 'Sakit', 'Dinas Luar', 'Cuti'];
var RE_TGL = /^\d{4}-\d{2}-\d{2}$/;
var RE_JAM = /^([01]\d|2[0-3]):[0-5]\d$/;
var RE_EMAIL = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

function normTgl(v) { return (v instanceof Date) ? tglStr(v) : String(v); }

function normSapaan(v) {
  v = String(v || '').trim().toLowerCase();
  return v === 'bapak' ? 'Bapak' : (v === 'ibu' ? 'Ibu' : '');
}

function petaNama() {
  var o = {};
  baris('Guru').forEach(function (r) { o[String(r[0])] = r[1]; });
  return o;
}

// guruId -> status manual yang berlaku pada tanggal tgl
function statusMap(tgl) {
  var o = {};
  baris('Status Manual').forEach(function (r) {
    if (normTgl(r[2]) <= tgl && tgl <= normTgl(r[3])) o[String(r[0])] = String(r[1]);
  });
  return o;
}

// ---------- Data guru ----------
function getGuruAdmin(token) {
  butuh(token, 'guru', 1);
  rapikanGuru();
  return baris('Guru').map(function (r) {
    return {
      id: String(r[0]), nama: r[1], nip: String(r[2]), mapel: r[3],
      aktif: String(r[4]) === 'Ya', statusWajah: r[5] || 'Belum', sapaan: String(r[6] || ''), libur: normLibur(r[7])
    };
  });
}

// daftar guru ringkas (id, nama, aktif) untuk menu Edit Absen, Status Manual, dan Rekap
// tidak butuh izin Data Guru; cukup salah satu dari menu-menu itu
function getGuruRingkas(token) {
  butuhSalahSatu(token, ['guru', 'absen', 'status', 'rekap'], 1);
  return baris('Guru').map(function (r) {
    return { id: String(r[0]), nama: r[1], aktif: String(r[4]) === 'Ya' };
  });
}

function idBaru() {
  var maks = 0;
  baris('Guru').forEach(function (r) {
    var n = parseInt(String(r[0]).replace(/\D/g, ''), 10) || 0;
    if (n > maks) maks = n;
  });
  return 'G' + ('000' + (maks + 1)).slice(-3);
}

function nipKembar(nip, kecualiId) {
  if (!nip) return false;
  return baris('Guru').some(function (r) {
    return String(r[2]) === nip && String(r[0]) !== String(kecualiId || '');
  });
}

function simpanGuru(token, g) {
  var s = butuh(token, 'guru', 2);
  var nama = String(g.nama || '').trim();
  var nip = String(g.nip || '').trim();
  var mapel = String(g.mapel || '').trim();
  var sapaan = normSapaan(g.sapaan);
  if (!nama) return { ok: false, pesan: 'Nama wajib diisi' };
  if (nipKembar(nip, g.id)) return { ok: false, pesan: 'NIP/NUPTK sudah dipakai guru lain' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (g.id) {
      var r = cariBaris('Guru', 1, g.id);
      if (!r) return { ok: false, pesan: 'Guru tidak ditemukan' };
      var rg = sh('Guru').getRange(r, 2, 1, 3);
      rg.setNumberFormat('@');
      rg.setValues([[nama, nip, mapel]]);
      var cs = sh('Guru').getRange(r, 7);
      cs.setNumberFormat('@');
      cs.setValue(sapaan);
      hapusCacheGuru();
      logAdmin(s.mode, 'Ubah guru: ' + nama);
      return { ok: true, pesan: 'Data ' + nama + ' diperbarui' };
    }
    appendTeks('Guru', [idBaru(), nama, nip, mapel, 'Ya', 'Belum', sapaan, '', PIN_DEF, '']);
    hapusCacheGuru();
    logAdmin(s.mode, 'Tambah guru: ' + nama);
    return { ok: true, pesan: nama + ' ditambahkan' };
  } finally {
    lock.releaseLock();
  }
}

// Simpan beberapa guru sekaligus (tab Daftar: mapel & sapaan, tab Edit: nama & NIP). Hanya kolom yang dikirim yang diubah.
function simpanGuruMassal(token, arr) {
  var s = butuh(token, 'guru', 2);
  arr = arr || [];
  if (!arr.length) return { ok: true, pesan: 'Tidak ada perubahan' };
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var cur = {};
    baris('Guru').forEach(function (r) {
      cur[String(r[0])] = { nama: String(r[1]), nip: String(r[2]), mapel: String(r[3]), sapaan: String(r[6] || ''), libur: normLibur(r[7]) };
    });
    var baru = {};
    for (var i = 0; i < arr.length; i++) {
      var x = arr[i], id = String(x.id || '');
      if (!cur[id]) return { ok: false, pesan: 'Guru tidak ditemukan' };
      var o = baru[id] || { nama: cur[id].nama, nip: cur[id].nip, mapel: cur[id].mapel, sapaan: cur[id].sapaan, libur: cur[id].libur };
      if (x.nama !== undefined) { o.nama = String(x.nama).trim(); if (!o.nama) return { ok: false, pesan: 'Nama tidak boleh kosong' }; }
      if (x.nip !== undefined) { o.nip = String(x.nip).trim(); if (!o.nip) return { ok: false, pesan: 'NIP/NUPTK tidak boleh kosong: ' + o.nama }; }
      if (x.mapel !== undefined) o.mapel = String(x.mapel).trim();
      if (x.sapaan !== undefined) o.sapaan = normSapaan(x.sapaan);
      if (x.libur !== undefined) o.libur = normLibur(x.libur);
      baru[id] = o;
    }
    var ids = Object.keys(baru);
    for (var a = 0; a < ids.length; a++) { // NIP kembar (keadaan akhir)
      var nip = baru[ids[a]].nip;
      if (!nip) continue;
      var lain = Object.keys(cur);
      for (var b = 0; b < lain.length; b++) {
        if (lain[b] === ids[a]) continue;
        if (((baru[lain[b]] || cur[lain[b]]).nip) === nip) {
          return { ok: false, pesan: 'NIP/NUPTK ' + nip + ' dipakai ' + (baru[lain[b]] || cur[lain[b]]).nama };
        }
      }
    }
    ids.forEach(function (id) {
      var r = cariBaris('Guru', 1, id), o = baru[id];
      var rg = sh('Guru').getRange(r, 2, 1, 3);
      rg.setNumberFormat('@');
      rg.setValues([[o.nama, o.nip, o.mapel]]);
      var cs = sh('Guru').getRange(r, 7);
      cs.setNumberFormat('@');
      cs.setValue(o.sapaan);
      var cl = sh('Guru').getRange(r, 8);
      cl.setNumberFormat('@');
      cl.setValue(o.libur);
    });
    pasangLibur(sh('Guru'));
    var hariIni = tglStr();
    ids.forEach(function (id) { if (baru[id].libur !== cur[id].libur) hapusPenandaLibur(hariIni, id); });
    CacheService.getScriptCache().remove('lbk_' + hariIni); // isi ulang libur hari ini sesuai pengaturan baru
    hapusCacheGuru();
    logAdmin(s.mode, 'Ubah data ' + ids.length + ' guru');
    return { ok: true, pesan: ids.length + ' guru diperbarui' };
  } finally {
    lock.releaseLock();
  }
}

function imporGuru(token, rows) {
  var s = butuh(token, 'guru', 2);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var baru = 0, lewat = 0;
    (rows || []).forEach(function (r) {
      var nama = String(r[0] || '').trim();
      var nip = String(r[1] || '').trim();
      var mapel = String(r[2] || '').trim();
      var sapaan = normSapaan(r[3]);
      if (!nama || nipKembar(nip, '')) { lewat++; return; }
      appendTeks('Guru', [idBaru(), nama, nip, mapel, 'Ya', 'Belum', sapaan, '', PIN_DEF, '']);
      baru++;
    });
    hapusCacheGuru();
    logAdmin(s.mode, 'Impor guru: ' + baru + ' baru, ' + lewat + ' dilewati');
    return { pesan: baru + ' guru ditambahkan, ' + lewat + ' dilewati (kosong/NIP kembar)' };
  } finally {
    lock.releaseLock();
  }
}

function aktifGuru(token, id) {
  var s = butuh(token, 'guru', 2);
  var r = cariBaris('Guru', 1, id);
  if (!r) return { ok: false, pesan: 'Guru tidak ditemukan' };
  var c = sh('Guru').getRange(r, 5);
  var baruNilai = (String(c.getValue()) === 'Ya') ? 'Tidak' : 'Ya';
  c.setNumberFormat('@');
  c.setValue(baruNilai);
  hapusCacheGuru();
  logAdmin(s.mode, (baruNilai === 'Ya' ? 'Aktifkan' : 'Nonaktifkan') + ' guru: ' + sh('Guru').getRange(r, 2).getValue());
  return { ok: true, pesan: baruNilai === 'Ya' ? 'Guru diaktifkan' : 'Guru dinonaktifkan' };
}

// ---------- Menu Absen (dulu Dashboard) ----------
function getDashboard(token, tgl) {
  var sd = butuh(token, 'dash', 1);
  isiLiburAman();
  tgl = String(tgl || tglStr());
  if (!RE_TGL.test(tgl)) throw new Error('Format tanggal salah');
  var lihat = sd.mode !== 'wakil'; // Wakil tidak melihat foto
  function bendera(v) { v = String(v || ''); return (!lihat || !v) ? '' : (v === 'kedaluwarsa' ? 'exp' : 'ada'); }
  var sm = statusMap(tgl), abs = {};
  baris('Absensi').forEach(function (r) {
    if (normTgl(r[0]) === tgl) abs[String(r[1])] = r;
  });
  return guruAktif().map(function (g) {
    var a = abs[g.id] || [];
    var kj = ketJenis(a[5]);
    return { id: g.id, nama: g.nama, pagi: String(a[2] || ''), stPagi: String(a[3] || ''),
             siang: String(a[4] || ''), ket: sm[g.id] || '', fp: bendera(a[6]), fs: bendera(a[7]),
             tsP: kj.tsPagi, tsS: kj.tsSiang, it: kj.izinTerlambat, pc: kj.pulangCepat, mp: kj.manualPagi, ms: kj.manualSiang,
             kp: kj.alasanPagi, ks: kj.alasanSiang };
  }).sort(function (x, y) { return String(x.nama).localeCompare(String(y.nama), 'id'); });
}

// ---------- Absensi harian ----------
function getAbsenHari(token, tgl) {
  butuh(token, 'absen', 1);
  isiLiburAman();
  tgl = String(tgl || tglStr());
  if (!RE_TGL.test(tgl)) throw new Error('Format tanggal salah');
  var sm = statusMap(tgl), abs = {};
  baris('Absensi').forEach(function (r) {
    if (normTgl(r[0]) === tgl) abs[String(r[1])] = r;
  });
  return guruAktif().map(function (g) {
    var a = abs[g.id] || [];
    return {
      id: g.id, nama: g.nama,
      pagi: String(a[2] || ''), stPagi: String(a[3] || ''),
      siang: String(a[4] || ''),
      manual: sm[g.id] || String(a[5] || '')
    };
  });
}

function editAbsen(token, tgl, guruId, sesiAbsen, jam) {
  var s = butuh(token, 'absen', 2);
  tgl = String(tgl); guruId = String(guruId); jam = String(jam || '').trim();
  if (!RE_TGL.test(tgl)) return { ok: false, pesan: 'Format tanggal salah' };
  if (jam && !RE_JAM.test(jam)) return { ok: false, pesan: 'Format jam harus HH:MM' };
  if (tgl !== tglStr() && s.mode !== 'super') return { ok: false, pesan: 'Hanya Super Admin yang bisa mengedit tanggal lain' };
  if (!cariBaris('Guru', 1, guruId)) return { ok: false, pesan: 'Guru tidak ditemukan' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (sesiAbsen === 'pagi') {
      var J = jamHari(tgl);
      var st = jam ? ((J && menitDari(jam) > menitDari(batasTelat(J, piketHari(guruId, tgl)))) ? 'terlambat' : 'tepat') : '';
      if (jam && ketJenis(jamPunya(tgl, guruId).ket).tsPagi) st = 'tepat'; // Tugas Sekolah tidak pernah terlambat
      if (!jam && !cariAbsen(tgl, guruId)) return { ok: true, pesan: 'Tidak ada data' };
      tulisAbsen(tgl, guruId, 3, jam);
      tulisAbsen(tgl, guruId, 4, st);
      if (!jam) bersihKetSesi(tgl, guruId, 'pagi');
    } else if (sesiAbsen === 'siang') {
      if (!jam && !cariAbsen(tgl, guruId)) return { ok: true, pesan: 'Tidak ada data' };
      tulisAbsen(tgl, guruId, 5, jam);
      if (!jam) bersihKetSesi(tgl, guruId, 'siang');
    } else {
      return { ok: false, pesan: 'Sesi tidak dikenal' };
    }
    var nama = petaNama()[guruId] || guruId;
    logAdmin(s.mode, 'Edit absen ' + tgl + ' ' + nama + ' ' + sesiAbsen + ': ' + (jam || 'dihapus'));
    return { ok: true, pesan: jam ? 'Absen disimpan' : 'Absen dihapus' };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Tugas Sekolah & Izin Pribadi (hanya hari ini, jam otomatis = waktu server) ----------
function bersihKet(v) {
  return String(v || '').replace(/[\r\n|]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 100);
}

// jenis: tsPagi | tsSiang | izinTerlambat | pulangCepat
function catatManual(s, d, jenis) {
  var id = String(d.id || ''), ket = bersihKet(d.ket);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!guruAktif().some(function (g) { return g.id === id; })) return { ok: false, pesan: 'Pilih guru aktif dulu' };
    var now = new Date(), tgl = tglStr(now), jam = jamStr(now), m = menitDari(jam);
    if (isLibur(tgl, now)) return { ok: false, pesan: 'Hari ini libur' };
    var J = jamHari(tgl);
    if (!J) return { ok: false, pesan: 'Jadwal hari ini belum diatur' };
    var stm = statusMap(tgl)[id];
    if (stm) return { ok: false, pesan: 'Guru ini berstatus ' + stm + ' hari ini' };
    var pagi = (jenis === 'tsPagi' || jenis === 'izinTerlambat');
    if (jenis === 'izinTerlambat' && m < menitDari(J.pagiTutup)) {
      return { ok: false, pesan: 'Absen pagi masih dibuka. Guru bisa absen wajah di tablet.' };
    }
    var ada = jamPunya(tgl, id);
    if (pagi && ada.pagi) return { ok: false, pesan: 'Sudah ada data absen pagi (jam ' + ada.pagi + ')' };
    if (!pagi && ada.siang) return { ok: false, pesan: 'Sudah ada data absen siang (jam ' + ada.siang + ')' };
    var label = { tsPagi: KET_TSP, tsSiang: KET_TSS, izinTerlambat: KET_IT, pulangCepat: KET_PC }[jenis];
    if (pagi) {
      tulisAbsen(tgl, id, 3, jam);
      tulisAbsen(tgl, id, 4, jenis === 'tsPagi' ? 'tepat' : 'terlambat'); // Tugas Sekolah tidak pernah terlambat
    } else {
      tulisAbsen(tgl, id, 5, jam);
    }
    tulisAbsen(tgl, id, 6, ketTambah(ada.ket, label + (ket ? ': ' + ket : '')));
    logAdmin(s.mode, label + ': ' + (petaNama()[id] || id) + (ket ? ' - ' + ket : ''));
    return { ok: true, pesan: label + ' dicatat, jam ' + jam };
  } finally {
    lock.releaseLock();
  }
}

function simpanTugasSekolah(token, d) { // d: {id, sesi: 'pagi'|'siang', ket}
  var s = butuh(token, 'status', 2);
  if (d.sesi !== 'pagi' && d.sesi !== 'siang') return { ok: false, pesan: 'Pilih sesi Pagi atau Siang' };
  return catatManual(s, d, d.sesi === 'pagi' ? 'tsPagi' : 'tsSiang');
}

function simpanIzinPribadi(token, d) { // d: {id, jenis: 'terlambat'|'pulang', ket}
  var s = butuh(token, 'status', 2);
  if (d.jenis !== 'terlambat' && d.jenis !== 'pulang') return { ok: false, pesan: 'Pilih jenis izin' };
  return catatManual(s, d, d.jenis === 'terlambat' ? 'izinTerlambat' : 'pulangCepat');
}

// ---------- Status manual ----------
function getStatusList(token) {
  butuh(token, 'status', 1);
  var nama = petaNama(), out = [];
  baris('Status Manual').forEach(function (r, i) {
    out.push({
      row: i + 2, id: String(r[0]), nama: nama[String(r[0])] || String(r[0]),
      status: String(r[1]), dari: normTgl(r[2]), sampai: normTgl(r[3]), ket: String(r[4] || '')
    });
  });
  return out;
}

function simpanStatus(token, d) {
  var s = butuh(token, 'status', 2);
  var id = String(d.id || ''), st = String(d.status || '');
  var dari = String(d.dari || ''), sampai = String(d.sampai || d.dari || '');
  if (!cariBaris('Guru', 1, id)) return { ok: false, pesan: 'Pilih guru dulu' };
  if (STATUS_OK.indexOf(st) < 0) return { ok: false, pesan: 'Status tidak valid' };
  if (!RE_TGL.test(dari) || !RE_TGL.test(sampai)) return { ok: false, pesan: 'Tanggal belum benar' };
  if (sampai < dari) return { ok: false, pesan: 'Tanggal sampai lebih awal dari dari' };
  appendTeks('Status Manual', ['', id, petaNama()[id] || '', st, dari, sampai, String(d.ket || '')]);
  nomori('Status Manual');
  logAdmin(s.mode, 'Status ' + st + ': ' + (petaNama()[id] || id) + ' ' + dari + ' s/d ' + sampai);
  return { ok: true, pesan: 'Status disimpan' };
}

function hapusStatus(token, row) {
  var s = butuh(token, 'status', 2);
  row = parseInt(row, 10);
  var sheet = sh('Status Manual');
  if (!row || row < 2 || row > sheet.getLastRow()) return { ok: false, pesan: 'Data tidak ditemukan' };
  var v = sheet.getRange(row, 2, 1, 3).getValues()[0];
  sheet.deleteRow(row);
  nomori('Status Manual');
  logAdmin(s.mode, 'Hapus status ' + v[2] + ': ' + (petaNama()[String(v[0])] || v[0]));
  return { ok: true, pesan: 'Status dihapus' };
}

// ---------- Hari libur ----------
function cariLibur(tgl) {
  var s = sh('Hari Libur'), n = s.getLastRow();
  if (n < 2) return 0;
  var v = s.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (normTgl(v[i][0]) === tgl) return i + 2;
  }
  return 0;
}

function getLibur(token) {
  butuh(token, 'libur', 1);
  return baris('Hari Libur').map(function (r) {
    return { tgl: normTgl(r[0]), ket: String(r[1] || '') };
  }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; });
}

function tambahLibur(token, tgl, ket) {
  var s = butuh(token, 'libur', 2);
  tgl = String(tgl || '');
  if (!RE_TGL.test(tgl)) return { ok: false, pesan: 'Tanggal belum benar' };
  if (cariLibur(tgl)) return { ok: false, pesan: 'Tanggal sudah ada' };
  appendTeks('Hari Libur', [tgl, String(ket || '')]);
  hapusCacheCfg();
  logAdmin(s.mode, 'Tambah libur: ' + tgl + ' ' + (ket || ''));
  return { ok: true, pesan: 'Hari libur ditambahkan' };
}

function hapusLibur(token, tgl) {
  var s = butuh(token, 'libur', 2);
  var r = cariLibur(String(tgl));
  if (r) sh('Hari Libur').deleteRow(r);
  hapusCacheCfg();
  logAdmin(s.mode, 'Hapus libur: ' + tgl);
  return { ok: true, pesan: 'Hari libur dihapus' };
}

// Tutup / buka kembali absen satu tanggal (libur mendadak) dari menu Pengaturan
function tutupAbsen(token, tgl, ket) {
  var s = butuh(token, 'set', 2);
  tgl = String(tgl || '');
  if (!RE_TGL.test(tgl)) return { ok: false, pesan: 'Tanggal belum benar' };
  if (tgl < tglStr()) return { ok: false, pesan: 'Tanggal sudah lewat' };
  if (cariLibur(tgl)) return { ok: false, pesan: 'Absen tanggal itu sudah ditutup' };
  ket = String(ket || '').trim() || 'Libur mendadak';
  appendTeks('Hari Libur', [tgl, ket]);
  hapusCacheCfg();
  logAdmin(s.mode, 'Tutup absen: ' + tgl + ' ' + ket);
  return { ok: true, pesan: 'Absen ' + tgl + ' ditutup' };
}

function bukaAbsen(token, tgl) {
  var s = butuh(token, 'set', 2);
  var r = cariLibur(String(tgl));
  if (!r) return { ok: false, pesan: 'Tanggal tidak ada di daftar libur' };
  sh('Hari Libur').deleteRow(r);
  hapusCacheCfg();
  logAdmin(s.mode, 'Buka absen: ' + tgl);
  return { ok: true, pesan: 'Absen ' + tgl + ' dibuka kembali' };
}
// ===== BAGIAN 4: REKAP, PENGATURAN, TABLET, LOG, BACKUP =====

var RE_BLN = /^\d{4}-\d{2}$/;
var CFG_KEYS = ['sekolah', 'threshold', 'adminGuruId', 'guruAdmin', 'guruKepsek', 'guruWakil', 'fotoHari', 'emailKepsek', 'emailWakil', 'notifMin'];

// ---------- Rekap bulanan (guru nonaktif tetap muncul bila punya data bulan itu) ----------
function getRekap(token, bulan, guruId) {
  butuh(token, 'rekap', 1);
  return rekapInti(bulan, guruId);
}

function rekapInti(bulan, guruId) {
  isiLiburAman();
  bulan = String(bulan || tglStr().slice(0, 7));
  if (!RE_BLN.test(bulan)) throw new Error('Format bulan salah');
  var y = parseInt(bulan.slice(0, 4), 10), mo = parseInt(bulan.slice(5), 10);
  var jmlHari = new Date(y, mo, 0).getDate();
  var now = new Date(), hariIni = tglStr(now), m = menitDari(jamStr(now));
  var jd = jadwalAll();
  var awal = bulan + '-01', akhir = bulan + '-' + ('0' + jmlHari).slice(-2);

  var libur = {};
  baris('Hari Libur').forEach(function (r) { libur[normTgl(r[0])] = 1; });
  var abs = {}, punya = {};
  baris('Absensi').forEach(function (r) {
    var t = normTgl(r[0]);
    if (t.slice(0, 7) === bulan) {
      abs[t + '|' + r[1]] = r;
      if (r[2] !== '' || r[4] !== '') punya[String(r[1])] = 1;
    }
  });
  var sts = baris('Status Manual').map(function (r) {
    return { id: String(r[0]), st: String(r[1]), dari: normTgl(r[2]), sampai: normTgl(r[3]) };
  });
  sts.forEach(function (x) { if (x.dari <= akhir && x.sampai >= awal) punya[x.id] = 1; });

  rapikanGuru();
  var daftar = baris('Guru').filter(function (r) { return String(r[1]).trim(); }).map(function (r) {
    return { id: String(r[0]), nama: String(r[1]), aktif: String(r[4]) === 'Ya' };
  }).filter(function (g) { return g.aktif || punya[g.id]; });
  if (guruId) daftar = daftar.filter(function (g) { return g.id === String(guruId); });

  var rows = daftar.map(function (g) {
    var h = 0, t = 0, it = 0, pc = 0, a = 0, ket = { 'Izin': 0, 'Sakit': 0, 'Dinas Luar': 0, 'Cuti': 0 };
    for (var d = 1; d <= jmlHari; d++) {
      var tgl = bulan + '-' + ('0' + d).slice(-2);
      if (tgl > hariIni) break;
      if (libur[tgl]) continue;
      var w = hariKe(tgl);
      if (w === 7) continue;
      var r = abs[tgl + '|' + g.id];
      if (r && (r[2] !== '' || r[4] !== '')) {
        var kj = ketJenis(r[5]);
        if (String(r[3]) === 'terlambat') { if (kj.izinTerlambat) it++; else t++; } else h++; // Izin Terlambat tidak dihitung ganda
        if (kj.pulangCepat) pc++;
        continue;
      }
      var J = jamHari(tgl, jd); // jadwal khusus tanggal didahulukan
      if (!J) continue; // hari tanpa jadwal bukan hari absen
      var st = '';
      for (var i = 0; i < sts.length; i++) {
        if (sts[i].id === g.id && sts[i].dari <= tgl && tgl <= sts[i].sampai) { st = sts[i].st; break; }
      }
      if (st && ket.hasOwnProperty(st)) { ket[st]++; continue; }
      if (tgl === hariIni && m < menitDari(J.siangTutup)) continue; // hari ini belum ditutup
      if (g.aktif) a++; // guru nonaktif: alpa tidak dihitung
    }
    return [g.nama + (g.aktif ? '' : ' (nonaktif)'), h, t, it, pc, a, ket['Izin'], ket['Sakit'], ket['Dinas Luar'], ket['Cuti']];
  });
  return { head: ['Nama', 'Hadir', 'Terlambat', 'Izin Terlambat', 'Pulang Cepat', 'Alpa', 'Izin', 'Sakit', 'Dinas', 'Cuti'], rows: rows };
}

// ---------- Pengaturan ----------
function getSettings(token) {
  var s = butuhSalahSatu(token, ['set', 'hak', 'pin', 'tablet', 'jadwal'], 1);
  var c = cfgAll(), jd = jadwalAll();
  var lihatPin = (s.mode === 'super' || s.mode === 'admin'); // PIN hanya terlihat oleh super & admin
  var cfg = {};
  CFG_KEYS.forEach(function (k) { cfg[k] = c[k] !== undefined ? c[k] : ''; });
  var akses = { kepsek: {}, wakil: {} };
  ['kepsek', 'wakil'].forEach(function (mode) {
    var a = aksesUntuk(mode, c);
    MENU.forEach(function (mn) {
      if (mn !== 'wajah' && mn !== 'saya') akses[mode][mn] = a[mn];
    });
  });
  var tablet = [];
  Object.keys(c).forEach(function (k) {
    if (k.indexOf('tablet:') === 0) tablet.push({ id: k.slice(7), nama: c[k] });
  });
  var jadwal = {}, kurang = [];
  for (var w = 1; w <= 6; w++) {
    var o = {};
    for (var q = 0; q < JDW_KEYS.length; q++) o[JDW_KEYS[q]] = jd[w][JDW_KEYS[q]];
    jadwal[w] = o;
    if (!jd[w].lengkap) kurang.push(HARI_JDW[w - 1]);
  }
  var NAMA_MODE = { 'super': 'Super Admin', admin: 'Admin', kepsek: 'Kepsek', wakil: 'Wakil Kepsek', daftar: 'Daftar Wajah' };
  var bawaan = Object.keys(NAMA_MODE).filter(function (m) { return c['pin_' + m] === DEF_CFG['pin_' + m]; })
    .map(function (m) { return NAMA_MODE[m]; });
  var hariIni = tglStr();
  var liburAkan = baris('Hari Libur').map(function (r) {
    return { tgl: normTgl(r[0]), ket: String(r[1] || '') };
  }).filter(function (x) { return x.tgl >= hariIni; }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; });
  return {
    cfg: cfg,
    pin: { 'super': (s.mode === 'super') ? (c.pin_super || '') : undefined,
      admin: lihatPin ? (c.pin_admin || '') : '', kepsek: lihatPin ? (c.pin_kepsek || '') : '',
      wakil: lihatPin ? (c.pin_wakil || '') : '', daftar: lihatPin ? (c.pin_daftar || '') : '' },
    akses: akses,
    mapel: daftarMapel(),
    tablet: tablet,
    jadwal: jadwal,
    jadwalKurang: kurang,
    pinBawaan: bawaan,
    ucapan: ucapanAll(),
    liburAkan: liburAkan,
    telatPiket: jd[1].telatPiket || '',
    khusus: khususList(jd)
  };
}

function simpanSettings(token, d) {
  var s = sesi(token);
  if (!s) throw new Error('Sesi habis, silakan login ulang');
  var akses0 = aksesUntuk(s.mode);
  d = d || {};
  var cfg = d.cfg || {}, pin = d.pin || {}, akses = d.akses || {}, jadwal = d.jadwal, ucapan = d.ucapan, telatPiket = d.telatPiket;

  // tiap bagian dicek dengan menunya sendiri
  if ((Object.keys(cfg).length || ucapan) && (akses0.set || 0) < 2) throw new Error('Tidak punya akses');
  if ((jadwal || telatPiket !== undefined) && (akses0.jadwal || 0) < 2) throw new Error('Tidak punya akses');
  if (Object.keys(pin).length && (akses0.pin || 0) < 2) throw new Error('Tidak punya akses');
  if (Object.keys(akses).length && (akses0.hak || 0) < 2) throw new Error('Tidak punya akses');

  // validasi
  if (cfg.threshold !== undefined) {
    var th = parseFloat(cfg.threshold);
    if (isNaN(th) || th < 0.3 || th > 0.7) return { ok: false, pesan: 'Threshold harus 0.3 - 0.7' };
    cfg.threshold = String(th);
  }
  if (cfg.fotoHari !== undefined && !/^([1-9]|[12][0-9]|30)$/.test(String(cfg.fotoHari))) return { ok: false, pesan: 'Lama simpan foto harus 1-30 hari' };
  if (cfg.adminGuruId && !cariBaris('Guru', 1, cfg.adminGuruId)) return { ok: false, pesan: 'ID guru Super Admin tidak ditemukan' };
  var idGuruKeys = ['guruAdmin', 'guruKepsek', 'guruWakil'];
  for (var gi = 0; gi < idGuruKeys.length; gi++) {
    var kg = idGuruKeys[gi];
    if (cfg[kg] && !cariBaris('Guru', 1, cfg[kg])) return { ok: false, pesan: 'ID guru ' + kg + ' tidak ditemukan' };
  }
  ['emailKepsek', 'emailWakil'].forEach(function (k) { if (cfg[k] !== undefined) cfg[k] = String(cfg[k]).trim(); });
  if (cfg.emailKepsek && !RE_EMAIL.test(cfg.emailKepsek)) return { ok: false, pesan: 'Email Kepsek tidak valid' };
  if (cfg.emailWakil && !RE_EMAIL.test(cfg.emailWakil)) return { ok: false, pesan: 'Email Wakil tidak valid' };
  if (cfg.notifMin !== undefined && !/^([1-9]|[1-9][0-9])$/.test(String(cfg.notifMin))) return { ok: false, pesan: 'Batas notifikasi harus 1-99 guru' };
  var modePin = ['admin', 'kepsek', 'wakil', 'daftar'];
  if (s.mode === 'super') modePin.unshift('super');
  for (var j = 0; j < modePin.length; j++) {
    var p = pin[modePin[j]];
    if (p === undefined || String(p) === '') continue; // kosong = tidak diubah
    if (!/^\d{4,6}$/.test(String(p))) return { ok: false, pesan: 'PIN ' + modePin[j] + ' harus 4-6 angka' };
  }
  var jadwalOk = {};
  for (var w = 1; w <= 6; w++) {
    var hj = jadwal && jadwal[w];
    if (!hj) continue;
    var vals = [], isi = 0;
    for (var q = 0; q < JDW_KEYS.length; q++) {
      var tj = String(hj[JDW_KEYS[q]] || '').trim();
      if (tj && !RE_JAM.test(tj)) return { ok: false, pesan: 'Format jam salah: ' + HARI_JDW[w - 1] };
      if (tj) isi++;
      vals.push(tj);
    }
    if (isi > 0 && isi < JDW_KEYS.length) return { ok: false, pesan: 'Jadwal ' + HARI_JDW[w - 1] + ': isi semua jam atau kosongkan semua' };
    if (isi === JDW_KEYS.length) {
      var m1 = menitDari(vals[0]), m2 = menitDari(vals[1]), mt = menitDari(vals[2]), m3 = menitDari(vals[3]), m4 = menitDari(vals[4]);
      if (!(m1 < m2 && m2 <= m3 && m3 < m4 && mt >= m1 && mt <= m2)) {
        return { ok: false, pesan: 'Urutan jam ' + HARI_JDW[w - 1] + ' tidak masuk akal (pagi buka < batas terlambat <= pagi tutup <= siang buka < siang tutup)' };
      }
    }
    jadwalOk[w] = vals;
  }
  var tpOk = null;
  if (telatPiket !== undefined) {
    tpOk = String(telatPiket || '').trim();
    if (tpOk && !RE_JAM.test(tpOk)) return { ok: false, pesan: 'Format batas terlambat piket salah (HH:MM)' };
    if (tpOk) {
      var cj = jadwalAll();
      for (var w3 = 1; w3 <= 6; w3++) {
        var pb = jadwalOk[w3] ? jadwalOk[w3][0] : (cj[w3].lengkap ? cj[w3].pagiBuka : '');
        if (pb && menitDari(tpOk) < menitDari(pb)) return { ok: false, pesan: 'Batas terlambat piket lebih awal dari jam pagi buka hari ' + HARI_JDW[w3 - 1] };
      }
    }
  }
  var uc = null;
  if (ucapan) {
    uc = { pagi: [], siang: [] };
    ['pagi', 'siang'].forEach(function (k) {
      (ucapan[k] || []).forEach(function (t) {
        t = String(t).trim().slice(0, 250);
        if (t) uc[k].push(t);
      });
    });
    if (uc.pagi.length + uc.siang.length > UC_N) return { ok: false, pesan: 'Maksimal ' + UC_N + ' kalimat ucapan' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ubah = [];
    CFG_KEYS.forEach(function (k) {
      var boleh = (k === 'adminGuruId') ? (s.mode === 'super')
        : (k === 'guruAdmin' || k === 'guruKepsek' || k === 'guruWakil') ? (s.mode === 'admin' || s.mode === 'super') : true;
      if (cfg[k] !== undefined && boleh) { cfgSet(k, cfg[k]); ubah.push(k); }
    });
    modePin.forEach(function (mp) {
      if (pin[mp] !== undefined && String(pin[mp]) !== '') {
        cfgSet('pin_' + mp, String(pin[mp]));
        ubah.push('pin_' + mp);
      }
    });
    ['kepsek', 'wakil'].forEach(function (mode) {
      var a = akses[mode];
      if (!a) return;
      MENU.forEach(function (mn) {
        if (mn === 'wajah' || mn === 'saya' || TERLARANG_KW.indexOf(mn) >= 0) return;
        var v = parseInt(a[mn], 10);
        if (a[mn] !== undefined && v >= 0 && v <= 2) cfgSet('akses:' + mode + ':' + mn, v);
      });
      ubah.push('akses_' + mode);
    });
    var sj = sh('Jadwal');
    Object.keys(jadwalOk).forEach(function (w2) {
      var rg = sj.getRange(parseInt(w2, 10) + 1, 2, 1, 5);
      rg.setNumberFormat('@');
      rg.setValues([jadwalOk[w2]]);
    });
    if (Object.keys(jadwalOk).length) ubah.push('jadwal');
    if (tpOk !== null) {
      sj.getRange(8, 1).setValue('Batas Terlambat Piket');
      var cb = sj.getRange(8, 2);
      cb.setNumberFormat('@');
      cb.setValue(tpOk);
      ubah.push('telatPiket');
    }
    if (uc) { simpanUcapan(uc); ubah.push('ucapan'); }
    hapusCacheCfg();
    CacheService.getScriptCache().remove('lbk_' + tglStr());
    logAdmin(s.mode, 'Ubah pengaturan: ' + ubah.join(', ')); // nilai PIN tidak dicatat
    return { ok: true, pesan: 'Pengaturan disimpan' };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Tablet terdaftar (kode 6 angka) ----------
function daftarTablet(token, id, nama) {
  var s = butuh(token, 'tablet', 2);
  id = String(id || '').trim(); nama = String(nama || '').trim();
  if (!/^\d{6}$/.test(id)) return { ok: false, pesan: 'Kode harus 6 angka (sesuai yang tampil di layar tablet)' };
  cfgSet('tablet:' + id, nama || 'Tablet');
  logAdmin(s.mode, 'Daftar tablet: ' + (nama || id));
  return { ok: true, pesan: 'Tablet didaftarkan' };
}

function hapusTablet(token, id) {
  var s = butuh(token, 'tablet', 2);
  cfgDel('tablet:' + id);
  logAdmin(s.mode, 'Hapus tablet: ' + id);
  return { ok: true, pesan: 'Tablet dihapus' };
}

// ---------- Log ----------
function getLog(token, jenis) {
  var sl = butuh(token, 'log', 1);
  var nama = (jenis === 'gagal') ? 'Log Gagal' : 'Log Admin';
  var s = sh(nama), n = s.getLastRow();
  var kol = HEAD[nama].length;
  if (n < 2) return { head: HEAD[nama], rows: [] };
  var mulai = Math.max(2, n - 399);
  var v = s.getRange(mulai, 1, n - mulai + 1, kol).getValues().map(function (r) {
    return r.map(function (x) { return String(x); });
  });
  if (jenis !== 'gagal' && sl.mode !== 'super') {
    v = v.filter(function (r) { return r[1] !== 'super'; });
  }
  v = v.slice(-200);
  return { head: HEAD[nama], rows: v.reverse() }; // terbaru di atas
}

// ---------- Backup ----------
function folderBackup() {
  var it = DriveApp.getFoldersByName('Backup Absensi Faceprint');
  return it.hasNext() ? it.next() : DriveApp.createFolder('Backup Absensi Faceprint');
}

function getBackups(token) {
  butuh(token, 'backup', 1);
  var out = [], fs = folderBackup().getFiles();
  while (fs.hasNext()) {
    var f = fs.next();
    out.push({ t: f.getDateCreated().getTime(), waktu: Utilities.formatDate(f.getDateCreated(), TZ, 'yyyy-MM-dd HH:mm'), nama: f.getName() });
  }
  out.sort(function (a, b) { return b.t - a.t; });
  return out.slice(0, 50).map(function (o) { return { waktu: o.waktu, nama: o.nama }; });
}

function buatBackup(token) {
  var s = butuh(token, 'backup', 2);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nama = 'Backup Absensi ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH-mm');
  DriveApp.getFileById(ss.getId()).makeCopy(nama, folderBackup());
  logAdmin(s.mode, 'Buat backup: ' + nama);
  return { pesan: 'Backup dibuat: ' + nama };
}

// Trigger mingguan (Minggu jam 02): backup otomatis, simpan 8 terbaru
function backupOtomatis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nama = 'Backup Otomatis ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var folder = folderBackup();
  DriveApp.getFileById(ss.getId()).makeCopy(nama, folder);
  var daftar = [], it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName().indexOf('Backup Otomatis ') === 0) daftar.push(f);
  }
  daftar.sort(function (a, b) { return b.getDateCreated().getTime() - a.getDateCreated().getTime(); });
  for (var i = 8; i < daftar.length; i++) daftar[i].setTrashed(true);
  logAdmin('sistem', 'Backup otomatis: ' + nama);
}

// ---------- Notifikasi email ke Kepsek & Wakil ----------
// Trigger tiap jam: setelah jam tutup siang, kirim SEKALI per hari bila guru yang belum absen >= batas
function cekNotifikasi() {
  var now = new Date(), tgl = tglStr(now), m = menitDari(jamStr(now));
  if (isLibur(tgl, now)) return;
  var J = jamHari(tgl);
  if (!J || m < menitDari(J.siangTutup)) return;
  isiLiburAman();
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('notifTerakhir') === tgl) return;
  props.setProperty('notifTerakhir', tgl);
  var c = cfgAll(), tujuan = [];
  [c.emailKepsek, c.emailWakil].forEach(function (e) { if (e && RE_EMAIL.test(e)) tujuan.push(e); });
  if (!tujuan.length) return;
  var sm = statusMap(tgl), abs = {};
  baris('Absensi').forEach(function (r) {
    if (normTgl(r[0]) === tgl && (r[2] !== '' || r[4] !== '')) abs[String(r[1])] = 1;
  });
  var belum = guruAktif().filter(function (g) { return !abs[g.id] && !sm[g.id]; })
    .map(function (g) { return String(g.nama); }).sort(function (a, b) { return a.localeCompare(b, 'id'); });
  var min = parseInt(c.notifMin, 10) || 3;
  if (belum.length < min) return;
  var body = 'Daftar guru yang belum absen pada ' + tgl + ' (' + belum.length + ' guru):\n\n' +
    belum.map(function (n, i) { return (i + 1) + '. ' + n; }).join('\n') +
    '\n\nGuru yang berstatus Izin/Sakit/Dinas Luar/Cuti tidak termasuk daftar ini.';
  try {
    MailApp.sendEmail(tujuan.join(','), 'Absensi ' + (c.sekolah || 'Guru') + ' ' + tgl + ': ' + belum.length + ' guru belum absen', body);
    logAdmin('sistem', 'Notifikasi email: ' + belum.length + ' guru belum absen');
  } catch (e) { Logger.log('Email gagal: ' + e); }
}

// ===== PENGATURAN (TABEL) & HAK AKSES =====
// Tabel 1 (A:C) baris 3-14 | Tabel 2 Tablet (A:B) | Tabel 3 Mapel (E:F) mulai baris 3 | Tabel 4 Kalimat Ucapan (H:I) mulai baris 3
var PENG = [
  ['sekolah', 'Nama Sekolah', 'Tampil di judul dan rekap'],
  ['threshold', 'Threshold Wajah', '0.3 - 0.7 (makin kecil makin ketat)'],
  ['adminGuruId', 'ID Guru Super Admin', 'Untuk menu Absen Saya (ID dari tab Guru)'],
  ['pin_super', 'PIN Super Admin', '4-6 angka'],
  ['pin_admin', 'PIN Admin', '4-6 angka'],
  ['pin_kepsek', 'PIN Kepsek', '4-6 angka'],
  ['pin_wakil', 'PIN Wakil Kepsek', '4-6 angka'],
  ['pin_daftar', 'PIN Daftar Wajah', '4-6 angka'],
  ['fotoHari', 'Lama Simpan Foto (hari)', 'Foto absen dihapus permanen setelah sekian hari (1-30)'],
  ['emailKepsek', 'Email Kepsek', 'Penerima notifikasi guru belum absen (boleh kosong)'],
  ['emailWakil', 'Email Wakil Kepsek', 'Penerima notifikasi guru belum absen (boleh kosong)'],
  ['notifMin', 'Notifikasi jika belum absen >=', 'Jumlah guru (1-99). Email dikirim sekali sesudah jam tutup siang']
];
var P0 = 3, T_HEAD = P0 + PENG.length + 2, T0 = T_HEAD + 1, M0 = 3;
var K0 = 3, KC = 8, UC_N = 40;
var UC_DEF = [
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen hari ini. Selamat mengajar, {sapaan}.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen. Semangat mengajar hari ini.'],
  ['Pagi', 'Selamat pagi {sapaan} {nama}, terima kasih sudah absen. Selamat mengajar.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen. Selamat mengajar.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen pagi ini. Semangat mengajar, {sapaan}.'],
  ['Pagi', 'Selamat pagi {sapaan} {nama}, terima kasih sudah absen. Selamat bertugas.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, absen pagi sudah tercatat. Selamat mengajar hari ini.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen. Semoga pelajaran berjalan lancar.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen. Semoga kesabaran {sapaan} semakin membaja menghadapi anak-anak kita.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen. Bersiap, {sapaan}, mendidik anak-anak supaya jadi wakil presiden sekarang.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen. Hari ini anak-anak sudah menunggu {sapaan} dengan semangat, semoga kita sama semangatnya.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen. Ayo, {sapaan}, buktikan lagi kenapa anak-anak kagum kepada {sapaan}.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen. Hari ini kita buat sekolah makin keren, mulai dari {sapaan}.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen. Semoga bel sekolah terasa lebih singkat hari ini, hahahaha.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen. Selamat berjuang, pahlawan tanpa tanda jasa.'],
  ['Pagi', 'Terima kasih {sapaan} {nama}, sudah absen. Ilmu hari ini sudah siap, tinggal {sapaan} bagikan.'],
  ['Pagi', 'Burung kenari terbang tinggi, hinggap sebentar di pucuk pagi. Terima kasih {sapaan} {nama}, sudah absen pagi ini, semangat mengajar tiada henti.'],
  ['Pagi', 'Pergi ke pasar membeli sayur, pulang singgah membeli cabai. Terima kasih {sapaan} {nama}, sudah absen pagi ini, selamat mengajar dan semangat selalu.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen. Hati-hati di jalan, {sapaan}.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen. Selamat beristirahat.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen. Sampai jumpa besok.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen siang. Selamat beristirahat.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, absen siang sudah tercatat. Hati-hati di jalan, {sapaan}.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen siang. Selamat pulang dan istirahat.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen. Terima kasih atas pengabdiannya hari ini.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen siang. Kerja {sapaan} hari ini luar biasa.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen siang. Dedikasi {sapaan} patut dicontoh.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen siang. Selamat beristirahat, {sapaan} sudah bekerja dengan baik.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen siang. Semangat {sapaan} menginspirasi banyak orang.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen siang. Terima kasih atas kesabaran dan pengabdian {sapaan} hari ini.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen siang. Anak-anak beruntung punya {sapaan} sebagai guru.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen siang. Hari yang hebat untuk {sapaan}, selamat istirahat.'],
  ['Siang', 'Terima kasih {sapaan} {nama}, sudah absen siang. Kontribusi {sapaan} hari ini sangat berarti untuk sekolah.']
];
var MAPEL_DEF = ['Pendidikan Agama dan Budi Pekerti', 'Pendidikan Pancasila', 'Bahasa Indonesia', 'Matematika',
  'IPA (Ilmu Pengetahuan Alam)', 'IPS (Ilmu Pengetahuan Sosial)', 'Bahasa Inggris',
  'PJOK (Pendidikan Jasmani, Olahraga, dan Kesehatan)', 'Informatika',
  'Seni dan Budaya (musik, rupa, teater, atau tari)', 'Prakarya', 'Muatan lokal', 'BTQ', 'BK'];
var HAK_MENU = [['dash', 'Absen'], ['absen', 'Edit Absen Hari Ini'], ['status', 'Status Manual'],
  ['rekap', 'Rekap & Ekspor'], ['libur', 'Hari Libur'], ['set', 'Pengaturan'], ['hak', 'Hak Akses'], ['pin', 'PIN'],
  ['tablet', 'Daftar Tablet'], ['jadwal', 'Jadwal Absen'], ['hariini', 'Jadwal Hari Ini'], ['reset', 'Reset Absen'],
  ['log', 'Log'], ['backup', 'Backup']];
var HAK_LABEL = ['Tidak ada', 'Lihat', 'Edit'];

function teksSel(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'HH:mm');
  }
  return String(v == null ? '' : v).trim();
}

function siapkanPengaturan(old, mapelOld) {
  old = old || {};
  var s = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Pengaturan');
  var data = PENG.map(function (p) {
    var v = old[p[0]];
    if (v === undefined || v === '') v = (DEF_CFG[p[0]] !== undefined) ? DEF_CFG[p[0]] : '';
    return [p[1], String(v), p[2]];
  });
  var nP = PENG.length;
  s.getRange(1, 1).setValue('PENGATURAN').setFontWeight('bold').setFontSize(12);
  s.getRange(2, 1, 1, 3).setValues([['Pengaturan', 'Nilai', 'Keterangan']]).setFontWeight('bold');
  s.getRange(P0, 1, nP, 3).setNumberFormat('@').setValues(data);
  hias(function () { s.getRange(2, 1, nP + 1, 3).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); });
  hias(function () { s.getRange(2, 1, nP + 1, 3).setBorder(true, true, true, true, true, true); });
  PENG.forEach(function (p, i) { validasiBaris(s, p, P0 + i); });
  // Tabel 2: Tablet Terdaftar (kode 6 angka)
  s.getRange(T_HEAD - 1, 1).setValue('TABLET TERDAFTAR').setFontWeight('bold').setFontSize(12);
  s.getRange(T_HEAD, 1, 1, 2).setValues([['Kode Tablet (6 angka)', 'Nama Tablet']]).setFontWeight('bold');
  s.getRange(T_HEAD, 1, 43, 2).setNumberFormat('@');
  pasangRumus(s.getRange(T0, 1, 42, 1), '=OR(A' + T0 + '=""§REGEXMATCH(A' + T0 + '§"^[0-9]{6}$"))', 'Kode tablet = 6 angka');
  hias(function () { s.getRange(T_HEAD, 1, 43, 2).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); });
  hias(function () { s.getRange(T_HEAD, 1, 43, 2).setBorder(true, true, true, true, true, true); });
  // Tabel 3: Daftar Mapel (kolom E:F)
  s.getRange(1, 5).setValue('DAFTAR MAPEL').setFontWeight('bold').setFontSize(12);
  s.getRange(2, 5, 1, 2).setValues([['No', 'Mapel']]).setFontWeight('bold');
  var nM = 60, fm = [];
  for (var i = 0; i < nM; i++) fm.push(['=IF(F' + (M0 + i) + '="","",ROW()-' + (M0 - 1) + ')']);
  s.getRange(M0, 5, nM, 1).setFormulas(fm);
  s.getRange(M0, 6, nM, 1).setNumberFormat('@');
  var dm = (mapelOld && mapelOld.length) ? mapelOld.slice(0, nM) : MAPEL_DEF;
  s.getRange(M0, 6, dm.length, 1).setValues(dm.map(function (m) { return [m]; }));
  hias(function () { s.getRange(2, 5, nM + 1, 2).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); });
  hias(function () { s.getRange(2, 5, nM + 1, 2).setBorder(true, true, true, true, true, true); });
  // Tabel 4: Kalimat Ucapan (kolom H:I)
  s.getRange(1, KC).setValue('KALIMAT UCAPAN').setFontWeight('bold').setFontSize(12);
  s.getRange(2, KC, 1, 2).setValues([['Sesi', 'Kalimat']]).setFontWeight('bold');
  s.getRange(2, KC + 1).setNote('Gunakan {sapaan} dan {nama}. Satu kalimat dipilih acak tiap absen.');
  s.getRange(K0, KC, UC_N, 2).setNumberFormat('@');
  s.getRange(K0, KC, UC_N, 1).setDataValidation(daftarNilai(['Pagi', 'Siang']));
  s.getRange(K0, KC, UC_DEF.length, 2).setValues(UC_DEF);
  hias(function () { s.getRange(2, KC, UC_N + 1, 2).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); });
  hias(function () { s.getRange(2, KC, UC_N + 1, 2).setBorder(true, true, true, true, true, true); });
  rapikanLebarPengaturan(s);
  return s;
}

function rapikanLebarPengaturan(s) {
  s.autoResizeColumns(1, 3);
  s.setColumnWidth(4, 30);
  s.autoResizeColumns(5, 2);
  s.setColumnWidth(7, 30);
  s.setColumnWidth(KC, 70);
  s.setColumnWidth(KC + 1, 520);
}

function ucapanAll() {
  var s = sh('Pengaturan'), o = { pagi: [], siang: [] };
  s.getRange(K0, KC, UC_N, 2).getValues().forEach(function (r) {
    var j = teksSel(r[0]).toLowerCase(), t = teksSel(r[1]);
    if (t && (j === 'pagi' || j === 'siang')) o[j].push(t);
  });
  return o;
}

function simpanUcapan(u) {
  var s = sh('Pengaturan'), rows = [];
  u.pagi.forEach(function (t) { rows.push(['Pagi', t]); });
  u.siang.forEach(function (t) { rows.push(['Siang', t]); });
  s.getRange(K0, KC, UC_N, 2).clearContent();
  if (rows.length) {
    var rg = s.getRange(K0, KC, rows.length, 2);
    rg.setNumberFormat('@');
    rg.setValues(rows);
  }
}

function siapkanHakAkses(old) {
  old = old || {};
  var s = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Hak Akses');
  var data = HAK_MENU.map(function (m) {
    return [m[1]].concat(['kepsek', 'wakil'].map(function (mode) {
      var v = old['akses:' + mode + ':' + m[0]];
      var lv = (v !== undefined && v !== '') ? parseInt(v, 10) : ((AKSES_DEF[mode] || {})[m[0]] || 0);
      return HAK_LABEL[lv] || HAK_LABEL[0];
    }));
  });
  s.getRange(1, 1, 1, 3).setValues([['Menu', 'Kepsek', 'Wakil']]).setFontWeight('bold');
  s.getRange(2, 1, data.length, 3).setNumberFormat('@').setValues(data);
  s.getRange(2, 2, data.length, 2).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(HAK_LABEL, true).setAllowInvalid(false).build());
  hias(function () { s.getRange(1, 1, data.length + 1, 3).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); });
  hias(function () { s.getRange(1, 1, data.length + 1, 3).setBorder(true, true, true, true, true, true); });
  s.setFrozenRows(1);
  s.autoResizeColumns(1, 3);
  return s;
}

// Jalankan SEKALI dari editor setelah menempel kode Ver. 004 (Sheet yang sudah berjalan)
// Membuat ulang tab Pengaturan (nilai lama dibawa), membuat tab Jadwal (kosong), memasang trigger.
function migrasiV4() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Jadwal')) { Logger.log('Sudah dimigrasi (tab Jadwal sudah ada).'); return; }
  var lama = ss.getSheetByName('Pengaturan'), old = {}, mapel = null;
  if (lama) {
    if (String(lama.getRange(1, 1).getValue()) === 'Kunci') {
      var n = lama.getLastRow();
      if (n >= 2) lama.getRange(2, 1, n - 1, 2).getValues().forEach(function (r) {
        old[String(r[0])] = (r[1] instanceof Date) ? '' : String(r[1]);
      });
    } else {
      var lb = {};
      PENG.forEach(function (p) { lb[p[1]] = p[0]; });
      lama.getRange(3, 1, 14, 2).getValues().forEach(function (r) {
        var k = lb[teksSel(r[0])];
        if (k) old[k] = teksSel(r[1]);
      });
      var jm = Math.min(60, lama.getMaxRows() - M0 + 1);
      mapel = [];
      lama.getRange(M0, 6, jm, 1).getValues().forEach(function (r) {
        var t = teksSel(r[0]);
        if (t) mapel.push(t);
      });
    }
    ss.deleteSheet(lama);
  }
  var th = parseFloat(old.threshold);
  if (isNaN(th) || th > 0.45) old.threshold = '0.45';
  siapkanPengaturan(old, mapel);
  sh('Jadwal');
  if (!ss.getSheetByName('Hak Akses')) siapkanHakAkses(old);
  rapikanTab('Guru');
  pasangTrigger();
  logAdmin('sistem', 'Migrasi Ver. 004: Pengaturan, Jadwal, trigger');
  Logger.log('Migrasi selesai. Isi tab Jadwal (atau lewat menu Pengaturan). Tablet lama harus didaftarkan ulang dengan kode 6 angka.');
}

function daftarMapel() {
  var s = sh('Pengaturan'), n = s.getLastRow(), out = [];
  if (n >= M0) s.getRange(M0, 6, n - M0 + 1, 1).getValues().forEach(function (r) {
    var t = teksSel(r[0]);
    if (t) out.push(t);
  });
  return out;
}

// ---------- Pengaturan (baca/tulis ke tabel) ----------
function cfgAll() {
  var s = sh('Pengaturan'), o = {};
  var v = s.getRange(P0, 1, PENG.length, 2).getValues();
  PENG.forEach(function (p, i) { o[p[0]] = teksSel(v[i][1]); });
  // ID guru Admin & Kepsek disimpan di Script Properties (tidak mengubah tata letak sheet)
  var sp = PropertiesService.getScriptProperties().getProperties();
  ['guruAdmin', 'guruKepsek', 'guruWakil'].forEach(function (k) { o[k] = sp[k] || ''; });
  var n = s.getLastRow();
  if (n >= T0) {
    s.getRange(T0, 1, n - T0 + 1, 2).getValues().forEach(function (r) {
      var id = teksSel(r[0]).toUpperCase();
      if (id) o['tablet:' + id] = teksSel(r[1]) || 'Tablet';
    });
  }
  var h = sh('Hak Akses'), hn = h.getLastRow();
  if (hn >= 2) {
    h.getRange(2, 1, hn - 1, 3).getValues().forEach(function (r) {
      var lb = teksSel(r[0]);
      if (lb === 'Jadwal Absen Harian') lb = 'Jadwal Absen'; // nama lama
      var mk = HAK_MENU.filter(function (m) { return m[1] === lb; })[0];
      if (!mk) return;
      ['kepsek', 'wakil'].forEach(function (mode, j) {
        o['akses:' + mode + ':' + mk[0]] = String(Math.max(0, HAK_LABEL.indexOf(teksSel(r[1 + j]))));
      });
    });
  }
  return o;
}

// Jalankan SEKALI dari editor setelah update: tambah baris menu baru ke sheet Hak Akses
function tambahBarisHakAkses() {
  var h = sh('Hak Akses');
  var n = h.getLastRow();
  var ada = n >= 2 ? h.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return teksSel(r[0]); }) : [];
  var tambah = HAK_MENU.filter(function (m) { return ada.indexOf(m[1]) < 0; });
  tambah.forEach(function (m, i) {
    h.getRange(n + 1 + i, 1, 1, 3).setNumberFormat('@').setValues([[m[1], HAK_LABEL[(AKSES_DEF.kepsek || {})[m[0]] || 0], HAK_LABEL[(AKSES_DEF.wakil || {})[m[0]] || 0]]]);
  });
  Logger.log('Baris Hak Akses ditambah: ' + tambah.length);
}

function cfgSet(k, v) {
  v = String(v);
  if (k === 'guruAdmin' || k === 'guruKepsek' || k === 'guruWakil') {
    PropertiesService.getScriptProperties().setProperty(k, v);
    return;
  }
  var s = sh('Pengaturan'), i, rg;
  if (k.indexOf('tablet:') === 0) {
    var id = k.slice(7).toUpperCase(), last = s.getLastRow();
    var vals = last >= T0 ? s.getRange(T0, 1, last - T0 + 1, 1).getValues() : [], r = 0;
    for (i = 0; i < vals.length; i++) {
      if (teksSel(vals[i][0]).toUpperCase() === id) { r = T0 + i; break; }
    }
    if (!r) {
      for (i = 0; i < vals.length; i++) {
        if (teksSel(vals[i][0]) === '') { r = T0 + i; break; }
      }
      if (!r) r = T0 + vals.length;
    }
    rg = s.getRange(r, 1, 1, 2);
    rg.setNumberFormat('@');
    rg.setValues([[id, v]]);
    return;
  }
  if (k.indexOf('akses:') === 0) {
    var p = k.split(':'), h = sh('Hak Akses');
    var mk = HAK_MENU.filter(function (m) { return m[0] === p[2]; })[0];
    if (!mk) return;
    var col = (p[1] === 'kepsek') ? 2 : 3, hn = h.getLastRow();
    var rows = hn >= 2 ? h.getRange(2, 1, hn - 1, 1).getValues() : [];
    for (i = 0; i < rows.length; i++) {
      var lbl = teksSel(rows[i][0]);
      if (lbl === 'Jadwal Absen Harian') lbl = 'Jadwal Absen';
      if (lbl === mk[1]) { h.getRange(i + 2, col).setValue(HAK_LABEL[parseInt(v, 10)] || HAK_LABEL[0]); return; }
    }
    // baris menu belum ada di sheet Hak Akses: tambah dengan nilai bawaan, lalu isi
    var nb = hn + 1;
    h.getRange(nb, 1, 1, 3).setNumberFormat('@').setValues([[mk[1], HAK_LABEL[(AKSES_DEF.kepsek || {})[mk[0]] || 0], HAK_LABEL[(AKSES_DEF.wakil || {})[mk[0]] || 0]]]);
    h.getRange(nb, col).setValue(HAK_LABEL[parseInt(v, 10)] || HAK_LABEL[0]);
    return;
  }
  for (i = 0; i < PENG.length; i++) {
    if (PENG[i][0] === k) {
      rg = s.getRange(P0 + i, 2);
      rg.setNumberFormat('@');
      rg.setValue(v);
      return;
    }
  }
}

function cfgDel(k) {
  if (k.indexOf('tablet:') !== 0) return;
  var s = sh('Pengaturan'), id = k.slice(7).toUpperCase(), last = s.getLastRow();
  if (last < T0) return;
  var vals = s.getRange(T0, 1, last - T0 + 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (teksSel(vals[i][0]).toUpperCase() === id) { s.getRange(T0 + i, 1, 1, 2).clearContent(); return; }
  }
}

// ===== TAB SHEET: Guru, Wajah, Absensi, Status Manual =====
var KOL_LAMA = { 'Absensi': [1, 2, 4, 5, 6, 7, 8, 9], 'Status Manual': [1, 3, 4, 5, 6] };

function baris(nama) {
  var s = sh(nama), n = s.getLastRow();
  if (n < 2) return [];
  var v = s.getRange(2, 1, n - 1, HEAD[nama].length).getValues();
  var k = KOL_LAMA[nama];
  if (!k) return v;
  return v.map(function (r) { return k.map(function (i) { return r[i]; }); }); // bentuk lama: tanpa No & Nama
}

function nomori(nama) { // No urut 1..n (Status Manual)
  var s = sh(nama), n = s.getLastRow();
  if (n < 2) return;
  var no = [];
  for (var i = 1; i < n; i++) no.push([String(i)]);
  var c = s.getRange(2, 1, n - 1, 1);
  c.setNumberFormat('@');
  c.setValues(no);
}

// ---------- Absensi: A No | B Tanggal | C GuruID | D Nama | E Jam Pagi | F Status Pagi | G Jam Siang | H Ket ----------
function cariAbsen(tgl, guruId) {
  var s = sh('Absensi'), n = s.getLastRow();
  if (n < 2) return 0;
  var v = s.getRange(2, 2, n - 1, 2).getValues();
  for (var i = 0; i < v.length; i++) {
    if (normTgl(v[i][0]) === tgl && String(v[i][1]) === String(guruId)) return i + 2;
  }
  return 0;
}

// Buat baris semua guru aktif untuk tanggal tgl (urut A-Z) bila belum ada; tanggal terbaru di atas
function siapkanAbsensi(tgl) {
  var s = sh('Absensi'), n = s.getLastRow(), ada = false, maks = '';
  if (n >= 2) {
    s.getRange(2, 2, n - 1, 1).getValues().forEach(function (r) {
      var t = normTgl(r[0]);
      if (t === tgl) ada = true;
      if (t > maks) maks = t;
    });
  }
  if (ada) return false;
  var g = guruAktif().sort(function (a, b) { return String(a.nama).localeCompare(String(b.nama), 'id'); });
  if (!g.length) return false;
  var rows = g.map(function (x, i) { return [String(i + 1), tgl, x.id, x.nama, '', '', '', '', '', '']; });
  s.insertRowsBefore(2, rows.length);
  var rg = s.getRange(2, 1, rows.length, 10);
  rg.setFontWeight('normal').setBackground(null);
  rg.setNumberFormat('@');
  rg.setValues(rows);
  s.getRange(2, 6, rows.length, 1).setDataValidation(daftarNilai(['tepat', 'terlambat', 'libur']));
  if (maks > tgl) urutkanAbsensi(); // tanggal lampau: tempatkan sesuai urutan
  return true;
}

function urutkanAbsensi() { // tanggal terbaru di atas, nama A-Z, No urut per tanggal
  var s = sh('Absensi'), n = s.getLastRow();
  if (n < 3) return;
  var rg = s.getRange(2, 1, n - 1, 10);
  rg.sort([{ column: 2, ascending: false }, { column: 4, ascending: true }]);
  var no = 0, prev = '', out = [];
  rg.getValues().forEach(function (r) {
    var t = normTgl(r[1]);
    no = (t === prev) ? no + 1 : 1;
    prev = t;
    out.push([String(no)]);
  });
  s.getRange(2, 1, out.length, 1).setValues(out);
}

// kol (penomoran lama): 3=JamPagi 4=StatusPagi 5=JamSiang 6=Ket
function tulisAbsen(tgl, guruId, kol, nilai) {
  var r = cariAbsen(tgl, guruId);
  if (!r) {
    siapkanAbsensi(tgl);
    r = cariAbsen(tgl, guruId);
    if (!r) { // guru baru / belum ada barisnya
      appendTeks('Absensi', ['', tgl, String(guruId), petaNama()[String(guruId)] || '', '', '', '', '', '', '']);
      urutkanAbsensi();
      r = cariAbsen(tgl, guruId);
    }
  }
  var c = sh('Absensi').getRange(r, kol + 2);
  c.setNumberFormat('@');
  c.setValue(nilai);
}

function jamPunya(tgl, guruId) { // jam, keterangan, dan foto yang sudah ada
  var r = cariAbsen(tgl, guruId);
  if (!r) return { pagi: '', siang: '', ket: '', fotoPagi: '', fotoSiang: '', liburPagi: false, liburSiang: false };
  var v = sh('Absensi').getRange(r, 1, 1, 10).getValues()[0];
  // penanda LIBUR dianggap kosong: guru yang datang di hari liburnya boleh absen, penanda diganti jam asli
  var lp = String(v[4] || '') === PENANDA_LIBUR, ls = String(v[6] || '') === PENANDA_LIBUR;
  return { pagi: lp ? '' : String(v[4] || ''), siang: ls ? '' : String(v[6] || ''), ket: String(v[7] || ''),
           fotoPagi: String(v[8] || ''), fotoSiang: String(v[9] || ''), liburPagi: lp, liburSiang: ls };
}

// ---------- Keterangan manual: penanda jenis (Tugas Sekolah, Izin Terlambat, Pulang Cepat) ----------
// Segmen dipisah ' | ', contoh: 'Tugas Sekolah (Pagi): dipanggil kepsek | Pulang Cepat: anak sakit'
var KET_TSP = 'Tugas Sekolah (Pagi)', KET_TSS = 'Tugas Sekolah (Siang)', KET_IT = 'Izin Terlambat', KET_PC = 'Pulang Cepat';
var KET_MNP = 'Absen Manual Pagi (Foto)', KET_MNS = 'Absen Manual Siang (Foto)';
function ketJenis(teks) {
  var o = { tsPagi: false, tsSiang: false, izinTerlambat: false, pulangCepat: false,
            manualPagi: false, manualSiang: false,
            alasanPagi: '', alasanSiang: '', teks: String(teks || '') };
  o.teks.split(' | ').forEach(function (seg) {
    seg = seg.trim();
    var alasan = seg.indexOf(': ') >= 0 ? seg.slice(seg.indexOf(': ') + 2) : '';
    if (seg.indexOf(KET_TSP) === 0) { o.tsPagi = true; o.alasanPagi = alasan; }
    else if (seg.indexOf(KET_IT) === 0) { o.izinTerlambat = true; o.alasanPagi = alasan; }
    else if (seg.indexOf(KET_TSS) === 0) { o.tsSiang = true; o.alasanSiang = alasan; }
    else if (seg.indexOf(KET_PC) === 0) { o.pulangCepat = true; o.alasanSiang = alasan; }
    else if (seg.indexOf(KET_MNP) === 0) { o.manualPagi = true; }
    else if (seg.indexOf(KET_MNS) === 0) { o.manualSiang = true; }
  });
  return o;
}
function ketTambah(lama, seg) {
  lama = String(lama || '').trim();
  return lama ? lama + ' | ' + seg : seg;
}
function ketBuang(lama, awalan) { // buang segmen yang diawali salah satu awalan
  return String(lama || '').split(' | ').filter(function (seg) {
    seg = seg.trim();
    return seg && !awalan.some(function (a) { return seg.indexOf(a) === 0; });
  }).join(' | ');
}
function bersihKetSesi(tgl, guruId, sesiAbsen) { // jam dihapus: buang penanda sesi itu
  var r = cariAbsen(tgl, guruId);
  if (!r) return;
  var c = sh('Absensi').getRange(r, 8), lama = String(c.getValue() || '');
  var baru = ketBuang(lama, sesiAbsen === 'pagi' ? [KET_TSP, KET_IT, KET_MNP] : [KET_TSS, KET_PC, KET_MNS]);
  if (baru !== lama) { c.setNumberFormat('@'); c.setValue(baru); }
}

function siapkanHariIni() { // dipanggil trigger harian
  var now = new Date(), tgl = tglStr(now);
  bersihkanKhusus();
  if (isLibur(tgl, now)) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { siapkanAbsensi(tgl); } finally { lock.releaseLock(); }
}

// Jalankan SEKALI dari editor: pasang semua trigger
// (baris absen jam 00, hapus foto lama jam 01, backup mingguan Minggu jam 02, cek notifikasi tiap jam, isi libur guru tiap 5 menit)
function pasangTrigger() {
  var daftar = ['siapkanHariIni', 'hapusFotoLama', 'backupOtomatis', 'cekNotifikasi', 'isiLiburHariIni'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (daftar.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('siapkanHariIni').timeBased().everyDays(1).atHour(0).create();
  ScriptApp.newTrigger('hapusFotoLama').timeBased().everyDays(1).atHour(1).create();
  ScriptApp.newTrigger('backupOtomatis').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();
  ScriptApp.newTrigger('cekNotifikasi').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('isiLiburHariIni').timeBased().everyMinutes(5).create();
  Logger.log('Trigger terpasang: baris absen (00), hapus foto lama (01), backup mingguan (Minggu 02), cek notifikasi (tiap jam), isi libur guru (tiap 5 menit).');
}

// ---------- Tampilan tab ----------
function daftarNilai(arr, izinLain) {
  return SpreadsheetApp.newDataValidation().requireValueInList(arr, true).setAllowInvalid(!!izinLain).build();
}

function rapikanTab(nama, s) {
  s = s || sh(nama);
  var nk = HEAD[nama].length, mx = s.getMaxRows();
  s.getRange(1, 1, 1, nk).setFontWeight('bold').setBackground('#e2e8f0');
  s.setFrozenRows(1);
  if (nama === 'Absensi') {
    if (!s.getFilter()) s.getRange(1, 1, mx, nk).createFilter();
    s.getRange(2, 6, mx - 1, 1).setDataValidation(daftarNilai(['tepat', 'terlambat', 'libur']));
  } else if (nama === 'Status Manual') {
    if (!s.getFilter()) s.getRange(1, 1, mx, nk).createFilter();
    s.getRange(2, 4, mx - 1, 1).setDataValidation(daftarNilai(STATUS_OK));
  } else if (nama === 'Guru') {
    var mp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengaturan');
    if (mp) {
      s.getRange(2, 4, mx - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(mp.getRange(M0, 6, 60, 1), true).setAllowInvalid(true)
        .setHelpText('Pilih mapel dari daftar. Untuk lebih dari satu, pisahkan dengan titik koma: Matematika; Informatika').build());
    }
    s.getRange(2, 5, mx - 1, 1).setDataValidation(daftarNilai(['Ya', 'Tidak']));
    s.getRange(2, 6, mx - 1, 1).setDataValidation(daftarNilai(['Belum', 'Terdaftar']));
    if (s.getBandings().length === 0) {
      hias(function () { s.getRange(1, 1, mx, nk).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); });
    }
  } else if (nama === 'Wajah') {
    s.getRange(1, 1, mx, nk).setBackground('#e5e7eb').setFontColor('#6b7280');
    s.setTabColor('#9ca3af');
    s.getRange(2, 2, mx - 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    var pr = s.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0] || s.protect();
    pr.setWarningOnly(true).setDescription('Data wajah: jangan diedit manual');
    s.setColumnWidth(1, 90); s.setColumnWidth(2, 320); s.setColumnWidth(3, 150);
    return;
  }
  s.autoResizeColumns(1, nk);
  if (nama === 'Absensi') {
    s.setColumnWidth(9, 150); s.setColumnWidth(10, 150);
    s.getRange(2, 9, mx - 1, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}

// Jalankan SEKALI dari editor setelah menempel kode Ver. 006 (sheet yang sudah berjalan): tambah kolom Libur di sheet Guru,
// izinkan status 'libur' di sheet Absensi, dan pasang ulang trigger. Data TIDAK dihapus.
function migrasiLibur() {
  pasangLibur(sh('Guru'));
  var s = sh('Absensi'), mx = s.getMaxRows();
  s.getRange(2, 6, mx - 1, 1).setDataValidation(daftarNilai(['tepat', 'terlambat', 'libur']));
  pasangTrigger();
  logAdmin('sistem', 'Migrasi Ver. 006: kolom Libur guru');
  Logger.log('Migrasi libur selesai.');
}

// Jalankan SEKALI dari editor: HAPUS data Absensi & Status Manual lama, buat ulang dengan kolom baru
function migrasiTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Absensi', 'Status Manual'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s) ss.deleteSheet(s);
    sh(n);
  });
  rapikanTab('Guru');
  rapikanTab('Wajah');
  siapkanHariIni();
  logAdmin('sistem', 'Migrasi tab Guru, Wajah, Absensi, Status Manual');
  Logger.log('Migrasi tab selesai.');
}

// ---------- Mapel ----------
function getMapel(token) {
  butuh(token, 'guru', 1);
  return daftarMapel();
}

// Buang mapel guru yang sudah tidak ada di daftar (lama -> baru = ganti nama mapel)
function bersihkanMapel(lama, baru) {
  var daftar = daftarMapel(), s = sh('Guru'), n = s.getLastRow();
  if (n < 2) return;
  var rg = s.getRange(2, 4, n - 1, 1), v = rg.getValues(), ubah = false;
  v.forEach(function (r) {
    var asal = String(r[0]).trim();
    if (!asal) return;
    var hasil = asal.split(/\s*;\s*/).map(function (m) { return (lama && m === lama) ? baru : m; })
      .filter(function (m) { return m && daftar.indexOf(m) >= 0; }).join('; ');
    if (hasil !== asal) { r[0] = hasil; ubah = true; }
  });
  if (ubah) { rg.setNumberFormat('@'); rg.setValues(v); }
}

// Lebar kolom otomatis + sinkron mapel saat Sheet diedit
function onEdit(e) {
  try {
    var s = e.range.getSheet(), n = s.getName();
    if (n === 'Pengaturan') {
      rapikanLebarPengaturan(s);
      if (e.range.getColumn() === 6 && e.range.getRow() >= M0) {
        var tunggal = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;
        var baru = String(e.value || '').trim();
        if (tunggal && e.oldValue && baru) bersihkanMapel(String(e.oldValue).trim(), baru); // ganti nama
        else bersihkanMapel(); // dihapus -> kosong di guru
      }
    } else if (n === 'Hak Akses') {
      s.autoResizeColumns(1, 3);
    } else if (n === 'Jadwal') {
      s.autoResizeColumns(1, 6);
    } else if (n === 'Guru' || n === 'Absensi' || n === 'Status Manual') {
      s.autoResizeColumns(1, Math.min(HEAD[n].length, 8));
    }
  } catch (x) {}
}

// ===== FOTO ABSEN (Drive) =====
var FOTO_FOLDER = 'Foto Absen';

function folderFoto() {
  var it = DriveApp.getFoldersByName(FOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOTO_FOLDER);
}

function catatAbsen(devId, guruId, sesiAbsen, token, foto, opsi) {
  var r = catatAbsenInti(devId, guruId, sesiAbsen, token, opsi);
  if (r && r.ok && foto) simpanFoto(tglStr(), String(guruId), sesiAbsen, r.jam, foto);
  return r;
}

function catatAbsenBatch(devId, arr) {
  var fq = [];
  var r = catatAbsenBatchInti(devId, arr, fq);
  fq.forEach(function (x) { simpanFoto(x.tgl, x.id, x.sesi, x.jam, x.foto); });
  return r;
}

// Simpan foto ke Drive (di luar kunci), lalu catat tautannya di kolom Foto Pagi / Foto Siang
function simpanFoto(tgl, guruId, sesiAbsen, jam, foto) {
  try {
    var m = String(foto || '').match(/^data:image\/jpeg;base64,([A-Za-z0-9+\/=]+)$/);
    if (!m || m[1].length > 60000) return;
    var nama = tgl + '_' + String(jam).replace(':', '') + '_' + sesiAbsen + '_' + guruId + '.jpg';
    var f = folderFoto().createFile(Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/jpeg', nama));
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try { tulisAbsen(tgl, guruId, sesiAbsen === 'pagi' ? 7 : 8, f.getUrl()); } finally { lock.releaseLock(); }
  } catch (e) { Logger.log('Foto gagal disimpan: ' + e); }
}

function bagi(arr, n) {
  var out = [];
  for (var i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
  return out;
}

// Foto sehari untuk menu Absen: {guruId: {pagi: dataURL, siang: dataURL}} (tidak untuk Wakil)
function getFoto(token, tgl) {
  var s = butuh(token, 'dash', 1);
  if (s.mode === 'wakil') throw new Error('Tidak punya akses');
  return fotoHari(tgl);
}

function fotoHari(tgl) {
  tgl = String(tgl || tglStr());
  if (!RE_TGL.test(tgl)) throw new Error('Format tanggal salah');
  var daftar = [];
  baris('Absensi').forEach(function (r) {
    if (normTgl(r[0]) !== tgl) return;
    [['pagi', r[6]], ['siang', r[7]]].forEach(function (x) {
      var m = String(x[1] || '').match(/\/d\/([\w-]+)/);
      if (m) daftar.push({ g: String(r[1]), s: x[0], f: m[1] });
    });
  });
  var cache = CacheService.getScriptCache(), ada = {};
  bagi(daftar.map(function (d) { return 'ft_' + d.f; }), 90).forEach(function (k) {
    var h = cache.getAll(k);
    Object.keys(h).forEach(function (x) { ada[x] = h[x]; });
  });
  var perlu = daftar.filter(function (d) { return !ada['ft_' + d.f]; });
  if (perlu.length) {
    var tok = ScriptApp.getOAuthToken();
    var res = UrlFetchApp.fetchAll(perlu.map(function (d) {
      return { url: 'https://www.googleapis.com/drive/v3/files/' + d.f + '?alt=media',
               headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true };
    }));
    var baru = {};
    res.forEach(function (r, i) {
      if (r.getResponseCode() !== 200) return;
      var v = 'data:image/jpeg;base64,' + Utilities.base64Encode(r.getContent());
      if (v.length < 90000) { baru['ft_' + perlu[i].f] = v; ada['ft_' + perlu[i].f] = v; }
    });
    bagi(Object.keys(baru), 90).forEach(function (ks) {
      var o = {};
      ks.forEach(function (k) { o[k] = baru[k]; });
      cache.putAll(o, 21600);
    });
  }
  var out = {};
  daftar.forEach(function (d) {
    var v = ada['ft_' + d.f];
    if (!v) return;
    if (!out[d.g]) out[d.g] = {};
    out[d.g][d.s] = v;
  });
  return out;
}

// Trigger harian: hapus PERMANEN foto yang lebih lama dari "Lama Simpan Foto" (default 7 hari)
function hapusFotoLama() {
  var hari = parseInt(cfgAll().fotoHari, 10) || 7;
  var batas = new Date(Date.now() - hari * 86400000), n = 0;
  var it = folderFoto().getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getDateCreated().getTime() >= batas.getTime()) continue;
    try {
      if (typeof Drive !== 'undefined') Drive.Files.remove(f.getId()); // permanen, tanpa tempat sampah
      else f.setTrashed(true); // cadangan bila layanan Drive API belum ditambahkan
      n++;
    } catch (e) { Logger.log('Gagal hapus ' + f.getName() + ': ' + e); }
  }
  var s = sh('Absensi'), last = s.getLastRow(), cut = tglStr(batas);
  if (last >= 2) {
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var tg = s.getRange(2, 2, last - 1, 1).getValues();
      var rg = s.getRange(2, 9, last - 1, 2), v = rg.getValues(), ubah = false;
      v.forEach(function (r, i) {
        if (normTgl(tg[i][0]) >= cut) return;
        for (var k = 0; k < 2; k++) {
          if (r[k] && r[k] !== 'kedaluwarsa') { r[k] = 'kedaluwarsa'; ubah = true; }
        }
      });
      if (ubah) { rg.setNumberFormat('@'); rg.setValues(v); }
    } finally { lock.releaseLock(); }
  }
  logAdmin('sistem', 'Hapus foto lama: ' + n + ' file');
}

// Validasi dengan rumus sering ditolak Sheets (beda locale), jadi tidak dipasang.
// Pemeriksaan PIN, jam, dan kode tablet tetap dilakukan server saat simpan dari aplikasi.
function pasangRumus(rg, rumus, bantuan) { return false; }

// Jalankan langkah tampilan; kegagalan tampilan tidak boleh menghentikan setup
function hias(fn) { try { fn(); } catch (e) { Logger.log('Tampilan dilewati: ' + e); } }

// Validasi sel baris Pengaturan (PIN, lama simpan foto, batas notifikasi)
function validasiBaris(s, p, r) {
  var f = null;
  if (p[0].indexOf('pin_') === 0) {
    f = '=REGEXMATCH(B' + r + '§"^[0-9]{4,6}$")';
  } else if (p[0] === 'fotoHari') {
    f = '=REGEXMATCH(B' + r + '§"^([1-9]|[12][0-9]|30)$")';
  } else if (p[0] === 'notifMin') {
    f = '=REGEXMATCH(B' + r + '§"^([1-9]|[1-9][0-9])$")';
  }
  if (f) pasangRumus(s.getRange(r, 2), f, p[2]);
}

// Isi baris Pengaturan yang belum ada
function lengkapiPengaturan() {
  var s = sh('Pengaturan');
  PENG.forEach(function (p, i) {
    var r = P0 + i;
    if (teksSel(s.getRange(r, 1).getValue()) !== '') return;
    var rg = s.getRange(r, 1, 1, 3);
    rg.setNumberFormat('@');
    rg.setValues([[p[1], (DEF_CFG[p[0]] !== undefined) ? String(DEF_CFG[p[0]]) : '', p[2]]]);
    rg.setBorder(true, true, true, true, true, true);
    validasiBaris(s, p, r);
  });
  rapikanLebarPengaturan(s);
}

// Jalankan SEKALI dari editor (hanya jika tab Absensi belum punya kolom foto)
function migrasiFoto() {
  var s = sh('Absensi'), mx = s.getMaxRows();
  s.getRange(1, 9, 1, 2).setValues([['Foto Pagi', 'Foto Siang']]).setFontWeight('bold').setBackground('#e2e8f0');
  s.getRange(1, 9, mx, 2).setNumberFormat('@');
  var f = s.getFilter();
  if (f) f.remove();
  s.getRange(1, 1, mx, 10).createFilter();
  s.getRange(2, 9, mx - 1, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  s.setColumnWidth(9, 150);
  s.setColumnWidth(10, 150);
  lengkapiPengaturan();
  folderFoto();
  pasangTrigger();
  logAdmin('sistem', 'Migrasi foto absen');
  Logger.log('Migrasi foto selesai.');
}

// ===== BAGIAN 5: VER. 007 =====
// Link guru (login nama + PIN), lokasi sekolah, PIN guru, piket & libur, jadwal khusus / hari ini,
// reset absen, Absen Saya untuk semua akun (Super, Admin, Kepsek, Wakil).

var PIN_DEF = '000000';
var GURU_TTL = 3600;      // sesi HP guru di server: 1 jam (di aplikasi sesi berakhir otomatis setelah 10 menit tanpa sentuhan)
var RE_PIN = /^\d{6}$/;
var MODE_MANDIRI = ['super', 'admin', 'kepsek', 'wakil', 'guru']; // boleh absen sendiri dari HP (tanpa kode perangkat)
var HARI_PIKET = { 'Senin': 1, 'Selasa': 2, 'Rabu': 3, 'Kamis': 4, 'Jumat': 5, 'Sabtu': 6 };

// ---------- Cache ----------
function hapusCacheCfg() { CacheService.getScriptCache().remove('cfg_pub'); }
function hapusCacheGuru() { CacheService.getScriptCache().remove('dng'); }

// ---------- PIN guru (kolom I tab Guru; kosong = 000000) ----------
function pinNilai(v) {
  var t = teksSel(v);
  if (/^\d{1,6}$/.test(t)) t = ('000000' + t).slice(-6); // angka yang kehilangan nol di depan
  return RE_PIN.test(t) ? t : PIN_DEF;
}

function pinGuruDari(id) {
  var p = PIN_DEF;
  baris('Guru').some(function (r) {
    if (String(r[0]) === String(id)) { p = pinNilai(r[8]); return true; }
    return false;
  });
  return p;
}

function tulisPin(id, pin) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var r = cariBaris('Guru', 1, id);
    if (!r) return false;
    var c = sh('Guru').getRange(r, 9);
    c.setNumberFormat('@');
    c.setValue(pin);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// kolom PIN (I) & Piket (J) di tab Guru + dropdown Senin-Sabtu
function pasangPinPiket(s) {
  var cache = CacheService.getScriptCache();
  if (cache.get('pp_ok')) return;
  var h = s.getRange(1, 9, 1, 2).getValues()[0];
  if (String(h[0]) !== 'PIN' || String(h[1]) !== 'Piket') {
    s.getRange(1, 9, 1, 2).setValues([['PIN', 'Piket']]).setFontWeight('bold');
    hias(function () { s.getRange(1, 9, 1, 2).setBackground('#e2e8f0'); });
  }
  if (s.getRange(2, 10).getDataValidation() === null) {
    s.getRange(1, 9, s.getMaxRows(), 2).setNumberFormat('@');
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'], true).setAllowInvalid(false).build();
    s.getRange(2, 10, Math.max(1, s.getMaxRows() - 1), 1).setDataValidation(rule);
  }
  cache.put('pp_ok', '1', 21600);
}

// ---------- Link guru (kode rahasia di link) ----------
function kodeBaru() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function linkKode() {
  var p = PropertiesService.getScriptProperties();
  var k = p.getProperty('linkKode');
  if (!k) { k = kodeBaru(); p.setProperty('linkKode', k); }
  return k;
}
function cekKode(kode) { return !!kode && String(kode) === linkKode(); }

function getLinkGuru(token) {
  butuh(token, 'link', 1);
  return { kode: linkKode() };
}

function buatUlangLink(token) {
  var s = butuh(token, 'link', 2);
  var k = kodeBaru();
  PropertiesService.getScriptProperties().setProperty('linkKode', k);
  logAdmin(s.mode, 'Buat ulang link guru');
  return { ok: true, kode: k, pesan: 'Link baru dibuat. Link lama tidak berlaku lagi.' };
}

// ---------- Lokasi sekolah (Script Properties, tidak mengubah tata letak sheet) ----------
function lokasiSekolah() {
  var p = PropertiesService.getScriptProperties().getProperties();
  var lat = parseFloat(p.lokLat), lng = parseFloat(p.lokLng), r = parseInt(p.lokRadius, 10) || 100;
  var ada = !isNaN(lat) && !isNaN(lng);
  return { ada: ada, lat: ada ? lat : '', lng: ada ? lng : '', radius: r };
}

function getLokasi(token) {
  butuh(token, 'lokasi', 1);
  return lokasiSekolah();
}

function simpanLokasi(token, d) {
  var s = butuh(token, 'lokasi', 2);
  d = d || {};
  var lat = parseFloat(d.lat), lng = parseFloat(d.lng), r = parseInt(d.radius, 10);
  if (isNaN(lat) || lat < -90 || lat > 90) return { ok: false, pesan: 'Latitude tidak valid (-90 sampai 90)' };
  if (isNaN(lng) || lng < -180 || lng > 180) return { ok: false, pesan: 'Longitude tidak valid (-180 sampai 180)' };
  if (isNaN(r) || r < 20 || r > 2000) return { ok: false, pesan: 'Radius harus 20 - 2000 meter' };
  PropertiesService.getScriptProperties().setProperties({ lokLat: String(lat), lokLng: String(lng), lokRadius: String(r) });
  logAdmin(s.mode, 'Ubah lokasi sekolah, radius ' + r + ' m');
  return { ok: true, pesan: 'Lokasi sekolah disimpan' };
}

function jarakM(lat1, lng1, lat2, lng2) { // meter (haversine)
  var R = 6371000, rad = function (x) { return x * Math.PI / 180; };
  var dl = rad(lat2 - lat1), dg = rad(lng2 - lng1);
  var h = Math.sin(dl / 2) * Math.sin(dl / 2) + Math.cos(rad(lat1)) * Math.cos(rad(lat2)) * Math.sin(dg / 2) * Math.sin(dg / 2);
  return 2 * R * Math.asin(Math.sqrt(h));
}

// o: {lat, lng, acc}. Toleransi akurasi GPS maksimal 30 m.
function cekLokasi(o) {
  var L = lokasiSekolah();
  if (!L.ada) return { ok: false, pesan: 'Lokasi sekolah belum diatur, hubungi admin' };
  o = o || {};
  var lat = parseFloat(o.lat), lng = parseFloat(o.lng);
  if (isNaN(lat) || isNaN(lng)) return { ok: false, pesan: 'Lokasi belum terbaca. Aktifkan GPS dan izinkan lokasi.' };
  var acc = Math.min(Math.max(parseFloat(o.acc) || 0, 0), 30);
  var j = jarakM(lat, lng, L.lat, L.lng);
  if (j - acc > L.radius) return { ok: false, pesan: 'Anda di luar area sekolah (' + Math.round(j) + ' m dari sekolah)' };
  return { ok: true, jarak: j };
}

// ---------- Data guru ringkas, piket ----------
function infoGuru(id) {
  var o = { nama: '', sapaan: '' };
  baris('Guru').some(function (r) {
    if (String(r[0]) === String(id)) { o.nama = String(r[1]); o.sapaan = normSapaan(r[6]); return true; }
    return false;
  });
  return o;
}

function normPiket(v) { v = String(v || '').trim(); return HARI_PIKET.hasOwnProperty(v) ? v : ''; }

function piketMap() { // guruId -> hari piket (1-6)
  var o = {};
  baris('Guru').forEach(function (r) {
    var h = HARI_PIKET[normPiket(r[9])];
    if (h) o[String(r[0])] = h;
  });
  return o;
}

function piketHari(guruId, tgl) { // true bila tgl adalah hari piket guru
  var h = piketMap()[String(guruId)];
  return !!h && h === hariKe(tgl);
}

// batas terlambat yang berlaku: guru piket memakai batas piket bila lebih awal dari batas biasa
function batasTelat(J, piket) {
  if (piket && J.telatPiket && menitDari(J.telatPiket) < menitDari(J.telat)) return J.telatPiket;
  return J.telat;
}

function getPiketLibur(token) {
  butuh(token, 'piket', 1);
  rapikanGuru();
  return baris('Guru').filter(function (r) { return String(r[1]).trim(); }).map(function (r) {
    return { id: String(r[0]), nama: String(r[1]), aktif: String(r[4]) === 'Ya', libur: normLibur(r[7]), piket: normPiket(r[9]) };
  });
}

// arr: [{id, libur?, piket?}]; libur Selasa-Sabtu, piket Senin-Sabtu, tidak boleh jatuh di hari yang sama
function simpanPiketLibur(token, arr) {
  var s = butuh(token, 'piket', 2);
  arr = arr || [];
  if (!arr.length) return { ok: true, pesan: 'Tidak ada perubahan' };
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var cur = {};
    baris('Guru').forEach(function (r) {
      cur[String(r[0])] = { nama: String(r[1]), libur: normLibur(r[7]), piket: normPiket(r[9]) };
    });
    var baru = {};
    for (var i = 0; i < arr.length; i++) {
      var x = arr[i], id = String(x.id || '');
      if (!cur[id]) return { ok: false, pesan: 'Guru tidak ditemukan' };
      var o = baru[id] || { libur: cur[id].libur, piket: cur[id].piket };
      if (x.libur !== undefined) o.libur = normLibur(x.libur);
      if (x.piket !== undefined) o.piket = normPiket(x.piket);
      baru[id] = o;
    }
    var ids = Object.keys(baru);
    for (var a = 0; a < ids.length; a++) {
      var o2 = baru[ids[a]];
      if (o2.libur && o2.libur === o2.piket) {
        return { ok: false, pesan: cur[ids[a]].nama + ': hari libur dan hari piket tidak boleh sama (' + o2.libur + ')' };
      }
    }
    var sg = sh('Guru');
    ids.forEach(function (id) {
      var r = cariBaris('Guru', 1, id), o3 = baru[id];
      var cl = sg.getRange(r, 8); cl.setNumberFormat('@'); cl.setValue(o3.libur);
      var cp = sg.getRange(r, 10); cp.setNumberFormat('@'); cp.setValue(o3.piket);
    });
    pasangLibur(sg);
    pasangPinPiket(sg);
    var hariIni = tglStr();
    ids.forEach(function (id) { if (baru[id].libur !== cur[id].libur) hapusPenandaLibur(hariIni, id); });
    CacheService.getScriptCache().remove('lbk_' + hariIni); // isi ulang libur hari ini sesuai pengaturan baru
    logAdmin(s.mode, 'Ubah piket/libur ' + ids.length + ' guru');
    return { ok: true, pesan: ids.length + ' guru diperbarui' };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Login guru (link + nama + PIN) ----------
function getDaftarNamaGuru(kode) {
  if (!cekKode(kode)) throw new Error('Link tidak berlaku. Minta link baru dari admin.');
  var cache = CacheService.getScriptCache(), hit = cache.get('dng');
  if (hit) return JSON.parse(hit);
  // baca langsung (tanpa rapikanGuru yang menulis ke sheet); cache 1 jam, dihapus tiap data guru berubah
  var out = baris('Guru').filter(function (r) { return String(r[1]).trim() && String(r[4]) === 'Ya'; })
    .map(function (r) { return { id: String(r[0]), nama: String(r[1]) }; })
    .sort(function (a, b) { return String(a.nama).localeCompare(String(b.nama), 'id'); });
  try { cache.put('dng', JSON.stringify(out), 3600); } catch (e) {}
  return out;
}

function loginGuru(kode, guruId, pin) {
  if (!cekKode(kode)) return { ok: false, kodeSalah: true, pesan: 'Link tidak berlaku. Minta link baru dari admin.' };
  guruId = String(guruId || '');
  var cache = CacheService.getScriptCache(), kk = 'gagal_g_' + guruId; // kunci percobaan PIN per guru
  var n = cekKunci(kk);
  if (n >= MAX_GAGAL) return { ok: false, pesan: 'Terkunci sementara. Coba lagi 5 menit atau lapor admin.' };
  var g = guruAktif().filter(function (x) { return x.id === guruId; })[0];
  if (!g) return { ok: false, pesan: 'Guru tidak ditemukan' };
  var benar = pinGuruDari(guruId);
  if (!RE_PIN.test(String(pin || '')) || String(pin) !== benar) {
    cache.put(kk, String(n + 1), KUNCI_DETIK);
    return { ok: false, pesan: (n + 1 >= MAX_GAGAL) ? 'Terkunci sementara. Coba lagi 5 menit atau lapor admin.' : 'PIN salah' };
  }
  cache.remove(kk);
  var token = buatToken({ mode: 'guru', devId: '', guruId: guruId }, GURU_TTL);
  return { ok: true, token: token, nama: String(g.nama), sapaan: normSapaan(g.sapaan), akses: aksesUntuk('guru'),
           guruId: guruId, pinDefault: benar === PIN_DEF, pesan: '' };
}

function keluarSesi(token) {
  if (token) CacheService.getScriptCache().remove('tk_' + String(token));
  return { ok: true };
}

function ubahPinGuru(token, lama, baru) {
  var s = sesi(token);
  if (!s || s.mode !== 'guru') throw new Error('Sesi habis, silakan login ulang');
  lama = String(lama || ''); baru = String(baru || '');
  if (!RE_PIN.test(baru)) return { ok: false, pesan: 'PIN baru harus 6 angka' };
  if (baru === PIN_DEF) return { ok: false, pesan: 'PIN baru tidak boleh 000000' };
  if (lama !== pinGuruDari(s.guruId)) return { ok: false, pesan: 'PIN lama salah' };
  tulisPin(s.guruId, baru);
  logAdmin('guru', 'Ubah PIN: ' + (infoGuru(s.guruId).nama || s.guruId)); // nilai PIN tidak dicatat
  return { ok: true, pesan: 'PIN diperbarui' };
}

// PIN semua guru (tab PIN di Data Guru). Menu Data Guru hanya untuk Super Admin & Admin.
function getPinGuru(token) {
  butuh(token, 'guru', 2);
  return baris('Guru').filter(function (r) { return String(r[1]).trim(); }).map(function (r) {
    return { id: String(r[0]), nama: String(r[1]), aktif: String(r[4]) === 'Ya', pin: pinNilai(r[8]) };
  });
}

function resetPinGuru(token, id) {
  var s = butuh(token, 'guru', 2);
  id = String(id || '');
  if (!cariBaris('Guru', 1, id)) return { ok: false, pesan: 'Guru tidak ditemukan' };
  tulisPin(id, PIN_DEF);
  CacheService.getScriptCache().remove('gagal_g_' + id); // sekaligus membuka kunci bila terkunci
  logAdmin(s.mode, 'Reset PIN guru: ' + (petaNama()[id] || id));
  return { ok: true, pesan: 'PIN direset menjadi 000000' };
}

// info absen untuk layar guru / Absen Saya: status hari ini + lokasi sekolah
function getInfoAbsen(token) {
  var s = sesi(token);
  if (!s) throw new Error('Sesi habis, silakan login ulang');
  if (!s.guruId) throw new Error('Akun Anda belum terhubung ke data guru');
  var tgl = tglStr(), ada = jamPunya(tgl, s.guruId), gi = infoGuru(s.guruId);
  return { nama: gi.nama, sapaan: gi.sapaan, pagi: ada.pagi, siang: ada.siang,
           piket: piketHari(s.guruId, tgl), lokasi: lokasiSekolah() };
}

// ---------- Rekap pribadi ----------
function getRekapSaya(token, bulan) {
  var s = sesi(token);
  if (!s || !s.guruId) throw new Error('Sesi habis, silakan login ulang');
  bulan = String(bulan || tglStr().slice(0, 7));
  if (!RE_BLN.test(bulan)) throw new Error('Format bulan salah');
  var r = rekapInti(bulan, s.guruId);
  var id = String(s.guruId), A = {}, sts = [], hari = [];
  baris('Absensi').forEach(function (x) {
    var t = normTgl(x[0]);
    if (t.slice(0, 7) === bulan && String(x[1]) === id) A[t] = x;
  });
  baris('Status Manual').forEach(function (x) {
    if (String(x[0]) === id) sts.push({ st: String(x[1]), dari: normTgl(x[2]), sampai: normTgl(x[3]) });
  });
  var y = parseInt(bulan.slice(0, 4), 10), mo = parseInt(bulan.slice(5), 10), n = new Date(y, mo, 0).getDate(), hi = tglStr();
  for (var d = n; d >= 1; d--) {
    var tgl = bulan + '-' + ('0' + d).slice(-2);
    if (tgl > hi) continue;
    var a = A[tgl], st = '';
    for (var i = 0; i < sts.length; i++) { if (sts[i].dari <= tgl && tgl <= sts[i].sampai) st = sts[i].st; }
    if (!a && !st) continue;
    var pg = a ? String(a[2] || '') : '', sg = a ? String(a[4] || '') : '';
    hari.push([tgl, pg, a ? String(a[3] || '') : '', sg, st || (a ? String(a[5] || '') : '')]);
  }
  return { head: r.head, row: r.rows[0] || [], hari: hari };
}

// ---------- Jadwal khusus tanggal & Jadwal Hari Ini ----------
function urutanJamOk(v) { // v: [pagiBuka, pagiTutup, telat, siangBuka, siangTutup]
  var m1 = menitDari(v[0]), m2 = menitDari(v[1]), mt = menitDari(v[2]), m3 = menitDari(v[3]), m4 = menitDari(v[4]);
  return m1 < m2 && m2 <= m3 && m3 < m4 && mt >= m1 && mt <= m2;
}

function khususList(jd) { // jadwal khusus hari ini dan mendatang
  var hi = tglStr();
  return Object.keys(jd.k || {}).filter(function (t) { return t >= hi; }).sort().map(function (t) {
    var h = jd.k[t];
    return { tgl: t, pagiBuka: h.pagiBuka, pagiTutup: h.pagiTutup, telat: h.telat, siangBuka: h.siangBuka, siangTutup: h.siangTutup, telatPiket: h.telatPiket };
  });
}

function simpanKhususInti(s, tgl, jam, tp) {
  var vals = [];
  jam = jam || {};
  for (var q = 0; q < JDW_KEYS.length; q++) {
    var t = String(jam[JDW_KEYS[q]] || '').trim();
    if (!RE_JAM.test(t)) return { ok: false, pesan: 'Isi kelima jam dengan format HH:MM' };
    vals.push(t);
  }
  if (!urutanJamOk(vals)) {
    return { ok: false, pesan: 'Urutan jam tidak masuk akal (pagi buka < batas terlambat <= pagi tutup <= siang buka < siang tutup)' };
  }
  tp = String(tp || '').trim();
  if (tp && !RE_JAM.test(tp)) return { ok: false, pesan: 'Format batas terlambat piket salah (HH:MM)' };
  if (tp && menitDari(tp) < menitDari(vals[0])) return { ok: false, pesan: 'Batas terlambat piket lebih awal dari jam pagi buka' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var r = cariBaris('Jadwal Khusus', 1, tgl), row = [tgl].concat(vals, [tp]);
    if (r) {
      var rg = sh('Jadwal Khusus').getRange(r, 1, 1, 7);
      rg.setNumberFormat('@');
      rg.setValues([row]);
    } else {
      appendTeks('Jadwal Khusus', row);
    }
    CacheService.getScriptCache().remove('lbk_' + tgl);
    hapusCacheCfg();
    logAdmin(s.mode, 'Jadwal khusus ' + tgl + ': ' + vals.join(' '));
    return { ok: true, pesan: 'Jadwal ' + tgl + ' disimpan' };
  } finally {
    lock.releaseLock();
  }
}

function hapusKhususInti(s, tgl) {
  var r = cariBaris('Jadwal Khusus', 1, tgl);
  if (!r) return { ok: true, pesan: 'Tidak ada jadwal khusus' };
  sh('Jadwal Khusus').deleteRow(r);
  CacheService.getScriptCache().remove('lbk_' + tgl);
  hapusCacheCfg();
  logAdmin(s.mode, 'Hapus jadwal khusus ' + tgl);
  return { ok: true, pesan: 'Kembali ke jadwal mingguan' };
}

// d: {tgl, jam:{pagiBuka,pagiTutup,telat,siangBuka,siangTutup}, telatPiket} (kartu Jadwal Absen, hari ini & mendatang)
function simpanJadwalKhusus(token, d) {
  var s = butuh(token, 'jadwal', 2);
  d = d || {};
  var tgl = String(d.tgl || '');
  if (!RE_TGL.test(tgl)) return { ok: false, pesan: 'Tanggal belum benar' };
  if (tgl < tglStr()) return { ok: false, pesan: 'Tanggal sudah lewat' };
  return simpanKhususInti(s, tgl, d.jam, d.telatPiket);
}

function hapusJadwalKhusus(token, tgl) {
  var s = butuh(token, 'jadwal', 2);
  tgl = String(tgl || '');
  if (!RE_TGL.test(tgl)) return { ok: false, pesan: 'Tanggal belum benar' };
  return hapusKhususInti(s, tgl);
}

// kartu Jadwal Hari Ini: hanya tanggal hari ini
function getJadwalHariIni(token) {
  butuh(token, 'hariini', 1);
  var now = new Date(), tgl = tglStr(now), w = hariKe(tgl);
  var jd = jadwalAll(), kh = (jd.k[tgl] && jd.k[tgl].lengkap) ? jd.k[tgl] : null;
  var base = (jd[w] && jd[w].lengkap) ? jd[w] : null;
  var pakai = kh || base, jam = null;
  if (pakai) { jam = {}; JDW_KEYS.forEach(function (k) { jam[k] = pakai[k]; }); }
  return { tgl: tgl, hari: HARI_JDW[w - 1] || 'Minggu', jam: jam, khusus: !!kh,
           telatPiket: pakai ? (pakai.telatPiket || '') : (jd[1].telatPiket || ''),
           libur: isLibur(tgl, now) ? 'Hari ini libur, jadwal tidak berlaku.' : '' };
}

function simpanJadwalHariIni(token, d) {
  var s = butuh(token, 'hariini', 2);
  d = d || {};
  var now = new Date(), tgl = tglStr(now);
  if (isLibur(tgl, now)) return { ok: false, pesan: 'Hari ini libur' };
  return simpanKhususInti(s, tgl, d.jam, d.telatPiket);
}

function hapusJadwalHariIni(token) {
  var s = butuh(token, 'hariini', 2);
  return hapusKhususInti(s, tglStr());
}

function bersihkanKhusus() { // buang jadwal khusus lebih dari 30 hari lalu
  try {
    var s = sh('Jadwal Khusus'), n = s.getLastRow();
    if (n < 2) return;
    var batas = tglStr(new Date(Date.now() - 30 * 86400000));
    var v = s.getRange(2, 1, n - 1, 1).getValues();
    for (var i = v.length - 1; i >= 0; i--) {
      if (normTgl(v[i][0]) < batas) s.deleteRow(i + 2);
    }
  } catch (e) { Logger.log('Bersih jadwal khusus dilewati: ' + e); }
}

// ---------- Reset absen (hari ini saja): hapus foto + jam + status satu sesi satu guru ----------
function getResetList(token) {
  butuh(token, 'reset', 1);
  var tgl = tglStr(), nama = petaNama(), out = [];
  baris('Absensi').forEach(function (r) {
    if (normTgl(r[0]) !== tgl) return;
    var id = String(r[1]);
    [['pagi', r[2], r[3], r[6]], ['siang', r[4], '', r[7]]].forEach(function (x) {
      var jam = String(x[1] || '');
      if (!jam || jam === PENANDA_LIBUR) return;
      var f = String(x[3] || '');
      out.push({ id: id, nama: String(nama[id] || id), sesi: x[0], jam: jam, st: String(x[2] || ''),
                 foto: f ? (f === 'kedaluwarsa' ? 'exp' : 'ada') : '' });
    });
  });
  out.sort(function (a, b) {
    var c = String(a.nama).localeCompare(String(b.nama), 'id');
    return c !== 0 ? c : (a.sesi < b.sesi ? 1 : -1); // pagi dulu
  });
  return out;
}

function getFotoReset(token) { // foto hari ini, termasuk untuk Wakil (khusus menu Reset Absen)
  butuh(token, 'reset', 1);
  return fotoHari(tglStr());
}

function hapusFileDrive(id) {
  try {
    CacheService.getScriptCache().remove('ft_' + id);
    if (typeof Drive !== 'undefined') Drive.Files.remove(id); // permanen, tanpa tempat sampah
    else DriveApp.getFileById(id).setTrashed(true);           // cadangan bila layanan Drive API belum ditambahkan
  } catch (e) { Logger.log('Hapus foto gagal: ' + e); }
}

// catat reset supaya tablet ikut menghapus tanda "sudah absen" di cache-nya
function catatReset(tgl, id, sesiAbsen) {
  var p = PropertiesService.getScriptProperties(), o = {};
  try { o = JSON.parse(p.getProperty('resetHariIni') || '{}'); } catch (e) { o = {}; }
  if (o.tgl !== tgl) o = { tgl: tgl, items: [] };
  o.items.push({ id: String(id), sesi: sesiAbsen, t: new Date().getTime() });
  p.setProperty('resetHariIni', JSON.stringify(o));
}

function getResetHariIni(devId) {
  if (!cekPerangkat(devId).ok) return [];
  try {
    var o = JSON.parse(PropertiesService.getScriptProperties().getProperty('resetHariIni') || '{}');
    return o.tgl === tglStr() ? (o.items || []) : [];
  } catch (e) { return []; }
}

function resetAbsen(token, guruId, sesiAbsen) {
  var s = butuh(token, 'reset', 2);
  guruId = String(guruId || '');
  if (sesiAbsen !== 'pagi' && sesiAbsen !== 'siang') return { ok: false, pesan: 'Sesi tidak dikenal' };
  var tgl = tglStr(), jam = '', fid = '';
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var r = cariAbsen(tgl, guruId);
    if (!r) return { ok: false, pesan: 'Data absen tidak ditemukan' };
    var v = sh('Absensi').getRange(r, 1, 1, 10).getValues()[0];
    jam = String(v[sesiAbsen === 'pagi' ? 4 : 6] || '');
    if (!jam || jam === PENANDA_LIBUR) return { ok: false, pesan: 'Tidak ada absen ' + sesiAbsen + ' untuk direset' };
    var m = String(v[sesiAbsen === 'pagi' ? 8 : 9] || '').match(/\/d\/([\w-]+)/);
    fid = m ? m[1] : '';
    if (sesiAbsen === 'pagi') {
      tulisAbsen(tgl, guruId, 3, '');
      tulisAbsen(tgl, guruId, 4, '');
      tulisAbsen(tgl, guruId, 7, '');
    } else {
      tulisAbsen(tgl, guruId, 5, '');
      tulisAbsen(tgl, guruId, 8, '');
    }
    bersihKetSesi(tgl, guruId, sesiAbsen);
    catatReset(tgl, guruId, sesiAbsen);
    CacheService.getScriptCache().remove('lbk_' + tgl); // penanda LIBUR guru yang libur diisi ulang bila perlu
  } finally {
    lock.releaseLock();
  }
  if (fid) hapusFileDrive(fid); // di luar kunci
  var nama = petaNama()[guruId] || guruId;
  logAdmin(s.mode, 'Reset absen ' + sesiAbsen + ' ' + tgl + ': ' + nama + ' (jam ' + jam + ')');
  return { ok: true, pesan: 'Absen ' + nama + ' (' + sesiAbsen + ') direset' };
}

// Jalankan SEKALI dari editor setelah menempel kode Ver. 007 (sheet yang sudah berjalan). Data TIDAK dihapus:
// menambah kolom PIN & Piket di tab Guru (PIN kosong jadi 000000), tab Jadwal Khusus, batas terlambat piket 07:15,
// baris Hak Akses baru, kalimat ucapan baru, dan kode rahasia link guru.
function migrasiV7() {
  var g = sh('Guru');
  pasangSapaan(g);
  pasangLibur(g);
  CacheService.getScriptCache().remove('pp_ok');
  pasangPinPiket(g);
  var n = g.getLastRow();
  if (n >= 2) {
    var nm = g.getRange(2, 2, n - 1, 1).getValues();
    var rg = g.getRange(2, 9, n - 1, 1), v = rg.getValues(), ubah = false;
    v.forEach(function (r, i) {
      if (!String(nm[i][0]).trim()) return;
      var p = pinNilai(r[0]);
      if (String(r[0]) !== p) { r[0] = p; ubah = true; }
    });
    if (ubah) { rg.setNumberFormat('@'); rg.setValues(v); }
  }
  var sj = sh('Jadwal');
  if (!RE_JAM.test(teksSel(sj.getRange(8, 2).getValue()))) {
    sj.getRange(8, 1, 1, 2).setNumberFormat('@').setValues([['Batas Terlambat Piket', '07:15']]);
  }
  sh('Jadwal Khusus');
  var h = sh('Hak Akses'), hn = h.getLastRow();
  if (hn >= 2) { // nama kartu berubah: Jadwal Absen Harian -> Jadwal Absen
    h.getRange(2, 1, hn - 1, 1).getValues().forEach(function (r, i) {
      if (teksSel(r[0]) === 'Jadwal Absen Harian') h.getRange(i + 2, 1).setValue('Jadwal Absen');
    });
  }
  tambahBarisHakAkses();
  var uc = ucapanAll(), ada = {}, tot = uc.pagi.length + uc.siang.length, tambah = 0;
  uc.pagi.forEach(function (t) { ada['pagi|' + t] = 1; });
  uc.siang.forEach(function (t) { ada['siang|' + t] = 1; });
  UC_DEF.forEach(function (x) {
    var k = String(x[0]).toLowerCase();
    if (ada[k + '|' + x[1]] || tot >= UC_N) return;
    uc[k].push(x[1]); tot++; tambah++;
  });
  if (tambah) simpanUcapan(uc);
  linkKode();
  hapusCacheCfg();
  hapusCacheGuru();
  logAdmin('sistem', 'Migrasi Ver. 007: PIN & Piket guru, Jadwal Khusus, Hak Akses, ucapan, link guru');
  Logger.log('Migrasi Ver. 007 selesai. Kalimat ucapan ditambah: ' + tambah + '. Atur Lokasi Sekolah dan salin Link Guru dari menu di aplikasi.');
}
