(function () {
'use strict';
/* ════════ Inti: state, util, RPC, modal, toast ════════ */
var INIT = (function () {
  /* Petunjuk awal dari alamat (?app=dashboard|kartu, ?aset=ID dari QR). Token sesi tidak pernah ada di alamat. */
  try {
    var q = new URLSearchParams(location.search), a = q.get('app'), s = q.get('aset');
    return { app: (a === 'dashboard' || a === 'kartu') ? a : '', aset: /^[A-Za-z0-9_\-]{1,40}$/.test(s || '') ? s : '' };
  } catch (e) { return {}; }
})();
var S = {
  token: null, user: null, cfg: { appName: 'SIGAP TANGGAP - Sistem Informasi Gedung Aman Preventif berbasis Risiko', logoUrl: '', webAppUrl: '' },
  master: { data: {}, lokasi: [] }, pic: [], notif: null, aset: [], app: 'dashboard', page: 'beranda', charts: [],
  f: { q: '', kat: '', ged: '', kon: '', kri: '', sta: '', urut: '', hal: 1, per: 25 },
  asetView: 'tabel', sel: {}, kPage: 'beranda', kAset: null, kTab: 'info', theme: 'light', adminTab: 'pengguna', adminKel: 'Kategori Aset',
  lbl: { kat: '', ged: '', lan: '', mt: 10, mx: 8, gap: 3, mode: 'zt411', W: 80, H: 24, mL: 2.5, mR: 2.5, mT: 2.5, mB: 2.5, ox: 0, oy: 0, salinan: 1 }, sesiHabis: false
};
var KELOMPOK = ['Kategori Aset', 'Kondisi Aset', 'Kritikalitas', 'Status Aset', 'Frekuensi PM', 'Pelaksana', 'Nama Teknisi', 'Nama Vendor', 'Verifikator', 'Jabatan PIC',
  'Sumber WO', 'Prioritas', 'Status WO', 'Status SLA', 'Jenis PM', 'Prioritas Perawatan', 'Status PM', 'Alasan Penundaan', 'Kategori Risiko', 'Level Risiko', 'Level Inherent',
  'Status Risiko', 'Kategori Bahaya', 'APD Wajib', 'Status Temuan', 'Target KPI'];
var KELOMPOK_NAMA = ['Nama Teknisi', 'Nama Vendor', 'Verifikator', 'Jabatan PIC'];
var PERAN = ['Master Admin', 'Kabag', 'Kasubbag', 'SPV', 'PIC', 'Teknisi'];
var PERAN_DASH = ['Master Admin', 'Kabag', 'Kasubbag', 'SPV', 'PIC'];
var DIHAPUS = 'Dihapus/Tidak Aktif';
var CDN_XLSX = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
var CDN_QR = 'https://cdn.jsdelivr.net/npm/html5-qrcode@2.3.8/html5-qrcode.min.js';

var $ = function (s, r) { return (r || document).querySelector(s); };
var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
var SS = {
  get: function (k) { try { return sessionStorage.getItem(k); } catch (e) { return null; } },
  set: function (k, v) { try { sessionStorage.setItem(k, v); } catch (e) {} },
  del: function (k) { try { sessionStorage.removeItem(k); } catch (e) {} }
};
var LS = {
  get: function (k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) {} },
  del: function (k) { try { localStorage.removeItem(k); } catch (e) {} }
};
function isAdmin() { return S.user && S.user.peran === 'Master Admin'; }
function bolehDash() { return S.user && PERAN_DASH.indexOf(S.user.peran) >= 0; }
function bisaLihatNilai() { return S.user && ['Master Admin', 'Kabag', 'Kasubbag'].indexOf(S.user.peran) >= 0; }
function bolehLabel() { return S.user && ['Master Admin', 'Kabag', 'Kasubbag'].indexOf(S.user.peran) >= 0; }
function fmtRp(n) { return (n === '' || n == null || isNaN(Number(n))) ? '—' : 'Rp ' + Number(n).toLocaleString('id-ID'); }
function fmtAngka(n) { return Number(n || 0).toLocaleString('id-ID'); }
function listMaster(k) { return ((S.master.data || {})[k] || []).map(function (x) { return x.nilai; }); }
function uniq(a) { var s = {}, o = []; a.forEach(function (x) { if (x && !s[x]) { s[x] = 1; o.push(x); } }); return o; }
function debounce(fn, ms) { var t; return function () { var a = arguments, c = this; clearTimeout(t); t = setTimeout(function () { fn.apply(c, a); }, ms); }; }
function inisial(n) { return (String(n || '?').trim().charAt(0) || '?').toUpperCase(); }

/* ── Tema ── */
function terapkanTema(t) {
  S.theme = t === 'dark' ? 'dark' : 'light';
  document.documentElement.setAttribute('data-theme', S.theme);
  document.documentElement.setAttribute('data-bs-theme', S.theme);
  var m = document.querySelector('meta[name=theme-color]'); if (m) m.setAttribute('content', S.theme === 'dark' ? '#0B1220' : '#1E5EFF');
  $$('[data-act=toggleTheme] i').forEach(function (i) { i.className = 'bi ' + (S.theme === 'dark' ? 'bi-sun' : 'bi-moon-stars') + (i.parentNode.classList.contains('btn-link-x') ? ' me-1' : ''); });
}
function temaAwal() {
  var t = LS.get('sma_theme');
  if (!t) { try { t = window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'; } catch (e) { t = 'light'; } }
  terapkanTema(t);
}

/* ── Toast & konfirmasi ── */
function toast(msg, tipe, ms) {
  var el = document.createElement('div');
  el.className = 'toast-x ' + (tipe || '');
  var ic = { ok: 'bi-check-circle-fill text-success', bad: 'bi-x-circle-fill text-danger', warn: 'bi-exclamation-triangle-fill text-warning' }[tipe] || 'bi-info-circle-fill text-primary';
  el.innerHTML = '<i class="bi ' + ic + '"></i><div class="flex-grow-1">' + esc(msg) + '</div>';
  $('#toasts').appendChild(el);
  setTimeout(function () { el.remove(); }, ms || (tipe === 'bad' ? 6500 : 3500));
}
function konfirm(judul, isi, oke, bahaya) {
  return new Promise(function (res) {
    var box = $('#konfirm'); $('#kfJudul').textContent = judul; $('#kfIsi').textContent = isi;
    var ok = $('#kfOk'), bt = $('#kfBatal'); ok.textContent = oke || 'Ya';
    ok.className = 'btn ' + (bahaya === false ? 'btn-primary' : 'btn-danger');
    box.classList.remove('d-none');
    function tutup(v) { box.classList.add('d-none'); ok.onclick = bt.onclick = null; res(v); }
    ok.onclick = function () { tutup(true); }; bt.onclick = function () { tutup(false); };
    bt.focus();
  });
}

/* ── Modal tunggal ── */
var mdlInst = null;
function bukaModal(o) {
  $('#mdlDlg').className = 'modal-dialog modal-dialog-centered modal-dialog-scrollable' + (o.size ? ' modal-' + o.size : '');
  $('#mdlContent').innerHTML =
    '<div class="modal-header"><h2 class="modal-title">' + esc(o.judul || '') + '</h2><button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Tutup"></button></div>' +
    '<div class="modal-body">' + (o.isi || '') + '</div>' + (o.footer ? '<div class="modal-footer">' + o.footer + '</div>' : '');
  var el = $('#mdl');
  mdlInst = bootstrap.Modal.getOrCreateInstance(el, { backdrop: 'static' });
  mdlInst.show();
}
function tutupModal() { if (mdlInst) mdlInst.hide(); }
function sibuk(btn, on, teks) {
  if (!btn) return;
  if (on) { btn.dataset.t = btn.innerHTML; btn.disabled = true; btn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>' + esc(teks || 'Memproses…'); }
  else { btn.disabled = false; if (btn.dataset.t != null) btn.innerHTML = btn.dataset.t; }
}

/* ── RPC: fetch ke API Google Apps Script (Api.gs) ── */
function alamatGas() { return (window.SIGAP_CONFIG && window.SIGAP_CONFIG.GAS_URL) || ''; }
function rpc(fn) {
  var args = Array.prototype.slice.call(arguments, 1);
  var url = alamatGas();
  if (!url || /XXXX/.test(url) || !/^https:\/\//.test(url)) return Promise.reject(new Error('Alamat server belum diisi. Admin: isi GAS_URL di berkas config.js.'));
  var baca = /^get/.test(fn); // hanya pembacaan yang boleh diulang otomatis
  function kirim(ulang) {
    var ctl = typeof AbortController !== 'undefined' ? new AbortController() : null;
    var timer = ctl ? setTimeout(function () { ctl.abort(); }, 120000) : null;
    return fetch(url, {
      method: 'POST',
      // WAJIB text/plain: tipe lain memicu preflight CORS yang tidak didukung Apps Script
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ fn: fn, args: args }),
      redirect: 'follow',
      signal: ctl ? ctl.signal : undefined
    }).then(function (r) { if (timer) clearTimeout(timer); return r.text(); })
      .then(function (txt) {
        try { return JSON.parse(txt); } catch (e) { var g = new Error('Server tidak mengembalikan data yang valid. Coba lagi sebentar lagi.'); g.ulang = true; throw g; }
      })
      .catch(function (e) {
        if (timer) clearTimeout(timer);
        var batas = e && e.name === 'AbortError';
        if (baca && !ulang && !batas) return kirim(true);
        if (e && e.ulang) throw e;
        throw new Error(batas ? 'Server terlalu lama merespons. Coba lagi.' : 'Gagal terhubung ke server. Periksa koneksi internet.');
      });
  }
  return kirim(false).then(function (x) {
    if (!x) throw new Error('Respon server kosong.');
    if (x.success) return x.data;
    if (x.code === 'AUTH') sesiBerakhir(x.message);
    var e = new Error(x.message || 'Terjadi kesalahan.'); e.code = x.code; throw e;
  });
}
var FN_TULIS = /^(laporWo|mulaiWo|mulaiSlaWo|laporLangsungWo|saveWo|deleteWo|laporPm|mulaiPm|laporLangsungPm|savePm|deletePm|saveRisk|deleteRisk|updateStatusRisk)$/;
/** Tambalan lokal (optimistik) setelah aksi tulis berhasil, agar daftar & kartu darurat langsung benar sebelum data server tiba. */
function tambalLokal(fn, a) {
  try {
    var id = a[0], p = a[1] || {}, st = null;
    if (fn === 'mulaiWo' || fn === 'mulaiPm') st = 'Sedang Dikerjakan';
    else if (fn === 'laporWo' || fn === 'laporPm') st = p.Status_Hasil === 'Ditunda' ? 'Ditunda' : (fn === 'laporWo' ? 'Selesai' : 'Terlaksana');
    if (!st) return;
    var isWo = /Wo$/.test(fn), arr = isWo ? WO.list : PM.list, kid = isWo ? 'ID_WO' : 'ID_PM', ks = isWo ? 'Status_WO' : 'Status_PM';
    arr.forEach(function (x) { if (x[kid] === id) x[ks] = st; });
    if (isWo) S.darurat = (S.darurat || []).filter(function (x) { return x.ID_WO !== id || st !== 'Selesai'; }).map(function (x) { if (x.ID_WO === id) x.Status_WO = st; return x; });
  } catch (e) { /* abaikan */ }
}
function call(fn) {
  var a = Array.prototype.slice.call(arguments, 1);
  var p = rpc.apply(null, [fn, S.token].concat(a));
  if (FN_TULIS.test(fn)) p.then(function () { tambalLokal(fn, a); S.kartuTs = 0; if (S.app === 'kartu') setTimeout(function () { muatKartu(true); }, 0); }, function () {});
  return p;
}
function gagal(e) { if (e && e.code === 'AUTH') return; toast(e && e.message ? e.message : 'Terjadi kesalahan.', 'bad'); }

/* ── Tampilan ── */
function tampil(id) {
  pasangLogo();
  ['v-loading', 'v-login', 'v-pilih', 'v-dash', 'v-kartu'].forEach(function (v) { $('#' + v).classList.toggle('d-none', v !== id); });
}
function sesiBerakhir(msg) {
  if (S.sesiHabis) return; S.sesiHabis = true;
  LS.del('sma_token'); S.token = null; S.user = null; matikanScan && matikanScan();
  hancurkanChart(); tutupModal(); tampil('v-login');
  var e = $('#lgErr'); e.textContent = msg || 'Sesi berakhir. Silakan login kembali.'; e.classList.remove('d-none');
  setTimeout(function () { S.sesiHabis = false; }, 500);
}

/* ── Pencarian pintar ──
   Kata kunci dipecah per kata; sebuah data tampil bila SATU kata atau lebih cocok (awalan kata, bagian kata, atau salah ketik ringan),
   diurutkan dari yang paling banyak & paling tepat cocok. Contoh: "AC Medvis" menemukan "AC Split Ruang Medvis". */
var CARI_STOP = { di: 1, ke: 1, dari: 1, dan: 1, yang: 1, untuk: 1, pada: 1, the: 1, of: 1 };
function cariNorm(t) { var s = String(t === null || t === undefined ? '' : t).toLowerCase(); try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { /* peramban lama */ } return s.replace(/[^a-z0-9]+/g, ' ').trim(); }
function cariToken(q) {
  var t = cariNorm(q).split(' ').filter(Boolean), u = [], seen = {};
  t.forEach(function (x) { if (!seen[x]) { seen[x] = 1; u.push(x); } });
  var bersih = u.filter(function (x) { return !CARI_STOP[x]; });
  return bersih.length ? bersih : u;
}
function cariLev(a, b, maks) {
  if (Math.abs(a.length - b.length) > maks) return maks + 1;
  var p = [], i, j, c, m;
  for (j = 0; j <= b.length; j++) p[j] = j;
  for (i = 1; i <= a.length; i++) {
    var prev = p[0]; p[0] = i; m = p[0];
    for (j = 1; j <= b.length; j++) { var tmp = p[j]; c = a.charCodeAt(i - 1) === b.charCodeAt(j - 1) ? 0 : 1; p[j] = Math.min(p[j] + 1, p[j - 1] + 1, prev + c); prev = tmp; if (p[j] < m) m = p[j]; }
    if (m > maks) return maks + 1;
  }
  return p[b.length];
}
/** Skor kecocokan satu data: { n: jumlah kata cocok, s: skor }. */
function cariSkor(tokens, teks) {
  var h = ' ' + cariNorm(teks) + ' ', kata = null, n = 0, skor = 0;
  tokens.forEach(function (t) {
    var v = 0;
    if (h.indexOf(' ' + t + ' ') >= 0) v = 4;
    else if (h.indexOf(' ' + t) >= 0) v = 3;
    else if (t.length >= 2 && h.indexOf(t) >= 0) v = 2;
    else if (t.length >= 4) {
      if (!kata) kata = h.trim().split(' ');
      var mk = t.length >= 8 ? 2 : 1;
      for (var i = 0; i < kata.length; i++) { var w = kata[i]; if (w.length >= 3 && (cariLev(t, w, mk) <= mk || cariLev(t, w.slice(0, t.length), mk) <= mk)) { v = 1; break; } }
    }
    if (v) { n++; skor += v; }
  });
  return { n: n, s: skor };
}
/** Saring + urutkan daftar menurut relevansi. q kosong → daftar apa adanya. teksFn(item) → teks yang dicari. */
function cariUrut(list, q, teksFn) {
  var tk = cariToken(q); if (!tk.length) return list;
  var hasil = [];
  list.forEach(function (x, i) { var r = cariSkor(tk, teksFn(x)); if (r.n) hasil.push({ x: x, n: r.n, s: r.s, i: i }); });
  hasil.sort(function (a, b) { return b.n - a.n || b.s - a.s || a.i - b.i; });
  return hasil.map(function (h) { return h.x; });
}
function cocokCari(q, teks) { var tk = cariToken(q); return !tk.length || cariSkor(tk, teks).n > 0; }
var TEKS_ASET = function (a) { return a.ID_Aset + ' ' + a.NUP + ' ' + a.Nama_Aset + ' ' + a.Merk_Type + ' ' + a.Kategori_Aset + ' ' + a.Gedung + ' ' + a.Lantai + ' ' + a.Ruangan + ' ' + a.Lokasi_Detail; };

/* ── Gambar ── */
/**
 * Foto Drive: tautan thumbnail drive.google.com sering gagal di Safari/iPhone (pelacakan lintas-situs memblokir cookie Google → ikon "?").
 * Karena itu foto ditampilkan lewat lh3.googleusercontent.com (tanpa pengalihan/cookie); foto yang baru diunggah memakai pratinjau lokal;
 * bila tetap gagal, foto diambil lewat server (getFotoData) sebagai data URI.
 */
var FOTO_LOKAL = {}, FOTO_LOKAL_N = 0, FOTO_CAD = {};
function idDrive(u) { var m = /[?&]id=([A-Za-z0-9_\-]{10,})|\/d\/([A-Za-z0-9_\-]{10,})/.exec(String(u || '')); return m ? (m[1] || m[2]) : ''; }
function imgU(u) {
  u = String(u || ''); if (!u) return '';
  if (FOTO_LOKAL[u]) return FOTO_LOKAL[u];
  if (u.indexOf('drive.google.com/thumbnail') < 0) return u;
  var id = idDrive(u), w = /[&?]sz=w(\d+)/.exec(u);
  return id ? 'https://lh3.googleusercontent.com/d/' + id + '=w' + (w ? w[1] : 1000) : u;
}
function catatFotoLokal(u, preview) { if (!u || !preview) return; if (FOTO_LOKAL_N > 40) { FOTO_LOKAL = {}; FOTO_LOKAL_N = 0; } FOTO_LOKAL[u] = preview; FOTO_LOKAL_N++; }
function cadanganFoto(id) {
  if (!FOTO_CAD[id]) FOTO_CAD[id] = call('getFotoData', id).then(function (d) { return d.data; });
  return FOTO_CAD[id];
}
function fotoHtml(a, cls) {
  if (a && a.Foto_Aset_URL) return '<img class="' + (cls || 'thumb') + '" src="' + esc(imgU(a.Foto_Aset_URL)) + '" alt="" loading="lazy" referrerpolicy="no-referrer" data-fb="1">';
  return '<span class="' + (cls || 'thumb') + ' ph"><i class="bi bi-image"></i></span>';
}
function gantiFotoPh(t) { if (!t.dataset.fb || !t.parentNode) return; var s = document.createElement('span'); s.className = t.className + ' ph'; s.innerHTML = '<i class="bi bi-image"></i>'; t.replaceWith(s); }
document.addEventListener('error', function (e) {
  var t = e.target;
  if (!t || t.tagName !== 'IMG') return;
  var id = /googleusercontent\.com\/d\/|drive\.google\.com\/thumbnail/.test(t.src || '') ? idDrive(t.src) : '';
  if (id && !t.dataset.cad && S && S.token) {
    t.dataset.cad = '1';
    cadanganFoto(id).then(function (d) { t.src = d; }, function () { gantiFotoPh(t); });
    return;
  }
  gantiFotoPh(t);
}, true);

function kompresFoto(file) {
  return new Promise(function (res, rej) {
    if (!file) { rej(new Error('Foto tidak ditemukan.')); return; }
    var tp = String(file.type || '').toLowerCase();
    if (tp && tp.indexOf('image/') !== 0) { rej(new Error('Berkas harus berupa foto.')); return; }
    var url = null, img = new Image();
    var beres = function () { if (url) { try { URL.revokeObjectURL(url); } catch (e) {} url = null; } };
    img.onerror = function () { beres(); rej(new Error('Foto tidak dapat dibuka. Coba ambil ulang atau pilih foto lain.')); };
    img.onload = function () {
      try {
        var maks = 1280, w = img.naturalWidth || img.width, h = img.naturalHeight || img.height, k = Math.min(1, maks / Math.max(w, h));
        var c = document.createElement('canvas'); c.width = Math.max(1, Math.round(w * k)); c.height = Math.max(1, Math.round(h * k));
        var x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height);
        var d = c.toDataURL('image/jpeg', 0.8);
        beres();
        if (d.indexOf('data:image/jpeg') !== 0) throw new Error('x');
        res({ name: 'foto.jpg', mime: 'image/jpeg', base64: d.split(',')[1], preview: d });
      } catch (e) { beres(); rej(new Error('Foto tidak dapat diproses. Coba foto lain.')); }
    };
    try { url = URL.createObjectURL(file); img.src = url; }
    catch (e) { var fr = new FileReader(); fr.onerror = function () { rej(new Error('Gagal membaca foto.')); }; fr.onload = function () { img.src = fr.result; }; fr.readAsDataURL(file); }
  });
}

/* ── QR ── */
function qrSvg(teks) {
  var q = qrcode(0, 'M'); q.addData(String(teks)); q.make();
  var n = q.getModuleCount(), d = '';
  for (var r = 0; r < n; r++) for (var c = 0; c < n; c++) if (q.isDark(r, c)) d += 'M' + (c + 4) + ',' + (r + 4) + 'h1v1h-1z';
  var m = n + 8;
  return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + m + ' ' + m + '" shape-rendering="crispEdges" role="img" aria-label="Kode QR"><rect width="' + m + '" height="' + m + '" fill="#fff"/><path d="' + d + '" fill="#000"/></svg>';
}
function qrTeks(a) {
  if (a.QR_URL) return a.QR_URL;
  return S.cfg.webAppUrl ? S.cfg.webAppUrl + (S.cfg.webAppUrl.indexOf('?') >= 0 ? '&' : '?') + 'aset=' + encodeURIComponent(a.ID_Aset) : '';
}
function idDariQr(teks) {
  teks = String(teks || '').trim();
  var m = /[?&]aset=([A-Za-z0-9_\-]{1,40})/.exec(teks);
  if (m) return m[1];
  return /^[A-Za-z0-9_\-]{1,40}$/.test(teks) ? teks : '';
}

/* ── Badge ── */
function badge(t, k) { return t ? '<span class="badge-x ' + (k || '') + '">' + esc(t) + '</span>' : '<span class="text-muted">—</span>'; }
function bKondisi(v) { return badge(v, { 'Baik': 'ok', 'Dalam Pemeliharaan': 'info', 'Rusak Ringan': 'warn', 'Rusak Berat': 'bad' }[v] || ''); }
function bKrit(v) { return badge(v, { 'Tinggi': 'bad', 'Sedang': 'warn', 'Rendah': 'ok' }[v] || ''); }
function bStatus(v) { return badge(v, { 'Aktif': 'ok', 'Siaga': 'warn', 'Tidak Aktif': '', 'Dihapus/Tidak Aktif': 'bad' }[v] || ''); }
/** WO dari pengaduan (belum terhubung ke aset): lokasi = "Lokasi / Nama pelapor" sesuai sheet pengaduan. */
function lokPengaduan(a) { return a.ID_Pengaduan && !a.Ruangan && !a.Gedung ? [a.Lokasi_Pengaduan, a.Nama_Pelapor].filter(Boolean).join(' / ') : ''; }
function lokasiStr(a) { return lokPengaduan(a) || a.Ruangan || ''; } // kolom Lokasi di kartu/tabel = Ruangan saja; rincian ada di detail
/** PIC disimpan sebagai JABATAN; tampilannya "Nama - Jabatan" dari akun PIC yang memegang jabatan itu saat ini. */
function picLabel(j) {
  if (!j) return '';
  var p = (S.pic || []).filter(function (x) { return x.jabatan.toLowerCase() === String(j).toLowerCase(); })[0];
  return p ? p.label : String(j);
}
function picOpsi(sel, kosong) {
  var l = (S.pic || []).slice(), ada = l.some(function (x) { return x.jabatan === sel; });
  var h = '<option value="">' + esc(kosong || '— pilih PIC —') + '</option>';
  l.forEach(function (x) { h += '<option value="' + esc(x.jabatan) + '"' + (x.jabatan === sel ? ' selected' : '') + '>' + esc(x.label) + '</option>'; });
  if (sel && !ada) h += '<option value="' + esc(sel) + '" selected>' + esc(sel) + ' (belum ada pemegang)</option>';
  return h;
}
function cariAset(id) { id = String(id || '').toLowerCase(); return S.aset.filter(function (a) { return a.ID_Aset.toLowerCase() === id; })[0] || null; }
function opsi(arr, sel, kosong) {
  return (kosong != null ? '<option value="">' + esc(kosong) + '</option>' : '') +
    arr.map(function (x) { return '<option value="' + esc(x) + '"' + (x === sel ? ' selected' : '') + '>' + esc(x) + '</option>'; }).join('');
}
function unduhBlob(nama, blob) {
  var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nama; document.body.appendChild(a); a.click();
  setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
function muatSkrip(src, cek) {
  return new Promise(function (res, rej) {
    if (cek()) { res(); return; }
    var s = document.createElement('script'); s.src = src; s.async = true;
    s.onload = function () { cek() ? res() : rej(new Error('Pustaka gagal dimuat.')); };
    s.onerror = function () { rej(new Error('Pustaka tidak dapat diunduh (periksa koneksi internet).')); };
    document.head.appendChild(s);
  });
}
function amanSel(v) { v = String(v == null ? '' : v); return /^[=+\-@\t\r]/.test(v) ? "'" + v : v; }

/* CSV */
function parseCsv(teks) {
  teks = String(teks).replace(/^﻿/, '');
  var baris1 = teks.split(/\r?\n/, 1)[0] || '';
  var delim = (baris1.split(';').length > baris1.split(',').length) ? ';' : ',';
  var rows = [], row = [], cur = '', q = false;
  for (var i = 0; i < teks.length; i++) {
    var ch = teks[i];
    if (q) {
      if (ch === '"') { if (teks[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch;
    } else if (ch === '"') q = true;
    else if (ch === delim) { row.push(cur); cur = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && teks[i + 1] === '\n') i++; row.push(cur); cur = ''; rows.push(row); row = []; }
    else cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows.filter(function (r) { return r.some(function (c) { return String(c).trim() !== ''; }); });
}
function csvKeString(rows) {
  return '﻿' + rows.map(function (r) { return r.map(function (c) { return '"' + amanSel(c).replace(/"/g, '""') + '"'; }).join(','); }).join('\r\n');
}
function bacaFile(file, sebagaiTeks) {
  return new Promise(function (res, rej) {
    var fr = new FileReader(); fr.onerror = function () { rej(new Error('Gagal membaca file.')); }; fr.onload = function () { res(fr.result); };
    sebagaiTeks ? fr.readAsText(file, 'utf-8') : fr.readAsDataURL(file);
  });
}

var ACTIONS = {}, SUBMITS = {}, CHANGES = {};
var matikanScan = null;
function hancurkanChart() { S.charts.forEach(function (c) { try { c.destroy(); } catch (e) {} }); S.charts = []; }

/* ════════ Auth, boot, shell ════════ */
var LOGO_SIGAP = 'data:image/png;base64,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';
var NAMA_SIGAP = 'SIGAP TANGGAP - Sistem Informasi Gedung Aman Preventif berbasis Risiko';
function logoHtml() { return '<img src="' + LOGO_SIGAP + '" alt="Logo SIGAP">'; }
function pasangLogo() { ['loginLogo', 'pilihLogo', 'sbLogo', 'kLogo'].forEach(function (id) { var e = document.getElementById(id); if (e) e.innerHTML = logoHtml(); }); }
function terapkanBranding() {
  var nama = (S.cfg.appName && S.cfg.appName !== 'Sistem Manajemen Aset' && S.cfg.appName !== 'SIGAP - Sistem Informasi Gedung Aman Preventif') ? S.cfg.appName : NAMA_SIGAP;
  var bagi = nama.split(' - '), kepala = bagi[0], ekor = bagi.length > 1 ? bagi.slice(1).join(' - ') : '';
  document.title = nama; pasangLogo();
  $('#sbName').textContent = kepala; $('#kName').textContent = kepala;
  $('#sbSub').textContent = ekor; $('#kSub').textContent = ekor;
  var u = S.user || {};
  $('#avD').textContent = inisial(u.nama); $('#uNamaD').textContent = u.nama || '—'; $('#uPeranD').textContent = u.peran || '—';
}
async function muatData() {
  var d = await call('getBootstrapData');
  S.user = d.user; S.cfg = d.config; S.master = d.master; S.aset = d.aset; S.pic = d.pic || [];
  S.darurat = d.darurat || []; S.daruratTs = Date.now();
}
function tampilPilih() {
  pasangLogo();
  $('#pilihSapa').textContent = 'Halo, ' + ((S.user.nama || '').split(' ')[0] || 'pengguna');
  tampil('v-pilih'); terapkanTema(S.theme);
}
async function masuk(tanyaModul) {
  terapkanBranding();
  var app = INIT.app;
  if (INIT.aset) app = 'kartu';
  if (!app) {
    if (!bolehDash()) app = 'kartu';
    else {
      var ingat = SS.get('sma_modul');
      if (tanyaModul && !ingat) { tampilPilih(); return; }
      app = ingat === 'kartu' ? 'kartu' : (ingat === 'dashboard' ? 'dashboard' : 'dashboard');
    }
  }
  if (app === 'dashboard' && !bolehDash()) app = 'kartu';
  S.app = app;
  if (bolehDash() && !INIT.aset) SS.set('sma_modul', app);
  if (app === 'kartu') {
    tampil('v-kartu'); S.kPage = 'beranda'; renderKartu(); muatKartu(true);
    if (INIT.aset) { var id = INIT.aset; INIT.aset = ''; bukaAsetKartu(id); }
  } else {
    tampil('v-dash'); S.page = 'beranda'; buatMenu(); go('beranda'); muatNotif();
  }
  terapkanTema(S.theme);
}
async function boot() {
  temaAwal();
  S.token = LS.get('sma_token');
  if (!S.token) { tampil('v-login'); return; }
  try { await muatData(); S.sesiHabis = false; await masuk(true); }
  catch (e) {
    if (e.code !== 'AUTH') { LS.del('sma_token'); S.token = null; tampil('v-login'); var el = $('#lgErr'); el.textContent = e.message; el.classList.remove('d-none'); }
  }
}
$('#formLogin').addEventListener('submit', async function (ev) {
  ev.preventDefault();
  var err = $('#lgErr'); err.classList.add('d-none');
  var u = $('#lgUser').value.trim(), p = $('#lgPass').value, btn = $('#btnLogin');
  if (!u || !p) { err.textContent = 'Isi username dan password.'; err.classList.remove('d-none'); return; }
  sibuk(btn, true, 'Masuk…');
  try {
    var d = await rpc('login', u, p);
    S.token = d.token; LS.set('sma_token', d.token); $('#lgPass').value = '';
    S.sesiHabis = false; SS.del('sma_modul');
    await muatData(); await masuk(true);
  } catch (e) { err.textContent = e.message; err.classList.remove('d-none'); }
  finally { sibuk(btn, false); }
});

ACTIONS.togglePass = function (el) {
  var i = $('#lgPass'); i.type = i.type === 'password' ? 'text' : 'password';
  el.querySelector('i').className = 'bi ' + (i.type === 'password' ? 'bi-eye' : 'bi-eye-slash');
};
ACTIONS.toggleTheme = function () {
  LS.set('sma_theme', S.theme === 'dark' ? 'light' : 'dark'); terapkanTema(LS.get('sma_theme') || (S.theme === 'dark' ? 'light' : 'dark'));
  if (S.app === 'dashboard' && S.token && S.page === 'beranda') renderBeranda();
  else if (S.app === 'dashboard' && S.token && S.page === 'aset' && S.asetView === 'grafik') renderAsetGrafik();
};
ACTIONS.pilihDash = function () { SS.set('sma_modul', 'dashboard'); masuk(); };
ACTIONS.pilihKartu = function () { SS.set('sma_modul', 'kartu'); masuk(); };
ACTIONS.logout = async function () {
  SS.del('sma_modul');
  try { await rpc('logout', S.token); } catch (e) {}
  LS.del('sma_token'); S.token = null; S.user = null; matikanScan(); hancurkanChart(); S.sel = {};
  $('#lgErr').classList.add('d-none'); tampil('v-login');
};
ACTIONS.toKartu = function () { SS.set('sma_modul', 'kartu'); S.app = 'kartu'; hancurkanChart(); tampil('v-kartu'); S.kPage = 'beranda'; renderKartu(); muatKartu(true); };
ACTIONS.toDash = function () { if (!bolehDash()) return; SS.set('sma_modul', 'dashboard'); matikanScan(); S.app = 'dashboard'; tampil('v-dash'); buatMenu(); go(S.page || 'beranda'); };
ACTIONS.refreshAll = async function (el) {
  sibuk(el, true, '');
  try { await muatData(); terapkanBranding(); go(S.page); muatNotif(); toast('Data dimuat ulang.', 'ok'); } catch (e) { gagal(e); } finally { sibuk(el, false); }
};
ACTIONS.ubahPassword = function () {
  bukaModal({
    judul: 'Ganti password', size: 'sm',
    isi: '<form id="fPw" data-submit="ubahPw" novalidate>' +
      '<div class="mb-3"><label class="form-label req">Password saat ini</label><input name="lama" type="password" class="form-control" autocomplete="current-password" required></div>' +
      '<div class="mb-3"><label class="form-label req">Password baru</label><input name="baru" type="password" class="form-control" autocomplete="new-password" minlength="8" required><div class="form-text">Minimal 8 karakter.</div></div>' +
      '<div class="mb-3"><label class="form-label req">Ulangi password baru</label><input name="ulang" type="password" class="form-control" autocomplete="new-password" required></div>' +
      '<div class="alert-x danger d-none" id="pwErr"></div></form>',
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fPw">Simpan</button>'
  });
};
SUBMITS.ubahPw = async function (f) {
  var e = $('#pwErr'); e.classList.add('d-none');
  var lama = f.lama.value, baru = f.baru.value;
  if (baru.length < 8) { e.textContent = 'Password baru minimal 8 karakter.'; e.classList.remove('d-none'); return; }
  if (baru !== f.ulang.value) { e.textContent = 'Konfirmasi password tidak sama.'; e.classList.remove('d-none'); return; }
  var btn = $('button[type=submit][form=fPw]'); sibuk(btn, true);
  try { await call('changeOwnPassword', lama, baru); tutupModal(); toast('Password diubah. Silakan login ulang.', 'ok'); ACTIONS.logout(); }
  catch (er) { e.textContent = er.message; e.classList.remove('d-none'); } finally { sibuk(btn, false); }
};

/* ════════ Dashboard: menu & navigasi ════════ */
var SEMUA_DASH = ['Master Admin', 'Kabag', 'Kasubbag', 'SPV', 'PIC'];
var OPS = ['Kabag', 'Kasubbag', 'SPV', 'PIC'];
// Hak akses menu mengikuti PRD §3.2; menu yang tak dapat diakses sebuah peran tidak ditampilkan.
var MENU = [
  { sec: 'Utama' },
  { id: 'beranda', ic: 'bi-grid-1x2', t: 'Beranda', peran: SEMUA_DASH },
  { id: 'aset', ic: 'bi-box-seam', t: 'Data Aset', peran: SEMUA_DASH },
  { sec: 'Operasional' },
  { id: 'wo', ic: 'bi-clipboard-check', t: 'Work Order', peran: OPS },
  { id: 'pm', ic: 'bi-calendar-check', t: 'Preventive Maint.', peran: OPS },
  { id: 'risk', ic: 'bi-exclamation-diamond', t: 'Risk Register', peran: OPS },
  { id: 'k3', ic: 'bi-shield-check', t: 'K3 & Safety', peran: OPS },
  { id: 'sla', ic: 'bi-stopwatch', t: 'SLA', peran: ['Kabag', 'Kasubbag'] },
  { id: 'kpi', ic: 'bi-graph-up-arrow', t: 'KPI & Laporan', peran: ['Master Admin', 'Kabag', 'Kasubbag'] },
  { sec: 'Sistem' },
  { id: 'admin', ic: 'bi-gear', t: 'Master Admin', peran: ['Master Admin'] }
];
var JUDUL = { beranda: 'Beranda', aset: 'Data Aset', label: 'Data Aset · Label QR', wo: 'Work Order', pm: 'Preventive Maintenance', risk: 'Risk Register', k3: 'K3 & Safety', sla: 'SLA', kpi: 'KPI & Laporan', admin: 'Master Admin' };
function nbadgeMenu(id) {
  if (!isAdmin() || !S.notif) return '';
  var n = id === 'admin' ? notifUsulan() : (id === 'aset' ? S.notif.rusak : 0);
  return n ? '<span class="nbadge ms-auto">' + n + '</span>' : '';
}
function buatMenu() {
  var peran = S.user && S.user.peran, out = [], sec = null;
  MENU.forEach(function (m) {
    if (m.sec) { sec = m; return; }
    if (m.peran.indexOf(peran) < 0) return;
    if (sec) { out.push('<div class="sb-sec">' + esc(sec.sec) + '</div>'); sec = null; }
    var aktif = S.page === m.id || (m.id === 'aset' && S.page === 'label');
    out.push('<button class="sb-item' + (aktif ? ' active' : '') + '" data-act="go" data-p="' + m.id + '"' + (m.soon ? ' disabled title="Tersedia ' + (m.soon === 'Segera' ? 'segera' : 'pada ' + m.soon) + '"' : '') + '>' +
      '<i class="bi ' + m.ic + '"></i><span>' + esc(m.t) + '</span>' + (m.soon ? '<span class="soon">' + m.soon + '</span>' : '') + nbadgeMenu(m.id) + '</button>');
  });
  $('#sbNav').innerHTML = out.join('');
}
ACTIONS.go = function (el) { go(el.dataset.p); };
ACTIONS.openSidebar = function () { $('#sidebar').classList.add('open'); $('.sb-backdrop').classList.add('open'); };
ACTIONS.closeSidebar = function () { $('#sidebar').classList.remove('open'); $('.sb-backdrop').classList.remove('open'); };
function go(p) {
  hancurkanChart(); ACTIONS.closeSidebar();
  if (p === 'aset' && S.page !== 'label') S.asetView = 'tabel';
  if (p === 'aset' && S.page === 'label') S.asetView = 'tabel';
  if (p === 'wo') WO.view = 'tabel';
  if (p === 'pm') PM.view = 'tabel';
  if (p === 'risk') RK.view = 'tabel';
  S.page = p; buatMenu();
  $('#topTitle').textContent = JUDUL[p] || '';
  var m = $('#main');
  if (p === 'beranda') renderBeranda();
  else if (p === 'aset') renderAset();
  else if (p === 'wo' && OPS.indexOf(S.user.peran) >= 0) renderWo();
  else if (p === 'pm' && OPS.indexOf(S.user.peran) >= 0) renderPm();
  else if (p === 'risk' && OPS.indexOf(S.user.peran) >= 0) renderRk();
  else if (p === 'k3' && OPS.indexOf(S.user.peran) >= 0) renderK3();
  else if (p === 'sla' && ['Kabag', 'Kasubbag'].indexOf(S.user.peran) >= 0) renderSla();
  else if (p === 'kpi' && AN_KPI_PERAN.indexOf(S.user.peran) >= 0) renderKpi();
  else if (p === 'label' && bolehLabel()) renderLabel();
  else if (p === 'admin' && isAdmin()) renderAdmin();
  else { S.page = 'beranda'; renderBeranda(); }
  window.scrollTo(0, 0);
}

/* ════════ Beranda ════════ */
var PALET = ['#1E5EFF', '#16A34A', '#F59E0B', '#DC2626', '#8B5CF6', '#06B6D4', '#EC4899', '#64748B'];
function hitung(arr, kunci) {
  var o = {}; arr.forEach(function (a) { var k = a[kunci] || '(kosong)'; o[k] = (o[k] || 0) + 1; });
  return Object.keys(o).map(function (k) { return [k, o[k]]; }).sort(function (a, b) { return b[1] - a[1]; });
}
function kpiHtml(ic, kls, v, l) { return '<div class="card-x kpi"><div class="ic ' + kls + '"><i class="bi ' + ic + '"></i></div><div><div class="v">' + v + '</div><div class="l">' + esc(l) + '</div></div></div>'; }
function renderBeranda() {
  hancurkanChart();
  var all = S.aset.filter(function (a) { return a.Status_Aset !== DIHAPUS; });
  var baik = all.filter(function (a) { return a.Kondisi === 'Baik'; }).length;
  var rusak = all.filter(function (a) { return a.Kondisi === 'Rusak Ringan' || a.Kondisi === 'Rusak Berat'; }).length;
  var tinggi = all.filter(function (a) { return a.Kritikalitas === 'Tinggi'; }).length;
  var aktif = all.filter(function (a) { return a.Status_Aset === 'Aktif'; }).length;
  var perhatian = all.filter(function (a) { return a.Kondisi !== 'Baik'; }).sort(function (a, b) {
    var w = { 'Tinggi': 0, 'Sedang': 1, 'Rendah': 2 }; return (w[a.Kritikalitas] || 3) - (w[b.Kritikalitas] || 3);
  }).slice(0, 8);
  var html = '<div class="page-head"><div><h2>Halo, ' + esc((S.user.nama || '').split(' ')[0]) + '</h2><div class="text-muted small">Ringkasan aset gedung &amp; fasilitas pendukung</div></div></div>' +
    '<div class="kpi-grid">' +
    kpiHtml('bi-box-seam', '', fmtAngka(all.length), 'Total Aset') +
    kpiHtml('bi-check2-circle', 'ok', fmtAngka(aktif), 'Status Aktif') +
    kpiHtml('bi-heart-pulse', 'ok', all.length ? Math.round(baik / all.length * 100) + '%' : '—', 'Kondisi Baik') +
    kpiHtml('bi-wrench-adjustable', 'bad', fmtAngka(rusak), 'Rusak (Ringan/Berat)') +
    kpiHtml('bi-exclamation-triangle', 'warn', fmtAngka(tinggi), 'Kritikalitas Tinggi') + '</div>' +
    '<div class="kpi-grid" id="anKartu"><div class="skeleton" style="height:84px"></div><div class="skeleton" style="height:84px"></div><div class="skeleton" style="height:84px"></div></div>' +
    '<div class="card-x mb-3" id="anAlert"><div class="card-b"><div class="skeleton" style="height:90px"></div></div></div>';
  if (!all.length) {
    html += '<div class="card-x"><div class="empty"><i class="bi bi-box-seam"></i>Belum ada aset.' + (isAdmin() ? '<div class="mt-3"><button class="btn btn-primary" data-act="go" data-p="aset">Tambah aset pertama</button></div>' : '') + '</div></div>';
  } else {
    html += '<div class="grid-2">' +
      '<div class="card-x"><div class="card-h">Aset per Kategori</div><div class="card-b"><div class="chart-box"><canvas id="chKat" role="img" aria-label="Grafik aset per kategori"></canvas></div></div></div>' +
      '<div class="card-x"><div class="card-h">Kondisi Aset</div><div class="card-b"><div class="chart-box"><canvas id="chKon" role="img" aria-label="Grafik kondisi aset"></canvas></div></div></div>' +
      '<div class="card-x"><div class="card-h">Aset per Gedung</div><div class="card-b"><div class="chart-box"><canvas id="chGed" role="img" aria-label="Grafik aset per gedung"></canvas></div></div></div>' +
      '<div class="card-x"><div class="card-h">Perlu Perhatian <span class="badge-x ms-auto">Kondisi bukan Baik</span></div>' +
      (perhatian.length ? '<div class="tbl-wrap"><table class="tbl"><tbody>' + perhatian.map(function (a) {
        return '<tr><td><div class="d-flex align-items-center gap-2">' + fotoHtml(a) + '<div><div class="nm">' + esc(a.Nama_Aset) + '</div><div class="sub">' + esc(a.ID_Aset) + ' · ' + esc(lokasiStr(a)) + '</div></div></div></td><td class="acts">' + bKondisi(a.Kondisi) + '</td></tr>';
      }).join('') + '</tbody></table></div>' : '<div class="empty"><i class="bi bi-emoji-smile"></i>Semua aset dalam kondisi baik.</div>') + '</div></div>' +
      (OPS.indexOf(S.user.peran) >= 0 ? '<div class="card-x" id="rkBeranda" style="grid-column:1/-1"><div class="card-b"><div class="skeleton" style="height:140px"></div></div></div>' : '') +
      '<div class="card-x"><div class="card-h"><i class="bi bi-bar-chart"></i>Tren Pekerjaan Pemeliharaan <span class="badge-x ms-auto">klik batang untuk rincian</span></div><div class="card-b"><div class="chart-box"><canvas id="anTren" role="img" aria-label="Grafik tren pekerjaan preventive dan corrective per bulan"></canvas></div></div></div>' +
      '<div class="card-x"><div class="card-h"><i class="bi bi-pie-chart"></i>Status Work Order</div><div class="card-b"><div class="chart-box"><canvas id="anStatus" role="img" aria-label="Grafik status work order"></canvas></div><div id="anStatusChip"></div></div></div>' +
      '<div class="card-x"><div class="card-h"><i class="bi bi-calendar-week"></i>Jadwal PM 7 Hari ke Depan</div><div id="anJadwal"><div class="card-b"><div class="skeleton" style="height:120px"></div></div></div></div>' +
      '<div class="card-x"><div class="card-h"><i class="bi bi-speedometer2"></i>KPI Pemeliharaan <span class="badge-x ms-auto">3 bulan terakhir</span></div><div class="card-b" id="anKpi"><div class="skeleton" style="height:120px"></div></div></div>';
  }
  $('#main').innerHTML = html;
  if (OPS.indexOf(S.user.peran) >= 0) berandaRisiko();
  berandaAn();
  if (!all.length || typeof Chart === 'undefined') return;
  var gelap = S.theme === 'dark';
  Chart.defaults.color = gelap ? '#98A5BF' : '#6B7280'; Chart.defaults.font.family = "'Plus Jakarta Sans',sans-serif";
  var grid = gelap ? '#22304D' : '#E2E8F0';
  var kat = hitung(all, 'Kategori_Aset'), kon = hitung(all, 'Kondisi'), ged = hitung(all, 'Gedung');
  var wk = { 'Baik': '#16A34A', 'Dalam Pemeliharaan': '#1E5EFF', 'Rusak Ringan': '#F59E0B', 'Rusak Berat': '#DC2626' };
  S.charts.push(new Chart($('#chKat'), { type: 'bar', data: { labels: kat.map(function (x) { return x[0]; }), datasets: [{ data: kat.map(function (x) { return x[1]; }), backgroundColor: '#1E5EFF', borderRadius: 6 }] },
    options: { indexAxis: 'y', maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { precision: 0 }, grid: { color: grid } }, y: { grid: { display: false } } } } }));
  S.charts.push(new Chart($('#chKon'), { type: 'doughnut', data: { labels: kon.map(function (x) { return x[0]; }), datasets: [{ data: kon.map(function (x) { return x[1]; }), backgroundColor: kon.map(function (x, i) { return wk[x[0]] || PALET[i % 8]; }), borderWidth: 0 }] },
    options: { maintainAspectRatio: false, cutout: '62%', plugins: { legend: { position: 'bottom' } }, onHover: anCursor, onClick: function (e, els) { if (els.length) anDrill('aset', { kondisi: kon[els[0].index][0] }); } } }));
  S.charts.push(new Chart($('#chGed'), { type: 'bar', data: { labels: ged.map(function (x) { return x[0]; }), datasets: [{ data: ged.map(function (x) { return x[1]; }), backgroundColor: PALET[1], borderRadius: 6 }] },
    options: { maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { ticks: { precision: 0 }, grid: { color: grid } }, x: { grid: { display: false } } } } }));
}

/* ════════ Data Aset ════════ */
function asetTersaring() {
  var f = S.f, q = f.q.trim();
  var hasil0 = S.aset.filter(function (a) {
    if (f.sta ? a.Status_Aset !== f.sta : a.Status_Aset === DIHAPUS) return false;
    if (f.kat && a.Kategori_Aset !== f.kat) return false;
    if (f.ged && a.Gedung !== f.ged) return false;
    if (f.kon && a.Kondisi !== f.kon) return false;
    if (f.kri && a.Kritikalitas !== f.kri) return false;
    return true;
  });
  if (q) return cariUrut(hasil0, q, function (a) { return TEKS_ASET(a) + ' ' + a.Nomor_Seri + ' ' + picLabel(a.PIC); });
  return hasil0.sort(function (a, b) {
    if (f.urut === 'nama' || f.urut === 'nama-') { var c = String(a.Nama_Aset).localeCompare(String(b.Nama_Aset), 'id', { sensitivity: 'base', numeric: true }); return f.urut === 'nama' ? c : -c; }
    return a.ID_Aset < b.ID_Aset ? -1 : (a.ID_Aset > b.ID_Aset ? 1 : 0);
  });
}
function renderAset() {
  if (S.asetView === 'grafik') { renderAsetGrafik(); return; }
  var f = S.f, adm = isAdmin();
  $('#main').innerHTML =
    '<div class="page-head"><h2>Data Aset</h2><div class="ms-auto d-flex flex-wrap gap-2">' +
    '<button class="btn btn-light-x" data-act="ekspor"><i class="bi bi-file-earmark-excel me-1"></i>Ekspor Excel</button>' +
    '<button class="btn btn-light-x" data-act="lihatGrafik"><i class="bi bi-bar-chart-line me-1"></i>Lihat Grafik</button>' +
    (bolehLabel() ? '<button class="btn btn-light-x" data-act="labelDariPilihan"><i class="bi bi-qr-code me-1"></i>Cetak Label QR</button>' : '') +
    (adm ? '<button class="btn btn-light-x" data-act="impor"><i class="bi bi-upload me-1"></i>Impor CSV</button><button class="btn btn-primary" data-act="tambahAset"><i class="bi bi-plus-lg me-1"></i>Tambah Aset</button>' : '') +
    '</div></div>' +
    (adm ? '<div id="panelMenunggu"></div>' : '') + '<div class="kpi-grid" id="asetStat"></div>' +
    '<div class="card-x"><div class="toolbar">' +
    '<div class="grow input-icon"><i class="bi bi-search"></i><input class="form-control" data-f="q" placeholder="Cari ID, NUP, nama, merk, ruangan…" value="' + esc(f.q) + '" aria-label="Cari aset"></div>' +
    '<select class="form-select fsel" data-f="kat" aria-label="Kategori Aset">' + opsi(listMaster('Kategori Aset'), f.kat, 'Semua kategori') + '</select>' +
    '<select class="form-select fsel" data-f="ged" aria-label="Gedung">' + opsi(uniq(S.master.lokasi.map(function (l) { return l.G; })), f.ged, 'Semua gedung') + '</select>' +
    '</div><div class="tbl-wrap"><table class="tbl stack"><thead><tr>' + (bolehLabel() ? '<th style="width:36px"><input type="checkbox" class="chk" id="chkAll" aria-label="Pilih semua di halaman"></th>' : '') +
    '<th><button type="button" class="th-sort" data-act="urutAset" title="Urutkan nama aset A–Z">Aset <i class="bi ' + (f.urut === 'nama' ? 'bi-sort-alpha-down' : (f.urut === 'nama-' ? 'bi-sort-alpha-up' : 'bi-arrow-down-up')) + '"></i></button></th><th>Kategori</th><th>Lokasi</th><th>Kondisi</th><th>Kritikalitas</th><th>Status</th><th></th></tr></thead><tbody id="asetBody"></tbody></table></div>' +
    '<div class="pager" id="asetPager"></div></div>';
  isiTabelAset(); isiStatAset();
  if (adm) { isiMenunggu(); muatMenunggu(); }
}
function asetHidup() { return S.aset.filter(function (a) { return a.Status_Aset !== DIHAPUS; }); }
function isiStatAset() {
  var all = asetHidup(), el = $('#asetStat'); if (!el) return;
  var n = function (fn) { return fmtAngka(all.filter(fn).length); };
  el.innerHTML = kpiHtml('bi-box-seam', '', fmtAngka(all.length), 'Total Aset') +
    kpiHtml('bi-check2-circle', 'ok', n(function (a) { return a.Kondisi === 'Baik'; }), 'Kondisi Baik') +
    kpiHtml('bi-tools', '', n(function (a) { return a.Kondisi === 'Dalam Pemeliharaan'; }), 'Dalam Pemeliharaan') +
    kpiHtml('bi-exclamation-circle', 'warn', n(function (a) { return a.Kondisi === 'Rusak Ringan'; }), 'Rusak Ringan') +
    kpiHtml('bi-x-octagon', 'bad', n(function (a) { return a.Kondisi === 'Rusak Berat'; }), 'Rusak Berat') +
    kpiHtml('bi-exclamation-triangle', 'warn', n(function (a) { return a.Kritikalitas === 'Tinggi'; }), 'Total Aset Kritis');
}

/* ── Grafik Data Aset (halaman terpisah; tombol Kembali) ── */
ACTIONS.lihatGrafik = function () { S.asetView = 'grafik'; hancurkanChart(); renderAsetGrafik(); window.scrollTo(0, 0); };
ACTIONS.kembaliAset = function () { S.asetView = 'tabel'; S.page = 'aset'; go('aset'); };
function umurBucket(all) {
  var th = new Date().getFullYear(), b = [['0–5 th', 0], ['6–10 th', 0], ['11–15 th', 0], ['16–20 th', 0], ['> 20 th', 0], ['Tanpa tahun', 0]];
  all.forEach(function (a) {
    if (!a.Tahun_Perolehan) { b[5][1]++; return; }
    var u = th - Number(a.Tahun_Perolehan); if (u < 0) u = 0;
    b[u <= 5 ? 0 : u <= 10 ? 1 : u <= 15 ? 2 : u <= 20 ? 3 : 4][1]++;
  });
  return b.filter(function (x, i) { return i < 5 || x[1] > 0; });
}
function renderAsetGrafik() {
  hancurkanChart();
  var all = asetHidup();
  var kotak = function (id, judul, ket) { return '<div class="card-x"><div class="card-h">' + judul + (ket ? '<span class="badge-x ms-auto">' + ket + '</span>' : '') + '</div><div class="card-b"><div class="chart-box"><canvas id="' + id + '" role="img" aria-label="' + esc(judul) + '"></canvas></div></div></div>'; };
  $('#main').innerHTML = '<div class="page-head"><button class="btn btn-light-x" data-act="kembaliAset"><i class="bi bi-arrow-left me-1"></i>Kembali ke Data Aset</button><h2>Grafik Data Aset</h2></div>' +
    (all.length ? '<div class="grid-2">' + kotak('gKat', 'Distribusi Aset per Kategori') + kotak('gKon', 'Kondisi Aset') + kotak('gKrit', 'Aset Kritis', 'Kritikalitas Tinggi') + kotak('gUmur', 'Umur Aset', 'tahun berjalan − tahun perolehan') + '</div>'
      : '<div class="card-x"><div class="empty"><i class="bi bi-bar-chart-line"></i>Belum ada aset untuk digrafikkan.</div></div>');
  if (!all.length || typeof Chart === 'undefined') return;
  var gelap = S.theme === 'dark', grid = gelap ? '#22304D' : '#E2E8F0';
  Chart.defaults.color = gelap ? '#98A5BF' : '#6B7280'; Chart.defaults.font.family = "'Plus Jakarta Sans',sans-serif";
  var wk = { 'Baik': '#16A34A', 'Dalam Pemeliharaan': '#1E5EFF', 'Rusak Ringan': '#F59E0B', 'Rusak Berat': '#DC2626' };
  var kat = hitung(all, 'Kategori_Aset'), kon = hitung(all, 'Kondisi');
  S.charts.push(new Chart($('#gKat'), { type: 'bar', data: { labels: kat.map(function (x) { return x[0]; }), datasets: [{ data: kat.map(function (x) { return x[1]; }), backgroundColor: '#1E5EFF', borderRadius: 6 }] },
    options: { indexAxis: 'y', maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { precision: 0 }, grid: { color: grid } }, y: { grid: { display: false } } } } }));
  S.charts.push(new Chart($('#gKon'), { type: 'doughnut', data: { labels: kon.map(function (x) { return x[0]; }), datasets: [{ data: kon.map(function (x) { return x[1]; }), backgroundColor: kon.map(function (x, i) { return wk[x[0]] || PALET[i % 8]; }), borderWidth: 0 }] },
    options: { maintainAspectRatio: false, cutout: '62%', plugins: { legend: { position: 'bottom' } } } }));
  var kritis = all.filter(function (a) { return a.Kritikalitas === 'Tinggi'; }), kk = hitung(kritis, 'Kategori_Aset').map(function (x) { return x[0]; });
  var kondisiList = listMaster('Kondisi Aset');
  S.charts.push(new Chart($('#gKrit'), { type: 'bar', data: { labels: kk, datasets: kondisiList.map(function (k, i) {
      return { label: k, backgroundColor: wk[k] || PALET[i % 8], borderRadius: 4, data: kk.map(function (c) { return kritis.filter(function (a) { return a.Kategori_Aset === c && a.Kondisi === k; }).length; }) };
    }) }, options: { maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, ticks: { precision: 0 }, grid: { color: grid } } } } }));
  var um = umurBucket(all);
  S.charts.push(new Chart($('#gUmur'), { type: 'bar', data: { labels: um.map(function (x) { return x[0]; }), datasets: [{ data: um.map(function (x) { return x[1]; }), backgroundColor: PALET[4], borderRadius: 6 }] },
    options: { maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { ticks: { precision: 0 }, grid: { color: grid } }, x: { grid: { display: false } } } } }));
}

function isiTabelAset() {
  var f = S.f, data = asetTersaring(), total = data.length, hal = Math.max(1, Math.ceil(total / f.per));
  if (f.hal > hal) f.hal = hal;
  var mulai = (f.hal - 1) * f.per, tampilData = data.slice(mulai, mulai + f.per), adm = isAdmin(), lbl = bolehLabel();
  $('#asetBody').innerHTML = tampilData.length ? tampilData.map(function (a) {
    return '<tr>' + (lbl ? '<td><input type="checkbox" class="chk" data-sel="' + esc(a.ID_Aset) + '"' + (S.sel[a.ID_Aset] ? ' checked' : '') + ' aria-label="Pilih ' + esc(a.ID_Aset) + '"></td>' : '') +
      '<td data-l="Aset"><div class="d-flex align-items-center gap-2">' + fotoHtml(a) + '<div class="min-w-0"><div class="nm">' + esc(a.Nama_Aset) + '</div><div class="sub">' + esc(a.ID_Aset) + (a.NUP ? ' · NUP ' + esc(a.NUP) : '') + '</div></div></div></td>' +
      '<td data-l="Kategori">' + esc(a.Kategori_Aset) + '</td><td data-l="Lokasi">' + esc(lokasiStr(a)) + '</td>' +
      '<td data-l="Kondisi">' + bKondisi(a.Kondisi) + '</td><td data-l="Kritikalitas">' + bKrit(a.Kritikalitas) + '</td><td data-l="Status">' + bStatus(a.Status_Aset) + '</td>' +
      '<td class="acts"><button class="icon-btn" data-act="lihatAset" data-id="' + esc(a.ID_Aset) + '" title="Detail" aria-label="Detail ' + esc(a.ID_Aset) + '"><i class="bi bi-eye"></i></button>' +
      (adm ? '<button class="icon-btn" data-act="ubahAset" data-id="' + esc(a.ID_Aset) + '" title="Ubah" aria-label="Ubah ' + esc(a.ID_Aset) + '"><i class="bi bi-pencil"></i></button><button class="icon-btn text-danger" data-act="hapusAset" data-id="' + esc(a.ID_Aset) + '" title="Hapus" aria-label="Hapus ' + esc(a.ID_Aset) + '"><i class="bi bi-trash"></i></button>' : '') + '</td></tr>';
  }).join('') : '<tr><td colspan="8"><div class="empty"><i class="bi bi-search"></i>Tidak ada aset yang cocok.</div></td></tr>';
  $('#asetPager').innerHTML = '<div>' + (total ? (mulai + 1) + '–' + (mulai + tampilData.length) + ' dari ' + fmtAngka(total) + ' aset' : '0 aset') + (Object.keys(S.sel).length ? ' · <b>' + Object.keys(S.sel).length + ' dipilih</b> <button class="btn-link-x" data-act="resetPilih">reset</button>' : '') + '</div>' +
    '<div class="d-flex align-items-center gap-2"><button class="btn btn-light-x btn-sm" data-act="halaman" data-d="-1"' + (f.hal <= 1 ? ' disabled' : '') + '><i class="bi bi-chevron-left"></i></button><span>Hal. ' + f.hal + ' / ' + hal + '</span><button class="btn btn-light-x btn-sm" data-act="halaman" data-d="1"' + (f.hal >= hal ? ' disabled' : '') + '><i class="bi bi-chevron-right"></i></button></div>';
}
ACTIONS.urutAset = function () { S.f.urut = S.f.urut === '' ? 'nama' : (S.f.urut === 'nama' ? 'nama-' : ''); S.f.hal = 1; renderAset(); };
ACTIONS.halaman = function (el) { S.f.hal += Number(el.dataset.d); isiTabelAset(); };
ACTIONS.resetPilih = function () { S.sel = {}; isiTabelAset(); };
document.addEventListener('input', debounce(function (e) {
  var t = e.target; if (t.dataset && t.dataset.f === 'q' && S.page === 'aset') { S.f.q = t.value; S.f.hal = 1; isiTabelAset(); }
}, 200));
document.addEventListener('change', function (e) {
  var t = e.target;
  if (t.dataset && t.dataset.f && t.dataset.f !== 'q' && S.page === 'aset') { S.f[t.dataset.f] = t.value; S.f.hal = 1; isiTabelAset(); }
  else if (t.dataset && t.dataset.sel) { if (t.checked) S.sel[t.dataset.sel] = 1; else delete S.sel[t.dataset.sel]; isiTabelAset(); }
  else if (t.id === 'chkAll') { $$('[data-sel]').forEach(function (c) { c.checked = t.checked; if (t.checked) S.sel[c.dataset.sel] = 1; else delete S.sel[c.dataset.sel]; }); isiTabelAset(); }
});

/* ════════ Detail, form, hapus ════════ */
function dlHtml(a) {
  var r = function (k, v) { return '<dt>' + k + '</dt><dd>' + (v === '' || v == null ? '<span class="text-muted">—</span>' : v) + '</dd>'; };
  return '<dl class="dl">' + r('ID Aset', esc(a.ID_Aset)) + r('NUP', esc(a.NUP)) + r('Kategori', esc(a.Kategori_Aset)) +
    r('Gedung', esc(a.Gedung)) + r('Lantai', esc(a.Lantai)) + r('Ruangan', esc(a.Ruangan)) +
    r('Tahun Perolehan', a.Tahun_Perolehan || '') + r('Merk / Type', esc(a.Merk_Type)) + r('Nomor Seri', esc(a.Nomor_Seri)) +
    r('Spesifikasi', esc(a.Spesifikasi).replace(/\n/g, '<br>')) + (bisaLihatNilai() ? r('Nilai Perolehan', a.Nilai_Perolehan === '' ? '' : fmtRp(a.Nilai_Perolehan)) : '') +
    r('Kondisi', bKondisi(a.Kondisi)) + r('Kritikalitas', bKrit(a.Kritikalitas)) + r('Status', bStatus(a.Status_Aset)) +
    r('PIC', esc(picLabel(a.PIC))) + r('Frekuensi PM', esc(a.Frekuensi_PM)) + '</dl>';
}
function fotoBesar(a) {
  return '<div class="foto-box">' + (a.Foto_Aset_URL ? '<img src="' + esc(imgU(a.Foto_Aset_URL)) + '" alt="Foto ' + esc(a.Nama_Aset) + '" data-act="lb" data-src="' + esc(a.Foto_Aset_URL) + '" referrerpolicy="no-referrer" data-fb="1">' : '<div class="text-center"><i class="bi bi-image fs-1"></i><div class="small">Belum ada foto</div></div>') + '</div>';
}
function qrBox(a) {
  var t = qrTeks(a);
  return t ? '<div class="qr-box">' + qrSvg(t) + '</div>' + (bolehLabel() ? '<div class="mt-2"><button class="btn btn-light-x btn-sm" data-act="unduhQr" data-id="' + esc(a.ID_Aset) + '"><i class="bi bi-download me-1"></i>Unduh QR (JPG)</button></div>' : '') : '<div class="alert-x warn small mb-0">URL Web App belum diatur (Master Admin → Pengaturan), QR belum dapat dibuat.</div>';
}
ACTIONS.unduhQr = function (el) {
  var a = cariAset(el.dataset.id); if (!a) return;
  var t = qrTeks(a); if (!t) { toast('URL Web App belum diatur, QR belum dapat dibuat.', 'bad'); return; }
  var svg = qrSvg(t).replace('<svg ', '<svg width="640" height="640" ');
  var img = new Image();
  img.onload = function () {
    var c = document.createElement('canvas'); c.width = 720; c.height = 800;
    var x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, 720, 800);
    x.drawImage(img, 40, 20, 640, 640);
    x.fillStyle = '#000'; x.textAlign = 'center';
    x.font = 'bold 42px "Courier New",monospace'; x.fillText(a.ID_Aset, 360, 722);
    x.font = '26px Arial,sans-serif'; var nm = a.Nama_Aset.length > 42 ? a.Nama_Aset.slice(0, 41) + '…' : a.Nama_Aset; x.fillText(nm, 360, 766);
    c.toBlob(function (b) { if (b) { unduhBlob('QR_' + a.ID_Aset + '.jpg', b); toast('QR ' + a.ID_Aset + ' diunduh.', 'ok'); } else toast('Gagal membuat gambar QR.', 'bad'); }, 'image/jpeg', 0.95);
  };
  img.onerror = function () { toast('Gagal membuat gambar QR.', 'bad'); };
  img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
};
ACTIONS.lb = function (el) { $('#lbImg').removeAttribute('data-cad'); $('#lbImg').src = imgU(el.dataset.src); $('#lbDl').href = el.dataset.src.replace('/thumbnail?id=', '/uc?export=download&id=').replace(/&sz=w\d+/, ''); $('#lightbox').classList.remove('d-none'); };
ACTIONS.lbClose = function () { $('#lightbox').classList.add('d-none'); $('#lbImg').removeAttribute('src'); };
document.addEventListener('keydown', function (e) { if (e.key === 'Escape') ACTIONS.lbClose(); });

ACTIONS.lihatAset = async function (el) {
  var a = cariAset(el.dataset.id); if (!a) return;
  bukaModal({
    judul: a.Nama_Aset, size: 'lg',
    isi: '<div class="row g-3"><div class="col-md-5">' + fotoBesar(a) + '<div class="text-center mt-3">' + qrBox(a) + '<div class="small text-muted mt-1">' + esc(a.ID_Aset) + '</div></div></div>' +
      '<div class="col-md-7">' + dlHtml(a) + '<hr><div class="fw-bold mb-2">Riwayat</div><div id="riw" class="small text-muted">Memuat…</div></div></div>',
    footer: (bolehLabel() ? '<button class="btn btn-light-x me-auto" data-act="labelSatu" data-id="' + esc(a.ID_Aset) + '"><i class="bi bi-qr-code me-1"></i>Label QR</button>' : '') +
      (isAdmin() ? '<button class="btn btn-light-x" data-act="ubahAset" data-id="' + esc(a.ID_Aset) + '"><i class="bi bi-pencil me-1"></i>Ubah</button>' : '') +
      '<button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>'
  });
  try {
    var d = await call('getAssetDetail', a.ID_Aset), r = $('#riw'); if (!r) return;
    r.innerHTML = riwayatHtml(d, a.ID_Aset);
  } catch (e) { var r2 = $('#riw'); if (r2) r2.textContent = e.message; }
};

/** Riwayat WO / PM sebagai kartu per laporan; kartu WO dapat diklik untuk melihat progres laporannya. */
function riwayatHtml(d, idAset) {
  var wo = (d.riwayatWO || []).slice().sort(function (a, b) { return String(b.Waktu_Dibuat) < String(a.Waktu_Dibuat) ? -1 : 1; });
  var pm = (d.riwayatPM || []).slice().sort(function (a, b) { return String(b.Tanggal_Jadwal) < String(a.Tanggal_Jadwal) ? -1 : 1; });
  if (!wo.length && !pm.length) return '<span class="text-muted">Belum ada riwayat Work Order / Preventive Maintenance.</span>';
  var h = '';
  if (wo.length) h += '<div class="riw-judul">Work Order · ' + wo.length + '</div>' + wo.map(function (x) {
    return '<button type="button" class="riw-card" data-act="riwayatWo" data-id="' + esc(x.ID_WO) + '" data-aset="' + esc(idAset) + '">' +
      '<div class="d-flex flex-wrap gap-1 align-items-center"><b>' + esc(x.ID_WO) + '</b>' + bStatusWo(x.Status_WO) + (x.Status_SLA ? badge(x.Status_SLA, x.Status_SLA === 'Terlambat' ? 'bad' : 'ok') : '') + '</div>' +
      '<div class="riw-uraian">' + esc(x.Uraian_Pekerjaan || '—') + '</div>' +
      '<div class="riw-meta"><i class="bi bi-calendar-event me-1"></i>' + esc(fmtTgl(x.Tanggal_Pengerjaan || x.Waktu_Dibuat, true)) + (x.Prioritas ? ' · ' + esc(x.Prioritas) : '') + '<i class="bi bi-chevron-right ms-auto"></i></div></button>';
  }).join('');
  if (pm.length) h += '<div class="riw-judul">Preventive Maintenance · ' + pm.length + '</div>' + pm.map(function (x) {
    return '<button type="button" class="riw-card" data-act="riwayatPm" data-id="' + esc(x.ID_PM) + '" data-aset="' + esc(idAset) + '"><div class="d-flex flex-wrap gap-1 align-items-center"><b>' + esc(x.ID_PM) + '</b>' + bStatusWo(x.Status_PM) + '</div><div class="riw-uraian">' + esc(x.Jenis_PM || '—') + '</div>' +
      '<div class="riw-meta"><i class="bi bi-calendar-event me-1"></i>' + esc(fmtTgl(x.Tanggal_Jadwal, true)) + '<i class="bi bi-chevron-right ms-auto"></i></div></button>';
  }).join('');
  return h;
}
ACTIONS.riwayatWo = async function (el) {
  var id = el.dataset.id, idAset = el.dataset.aset;
  bukaModal({ judul: 'Work Order ' + id, size: 'lg', isi: '<div class="skeleton" style="height:240px"></div>', footer: '<button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>' });
  try {
    var d = await call('getWoDetail', id);
    $('#mdlContent .modal-body').innerHTML = detailWoHtml(d);
    $('#mdlContent .modal-footer').innerHTML = '<button class="btn btn-light-x me-auto" data-act="lihatAset" data-id="' + esc(idAset) + '"><i class="bi bi-arrow-left me-1"></i>Kembali ke aset</button><button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>';
  } catch (e) { tutupModal(); gagal(e); }
};

var LK = ['J', 'G', 'L', 'R', 'D'], LANTAI_PILIH = '__pilih__';
/** pilih[i]: null = belum dipilih; string = terpilih ('' = memang kosong, hanya sah untuk Lantai & Detail). */
function lokNilai(level, pilih) {
  var rows = S.master.lokasi.filter(function (l) {
    for (var i = 0; i < level; i++) { if (pilih[i] === null || pilih[i] === undefined || l[LK[i]] !== pilih[i]) return false; }
    return true;
  });
  var s = {}, o = [];
  rows.forEach(function (l) { var v = l[LK[level]]; if (!s.hasOwnProperty(v)) { s[v] = 1; o.push(v); } });
  return o;
}
var RUANG_BARU = '__baru__';
/** Tingkat 0 (Jenis Ruangan) sudah dihapus: selalu bernilai '' dan tidak ditampilkan. izinBaru: pada level Ruangan (3) tambahkan pilihan "+ Tambah ruangan baru" bila Gedung dan Lantai sudah dipilih. */
function lokOpsiHtml(level, pilih, terpilih, izinBaru) {
  var vals = lokNilai(level, pilih), h = '';
  if (level === 4) h = '<option value="">(tanpa detail)</option>';
  else if (level === 2) h = '<option value="' + LANTAI_PILIH + '">— pilih —</option>';
  else h = '<option value="">— pilih —</option>';
  vals.forEach(function (v) {
    if ((level === 4 || level === 3) && v === '') return;
    h += '<option value="' + esc(v) + '"' + (terpilih !== null && terpilih !== undefined && v === terpilih ? ' selected' : '') + '>' + esc(v === '' ? '(tanpa lantai)' : v) + '</option>';
  });
  var terisi = function (x) { return x !== null && x !== undefined; };
  if (izinBaru && level === 3 && terisi(pilih[0]) && terisi(pilih[1]) && terisi(pilih[2])) h += '<option value="' + RUANG_BARU + '"' + (terpilih === RUANG_BARU ? ' selected' : '') + '>＋ Tambah ruangan baru…</option>';
  return h;
}
function formAset(a, idBaru, kunci, paksaBaru) {
  var baru = !a || !!paksaBaru; a = a || {};
  var o = function (k, v) { return opsi(listMaster(k), v, '— pilih —'); };
  var fld = function (lbl, nm, val, extra, req) { return '<div class="col-md-6"><label class="form-label' + (req ? ' req' : '') + '">' + lbl + '</label><input class="form-control" name="' + nm + '" value="' + esc(val) + '" ' + (extra || '') + '></div>'; };
  var pil = (baru && !a.Gedung) ? ['', null, null, null, ''] : ['', a.Gedung || '', a.Lantai || '', a.Ruangan || '', a.Lokasi_Detail || ''];
  var sel = function (i, lbl, nm) {
    var cur = pil[i], pp = pil;
    if (i === 4) return '<div class="col-md-6"><label class="form-label">' + lbl + ' <span class="text-muted fw-normal">(ketik bebas, boleh kosong)</span></label><input class="form-control" name="' + nm + '" maxlength="150" value="' + esc(cur || '') + '" placeholder="mis. dekat jendela barat"></div>';
    if (i === 0) return '<input type="hidden" name="' + nm + '" value="">';
    return '<div class="col-md-6"><label class="form-label' + (i === 1 ? ' req' : '') + '">' + lbl + (i === 2 ? ' <span class="text-muted fw-normal">(boleh tanpa lantai)</span>' : '') + (i === 3 ? ' <span class="text-muted fw-normal">(boleh dikosongkan)</span>' : '') + '</label><select class="form-select" name="' + nm + '" data-lok="' + i + '">' + lokOpsiHtml(i, pp, cur, true) + '</select></div>' + (i === 3 ? '<div class="col-md-6 d-none" id="boxRuBaru"><label class="form-label req">Nama ruangan baru</label><input class="form-control" name="RuanganBaruNama" maxlength="100" autocomplete="off"><div class="form-text">Otomatis didaftarkan ke Master Lokasi saat aset disimpan.</div></div>' : '');
  };
  var picNow = a.PIC || '';
  return '<form id="fAset" data-submit="simpanAset" novalidate><input type="hidden" name="mode" value="' + (baru ? 'baru' : 'ubah') + '">' +
    (a._hint ? '<div class="alert-x info small mb-3">' + a._hint + '</div>' : '') +
    '<div class="row g-3">' +
    '<div class="col-md-6"><label class="form-label">ID Aset</label><input class="form-control" name="ID_Aset" value="' + esc(baru ? idBaru : a.ID_Aset) + '" maxlength="30" autocapitalize="characters" spellcheck="false"' + (baru ? '' : ' readonly') + '>' +
    (baru ? '<div class="form-text">Boleh diganti sendiri (huruf, angka, - atau _). Kosongkan untuk ID otomatis.</div>' : '<div class="form-text">ID tidak dapat diubah.</div>') + '</div>' +
    fld('NUP', 'NUP', a.NUP, 'maxlength="60"') +
    '<div class="col-12"><label class="form-label req">Nama Aset</label><input class="form-control" name="Nama_Aset" value="' + esc(a.Nama_Aset) + '" maxlength="150" required></div>' +
    '<div class="col-md-6"><label class="form-label req">Kategori Aset</label><select class="form-select" name="Kategori_Aset" required>' + o('Kategori Aset', a.Kategori_Aset) + '</select></div>' +
    '<div class="col-md-6"><label class="form-label">Foto Aset</label><input type="file" class="form-control" id="fotoIn" accept="image/*"><div class="form-text">Foto dikompres otomatis (±1280 px).</div></div>' +
    '<div class="col-12 d-flex align-items-center gap-3" id="fotoPrev">' + (a.Foto_Aset_URL ? fotoHtml(a, 'thumb') + '<button type="button" class="btn btn-light-x btn-sm" data-act="hapusFoto">Hapus foto</button>' : '') + '</div>' +
    '<input type="hidden" name="Foto_Aset_URL" value="' + esc(a.Foto_Aset_URL) + '">' +
    '<div class="col-12"><div class="fw-bold small text-muted text-uppercase mt-2">Lokasi</div></div>' +
    sel(0, '', 'Jenis_Ruangan') + sel(1, 'Gedung', 'Gedung') + sel(2, 'Lantai', 'Lantai') + sel(3, 'Ruangan', 'Ruangan') + sel(4, 'Lokasi Detail', 'Lokasi_Detail') +
    '<div class="col-12"><div class="fw-bold small text-muted text-uppercase mt-2">Spesifikasi</div></div>' +
    '<div class="col-md-6"><label class="form-label">Tahun Perolehan <span class="text-muted fw-normal">(boleh kosong)</span></label><input class="form-control" name="Tahun_Perolehan" type="number" min="1900" max="' + (new Date().getFullYear() + 1) + '" value="' + esc(a.Tahun_Perolehan || '') + '"></div>' +
    fld('Merk / Type', 'Merk_Type', a.Merk_Type, 'maxlength="150"') + fld('Nomor Seri', 'Nomor_Seri', a.Nomor_Seri, 'maxlength="100"') +
    '<div class="col-md-6"><label class="form-label">Nilai Perolehan (Rp)</label><input class="form-control" name="Nilai_Perolehan" inputmode="numeric" value="' + esc(a.Nilai_Perolehan) + '" placeholder="mis. 125000000"></div>' +
    '<div class="col-12"><label class="form-label">Spesifikasi</label><textarea class="form-control" name="Spesifikasi" rows="3" maxlength="1000">' + esc(a.Spesifikasi) + '</textarea></div>' +
    '<div class="col-md-4"><label class="form-label req">Kondisi</label><select class="form-select" name="Kondisi" required' + (kunci ? ' disabled' : '') + '>' + o('Kondisi Aset', a.Kondisi || 'Baik') + '</select></div>' +
    '<div class="col-md-4"><label class="form-label">Kritikalitas <span class="text-muted fw-normal">(boleh kosong)</span></label><select class="form-select" name="Kritikalitas">' + opsi(listMaster('Kritikalitas'), a.Kritikalitas, '— belum diketahui —') + '</select></div>' +
    '<div class="col-md-4"><label class="form-label req">Status</label><select class="form-select" name="Status_Aset" required' + (kunci ? ' disabled' : '') + '>' + o('Status Aset', a.Status_Aset || 'Aktif') + '</select></div>' +
    (kunci ? '<div class="col-12"><div class="alert-x warn small mb-0"><i class="bi bi-lock me-1"></i>Kondisi dan Status Aset terkunci: <b>menunggu status Work Order</b> ' + esc(kunci.wo.ID_WO) + ' (' + esc(kunci.wo.Status_WO) + '). Setelah teknisi memilih status WO, perbarui lewat panel &quot;Menunggu Pembaruan Kondisi Aset&quot;.</div></div>' : '') +
    '<div class="col-md-6"><label class="form-label">PIC</label><select class="form-select" name="PIC">' + picOpsi(picNow, '— tanpa PIC —') + '</select>' +
    ((S.pic || []).length ? '<div class="form-text">Pilihan berasal dari akun PIC (Master Admin → Pengguna).</div>' : '<div class="form-text text-warning">Belum ada akun PIC berjabatan. Tambahkan di Master Admin → Pengguna (Peran PIC + Jabatan).</div>') + '</div>' +
    '<div class="col-md-6"><label class="form-label">Frekuensi PM</label><select class="form-select" name="Frekuensi_PM">' + opsi(listMaster('Frekuensi PM'), a.Frekuensi_PM, '— belum diatur —') + '</select></div>' +
    '<div class="col-12"><div class="alert-x danger d-none mb-0" id="asetErr"></div></div></div></form>';
}
var fotoBaru = null;
async function bukaFormAset(a, prefill) {
  fotoBaru = null;
  var id = '', kunci = null;
  if (!a) { try { id = await call('getNextAssetId'); } catch (e) { return gagal(e); } }
  else { try { var kk = await call('getKunciAset', a.ID_Aset); if (kk.terkunci) kunci = kk; } catch (e) { return gagal(e); } }
  bukaModal({ judul: a ? 'Ubah aset' : 'Tambah aset', size: 'lg', isi: formAset(a || prefill || null, id, kunci, !a),
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fAset">Simpan</button>' });
  var fm = $('#fAset'); if (fm) fm.dataset.idSaran = id;
}
ACTIONS.tambahAset = function () { USL_TERKAIT = null; bukaFormAset(null); };
var USL_TERKAIT = null;
ACTIONS.ubahAset = function (el) { var a = cariAset(el.dataset.id); if (a) { tutupModal(); setTimeout(function () { bukaFormAset(a); }, 250); } };
ACTIONS.hapusFoto = function () { fotoBaru = null; var f = $('#fAset'); f.Foto_Aset_URL.value = ''; $('#fotoPrev').innerHTML = ''; $('#fotoIn').value = ''; };
document.addEventListener('change', async function (e) {
  var t = e.target;
  if (t.dataset && t.dataset.lok !== undefined && t.closest('#fAset')) { // dropdown lokasi berjenjang
    var f = $('#fAset'), i = Number(t.dataset.lok), nm = ['Jenis_Ruangan', 'Gedung', 'Lantai', 'Ruangan', 'Lokasi_Detail'];
    var pil = nm.map(function (n, k) { var v = f[n].value; if (k === 0) return ''; if (k === 2) return v === LANTAI_PILIH ? null : v; if (k === 4) return v; return v === '' ? null : v; });
    for (var j = i + 1; j < 4; j++) { pil[j] = null; f[nm[j]].innerHTML = lokOpsiHtml(j, pil, null, true); }
    $('#boxRuBaru').classList.toggle('d-none', f.Ruangan.value !== RUANG_BARU);
    if (f.Ruangan.value === RUANG_BARU) f.RuanganBaruNama.focus();
  }
  if (t.id === 'fotoIn' && t.files && t.files[0]) {
    try { fotoBaru = await kompresFoto(t.files[0]); $('#fotoPrev').innerHTML = '<img class="thumb" style="width:72px;height:72px" src="' + fotoBaru.preview + '" alt="Pratinjau"><span class="small text-muted">Foto baru siap diunggah</span>'; }
    catch (er) { toast(er.message, 'bad'); t.value = ''; fotoBaru = null; }
  }
});
SUBMITS.simpanAset = async function (f) {
  var err = $('#asetErr'); err.classList.add('d-none');
  var p = {}; ['ID_Aset', 'NUP', 'Nama_Aset', 'Kategori_Aset', 'Foto_Aset_URL', 'Gedung', 'Lantai', 'Ruangan', 'Lokasi_Detail', 'Tahun_Perolehan', 'Merk_Type', 'Nomor_Seri', 'Spesifikasi', 'Nilai_Perolehan', 'Kondisi', 'Kritikalitas', 'Status_Aset', 'PIC', 'Frekuensi_PM']
    .forEach(function (k) { p[k] = f[k].value; });
  var mode = f.mode.value;
  if (mode === 'ubah') p.ID_Asal = p.ID_Aset;
  var salah = function (m) { err.textContent = m; err.classList.remove('d-none'); };
  if (p.Lantai === LANTAI_PILIH) { salah('Pilih Lantai (atau "(tanpa lantai)" bila lokasi tidak memiliki lantai).'); return; }
  if (p.Ruangan === RUANG_BARU) {
    p.Ruangan = f.RuanganBaruNama.value.trim().replace(/\s+/g, ' ');
    if (!p.Ruangan) { salah('Isi nama ruangan baru.'); return; }
    p.Ruangan_Baru = true;
  }
  if (!p.Nama_Aset.trim() || !p.Kategori_Aset || !p.Gedung || !p.Kondisi || !p.Status_Aset) { salah('Lengkapi semua kolom bertanda *.'); return; }
  if (mode === 'baru') {
    p.ID_Aset = p.ID_Aset.trim();
    if (!p.ID_Aset) p.ID_Aset = $('#fAset').dataset.idSaran || '';
    if (p.ID_Aset && !/^[A-Za-z0-9][A-Za-z0-9_\-]{1,29}$/.test(p.ID_Aset)) { salah('ID Aset hanya boleh huruf, angka, "-" atau "_" (2–30 karakter).'); return; }
    if (p.ID_Aset && cariAset(p.ID_Aset)) { salah('ID Aset "' + p.ID_Aset + '" sudah digunakan.'); return; }
  }
  var btn = $('button[type=submit][form=fAset]'); sibuk(btn, true, 'Menyimpan…');
  try {
    if (fotoBaru) { var u = await call('uploadFile', 'aset', p.ID_Aset, { name: fotoBaru.name, mime: fotoBaru.mime, base64: fotoBaru.base64 }); p.Foto_Aset_URL = u.url; }
    var d = await call('saveAsset', p, mode);
    if (d.master) S.master = d.master;
    var idx = -1; S.aset.forEach(function (a, i) { if (a.ID_Aset === d.aset.ID_Aset) idx = i; });
    if (idx >= 0) S.aset[idx] = d.aset; else S.aset.push(d.aset);
    S.aset.sort(function (a, b) { return a.ID_Aset < b.ID_Aset ? -1 : 1; });
    tutupModal(); toast((mode === 'ubah' ? 'Aset diperbarui.' : 'Aset ditambahkan.') + (p.Ruangan_Baru ? ' Ruangan baru masuk Master Lokasi.' : ''), 'ok');
    if (USL_TERKAIT && mode === 'baru') {
      var uid = USL_TERKAIT; USL_TERKAIT = null;
      try { await call('prosesUsulan', uid, 'terima', { ID_Aset: d.aset.ID_Aset }); toast('Usulan aset diterima dan WO ditautkan ke ' + d.aset.ID_Aset + '.', 'ok', 5000); } catch (e2) { gagal(e2); }
      muatNotif(); if (S.page === 'admin' && S.adminTab === 'usulan') adminUsulan();
    }
    if (S.page === 'aset') { isiTabelAset(); isiStatAset(); } else if (S.page === 'beranda') renderBeranda();
  } catch (e) { err.textContent = e.message; err.classList.remove('d-none'); } finally { sibuk(btn, false); }
};
ACTIONS.hapusAset = async function (el) {
  var a = cariAset(el.dataset.id); if (!a) return;
  if (!(await konfirm('Hapus aset?', a.ID_Aset + ' — ' + a.Nama_Aset + '. Aset yang memiliki riwayat WO/PM tidak dihapus permanen, melainkan diberi status "' + DIHAPUS + '".', 'Hapus'))) return;
  try {
    var d = await call('deleteAsset', a.ID_Aset);
    if (d.mode === 'soft') a.Status_Aset = DIHAPUS; else S.aset = S.aset.filter(function (x) { return x.ID_Aset !== a.ID_Aset; });
    delete S.sel[a.ID_Aset]; toast(d.mode === 'soft' ? 'Aset memiliki riwayat — status diubah menjadi "' + DIHAPUS + '".' : 'Aset dihapus.', 'ok');
    isiTabelAset(); isiStatAset();
  } catch (e) { gagal(e); }
};

/* ════════ Impor CSV ════════ */
var KOL_ASET = ['ID_Aset', 'NUP', 'Nama_Aset', 'Kategori_Aset', 'Gedung', 'Lantai', 'Ruangan', 'Lokasi_Detail', 'Tahun_Perolehan', 'Merk_Type', 'Nomor_Seri', 'Spesifikasi', 'Nilai_Perolehan', 'Kondisi', 'Kritikalitas', 'Status_Aset', 'PIC', 'Frekuensi_PM'];
var impor = { rows: [], hasil: [] };
ACTIONS.impor = function () {
  impor = { rows: [], hasil: [] };
  bukaModal({ judul: 'Impor aset dari CSV', size: 'lg',
    isi: '<div class="alert-x info small">Kolom: <code>' + KOL_ASET.join(', ') + '</code>. ID_Aset boleh kosong (dibuat otomatis). Foto tidak diimpor lewat CSV. Nilai harus sesuai Data Master (kategori, kondisi, dst.). <b>Ruangan dan Kritikalitas boleh dikosongkan</b> bila belum diketahui. Gedung, Lantai, dan Ruangan yang belum ada otomatis didaftarkan ke Master Lokasi, jadi periksa ejaannya di pratinjau. Maksimal 2.000 baris.</div>' +
      '<div class="d-flex gap-2 flex-wrap mb-2"><button class="btn btn-light-x btn-sm" data-act="templateCsv"><i class="bi bi-download me-1"></i>Unduh template</button></div>' +
      '<div class="form-check mb-3"><input class="form-check-input" type="checkbox" id="imporUbah"><label class="form-check-label small" for="imporUbah"><b>Perbarui aset yang ID_Aset-nya sudah ada</b> (kolom yang dikosongkan di CSV tidak mengubah isian lama; foto tidak berubah). Tanpa centang, ID yang sudah ada ditolak.</label></div>' +
      '<div class="drop"><input type="file" class="form-control" id="csvIn" accept=".csv,text/csv"></div><div id="imporPrev" class="mt-3"></div>',
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Tutup</button><button class="btn btn-primary" id="btnImporGo" data-act="imporGo" disabled>Impor baris valid</button>' });
};
ACTIONS.templateCsv = function () {
  var contoh = ['', '1234/5', 'AC Split Ruang Rapat', listMaster('Kategori Aset')[0] || '', S.master.lokasi[0] ? S.master.lokasi[0].G : '', S.master.lokasi[0] ? S.master.lokasi[0].L : '', S.master.lokasi[0] ? S.master.lokasi[0].R : '', S.master.lokasi[0] ? S.master.lokasi[0].D : '', '2020', 'Daikin', 'SN123', '2 PK', '8500000', 'Baik', 'Sedang', 'Aktif', 'Teknisi A', 'Bulanan'];
  unduhBlob('template_impor_aset.csv', new Blob([csvKeString([KOL_ASET, contoh])], { type: 'text/csv;charset=utf-8' }));
};
document.addEventListener('change', async function (e) {
  if (e.target.id === 'imporUbah') { var fi = $('#csvIn'); if (fi && fi.files && fi.files[0]) fi.dispatchEvent(new Event('change', { bubbles: true })); return; }
  if (e.target.id !== 'csvIn' || !e.target.files[0]) return;
  impor.ubah = !!($('#imporUbah') && $('#imporUbah').checked);
  var box = $('#imporPrev'); box.innerHTML = '<div class="text-muted"><span class="spinner-border spinner-border-sm me-2"></span>Membaca &amp; memeriksa…</div>';
  try {
    var tabel = parseCsv(await bacaFile(e.target.files[0], true));
    if (tabel.length < 2) throw new Error('File kosong atau tidak memiliki baris data.');
    if (tabel.length > 2001) throw new Error('Maksimal 2.000 baris per impor.');
    var map = {}; tabel[0].forEach(function (h, i) { map[String(h).trim().toLowerCase().replace(/[\s]+/g, '_')] = i; });
    var kolom = KOL_ASET.filter(function (k) { return map[k.toLowerCase()] !== undefined; });
    if (kolom.indexOf('Nama_Aset') < 0) throw new Error('Kolom Nama_Aset tidak ditemukan. Gunakan template.');
    impor.rows = tabel.slice(1).map(function (r, i) { var o = { _no: i + 2 }; kolom.forEach(function (k) { o[k] = String(r[map[k.toLowerCase()]] == null ? '' : r[map[k.toLowerCase()]]).trim(); }); return o; });
    impor.hasil = [];
    for (var i = 0; i < impor.rows.length; i += 200) { var h = await call('validateAssetRows', impor.rows.slice(i, i + 200), impor.ubah); impor.hasil = impor.hasil.concat(h); }
    var ok = impor.hasil.filter(function (x) { return x.ok; }).length, bad = impor.hasil.length - ok;
    var nUbah = impor.hasil.filter(function (x) { return x.ok && x.ubah && x.diubah.length; }).length, nSama = impor.hasil.filter(function (x) { return x.ok && x.ubah && !x.diubah.length; }).length;
    impor.ringkas = '<div class="mb-2"><span class="badge-x ok me-1">' + ok + ' valid</span><span class="badge-x bad">' + bad + ' bermasalah</span>' + (impor.ubah ? ' <span class="small text-muted ms-1">' + (ok - nUbah - nSama) + ' ditambah · ' + nUbah + ' diperbarui · ' + nSama + ' tanpa perubahan</span>' : '') + '</div>';
    impor.hanyaMasalah = false;
    renderTabelImpor();
    $('#btnImporGo').disabled = ok === 0;
  } catch (er) { box.innerHTML = '<div class="alert-x danger">' + esc(er.message) + '</div>'; }
});
function renderTabelImpor() {
  var box = $('#imporPrev'); if (!box) return;
  var idx = impor.hasil.map(function (x, i) { return i; });
  var bad = idx.filter(function (i) { return !impor.hasil[i].ok; }), good = idx.filter(function (i) { return impor.hasil[i].ok; });
  var urut = impor.hanyaMasalah ? bad : bad.concat(good);
  box.innerHTML = impor.ringkas +
    (bad.length ? '<div class="form-check mb-2"><input class="form-check-input" type="checkbox" id="imporMasalah"' + (impor.hanyaMasalah ? ' checked' : '') + '><label class="form-check-label small" for="imporMasalah">Tampilkan hanya baris bermasalah (' + bad.length + ')</label></div>' : '') +
    '<div class="prev-scroll"><table class="tbl"><thead><tr><th>Baris</th><th>Nama</th><th>Hasil</th></tr></thead><tbody>' + urut.slice(0, 300).map(function (i) {
      var x = impor.hasil[i];
      return '<tr><td>' + esc(x.no) + '</td><td>' + esc(impor.rows[i].Nama_Aset) + '</td><td>' + (x.ok ? (x.ubah ? (x.diubah.length ? '<span class="badge-x warn">Perbarui</span> <span class="small text-muted">' + esc(x.diubah.join(', ')) + '</span>' : '<span class="badge-x ok">Tanpa perubahan</span>') : '<span class="badge-x ok">OK</span>') + (x.ruanganBaru ? ' <span class="small text-muted">Lokasi baru didaftarkan (' + esc((x.levelBaru || []).join(', ') || 'Ruangan') + '): ' + esc(x.ruanganBaru) + '</span>' : '') : '<span class="text-danger small">' + esc(x.errors.join(' ')) + '</span>') + '</td></tr>';
    }).join('') + '</tbody></table></div>' + (urut.length > 300 ? '<div class="small text-muted mt-1">Menampilkan 300 baris pertama (baris bermasalah ditampilkan lebih dulu).</div>' : '');
}
document.addEventListener('change', function (e) { if (e.target.id === 'imporMasalah') { impor.hanyaMasalah = e.target.checked; renderTabelImpor(); } });
ACTIONS.imporGo = async function (el) {
  var valid = impor.rows.filter(function (r, i) { return impor.hasil[i] && impor.hasil[i].ok; });
  if (!valid.length) return;
  sibuk(el, true, 'Mengimpor…');
  var berhasil = 0, diperbarui = 0, sama = 0, gagalList = [], ruangBaru = 0, total = valid.length, prev = $('#imporPrev');
  var progres = function (n) { var pct = Math.round(n / total * 100); prev.innerHTML = '<div class="small text-muted mb-1">Mengimpor ' + n + ' dari ' + total + ' baris… jangan tutup jendela ini.</div><div class="progress" style="height:10px"><div class="progress-bar" style="width:' + pct + '%"></div></div>'; };
  progres(0);
  try {
    for (var i = 0; i < valid.length; i += 200) {
      var d = await call('importAssetRows', valid.slice(i, i + 200), impor.ubah); berhasil += d.berhasil; diperbarui += d.diperbarui || 0; sama += d.tanpaPerubahan || 0; gagalList = gagalList.concat(d.gagal); ruangBaru += d.ruanganBaru || 0; if (d.master) S.master = d.master;
      progres(Math.min(total, i + 200));
    }
    S.aset = await call('getAssetList'); isiTabelAset(); isiStatAset();
    var adaGagal = gagalList.length > 0;
    prev.innerHTML = '<div class="impor-selesai' + (adaGagal ? ' sebagian' : '') + '" role="status"><i class="bi ' + (adaGagal ? 'bi-exclamation-circle-fill' : 'bi-check-circle-fill') + '"></i><div><div class="h5 fw-bold mb-1">' + (adaGagal ? 'Impor selesai dengan catatan' : 'Impor selesai') + '</div>' +
      '<div><b>' + berhasil + '</b> aset ditambah' + (impor.ubah ? ', <b>' + diperbarui + '</b> diperbarui, ' + sama + ' tanpa perubahan' : '') + (adaGagal ? ', <b>' + gagalList.length + '</b> baris gagal' : '') + '.' + (ruangBaru ? ' ' + ruangBaru + ' lokasi baru didaftarkan ke Master Lokasi.' : '') + '</div></div></div>' +
      (adaGagal ? '<div class="alert-x danger mt-2"><b>Baris gagal:</b><br>' + gagalList.slice(0, 20).map(function (g) { return 'Baris ' + esc(g.no) + ': ' + esc(g.errors.join(' ')); }).join('<br>') + '</div>' : '');
    el.textContent = 'Impor selesai'; el.disabled = true; el.classList.remove('btn-primary'); el.classList.add('btn-success');
    var mb = $('#mdlContent .modal-body'); if (mb) mb.scrollTop = mb.scrollHeight;
    toast('Impor selesai: ' + berhasil + ' ditambah' + (impor.ubah ? ', ' + diperbarui + ' diperbarui' : '') + (adaGagal ? ', ' + gagalList.length + ' gagal' : '') + '.', adaGagal ? 'warn' : 'ok', 7000);
  } catch (e) { prev.innerHTML = '<div class="alert-x danger">Impor terhenti: ' + esc(e.message) + (berhasil ? ' (' + berhasil + ' baris sudah masuk sebelum terhenti.)' : '') + '</div>'; gagal(e); sibuk(el, false); }
};

/* ════════ Ekspor Excel ════════ */
ACTIONS.ekspor = async function (el) {
  var data = asetTersaring(); if (!data.length) { toast('Tidak ada data untuk diekspor.', 'warn'); return; }
  var kolom = ['ID_Aset', 'NUP', 'Nama_Aset', 'Kategori_Aset', 'Gedung', 'Lantai', 'Ruangan', 'Lokasi_Detail', 'Tahun_Perolehan', 'Merk_Type', 'Nomor_Seri', 'Spesifikasi', 'Nilai_Perolehan', 'Kondisi', 'Kritikalitas', 'Status_Aset', 'PIC', 'Frekuensi_PM', 'Dibuat_Pada', 'Diubah_Pada'];
  if (!bisaLihatNilai()) kolom = kolom.filter(function (k) { return k !== 'Nilai_Perolehan'; });
  var tabel = [kolom].concat(data.map(function (a) { return kolom.map(function (k) { var v = k === 'PIC' ? picLabel(a.PIC) : a[k]; return typeof v === 'number' ? v : amanSel(v); }); }));
  sibuk(el, true, 'Menyiapkan…');
  try {
    await muatSkrip(CDN_XLSX, function () { return typeof XLSX !== 'undefined'; });
    var ws = XLSX.utils.aoa_to_sheet(tabel), wb = XLSX.utils.book_new();
    ws['!cols'] = kolom.map(function (k) { return { wch: Math.min(40, Math.max(k.length + 2, 14)) }; });
    XLSX.utils.book_append_sheet(wb, ws, 'Data Aset');
    var out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    unduhBlob('Data_Aset_' + new Date().toISOString().slice(0, 10) + '.xlsx', new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    toast('Berkas Excel diunduh (' + data.length + ' aset).', 'ok');
  } catch (e) {
    unduhBlob('Data_Aset_' + new Date().toISOString().slice(0, 10) + '.csv', new Blob([csvKeString(tabel.map(function (r) { return r.map(function (c) { return c; }); }))], { type: 'text/csv;charset=utf-8' }));
    toast('Pustaka Excel tidak tersedia — diunduh sebagai CSV (dapat dibuka di Excel).', 'warn', 6000);
  } finally { sibuk(el, false); }
};

/* ════════ Label QR ════════ */
ACTIONS.labelSatu = function (el) { S.sel = {}; S.sel[el.dataset.id] = 1; tutupModal(); setTimeout(function () { go('label'); }, 250); };
ACTIONS.labelDariPilihan = function () { go('label'); };
function asetUntukLabel() {
  var l = S.lbl, ids = Object.keys(S.sel);
  if (ids.length) return S.aset.filter(function (a) { return S.sel[a.ID_Aset]; });
  return S.aset.filter(function (a) { return a.Status_Aset !== DIHAPUS && (!l.kat || a.Kategori_Aset === l.kat) && (!l.ged || a.Gedung === l.ged) && (!l.lan || a.Lantai === l.lan); });
}
function lembarLabel(list) {
  var l = S.lbl, h = '';
  for (var i = 0; i < list.length; i += 18) {
    h += '<div class="lbl-sheet" style="--mt:' + Number(l.mt) + 'mm;--mx:' + Number(l.mx) + 'mm;--gap:' + Number(l.gap) + 'mm">' + list.slice(i, i + 18).map(function (a) {
      var t = qrTeks(a);
      return '<div class="lbl"><div class="q">' + (t ? qrSvg(t) : '') + '</div><div class="t"><div class="id">' + esc(a.ID_Aset) + '</div><b>' + esc(a.Nama_Aset) + '</b><small>' + esc([a.Gedung, a.Lantai, a.Ruangan].filter(Boolean).join(' · ')) + '</small></div></div>';
    }).join('') + '</div>';
  }
  return h;
}
var LBL_BAWAAN = { mode: 'zt411', W: 80, H: 24, mL: 2.5, mR: 2.5, mT: 2.5, mB: 2.5, ox: 0, oy: 0, salinan: 1 };
var LBL_LOGO = 'data:image/png;base64,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';
var LBL_JUDUL = 'KARTU PEMELIHARAAN DIGITAL', LBL_INSTANSI = 'BAGIAN PEMELIHARAAN, INSTALASI, DAN ANGKUTAN<br>SEKRETARIAT JENDERAL MPR RI';
(function () { try { var o = JSON.parse(LS.get('sma_lbl_zt2') || 'null'); if (o) Object.keys(LBL_BAWAAN).forEach(function (k) { if (o[k] !== undefined && o[k] !== null && o[k] !== '') S.lbl[k] = o[k]; }); } catch (e) {} })();
function simpanLblZt() { try { var o = {}; Object.keys(LBL_BAWAAN).forEach(function (k) { o[k] = S.lbl[k]; }); LS.set('sma_lbl_zt2', JSON.stringify(o)); } catch (e) {} }
/** QR rapat untuk label kecil: zona tenang 1 modul (sisanya disediakan margin label). */
function qrSvgRapat(teks) {
  var q = qrcode(0, 'M'); q.addData(String(teks)); q.make();
  var n = q.getModuleCount(), d = '';
  for (var r = 0; r < n; r++) for (var c = 0; c < n; c++) if (q.isDark(r, c)) d += 'M' + (c + 1) + ',' + (r + 1) + 'h1v1h-1z';
  return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="1 1 ' + n + ' ' + n + '" shape-rendering="crispEdges" role="img" aria-label="Kode QR"><rect x="1" y="1" width="' + n + '" height="' + n + '" fill="#fff"/><path d="' + d + '" fill="#000"/></svg>';
}
/** Satu label printer ZT411 (bawaan 80 × 24 mm). Tata letak mengikuti desain: logo + judul, ID Aset, nama aset, gedung, QR di kanan. */
function labelZt(a) {
  var l = S.lbl, t = qrTeks(a), id = String(a.ID_Aset || ''), fs = id.length > 28 ? Math.max(6, 12 * 28 / id.length) : 12;
  var n = function (v) { return Number(v) || 0; };
  return '<div class="zl" style="width:' + n(l.W) + 'mm;height:' + n(l.H) + 'mm;padding:' + n(l.mT) + 'mm ' + n(l.mR) + 'mm ' + n(l.mB) + 'mm ' + n(l.mL) + 'mm"><div class="zl-in" style="transform:translate(' + n(l.ox) + 'mm,' + n(l.oy) + 'mm)">' +
    '<div class="zl-t"><div class="zl-h"><img src="' + LBL_LOGO + '" alt=""><div class="zl-ht"><b>' + LBL_JUDUL + '</b><small>' + LBL_INSTANSI + '</small></div></div>' +
    '<div class="zl-m"><div class="zl-id" style="font-size:' + fs.toFixed(1) + 'pt">' + esc(id) + '</div><div class="zl-nm">' + esc(a.Nama_Aset) + '</div></div>' +
    '<div class="zl-gd">' + esc(a.Gedung) + '</div></div>' +
    '<div class="zl-q">' + (t ? qrSvgRapat(t) : '') + '</div></div></div>';
}
function labelZtBanyak(list) {
  var h = '', k = Math.max(1, Math.min(20, Math.round(Number(S.lbl.salinan) || 1)));
  list.forEach(function (a) { for (var i = 0; i < k; i++) h += labelZt(a); });
  return h;
}
function renderLabel() {
  var l = S.lbl, nSel = Object.keys(S.sel).length, zt = true;
  var num = function (lbl, k, min, max, step, kol) { return '<div class="' + (kol || 'col-6 col-md-3') + '"><label class="form-label">' + lbl + '</label><input class="form-control" type="number" min="' + min + '" max="' + max + '" step="' + step + '" data-lbl="' + k + '" value="' + l[k] + '"></div>'; };
  $('#main').innerHTML = '<div class="page-head"><button class="btn btn-light-x" data-act="kembaliAset"><i class="bi bi-arrow-left me-1"></i>Kembali ke Data Aset</button><h2>Label QR Aset</h2><div class="ms-auto d-flex gap-2">' +
    (zt ? '<button class="btn btn-light-x" data-act="cetakLabelUji"><i class="bi bi-printer me-1"></i>Cetak 1 label (uji)</button>' : '') + '<button class="btn btn-primary" data-act="cetakLabel"><i class="bi bi-printer me-1"></i>Cetak</button></div></div>' +
    (S.cfg.webAppUrl ? '' : '<div class="alert-x warn">URL Web App belum terdeteksi. Isi di <b>Master Admin → Pengaturan</b> sebelum mencetak, agar QR mengarah ke aplikasi.</div>') +
    '<div class="card-x mb-3"><div class="card-b"><div class="row g-3 align-items-end">' +
    (nSel ? '<div class="col-12"><div class="alert-x info mb-0">Mencetak <b>' + nSel + '</b> aset yang dipilih di Data Aset. <button class="btn-link-x" data-act="resetLabelPilih">Pakai filter</button></div></div>' :
      '<div class="col-md-4"><label class="form-label">Kategori</label><select class="form-select" data-lbl="kat">' + opsi(listMaster('Kategori Aset'), l.kat, 'Semua') + '</select></div>' +
      '<div class="col-md-4"><label class="form-label">Gedung</label><select class="form-select" data-lbl="ged">' + opsi(uniq(S.master.lokasi.map(function (x) { return x.G; })), l.ged, 'Semua') + '</select></div>' +
      '<div class="col-md-4"><label class="form-label">Lantai</label><select class="form-select" data-lbl="lan">' + opsi(uniq(S.master.lokasi.filter(function (x) { return !l.ged || x.G === l.ged; }).map(function (x) { return x.L; })), l.lan, 'Semua') + '</select></div>') +
    (zt ?
      num('Lebar label (mm)', 'W', 20, 200, 0.5) + num('Tinggi label (mm)', 'H', 10, 200, 0.5) + num('Jumlah salinan per aset', 'salinan', 1, 20, 1) +
      '<div class="col-6 col-md-3 d-flex align-items-end"><button class="btn btn-light-x w-100" data-act="lblReset"><i class="bi bi-arrow-counterclockwise me-1"></i>Kembalikan bawaan</button></div>' +
      '<div class="col-12"><div class="fw-bold small text-muted text-uppercase">Gap / margin dalam label (mm)</div></div>' +
      num('Gap kiri', 'mL', 0, 20, 0.5) + num('Gap kanan', 'mR', 0, 20, 0.5) + num('Gap atas', 'mT', 0, 20, 0.5) + num('Gap bawah', 'mB', 0, 20, 0.5) +
      '<div class="col-12"><div class="fw-bold small text-muted text-uppercase">Geser hasil cetak (kalibrasi, mm) — negatif = ke kiri / ke atas</div></div>' +
      num('Geser horizontal', 'ox', -10, 10, 0.1) + num('Geser vertikal', 'oy', -10, 10, 0.1) +
      '<div class="col-12 small text-muted">Pengaturan tersimpan di perangkat ini. <b>Cara cetak:</b> pilih printer <b>Zebra ZT411</b>, ukuran kertas <b>' + esc(l.W) + ' × ' + esc(l.H) + ' mm</b> (atur di Preferences driver), <b>Margin: Tidak ada</b>, skala <b>100%</b>, matikan <b>Header &amp; footer</b>. Cetak 1 label dulu; bila posisi kurang pas, ubah Gap atau Geser.</div>' :
      '<div class="col-4 col-md-2"><label class="form-label">Margin atas (mm)</label><input class="form-control" type="number" min="0" max="30" data-lbl="mt" value="' + l.mt + '"></div>' +
      '<div class="col-4 col-md-2"><label class="form-label">Margin kiri/kanan</label><input class="form-control" type="number" min="0" max="30" data-lbl="mx" value="' + l.mx + '"></div>' +
      '<div class="col-4 col-md-2"><label class="form-label">Jarak (mm)</label><input class="form-control" type="number" min="0" max="15" data-lbl="gap" value="' + l.gap + '"></div>' +
      '<div class="col-md-6 small text-muted">A4, 3 kolom × 6 baris (18 label per lembar). Saat mencetak, pilih skala 100% dan matikan header/footer browser.</div>') +
    '</div></div></div><div class="small text-muted mb-2" id="lblInfo"></div><div class="lbl-preview" id="lblPrev"></div>';
  isiLabel();
}
ACTIONS.lblReset = function () { Object.keys(LBL_BAWAAN).forEach(function (k) { if (k !== 'mode') S.lbl[k] = LBL_BAWAAN[k]; }); simpanLblZt(); renderLabel(); toast('Pengaturan label dikembalikan ke bawaan.', 'ok'); };
function isiLabel() {
  var list = asetUntukLabel(), zt = true;
  if (zt) {
    var total = list.length * Math.max(1, Math.min(20, Math.round(Number(S.lbl.salinan) || 1)));
    $('#lblInfo').textContent = list.length + ' aset · ' + total + ' label (1 label = 1 lembar pada printer label)';
    $('#lblPrev').innerHTML = list.length ? '<div class="zl-prev">' + list.slice(0, 12).map(labelZt).join('') + '</div>' + (list.length > 12 ? '<div class="small text-muted text-center mt-2">Pratinjau 12 label pertama; semua label ikut tercetak.</div>' : '') : '<div class="empty"><i class="bi bi-qr-code"></i>Tidak ada aset sesuai filter.</div>';
    return;
  }
  $('#lblInfo').textContent = list.length + ' label · ' + Math.ceil(list.length / 18) + ' lembar';
  $('#lblPrev').innerHTML = list.length ? lembarLabel(list.slice(0, 36)) + (list.length > 36 ? '<div class="small text-muted text-center">Pratinjau 2 lembar pertama; semua label ikut tercetak.</div>' : '') : '<div class="empty"><i class="bi bi-qr-code"></i>Tidak ada aset sesuai filter.</div>';
}
ACTIONS.resetLabelPilih = function () { S.sel = {}; renderLabel(); };
document.addEventListener('change', function (e) {
  var t = e.target; if (!t.dataset || !t.dataset.lbl || S.page !== 'label') return;
  var kz = ['ox', 'oy'].indexOf(t.dataset.lbl) >= 0;
  S.lbl[t.dataset.lbl] = t.type === 'number' ? (kz ? (Number(t.value) || 0) : Math.max(0, Number(t.value) || 0)) : t.value;
  if (['W', 'H'].indexOf(t.dataset.lbl) >= 0 && S.lbl[t.dataset.lbl] < 10) S.lbl[t.dataset.lbl] = 10;
  if (t.dataset.lbl === 'salinan') S.lbl.salinan = Math.max(1, Math.min(20, Math.round(S.lbl.salinan) || 1));
  if (t.dataset.lbl === 'ged') S.lbl.lan = '';
  simpanLblZt();
  if (t.dataset.lbl === 'ged' || t.dataset.lbl === 'W' || t.dataset.lbl === 'H') renderLabel(); else isiLabel();
});
/** Cetak: ZT411 → satu label per halaman ukuran persis lebar × tinggi label; A4 → lembar 18 label. */
function cetakIsi(html, ukuran) {
  var st = document.getElementById('pgDyn'); if (st) st.remove();
  st = document.createElement('style'); st.id = 'pgDyn'; st.textContent = '@media print { @page { size:' + ukuran + '; margin:0; } }'; document.head.appendChild(st);
  $('#printRoot').innerHTML = html;
  var tunggu = function (p) { return Promise.race([p, new Promise(function (r) { setTimeout(r, 2500); })]); };
  var f = document.fonts && document.fonts.load ? Promise.all([document.fonts.load('700 10px Poppins'), document.fonts.load('400 10px Poppins'), document.fonts.load('400 12px "Bebas Neue"')]).catch(function () {}) : Promise.resolve();
  tunggu(f).then(function () { setTimeout(function () { window.print(); }, 150); });
}
function cetakLabelAset(list, uji) {
  if (!list.length) { toast('Tidak ada label untuk dicetak.', 'warn'); return; }
  if (!S.cfg.webAppUrl && list.some(function (a) { return !a.QR_URL; })) { toast('Atur URL Web App dulu agar QR valid.', 'bad'); return; }
  var l = S.lbl, daftar = uji ? list.slice(0, 1) : list, k = uji ? 1 : Math.max(1, Math.min(20, Math.round(Number(l.salinan) || 1)));
  if (!uji && daftar.length * k > 300) { konfirm('Cetak ' + (daftar.length * k) + ' label?', 'Jumlah label besar; pastikan gulungan label cukup.', 'Cetak', false).then(function (ya) { if (ya) cetakIsi(labelZtBanyak(daftar), l.W + 'mm ' + l.H + 'mm'); }); return; }
  cetakIsi(uji ? labelZt(daftar[0]) : labelZtBanyak(daftar), l.W + 'mm ' + l.H + 'mm');
}
ACTIONS.cetakLabel = function () { cetakLabelAset(asetUntukLabel(), false); };
ACTIONS.cetakLabelUji = function () { cetakLabelAset(asetUntukLabel(), true); };
window.addEventListener('afterprint', function () { var p = $('#printRoot'); if (p) p.innerHTML = ''; var st = document.getElementById('pgDyn'); if (st) st.remove(); });

/* ════════ Master Admin ════════ */
function renderAdmin() {
  var tabs = [['pengguna', 'Pengguna'], ['usulan', 'Usulan Data'], ['master', 'Data Master'], ['lokasi', 'Master Lokasi'], ['pengaturan', 'Pengaturan'], ['log', 'Log Aktivitas']];
  $('#main').innerHTML = '<div class="page-head"><h2>Master Admin</h2></div><ul class="nav nav-tabs mb-3" style="overflow-x:auto;flex-wrap:nowrap">' +
    tabs.map(function (t) { return '<li class="nav-item"><button class="nav-link text-nowrap' + (S.adminTab === t[0] ? ' active' : '') + '" data-act="adminTab" data-t="' + t[0] + '">' + t[1] + (t[0] === 'usulan' && notifUsulan() ? ' <span class="nbadge">' + notifUsulan() + '</span>' : '') + '</button></li>'; }).join('') + '</ul><div id="adminBody"><div class="skeleton" style="height:200px"></div></div>';
  var t = S.adminTab;
  if (t === 'pengguna') adminUsers(); else if (t === 'usulan') adminUsulan(); else if (t === 'master') adminMaster(); else if (t === 'lokasi') adminLokasi(); else if (t === 'pengaturan') adminPengaturan(); else adminLog();
}
ACTIONS.adminTab = function (el) { S.adminTab = el.dataset.t; renderAdmin(); };

/* Pemberitahuan & Usulan Data (Master Admin) */
function notifUsulan() { return S.notif ? (S.notif.vendor + S.notif.asetBaru + (S.notif.ruangan || 0)) : 0; }
async function muatNotif() {
  if (!isAdmin() || S.app !== 'dashboard' || !S.token) return;
  try { S.notif = await call('getNotifAdmin'); } catch (e) { return; }
  var b = $('#btnNotif'), n = $('#notifN'), tot = S.notif ? S.notif.total : 0;
  if (b) b.classList.remove('d-none'); if (n) { n.textContent = tot > 99 ? '99+' : tot; n.classList.toggle('d-none', !tot); }
  buatMenu();
  var tb = $('[data-act=adminTab][data-t=usulan]'); if (tb) tb.innerHTML = 'Usulan Data' + (notifUsulan() ? ' <span class="nbadge">' + notifUsulan() + '</span>' : '');
}
setInterval(function () { if (!document.hidden) muatNotif(); }, 90000);
ACTIONS.notifBuka = async function () {
  await muatNotif(); var n = S.notif || { vendor: 0, asetBaru: 0, ruangan: 0, rusak: 0 };
  var baris = function (ic, kls, jml, judul, ket, act, p) { return '<button type="button" class="notif-row' + (jml ? '' : ' kosong') + '" data-act="' + act + '"' + (p ? ' data-p="' + p + '"' : '') + '><span class="ic ' + kls + '"><i class="bi ' + ic + '"></i></span><span class="flex-grow-1"><b>' + judul + '</b><span class="d-block small text-muted">' + ket + '</span></span><span class="badge-x ' + (jml ? 'warn' : '') + '">' + jml + '</span></button>'; };
  bukaModal({ judul: 'Pemberitahuan', size: 'md', isi: '<div class="d-grid gap-2">' +
    baris('bi-exclamation-octagon', 'bad', n.rusak, 'Aset dilaporkan rusak', 'Teknisi memilih Rusak Ringan / Rusak Berat. Perbarui kondisi aset di Data Aset.', 'notifKe', 'aset') +
    baris('bi-box-seam', 'info', n.asetBaru, 'Aset belum terdaftar', 'Teknisi menemukan aset yang belum ada di Data Aset. Daftarkan atau tautkan.', 'notifKe', 'usulan') +
    baris('bi-door-open', 'info', n.ruangan || 0, 'Ruangan baru', 'Ruangan yang diketik teknisi saat membuat WO dan belum ada di Master Lokasi.', 'notifKe', 'usulan') +
    baris('bi-building', 'info', n.vendor, 'Vendor baru', 'Nama vendor yang diketik saat membuat WO dan belum terdaftar.', 'notifKe', 'usulan') + '</div>',
    footer: '<button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>' });
};
ACTIONS.notifKe = function (el) { tutupModal(); if (el.dataset.p === 'usulan') { S.adminTab = 'usulan'; go('admin'); } else go('aset'); };
var USL = { d: null };
async function adminUsulan() {
  try { USL.d = await call('getUsulan'); } catch (e) { return gagal(e); }
  var m = USL.d.menunggu, r = USL.d.riwayat;
  var lokTeks = function (x) { var d = x.detail || {}; return [d.Gedung, d.Lantai, d.Ruangan, d.Lokasi_Detail].filter(Boolean).join(' › '); };
  var kartu = function (u) {
    if (u.Jenis === 'Vendor') return '<div class="usul-card"><div class="d-flex flex-wrap align-items-center gap-2"><span class="ic info"><i class="bi bi-building"></i></span><div class="flex-grow-1"><div class="fw-bold">' + esc(u.Nilai) + '</div>' +
      '<div class="small text-muted">Vendor baru · diketik oleh ' + esc(u.Diusulkan_Oleh) + ' pada WO ' + esc(u.ID_WO) + ' · ' + esc(fmtTgl(u.Diusulkan_Pada)) + '</div></div></div>' +
      '<div class="d-flex flex-wrap gap-2 mt-2"><button class="btn btn-primary btn-sm" data-act="usulTerimaVendor" data-id="' + esc(u.ID_Usulan) + '"><i class="bi bi-check2 me-1"></i>Daftarkan ke Data Master</button><button class="btn btn-light-x btn-sm text-danger" data-act="usulTolak" data-id="' + esc(u.ID_Usulan) + '">Tolak</button></div></div>';
    if (u.Jenis === 'Ruangan') return '<div class="usul-card"><div class="d-flex flex-wrap align-items-center gap-2"><span class="ic info"><i class="bi bi-door-open"></i></span><div class="flex-grow-1"><div class="fw-bold">' + esc(u.Nilai) + '</div>' +
      '<div class="small text-muted">Ruangan baru · ' + esc(lokTeks(u)) + '</div><div class="small text-muted">Diusulkan oleh ' + esc(u.Diusulkan_Oleh) + ' pada WO ' + esc(u.ID_WO) + ' · ' + esc(fmtTgl(u.Diusulkan_Pada)) + '</div></div></div>' +
      '<div class="d-flex flex-wrap gap-2 mt-2"><button class="btn btn-primary btn-sm" data-act="usulTerimaRuangan" data-id="' + esc(u.ID_Usulan) + '"><i class="bi bi-check2 me-1"></i>Daftarkan ke Master Lokasi</button><button class="btn btn-light-x btn-sm text-danger" data-act="usulTolak" data-id="' + esc(u.ID_Usulan) + '">Tolak</button></div></div>';
    return '<div class="usul-card"><div class="d-flex flex-wrap gap-3">' + (u.Foto_URL ? fotoThumb(u.Foto_URL, 'Foto aset') : '<div class="foto-th kosong"><span>Tanpa foto</span></div>') +
      '<div class="flex-grow-1" style="min-width:200px"><div class="fw-bold">' + esc(u.Nilai) + '</div><div class="small text-muted">Aset belum terdaftar · oleh ' + esc(u.Diusulkan_Oleh) + ' pada WO ' + esc(u.ID_WO) + ' · ' + esc(fmtTgl(u.Diusulkan_Pada)) + '</div>' +
      '<dl class="dl mt-2 mb-0">' + dlRow('Kategori', esc((u.detail || {}).Kategori_Aset)) + dlRow('Lokasi', esc(lokTeks(u))) + '</dl></div></div>' +
      '<div class="d-flex flex-wrap gap-2 mt-2"><button class="btn btn-primary btn-sm" data-act="usulDaftar" data-id="' + esc(u.ID_Usulan) + '"><i class="bi bi-plus-lg me-1"></i>Daftarkan aset</button><button class="btn btn-light-x btn-sm" data-act="usulTaut" data-id="' + esc(u.ID_Usulan) + '"><i class="bi bi-link-45deg me-1"></i>Tautkan ke aset yang ada</button><button class="btn btn-light-x btn-sm text-danger" data-act="usulTolak" data-id="' + esc(u.ID_Usulan) + '">Tolak</button></div></div>';
  };
  $('#adminBody').innerHTML = '<div class="card-x"><div class="card-h">Menunggu diproses <span class="badge-x ' + (m.length ? 'warn' : '') + '">' + m.length + '</span></div><div class="card-b">' +
    '<div class="small text-muted mb-3">Data yang diketik atau ditemukan teknisi/pembuat WO dan belum ada di Data Master / Data Aset. Terima agar masuk ke data resmi, atau tolak.</div>' +
    (m.length ? '<div class="d-grid gap-3">' + m.map(kartu).join('') + '</div>' : '<div class="empty"><i class="bi bi-check2-circle"></i>Tidak ada usulan yang menunggu.</div>') + '</div></div>' +
    (r.length ? '<div class="card-x mt-3"><div class="card-h">Riwayat terakhir</div><div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Jenis</th><th>Usulan</th><th>Status</th><th>Diproses</th></tr></thead><tbody>' + r.map(function (u) {
      return '<tr><td data-l="Jenis">' + esc(u.Jenis) + '</td><td data-l="Usulan"><span class="nm">' + esc(u.Nilai) + '</span><div class="sub">WO ' + esc(u.ID_WO) + (u.ID_Aset ? ' → ' + esc(u.ID_Aset) : '') + '</div></td><td data-l="Status">' + badge(u.Status, u.Status === 'Diterima' ? 'ok' : 'bad') + '</td><td data-l="Diproses">' + esc(u.Diproses_Oleh) + '<div class="sub">' + esc(fmtTgl(u.Diproses_Pada)) + '</div></td></tr>';
    }).join('') + '</tbody></table></div></div>' : '');
}
function usulanById(id) { return ((USL.d && USL.d.menunggu) || []).filter(function (x) { return x.ID_Usulan === id; })[0]; }
ACTIONS.usulTerimaVendor = async function (el) {
  var u = usulanById(el.dataset.id); if (!u) return;
  if (!(await konfirm('Daftarkan vendor?', '"' + u.Nilai + '" akan masuk ke Data Master → Nama Vendor.', 'Daftarkan', false))) return;
  sibuk(el, true);
  try { await call('prosesUsulan', u.ID_Usulan, 'terima', {}); var l = S.master.data['Nama Vendor'] = S.master.data['Nama Vendor'] || []; if (!l.some(function (x) { return x.nilai.toLowerCase() === u.Nilai.toLowerCase(); })) l.push({ nilai: u.Nilai, urutan: l.length + 1, keterangan: '' }); toast('Vendor didaftarkan.', 'ok'); await muatNotif(); adminUsulan(); }
  catch (e) { gagal(e); sibuk(el, false); }
};
ACTIONS.usulTerimaRuangan = async function (el) {
  var u = usulanById(el.dataset.id); if (!u) return;
  if (!(await konfirm('Daftarkan ruangan?', '"' + u.Nilai + '" akan masuk ke Master Lokasi (' + lokTeksUsul(u) + ').', 'Daftarkan', false))) return;
  sibuk(el, true);
  try { var d = await call('prosesUsulan', u.ID_Usulan, 'terima', {}); if (d && d.master) S.master = d.master; toast('Ruangan didaftarkan ke Master Lokasi.', 'ok'); await muatNotif(); adminUsulan(); }
  catch (e) { gagal(e); sibuk(el, false); }
};
function lokTeksUsul(x) { var d = x.detail || {}; return [d.Gedung, d.Lantai].filter(Boolean).join(' › '); }
ACTIONS.usulTolak = async function (el) {
  var u = usulanById(el.dataset.id); if (!u) return;
  if (!(await konfirm('Tolak usulan?', '"' + u.Nilai + '" tidak akan didaftarkan. Nama tetap tertulis apa adanya pada WO ' + u.ID_WO + '.', 'Tolak'))) return;
  try { await call('prosesUsulan', u.ID_Usulan, 'tolak', {}); toast('Usulan ditolak.', 'ok'); await muatNotif(); adminUsulan(); } catch (e) { gagal(e); }
};
ACTIONS.usulDaftar = async function (el) {
  var u = usulanById(el.dataset.id); if (!u) return; var d = u.detail || {};
  USL_TERKAIT = u.ID_Usulan;
  bukaFormAset(null, { Nama_Aset: d.Nama_Aset, Kategori_Aset: d.Kategori_Aset, Gedung: d.Gedung, Lantai: d.Lantai, Ruangan: d.Ruangan,
    Lokasi_Detail: d.Lokasi_Detail || '', Foto_Aset_URL: u.Foto_URL, Tahun_Perolehan: d.Tahun_Perolehan || '', NUP: d.NUP || '', Merk_Type: d.Merk_Type || '', Nomor_Seri: d.Nomor_Seri || '', Kondisi: d.Kondisi || '', Status_Aset: d.Status_Aset || '',
    _hint: 'Diisi dari laporan teknisi pada WO <b>' + esc(u.ID_WO) + '</b>. Lengkapi data lain secara bertahap (Tahun, Merk, Seri, Nilai, dan Spesifikasi boleh kosong).' });
};
ACTIONS.usulTaut = function (el) {
  var u = usulanById(el.dataset.id); if (!u) return;
  bukaModal({ judul: 'Tautkan ke aset yang ada', size: 'md', isi: '<form id="fTaut" data-submit="usulTautKirim" data-id="' + esc(u.ID_Usulan) + '" novalidate><div class="alert-x info small">WO <b>' + esc(u.ID_WO) + '</b> (' + esc(u.Nilai) + ') akan ditautkan ke aset yang Anda pilih.</div>' +
    '<label class="form-label req">Aset</label><input class="form-control" name="aset" list="dlAset" placeholder="Ketik ID atau nama aset…" autocomplete="off"><datalist id="dlAset">' +
    S.aset.filter(function (a) { return a.Status_Aset !== DIHAPUS; }).slice(0, 1500).map(function (a) { return '<option value="' + esc(a.ID_Aset + ' — ' + a.Nama_Aset) + '">'; }).join('') + '</datalist>' +
    '<div class="alert-x danger d-none mt-3 mb-0" id="tautErr"></div></form>', footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fTaut">Tautkan</button>' });
};
SUBMITS.usulTautKirim = async function (f) {
  var er = $('#tautErr'); er.classList.add('d-none');
  var id = String(f.aset.value).split(' — ')[0].trim(); if (!cariAset(id)) { er.textContent = 'Pilih aset dari daftar yang muncul saat mengetik.'; er.classList.remove('d-none'); return; }
  var btn = $('button[type=submit][form=fTaut]'); sibuk(btn, true);
  try { await call('prosesUsulan', f.dataset.id, 'terima', { ID_Aset: id }); tutupModal(); toast('WO ditautkan ke ' + id + '.', 'ok'); await muatNotif(); adminUsulan(); }
  catch (e) { er.textContent = e.message; er.classList.remove('d-none'); } finally { sibuk(btn, false); }
};

/* Pengguna */
var daftarUser = [];
async function adminUsers() {
  try { daftarUser = await call('listUsers'); } catch (e) { return gagal(e); }
  $('#adminBody').innerHTML = '<div class="card-x"><div class="card-h">Pengguna <span class="badge-x">' + daftarUser.length + '</span><button class="btn btn-primary btn-sm ms-auto" data-act="userBaru"><i class="bi bi-plus-lg me-1"></i>Tambah</button></div>' +
    '<div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Nama</th><th>Username</th><th>Peran / Jabatan</th><th>Status</th><th>Login terakhir</th><th></th></tr></thead><tbody>' +
    daftarUser.map(function (u) {
      return '<tr><td data-l="Nama"><span class="nm">' + esc(u.Nama_Lengkap) + '</span>' + (u.No_WA ? '<div class="sub"><i class="bi bi-whatsapp me-1"></i>' + esc(u.No_WA) + '</div>' : '') + '</td><td data-l="Username">' + esc(u.Username) + '</td><td data-l="Peran">' + badge(u.Peran, 'info') + (u.Jabatan ? '<div class="sub mt-1">' + esc(u.Jabatan) + '</div>' : (u.Peran === 'PIC' ? '<div class="sub mt-1 text-warning">Belum ada jabatan</div>' : '')) + '</td><td data-l="Status">' + badge(u.Status, u.Status === 'Aktif' ? 'ok' : '') + '</td><td data-l="Login terakhir">' + esc(u.Login_Terakhir || '—') + '</td>' +
        '<td class="acts"><button class="icon-btn" data-act="userUbah" data-id="' + esc(u.ID_User) + '" title="Ubah"><i class="bi bi-pencil"></i></button><button class="icon-btn" data-act="userReset" data-id="' + esc(u.ID_User) + '" title="Reset password"><i class="bi bi-key"></i></button>' +
        (u.ID_User === S.user.id ? '' : '<button class="icon-btn text-danger" data-act="userHapus" data-id="' + esc(u.ID_User) + '" title="Hapus pengguna" aria-label="Hapus ' + esc(u.Username) + '"><i class="bi bi-trash"></i></button>') + '</td></tr>';
    }).join('') + '</tbody></table></div></div>';
}
function opsiJabatan(cur) {
  var l = listMaster('Jabatan PIC').slice(), h = '<option value="">— pilih jabatan —</option>';
  if (cur && l.indexOf(cur) < 0) l.push(cur);
  return h + l.map(function (v) { return '<option value="' + esc(v) + '"' + (v === cur ? ' selected' : '') + '>' + esc(v) + (listMaster('Jabatan PIC').indexOf(v) < 0 ? ' (belum ada di Data Master)' : '') + '</option>'; }).join('');
}
function formUser(u) {
  var baru = !u; u = u || {};
  return '<form id="fUser" data-submit="simpanUser" novalidate><input type="hidden" name="mode" value="' + (baru ? 'baru' : 'ubah') + '"><input type="hidden" name="ID_User" value="' + esc(u.ID_User || '') + '">' +
    '<div class="mb-3"><label class="form-label req">Nama lengkap</label><input class="form-control" name="Nama_Lengkap" value="' + esc(u.Nama_Lengkap) + '" maxlength="100" required></div>' +
    '<div class="mb-3"><label class="form-label req">Username</label><input class="form-control" name="Username" value="' + esc(u.Username) + '" maxlength="30" autocapitalize="none" spellcheck="false" required><div class="form-text">Huruf kecil, angka, titik, - atau _ (3–30 karakter).' + (baru ? '' : ' Jika diubah, pengguna tersebut harus login ulang.') + '</div></div>' +
    (baru ? '<div class="mb-3"><label class="form-label req">Password awal</label><input class="form-control" name="Password" type="text" minlength="8" autocomplete="off" required><div class="form-text">Minimal 8 karakter. Sampaikan ke pengguna secara aman.</div></div>' : '') +
    '<div class="row g-3"><div class="col-6"><label class="form-label req">Peran</label><select class="form-select" name="Peran">' + opsi(PERAN, u.Peran || 'Teknisi') + '</select></div>' +
    '<div class="col-6"><label class="form-label req">Status</label><select class="form-select" name="Status">' + opsi(['Aktif', 'Nonaktif'], u.Status || 'Aktif') + '</select></div></div>' +
    '<div class="mt-3"><label class="form-label">No. WhatsApp</label><input class="form-control" name="No_WA" value="' + esc(u.No_WA) + '" maxlength="20" inputmode="tel" placeholder="081234567890"><div class="form-text">Untuk notifikasi otomatis (pengingat PM H-1 dan risiko tinggi bagi Teknisi). Boleh dikosongkan.</div></div>' +
    '<div class="mt-3' + ((u.Peran || 'Teknisi') === 'PIC' ? '' : ' d-none') + '" id="wrapJabatan"><label class="form-label req">Jabatan PIC</label><select class="form-select" name="Jabatan">' + opsiJabatan(u.Jabatan) + '</select><div class="form-text">' + (listMaster('Jabatan PIC').length ? '' : '<span class="text-warning">Belum ada pilihan jabatan. Tambahkan dulu di Data Master → Jabatan PIC.</span> ') + 'Daftar jabatan dikelola di <b>Data Master → Jabatan PIC</b>. Aset memakai <b>jabatan</b> ini sebagai PIC. Tampilan di semua pilihan: <i>Nama - Jabatan</i>. Jika jabatan dipindah ke orang lain, seluruh aset ikut berganti nama PIC, sedangkan riwayat WO/PM yang sudah selesai tetap memakai nama lama.</div></div>' +
    '<div class="alert-x danger d-none mt-3 mb-0" id="usrErr"></div></form>';
}
document.addEventListener('change', function (e) { var t = e.target; if (t.name === 'Peran' && t.closest('#fUser')) { var w = $('#wrapJabatan'); if (w) w.classList.toggle('d-none', t.value !== 'PIC'); } });
ACTIONS.userBaru = function () { bukaModal({ judul: 'Tambah pengguna', size: 'md', isi: formUser(null), footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fUser">Simpan</button>' }); };
ACTIONS.userUbah = function (el) { var u = daftarUser.filter(function (x) { return x.ID_User === el.dataset.id; })[0]; if (u) bukaModal({ judul: 'Ubah pengguna', size: 'md', isi: formUser(u), footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fUser">Simpan</button>' }); };
SUBMITS.simpanUser = async function (f) {
  var err = $('#usrErr'); err.classList.add('d-none');
  var p = { ID_User: f.ID_User.value, Nama_Lengkap: f.Nama_Lengkap.value, Username: f.Username.value, Peran: f.Peran.value, Status: f.Status.value, Jabatan: f.Jabatan.value, No_WA: f.No_WA.value };
  if (f.Password) p.Password = f.Password.value;
  if (p.Peran === 'PIC' && !p.Jabatan.trim()) { err.textContent = 'Jabatan wajib diisi untuk akun PIC.'; err.classList.remove('d-none'); return; }
  var lama = daftarUser.filter(function (x) { return x.ID_User === p.ID_User; })[0];
  var ubahSendiri = f.mode.value === 'ubah' && p.ID_User === S.user.id && p.Username.trim().toLowerCase() !== String(S.user.username).toLowerCase();
  var btn = $('button[type=submit][form=fUser]'); sibuk(btn, true);
  try {
    var jab = p.Jabatan.trim(), jabLama = lama ? lama.Jabatan : '';
    // Ganti nama jabatan sendiri (bukan memindahkan dari orang lain): tanyakan apakah aset ikut berpindah
    if (lama && p.Peran === 'PIC' && jabLama && jab && jab !== jabLama && !(S.pic || []).some(function (x) { return x.jabatan.toLowerCase() === jab.toLowerCase() && x.nama !== lama.Nama_Lengkap; })) {
      var n = S.aset.filter(function (a) { return a.PIC === jabLama; }).length;
      if (n && await konfirm('Aset ikut pindah jabatan?', n + ' aset saat ini berjabatan "' + jabLama + '". Pilih "Ya" bila ini hanya perubahan nama jabatan (aset ikut menjadi "' + jab + '"). Pilih "Tidak" bila jabatan lama tetap ada untuk orang lain.', 'Ya, pindahkan aset', false)) p.RenameJabatan = true;
    }
    var hasil;
    try { hasil = await call('saveUser', p, f.mode.value); }
    catch (e1) {
      if (e1.code !== 'JABATAN') throw e1;
      if (!(await konfirm('Pindahkan jabatan?', e1.message + ' Orang tersebut akan menjadi tanpa jabatan; seluruh aset berjabatan ini otomatis menampilkan nama PIC yang baru.', 'Ya, pindahkan', false))) { return; }
      p.PindahJabatan = true; hasil = await call('saveUser', p, f.mode.value);
    }
    tutupModal(); toast('Pengguna disimpan.', 'ok');
    try { S.pic = await call('getPicList'); } catch (e2) {}
    if (p.RenameJabatan) { try { S.aset = await call('getAssetList'); } catch (e3) {} }
    if (ubahSendiri) { toast('Username Anda berubah. Silakan login ulang.', 'warn', 6000); ACTIONS.logout(); return; }
    adminUsers();
  } catch (e) { err.textContent = e.message; err.classList.remove('d-none'); } finally { sibuk(btn, false); }
};
ACTIONS.userHapus = async function (el) {
  var u = daftarUser.filter(function (x) { return x.ID_User === el.dataset.id; })[0]; if (!u) return;
  if (!(await konfirm('Hapus pengguna?', u.Nama_Lengkap + ' (@' + u.Username + ') akan dihapus permanen dan tidak dapat login lagi. Untuk menonaktifkan sementara, gunakan Ubah → Status Nonaktif.', 'Hapus'))) return;
  try { await call('deleteUser', u.ID_User); toast('Pengguna dihapus.', 'ok'); adminUsers(); } catch (e) { gagal(e); }
};
ACTIONS.userReset = function (el) {
  var u = daftarUser.filter(function (x) { return x.ID_User === el.dataset.id; })[0]; if (!u) return;
  bukaModal({ judul: 'Reset password — ' + u.Username, size: 'sm',
    isi: '<form id="fReset" data-submit="resetPw"><input type="hidden" name="id" value="' + esc(u.ID_User) + '"><label class="form-label req">Password baru</label><input class="form-control" name="pw" minlength="8" autocomplete="off" required><div class="form-text">Pengguna harus login ulang dengan password ini.</div><div class="alert-x danger d-none mt-3 mb-0" id="rsErr"></div></form>',
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fReset">Reset</button>' });
};
SUBMITS.resetPw = async function (f) {
  var btn = $('button[type=submit][form=fReset]'); sibuk(btn, true);
  try { await call('resetUserPassword', f.id.value, f.pw.value); tutupModal(); toast('Password direset.', 'ok'); adminUsers(); }
  catch (e) { var er = $('#rsErr'); er.textContent = e.message; er.classList.remove('d-none'); } finally { sibuk(btn, false); }
};

/* Data Master */
function adminMaster() {
  var kel = S.adminKel, rows = (S.master.data[kel] || []), nama = KELOMPOK_NAMA.indexOf(kel) >= 0, ket = kel === 'Prioritas' ? 'Durasi SLA (jam)' : (kel === 'Target KPI' ? 'Target (%)' : (kel === 'Level Risiko' ? 'Rentang skor' : 'Keterangan'));
  $('#adminBody').innerHTML = '<div class="card-x"><div class="toolbar"><select class="form-select fsel" style="flex-basis:260px" id="selKel" aria-label="Kelompok">' + opsi(KELOMPOK.slice().sort(function (a, b) { return a.localeCompare(b, 'id'); }), kel) + '</select>' +
    '<button class="btn btn-primary btn-sm ms-auto" data-act="masterBaru"><i class="bi bi-plus-lg me-1"></i>Tambah nilai</button></div>' +
    '<div class="tbl-wrap"><table class="tbl"><thead><tr><th style="width:84px">Urutan</th><th>' + (nama ? 'Nama' : 'Nilai') + '</th><th>' + ket + '</th><th></th></tr></thead><tbody id="msBody">' +
    (rows.length ? rows.map(function (r, i) {
      return '<tr draggable="true" data-i="' + i + '" data-nilai="' + esc(r.nilai) + '"><td class="drag-cell"><span class="drag-h" title="Seret untuk mengubah urutan" aria-hidden="true"><i class="bi bi-grip-vertical"></i></span>' +
        '<span class="drag-no">' + (i + 1) + '</span></td><td class="nm">' + esc(r.nilai) + '</td><td>' + esc(r.keterangan || '—') + '</td><td class="acts">' +
        '<button class="icon-btn" data-act="masterNaik" data-i="' + i + '" title="Naikkan" aria-label="Naikkan"' + (i === 0 ? ' disabled' : '') + '><i class="bi bi-arrow-up"></i></button>' +
        '<button class="icon-btn" data-act="masterTurun" data-i="' + i + '" title="Turunkan" aria-label="Turunkan"' + (i === rows.length - 1 ? ' disabled' : '') + '><i class="bi bi-arrow-down"></i></button>' +
        '<button class="icon-btn" data-act="masterUbah" data-i="' + i + '" title="Ubah"><i class="bi bi-pencil"></i></button><button class="icon-btn text-danger" data-act="masterHapus" data-i="' + i + '" title="Hapus"><i class="bi bi-trash"></i></button></td></tr>';
    }).join('') : '<tr><td colspan="4"><div class="empty">' + (nama ? 'Belum ada nama. Klik "Tambah nilai" untuk mengisi.' : 'Belum ada nilai.') + '</div></td></tr>') + '</tbody></table></div>' +
    '<div class="pager">Nilai baru otomatis ditaruh di urutan terakhir. Seret baris (ikon <i class="bi bi-grip-vertical"></i>) atau gunakan tombol panah untuk mengubah urutan; tersimpan otomatis. Nilai yang sedang dipakai aset tidak dapat diganti nama atau dihapus.</div></div>';
}
async function simpanUrutMaster(daftarNilai) {
  try { S.master = await call('reorderMasterData', S.adminKel, daftarNilai); toast('Urutan disimpan.', 'ok', 1800); }
  catch (e) { gagal(e); try { S.master = (await call('getBootstrapData')).master; } catch (e2) {} }
  adminMaster();
}
var dragNilai = null;
document.addEventListener('dragstart', function (e) { var tr = e.target.closest && e.target.closest('#msBody tr[data-nilai]'); if (!tr) return; dragNilai = tr.dataset.nilai; tr.classList.add('dragging'); try { e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', dragNilai); } catch (x) {} });
document.addEventListener('dragover', function (e) {
  var tr = e.target.closest && e.target.closest('#msBody tr[data-nilai]'); if (!tr || dragNilai === null) return;
  e.preventDefault(); $$('#msBody tr').forEach(function (r) { r.classList.remove('drop-atas', 'drop-bawah'); });
  var rc = tr.getBoundingClientRect(); tr.classList.add(e.clientY < rc.top + rc.height / 2 ? 'drop-atas' : 'drop-bawah');
});
document.addEventListener('dragend', function () { dragNilai = null; $$('#msBody tr').forEach(function (r) { r.classList.remove('dragging', 'drop-atas', 'drop-bawah'); }); });
document.addEventListener('drop', function (e) {
  var tr = e.target.closest && e.target.closest('#msBody tr[data-nilai]'); if (!tr || dragNilai === null) return;
  e.preventDefault();
  var rc = tr.getBoundingClientRect(), sesudah = e.clientY >= rc.top + rc.height / 2;
  var daftar = (S.master.data[S.adminKel] || []).map(function (r) { return r.nilai; }), dari = daftar.indexOf(dragNilai), ke = daftar.indexOf(tr.dataset.nilai);
  dragNilai = null; if (dari < 0 || ke < 0 || dari === ke) { adminMaster(); return; }
  daftar.splice(dari, 1); ke = daftar.indexOf(tr.dataset.nilai); daftar.splice(sesudah ? ke + 1 : ke, 0, S.master.data[S.adminKel][dari].nilai);
  simpanUrutMaster(daftar);
});
function geserMaster(i, d) {
  var daftar = (S.master.data[S.adminKel] || []).map(function (r) { return r.nilai; }), j = i + d;
  if (j < 0 || j >= daftar.length) return; var t = daftar[i]; daftar[i] = daftar[j]; daftar[j] = t; simpanUrutMaster(daftar);
}
ACTIONS.masterNaik = function (el) { geserMaster(Number(el.dataset.i), -1); };
ACTIONS.masterTurun = function (el) { geserMaster(Number(el.dataset.i), 1); };
document.addEventListener('change', function (e) { if (e.target.id === 'selKel') { S.adminKel = e.target.value; adminMaster(); } });
function formMaster(r) {
  var kel = S.adminKel, perluAngka = kel === 'Prioritas' || kel === 'Target KPI', nama = KELOMPOK_NAMA.indexOf(kel) >= 0;
  return '<form id="fMaster" data-submit="simpanMaster"><input type="hidden" name="lama" value="' + esc(r ? r.nilai : '') + '">' +
    '<div class="mb-3"><label class="form-label">Kelompok</label><input class="form-control" value="' + esc(kel) + '" readonly></div>' +
    '<div class="mb-3"><label class="form-label req">' + (nama ? 'Nama' : 'Nilai') + '</label><input class="form-control" name="Nilai" value="' + esc(r ? r.nilai : '') + '" maxlength="100" required></div>' +
    '<div class="mb-1"><label class="form-label' + (perluAngka ? ' req' : '') + '">Keterangan' + (kel === 'Prioritas' ? ' (SLA, jam)' : '') + '</label><input class="form-control" name="Keterangan" value="' + esc(r ? r.keterangan : '') + '" maxlength="200"' + (perluAngka ? ' inputmode="decimal" required' : '') + '></div>' +
    (r ? '' : '<div class="form-text">Nilai baru ditaruh di urutan terakhir. Atur urutan dengan menyeret barisnya setelah disimpan.</div>') +
    '<div class="alert-x danger d-none mt-3 mb-0" id="msErr"></div></form>';
}
ACTIONS.masterBaru = function () { bukaModal({ judul: 'Tambah nilai master', size: 'md', isi: formMaster(null), footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fMaster">Simpan</button>' }); };
ACTIONS.masterUbah = function (el) { var r = (S.master.data[S.adminKel] || [])[Number(el.dataset.i)]; if (r) bukaModal({ judul: 'Ubah nilai master', size: 'md', isi: formMaster(r), footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fMaster">Simpan</button>' }); };
SUBMITS.simpanMaster = async function (f) {
  var btn = $('button[type=submit][form=fMaster]'); sibuk(btn, true);
  try {
    var lama = f.lama.value ? { Kelompok: S.adminKel, Nilai: f.lama.value } : null;
    S.master = await call('saveMasterData', lama, { Kelompok: S.adminKel, Nilai: f.Nilai.value, Keterangan: f.Keterangan.value });
    tutupModal(); toast('Data master disimpan.', 'ok'); adminMaster();
  } catch (e) { var er = $('#msErr'); er.textContent = e.message; er.classList.remove('d-none'); } finally { sibuk(btn, false); }
};
ACTIONS.masterHapus = async function (el) {
  var r = (S.master.data[S.adminKel] || [])[Number(el.dataset.i)]; if (!r) return;
  if (!(await konfirm('Hapus nilai?', '"' + r.nilai + '" dari ' + S.adminKel + '.', 'Hapus'))) return;
  try { S.master = await call('deleteMasterData', S.adminKel, r.nilai); toast('Dihapus.', 'ok'); adminMaster(); } catch (e) { gagal(e); }
};

/* Master Lokasi */
var lokHal = 1;
function adminLokasi() {
  var data = S.master.lokasi, per = 25, hal = Math.max(1, Math.ceil(data.length / per)); if (lokHal > hal) lokHal = hal;
  var mulai = (lokHal - 1) * per;
  $('#adminBody').innerHTML = '<div class="card-x"><div class="card-h">Master Lokasi <span class="badge-x">' + data.length + '</span><div class="ms-auto d-flex gap-2"><button class="btn btn-light-x btn-sm" data-act="lokImpor"><i class="bi bi-upload me-1"></i>Impor CSV</button><button class="btn btn-primary btn-sm" data-act="lokBaru"><i class="bi bi-plus-lg me-1"></i>Tambah</button></div></div>' +
    '<div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Gedung</th><th>Lantai</th><th>Ruangan</th><th></th></tr></thead><tbody>' +
    (data.length ? data.slice(mulai, mulai + per).map(function (l, i) {
      return '<tr><td data-l="Gedung">' + esc(l.G) + '</td><td data-l="Lantai">' + esc(l.L || '—') + '</td><td data-l="Ruangan">' + esc(l.R || '—') + '</td><td class="acts"><button class="icon-btn" data-act="lokUbah" data-i="' + (mulai + i) + '"><i class="bi bi-pencil"></i></button><button class="icon-btn text-danger" data-act="lokHapus" data-i="' + (mulai + i) + '"><i class="bi bi-trash"></i></button></td></tr>';
    }).join('') : '<tr><td colspan="4"><div class="empty">Belum ada lokasi.</div></td></tr>') + '</tbody></table></div>' +
    '<div class="pager"><div>Lokasi yang dipakai aset tidak dapat diubah atau dihapus.</div><div class="d-flex align-items-center gap-2"><button class="btn btn-light-x btn-sm" data-act="lokHal" data-d="-1"' + (lokHal <= 1 ? ' disabled' : '') + '><i class="bi bi-chevron-left"></i></button><span>Hal. ' + lokHal + ' / ' + hal + '</span><button class="btn btn-light-x btn-sm" data-act="lokHal" data-d="1"' + (lokHal >= hal ? ' disabled' : '') + '><i class="bi bi-chevron-right"></i></button></div></div></div>';
}
ACTIONS.lokHal = function (el) { lokHal += Number(el.dataset.d); adminLokasi(); };
function formLokasi(l, idx) {
  l = l || {};
  var f = function (lbl, n, v, r) { return '<div class="col-md-6"><label class="form-label' + (r ? ' req' : '') + '">' + lbl + '</label><input class="form-control" name="' + n + '" value="' + esc(v) + '" maxlength="150"' + (r ? ' required' : '') + '></div>'; };
  return '<form id="fLok" data-submit="simpanLok"><input type="hidden" name="idx" value="' + (idx == null ? '' : idx) + '"><div class="row g-3">' +
    f('Gedung', 'G', l.G, 1) + f('Lantai (boleh kosong)', 'L', l.L, 0) + f('Ruangan', 'R', l.R, 1) +
    '<div class="col-12 form-text mt-0">Lokasi Detail tidak diisi di sini. Isi pada setiap aset di menu Data Aset, karena berbeda-beda untuk tiap aset.</div></div><div class="alert-x danger d-none mt-3 mb-0" id="lkErr"></div></form>';
}
ACTIONS.lokBaru = function () { bukaModal({ judul: 'Tambah lokasi', isi: formLokasi(null), footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fLok">Simpan</button>' }); };
ACTIONS.lokUbah = function (el) { var i = Number(el.dataset.i); bukaModal({ judul: 'Ubah lokasi', isi: formLokasi(S.master.lokasi[i], i), footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fLok">Simpan</button>' }); };
SUBMITS.simpanLok = async function (f) {
  var baru = { G: f.G.value, L: f.L.value, R: f.R.value }, lama = f.idx.value === '' ? null : S.master.lokasi[Number(f.idx.value)];
  var btn = $('button[type=submit][form=fLok]'); sibuk(btn, true);
  try { S.master = await call('saveLokasi', lama, baru); tutupModal(); toast('Lokasi disimpan.', 'ok'); adminLokasi(); }
  catch (e) { var er = $('#lkErr'); er.textContent = e.message; er.classList.remove('d-none'); } finally { sibuk(btn, false); }
};
ACTIONS.lokHapus = async function (el) {
  var l = S.master.lokasi[Number(el.dataset.i)]; if (!l) return;
  if (!(await konfirm('Hapus lokasi?', [l.G, l.L, l.R].filter(Boolean).join(' / '), 'Hapus'))) return;
  try { S.master = await call('deleteLokasi', l); toast('Lokasi dihapus.', 'ok'); adminLokasi(); } catch (e) { gagal(e); }
};
ACTIONS.lokImpor = function () {
  bukaModal({ judul: 'Impor Master Lokasi (CSV)', size: 'lg',
    isi: '<div class="alert-x info small">Kolom: <code>Gedung, Lantai, Ruangan</code> (Lantai boleh kosong; kolom Lokasi_Detail, jika ada, diabaikan). Duplikat dilewati. Maksimal 2.000 baris.</div>' +
      '<button class="btn btn-light-x btn-sm mb-3" data-act="templateLok"><i class="bi bi-download me-1"></i>Unduh template</button><input type="file" class="form-control" id="lokCsv" accept=".csv,text/csv"><div id="lokHasil" class="mt-3"></div>',
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Tutup</button>' });
};
ACTIONS.templateLok = function () { unduhBlob('template_lokasi.csv', new Blob([csvKeString([['Gedung', 'Lantai', 'Ruangan'], ['Gedung Utama', 'Lantai 2', 'Ruang Rapat']])], { type: 'text/csv;charset=utf-8' })); };
document.addEventListener('change', async function (e) {
  if (e.target.id !== 'lokCsv' || !e.target.files[0]) return;
  var box = $('#lokHasil'); box.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Mengimpor…';
  try {
    var t = parseCsv(await bacaFile(e.target.files[0], true)); if (t.length < 2) throw new Error('File kosong.'); if (t.length > 2001) throw new Error('Maksimal 2.000 baris.');
    var map = {}; t[0].forEach(function (h, i) { map[String(h).trim().toLowerCase().replace(/\s+/g, '_')] = i; });
    ['gedung', 'ruangan'].forEach(function (k) { if (map[k] === undefined) throw new Error('Kolom ' + k + ' tidak ditemukan.'); });
    var g = function (r, k) { return map[k] === undefined ? '' : String(r[map[k]] || '').trim(); };
    var rows = t.slice(1).map(function (r) { return { Gedung: g(r, 'gedung'), Lantai: g(r, 'lantai'), Ruangan: g(r, 'ruangan'), Lokasi_Detail: g(r, 'lokasi_detail') }; });
    var tot = { berhasil: 0, duplikat: 0, gagal: [] };
    for (var i = 0; i < rows.length; i += 200) { var d = await call('importLokasi', rows.slice(i, i + 200)); tot.berhasil += d.berhasil; tot.duplikat += d.duplikat; tot.gagal = tot.gagal.concat(d.gagal); }
    var m = await call('getBootstrapData'); S.master = m.master; adminLokasi();
    box.innerHTML = '<div class="alert-x ok">' + tot.berhasil + ' lokasi ditambahkan, ' + tot.duplikat + ' duplikat dilewati.</div>' + (tot.gagal.length ? '<div class="alert-x danger">' + tot.gagal.length + ' baris gagal.<br>' + tot.gagal.slice(0, 15).map(function (x) { return esc(JSON.stringify(x)); }).join('<br>') + '</div>' : '');
  } catch (er) { box.innerHTML = '<div class="alert-x danger">' + esc(er.message) + '</div>'; }
});

/* Pengaturan */
function adminPengaturan() {
  $('#adminBody').innerHTML = '<div class="card-x"><div class="card-h">Pengaturan aplikasi</div><div class="card-b"><form id="fCfg" data-submit="simpanCfg" style="max-width:640px">' +
    '<div class="mb-3"><label class="form-label req">Nama aplikasi</label><input class="form-control" name="appName" value="' + esc(S.cfg.appName) + '" maxlength="80" required></div>' +
    '<input type="hidden" name="logoUrl" value="' + esc(S.cfg.logoUrl) + '">' +
    '<div class="mb-3"><label class="form-label">URL Web App (untuk QR)</label><input class="form-control" name="webAppUrl" value="' + esc(S.cfg.webAppUrl) + '" maxlength="500" placeholder="https://script.google.com/macros/s/…/exec"><div class="form-text">Wajib tepat sama dengan URL deployment yang dibagikan. Gunakan URL yang berakhiran <code>/exec</code>.</div></div>' +
    '<div class="alert-x danger d-none" id="cfgErr"></div><div class="d-flex flex-wrap gap-2"><button class="btn btn-primary" type="submit">Simpan</button>' +
    '<button class="btn btn-light-x" type="button" data-act="regenQr"><i class="bi bi-qr-code me-1"></i>Perbarui URL QR semua aset</button></div></form></div></div>' +
    '<div class="card-x mt-3"><div class="card-h">Sambungan spreadsheet pengaduan</div><div class="card-b" id="cfgPeng"><div class="skeleton" style="height:120px"></div></div></div>' +
    '<div class="card-x mt-3"><div class="card-h"><i class="bi bi-whatsapp me-2"></i>Notifikasi WhatsApp (Fonnte)</div><div class="card-b" id="cfgWa"><div class="skeleton" style="height:160px"></div></div></div>';
  adminPengaduan(); adminWa();
}
async function adminWa() {
  var box = $('#cfgWa'); if (!box) return;
  var c; try { c = await call('getWaConfig'); } catch (e) { box.innerHTML = '<div class="alert-x danger">' + esc(e.message) + '</div>'; return; }
  box.innerHTML = '<form id="fWa" data-submit="simpanWa" style="max-width:640px">' +
    '<div class="form-check form-switch mb-3"><input class="form-check-input" type="checkbox" role="switch" name="aktif" id="waAktif"' + (c.aktif ? ' checked' : '') + '><label class="form-check-label" for="waAktif">Kirim notifikasi WhatsApp otomatis</label></div>' +
    '<div class="mb-3"><label class="form-label">Token Fonnte</label><input class="form-control" name="token" type="password" autocomplete="off" placeholder="' + (c.tokenAda ? '•••• ' + esc(c.tokenAkhir) + ' (kosongkan bila tidak diganti)' : 'Tempel token perangkat Fonnte') + '"><div class="form-text">Disimpan di properti skrip (tidak di sheet).</div></div>' +
    '<div class="mb-3"><label class="form-label">ID grup WhatsApp</label><input class="form-control" name="grup" value="' + esc(c.grup) + '" maxlength="60" placeholder="120363123456789012@g.us"><div class="form-text">Tujuan notifikasi Work Order baru dan risiko tinggi. Nomor perangkat Fonnte harus sudah menjadi anggota grup.</div></div>' +
    '<div class="small text-muted mb-3">Pengingat PM (H-1) dan peringatan risiko (skor 10 ke atas atau kategori Listrik) dikirim ke grup saja, tidak ke nomor tiap Teknisi.</div>' +
    '<div class="alert-x danger d-none" id="waErr"></div><div class="d-flex flex-wrap gap-2 mb-3"><button class="btn btn-primary" type="submit">Simpan</button></div></form>' +
    '<div class="row g-2 align-items-end mb-3" style="max-width:640px"><div class="col-sm-6"><label class="form-label">Nomor untuk tes (opsional)</label><input class="form-control" id="waTes" placeholder="081234567890" inputmode="tel"></div>' +
    '<div class="col-sm-6"><button class="btn btn-light-x" data-act="ujiWa"><i class="bi bi-send me-1"></i>Kirim pesan tes</button></div></div>' +
    (c.log.length ? '<div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Waktu</th><th>Jenis</th><th>Tujuan</th><th>Status</th></tr></thead><tbody>' + c.log.map(function (l) {
      return '<tr><td data-l="Waktu">' + esc(fmtTgl(l.Waktu)) + '</td><td data-l="Jenis">' + esc(l.Jenis) + '</td><td data-l="Tujuan">' + esc(l.Target) + '</td><td data-l="Status">' + badge(l.Status === 'OK' ? 'Terkirim' : 'Gagal', l.Status === 'OK' ? 'ok' : 'bad') + (l.Status === 'OK' ? '' : '<div class="sub">' + esc(l.Keterangan) + '</div>') + '</td></tr>';
    }).join('') + '</tbody></table></div>' : '<div class="small text-muted">Belum ada riwayat pengiriman.</div>');
}
SUBMITS.simpanWa = async function (f) {
  var err = $('#waErr'); err.classList.add('d-none');
  try { await call('saveWaConfig', { aktif: f.aktif.checked, grup: f.grup.value, token: f.token.value }); toast('Pengaturan WhatsApp disimpan.', 'ok'); adminWa(); }
  catch (e) { err.textContent = e.message; err.classList.remove('d-none'); }
};
ACTIONS.ujiWa = async function (el) {
  sibuk(el, true, 'Mengirim…');
  try { var d = await call('ujiWa', { nomor: ($('#waTes') || {}).value || '' }); toast(d.hasil.join(' · '), /Terkirim/.test(d.hasil.join()) ? 'ok' : 'warn'); adminWa(); }
  catch (e) { gagal(e); } finally { sibuk(el, false); }
};
async function adminPengaduan() {
  var el = $('#cfgPeng'); if (!el) return;
  try {
    var c = await call('getPengaduanConfig'); if (!$('#cfgPeng')) return;
    el.innerHTML = '<form id="fPeng" data-submit="simpanPeng" style="max-width:640px"><div class="small text-muted mb-3">Pengaduan dibaca dari sheet ini (hanya baris dengan Tindak Lanjut kosong). Akun pemilik aplikasi harus punya akses <b>Editor</b> ke spreadsheet tersebut.</div>' +
      '<div class="mb-3"><label class="form-label req">Tautan spreadsheet pengaduan</label><input class="form-control" name="pengaduanSheetUrl" value="' + esc(c.pengaduanSheetUrl) + '" maxlength="400" placeholder="https://docs.google.com/spreadsheets/d/…/edit" spellcheck="false"><div class="form-text">Salin tautan lengkap Google Sheets. Ganti di sini bila tautan atau nama tab berubah, lalu klik <b>Uji sambungan</b>.</div></div>' +
      '<div class="row g-3 mb-3"><div class="col-md-6"><label class="form-label req">Nama tab</label><input class="form-control" name="pengaduanTab" value="' + esc(c.pengaduanTab) + '" maxlength="60"></div>' +
      '<div class="col-md-6"><label class="form-label req">Status pengaduan baru</label><input class="form-control" name="pengaduanStatusBaru" value="' + esc(c.pengaduanStatusBaru) + '" maxlength="40"><div class="form-text"><code>(kosong)</code> = Tindak Lanjut belum dipilih.</div></div>' +
      '<div class="col-md-6"><label class="form-label req">Status saat WO dibuat</label><input class="form-control" name="pengaduanStatusProses" value="' + esc(c.pengaduanStatusProses) + '" maxlength="40"><div class="form-text">Ditulis ke kolom &quot;Tindak Lanjut&quot;. Sesuaikan dengan isi dropdown di sheet.</div></div>' +
      '<div class="col-md-6"><label class="form-label req">Status saat WO selesai</label><input class="form-control" name="pengaduanStatusSelesai" value="' + esc(c.pengaduanStatusSelesai) + '" maxlength="40"></div></div>' +
      '<div class="alert-x danger d-none" id="pengErr"></div><div class="d-flex flex-wrap gap-2"><button class="btn btn-primary" type="submit">Simpan</button><button class="btn btn-light-x" type="button" data-act="ujiPeng"><i class="bi bi-plug me-1"></i>Uji sambungan</button></div></form>';
  } catch (e) { el.innerHTML = '<div class="alert-x danger mb-0">' + esc(e.message) + '</div>'; }
}
SUBMITS.simpanPeng = async function (f) {
  var btn = $('#fPeng button[type=submit]'), er = $('#pengErr'); er.classList.add('d-none'); sibuk(btn, true);
  try { await call('savePengaduanConfig', { pengaduanSheetUrl: f.pengaduanSheetUrl.value, pengaduanTab: f.pengaduanTab.value, pengaduanStatusBaru: f.pengaduanStatusBaru.value, pengaduanStatusProses: f.pengaduanStatusProses.value, pengaduanStatusSelesai: f.pengaduanStatusSelesai.value }); toast('Pengaturan pengaduan disimpan.', 'ok'); adminPengaduan(); }
  catch (e) { er.textContent = e.message; er.classList.remove('d-none'); } finally { sibuk(btn, false); }
};
ACTIONS.ujiPeng = async function (el) {
  sibuk(el, true, 'Menguji…');
  try { var d = await call('ujiPengaduan'); toast('Tersambung. Pengaduan baru yang belum dibuatkan WO: ' + d.jumlahBaru + '.', 'ok', 6000); } catch (e) { toast(e.message, 'bad', 9000); } finally { sibuk(el, false); }
};
SUBMITS.simpanCfg = async function (f) {
  var btn = $('#fCfg button[type=submit]'); sibuk(btn, true);
  try { S.cfg = await call('saveConfig', { appName: f.appName.value, logoUrl: f.logoUrl.value, webAppUrl: f.webAppUrl.value }); terapkanBranding(); toast('Pengaturan disimpan. Jika URL Web App berubah, klik "Perbarui URL QR".', 'ok', 5000); }
  catch (e) { var er = $('#cfgErr'); er.textContent = e.message; er.classList.remove('d-none'); } finally { sibuk(btn, false); }
};
ACTIONS.regenQr = async function (el) {
  if (!(await konfirm('Perbarui URL QR?', 'URL QR semua aset akan ditulis ulang berdasarkan URL Web App saat ini. Label yang sudah tercetak dengan URL lama perlu dicetak ulang jika URL berubah.', 'Perbarui', false))) return;
  sibuk(el, true);
  try { var d = await call('regenerateQrUrls'); S.aset = await call('getAssetList'); toast(d.jumlah + ' aset diperbarui.', 'ok'); } catch (e) { gagal(e); } finally { sibuk(el, false); }
};

/* Log */
async function adminLog() {
  try {
    var rows = await call('getAuditLog');
    $('#adminBody').innerHTML = '<div class="card-x"><div class="card-h">Log aktivitas <span class="badge-x">300 terakhir</span></div><div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Waktu</th><th>Pengguna</th><th>Aksi</th><th>Objek</th><th>Detail</th></tr></thead><tbody>' +
      (rows.length ? rows.map(function (r) { return '<tr><td data-l="Waktu">' + esc(r.Waktu) + '</td><td data-l="User">' + esc(r.User) + '</td><td data-l="Aksi">' + badge(r.Aksi, 'info') + '</td><td data-l="Objek">' + esc(r.Objek + (r.ID_Objek ? ' · ' + r.ID_Objek : '')) + '</td><td data-l="Detail">' + esc(r.Detail) + '</td></tr>'; }).join('') : '<tr><td colspan="5"><div class="empty">Belum ada aktivitas.</div></td></tr>') + '</tbody></table></div></div>';
  } catch (e) { gagal(e); }
}

/* ════════ Kartu Pemeliharaan (mobile) ════════ */
var scanner = null, kCari = '';
function recent() { try { return JSON.parse(LS.get('sma_recent') || '[]'); } catch (e) { return []; } }
function catatRecent(id) { var r = recent().filter(function (x) { return x !== id; }); r.unshift(id); LS.set('sma_recent', JSON.stringify(r.slice(0, 8))); }
function renderKartu() {
  matikanScan(); var m = $('#kmain'), p = S.kPage;
  $$('.k-nav button').forEach(function (b) { b.classList.toggle('active', b.dataset.p === (p === 'wodetail' ? 'wo' : (p === 'pmdetail' ? 'pm' : p))); });
  if (p === 'beranda') kBeranda(m); else if (p === 'scan') kScan(m); else if (p === 'profil') kProfil(m);
  else if (p === 'aset') kAsetView(m);
  else if (p === 'wo') kWoList(m);
  else if (p === 'wodetail') kWoDetail(m);
  else if (p === 'pm') kPmList(m);
  else if (p === 'pmdetail') kPmDetail(m);
  else m.innerHTML = '<div class="card-x k-card text-center"><i class="bi ' + 'bi-calendar-check' + ' fs-1 text-primary"></i><h2 class="h5 fw-bold mt-2">' + 'Preventive Maintenance' + '</h2><p class="text-muted mb-0">Modul Preventive Maintenance menyusul pada tahap berikutnya. Untuk sekarang, scan QR aset untuk melihat informasinya.</p></div>' +
    '<button class="btn btn-primary k-btn mt-2" data-act="kNav" data-p="scan"><i class="bi bi-qr-code-scan me-2"></i>Scan QR Aset</button>';
  window.scrollTo(0, 0);
}
ACTIONS.kNav = function (el) { S.kPage = el.dataset.p; if (el.dataset.p !== 'aset') S.kAset = null; renderKartu(); };
function kItem(a) {
  return '<button class="k-aset" data-act="kBuka" data-id="' + esc(a.ID_Aset) + '">' + fotoHtml(a) + '<div class="min-w-0 flex-grow-1"><div class="fw-bold text-truncate">' + esc(a.Nama_Aset) + '</div><div class="small text-muted text-truncate">' + esc(a.ID_Aset) + ' · ' + esc(lokasiStr(a)) + '</div></div>' + bKondisi(a.Kondisi) + '</button>';
}
function kBeranda(m) {
  var rec = recent().map(cariAset).filter(Boolean);
  m.innerHTML = '<div class="k-hero"><div class="small">Selamat bekerja,</div><div class="h5 fw-bold mb-0">' + esc(S.user.nama) + '</div></div>' +
    '<div class="input-icon mb-3"><i class="bi bi-search"></i><input class="form-control" id="kCari" placeholder="Cari ID, nama, atau ruangan…" value="' + esc(kCari) + '" aria-label="Cari aset" style="min-height:48px"></div>' +
    '<div id="kDarurat">' + kDaruratHtml() + '</div><div id="kPmBesok">' + kPmBesokHtml() + '</div><div id="kHasil"></div>';
  kIsiHasil(); muatKartu(false);
}
/** Pengingat di Beranda Kartu: jadwal Preventive Maintenance yang jatuh BESOK (H-1). Hilang sendiri bila tidak ada jadwal besok. */
function kPmBesokHtml() {
  var besok = fmtWib(Date.now() + 86400000).slice(0, 10);
  var daftar = (PM.list || []).filter(function (p) { return (p.Status_PM === 'Terjadwal' || p.Status_PM === 'Ditunda') && String(p.Tanggal_Jadwal || p.Due_Date).slice(0, 10) === besok; });
  if (!daftar.length) return '';
  return '<div class="pm-besok" role="status"><div class="d-flex align-items-start gap-2"><i class="bi bi-calendar2-check pm-besok-ic"></i><div class="flex-grow-1 min-w-0">' +
    '<div class="fw-bold">Besok ada ' + daftar.length + ' jadwal Preventive Maintenance</div><div class="small text-muted mb-1">' + esc(fmtTgl(besok, true)) + '</div>' +
    daftar.slice(0, 3).map(function (p) { return '<div class="small text-truncate">• ' + esc(p.Nama_Aset) + (p.Ruangan ? ' — ' + esc(p.Ruangan) : '') + ' <span class="text-muted">(' + esc(p.ID_PM) + ')</span></div>'; }).join('') +
    (daftar.length > 3 ? '<div class="small text-muted">dan ' + (daftar.length - 3) + ' lainnya</div>' : '') +
    '<button class="btn btn-primary k-btn mt-2" data-act="kNav" data-p="pm">Lihat jadwal PM</button></div></div></div>';
}
/** Kartu darurat dari Risk Register (skor ≥ 10 oranye; skor ≥ 17 atau kategori Listrik merah): muncul di Beranda sampai WO-nya selesai/dibatalkan. Dirender sinkron dari S.darurat (bootstrap/cache), lalu disegarkan di latar belakang. */
function kDaruratHtml() {
  var tek = adalahTeknisi();
  return (S.darurat || []).map(function (x) {
    var ds = String(x.Deskripsi_Risiko || '');
    var og = x.Tingkat === 'oranye';
    return '<div class="darurat-card' + (og ? ' oranye' : '') + (x.Terlambat ? ' lewat' : '') + '" role="alert"><div class="d-flex align-items-start gap-2"><i class="bi bi-exclamation-octagon-fill darurat-ic"></i><div class="flex-grow-1 min-w-0">' +
      '<div class="darurat-t">' + (og ? 'RISIKO TINGGI' : 'DARURAT') + ' · Risk Register ' + esc(x.ID_Risiko) + '</div>' +
      '<div class="fw-bold">' + esc(x.Nama_Aset) + '</div><div class="small">' + esc([x.Gedung, x.Lantai, x.Ruangan].filter(Boolean).join(' · ')) + '</div>' +
      '<div class="small mt-1">' + esc(ds.length > 140 ? ds.slice(0, 140) + '…' : ds) + '</div>' +
      '<div class="d-flex flex-wrap gap-1 mt-2"><span class="badge-x bad">Skor ' + x.Skor_Inherent + ' · ' + esc(x.Level_Inherent) + '</span><span class="badge-x">' + esc(x.ID_WO) + ' · ' + esc(x.Status_WO) + '</span>' +
      '<span class="badge-x ' + (x.Terlambat ? 'bad' : 'warn') + '">' + (x.Terlambat ? 'Melewati batas SLA ' : 'Batas SLA ') + esc(fmtTgl(x.Batas_SLA)) + '</span></div>' +
      '<button class="btn ' + (og ? 'btn-warning' : 'btn-danger') + ' k-btn mt-2" data-act="kwBuka" data-id="' + esc(x.ID_WO) + '" data-dari="beranda">' + (tek ? 'Kerjakan sekarang' : 'Lihat Work Order') + '</button></div></div></div>';
  }).join('');
}
var KARTU_TTL = 3000;
/** Satu panggilan gabungan (WO + PM + darurat + opsi teknisi). Dipanggil saat Kartu dibuka, saat pindah menu (bila data > 3 dtk), dan setelah setiap aksi tulis. */
function muatKartu(paksa) {
  if (S.kartuBusy) { if (paksa) S.kartuUlang = true; return S.kartuBusy; }
  if (!paksa && S.kartuTs && Date.now() - S.kartuTs < KARTU_TTL) {
    if (!S.kartuTunda) S.kartuTunda = setTimeout(function () { S.kartuTunda = null; if (S.app === 'kartu') muatKartu(false); }, KARTU_TTL - (Date.now() - S.kartuTs) + 60);
    return Promise.resolve();
  }
  S.kartuBusy = call('getKartuData').then(function (d) {
    var sigBaru = ''; try { sigBaru = JSON.stringify([d.wo, d.pm, d.darurat]); } catch (e) { sigBaru = String(Date.now()); }
    var berubah = sigBaru !== S.kartuSig; S.kartuSig = sigBaru;   // tampilan hanya digambar ulang bila data benar-benar berubah (penyegaran senyap)
    WO.list = d.wo; WO.dimuat = true; PM.list = d.pm; PM.dimuat = true; S.darurat = d.darurat || []; S.kartuTs = Date.now();
    if (d.opsi && !WO.opsi) { WO.opsi = d.opsi; WO.verif = d.opsi.verifikator; }
    if (S.app !== 'kartu' || !berubah) return;
    if (S.kPage === 'beranda') { var el = $('#kDarurat'); if (el) el.innerHTML = kDaruratHtml(); var pb = $('#kPmBesok'); if (pb) pb.innerHTML = kPmBesokHtml(); }
    else if (S.kPage === 'wo' && $('#kwList')) { kWoPerbarui(); }
    else if (S.kPage === 'pm' && $('#kpList')) { kPmPerbarui(); }
  }).catch(function () { /* data tambahan di latar belakang: tampilan cache tetap dipakai */ }).then(function () { S.kartuBusy = null; if (S.kartuUlang) { S.kartuUlang = false; muatKartu(true); } });
  return S.kartuBusy;
}
/* Penyegaran senyap Kartu: tiap 8 detik cek penanda versi (ringan, tanpa membaca spreadsheet). Data penuh hanya dimuat bila ada perubahan
   atau sudah > 45 detik; tanpa spinner, tanpa menutup formulir. Berhenti saat tab tersembunyi atau saat membuka detail/formulir. */
var kVer = null, kCekBusy = false;
setInterval(async function () {
  if (S.app !== 'kartu' || !S.token || S.sesiHabis || document.hidden || kCekBusy || S.kartuBusy) return;
  if (['beranda', 'wo', 'pm'].indexOf(S.kPage) < 0) return;
  if (document.querySelector('.modal.show')) return;
  var ae = document.activeElement; if (ae && /^(INPUT|TEXTAREA|SELECT)$/.test(ae.tagName) && ae.id !== 'kwQ' && ae.id !== 'kCari') return;
  kCekBusy = true;
  try {
    var d = await rpc('getKartuVersi', S.token);
    var lama = kVer; kVer = d.v;
    if ((lama !== null && d.v !== lama) || Date.now() - (S.kartuTs || 0) > 45000) await muatKartu(true);
  } catch (e) { /* diam: jaringan putus sementara */ } finally { kCekBusy = false; }
}, 8000);
function kIsiHasil() {
  var q = kCari.trim(), el = $('#kHasil'); if (!el) return;
  if (!q) {
    var rec = recent().map(cariAset).filter(Boolean);
    el.innerHTML = rec.length ? '<div class="fw-bold small text-muted text-uppercase mb-2">Terakhir dilihat</div>' + rec.map(kItem).join('') : '<div class="empty"><i class="bi bi-qr-code"></i>Scan QR pada aset atau cari manual untuk melihat Kartu Pemeliharaan.</div>';
    return;
  }
  var hasil = cariUrut(S.aset.filter(function (a) { return a.Status_Aset !== DIHAPUS; }), q, TEKS_ASET);
  el.innerHTML = hasil.length ? '<div class="small text-muted mb-2">' + hasil.length + ' aset' + (hasil.length > 30 ? ' (30 pertama)' : '') + '</div>' + hasil.slice(0, 30).map(kItem).join('') : '<div class="empty"><i class="bi bi-search"></i>Aset tidak ditemukan.</div>';
}
document.addEventListener('input', debounce(function (e) { if (e.target.id === 'kCari') { kCari = e.target.value; kIsiHasil(); } }, 150));
ACTIONS.kBuka = function (el) { bukaAsetKartu(el.dataset.id); };
async function bukaAsetKartu(id) {
  id = idDariQr(id); if (!id) { toast('Kode QR tidak dikenali.', 'bad'); return; }
  S.kPage = 'aset'; S.kTab = 'info'; var cache = cariAset(id); S.kAset = cache ? { aset: cache, riwayatPM: [], sebagian: true } : { loading: true, id: id }; renderKartu(); muatKartu(false);
  try {
    var d = await call('getAssetDetail', id); S.kAset = d; catatRecent(d.aset.ID_Aset);
    if (!cariAset(d.aset.ID_Aset)) S.aset.push(d.aset);
  } catch (e) { S.kAset = { error: e.message, id: id }; }
  if (S.kPage === 'aset') kAsetView($('#kmain'));
}
function kAsetView(m) {
  var d = S.kAset;
  if (!d || d.loading) { m.innerHTML = '<div class="skeleton" style="height:220px"></div><div class="skeleton mt-3" style="height:120px"></div>'; return; }
  if (d.error) { m.innerHTML = '<div class="card-x k-card text-center"><i class="bi bi-exclamation-circle fs-1 text-danger"></i><p class="mt-2 mb-3">' + esc(d.error) + '</p><button class="btn btn-primary k-btn" data-act="kNav" data-p="scan">Scan ulang</button></div>'; return; }
  var a = d.aset, tab = S.kTab;
  var pm = d.riwayatPM || [];
  if (tab === 'pm' && !PM.dimuat) muatPm().then(function () { if (S.kPage === 'aset' && S.kTab === 'pm') kAsetView($('#kmain')); }).catch(function (e) { PM.dimuat = true; gagal(e); });
  if (tab === 'wo' && !WO.dimuat) muatWo().then(function () { if (S.kPage === 'aset' && S.kTab === 'wo') kAsetView($('#kmain')); }).catch(function (e) { WO.dimuat = true; gagal(e); });
  var isi = tab === 'info' ? '<div class="card-x k-card">' + dlHtml(a) + '</div>' :
    tab === 'wo' ? kAsetWoHtml(a) :
    kAsetPmHtml(a);
  var bawah = adalahTeknisi() ? '<div class="mt-3"><button class="btn btn-primary k-btn mb-2" data-act="kwLapor"><i class="bi bi-clipboard-plus me-2"></i>Lapor Work Order</button><button class="btn btn-light-x k-btn" data-act="kpLapor"><i class="bi bi-calendar-check me-2"></i>Lapor Preventive Maintenance</button></div>' : '';
  m.innerHTML = '<button class="btn-link-x mb-2" data-act="kNav" data-p="beranda"><i class="bi bi-arrow-left me-1"></i>Kembali</button>' +
    fotoBesar(a) + '<div class="mt-3 mb-2"><div class="h5 fw-bold mb-1">' + esc(a.Nama_Aset) + '</div><div class="text-muted small mb-2">' + esc(a.ID_Aset) + ' · ' + esc(lokasiStr(a)) + '</div></div>' +
    '<div class="k-tabs mt-3">' + [['info', 'Info Aset'], ['wo', 'Work Order'], ['pm', 'Preventive']].map(function (t) { return '<button data-act="kTab" data-t="' + t[0] + '" class="' + (tab === t[0] ? 'active' : '') + '">' + t[1] + '</button>'; }).join('') + '</div>' + isi + bawah;
}
ACTIONS.kTab = function (el) { S.kTab = el.dataset.t; kAsetView($('#kmain')); };

/* Scan QR */
function kScan(m) {
  m.innerHTML = '<div class="h5 fw-bold mb-3">Scan QR Aset</div><div class="scan-view mb-3" id="scanBox"><div id="qrReader"></div><div class="text-center text-white-50 p-5" id="scanIdle"><i class="bi bi-camera fs-1 d-block mb-2"></i>Arahkan kamera ke QR pada aset</div></div>' +
    '<button class="btn btn-primary k-btn mb-2" data-act="mulaiScan"><i class="bi bi-camera me-2"></i>Buka kamera</button>' +
    '<label class="btn btn-light-x k-btn d-flex align-items-center justify-content-center mb-3" style="cursor:pointer"><i class="bi bi-image me-2"></i>Ambil foto QR<input type="file" id="scanFile" accept="image/*" capture="environment" hidden></label>' +
    '<div id="scanErr" class="alert-x danger d-none mt-3"></div>';
}
function scanErr(msg) { var e = $('#scanErr'); if (e) { e.textContent = msg; e.classList.remove('d-none'); } }
matikanScan = function () { if (scanner) { var s = scanner; scanner = null; try { s.stop().then(function () { s.clear(); }).catch(function () {}); } catch (e) {} } };
ACTIONS.mulaiScan = async function (el) {
  sibuk(el, true, 'Membuka kamera…');
  try {
    await muatSkrip(CDN_QR, function () { return typeof Html5Qrcode !== 'undefined'; });
    var idle = $('#scanIdle'); if (idle) idle.remove();
    scanner = new Html5Qrcode('qrReader');
    await scanner.start({ facingMode: 'environment' }, { fps: 10, qrbox: { width: 240, height: 240 } }, function (txt) {
      var id = idDariQr(txt); if (!id) return;
      matikanScan(); if (navigator.vibrate) navigator.vibrate(60); bukaAsetKartu(id);
    }, function () {});
  } catch (e) {
    scanner = null;
    scanErr('Kamera tidak dapat dibuka (' + (e && e.message ? e.message : e) + '). Izinkan akses kamera, atau gunakan "Ambil foto QR".');
  } finally { sibuk(el, false); }
};
document.addEventListener('change', async function (e) {
  if (e.target.id !== 'scanFile' || !e.target.files[0]) return;
  try {
    await muatSkrip(CDN_QR, function () { return typeof Html5Qrcode !== 'undefined'; });
    var tmp = new Html5Qrcode('qrReader'), txt = await tmp.scanFile(e.target.files[0], false); try { tmp.clear(); } catch (x) {}
    var id = idDariQr(txt); if (!id) throw new Error('QR bukan milik aplikasi ini.'); bukaAsetKartu(id);
  } catch (er) { scanErr('QR tidak terbaca. Coba foto lebih dekat dan terang.'); }
  e.target.value = '';
});

function kProfil(m) {
  var u = S.user;
  m.innerHTML = '<div class="card-x k-card text-center"><span class="avatar mx-auto mb-2" style="width:64px;height:64px;font-size:1.5rem">' + esc(inisial(u.nama)) + '</span><div class="h5 fw-bold mb-0">' + esc(u.nama) + '</div><div class="text-muted small">@' + esc(u.username) + '</div><div class="mt-2">' + badge(u.peran, 'info') + '</div></div>' +
    '<button class="btn btn-light-x k-btn mb-2" data-act="ubahPassword"><i class="bi bi-key me-2"></i>Ganti password</button>' +
    '<button class="btn btn-light-x k-btn mb-2" data-act="toggleTheme"><i class="bi bi-moon-stars me-2"></i>Ganti tema</button>' +
    (bolehDash() ? '<button class="btn btn-light-x k-btn mb-2" data-act="toDash"><i class="bi bi-grid-1x2 me-2"></i>Buka Dashboard</button>' : '') +
    '<button class="btn btn-outline-danger k-btn" data-act="logout"><i class="bi bi-box-arrow-right me-2"></i>Keluar</button>';
  terapkanTema(S.theme);
}

/* ════════ Work Order — bersama (Dashboard & Kartu) ════════ */
var WO = { list: [], dimuat: false, peng: [], f: { q: '', st: '', kat: '', pel: '', dari: '', smp: '', urut: '', hal: 1, per: 25 }, view: 'tabel', verif: null };
var BLN = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
var WO_DIKERJAKAN = ['Open', 'Sedang Dikerjakan', 'Ditunda'];
function msWib(s) { var m = /^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2})(:\d{2})?/.exec(String(s || '')); return m ? Date.parse(m[1] + 'T' + m[2] + (m[3] || ':00') + '+07:00') : NaN; }
function fmtTgl(s, tanpaJam) {
  var m = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/.exec(String(s || '')); if (!m) return s ? String(s) : '—';
  var t = Number(m[3]) + ' ' + BLN[Number(m[2]) - 1] + ' ' + m[1]; return (tanpaJam || !m[4]) ? t : t + ' ' + m[4] + ':' + m[5];
}
function fmtDur(ms) { var a = Math.abs(ms), mnt = Math.floor(a / 60000), h = Math.floor(mnt / 60), d = Math.floor(h / 24); if (d >= 1) return d + 'h ' + (h % 24) + 'j'; if (h >= 1) return h + 'j ' + (mnt % 60) + 'm'; return mnt + 'm'; }
function fmtJam(j) { if (j === '' || j == null || isNaN(Number(j))) return '—'; j = Number(j); return j < 1 ? Math.round(j * 60) + ' menit' : (Math.round(j * 10) / 10).toString().replace('.', ',') + ' jam'; }
function hariIniLokal() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
function jamLokal() { var d = new Date(); return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2); }
function bStatusWo(v) { return badge(v, { 'Open': 'info', 'Sedang Dikerjakan': 'warn', 'Ditunda': '', 'Selesai': 'ok', 'Terlaksana': 'ok', 'Terjadwal': 'info', 'Dibatalkan': 'bad' }[v] || ''); }
function bPrio(v) { return badge(v, { 'Darurat': 'bad', 'Tinggi': 'warn', 'Sedang': 'info', 'Rendah': '' }[v] || ''); }
function progresWo(w) {
  if (w.Status_WO === 'Selesai') return { t: w.Status_SLA || 'Selesai', k: w.Status_SLA === 'Terlambat' ? 'bad' : 'ok' };
  if (w.Status_WO === 'Dibatalkan') return { t: 'Dibatalkan', k: '' };
  if (!w.Batas_SLA) return { t: '—', k: '' };
  var b = msWib(w.Batas_SLA); if (isNaN(b)) return { t: '—', k: '' };
  var sisa = b - Date.now(), total = Number(w.SLA_Jam) * 3600000, pre = w.Status_WO === 'Ditunda' ? 'Ditunda · ' : '';
  if (sisa < 0) return { t: pre + 'Terlambat ' + fmtDur(sisa), k: 'bad' };
  return { t: pre + 'Sisa ' + fmtDur(sisa), k: sisa < total * 0.25 ? 'warn' : 'ok' };
}
function bProgres(w) { var p = progresWo(w); return badge(p.t, p.k); }
function lokasiLengkap(w) { return lokPengaduan(w) || [w.Gedung, w.Lantai, w.Ruangan, w.Lokasi_Detail].filter(Boolean).join(' · '); }
async function muatWo() { var d = await call('getWoList'); WO.list = d.wo; WO.dimuat = true; return d; }
function gantiWoLokal(w) {
  var ada = false; WO.list = WO.list.map(function (x) { if (x.ID_WO === w.ID_WO) { ada = true; return w; } return x; });
  if (!ada) WO.list.push(w);
}

/* Bagian-bagian detail WO (dipakai modal dashboard dan halaman Kartu) */
function woSeksi(judul, isi) { return '<div class="wo-sec"><div class="wo-sec-h">' + esc(judul) + '</div><div class="wo-sec-b">' + isi + '</div></div>'; }
function dlRow(k, v) { return '<dt>' + k + '</dt><dd>' + (v === '' || v == null ? '<span class="text-muted">—</span>' : v) + '</dd>'; }
function teksPanjang(s) { return esc(s).replace(/\n/g, '<br>'); }
function stepsWo(w) {
  var s = w.Status_WO, mulai = !!w.Waktu_Mulai;
  var langkah2 = s === 'Ditunda' ? { n: 'Ditunda', c: 'warn', ic: 'bi-pause-fill' } : { n: 'Sedang Dikerjakan', c: mulai ? 'done' : '', ic: 'bi-check-lg' };
  var langkah3 = s === 'Dibatalkan' ? { n: 'Dibatalkan', c: 'bad', ic: 'bi-x-lg', t: w.Waktu_Selesai } : { n: 'Selesai', c: s === 'Selesai' ? 'done' : '', ic: 'bi-check-lg', t: s === 'Selesai' ? w.Waktu_Selesai : '' };
  var st = function (n, c, ic, t) { return '<div class="st ' + c + '"><div class="dot"><i class="bi ' + ic + '"></i></div><b>' + esc(n) + '</b>' + (t ? esc(fmtTgl(t)) : '') + '</div>'; };
  return '<div class="steps">' + st('Open', 'done', 'bi-check-lg', w.Waktu_Dibuat) + st(langkah2.n, langkah2.c, langkah2.ic, w.Waktu_Mulai) + st(langkah3.n, langkah3.c, langkah3.ic, langkah3.t) + '</div>';
}
function timelineWo(log) {
  if (!log || !log.length) return '<div class="text-muted small">Belum ada riwayat.</div>';
  return '<ul class="tl">' + log.map(function (x) {
    var ubahStatus = x.Status_Lama !== x.Status_Baru, kls = ubahStatus ? ({ 'Selesai': 'ok', 'Terlaksana': 'ok', 'Ditunda': 'warn', 'Dibatalkan': 'bad' }[x.Status_Baru] || '') : 'note';
    return '<li class="' + kls + '"><div class="t1">' + esc(ubahStatus ? x.Status_Baru : 'Catatan') + '</div><div class="t2">' + esc(fmtTgl(x.Waktu)) + ' · ' + esc(x.Oleh) + '</div>' + (x.Keterangan ? '<div class="t3">' + esc(x.Keterangan) + '</div>' : '') + '</li>';
  }).join('') + '</ul>';
}
function fotoThumb(u, label) { return u ? '<div class="foto-th"><img src="' + esc(imgU(u)) + '" alt="' + esc(label) + '" data-act="lb" data-src="' + esc(u) + '" referrerpolicy="no-referrer" data-fb="1"><span>' + esc(label) + '</span></div>' : ''; }
function sparepartWo(w) { try { var a = JSON.parse(w.Foto_Sparepart || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
function sparepartHtml(w) {
  var a = sparepartWo(w); if (!a.length) return '';
  return '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Foto sparepart yang diganti (' + a.length + ')</div><div class="foto-grid mb-2">' +
    a.map(function (x, i) { return '<div>' + fotoThumb(x.u, 'Sparepart ' + (i + 1)) + '<div class="small mt-1">' + esc(x.c) + '</div></div>'; }).join('') + '</div>';
}
function dokumenWo(w) { try { return JSON.parse(w.Dokumen_URLs || '[]'); } catch (e) { return []; } }
function laporanWoHtml(w, admin) {
  var p = progresWo(w), dok = dokumenWo(w);
  var ada = w.Foto_Before_1 || w.Foto_Sparepart || w.Kondisi_Dilaporkan || w.Foto_Aset_URL || w.Penyebab_Kerusakan || w.Foto_After_1 || w.Tindakan || w.Catatan || w.Verifikator || dok.length;
  if (!ada) return '';
  var foto = function (a, b, l) { var h = fotoThumb(a, l + ' 1') + fotoThumb(b, l + ' 2'); return h ? '<div class="foto-grid mb-2">' + h + '</div>' : ''; };
  return woSeksi('Laporan Teknisi',
    foto(w.Foto_Before_1, w.Foto_Before_2, 'Sebelum') + '<dl class="dl">' + dlRow('Penyebab Kerusakan', teksPanjang(w.Penyebab_Kerusakan)) + '</dl>' + sparepartHtml(w) +
    '<hr class="my-2">' + foto(w.Foto_After_1, w.Foto_After_2, 'Sesudah') +
    '<dl class="dl">' + dlRow('Tindakan', teksPanjang(w.Tindakan)) +
    dlRow('Dokumen', dok.length ? dok.map(function (x) { return '<a href="' + esc(x.u) + '" target="_blank" rel="noopener"><i class="bi bi-paperclip me-1"></i>' + esc(x.n) + '</a>'; }).join('<br>') : '') +
    dlRow('Kondisi Aset (menurut teknisi)', w.Kondisi_Dilaporkan ? bKondisi(w.Kondisi_Dilaporkan) : '') + (w.Foto_Aset_URL ? dlRow('Foto Aset (belum terdaftar)', fotoThumb(w.Foto_Aset_URL, 'Foto aset')) : '') + dlRow('Catatan', teksPanjang(w.Catatan)) + dlRow('Verifikator', esc(w.Verifikator)) + dlRow('Tanggal Verifikasi', w.Tanggal_Verifikasi ? esc(fmtTgl(w.Tanggal_Verifikasi, true)) : '') + dlRow('Status Verifikasi', w.Status_Verifikasi ? badge(w.Status_Verifikasi, 'ok') : '') +
    (admin ? dlRow('Kondisi Aset Sebelum', esc(w.Kondisi_Aset_Sebelum)) + dlRow('Kondisi Aset Sesudah', esc(w.Kondisi_Aset_Sesudah)) + dlRow('Status Aset Sesudah', esc(w.Status_Aset_Sesudah)) : '') + '</dl>');
}
/** Seluruh detail WO. opsi.kartu = tata letak ringkas untuk HP. */
function detailWoHtml(d) {
  var w = d.wo, a = d.aset, p = progresWo(w);
  var peng = w.ID_Pengaduan ? dlRow('ID Pengaduan', esc(w.ID_Pengaduan)) + dlRow('Pelapor', esc(w.Nama_Pelapor)) + dlRow('Lokasi (versi pelapor)', esc(w.Lokasi_Pengaduan)) + dlRow('Deskripsi Kerusakan', teksPanjang(w.Deskripsi_Kerusakan)) + dlRow('Jenis Kerusakan', esc(w.Jenis_Kerusakan)) : '';
  return stepsWo(w) +
    woSeksi('Identitas & Status', '<dl class="dl">' + dlRow('ID WO', '<b>' + esc(w.ID_WO) + '</b>') + dlRow('Status', bStatusWo(w.Status_WO) + ' ' + bProgres(w)) + dlRow('Sumber WO', esc(w.Sumber_WO)) +
      dlRow('ID Risiko', esc(w.ID_Risiko)) + dlRow('ID PM', esc(w.ID_PM)) + peng + dlRow('Dibuat', esc(fmtTgl(w.Waktu_Dibuat))) + '</dl>') +
    woSeksi('Objek Aset & Lokasi', '<dl class="dl">' + dlRow('ID Aset', esc(w.ID_Aset)) + dlRow('Nama Aset', esc(w.Nama_Aset)) + dlRow('Kategori', esc(w.Kategori_Aset)) +
      (a ? dlRow('Kritikalitas', bKrit(a.Kritikalitas)) + dlRow('Kondisi', bKondisi(a.Kondisi)) + dlRow('Status Aset', bStatus(a.Status_Aset)) : '') +
      dlRow('Gedung', esc(w.Gedung)) + dlRow('Lantai', esc(w.Lantai)) + dlRow('Ruangan', esc(w.Ruangan)) + dlRow('Lokasi Detail', esc(w.Lokasi_Detail)) +
      dlRow('PIC', esc(w.PIC_Aset || (a ? picLabel(a.PIC) : ''))) + '</dl>') +
    woSeksi('Deskripsi Masalah & Urgensi', '<dl class="dl">' + dlRow('Uraian Pekerjaan', teksPanjang(w.Uraian_Pekerjaan)) + dlRow('Temuan / Indikasi', teksPanjang(w.Temuan_Indikasi)) + dlRow('Prioritas', bPrio(w.Prioritas)) + '</dl>') +
    woSeksi('Penugasan & Target Waktu', '<dl class="dl">' + dlRow('Pelaksana', esc(w.Pelaksana)) + dlRow(w.Pelaksana === 'Vendor' ? 'Vendor' : 'Teknisi', esc(w.Nama_Teknisi_atau_Vendor)) + dlRow('Dimulai oleh', esc(w.Dimulai_Oleh)) +
      dlRow('Rencana Pengerjaan', esc(fmtTgl(w.Tanggal_Pengerjaan, true) + (w.Jam_Mulai_Rencana ? ' · ' + w.Jam_Mulai_Rencana : ''))) +
      dlRow('SLA', w.SLA_Jam === '' ? '' : esc(w.SLA_Jam + ' jam')) + dlRow('SLA dihitung dari', w.Dasar_SLA ? esc(fmtTgl(w.Dasar_SLA)) + ' <span class="text-muted">(laporan/WO dibuat)</span>' : '') + dlRow('Batas SLA', esc(fmtTgl(w.Batas_SLA))) + dlRow('Status SLA', w.Status_SLA ? badge(w.Status_SLA, w.Status_SLA === 'Terlambat' ? 'bad' : 'ok') : bProgres(w)) +
      dlRow('Waktu Respons', w.Waktu_Respons_Jam === '' ? '' : esc(fmtJam(w.Waktu_Respons_Jam))) + dlRow('Durasi Penyelesaian', w.Durasi_Penyelesaian_Jam === '' ? '' : esc(fmtJam(w.Durasi_Penyelesaian_Jam))) + '</dl>') +
    (w.Status_WO === 'Ditunda' ? woSeksi('Alasan Penundaan', '<dl class="dl">' + dlRow('Alasan', esc(w.Alasan_Tunda)) + dlRow('Keterangan', teksPanjang(w.Deskripsi_Tunda)) + '</dl>') : '') +
    laporanWoHtml(w, isAdmin() || bolehDash()) +
    woSeksi('Progres', timelineWo(d.log));
}

/* ════════ Pemilih aset (dipakai bersama; PRD §5.4) ════════ */
var AP = null;
function apBaru(w, pindai) {
  AP = { mode: 'pilih', aset: null, belum: false, man: { Nama_Aset: '', Kategori_Aset: '', pil: ['', null, null, null, ''], D: '' }, kategoriDulu: false };
  var a = null;
  if (w && w.ID_Aset && w.ID_Aset !== 'Belum Tahu') a = cariAset(w.ID_Aset);
  else if (pindai) a = cariAset(pindai);
  if (a) { AP.aset = a; AP.mode = 'terpilih'; }
  else if (w && (w.Nama_Aset || w.Gedung)) {
    AP.mode = 'lokasi'; AP.belum = true;
    AP.man = { Nama_Aset: w.Nama_Aset && w.Nama_Aset !== 'Belum Tahu' ? w.Nama_Aset : '', Kategori_Aset: w.Kategori_Aset || '', pil: ['', w.Gedung || null, w.Gedung ? (w.Lantai || '') : null, w.Ruangan || null, ''], D: w.Lokasi_Detail || '' };
  }
}
function apHtml() { return '<div class="ap" id="ap">' + apIsi() + '</div>'; }
function apRender() { var el = $('#ap'); if (el) el.innerHTML = apIsi(); }
function apIsi() {
  if (AP.mode === 'terpilih' && AP.aset) {
    var a = AP.aset;
    return '<div class="ap-sel">' + fotoHtml(a) + '<div class="min-w-0 flex-grow-1"><div class="fw-bold">' + esc(a.Nama_Aset) + '</div><div class="small text-muted">' + esc(a.ID_Aset) + (a.NUP ? ' · NUP ' + esc(a.NUP) : '') + ' · ' + esc(a.Kategori_Aset) + '</div><div class="small text-muted">' + esc(lokasiLengkap(a)) + '</div></div>' +
      '<button type="button" class="btn btn-light-x btn-sm" data-act="apGanti">Ganti</button></div>';
  }
  if (AP.mode === 'pilih') {
    return '<label class="form-label">Cari aset (nama, ID, NUP, atau lokasi)</label><div class="input-icon"><i class="bi bi-search"></i><input class="form-control" id="apQ" placeholder="Ketik nama, ID, atau ruangan — mis. AC Medvis" autocomplete="off"></div><div class="ap-res" id="apRes"></div>' +
      '<div class="d-flex flex-wrap gap-3 mt-2"><button type="button" class="btn-link-x" data-act="apLokasi"><i class="bi bi-geo-alt me-1"></i>Pilih lewat lokasi &amp; kategori</button>' + (AP.tdSaja ? '' : '<button type="button" class="btn-link-x" data-act="apBelum"><i class="bi bi-question-circle me-1"></i>ID Aset belum diketahui</button>') + '</div>';
  }
  // mode 'lokasi'
  if (!AP.belum) return apIsiTerdaftar();
  var pil = AP.man.pil, nm = ['', 'Gedung', 'Lantai', 'Ruangan'];
  var sel = function (i) { if (i === 0) return ''; return '<div class="col-md-6"><label class="form-label req">' + nm[i] + (i === 2 ? ' <span class="text-muted fw-normal">(boleh tanpa lantai)</span>' : '') + '</label><select class="form-select" data-ap-lok="' + i + '">' + lokOpsiHtml(i, pil, pil[i], i === 3 && !AP.tdSaja) + '</select></div>' + (i === 3 && pil[3] === RUANG_BARU ? '<div class="col-12"><label class="form-label req">Nama ruangan baru</label><input class="form-control" data-ap="rb" maxlength="80" value="' + esc(AP.man.RB || '') + '" placeholder="mis. Ruang Arsip Lt. 2"><div class="form-text">Ruangan ini dikirim sebagai usulan ke Master Admin untuk didaftarkan ke Master Lokasi.</div></div>' : ''); };
  var kat = '<div class="col-md-6"><label class="form-label req">Kategori Aset</label><select class="form-select" data-ap="kat">' + opsi(listMaster('Kategori Aset'), AP.man.Kategori_Aset, '— pilih —') + '</select></div>';
  var belum = '<div class="col-12"><label class="form-label req">Nama / deskripsi aset (ID Aset: "Belum Tahu")</label><input class="form-control" data-ap="nama" maxlength="150" value="' + esc(AP.man.Nama_Aset) + '" placeholder="mis. AC split ruang rapat lantai 3"></div>' +
    '<div class="col-12"><label class="form-label">Lokasi Detail <span class="text-muted fw-normal">(ketik bebas, boleh kosong)</span></label><input class="form-control" data-ap="lokd" maxlength="150" value="' + esc(AP.man.D) + '" placeholder="mis. dekat jendela barat, rak paling atas"></div>';
  var urut = AP.kategoriDulu ? kat + sel(0) + sel(1) + sel(2) + sel(3) : sel(0) + sel(1) + sel(2) + sel(3) + kat;
  return '<div class="row g-2">' + urut + belum + '</div>' +
    '<div class="d-flex flex-wrap gap-3 mt-2"><button type="button" class="btn-link-x" data-act="apCari"><i class="bi bi-search me-1"></i>Coba fitur pencarian</button><button type="button" class="btn-link-x" data-act="apLokasi"><i class="bi bi-geo-alt me-1"></i>Pilih dari daftar aset</button></div>';
}
/* Pilih lewat lokasi & kategori: pilihan hanya yang ADA asetnya (Kategori → Gedung → Lantai → Ruangan → Nama Aset). */
var TANPA_RUANG = '__tr__';
function apDaftar(kat, ged, lan, ru) {
  return S.aset.filter(function (a) {
    if (a.Status_Aset === DIHAPUS) return false;
    if (kat && a.Kategori_Aset !== kat) return false;
    if (ged && a.Gedung !== ged) return false;
    if (lan !== null && lan !== undefined && (a.Lantai || '') !== lan) return false;
    if (ru !== null && ru !== undefined && (ru === TANPA_RUANG ? a.Ruangan : a.Ruangan !== ru)) return false;
    return true;
  });
}
function apUnik(list, fn) {
  var s = {}, o = [];
  list.forEach(function (a) { var v = fn(a); if (!s.hasOwnProperty(v)) { s[v] = 1; o.push(v); } });
  return o.sort(function (x, y) { return String(x).localeCompare(String(y), 'id', { numeric: true, sensitivity: 'base' }); });
}
function apIsiTerdaftar() {
  var m = AP.man, pil = m.pil, kat = m.Kategori_Aset, ged = pil[1], lan = pil[2], ru = pil[3];
  var semuaKat = apUnik(S.aset.filter(function (a) { return a.Status_Aset !== DIHAPUS; }), function (a) { return a.Kategori_Aset; }).filter(Boolean);
  var urutMaster = listMaster('Kategori Aset'); semuaKat.sort(function (x, y) { var i = urutMaster.indexOf(x), j = urutMaster.indexOf(y); return (i < 0 ? 999 : i) - (j < 0 ? 999 : j) || String(x).localeCompare(String(y), 'id'); });
  var o = function (v, t, sel) { return '<option value="' + esc(v) + '"' + (sel ? ' selected' : '') + '>' + esc(t) + '</option>'; };
  var hKat = '<option value="">— pilih kategori —</option>' + semuaKat.map(function (k) { return o(k, k, k === kat); }).join('');
  var gedList = kat ? apUnik(apDaftar(kat), function (a) { return a.Gedung; }).filter(Boolean) : [];
  var hGed = '<option value="">' + (kat ? '— pilih gedung (' + gedList.length + ') —' : '— pilih kategori dulu —') + '</option>' + gedList.map(function (g) { return o(g, g, g === ged); }).join('');
  var lanList = kat && ged ? apUnik(apDaftar(kat, ged), function (a) { return a.Lantai || ''; }) : [];
  var hLan = '<option value="' + LANTAI_PILIH + '">' + (ged ? '— pilih lantai (' + lanList.length + ') —' : '— pilih gedung dulu —') + '</option>' + lanList.map(function (l) { return o(l, l === '' ? '(tanpa lantai)' : l, lan !== null && lan !== undefined && l === lan); }).join('');
  var ruList = kat && ged && lan !== null && lan !== undefined ? apUnik(apDaftar(kat, ged, lan), function (a) { return a.Ruangan ? a.Ruangan : TANPA_RUANG; }) : [];
  var hRu = '<option value="">' + (lan !== null && lan !== undefined ? '— pilih ruangan (' + ruList.length + ') —' : '— pilih lantai dulu —') + '</option>' + ruList.map(function (r) { return o(r, r === TANPA_RUANG ? '(tanpa ruangan)' : r, r === ru); }).join('');
  var lokasiSiap = kat && ged && lan !== null && lan !== undefined && ru;
  var cocok = lokasiSiap ? apDaftar(kat, ged, lan, ru) : [];
  var sl = function (lbl, att, html, off) { return '<div class="col-md-6"><label class="form-label req">' + lbl + '</label><select class="form-select" ' + att + (off ? ' disabled' : '') + '>' + html + '</select></div>'; };
  var pilihAset = '<div class="col-12"><label class="form-label req">Nama Aset</label><select class="form-select" data-ap="aset"' + (lokasiSiap ? '' : ' disabled') + '><option value="">— ' + (lokasiSiap ? (cocok.length ? 'pilih aset (' + cocok.length + ')' : 'tidak ada aset') : 'lengkapi pilihan di atas') + ' —</option>' +
    cocok.slice(0, 300).map(function (a) { return '<option value="' + esc(a.ID_Aset) + '">' + esc(a.Nama_Aset) + ' — ' + esc(a.ID_Aset) + '</option>'; }).join('') + '</select></div>';
  return '<div class="row g-2">' + sl('Kategori Aset', 'data-ap="kat"', hKat, false) + sl('Gedung', 'data-ap-lok="1"', hGed, !kat) + sl('Lantai <span class="text-muted fw-normal">(boleh tanpa lantai)</span>', 'data-ap-lok="2"', hLan, !ged) + sl('Ruangan', 'data-ap-lok="3"', hRu, !(lan !== null && lan !== undefined)) + pilihAset + '</div>' +
    '<div class="form-text mt-1">Pilihan hanya menampilkan gedung, lantai, dan ruangan yang memiliki aset tersebut. Aset tidak ada di daftar? Pilih <b>ID Aset belum diketahui</b>.</div>' +
    '<div class="d-flex flex-wrap gap-3 mt-2"><button type="button" class="btn-link-x" data-act="apCari"><i class="bi bi-search me-1"></i>Coba fitur pencarian</button>' + (AP.tdSaja ? '' : '<button type="button" class="btn-link-x" data-act="apBelum"><i class="bi bi-question-circle me-1"></i>ID Aset belum diketahui</button>') + '</div>';
}
function apNilai() {
  if (AP.aset) return { ID_Aset: AP.aset.ID_Aset };
  if (AP.mode === 'lokasi' && AP.belum) {
    var m = AP.man, p = m.pil;
    if (!m.Nama_Aset.trim()) throw new Error('Isi nama/deskripsi aset.');
    if (!m.Kategori_Aset) throw new Error('Pilih Kategori Aset.');
    if (!p[1] || p[2] === null || !p[3]) throw new Error('Lengkapi Gedung, Lantai, dan Ruangan.');
    var rb = p[3] === RUANG_BARU, nmRu = rb ? String(m.RB || '').trim() : p[3];
    if (rb && !nmRu) throw new Error('Isi nama ruangan baru.');
    var hs = { ID_Aset: '', Nama_Aset: m.Nama_Aset.trim(), Kategori_Aset: m.Kategori_Aset, Gedung: p[1], Lantai: p[2], Ruangan: nmRu, Lokasi_Detail: (m.D || '').trim() };
    if (rb) hs.Ruangan_Baru = true;
    return hs;
  }
  throw new Error('Pilih aset, atau pilih "ID Aset belum diketahui" lalu lengkapi lokasinya.');
}
ACTIONS.apGanti = function () { AP.aset = null; AP.mode = 'pilih'; AP.belum = false; apRender(); };
ACTIONS.apLokasi = function () { AP.mode = 'lokasi'; AP.belum = false; AP.man.pil = ['', null, null, null, '']; apRender(); };
ACTIONS.apBelum = function () { AP.mode = 'lokasi'; AP.belum = true; var p = AP.man.pil; if (p[3] === TANPA_RUANG) p[3] = null; apRender(); };
ACTIONS.apCari = function () { AP.mode = 'pilih'; AP.belum = false; apRender(); };
ACTIONS.apPilih = function (el) { var a = cariAset(el.dataset.id); if (a) { AP.aset = a; AP.mode = 'terpilih'; apRender(); } };
document.addEventListener('input', function (e) {
  var t = e.target; if (!AP || !t.closest || !t.closest('#ap')) return;
  if (t.id === 'apQ') {
    var q = t.value.trim(), el = $('#apRes'); if (!el) return;
    if (q.length < 2) { el.innerHTML = ''; return; }
    var hasil = cariUrut(S.aset.filter(function (a) { return a.Status_Aset !== DIHAPUS; }), q, TEKS_ASET).slice(0, 12);
    el.innerHTML = hasil.length ? hasil.map(function (a) { return '<button type="button" class="ap-item" data-act="apPilih" data-id="' + esc(a.ID_Aset) + '">' + fotoHtml(a) + '<div class="min-w-0"><div class="fw-bold text-truncate">' + esc(a.Nama_Aset) + '</div><div class="small text-muted text-truncate">' + esc(a.ID_Aset) + ' · ' + esc(lokasiStr(a)) + '</div></div></button>'; }).join('') : '<div class="small text-muted p-2">Tidak ditemukan. Coba kata lain, atau pilih lewat lokasi &amp; kategori.</div>';
  } else if (t.dataset.ap === 'nama') AP.man.Nama_Aset = t.value;
  else if (t.dataset.ap === 'lokd') AP.man.D = t.value;
  else if (t.dataset.ap === 'rb') AP.man.RB = t.value;
});
document.addEventListener('change', function (e) {
  var t = e.target; if (!AP || !t.closest || !t.closest('#ap')) return;
  if (t.dataset.apLok !== undefined) {
    var i = Number(t.dataset.apLok), v = t.value, pil = AP.man.pil;
    pil[i] = i === 2 ? (v === LANTAI_PILIH ? null : v) : i === 4 ? v : (v === '' ? null : v);
    for (var j = i + 1; j < 5; j++) pil[j] = j === 4 ? '' : null;
    if (i === 3 && v === RUANG_BARU) AP.belum = true;
    apRender();
  } else if (t.dataset.ap === 'kat') { AP.man.Kategori_Aset = t.value; if (!AP.belum) AP.man.pil = ['', null, null, null, '']; apRender(); }
  else if (t.dataset.ap === 'aset') {
    if (t.value === '__belum__') { AP.belum = true; apRender(); }
    else if (t.value) { var a = cariAset(t.value); if (a) { AP.aset = a; AP.mode = 'terpilih'; AP.belum = false; apRender(); } }
    else { AP.belum = false; apRender(); }
  }
});

/* ════════ Dashboard: halaman Work Order ════════ */
function woTersaring() {
  var f = WO.f, q = f.q.trim();
  var l0 = WO.list.filter(function (w) {
    if (f.st && w.Status_WO !== f.st) return false;
    if (f.kat && w.Kategori_Aset !== f.kat) return false;
    if (f.pel && w.Pelaksana !== f.pel) return false;
    var tg = String(w.Waktu_Dibuat).slice(0, 10);
    if (f.dari && tg < f.dari) return false;
    if (f.smp && tg > f.smp) return false;
    return true;
  });
  if (q) return cariUrut(l0, q, function (w) { return w.ID_WO + ' ' + w.Nama_Aset + ' ' + w.ID_Aset + ' ' + w.Kategori_Aset + ' ' + w.Uraian_Pekerjaan + ' ' + w.Lantai + ' ' + w.Ruangan + ' ' + w.Gedung + ' ' + w.Nama_Teknisi_atau_Vendor + ' ' + w.ID_Pengaduan + ' ' + w.Nama_Pelapor + ' ' + w.Lokasi_Pengaduan + ' ' + w.Jenis_Kerusakan; });
  return l0.sort(function (a, b) {
    if (f.urut === 'tgl' || f.urut === 'tgl-') { var c = kunciMulai(a) < kunciMulai(b) ? -1 : (kunciMulai(a) > kunciMulai(b) ? 1 : 0); return f.urut === 'tgl' ? c : -c; }
    return a.Waktu_Dibuat < b.Waktu_Dibuat ? 1 : (a.Waktu_Dibuat > b.Waktu_Dibuat ? -1 : 0);
  });
}
/** Kunci urut: tanggal + jam mulai pengerjaan (WO tanpa tanggal ditaruh paling akhir). */
function kunciMulai(w) { return (w.Tanggal_Pengerjaan ? w.Tanggal_Pengerjaan + ' ' + (w.Jam_Mulai_Rencana || '00:00') : '9999-12-31 00:00'); }
function waktuMulaiHtml(w) { return w.Tanggal_Pengerjaan ? '<div class="nm">' + esc(fmtTgl(w.Tanggal_Pengerjaan, true)) + '</div><div class="sub">' + (w.Jam_Mulai_Rencana ? esc(w.Jam_Mulai_Rencana) + ' WIB' : '—') + '</div>' : '<span class="text-muted">—</span>'; }
ACTIONS.woUrut = function () { WO.f.urut = WO.f.urut === '' ? 'tgl' : (WO.f.urut === 'tgl' ? 'tgl-' : ''); WO.f.hal = 1; renderWo(); };
async function renderWo() {
  if (WO.view === 'grafik') { renderWoGrafik(); return; }
  $('#main').innerHTML = '<div class="skeleton" style="height:90px"></div><div class="skeleton mt-3" style="height:360px"></div>';
  try { await muatWo(); } catch (e) {
    if (S.page === 'wo') $('#main').innerHTML = '<div class="card-x"><div class="empty"><i class="bi bi-exclamation-circle"></i>' + esc(e.message) + '<div class="mt-3"><button class="btn btn-primary" data-act="woMuat">Coba lagi</button></div></div></div>';
    return;
  }
  if (S.page !== 'wo' || WO.view === 'grafik') return;
  var f = WO.f, kelola = OPS.indexOf(S.user.peran) >= 0;
  var statusList = listMaster('Status WO'); if (!statusList.length) statusList = ['Open', 'Sedang Dikerjakan', 'Ditunda', 'Selesai', 'Dibatalkan'];
  $('#main').innerHTML =
    '<div class="page-head"><h2>Work Order</h2><div class="ms-auto d-flex flex-wrap gap-2">' +
    '<button class="btn btn-light-x" data-act="woEkspor"><i class="bi bi-file-earmark-excel me-1"></i>Ekspor Excel</button>' +
    '<button class="btn btn-light-x" data-act="woGrafik"><i class="bi bi-bar-chart-line me-1"></i>Lihat Grafik</button>' +
    (kelola ? '<button class="btn btn-primary" data-act="woBaru"><i class="bi bi-plus-lg me-1"></i>WO Langsung</button>' : '') +
    '</div></div><div class="kpi-grid" id="woStat"></div>' +
    '<div class="card-x"><div class="toolbar">' +
    '<div class="grow input-icon"><i class="bi bi-search"></i><input class="form-control" data-wf="q" placeholder="Cari ID WO, aset, uraian, teknisi…" value="' + esc(f.q) + '" aria-label="Cari Work Order"></div>' +
    '<select class="form-select fsel" data-wf="st" aria-label="Status">' + opsi(statusList, f.st, 'Semua status') + '</select>' +
    '<select class="form-select fsel" data-wf="kat" aria-label="Kategori">' + opsi(listMaster('Kategori Aset'), f.kat, 'Semua kategori') + '</select>' +
    '<select class="form-select fsel" data-wf="pel" aria-label="Pelaksana">' + opsi(listMaster('Pelaksana'), f.pel, 'Semua pelaksana') + '</select>' +
    '<div class="d-flex align-items-center gap-1"><input type="date" class="form-control" style="max-width:150px" data-wf="dari" value="' + esc(f.dari) + '" aria-label="Dari tanggal"><span class="text-muted">–</span><input type="date" class="form-control" style="max-width:150px" data-wf="smp" value="' + esc(f.smp) + '" aria-label="Sampai tanggal"></div>' +
    '<button class="btn btn-light-x btn-sm" data-act="woReset">Reset</button>' +
    '</div><div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Work Order</th><th><button type="button" class="th-sort" data-act="woUrut" title="Urutkan berdasarkan tanggal pengerjaan">Mulai Pengerjaan <i class="bi ' + (f.urut === 'tgl' ? 'bi-sort-numeric-down' : (f.urut === 'tgl-' ? 'bi-sort-numeric-up' : 'bi-arrow-down-up')) + '"></i></button></th><th>Aset &amp; Lokasi</th><th>Pelaksana</th><th>Status</th><th>Progres SLA</th><th></th></tr></thead><tbody id="woBody"></tbody></table></div>' +
    '<div class="pager" id="woPager"></div></div>';
  isiWo();
}
ACTIONS.woMuat = function () { renderWo(); };
ACTIONS.woReset = function () { WO.f = { q: '', st: '', kat: '', pel: '', dari: '', smp: '', urut: '', hal: 1, per: 25 }; renderWo(); };
function isiWo() {
  var f = WO.f, data = woTersaring(), total = data.length, hal = Math.max(1, Math.ceil(total / f.per));
  if (f.hal > hal) f.hal = hal;
  var mulai = (f.hal - 1) * f.per, tmpl = data.slice(mulai, mulai + f.per), kelola = OPS.indexOf(S.user.peran) >= 0;
  var n = function (s) { return data.filter(function (w) { return w.Status_WO === s; }).length; };
  var dur = data.filter(function (w) { return w.Status_WO === 'Selesai' && w.Durasi_Penyelesaian_Jam !== ''; }).map(function (w) { return Number(w.Durasi_Penyelesaian_Jam); });
  var rata = dur.length ? dur.reduce(function (a, b) { return a + b; }, 0) / dur.length : '';
  $('#woStat').innerHTML = kpiHtml('bi-clipboard-data', '', fmtAngka(total), 'Total WO') + kpiHtml('bi-check2-circle', 'ok', fmtAngka(n('Selesai')), 'Selesai') +
    kpiHtml('bi-tools', 'warn', fmtAngka(n('Sedang Dikerjakan')), 'Sedang Dikerjakan') + kpiHtml('bi-pause-circle', '', fmtAngka(n('Ditunda')), 'Ditunda') +
    kpiHtml('bi-folder2-open', '', fmtAngka(n('Open')), 'Open') + kpiHtml('bi-hourglass-split', '', esc(fmtJam(rata)), 'Rata-rata Waktu Penyelesaian');
  $('#woBody').innerHTML = tmpl.length ? tmpl.map(function (w) {
    var p = progresWo(w);
    return '<tr><td data-l="WO"><div class="nm">' + esc(w.ID_WO) + '</div><div class="sub">' + esc(w.Sumber_WO) + ' · dibuat ' + esc(fmtTgl(w.Waktu_Dibuat, true)) + '</div></td>' +
      '<td data-l="Mulai">' + waktuMulaiHtml(w) + '</td>' +
      '<td data-l="Aset"><div class="nm">' + esc(w.Nama_Aset) + '</div><div class="sub">' + esc([w.ID_Aset === 'Belum Tahu' ? '' : w.ID_Aset, lokasiStr(w) || '—'].filter(Boolean).join(' · ')) + '</div></td>' +
      '<td data-l="Pelaksana"><div>' + esc(w.Pelaksana) + '</div><div class="sub">' + esc(w.Nama_Teknisi_atau_Vendor) + '</div></td>' +
      '<td data-l="Status">' + bStatusWo(w.Status_WO) + '</td><td data-l="Progres">' + badge(p.t, p.k) + '</td>' +
      '<td class="acts"><button class="icon-btn" data-act="woDetail" data-id="' + esc(w.ID_WO) + '" title="Detail &amp; progres" aria-label="Detail ' + esc(w.ID_WO) + '"><i class="bi bi-eye"></i></button>' +
      (kelola && (w.Status_WO === 'Open' || w.Status_WO === 'Ditunda') ? '<button class="icon-btn" data-act="woUbah" data-id="' + esc(w.ID_WO) + '" title="Ubah" aria-label="Ubah ' + esc(w.ID_WO) + '"><i class="bi bi-pencil"></i></button>' : '') +
      (kelola && w.Status_WO === 'Open' && !w.Waktu_Mulai ? '<button class="icon-btn text-danger" data-act="woHapus" data-id="' + esc(w.ID_WO) + '" title="Hapus" aria-label="Hapus ' + esc(w.ID_WO) + '"><i class="bi bi-trash"></i></button>' : '') + '</td></tr>';
  }).join('') : '<tr><td colspan="7"><div class="empty"><i class="bi bi-clipboard-x"></i>' + (WO.list.length ? 'Tidak ada Work Order yang cocok dengan filter.' : 'Belum ada Work Order.') + '</div></td></tr>';
  $('#woPager').innerHTML = '<div>' + (total ? (mulai + 1) + '–' + (mulai + tmpl.length) + ' dari ' + fmtAngka(total) + ' WO' : '0 WO') + '</div>' +
    '<div class="d-flex align-items-center gap-2"><button class="btn btn-light-x btn-sm" data-act="woHal" data-d="-1"' + (f.hal <= 1 ? ' disabled' : '') + '><i class="bi bi-chevron-left"></i></button><span>Hal. ' + f.hal + ' / ' + hal + '</span><button class="btn btn-light-x btn-sm" data-act="woHal" data-d="1"' + (f.hal >= hal ? ' disabled' : '') + '><i class="bi bi-chevron-right"></i></button></div>';
}
ACTIONS.woHal = function (el) { WO.f.hal += Number(el.dataset.d); isiWo(); };
document.addEventListener('input', debounce(function (e) { var t = e.target; if (S.page === 'wo' && t.dataset && t.dataset.wf === 'q') { WO.f.q = t.value; WO.f.hal = 1; isiWo(); } }, 200));
document.addEventListener('change', function (e) { var t = e.target; if (S.page === 'wo' && t.dataset && t.dataset.wf && t.dataset.wf !== 'q') { WO.f[t.dataset.wf] = t.value; WO.f.hal = 1; isiWo(); } });

/* Grafik (halaman terpisah + tombol Kembali) */
ACTIONS.woGrafik = function () { WO.view = 'grafik'; hancurkanChart(); renderWoGrafik(); window.scrollTo(0, 0); };
ACTIONS.kembaliWo = function () { WO.view = 'tabel'; renderWo(); };
function renderWoGrafik() {
  hancurkanChart();
  var data = woTersaring();
  var kotak = function (id, judul, ket) { return '<div class="card-x"><div class="card-h">' + judul + (ket ? '<span class="badge-x ms-auto">' + ket + '</span>' : '') + '</div><div class="card-b"><div class="chart-box"><canvas id="' + id + '" role="img" aria-label="' + esc(judul) + '"></canvas></div></div></div>'; };
  $('#main').innerHTML = '<div class="page-head"><button class="btn btn-light-x" data-act="kembaliWo"><i class="bi bi-arrow-left me-1"></i>Kembali ke Work Order</button><h2>Grafik Work Order</h2></div>' +
    '<div class="alert-x info small"><i class="bi bi-funnel me-1"></i>Grafik mengikuti filter di halaman Work Order (' + fmtAngka(data.length) + ' WO).</div>' +
    (data.length ? '<div class="grid-2">' + kotak('gwStatus', 'Status Work Order') + kotak('gwKat', 'WO per Kategori Pekerjaan', 'Kategori Aset') + '</div>' : '<div class="card-x"><div class="empty"><i class="bi bi-bar-chart-line"></i>Belum ada Work Order untuk digrafikkan.</div></div>');
  if (!data.length || typeof Chart === 'undefined') return;
  var gelap = S.theme === 'dark', grid = gelap ? '#22304D' : '#E2E8F0';
  Chart.defaults.color = gelap ? '#98A5BF' : '#6B7280'; Chart.defaults.font.family = "'Plus Jakarta Sans',sans-serif";
  var warna = { 'Open': '#1E5EFF', 'Sedang Dikerjakan': '#F59E0B', 'Ditunda': '#94A3B8', 'Selesai': '#16A34A', 'Dibatalkan': '#DC2626' };
  var st = hitung(data, 'Status_WO');
  S.charts.push(new Chart($('#gwStatus'), { type: 'doughnut', data: { labels: st.map(function (x) { return x[0]; }), datasets: [{ data: st.map(function (x) { return x[1]; }), backgroundColor: st.map(function (x, i) { return warna[x[0]] || PALET[i % 8]; }), borderWidth: 0 }] },
    options: { maintainAspectRatio: false, cutout: '62%', plugins: { legend: { position: 'bottom' } } } }));
  var kat = hitung(data, 'Kategori_Aset'), labels = kat.map(function (x) { return x[0]; });
  var urut = ['Open', 'Sedang Dikerjakan', 'Ditunda', 'Selesai', 'Dibatalkan'];
  S.charts.push(new Chart($('#gwKat'), { type: 'bar', data: { labels: labels, datasets: urut.map(function (s) { return { label: s, backgroundColor: warna[s], borderRadius: 4, data: labels.map(function (k) { return data.filter(function (w) { return w.Kategori_Aset === k && w.Status_WO === s; }).length; }) }; }) },
    options: { indexAxis: 'y', maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, scales: { x: { stacked: true, ticks: { precision: 0 }, grid: { color: grid } }, y: { stacked: true, grid: { display: false } } } } }));
}

/* Ekspor Excel */
ACTIONS.woEkspor = async function (el) {
  var data = woTersaring(); if (!data.length) { toast('Tidak ada data untuk diekspor.', 'warn'); return; }
  var kolom = ['ID_WO', 'Sumber_WO', 'ID_Pengaduan', 'ID_Aset', 'Nama_Aset', 'Kategori_Aset', 'Gedung', 'Lantai', 'Ruangan', 'Lokasi_Detail', 'Tanggal_Pengerjaan', 'Jam_Mulai_Rencana', 'Pelaksana', 'Nama_Teknisi_atau_Vendor',
    'Uraian_Pekerjaan', 'Temuan_Indikasi', 'Prioritas', 'SLA_Jam', 'Batas_SLA', 'Status_WO', 'Waktu_Dibuat', 'Waktu_Mulai', 'Waktu_Selesai', 'Waktu_Respons_Jam', 'Durasi_Penyelesaian_Jam', 'Status_SLA'];
  var tabel = [kolom].concat(data.map(function (w) { return kolom.map(function (k) { var v = w[k]; return typeof v === 'number' ? v : amanSel(v); }); }));
  sibuk(el, true, 'Menyiapkan…');
  var nama = 'Work_Order_' + new Date().toISOString().slice(0, 10);
  try {
    await muatSkrip(CDN_XLSX, function () { return typeof XLSX !== 'undefined'; });
    var ws = XLSX.utils.aoa_to_sheet(tabel), wb = XLSX.utils.book_new();
    ws['!cols'] = kolom.map(function (k) { return { wch: Math.min(40, Math.max(k.length + 2, 14)) }; });
    XLSX.utils.book_append_sheet(wb, ws, 'Work Order');
    unduhBlob(nama + '.xlsx', new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    toast('Berkas Excel diunduh (' + data.length + ' WO).', 'ok');
  } catch (e) {
    unduhBlob(nama + '.csv', new Blob([csvKeString(tabel)], { type: 'text/csv;charset=utf-8' }));
    toast('Pustaka Excel tidak tersedia — diunduh sebagai CSV (dapat dibuka di Excel).', 'warn', 6000);
  } finally { sibuk(el, false); }
};

/* Detail + progres */
ACTIONS.woDetail = async function (el) {
  var id = el.dataset.id;
  bukaModal({ judul: 'Work Order ' + id, size: 'lg', isi: '<div class="skeleton" style="height:240px"></div>', footer: '<button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>' });
  try {
    var d = await call('getWoDetail', id), w = d.wo, kelola = OPS.indexOf(S.user.peran) >= 0;
    $('#mdlContent .modal-body').innerHTML = detailWoHtml(d);
    $('#mdlContent .modal-footer').innerHTML =
      (kelola && w.Status_WO === 'Open' && !w.Waktu_Mulai ? '<button class="btn btn-light-x text-danger me-auto" data-act="woHapus" data-id="' + esc(id) + '"><i class="bi bi-trash me-1"></i>Hapus</button>' : '') +
      (kelola && (w.Status_WO === 'Open' || w.Status_WO === 'Ditunda') ? '<button class="btn btn-light-x" data-act="woUbah" data-id="' + esc(id) + '"><i class="bi bi-pencil me-1"></i>Ubah</button>' : '') +
      (kelola ? '<button class="btn btn-light-x" data-act="woPdf" data-id="' + esc(id) + '"><i class="bi bi-file-earmark-pdf me-1"></i>Cetak PDF</button>' : '') +
      '<button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>';
  } catch (e) { tutupModal(); gagal(e); }
};
ACTIONS.woHapus = async function (el) {
  var id = el.dataset.id;
  if (!(await konfirm('Hapus Work Order?', id + ' belum dikerjakan dan akan dihapus permanen. Bila berasal dari pengaduan, status di sheet pengaduan dikosongkan kembali (belum dipilih).', 'Hapus'))) return;
  try {
    var d = await call('deleteWo', id); WO.list = WO.list.filter(function (w) { return w.ID_WO !== id; });
    tutupModal(); toast(d.peringatan || 'Work Order dihapus.', d.peringatan ? 'warn' : 'ok', d.peringatan ? 7000 : 3500);
    if (S.page === 'wo') isiWo();
  } catch (e) { gagal(e); }
};

/* Daftar pengaduan baru */
ACTIONS.woPengaduan = async function () {
  bukaModal({ judul: 'Pengaduan Baru', size: 'lg', isi: '<div class="skeleton" style="height:200px"></div>', footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Tutup</button>' });
  try {
    WO.peng = await call('getPengaduanBaru');
    $('#mdlContent .modal-body').innerHTML = WO.peng.length ?
      '<div class="small text-muted mb-2">' + WO.peng.length + ' pengaduan yang kolom Tindak Lanjut-nya belum dipilih dan belum dibuatkan WO. Pilih satu untuk membuat Work Order.</div>' + WO.peng.map(function (p) {
        return '<div class="peng-item"><div class="d-flex flex-wrap gap-2 align-items-center"><b>' + esc(p.id) + '</b><span class="text-muted small ms-auto" title="Acuan perhitungan SLA"><i class="bi bi-clock me-1"></i>' + esc(fmtTgl(p.tanggal)) + '</span></div>' +
          '<div class="small mt-1"><i class="bi bi-person me-1"></i>' + esc(p.nama) + ' &nbsp;<i class="bi bi-geo-alt me-1"></i>' + esc(p.lokasi) + (p.unit ? ' &nbsp;<span class="text-muted">· ' + esc(p.unit) + '</span>' : '') + '</div>' +
          '<div class="peng-f"><span class="peng-l">Deskripsi Kerusakan</span><div class="desk">' + teksPanjang(p.deskripsi) + '</div></div><div class="peng-f"><span class="peng-l">Jenis Kerusakan</span><div class="desk">' + (p.jenis ? teksPanjang(p.jenis) : '—') + '</div></div><button class="btn btn-primary btn-sm mt-1" data-act="woDariPengaduan" data-id="' + esc(p.id) + '"><i class="bi bi-clipboard-plus me-1"></i>Buat Work Order</button></div>';
      }).join('') : '<div class="empty"><i class="bi bi-inbox"></i>Tidak ada pengaduan baru (Tindak Lanjut belum dipilih).</div>';
  } catch (e) { var b = $('#mdlContent .modal-body'); if (b) b.innerHTML = '<div class="alert-x danger mb-0">' + esc(e.message) + '</div>'; }
};

/* Formulir WO */
var WOF = null;
function opsiPrioritas(sel) {
  return '<option value="">— pilih —</option>' + (S.master.data['Prioritas'] || []).map(function (x) { return '<option value="' + esc(x.nilai) + '"' + (x.nilai === sel ? ' selected' : '') + '>' + esc(x.nilai) + ' (SLA ' + esc(x.keterangan) + ' jam)</option>'; }).join('');
}
function infoSla() {
  var el = $('#slaInfo'), f = $('#fWo'); if (!el || !f) return;
  var it = (S.master.data['Prioritas'] || []).filter(function (x) { return x.nilai === f.Prioritas.value; })[0];
  if (!it) { el.textContent = 'Pilih prioritas untuk menentukan SLA.'; return; }
  var jam = Number(it.keterangan) || 0, mulai = WOF && WOF.mode === 'ubah' && WOF.dasar ? msWib(WOF.dasar) : NaN;
  el.textContent = 'SLA ' + jam + ' jam · dihitung sejak laporan/WO dibuat' + (isNaN(mulai) ? '' : ' · Batas SLA ' + fmtTgl(fmtWib(mulai + jam * 3600000)));
}
function fmtWib(ms) { var d = new Date(ms + 7 * 3600000), p2 = function (n) { return ('0' + n).slice(-2); }; return d.getUTCFullYear() + '-' + p2(d.getUTCMonth() + 1) + '-' + p2(d.getUTCDate()) + ' ' + p2(d.getUTCHours()) + ':' + p2(d.getUTCMinutes()); }
/** Pilihan vendor: terdaftar (Data Master) atau ketik sendiri; yang diketik dikirim ke Master Admin untuk didaftarkan. */
function vendorBoxHtml(w) {
  var daftar = listMaster('Nama Vendor'), nm = (w && w.Pelaksana === 'Vendor') ? String(w.Nama_Teknisi_atau_Vendor || '') : '';
  var terdaftar = nm && daftar.indexOf(nm) >= 0, lain = (nm && !terdaftar) || !daftar.length;
  return '<div class="col-12' + ((w && w.Pelaksana === 'Vendor') ? '' : ' d-none') + '" id="boxVendor"><label class="form-label req">Nama Vendor</label>' +
    (daftar.length ? '<select class="form-select mb-2" name="VendorPilih"><option value="">— pilih vendor terdaftar —</option>' + daftar.map(function (v) { return '<option value="' + esc(v) + '"' + (v === nm ? ' selected' : '') + '>' + esc(v) + '</option>'; }).join('') + '<option value="__lain__"' + (lain ? ' selected' : '') + '>➕ Vendor lain (ketik nama)…</option></select>' : '<input type="hidden" name="VendorPilih" value="__lain__">') +
    '<input class="form-control' + (lain ? '' : ' d-none') + '" name="Nama_Vendor" maxlength="100" placeholder="Ketik nama vendor" value="' + esc(terdaftar ? '' : nm) + '">' +
    '<div class="form-text">Nama vendor yang belum terdaftar akan dikirim ke Master Admin untuk didaftarkan, dan tetap dipakai pada WO ini.</div></div>';
}
function formWoHtml(w, peng) {
  var ubah = !!(WOF && WOF.mode === 'ubah'), sumberList = listMaster('Sumber WO').filter(function (s) { return s !== 'Pengaduan' && s !== 'Lapor Langsung Teknisi'; });
  var sumber = ubah ? w.Sumber_WO : (peng ? 'Pengaduan' : (w && w.Sumber_WO) || 'Langsung');
  var info = peng ? '<div class="alert-x info"><div class="fw-bold mb-1"><i class="bi bi-chat-left-text me-1"></i>Pengaduan ' + esc(peng.id) + '</div><div class="small">Pelapor: <b>' + esc(peng.nama) + '</b> · Lokasi: ' + esc(peng.lokasi) + ' · ' + esc(fmtTgl(peng.tanggal)) + '</div><div class="small mt-1"><b>Deskripsi Kerusakan</b><br>' + teksPanjang(peng.deskripsi) + '</div><div class="small mt-1"><b>Jenis Kerusakan</b><br>' + esc(peng.jenis || '—') + '</div></div>' :
    (ubah && w.ID_Pengaduan ? '<div class="alert-x info small"><b>' + esc(w.ID_Pengaduan) + '</b> — ' + esc(w.Nama_Pelapor) + '<div class="mt-1"><b>Deskripsi Kerusakan</b><br>' + teksPanjang(w.Deskripsi_Kerusakan) + '</div><div class="mt-1"><b>Jenis Kerusakan</b><br>' + esc(w.Jenis_Kerusakan || '—') + '</div></div>' : '');
  return '<form id="fWo" data-submit="simpanWo" novalidate>' + info + '<div class="row g-3">' +
    '<div class="col-md-6"><label class="form-label req">Sumber WO</label>' + (peng || ubah ? '<input class="form-control" value="' + esc(sumber) + '" readonly><input type="hidden" name="Sumber_WO" value="' + esc(sumber) + '">' :
      '<select class="form-select" name="Sumber_WO" data-wsr="1">' + opsi(sumberList, sumber, null) + '</select>') + '</div>' +
    '<div class="col-md-6' + (sumber === 'Risk Register' ? '' : ' d-none') + '" id="boxRisiko"><label class="form-label req">ID Risiko</label><input class="form-control" name="ID_Risiko" maxlength="30" value="' + esc(w && w.ID_Risiko) + '" placeholder="mis. RSK-0001"></div>' +
    '<div class="col-md-6' + (sumber === 'Preventive Maintenance' ? '' : ' d-none') + '" id="boxPm"><label class="form-label req">ID PM</label><input class="form-control" name="ID_PM" maxlength="30" value="' + esc(w && w.ID_PM) + '" placeholder="mis. PM-0001"></div>' +
    '<div class="col-12"><div class="fw-bold small text-muted text-uppercase">Aset</div>' + apHtml() + '</div>' +
    '<div class="col-md-4"><label class="form-label req">Tanggal Pengerjaan</label><input type="date" class="form-control" name="Tanggal_Pengerjaan" value="' + esc((w && w.Tanggal_Pengerjaan) || hariIniLokal()) + '" required></div>' +
    '<div class="col-md-4"><label class="form-label req">Jam Mulai</label><input type="time" class="form-control" name="Jam_Mulai_Rencana" value="' + esc((w && w.Jam_Mulai_Rencana) || jamLokal()) + '" required></div>' +
    '<div class="col-md-4"><label class="form-label req">Pelaksana</label><select class="form-select" name="Pelaksana" required>' + opsi(listMaster('Pelaksana'), w && w.Pelaksana, '— pilih —') + '</select></div>' +
    vendorBoxHtml(w) +
    '<div class="col-12"><label class="form-label req">Uraian Pekerjaan</label><textarea class="form-control" name="Uraian_Pekerjaan" rows="3" maxlength="1000" required>' + esc((w && w.Uraian_Pekerjaan) || (peng ? peng.deskripsi : '')) + '</textarea></div>' +
    '<div class="col-12"><label class="form-label">Temuan / Indikasi</label><textarea class="form-control" name="Temuan_Indikasi" rows="2" maxlength="500">' + esc(w && w.Temuan_Indikasi) + '</textarea></div>' +
    '<div class="col-md-6"><label class="form-label req">Prioritas</label><select class="form-select" name="Prioritas" required>' + opsiPrioritas(w && w.Prioritas) + '</select></div>' +
    '<div class="col-md-6 d-flex align-items-end"><div class="small text-muted" id="slaInfo"></div></div>' +
    '<div class="col-12"><div class="alert-x danger d-none mb-0" id="woErr"></div></div></div></form>';
}
function bukaFormWo(w, peng, ubah) {
  WOF = { mode: ubah ? 'ubah' : 'baru', idWo: w && w.ID_WO, peng: peng || null, dibuat: ubah ? w.Waktu_Dibuat : '', pengaduan: !!(ubah && w.Sumber_WO === 'Pengaduan'), dasar: ubah ? (w.Dasar_SLA || '') : '' };
  apBaru(w, '');
  bukaModal({ judul: ubah ? 'Ubah ' + w.ID_WO : (peng ? 'Work Order dari Pengaduan ' + peng.id : 'Work Order Langsung'), size: 'lg', isi: formWoHtml(w, peng),
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fWo">' + (ubah ? 'Simpan perubahan' : 'Buat Work Order') + '</button>' });
  infoSla();
}
ACTIONS.woBaru = function () { bukaFormWo(null, null, false); };
ACTIONS.woDariPengaduan = function (el) {
  var p = WO.peng.filter(function (x) { return x.id === el.dataset.id; })[0]; if (!p) return;
  bukaFormWo(null, p, false);
};
ACTIONS.woUbah = async function (el) {
  var id = el.dataset.id;
  try { var d = await call('getWoDetail', id); bukaFormWo(d.wo, null, true); } catch (e) { gagal(e); }
};
document.addEventListener('change', function (e) {
  var t = e.target; if (!t.closest || !t.closest('#fWo')) return;
  if (t.dataset.wsr) { $('#boxRisiko').classList.toggle('d-none', t.value !== 'Risk Register'); $('#boxPm').classList.toggle('d-none', t.value !== 'Preventive Maintenance'); }
  if (t.name === 'Prioritas' || t.name === 'Tanggal_Pengerjaan' || t.name === 'Jam_Mulai_Rencana') infoSla();
  if (t.name === 'Pelaksana') { $('#boxVendor').classList.toggle('d-none', t.value !== 'Vendor'); }
  if (t.name === 'VendorPilih') { var ni = $('#fWo [name=Nama_Vendor]'); ni.classList.toggle('d-none', t.value !== '__lain__'); if (t.value === '__lain__') ni.focus(); }
});
SUBMITS.simpanWo = async function (f) {
  var err = $('#woErr'); err.classList.add('d-none');
  var salah = function (m) { err.textContent = m; err.classList.remove('d-none'); };
  var p = {};
  try { p = apNilai(); } catch (e) { salah(e.message); return; }
  ['Sumber_WO', 'Tanggal_Pengerjaan', 'Jam_Mulai_Rencana', 'Pelaksana', 'Uraian_Pekerjaan', 'Temuan_Indikasi', 'Prioritas'].forEach(function (k) { p[k] = f[k].value; });
  if (p.Sumber_WO === 'Risk Register') p.ID_Risiko = f.ID_Risiko.value;
  if (p.Sumber_WO === 'Preventive Maintenance') p.ID_PM = f.ID_PM.value;
  if (p.Pelaksana === 'Vendor') {
    var vp = f.VendorPilih.value, vn = vp === '__lain__' ? f.Nama_Vendor.value.trim() : vp;
    if (!vn) { salah('Nama vendor wajib dipilih atau diketik.'); return; }
    p.Nama_Vendor = vn;
  }
  if (WOF.peng) { p.ID_Pengaduan = WOF.peng.id; }
  if (WOF.mode === 'ubah') p.ID_WO = WOF.idWo;
  if (!p.Tanggal_Pengerjaan || !p.Jam_Mulai_Rencana || !p.Pelaksana || !p.Uraian_Pekerjaan.trim() || !p.Prioritas) { salah('Lengkapi semua kolom bertanda *.'); return; }
  if (p.Sumber_WO === 'Risk Register' && !p.ID_Risiko.trim()) { salah('ID Risiko wajib diisi.'); return; }
  if (p.Sumber_WO === 'Preventive Maintenance' && !p.ID_PM.trim()) { salah('ID PM wajib diisi.'); return; }
  var btn = $('button[type=submit][form=fWo]'); sibuk(btn, true, 'Menyimpan…');
  try {
    var d = await call('saveWo', p, WOF.mode);
    gantiWoLokal(d.wo); tutupModal();
    toast(d.peringatan || (WOF.mode === 'ubah' ? 'Work Order diperbarui.' : 'Work Order ' + d.wo.ID_WO + ' dibuat.'), d.peringatan ? 'warn' : 'ok', d.peringatan ? 8000 : 3500);
    if (S.page === 'wo') isiWo();
  } catch (e) { salah(e.message); } finally { sibuk(btn, false); }
};

/* ════════ US-24: Menunggu Pembaruan Kondisi Aset (Master Admin, di Data Aset) ════════ */
async function muatMenunggu() {
  try { S.menunggu = await call('getWoMenunggu'); } catch (e) { return; }
  isiMenunggu();
}
function isiMenunggu() {
  var el = $('#panelMenunggu'); if (!el) return;
  var l = S.menunggu || [];
  if (!l.length) { el.innerHTML = ''; return; }
  el.innerHTML = '<div class="card-x panel-menunggu mb-3"><div class="card-h"><i class="bi bi-hourglass-split text-warning"></i>Menunggu Pembaruan Kondisi Aset <span class="badge-x warn">' + l.length + ' WO</span>' +
    '<span class="text-muted small fw-normal ms-auto">Teknisi sudah memilih status WO. Perbarui Kondisi/Status Aset berdasarkan catatan teknisi.</span></div>' +
    '<div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Work Order</th><th>Aset</th><th>Status WO</th><th>Teknisi</th><th class="th-cat">Catatan Teknisi</th><th>Kondisi Saat Ini</th><th>Dilaporkan Teknisi</th><th></th></tr></thead><tbody>' +
    l.map(function (x) {
      return '<tr><td data-l="WO"><div class="nm">' + esc(x.ID_WO) + '</div><div class="sub">' + esc(fmtTgl(x.Waktu)) + '</div></td>' +
        '<td data-l="Aset"><div class="nm">' + esc(x.Nama_Aset) + '</div><div class="sub">' + esc([x.ID_Aset === 'Belum Tahu' ? '' : x.ID_Aset, lokasiStr(x)].filter(Boolean).join(' · ')) + '</div></td>' +
        '<td data-l="Status">' + bStatusWo(x.Status_WO) + (x.Alasan_Tunda ? '<div class="sub">' + esc(x.Alasan_Tunda) + '</div>' : '') + '</td><td data-l="Teknisi">' + esc(x.Teknisi) + '</td>' +
        '<td class="cat" data-l="Catatan">' + (x.Catatan ? teksPanjang(x.Catatan) : '<span class="text-muted fw-normal">Tidak ada catatan</span>') + '</td>' +
        '<td data-l="Kondisi">' + bKondisi(x.Kondisi_Sekarang) + '<div class="sub mt-1">' + esc(x.Status_Aset_Sekarang) + '</div></td>' +
        '<td data-l="Dilaporkan">' + (x.Kondisi_Dilaporkan ? bKondisi(x.Kondisi_Dilaporkan) + (x.Kondisi_Dilaporkan.indexOf('Rusak') === 0 ? '<div class="sub text-danger mt-1"><i class="bi bi-bell-fill me-1"></i>Perlu tindakan</div>' : '') : '<span class="text-muted">—</span>') + '</td>' +
        '<td class="acts"><button class="btn btn-primary btn-sm" data-act="menungguBuka" data-id="' + esc(x.ID_WO) + '">Perbarui</button></td></tr>';
    }).join('') + '</tbody></table></div></div>';
}
ACTIONS.menungguBuka = function (el) {
  var x = (S.menunggu || []).filter(function (m) { return m.ID_WO === el.dataset.id; })[0]; if (!x) return;
  bukaModal({ judul: 'Perbarui kondisi aset', size: 'lg',
    isi: '<form id="fKon" data-submit="simpanKondisi" data-wo="' + esc(x.ID_WO) + '" novalidate>' +
      '<div class="alert-x info"><b>' + esc(x.ID_WO) + '</b> · ' + bStatusWo(x.Status_WO) + ' · ' + esc(x.Nama_Aset) + ' (' + esc(x.ID_Aset) + ')</div>' +
      '<dl class="dl mb-3">' + dlRow('Teknisi / Vendor', esc(x.Teknisi)) + dlRow('Alasan Penundaan', esc(x.Alasan_Tunda ? x.Alasan_Tunda + (x.Deskripsi_Tunda ? ' — ' + x.Deskripsi_Tunda : '') : '')) + dlRow('Penyebab', teksPanjang(x.Penyebab)) + dlRow('Tindakan', teksPanjang(x.Tindakan)) + '</dl>' +
      '<div class="alert-x warn"><div class="fw-bold small mb-1">Catatan Teknisi</div>' + (x.Catatan ? teksPanjang(x.Catatan) : 'Tidak ada catatan.') + '</div>' +
      '<div class="row g-3"><div class="col-md-6"><label class="form-label req">Kondisi Aset</label><select class="form-select" name="Kondisi">' + opsi(listMaster('Kondisi Aset'), x.Kondisi_Dilaporkan || x.Kondisi_Sekarang) + '</select>' + (x.Kondisi_Dilaporkan ? '<div class="form-text">Teknisi melaporkan: <b>' + esc(x.Kondisi_Dilaporkan) + '</b> (kondisi saat ini: ' + esc(x.Kondisi_Sekarang) + ').</div>' : '') + '</div>' +
      '<div class="col-md-6"><label class="form-label req">Status Aset</label><select class="form-select" name="Status_Aset">' + opsi(listMaster('Status Aset'), x.Status_Aset_Sekarang) + '</select></div>' +
      '<div class="col-12 small text-muted">Perubahan dicatat di riwayat WO dan Log Aktivitas (sebelum → sesudah). Biarkan nilai apa adanya bila kondisi aset tidak berubah, lalu simpan untuk menandai WO selesai ditinjau.</div>' +
      '<div class="col-12"><div class="alert-x danger d-none mb-0" id="konErr"></div></div></div></form>',
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fKon">Simpan</button>' });
};
SUBMITS.simpanKondisi = async function (f) {
  var id = f.dataset.wo, btn = $('button[type=submit][form=fKon]'); sibuk(btn, true, 'Menyimpan…');
  try {
    await call('simpanKondisiDariWo', id, { Kondisi: f.Kondisi.value, Status_Aset: f.Status_Aset.value });
    var x = (S.menunggu || []).filter(function (m) { return m.ID_WO === id; })[0];
    if (x) { var a = cariAset(x.ID_Aset); if (a) { a.Kondisi = f.Kondisi.value; a.Status_Aset = f.Status_Aset.value; } }
    S.menunggu = (S.menunggu || []).filter(function (m) { return m.ID_WO !== id; });
    tutupModal(); toast('Kondisi aset diperbarui.', 'ok'); isiMenunggu(); muatNotif(); if (S.page === 'aset') { isiTabelAset(); isiStatAset(); }
  } catch (e) { var er = $('#konErr'); er.textContent = e.message; er.classList.remove('d-none'); } finally { sibuk(btn, false); }
};

/* ════════ Kartu Pemeliharaan: Work Order (teknisi) ════════ */
var KW_TAB = { open: ['Open', 'Sedang Dikerjakan'], ditunda: ['Ditunda'], selesai: ['Selesai', 'Dibatalkan'] };
S.kw = { tab: 'open', q: '', kat: '', waktu: '', pel: '', hal: 1, mulai: {}, draftLapor: false, draftMulai: false, id: null, d: null, err: '', langkah: 'info', dari: 'list', pindai: '', form: {}, foto: {}, docs: [], sp: [] };
function adalahTeknisi() { return S.user && S.user.peran === 'Teknisi'; }

/** Segera = Tanggal Pengerjaan hari ini; Terjadwal = bukan hari ini. Hanya untuk WO yang belum/sedang ditunda (Open & Ditunda). */
function bWaktuWo(w) {
  if (['Open', 'Ditunda'].indexOf(w.Status_WO) < 0 || !w.Tanggal_Pengerjaan) return '';
  if (w.Sumber_WO === 'Pengaduan') return badge('Segera', 'bad'); // laporan dari pengaduan selalu Segera
  return w.Tanggal_Pengerjaan === hariIniLokal() ? badge('Segera', 'bad') : badge('Terjadwal', 'info');
}
function waktuWo(w) { return w.Sumber_WO === 'Pengaduan' ? 'Segera' : (!w.Tanggal_Pengerjaan) ? '' : (w.Tanggal_Pengerjaan === hariIniLokal() ? 'Segera' : 'Terjadwal'); }

/* Draft laporan/mulai: tersimpan otomatis di perangkat agar tidak hilang bila teknisi tidak sengaja pindah menu atau halaman dimuat ulang */
function dkey(id) { return 'sma_dr_' + (S.user ? S.user.username : '') + '_' + id; }
function muatDraft(id) { try { var o = JSON.parse(LS.get(dkey(id)) || 'null'); if (o && o.ts && Date.now() - o.ts < 3 * 86400000) return o; } catch (e) {} return null; }
function simpanDraft() {
  var k = S.kw; if (!k.id || (!k.draftLapor && !k.draftMulai)) return;
  var o = muatDraft(k.id) || {};
  o.ts = Date.now();
  if (k.draftLapor) o.lapor = { form: k.form, foto: k.foto, docs: k.docs, sp: k.sp };
  if (k.draftMulai) o.mulai = k.mulai;
  try { LS.set(dkey(k.id), JSON.stringify(o)); } catch (e) {}
}
function hapusDraft(id) { try { LS.del(dkey(id)); } catch (e) {} }

function kwCard(w, dari, pindai) {
  return '<button class="wo-card" data-act="kwBuka" data-id="' + esc(w.ID_WO) + '" data-dari="' + (dari || 'list') + '" data-pindai="' + esc(pindai || '') + '">' +
    '<div class="d-flex flex-wrap gap-1 align-items-center mb-1"><b>' + esc(w.ID_WO) + '</b>' + bStatusWo(w.Status_WO) + bWaktuWo(w) + '</div>' +
    '<div class="fw-bold">' + esc(w.Nama_Aset) + '</div><div class="small text-muted">' + esc(w.Kategori_Aset) + '</div>' +
    '<div class="small mt-1"><i class="bi bi-geo-alt me-1"></i>' + esc(lokasiStr(w) || '—') + '</div>' +
    '<div class="small"><i class="bi bi-person-gear me-1"></i>' + esc(w.Pelaksana) + (w.Nama_Teknisi_atau_Vendor ? ' · ' + esc(w.Nama_Teknisi_atau_Vendor) : '') + '</div>' +
    '<div class="small"><i class="bi bi-calendar-event me-1"></i>' + esc(fmtTgl(w.Tanggal_Pengerjaan, true)) + (w.Jam_Mulai_Rencana ? ' · ' + esc(w.Jam_Mulai_Rencana) : '') + '</div>' +
    '<div class="u">' + esc(w.Uraian_Pekerjaan) + '</div>' + (muatDraft(w.ID_WO) ? '<div class="small text-muted mt-2"><i class="bi bi-save me-1"></i>Draft laporan tersimpan di perangkat ini</div>' : '') + '</button>';
}
async function kWoList(m) {
  if (WO.dimuat) { kWoGambar(m); muatKartu(false); return; }   // tampil instan dari cache, segarkan di latar belakang
  m.innerHTML = '<div class="skeleton" style="height:48px"></div><div class="skeleton mt-3" style="height:180px"></div>';
  try { await muatKartu(true); if (!WO.dimuat) await muatWo(); } catch (e) {}
  if (!WO.dimuat) {
    if (S.kPage === 'wo') m.innerHTML = '<div class="card-x k-card text-center"><i class="bi bi-exclamation-circle fs-1 text-danger"></i><p class="mt-2">Gagal memuat Work Order.</p><button class="btn btn-primary k-btn" data-act="kNav" data-p="wo">Coba lagi</button></div>';
    return;
  }
  if (S.kPage !== 'wo') return;
  kWoGambar(m);
}
/** Segarkan hanya hitungan tab & daftar (tanpa menghapus isian pencarian/filter yang sedang diketik). */
function kWoPerbarui() {
  $$('.k-tabs button[data-t]').forEach(function (b) { var t = b.dataset.t; if (KW_TAB[t]) b.textContent = ({ open: 'Open', ditunda: 'Ditunda', selesai: 'Selesai' })[t] + ' (' + WO.list.filter(function (w) { return KW_TAB[t].indexOf(w.Status_WO) >= 0; }).length + ')'; });
  kwIsi();
}
function kWoGambar(m) {
  var k = S.kw, jml = function (t) { return WO.list.filter(function (w) { return KW_TAB[t].indexOf(w.Status_WO) >= 0; }).length; };
  m.innerHTML = '<div class="d-flex align-items-center mb-3"><div class="h5 fw-bold mb-0">Work Order</div>' + (adalahTeknisi() ? '<button class="btn btn-primary btn-sm ms-auto" data-act="kwLaporBaru"><i class="bi bi-plus-lg me-1"></i>Buat Work Order</button>' : '') + '</div>' +
    '<div class="k-tabs">' + [['open', 'Open'], ['ditunda', 'Ditunda'], ['selesai', 'Selesai']].map(function (t) { return '<button data-act="kwTab" data-t="' + t[0] + '" class="' + (k.tab === t[0] ? 'active' : '') + '">' + t[1] + ' (' + jml(t[0]) + ')</button>'; }).join('') + '</div>' +
    '<div class="input-icon mb-2"><i class="bi bi-search"></i><input class="form-control" id="kwQ" placeholder="Cari ID WO, nama aset, uraian…" value="' + esc(k.q) + '" style="min-height:48px" aria-label="Cari Work Order"></div>' +
    '<div class="row g-2 mb-3" id="kwFilter">' + kwFilterHtml() + '</div>' +
    '<div id="kwList"></div><div id="kwPager"></div>';
  kwIsi();
}
/** Tab Open/Ditunda: filter Waktu Pengerjaan (Segera/Terjadwal). Tab Selesai: filter Kategori Aset. */
function kwFilterHtml() {
  var k = S.kw;
  return '<div class="col-6">' + (k.tab === 'selesai' ? '<select class="form-select" data-kwf="kat" aria-label="Kategori Aset">' + opsi(listMaster('Kategori Aset'), k.kat, 'Semua kategori') + '</select>' :
    '<select class="form-select" data-kwf="waktu" aria-label="Waktu pengerjaan">' + opsi(['Segera', 'Terjadwal'], k.waktu, 'Semua waktu pengerjaan') + '</select>') + '</div>' +
    '<div class="col-6"><select class="form-select" data-kwf="pel" aria-label="Pelaksana">' + opsi(listMaster('Pelaksana'), k.pel, 'Semua pelaksana') + '</select></div>';
}
function kwIsi() {
  var k = S.kw, q = k.q.trim(), el = $('#kwList'); if (!el) return;
  var data = WO.list.filter(function (w) {
    if (KW_TAB[k.tab].indexOf(w.Status_WO) < 0) return false;
    if (k.tab === 'selesai' && k.kat && w.Kategori_Aset !== k.kat) return false;
    if (k.tab !== 'selesai' && k.waktu && waktuWo(w) !== k.waktu) return false;
    if (k.pel && w.Pelaksana !== k.pel) return false;
    return true;
  });
  var teksWo = function (w) { return w.ID_WO + ' ' + w.Nama_Aset + ' ' + w.ID_Aset + ' ' + w.Kategori_Aset + ' ' + w.Uraian_Pekerjaan + ' ' + w.Lantai + ' ' + w.Ruangan + ' ' + w.Gedung + ' ' + w.Nama_Pelapor + ' ' + w.Lokasi_Pengaduan + ' ' + w.Jenis_Kerusakan; };
  if (q) data = cariUrut(data, q, teksWo);
  else if (k.tab === 'selesai') data.sort(function (a, b) { return a.Waktu_Selesai < b.Waktu_Selesai ? 1 : -1; });
  else if (!q) data.sort(function (a, b) { var x = String(a.Waktu_Dibuat), y = String(b.Waktu_Dibuat); return x < y ? 1 : (x > y ? -1 : (a.ID_WO < b.ID_WO ? 1 : -1)); }); // laporan terbaru di atas
  var per = 5, hal = Math.max(1, Math.ceil(data.length / per)); if (k.hal > hal) k.hal = hal;
  var mulai = (k.hal - 1) * per;
  el.innerHTML = data.length ? data.slice(mulai, mulai + per).map(function (w) { return kwCard(w); }).join('') : '<div class="empty"><i class="bi bi-clipboard-check"></i>Tidak ada Work Order di tab ini.</div>';
  $('#kwPager').innerHTML = data.length > per ? '<div class="d-flex align-items-center justify-content-between mt-2"><button class="btn btn-light-x" data-act="kwHal" data-d="-1"' + (k.hal <= 1 ? ' disabled' : '') + '><i class="bi bi-chevron-left"></i></button><span class="small text-muted">Hal. ' + k.hal + ' / ' + hal + ' · ' + data.length + ' WO</span><button class="btn btn-light-x" data-act="kwHal" data-d="1"' + (k.hal >= hal ? ' disabled' : '') + '><i class="bi bi-chevron-right"></i></button></div>' : '';
}
ACTIONS.kwTab = function (el) { S.kw.tab = el.dataset.t; S.kw.hal = 1; $$('.k-tabs button').forEach(function (b) { b.classList.toggle('active', b.dataset.t === S.kw.tab); }); var fe = $('#kwFilter'); if (fe) fe.innerHTML = kwFilterHtml(); kwIsi(); };
ACTIONS.kwHal = function (el) { S.kw.hal += Number(el.dataset.d); kwIsi(); window.scrollTo(0, 0); };
document.addEventListener('input', debounce(function (e) { if (e.target.id === 'kwQ') { S.kw.q = e.target.value; S.kw.hal = 1; kwIsi(); } }, 150));
document.addEventListener('change', function (e) { var t = e.target; if (t.dataset && t.dataset.kwf && S.kPage === 'wo') { S.kw[t.dataset.kwf] = t.value; S.kw.hal = 1; kwIsi(); } });

/* Detail & alur */
ACTIONS.kwBuka = function (el) { S.kw.id = el.dataset.id; S.kw.dari = el.dataset.dari || 'list'; S.kw.pindai = el.dataset.pindai || ''; kwMuatDetail('info'); };
async function kwMuatDetail(langkah) {
  S.kPage = 'wodetail'; S.kw.d = null; S.kw.err = ''; renderKartu();
  try {
    var pr = [call('getWoDetail', S.kw.id)];
    if (adalahTeknisi() && !WO.opsi) pr.push(call('getWoOpsi'));
    var hs = await Promise.all(pr); S.kw.d = hs[0];
    if (hs[1]) { WO.opsi = hs[1]; WO.verif = WO.opsi.verifikator; }
  } catch (e) { S.kw.err = e.message; }
  S.kw.langkah = langkah || 'info'; S.kw.draftLapor = false; S.kw.draftMulai = false;
  if (S.kw.d && langkah === 'lapor') kwInitLapor(S.kw.d.wo);
  if (S.kw.d && langkah === 'mulai') kwInitMulai(S.kw.d.wo);
  if (S.kPage === 'wodetail') kWoDetail($('#kmain'));
}
ACTIONS.kwKembali = function () { if (S.kw.dari === 'beranda') { S.kPage = 'beranda'; renderKartu(); } else if (S.kw.dari === 'aset' && S.kAset && S.kAset.aset) { S.kPage = 'aset'; renderKartu(); } else { S.kPage = 'wo'; renderKartu(); } };
ACTIONS.kwInfo = function () { S.kw.langkah = 'info'; kWoDetail($('#kmain')); };
/** "Mulai Pengerjaan": status WO langsung menjadi Sedang Dikerjakan (SLA tetap dihitung sejak laporan dibuat); lalu formulir data pengerjaan dibuka (tombol "Berikutnya"). */
ACTIONS.kwLangkahMulai = async function (el) {
  var w = S.kw.d.wo;
  if (w.Status_WO === 'Open' || w.Status_WO === 'Ditunda') {
    sibuk(el, true, 'Memulai…');
    try {
      var r = await call('mulaiSlaWo', w.ID_WO); gantiWoLokal(r.wo); Object.assign(S.kw.d.wo, r.wo);
      if (r.peringatan) toast(r.peringatan, 'warn', 8000);
    } catch (e) { sibuk(el, false); toast('Gagal memulai: ' + e.message, 'warn', 6000); return; }
    sibuk(el, false);
  }
  S.kw.langkah = 'mulai'; kwInitMulai(S.kw.d.wo); kWoDetail($('#kmain')); window.scrollTo(0, 0);
};
ACTIONS.kwLangkahLapor = function () { if (!kwBolehLapor(S.kw.d.wo)) { toast('Laporan hanya dapat diisi oleh teknisi yang menekan Mulai.', 'warn'); return; } S.kw.langkah = 'lapor'; kwInitLapor(S.kw.d.wo); kWoDetail($('#kmain')); window.scrollTo(0, 0); };
/** Laporan hanya boleh diisi akun teknisi yang menekan "Mulai" (WO lama tanpa catatan Dimulai_Oleh tetap boleh). */
function kwBolehLapor(w) { return adalahTeknisi() && (!w.Dimulai_Oleh || w.Dimulai_Oleh === S.user.username); }
function kwRingkas(w) {
  var p = progresWo(w);
  return '<div class="card-x k-card"><div class="d-flex flex-wrap gap-1 align-items-center mb-1"><b>' + esc(w.ID_WO) + '</b>' + bStatusWo(w.Status_WO) + bPrio(w.Prioritas) + '</div><div class="fw-bold">' + esc(w.Nama_Aset) + '</div>' +
    '<div class="small text-muted">' + esc(lokasiLengkap(w)) + '</div><div class="small mt-1">' + teksPanjang(w.Uraian_Pekerjaan) + '</div><div class="mt-2">' + badge(p.t, p.k) + ' <button class="btn-link-x" data-act="kwInfo">Lihat detail</button></div></div>';
}
function kWoDetail(m) {
  var k = S.kw, d = k.d;
  if (k.err) { m.innerHTML = '<button class="btn-link-x mb-2" data-act="kwKembali"><i class="bi bi-arrow-left me-1"></i>Kembali</button><div class="card-x k-card text-center"><i class="bi bi-exclamation-circle fs-1 text-danger"></i><p class="mt-2 mb-0">' + esc(k.err) + '</p></div>'; return; }
  if (!d) { var li = WO.list.filter(function (x) { return x.ID_WO === k.id; })[0]; m.innerHTML = '<button class="btn-link-x mb-2" data-act="kwKembali"><i class="bi bi-arrow-left me-1"></i>Kembali</button>' + (li ? kwRingkas(li).replace(/<button class="btn-link-x" data-act="kwInfo">Lihat detail<\/button>/, '') + '<div class="skeleton mt-3" style="height:140px"></div>' : '<div class="skeleton" style="height:60px"></div><div class="skeleton mt-3" style="height:260px"></div>'); return; }
  var w = d.wo, tek = adalahTeknisi(), balik = '<button class="btn-link-x mb-2" data-act="kwKembali"><i class="bi bi-arrow-left me-1"></i>Kembali</button>';
  if (k.langkah === 'mulai' && tek && (w.Status_WO === 'Open' || w.Status_WO === 'Ditunda' || (w.Status_WO === 'Sedang Dikerjakan' && kwBolehLapor(w)))) { m.innerHTML = balik + kwRingkas(w) + kwMulaiHtml(w); return; }
  if (k.langkah === 'lapor' && kwBolehLapor(w) && w.Status_WO === 'Sedang Dikerjakan') { m.innerHTML = balik + kwRingkas(w) + '<div id="kwLaporBox">' + kwLaporIsi(w) + '</div>'; kwCekTombol(); return; }
  var aksi = '';
  if (tek && (w.Status_WO === 'Open' || w.Status_WO === 'Ditunda')) aksi = '<button class="btn btn-primary k-btn mt-2" data-act="kwLangkahMulai"><i class="bi bi-play-circle me-2"></i>Mulai Pengerjaan</button>';
  else if (tek && w.Status_WO === 'Sedang Dikerjakan' && kwBolehLapor(w) && !w.Nama_Teknisi_atau_Vendor) aksi = '<button class="btn btn-primary k-btn mt-2" data-act="kwLangkahMulai"><i class="bi bi-pencil-square me-2"></i>Lengkapi Data Pengerjaan</button>';
  else if (tek && w.Status_WO === 'Sedang Dikerjakan' && kwBolehLapor(w)) aksi = (muatDraft(w.ID_WO) ? '<div class="alert-x warn small mt-2"><i class="bi bi-save me-1"></i>Ada draft laporan yang tersimpan di perangkat ini.</div>' : '') + '<button class="btn btn-primary k-btn mt-2" data-act="kwLangkahLapor"><i class="bi bi-clipboard2-check me-2"></i>' + (muatDraft(w.ID_WO) ? 'Lanjutkan Laporan' : 'Lapor Hasil Pekerjaan') + '</button>';
  else if (tek && w.Status_WO === 'Sedang Dikerjakan') aksi = '<div class="alert-x info small mt-2"><i class="bi bi-lock me-1"></i>WO ini dimulai oleh <b>' + esc(w.Nama_Teknisi_atau_Vendor || w.Dimulai_Oleh) + '</b> (akun <b>' + esc(w.Dimulai_Oleh) + '</b>). Laporan hasil hanya dapat diisi oleh akun yang menekan Mulai.</div>';
  else if (WO_DIKERJAKAN.indexOf(w.Status_WO) >= 0) aksi = '<div class="alert-x info small mt-2"><i class="bi bi-info-circle me-1"></i>Hanya Teknisi yang dapat memulai dan melaporkan Work Order.</div>';
  m.innerHTML = balik + '<div class="h5 fw-bold mb-3">' + esc(w.Nama_Aset) + '</div>' + detailWoHtml(d) + aksi;
}

/* Langkah 1: lengkapi data & tekan Mulai */
function kwInitMulai(w) {
  apBaru(w, S.kw.pindai);
  var dr = muatDraft(w.ID_WO), m = dr && dr.mulai;
  S.kw.mulai = { tgl: w.Tanggal_Pengerjaan || hariIniLokal(), jam: w.Jam_Mulai_Rencana || jamLokal(), pel: w.Pelaksana || 'Internal',
    tek: (w.Pelaksana !== 'Vendor' && w.Nama_Teknisi_atau_Vendor) ? String(w.Nama_Teknisi_atau_Vendor).split(/\s*,\s*/) : [S.user.nama],
    vendor: w.Pelaksana === 'Vendor' ? (w.Nama_Teknisi_atau_Vendor || '') : '' };
  if (m) { Object.keys(m).forEach(function (k) { if (k !== 'ap') S.kw.mulai[k] = m[k]; }); if (m.ap && !(w.ID_Aset && w.ID_Aset !== 'Belum Tahu')) AP = m.ap; }
  S.kw.draftMulai = true;
}
function kwSimpanMulai() {
  var f = $('#fMulai'); if (!f || !S.kw.draftMulai) return;
  var m = S.kw.mulai; m.tgl = f.tgl.value; m.jam = f.jam.value; m.pel = f.pel.value;
  if (m.pel === 'Vendor') { if (f.nama) m.vendor = f.nama.value; } else m.tek = $$('#fMulai [name=tek]:checked').map(function (c) { return c.value; });
  m.ap = AP; simpanDraft();
}
function kwNamaHtml(pel, m) {
  if (pel === 'Vendor') {
    var dv = uniq(listMaster('Nama Vendor').concat((WO.opsi && WO.opsi.vendor) || []));
    return '<label class="form-label req">Nama Vendor</label><input class="form-control" name="nama" list="dlVendor" maxlength="100" value="' + esc(m.vendor) + '" placeholder="Pilih atau ketik nama vendor"><datalist id="dlVendor">' + dv.map(function (x) { return '<option value="' + esc(x) + '">'; }).join('') + '</datalist>' +
      '<div class="form-text">' + (m.vendor ? 'Terisi otomatis dari WO yang dibuat di dashboard. ' : '') + 'Vendor yang belum terdaftar dikirim ke Master Admin untuk didaftarkan.</div>';
  }
  var daftar = uniq(((WO.opsi && WO.opsi.teknisi) || []).concat(listMaster('Nama Teknisi')));
  if (!daftar.length) return '<label class="form-label req">Nama Teknisi</label><input class="form-control" name="tekBebas" maxlength="100" value="' + esc((m.tek || [])[0] || S.user.nama) + '">';
  return '<label class="form-label req">Nama Teknisi <span class="text-muted fw-normal">(boleh lebih dari satu)</span></label><div class="chk-list" id="kwTek">' +
    daftar.map(function (n) { return '<label class="chk-item"><input type="checkbox" name="tek" value="' + esc(n) + '"' + ((m.tek || []).indexOf(n) >= 0 ? ' checked' : '') + '><span>' + esc(n) + '</span></label>'; }).join('') + '</div>';
}
function kwMulaiHtml(w) {
  var tahu = w.ID_Aset && w.ID_Aset !== 'Belum Tahu', m = S.kw.mulai, pel = m.pel;
  AP && (AP.kategoriDulu = true);
  return '<form id="fMulai" class="k-form" data-submit="kwMulai" novalidate><div class="card-x k-card"><div class="fw-bold mb-2">Lengkapi data pengerjaan lalu tekan Berikutnya</div>' +
    (muatDraft(w.ID_WO) && muatDraft(w.ID_WO).mulai ? '<div class="alert-x warn small"><i class="bi bi-save me-1"></i>Isian sebelumnya dipulihkan dari draft.</div>' : '') +
    '<div class="row g-2"><div class="col-6"><label class="form-label req">Tanggal Pengerjaan</label><input type="date" class="form-control" name="tgl" value="' + esc(m.tgl) + '"></div>' +
    '<div class="col-6"><label class="form-label req">Jam Mulai</label><input type="time" class="form-control" name="jam" value="' + esc(m.jam) + '"></div>' +
    '<div class="col-12"><button type="button" class="btn btn-light-x btn-sm" data-act="kwSekarang"><i class="bi bi-clock me-1"></i>Pakai tanggal &amp; jam sekarang</button></div></div>' +
    '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Data aset</div>' +
    (tahu ? '<div class="small"><b>' + esc(w.Nama_Aset) + '</b><br><span class="text-muted">' + esc(w.ID_Aset) + ' · ' + esc(w.Kategori_Aset) + '<br>' + esc(lokasiLengkap(w)) + '</span></div>' : apHtml()) +
    '<div class="row g-2 mt-2"><div class="col-12"><label class="form-label req">Pelaksana</label><select class="form-select" name="pel" data-kpel="1">' + opsi(listMaster('Pelaksana'), pel, null) + '</select></div>' +
    '<div class="col-12" id="kwNama">' + kwNamaHtml(pel, m) + '</div></div>' +
    '<div class="alert-x danger d-none mt-3 mb-0" id="mulaiErr"></div></div>' +
    '<button class="btn btn-primary k-btn" type="submit">Berikutnya<i class="bi bi-arrow-right ms-2"></i></button></form>';
}
ACTIONS.kwSekarang = function () { var f = $('#fMulai'); if (!f) return; f.tgl.value = hariIniLokal(); f.jam.value = jamLokal(); kwSimpanMulai(); toast('Diisi dengan tanggal & jam sekarang.', 'ok', 1800); };
document.addEventListener('change', function (e) {
  var t = e.target; if (!t.closest || !t.closest('#fMulai')) return;
  if (t.dataset && t.dataset.kpel) { kwSimpanMulai(); S.kw.mulai.pel = t.value; $('#kwNama').innerHTML = kwNamaHtml(t.value, S.kw.mulai); }
  kwSimpanMulai();
});
document.addEventListener('input', function (e) { var t = e.target; if (t.closest && t.closest('#fMulai')) kwSimpanMulai(); });
SUBMITS.kwMulai = async function (f) {
  var w = S.kw.d.wo, err = $('#mulaiErr'); err.classList.add('d-none');
  var salah = function (m) { err.textContent = m; err.classList.remove('d-none'); };
  var p = { Tanggal_Pengerjaan: f.tgl.value, Jam_Mulai_Rencana: f.jam.value, Pelaksana: f.pel.value };
  if (!p.Tanggal_Pengerjaan || !p.Jam_Mulai_Rencana) { salah('Isi tanggal dan jam mulai.'); return; }
  if (p.Pelaksana === 'Vendor') { p.Nama_Vendor = f.nama.value.trim(); if (!p.Nama_Vendor) { salah('Nama vendor wajib diisi.'); return; } }
  else {
    p.Nama_Teknisi = f.tekBebas ? [f.tekBebas.value.trim()] : $$('#fMulai [name=tek]:checked').map(function (c) { return c.value; });
    if (!p.Nama_Teknisi.length || !p.Nama_Teknisi[0]) { salah('Pilih minimal satu nama teknisi.'); return; }
  }
  if (!(w.ID_Aset && w.ID_Aset !== 'Belum Tahu')) { try { Object.assign(p, apNilai()); } catch (e) { salah(e.message); return; } }
  var btn = $('#fMulai button[type=submit]'); sibuk(btn, true, 'Memproses…');
  try { var rm = await call('mulaiWo', w.ID_WO, p); hapusDraftMulai(w.ID_WO); toast(rm && rm.peringatan ? rm.peringatan : 'Pekerjaan dimulai.', rm && rm.peringatan ? 'warn' : 'ok', rm && rm.peringatan ? 8000 : undefined); await kwMuatDetail('lapor'); }
  catch (e) { salah(e.message); } finally { sibuk(btn, false); }
};
function hapusDraftMulai(id) { var o = muatDraft(id); if (!o) return; delete o.mulai; try { if (o.lapor) LS.set(dkey(id), JSON.stringify(o)); else LS.del(dkey(id)); } catch (e) {} S.kw.draftMulai = false; }

/* Langkah 2: laporan hasil */
function kwAsetTerdaftar(w) { return !!(w.ID_Aset && w.ID_Aset !== 'Belum Tahu' && cariAset(w.ID_Aset)); }
function kwInitLapor(w) {
  var f = { hasil: '', alasan: '', desk: '', penyebab: w.Penyebab_Kerusakan || '', tindakan: w.Tindakan || '', catatan: '', verif: '', tgl: hariIniLokal(), kondisi: '', lokDetail: w.Lokasi_Detail || '', iThn: '', iNup: '', iMerk: '', iSeri: '', iStatus: '' };
  try { var ib = JSON.parse(w.Info_Aset_Baru || '{}') || {}; f.iThn = ib.Tahun_Perolehan || ''; f.iNup = ib.NUP || ''; f.iMerk = ib.Merk_Type || ''; f.iSeri = ib.Nomor_Seri || ''; f.iStatus = ib.Status_Aset || ''; f.kondisi = w.Kondisi_Dilaporkan || ''; } catch (e) {}
  var fo = { b1: w.Foto_Before_1 || '', b2: w.Foto_Before_2 || '', a1: w.Foto_After_1 || '', a2: w.Foto_After_2 || '', ast: w.Foto_Aset_URL || '' };
  var docs = dokumenWo(w), sp = sparepartWo(w).map(function (x) { return { u: x.u, c: x.c }; }), dr = muatDraft(w.ID_WO);
  S.kw.dipulihkan = false;
  if (dr && dr.lapor) {
    Object.keys(dr.lapor.form || {}).forEach(function (k) { if (dr.lapor.form[k] !== undefined && dr.lapor.form[k] !== '') f[k] = dr.lapor.form[k]; });
    Object.keys(dr.lapor.foto || {}).forEach(function (k) { if (dr.lapor.foto[k]) fo[k] = dr.lapor.foto[k]; });
    if (dr.lapor.docs && dr.lapor.docs.length) docs = dr.lapor.docs;
    if (dr.lapor.sp && dr.lapor.sp.length) sp = dr.lapor.sp;
    S.kw.dipulihkan = true;
  }
  S.kw.form = f; S.kw.foto = fo; S.kw.docs = docs; S.kw.sp = sp; S.kw.draftLapor = true;
}
/** Kotak foto: ketuk = kamera langsung (HP), tombol "Galeri" = pilih dari galeri. */
function slotHtml(key, label, u) {
  return '<div class="foto-slot' + (u ? ' terisi' : '') + '" data-slot="' + key + '">' + (u ? '<img src="' + esc(imgU(u)) + '" alt="' + esc(label) + '" referrerpolicy="no-referrer">' : '<i class="bi bi-camera"></i><span>Ketuk untuk ambil foto</span>') +
    '<span class="lbl-s">' + esc(label) + '</span><input type="file" accept="image/*" capture="environment" data-slot-in="' + key + '" aria-label="Ambil foto ' + esc(label) + '">' +
    '<label class="gal-btn"><i class="bi bi-images me-1"></i>Galeri<input type="file" accept="image/*" data-slot-in="' + key + '" aria-label="Pilih foto dari galeri ' + esc(label) + '"></label></div>';
}
/** Tombol ambil foto (kamera) + pilih dari galeri untuk daftar foto (sparepart). */
function inputFotoHtml(idGal, idCam, off) {
  return '<div class="d-flex flex-wrap gap-2"><label class="btn btn-primary btn-sm' + (off ? ' disabled' : '') + '"><i class="bi bi-camera-fill me-1"></i>Ambil foto<input type="file" id="' + idCam + '" accept="image/*" capture="environment" hidden' + (off ? ' disabled' : '') + '></label>' +
    '<label class="btn btn-light-x btn-sm' + (off ? ' disabled' : '') + '"><i class="bi bi-images me-1"></i>Pilih dari galeri<input type="file" id="' + idGal + '" accept="image/*" multiple hidden' + (off ? ' disabled' : '') + '></label></div>';
}
function kwSlot(key, label) { return slotHtml(key, label, S.kw.foto[key]); }
var SP_MIN = 4, SP_MAKS = 10;
/** WO yang sebelumnya Ditunda karena Menunggu Sparepart → laporan Selesai wajib melampirkan foto sparepart yang diganti. */
function kwPerluSp() {
  var d = S.kw.d; if (!d) return false;
  return d.wo.Alasan_Tunda === 'Menunggu Sparepart' || (d.log || []).some(function (x) { return x.Status_Baru === 'Ditunda' && String(x.Keterangan || '').indexOf('Menunggu Sparepart') === 0; });
}
function kpPerluSp() {
  var d = S.kp.d; if (!d) return false;
  return d.pm.Alasan_Tunda === 'Menunggu Sparepart' || (d.log || []).some(function (x) { return x.Status_Baru === 'Ditunda' && String(x.Keterangan || '').indexOf('Menunggu Sparepart') === 0; });
}
function kwSpHtml() {
  var a = S.kw.sp || [];
  return '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Foto sparepart yang diganti</div>' +
    '<div class="form-text mb-2">Minimal <b>' + SP_MIN + '</b> dan maksimal <b>' + SP_MAKS + '</b> foto, masing-masing wajib diberi keterangan. Terisi: <b id="kwSpJml">' + a.length + '</b>/' + SP_MAKS + '.</div>' +
    inputFotoHtml('kwSpIn', 'kwSpCam', a.length >= SP_MAKS) +
    '<div id="kwSpList" class="mt-2">' + a.map(function (x, i) {
      return '<div class="sp-item"><div class="foto-th"><img src="' + esc(imgU(x.u)) + '" alt="Sparepart ' + (i + 1) + '" referrerpolicy="no-referrer" data-fb="1"></div>' +
        '<div class="flex-grow-1"><label class="form-label small mb-1 req">Keterangan foto ' + (i + 1) + '</label><input class="form-control form-control-sm" maxlength="150" data-spc="' + i + '" value="' + esc(x.c) + '" placeholder="mis. Kapasitor 35 µF lama yang diganti">' +
        '<button type="button" class="btn-link-x text-danger small mt-1" data-act="kwSpHapus" data-i="' + i + '">hapus foto</button></div></div>';
    }).join('') + '</div>';
}
ACTIONS.kwSpHapus = function (el) { S.kw.sp.splice(Number(el.dataset.i), 1); kwGambarLapor(); };
function kwAlasanBoleh() {
  var pernah = S.kw.d.log.some(function (x) { return x.Status_Baru === 'Ditunda'; }), semua = listMaster('Alasan Penundaan').filter(function (a) { return a !== 'Dialihkan ke Pekerjaan Darurat'; });
  return pernah ? semua.filter(function (a) { return ['Kendala Akses, Cuaca, atau Jadwal', 'Menunggu Persetujuan'].indexOf(a) >= 0; }) : semua;
}
function kwLaporIsi(w) {
  var f = S.kw.form, h = f.hasil, sparepart = h === 'Ditunda' && f.alasan === 'Menunggu Sparepart', terdaftar = kwAsetTerdaftar(w), a0 = terdaftar ? cariAset(w.ID_Aset) : null;
  var seg = '<div class="seg">' + [['Selesai', 's-ok'], ['Ditunda', 's-warn'], ['Dibatalkan', 's-bad']].map(function (x) { return '<button type="button" class="' + x[1] + (h === x[0] ? ' on' : '') + '" data-act="kwHasil" data-v="' + x[0] + '">' + x[0] + '</button>'; }).join('') + '</div>';
  var sebelum = '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Kondisi sebelum</div><div class="foto-grid">' + kwSlot('b1', 'Sebelum 1') + kwSlot('b2', 'Sebelum 2') + '</div>' +
    '<label class="form-label req mt-2">Penyebab kerusakan</label><textarea class="form-control" rows="3" maxlength="1000" data-kwf2="penyebab">' + esc(f.penyebab) + '</textarea>';
  var sesudah = '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Hasil pekerjaan</div><div class="foto-grid">' + kwSlot('a1', 'Sesudah 1') + kwSlot('a2', 'Sesudah 2') + '</div>' +
    '<label class="form-label req mt-2">Tindakan yang dilakukan</label><textarea class="form-control" rows="3" maxlength="1000" data-kwf2="tindakan">' + esc(f.tindakan) + '</textarea>' +
    '<label class="form-label mt-2">Dokumen (opsional, JPG/PNG/PDF)</label><div class="mb-2"><label class="btn btn-light-x btn-sm"><i class="bi bi-camera-fill me-1"></i>Ambil foto dokumen<input type="file" id="kwDokCam" accept="image/*" capture="environment" hidden></label></div><input type="file" class="form-control" id="kwDok" multiple accept="image/*,application/pdf"><div class="form-text">Dokumen PDF maksimal <b>2 MB</b> per berkas. Foto dikompres otomatis.</div>' +
    '<div id="kwDokList" class="mt-2">' + kwDokHtml() + '</div>';
  var cabang = '';
  if (h === 'Selesai') cabang = sebelum + (kwPerluSp() ? kwSpHtml() : '') + sesudah;
  else if (h === 'Ditunda') {
    cabang = '<label class="form-label req mt-3">Alasan penundaan</label><select class="form-select" data-kwf2="alasan">' + opsi(kwAlasanBoleh(), f.alasan, '— pilih alasan —') + '</select>' +
      '<label class="form-label req mt-2">Deskripsi penundaan</label><textarea class="form-control" rows="3" maxlength="1000" data-kwf2="desk" placeholder="Jelaskan alasan penundaan">' + esc(f.desk) + '</textarea>' + (sparepart ? sebelum : '');
  }
  var kondisi = '', asetBaru = '';
  if (h && terdaftar) kondisi = '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Kondisi aset</div><label class="form-label' + (h === 'Dibatalkan' ? '' : ' req') + '">Kondisi aset saat ini (menurut Anda)</label><select class="form-select" data-kwf2="kondisi">' + opsi(listMaster('Kondisi Aset'), f.kondisi, '— pilih kondisi —') + '</select>' +
    '<div class="form-text">Sebelumnya tercatat: <b>' + esc(a0.Kondisi) + '</b>. Jika memilih <b>Rusak Ringan</b> atau <b>Rusak Berat</b>, Master Admin otomatis mendapat pemberitahuan untuk mengubah kondisi aset.</div>';
  if (h && !terdaftar) asetBaru = '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Aset belum terdaftar</div><div class="alert-x info small mb-2"><i class="bi bi-info-circle me-1"></i>Aset ini belum ada di Data Aset. ' + (h === 'Selesai' ? 'Unggah foto aset' : 'Foto aset') + ' dan isi lokasi detail; datanya dikirim ke Master Admin untuk didaftarkan.</div>' +
    '<div class="foto-grid">' + kwSlot('ast', 'Foto aset') + '</div><label class="form-label mt-2">Lokasi detail <span class="text-muted fw-normal">(ketik bebas)</span></label><input class="form-control" data-kwf2="lokDetail" maxlength="150" value="' + esc(f.lokDetail) + '" placeholder="mis. dekat jendela barat">' +
    '<div class="row g-2 mt-1"><div class="col-6"><label class="form-label">Tahun perolehan <span class="text-muted fw-normal">(jika ada)</span></label><input class="form-control" inputmode="numeric" maxlength="4" data-kwf2="iThn" value="' + esc(f.iThn) + '" placeholder="mis. 2019"></div>' +
    '<div class="col-6"><label class="form-label">NUP <span class="text-muted fw-normal">(jika ada)</span></label><input class="form-control" maxlength="60" data-kwf2="iNup" value="' + esc(f.iNup) + '"></div>' +
    '<div class="col-6"><label class="form-label">Merk / Type <span class="text-muted fw-normal">(jika ada)</span></label><input class="form-control" maxlength="150" data-kwf2="iMerk" value="' + esc(f.iMerk) + '"></div>' +
    '<div class="col-6"><label class="form-label">Nomor seri <span class="text-muted fw-normal">(jika ada)</span></label><input class="form-control" maxlength="100" data-kwf2="iSeri" value="' + esc(f.iSeri) + '"></div>' +
    '<div class="col-6"><label class="form-label' + (h === 'Dibatalkan' ? '' : ' req') + '">Kondisi aset</label><select class="form-select" data-kwf2="kondisi">' + opsi(listMaster('Kondisi Aset'), f.kondisi, '— pilih —') + '</select></div>' +
    '<div class="col-6"><label class="form-label' + (h === 'Dibatalkan' ? '' : ' req') + '">Status aset</label><select class="form-select" data-kwf2="iStatus">' + opsi(listMaster('Status Aset').filter(function (x) { return x !== DIHAPUS; }), f.iStatus, '— pilih —') + '</select></div></div>' +
    '<div class="form-text">Jika Rusak Ringan atau Rusak Berat dipilih, Master Admin otomatis mendapat pemberitahuan.</div>';
  var verif = '';
  if (h) verif = '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Verifikasi</div><div class="row g-2"><div class="col-12"><label class="form-label req">Verifikator (SPV)</label><select class="form-select" data-kwf2="verif">' + opsi(WO.verif || [], f.verif, '— pilih verifikator —') + '</select></div>' +
    '<div class="col-6"><label class="form-label req">Tanggal verifikasi</label><input type="date" class="form-control" data-kwf2="tgl" max="' + hariIniLokal() + '" value="' + esc(f.tgl) + '"><button type="button" class="btn-link-x small mt-1" data-act="kwHariIni">Pakai hari ini</button></div>' +
    '<div class="col-6"><label class="form-label">Status</label><select class="form-select" disabled><option>Terverifikasi</option></select></div></div>';
  var catatan = h ? '<label class="form-label mt-3">Catatan (opsional)</label><textarea class="form-control" rows="2" maxlength="1000" data-kwf2="catatan" placeholder="Catatan tambahan untuk verifikator / Master Admin">' + esc(f.catatan) + '</textarea>' : '';
  return '<div class="card-x k-card k-form"><div class="fw-bold mb-2">Lapor hasil pekerjaan</div>' + (S.kw.dipulihkan ? '<div class="alert-x warn small"><i class="bi bi-save me-1"></i>Draft laporan dipulihkan. Isian Anda tersimpan otomatis di perangkat ini.</div>' : '') +
    '<label class="form-label req">Status hasil</label>' + seg + cabang + kondisi + asetBaru + catatan + verif +
    '<div class="small text-warning mt-3 d-none" id="kwKurang"></div></div>' +
    '<button class="btn btn-primary k-btn" id="kwSimpan" data-act="kwSimpan" disabled><i class="bi bi-save me-2"></i>Simpan Laporan</button>';
}
function kwDokHtml() {
  return S.kw.docs.map(function (x, i) { return '<div class="d-flex align-items-center gap-2 small mb-1"><i class="bi bi-paperclip"></i><span class="text-truncate flex-grow-1">' + esc(x.n) + '</span><button type="button" class="icon-btn text-danger" data-act="kwDokHapus" data-i="' + i + '" title="Hapus dokumen" aria-label="Hapus dokumen ke-' + (i + 1) + '"><i class="bi bi-trash"></i></button></div>'; }).join('');
}
function kwGambarLapor() { var b = $('#kwLaporBox'); if (b) { b.innerHTML = kwLaporIsi(S.kw.d.wo); kwCekTombol(); } }
ACTIONS.kwHasil = function (el) { S.kw.form.hasil = el.dataset.v; if (el.dataset.v !== 'Ditunda') { S.kw.form.alasan = ''; } kwGambarLapor(); };
ACTIONS.kwDokHapus = function (el) { S.kw.docs.splice(Number(el.dataset.i), 1); $('#kwDokList').innerHTML = kwDokHtml(); simpanDraft(); };
ACTIONS.kwHariIni = function () { S.kw.form.tgl = hariIniLokal(); var i = $('[data-kwf2=tgl]'); if (i) i.value = S.kw.form.tgl; kwCekTombol(); };
function kwKurang() {
  var f = S.kw.form, fo = S.kw.foto, k = [];
  if (!f.hasil) return ['Pilih status hasil'];
  if (f.hasil === 'Selesai') {
    if (!fo.b1 || !fo.b2) k.push('2 foto sebelum'); if (!f.penyebab.trim()) k.push('penyebab kerusakan');
    if (!fo.a1 || !fo.a2) k.push('2 foto sesudah'); if (!f.tindakan.trim()) k.push('tindakan');
    if (kwPerluSp()) {
      var sp = S.kw.sp || [];
      if (sp.length < SP_MIN) k.push('foto sparepart minimal ' + SP_MIN + ' (baru ' + sp.length + ')');
      if (sp.some(function (x) { return !String(x.c || '').trim(); })) k.push('keterangan tiap foto sparepart');
    }
  } else if (f.hasil === 'Ditunda') {
    if (!f.alasan) k.push('alasan penundaan'); if (!f.desk.trim()) k.push('deskripsi penundaan');
    if (f.alasan === 'Menunggu Sparepart') { if (!fo.b1 || !fo.b2) k.push('2 foto sebelum'); if (!f.penyebab.trim()) k.push('penyebab kerusakan'); }
  }
  if (f.hasil !== 'Dibatalkan' && !f.kondisi) k.push('kondisi aset');
  if (f.hasil !== 'Dibatalkan' && !kwAsetTerdaftar(S.kw.d.wo) && !f.iStatus) k.push('status aset');
  if (!kwAsetTerdaftar(S.kw.d.wo) && f.iThn && !/^\d{4}$/.test(f.iThn.trim())) k.push('tahun perolehan 4 digit');
  if (f.hasil === 'Selesai' && !kwAsetTerdaftar(S.kw.d.wo) && !fo.ast) k.push('foto aset');
  if (!f.verif) k.push('verifikator'); if (!f.tgl) k.push('tanggal verifikasi');
  return k;
}
function kwCekTombol() {
  simpanDraft();
  var b = $('#kwSimpan'); if (!b) return; var k = kwKurang();
  b.disabled = k.length > 0;
  var el = $('#kwKurang'); if (el) { el.classList.toggle('d-none', !k.length || !S.kw.form.hasil); el.textContent = k.length ? 'Belum lengkap: ' + k.join(', ') + '.' : ''; }
}
document.addEventListener('input', function (e) {
  var t = e.target;
  if (t && t.dataset && t.dataset.spc !== undefined && S.kPage === 'wodetail') { var it = S.kw.sp[Number(t.dataset.spc)]; if (it) it.c = t.value; kwCekTombol(); return; }
  if (!t.dataset || t.dataset.kwf2 === undefined || S.kPage !== 'wodetail') return;
  S.kw.form[t.dataset.kwf2] = t.value; kwCekTombol();
});
document.addEventListener('change', async function (e) {
  var t = e.target; if (S.kPage !== 'wodetail' || !t.dataset) return;
  if (t.dataset.kwf2 !== undefined && t.tagName === 'SELECT') { S.kw.form[t.dataset.kwf2] = t.value; if (t.dataset.kwf2 === 'alasan') kwGambarLapor(); else kwCekTombol(); }
  else if (t.dataset.kwf2 === 'tgl') { S.kw.form.tgl = t.value; kwCekTombol(); }
  else if (t.dataset.slotIn && t.files && t.files[0]) {
    var key = t.dataset.slotIn, slot = t.closest('.foto-slot'); slot.classList.add('naik');
    try {
      var foto = await kompresFoto(t.files[0]);
      var u = await call('uploadFile', 'wo', S.kw.id, { name: foto.name, mime: foto.mime, base64: foto.base64 });
      catatFotoLokal(u.url, foto.preview); S.kw.foto[key] = u.url; var lab = slot.querySelector('.lbl-s').textContent; slot.outerHTML = kwSlot(key, lab); simpanDraft();
    } catch (er) { slot.classList.remove('naik'); t.value = ''; toast(er.message, 'bad'); }
    kwCekTombol();
  } else if ((t.id === 'kwSpIn' || t.id === 'kwSpCam') && t.files && t.files.length) {
    var daftar = Array.prototype.slice.call(t.files), sisa = SP_MAKS - S.kw.sp.length; t.disabled = true;
    if (daftar.length > sisa) toast('Maksimal ' + SP_MAKS + ' foto sparepart; ' + (daftar.length - sisa) + ' foto terakhir diabaikan.', 'warn');
    for (var si = 0; si < Math.min(daftar.length, sisa); si++) {
      try {
        var cf = await kompresFoto(daftar[si]);
        var ur = await call('uploadFile', 'wo', S.kw.id, { name: cf.name, mime: cf.mime, base64: cf.base64 });
        catatFotoLokal(ur.url, cf.preview); S.kw.sp.push({ u: ur.url, c: '' });
      } catch (er3) { toast(er3.message, 'bad'); }
    }
    simpanDraft(); kwGambarLapor();
  } else if ((t.id === 'kwDok' || t.id === 'kwDokCam') && t.files && t.files.length) {
    var list = Array.prototype.slice.call(t.files); t.disabled = true;
    for (var i = 0; i < list.length; i++) {
      var file = list[i];
      try {
        var body;
        if (file.type === 'application/pdf') {
          if (file.size > 2 * 1024 * 1024) throw new Error(file.name + ': ukuran dokumen PDF maksimal 2 MB (ukuran ' + (file.size / 1048576).toFixed(1).replace('.', ',') + ' MB).');
          body = { name: file.name, mime: 'application/pdf', base64: String(await bacaFile(file, false)).split(',')[1] };
        } else { var c = await kompresFoto(file); body = { name: c.name, mime: c.mime, base64: c.base64 }; }
        var r = await call('uploadFile', 'wo', S.kw.id, body); S.kw.docs.push({ n: file.name.slice(0, 80), u: r.unduh });
      } catch (er2) { toast(er2.message, 'bad'); }
    }
    t.value = ''; t.disabled = false; var dl = $('#kwDokList'); if (dl) dl.innerHTML = kwDokHtml(); simpanDraft();
  }
});
ACTIONS.kwSimpan = async function (el) {
  var f = S.kw.form, fo = S.kw.foto, w = S.kw.d.wo, k = kwKurang(); if (k.length) { toast('Belum lengkap: ' + k.join(', ') + '.', 'warn'); return; }
  var p = { Status_Hasil: f.hasil, Catatan: f.catatan, Verifikator: f.verif, Tanggal_Verifikasi: f.tgl };
  if (f.hasil === 'Selesai' || (f.hasil === 'Ditunda' && f.alasan === 'Menunggu Sparepart')) { p.Foto_Before_1 = fo.b1; p.Foto_Before_2 = fo.b2; p.Penyebab_Kerusakan = f.penyebab; }
  if (f.hasil === 'Selesai') { p.Foto_After_1 = fo.a1; p.Foto_After_2 = fo.a2; p.Tindakan = f.tindakan; p.Dokumen = S.kw.docs; if (kwPerluSp()) p.Foto_Sparepart = S.kw.sp.map(function (x) { return { u: x.u, c: x.c.trim() }; }); }
  if (f.hasil === 'Ditunda') { p.Alasan_Tunda = f.alasan; p.Deskripsi_Tunda = f.desk; }
  if (f.kondisi) p.Kondisi_Dilaporkan = f.kondisi;
  if (!kwAsetTerdaftar(w)) { if (fo.ast) p.Foto_Aset_URL = fo.ast; p.Lokasi_Detail = f.lokDetail; p.Info_Aset = { Tahun_Perolehan: f.iThn.trim(), NUP: f.iNup, Merk_Type: f.iMerk, Nomor_Seri: f.iSeri, Status_Aset: f.iStatus }; }
  if (!(await konfirm('Simpan laporan?', 'Status Work Order akan menjadi "' + f.hasil + '". Pastikan data sudah benar.', 'Simpan', false))) return;
  sibuk(el, true, 'Menyimpan…');
  try {
    var d = await call('laporWo', w.ID_WO, p); hapusDraft(w.ID_WO); S.kw.draftLapor = false; toast(d.peringatan || 'Laporan disimpan.', d.peringatan ? 'warn' : 'ok', d.peringatan ? 8000 : 3500);
    S.kw.tab = f.hasil === 'Ditunda' ? 'ditunda' : 'selesai'; await kwMuatDetail('info');
  } catch (e) { gagal(e); sibuk(el, false); }
};

/* Teknisi membuat Work Order dari daftar WO: pilih aset (cari / lokasi & kategori / ID Aset belum diketahui) */
ACTIONS.kwLaporBaru = function () {
  apBaru(null, ''); AP.kategoriDulu = true;
  bukaModal({ judul: 'Buat Work Order', isi: '<form id="fLB" class="k-form" data-submit="kwLaporBaruKirim" novalidate>' +
    '<div class="alert-x info small"><i class="bi bi-info-circle me-1"></i>Cari aset, atau pilih lewat lokasi &amp; kategori. Aset belum ada di Data Aset? Pilih <b>ID Aset belum diketahui</b>; foto aset diunggah saat melapor dan datanya dikirim ke Master Admin untuk didaftarkan.</div>' +
    apHtml() +
    '<div class="my-3"><label class="form-label req">Uraian masalah / pekerjaan</label><textarea class="form-control" name="uraian" rows="3" maxlength="1000"></textarea></div>' +
    '<div class="mb-3"><label class="form-label">Temuan / indikasi</label><textarea class="form-control" name="temuan" rows="2" maxlength="500"></textarea></div>' +
    '<div class="mb-2"><label class="form-label req">Prioritas</label><select class="form-select" name="prio">' + opsiPrioritas('') + '</select><div class="form-text">Prioritas menentukan lama SLA; SLA mulai dihitung saat Anda menekan Mulai Pengerjaan.</div></div>' +
    '<div class="alert-x danger d-none mb-0" id="lbErr"></div></form>',
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fLB">Buat Work Order</button>' });
};
SUBMITS.kwLaporBaruKirim = async function (f) {
  var err = $('#lbErr'); err.classList.add('d-none');
  var salah = function (m) { err.textContent = m; err.classList.remove('d-none'); };
  var p; try { p = apNilai(); } catch (e) { salah(e.message); return; }
  if (!f.uraian.value.trim() || !f.prio.value) { salah('Isi uraian dan pilih prioritas.'); return; }
  p.Uraian_Pekerjaan = f.uraian.value; p.Temuan_Indikasi = f.temuan.value; p.Prioritas = f.prio.value;
  var btn = $('button[type=submit][form=fLB]'); sibuk(btn, true, 'Membuat…');
  try {
    var d = await call('laporLangsungWo', p); gantiWoLokal(d.wo); tutupModal(); toast('Work Order ' + d.wo.ID_WO + ' dibuat.', 'ok');
    S.kw.id = d.wo.ID_WO; S.kw.dari = 'list'; S.kw.pindai = ''; await kwMuatDetail('info');
  } catch (e) { salah(e.message); } finally { sibuk(btn, false); }
};

/* Halaman aset: tab Work Order + Lapor Langsung */
function kAsetWoHtml(a) {
  if (!WO.dimuat) return '<div class="skeleton" style="height:120px"></div>';
  var tek = adalahTeknisi(), hidup = function (w) { return WO_DIKERJAKAN.indexOf(w.Status_WO) >= 0; };
  var sama = WO.list.filter(function (w) { return w.ID_Aset === a.ID_Aset && hidup(w); });
  var belum = WO.list.filter(function (w) { return (w.ID_Aset === 'Belum Tahu' || !w.ID_Aset) && (w.Status_WO === 'Open' || w.Status_WO === 'Ditunda'); });
  var riw = WO.list.filter(function (w) { return w.ID_Aset === a.ID_Aset && !hidup(w); }).sort(function (x, y) { return x.Waktu_Dibuat < y.Waktu_Dibuat ? 1 : -1; });
  var blok = function (jud, arr, pindai, kosong) { return '<div class="fw-bold small text-muted text-uppercase mb-2 mt-3">' + jud + ' (' + arr.length + ')</div>' + (arr.length ? arr.map(function (w) { return kwCard(w, 'aset', pindai); }).join('') : '<div class="small text-muted mb-2">' + kosong + '</div>'); };
  return blok('WO untuk aset ini', sama, '', 'Tidak ada WO aktif untuk aset ini.') +
    blok('WO tanpa ID Aset ("Belum Tahu")', belum, a.ID_Aset, 'Tidak ada. Bila Anda memilih salah satu, ID aset yang dipindai otomatis diisikan.') +
    blok('Riwayat WO aset ini', riw.slice(0, 10), '', 'Belum ada riwayat.');
}
ACTIONS.kwLapor = function () {
  var a = S.kAset && S.kAset.aset; if (!a) return;
  bukaModal({ judul: 'Lapor Work Order', isi: '<form id="fLL" class="k-form" data-submit="kwLaporLangsung" novalidate><div class="alert-x info small"><b>' + esc(a.Nama_Aset) + '</b><br>' + esc(a.ID_Aset) + ' · ' + esc(lokasiStr(a)) + '</div>' +
    '<div class="mb-3"><label class="form-label req">Uraian masalah / pekerjaan</label><textarea class="form-control" name="uraian" rows="3" maxlength="1000" required></textarea></div>' +
    '<div class="mb-3"><label class="form-label">Temuan / indikasi</label><textarea class="form-control" name="temuan" rows="2" maxlength="500"></textarea></div>' +
    '<div class="mb-2"><label class="form-label req">Prioritas</label><select class="form-select" name="prio">' + opsiPrioritas('') + '</select><div class="form-text">Prioritas menentukan SLA.</div></div>' +
    '<div class="alert-x danger d-none mb-0" id="llErr"></div></form>',
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fLL">Buat &amp; lanjut</button>' });
};
SUBMITS.kwLaporLangsung = async function (f) {
  var a = S.kAset.aset, err = $('#llErr'); err.classList.add('d-none');
  if (!f.uraian.value.trim() || !f.prio.value) { err.textContent = 'Isi uraian dan pilih prioritas.'; err.classList.remove('d-none'); return; }
  var btn = $('button[type=submit][form=fLL]'); sibuk(btn, true, 'Membuat…');
  try {
    var d = await call('laporLangsungWo', { ID_Aset: a.ID_Aset, Uraian_Pekerjaan: f.uraian.value, Temuan_Indikasi: f.temuan.value, Prioritas: f.prio.value });
    gantiWoLokal(d.wo); tutupModal(); toast('Work Order ' + d.wo.ID_WO + ' dibuat.', 'ok');
    S.kw.id = d.wo.ID_WO; S.kw.dari = 'aset'; S.kw.pindai = ''; await kwMuatDetail('info');
  } catch (e) { err.textContent = e.message; err.classList.remove('d-none'); } finally { sibuk(btn, false); }
};

/* ════════════════════════════════════════════════════════
   PREVENTIVE MAINTENANCE — dashboard (jadwal, kalender, tabel, grafik, impor) dan Kartu (teknisi)
   ════════════════════════════════════════════════════════ */
var PM = { list: [], dimuat: false, f: { q: '', st: '', kat: '', pel: '', dari: '', smp: '', urut: '', hal: 1, per: 25 }, view: 'tabel', bulan: '', tgl: '' };
var PM_AKTIF = ['Terjadwal', 'Sedang Dikerjakan', 'Ditunda'];
var KP_TAB = { terjadwal: ['Terjadwal', 'Sedang Dikerjakan'], ditunda: ['Ditunda'], selesai: ['Terlaksana', 'Dibatalkan'] };
S.kp = { tab: 'terjadwal', q: '', kat: '', lok: '', pel: '', tgl: '', hal: 1, id: null, d: null, err: '', langkah: 'info', dari: 'list', mulai: {}, form: {}, foto: {}, docs: [], sp: [], draftLapor: false, draftMulai: false, dipulihkan: false };

async function muatPm() { var d = await call('getPmList'); PM.list = d.pm; PM.dimuat = true; return d; }
function gantiPmLokal(p) {
  var ada = false; PM.list = PM.list.map(function (x) { if (x.ID_PM === p.ID_PM) { ada = true; return p; } return x; });
  if (!ada) PM.list.push(p);
}
function selisihHari(a, b) { // b - a dalam hari (yyyy-mm-dd)
  var pa = /^(\d{4})-(\d{2})-(\d{2})/.exec(a || ''), pb = /^(\d{4})-(\d{2})-(\d{2})/.exec(b || ''); if (!pa || !pb) return NaN;
  return Math.round((Date.UTC(+pb[1], +pb[2] - 1, +pb[3]) - Date.UTC(+pa[1], +pa[2] - 1, +pa[3])) / 86400000);
}
/** Progres PM terhadap Due Date. */
function progresPm(p) {
  if (p.Status_PM === 'Dibatalkan') return { t: 'Dibatalkan', k: '' };
  if (p.Status_PM === 'Terlaksana') {
    var t = selisihHari(p.Due_Date, String(p.Waktu_Selesai).slice(0, 10));
    return isNaN(t) ? { t: 'Terlaksana', k: 'ok' } : (t > 0 ? { t: 'Terlambat ' + t + ' hari', k: 'bad' } : { t: 'Tepat waktu', k: 'ok' });
  }
  var d = selisihHari(hariIniLokal(), p.Due_Date), pre = p.Status_PM === 'Ditunda' ? 'Ditunda · ' : '';
  if (isNaN(d)) return { t: '—', k: '' };
  if (d < 0) return { t: pre + 'Terlambat ' + (-d) + ' hari', k: 'bad' };
  if (d === 0) return { t: pre + 'Jatuh tempo hari ini', k: 'warn' };
  return { t: pre + 'Sisa ' + d + ' hari', k: d <= 3 ? 'warn' : 'ok' };
}
function bProgresPm(p) { var x = progresPm(p); return badge(x.t, x.k); }
function bPrioPm(v) { return badge(v, { 'Tinggi': 'bad', 'Sedang': 'warn', 'Rendah': '' }[v] || ''); }
function dokumenPm(p) { try { return JSON.parse(p.Dokumen_URLs || '[]'); } catch (e) { return []; } }
function infoAsetBaru(p) { try { return JSON.parse(p.Info_Aset_Baru || '{}') || {}; } catch (e) { return {}; } }

/* ── Detail (dipakai modal dashboard dan halaman Kartu) ── */
function stepsPm(p) {
  var s = p.Status_PM, mulai = !!p.Waktu_Mulai;
  var l2 = s === 'Ditunda' ? { n: 'Ditunda', c: 'warn', ic: 'bi-pause-fill' } : { n: 'Sedang Dikerjakan', c: mulai ? 'done' : '', ic: 'bi-check-lg' };
  var l3 = s === 'Dibatalkan' ? { n: 'Dibatalkan', c: 'bad', ic: 'bi-x-lg', t: p.Waktu_Selesai } : { n: 'Terlaksana', c: s === 'Terlaksana' ? 'done' : '', ic: 'bi-check-lg', t: s === 'Terlaksana' ? p.Waktu_Selesai : '' };
  var st = function (n, c, ic, t) { return '<div class="st ' + c + '"><div class="dot"><i class="bi ' + ic + '"></i></div><b>' + esc(n) + '</b>' + (t ? esc(fmtTgl(t)) : '') + '</div>'; };
  return '<div class="steps">' + st('Terjadwal', 'done', 'bi-check-lg', p.Dibuat_Pada) + st(l2.n, l2.c, l2.ic, p.Waktu_Mulai) + st(l3.n, l3.c, l3.ic, l3.t) + '</div>';
}
function laporanPmHtml(p, admin) {
  var dok = dokumenPm(p), ib = infoAsetBaru(p);
  var ada = p.Foto_Before_1 || p.Foto_Sparepart || p.Hasil_Pemeriksaan || p.Kondisi_Dilaporkan || p.Foto_Aset_URL || p.Foto_After_1 || p.Tindakan || p.Catatan || p.Verifikator || dok.length;
  if (!ada) return '';
  var foto = function (a, b, l) { var h = fotoThumb(a, l + ' 1') + fotoThumb(b, l + ' 2'); return h ? '<div class="foto-grid mb-2">' + h + '</div>' : ''; };
  return woSeksi('Laporan Teknisi',
    foto(p.Foto_Before_1, p.Foto_Before_2, 'Sebelum') + sparepartHtml(p) +
    '<dl class="dl">' + dlRow('Hasil Pemeriksaan', teksPanjang(p.Hasil_Pemeriksaan)) + dlRow('Temuan', p.Ada_Temuan ? (p.Ada_Temuan === 'Ya' ? badge('Ada temuan', 'warn') : badge('Tidak ada temuan', 'ok')) : '') + dlRow('Deskripsi Temuan', teksPanjang(p.Temuan)) + '</dl>' +
    '<hr class="my-2">' + foto(p.Foto_After_1, p.Foto_After_2, 'Sesudah') +
    '<dl class="dl">' + dlRow('Tindakan', teksPanjang(p.Tindakan)) +
    dlRow('Dokumen', dok.length ? dok.map(function (x) { return '<a href="' + esc(x.u) + '" target="_blank" rel="noopener"><i class="bi bi-paperclip me-1"></i>' + esc(x.n) + '</a>'; }).join('<br>') : '') +
    dlRow('Kondisi Aset (menurut teknisi)', p.Kondisi_Dilaporkan ? bKondisi(p.Kondisi_Dilaporkan) : '') +
    (p.Foto_Aset_URL ? dlRow('Foto Aset (belum terdaftar)', fotoThumb(p.Foto_Aset_URL, 'Foto aset')) : '') +
    (ib.Status_Aset || ib.Merk_Type || ib.Nomor_Seri || ib.NUP || ib.Tahun_Perolehan ? dlRow('Data Aset Baru (dari teknisi)', esc([ib.Tahun_Perolehan && 'Tahun ' + ib.Tahun_Perolehan, ib.NUP && 'NUP ' + ib.NUP, ib.Merk_Type, ib.Nomor_Seri && 'SN ' + ib.Nomor_Seri, ib.Status_Aset && 'Status ' + ib.Status_Aset].filter(Boolean).join(' · '))) : '') +
    dlRow('Catatan', teksPanjang(p.Catatan)) + dlRow('Verifikator', esc(p.Verifikator)) + dlRow('Tanggal Verifikasi', p.Tanggal_Verifikasi ? esc(fmtTgl(p.Tanggal_Verifikasi, true)) : '') + dlRow('Status Verifikasi', p.Status_Verifikasi ? badge(p.Status_Verifikasi, 'ok') : '') +
    (admin ? dlRow('Kondisi Aset Sebelum', esc(p.Kondisi_Aset_Sebelum)) + dlRow('Kondisi Aset Sesudah', esc(p.Kondisi_Aset_Sesudah)) + dlRow('Status Aset Sesudah', esc(p.Status_Aset_Sesudah)) : '') + '</dl>');
}
function detailPmHtml(d) {
  var p = d.pm, a = d.aset;
  return stepsPm(p) +
    woSeksi('Identitas & Status', '<dl class="dl">' + dlRow('ID PM', '<b>' + esc(p.ID_PM) + '</b>') + dlRow('Status', bStatusWo(p.Status_PM) + ' ' + bProgresPm(p)) + dlRow('Sumber', esc(p.Sumber_PM || 'Terjadwal')) + dlRow('Dibuat', esc(fmtTgl(p.Dibuat_Pada))) + '</dl>') +
    woSeksi('Objek Aset & Lokasi', '<dl class="dl">' + dlRow('ID Aset', esc(p.ID_Aset)) + dlRow('Nama Aset', esc(p.Nama_Aset)) + dlRow('Kategori', esc(p.Kategori_Aset)) +
      (a ? dlRow('Kritikalitas', bKrit(a.Kritikalitas)) + dlRow('Kondisi', bKondisi(a.Kondisi)) + dlRow('Status Aset', bStatus(a.Status_Aset)) : '') +
      dlRow('Gedung', esc(p.Gedung)) + dlRow('Lantai', esc(p.Lantai)) + dlRow('Ruangan', esc(p.Ruangan)) + dlRow('Lokasi Detail', esc(p.Lokasi_Detail)) +
      dlRow('PIC', esc(p.PIC_Aset || (a ? picLabel(a.PIC) : ''))) + '</dl>') +
    woSeksi('Rencana Pemeliharaan', '<dl class="dl">' + dlRow('Tanggal Jadwal', esc(fmtTgl(p.Tanggal_Jadwal, true))) + dlRow('Due Date', esc(fmtTgl(p.Due_Date, true))) + dlRow('Jenis PM', esc(p.Jenis_PM)) +
      dlRow('Frekuensi PM', esc(p.Frekuensi_PM || (a && a.Frekuensi_PM) || '')) + dlRow('Prioritas Perawatan', bPrioPm(p.Prioritas_Perawatan)) + dlRow('Aktivitas / Checklist', teksPanjang(p.Aktivitas_Checklist)) + '</dl>') +
    woSeksi('Penugasan', '<dl class="dl">' + dlRow('Pelaksana', esc(p.Pelaksana)) + dlRow(p.Pelaksana === 'Vendor' ? 'Vendor' : 'Teknisi', esc(p.Nama_Teknisi_atau_Vendor)) + dlRow('Dimulai oleh', esc(p.Dimulai_Oleh)) +
      dlRow('Waktu Mulai', p.Waktu_Mulai ? esc(fmtTgl(p.Waktu_Mulai)) : '') + dlRow('Waktu Selesai', p.Waktu_Selesai ? esc(fmtTgl(p.Waktu_Selesai)) : '') + '</dl>') +
    (p.Status_PM === 'Ditunda' ? woSeksi('Alasan Penundaan', '<dl class="dl">' + dlRow('Alasan', esc(p.Alasan_Tunda)) + dlRow('Keterangan', teksPanjang(p.Deskripsi_Tunda)) + '</dl>') : '') +
    laporanPmHtml(p, isAdmin() || bolehDash()) +
    woSeksi('Progres', timelineWo(d.log));
}

/* ════════ Dashboard: halaman Preventive Maintenance ════════ */
function pmTersaring() {
  var f = PM.f, q = f.q.trim();
  var l0 = PM.list.filter(function (p) {
    if (f.st && p.Status_PM !== f.st) return false;
    if (f.kat && p.Kategori_Aset !== f.kat) return false;
    if (f.pel && p.Pelaksana !== f.pel) return false;
    if (f.dari && p.Tanggal_Jadwal < f.dari) return false;
    if (f.smp && p.Tanggal_Jadwal > f.smp) return false;
    return true;
  });
  if (q) return cariUrut(l0, q, function (p) { return p.ID_PM + ' ' + p.Nama_Aset + ' ' + p.ID_Aset + ' ' + p.Kategori_Aset + ' ' + p.Jenis_PM + ' ' + p.Lantai + ' ' + p.Ruangan + ' ' + p.Gedung + ' ' + p.Nama_Teknisi_atau_Vendor; });
  return l0.sort(function (a, b) {
    var c = a.Tanggal_Jadwal < b.Tanggal_Jadwal ? -1 : (a.Tanggal_Jadwal > b.Tanggal_Jadwal ? 1 : 0);
    if (f.urut === 'tgl') return c; if (f.urut === 'tgl-') return -c;
    return a.Tanggal_Jadwal < b.Tanggal_Jadwal ? 1 : (a.Tanggal_Jadwal > b.Tanggal_Jadwal ? -1 : 0);
  });
}
ACTIONS.pmUrut = function () { PM.f.urut = PM.f.urut === '' ? 'tgl' : (PM.f.urut === 'tgl' ? 'tgl-' : ''); PM.f.hal = 1; renderPm(); };
async function renderPm() {
  if (PM.view === 'grafik') { renderPmGrafik(); return; }
  $('#main').innerHTML = '<div class="skeleton" style="height:90px"></div><div class="skeleton mt-3" style="height:360px"></div>';
  try { await muatPm(); } catch (e) {
    if (S.page === 'pm') $('#main').innerHTML = '<div class="card-x"><div class="empty"><i class="bi bi-exclamation-circle"></i>' + esc(e.message) + '<div class="mt-3"><button class="btn btn-primary" data-act="pmMuat">Coba lagi</button></div></div></div>';
    return;
  }
  if (S.page !== 'pm' || PM.view === 'grafik') return;
  if (!PM.bulan) PM.bulan = hariIniLokal().slice(0, 7);
  var f = PM.f, statusList = listMaster('Status PM'); if (!statusList.length) statusList = ['Terjadwal', 'Sedang Dikerjakan', 'Ditunda', 'Terlaksana', 'Dibatalkan'];
  $('#main').innerHTML =
    '<div class="page-head"><h2>Preventive Maintenance</h2><div class="ms-auto d-flex flex-wrap gap-2">' +
    '<button class="btn btn-light-x" data-act="pmEkspor"><i class="bi bi-file-earmark-excel me-1"></i>Ekspor Excel</button>' +
    '<button class="btn btn-light-x" data-act="pmGrafik"><i class="bi bi-bar-chart-line me-1"></i>Lihat Grafik</button>' +
    '<button class="btn btn-light-x" data-act="pmImpor"><i class="bi bi-upload me-1"></i>Impor CSV</button>' +
    '<button class="btn btn-primary" data-act="pmBaru"><i class="bi bi-plus-lg me-1"></i>Jadwal PM</button></div></div>' +
    '<div class="kpi-grid" id="pmStat"></div>' +
    '<div class="grid-2"><div class="card-x"><div class="card-h"><i class="bi bi-calendar3"></i>Kalender PM<div class="ms-auto d-flex gap-1"><button class="icon-btn" data-act="pmBulan" data-d="-1" aria-label="Bulan sebelumnya"><i class="bi bi-chevron-left"></i></button><button class="icon-btn" data-act="pmBulan" data-d="1" aria-label="Bulan berikutnya"><i class="bi bi-chevron-right"></i></button></div></div><div class="card-b" id="pmKal"></div></div>' +
    '<div class="card-x"><div class="card-h" id="pmHariJud"></div><div class="card-b" id="pmHari"></div></div></div>' +
    '<div class="card-x"><div class="toolbar">' +
    '<div class="grow input-icon"><i class="bi bi-search"></i><input class="form-control" data-pf="q" placeholder="Cari ID PM, aset, jenis, teknisi…" value="' + esc(f.q) + '" aria-label="Cari PM"></div>' +
    '<select class="form-select fsel" data-pf="st" aria-label="Status">' + opsi(statusList, f.st, 'Semua status') + '</select>' +
    '<select class="form-select fsel" data-pf="kat" aria-label="Kategori">' + opsi(listMaster('Kategori Aset'), f.kat, 'Semua kategori') + '</select>' +
    '<select class="form-select fsel" data-pf="pel" aria-label="Pelaksana">' + opsi(listMaster('Pelaksana'), f.pel, 'Semua pelaksana') + '</select>' +
    '<div class="d-flex align-items-center gap-1"><input type="date" class="form-control" style="max-width:150px" data-pf="dari" value="' + esc(f.dari) + '" aria-label="Jadwal dari"><span class="text-muted">–</span><input type="date" class="form-control" style="max-width:150px" data-pf="smp" value="' + esc(f.smp) + '" aria-label="Jadwal sampai"></div>' +
    '<button class="btn btn-light-x btn-sm" data-act="pmReset">Reset</button></div>' +
    '<div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>PM</th><th><button type="button" class="th-sort" data-act="pmUrut" title="Urutkan berdasarkan tanggal jadwal">Jadwal / Due Date <i class="bi ' + (f.urut === 'tgl' ? 'bi-sort-numeric-down' : (f.urut === 'tgl-' ? 'bi-sort-numeric-up' : 'bi-arrow-down-up')) + '"></i></button></th><th>Aset &amp; Lokasi</th><th>Pelaksana</th><th>Status</th><th>Progres</th><th></th></tr></thead><tbody id="pmBody"></tbody></table></div><div class="pager" id="pmPager"></div></div>';
  isiPm();
}
ACTIONS.pmMuat = function () { renderPm(); };
ACTIONS.pmReset = function () { PM.f = { q: '', st: '', kat: '', pel: '', dari: '', smp: '', urut: '', hal: 1, per: 25 }; renderPm(); };
function pmKalender() {
  var bl = PM.bulan, th = Number(bl.slice(0, 4)), bn = Number(bl.slice(5, 7)) - 1;
  var pertama = new Date(th, bn, 1), mulaiKol = (pertama.getDay() + 6) % 7, hari = new Date(th, bn + 1, 0).getDate(), ada = {};
  PM.list.forEach(function (p) { if (p.Tanggal_Jadwal && p.Tanggal_Jadwal.slice(0, 7) === bl) ada[p.Tanggal_Jadwal] = (ada[p.Tanggal_Jadwal] || 0) + 1; });
  var h = '<div class="fw-bold text-center mb-2">' + BLN[bn] + ' ' + th + '</div><div class="kal">' + ['Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab', 'Min'].map(function (x) { return '<div class="kal-h">' + x + '</div>'; }).join('');
  for (var i = 0; i < mulaiKol; i++) h += '<div></div>';
  for (var d = 1; d <= hari; d++) {
    var tg = bl + '-' + ('0' + d).slice(-2);
    h += '<button type="button" class="kal-d' + (tg === hariIniLokal() ? ' hari' : '') + (tg === PM.tgl ? ' pilih' : '') + '" data-act="pmTgl" data-t="' + tg + '" aria-label="' + d + ' ' + BLN[bn] + (ada[tg] ? ', ' + ada[tg] + ' PM' : '') + '">' + d + (ada[tg] ? '<i class="dot-pm"></i>' : '') + '</button>';
  }
  return h + '</div><div class="small text-muted mt-2"><i class="dot-pm d-inline-block me-1"></i>Ada jadwal PM</div>';
}
ACTIONS.pmBulan = function (el) { var th = Number(PM.bulan.slice(0, 4)), bn = Number(PM.bulan.slice(5, 7)) - 1 + Number(el.dataset.d), t = new Date(th, bn, 1); PM.bulan = t.getFullYear() + '-' + ('0' + (t.getMonth() + 1)).slice(-2); PM.tgl = ''; isiPm(); };
ACTIONS.pmTgl = function (el) { PM.tgl = PM.tgl === el.dataset.t ? '' : el.dataset.t; isiPm(); };
function pmHariHtml() {
  var tg = PM.tgl || hariIniLokal(), l = PM.list.filter(function (p) { return p.Tanggal_Jadwal === tg; });
  $('#pmHariJud').innerHTML = '<i class="bi bi-list-check"></i>PM ' + (tg === hariIniLokal() && !PM.tgl ? 'hari ini' : esc(fmtTgl(tg, true))) + ' <span class="badge-x">' + l.length + '</span>';
  $('#pmHari').innerHTML = l.length ? l.map(function (p) {
    return '<button type="button" class="riw-card" data-act="pmDetail" data-id="' + esc(p.ID_PM) + '"><div class="d-flex flex-wrap gap-1 align-items-center"><b>' + esc(p.ID_PM) + '</b>' + bStatusWo(p.Status_PM) + '</div><div class="riw-uraian">' + esc(p.Nama_Aset) + ' — ' + esc(p.Jenis_PM) + '</div><div class="riw-meta">' + esc(lokasiStr(p)) + '<i class="bi bi-chevron-right ms-auto"></i></div></button>';
  }).join('') : '<div class="small text-muted">Tidak ada jadwal PM pada tanggal ini.</div>';
}
function isiPm() {
  var f = PM.f, data = pmTersaring(), total = data.length, hal = Math.max(1, Math.ceil(total / f.per));
  if (f.hal > hal) f.hal = hal;
  var mulai = (f.hal - 1) * f.per, tmpl = data.slice(mulai, mulai + f.per), kelola = OPS.indexOf(S.user.peran) >= 0;
  var th = hariIniLokal().slice(0, 4), tahun = PM.list.filter(function (p) { return String(p.Tanggal_Jadwal).slice(0, 4) === th; });
  var n = function (s) { return tahun.filter(function (p) { return p.Status_PM === s; }).length; };
  var efektif = tahun.length - n('Dibatalkan'), krit = S.aset.filter(function (a) { return a.Status_Aset !== DIHAPUS && a.Kritikalitas === 'Tinggi'; }).length;
  $('#pmStat').innerHTML = kpiHtml('bi-calendar-check', '', fmtAngka(tahun.length), 'Total PM (' + th + ')') + kpiHtml('bi-check2-circle', 'ok', fmtAngka(n('Terlaksana')), 'Terlaksana') +
    kpiHtml('bi-tools', 'warn', fmtAngka(n('Sedang Dikerjakan')), 'Sedang Dikerjakan') + kpiHtml('bi-exclamation-triangle', 'bad', fmtAngka(krit), 'Total Aset Kritis') +
    kpiHtml('bi-percent', 'ok', efektif ? Math.round(n('Terlaksana') / efektif * 100) + '%' : '—', 'PM Terselesaikan');
  $('#pmKal').innerHTML = pmKalender(); pmHariHtml();
  $('#pmBody').innerHTML = tmpl.length ? tmpl.map(function (p) {
    return '<tr><td data-l="PM"><div class="nm">' + esc(p.ID_PM) + '</div><div class="sub">' + esc(p.Jenis_PM) + (p.Sumber_PM === 'Lapor Langsung Teknisi' ? ' · lapor langsung' : '') + '</div></td>' +
      '<td data-l="Jadwal"><div class="nm">' + esc(fmtTgl(p.Tanggal_Jadwal, true)) + '</div><div class="sub">Due ' + esc(fmtTgl(p.Due_Date, true)) + '</div></td>' +
      '<td data-l="Aset"><div class="nm">' + esc(p.Nama_Aset) + '</div><div class="sub">' + esc(p.ID_Aset) + ' · ' + esc(lokasiStr(p) || '—') + '</div></td>' +
      '<td data-l="Pelaksana"><div>' + esc(p.Pelaksana) + '</div><div class="sub">' + esc(p.Nama_Teknisi_atau_Vendor) + '</div></td>' +
      '<td data-l="Status">' + bStatusWo(p.Status_PM) + '</td><td data-l="Progres">' + bProgresPm(p) + '</td>' +
      '<td class="acts"><button class="icon-btn" data-act="pmDetail" data-id="' + esc(p.ID_PM) + '" title="Detail &amp; progres" aria-label="Detail ' + esc(p.ID_PM) + '"><i class="bi bi-eye"></i></button>' +
      (kelola && (p.Status_PM === 'Terjadwal' || p.Status_PM === 'Ditunda') ? '<button class="icon-btn" data-act="pmUbah" data-id="' + esc(p.ID_PM) + '" title="Ubah" aria-label="Ubah ' + esc(p.ID_PM) + '"><i class="bi bi-pencil"></i></button>' : '') +
      (kelola && p.Status_PM === 'Terjadwal' && !p.Waktu_Mulai ? '<button class="icon-btn text-danger" data-act="pmHapus" data-id="' + esc(p.ID_PM) + '" title="Hapus" aria-label="Hapus ' + esc(p.ID_PM) + '"><i class="bi bi-trash"></i></button>' : '') + '</td></tr>';
  }).join('') : '<tr><td colspan="7"><div class="empty"><i class="bi bi-calendar-x"></i>' + (PM.list.length ? 'Tidak ada PM yang cocok dengan filter.' : 'Belum ada jadwal PM.') + '</div></td></tr>';
  $('#pmPager').innerHTML = '<div>' + (total ? (mulai + 1) + '–' + (mulai + tmpl.length) + ' dari ' + fmtAngka(total) + ' PM' : '0 PM') + '</div>' +
    '<div class="d-flex align-items-center gap-2"><button class="btn btn-light-x btn-sm" data-act="pmHal" data-d="-1"' + (f.hal <= 1 ? ' disabled' : '') + '><i class="bi bi-chevron-left"></i></button><span>Hal. ' + f.hal + ' / ' + hal + '</span><button class="btn btn-light-x btn-sm" data-act="pmHal" data-d="1"' + (f.hal >= hal ? ' disabled' : '') + '><i class="bi bi-chevron-right"></i></button></div>';
}
ACTIONS.pmHal = function (el) { PM.f.hal += Number(el.dataset.d); isiPm(); };
document.addEventListener('input', debounce(function (e) { var t = e.target; if (S.page === 'pm' && t.dataset && t.dataset.pf === 'q') { PM.f.q = t.value; PM.f.hal = 1; isiPm(); } }, 200));
document.addEventListener('change', function (e) { var t = e.target; if (S.page === 'pm' && t.dataset && t.dataset.pf && t.dataset.pf !== 'q') { PM.f[t.dataset.pf] = t.value; PM.f.hal = 1; isiPm(); } });

/* Grafik */
ACTIONS.pmGrafik = function () { PM.view = 'grafik'; hancurkanChart(); renderPmGrafik(); window.scrollTo(0, 0); };
ACTIONS.kembaliPm = function () { PM.view = 'tabel'; renderPm(); };
function renderPmGrafik() {
  hancurkanChart();
  var data = pmTersaring();
  var kotak = function (id, judul, ket) { return '<div class="card-x"><div class="card-h">' + judul + (ket ? '<span class="badge-x ms-auto">' + ket + '</span>' : '') + '</div><div class="card-b"><div class="chart-box"><canvas id="' + id + '" role="img" aria-label="' + esc(judul) + '"></canvas></div></div></div>'; };
  $('#main').innerHTML = '<div class="page-head"><button class="btn btn-light-x" data-act="kembaliPm"><i class="bi bi-arrow-left me-1"></i>Kembali ke Preventive Maintenance</button><h2>Grafik Preventive Maintenance</h2></div>' +
    '<div class="alert-x info small"><i class="bi bi-funnel me-1"></i>Grafik mengikuti filter di halaman PM (' + fmtAngka(data.length) + ' PM).</div>' +
    (data.length ? '<div class="grid-2">' + kotak('gpBulan', 'Realisasi PM per Bulan', 'Jadwal vs Terlaksana') + kotak('gpJenis', 'Jenis Pekerjaan PM') + '</div>' + kotak('gpKat', 'PM Compliance per Kategori', '% Terlaksana') : '<div class="card-x"><div class="empty"><i class="bi bi-bar-chart-line"></i>Belum ada jadwal PM untuk digrafikkan.</div></div>');
  if (!data.length || typeof Chart === 'undefined') return;
  var gelap = S.theme === 'dark', grid = gelap ? '#22304D' : '#E2E8F0';
  Chart.defaults.color = gelap ? '#98A5BF' : '#6B7280'; Chart.defaults.font.family = "'Plus Jakarta Sans',sans-serif";
  var bulan = uniq(data.map(function (p) { return String(p.Tanggal_Jadwal).slice(0, 7); })).sort();
  S.charts.push(new Chart($('#gpBulan'), { type: 'bar', data: { labels: bulan.map(function (b) { return BLN[Number(b.slice(5)) - 1].slice(0, 3) + ' ' + b.slice(2, 4); }), datasets: [
    { label: 'Dijadwalkan', backgroundColor: '#94A3B8', borderRadius: 4, data: bulan.map(function (b) { return data.filter(function (p) { return p.Tanggal_Jadwal.slice(0, 7) === b && p.Status_PM !== 'Dibatalkan'; }).length; }) },
    { label: 'Terlaksana', backgroundColor: '#16A34A', borderRadius: 4, data: bulan.map(function (b) { return data.filter(function (p) { return p.Tanggal_Jadwal.slice(0, 7) === b && p.Status_PM === 'Terlaksana'; }).length; }) }] },
    options: { maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, scales: { y: { ticks: { precision: 0 }, grid: { color: grid } }, x: { grid: { display: false } } } } }));
  var jn = hitung(data, 'Jenis_PM');
  S.charts.push(new Chart($('#gpJenis'), { type: 'doughnut', data: { labels: jn.map(function (x) { return x[0]; }), datasets: [{ data: jn.map(function (x) { return x[1]; }), backgroundColor: jn.map(function (x, i) { return PALET[i % 8]; }), borderWidth: 0 }] }, options: { maintainAspectRatio: false, cutout: '62%', plugins: { legend: { position: 'bottom' } } } }));
  var kat = hitung(data, 'Kategori_Aset').map(function (x) { return x[0]; });
  var pers = kat.map(function (k) { var s = data.filter(function (p) { return p.Kategori_Aset === k && p.Status_PM !== 'Dibatalkan'; }); return s.length ? Math.round(s.filter(function (p) { return p.Status_PM === 'Terlaksana'; }).length / s.length * 100) : 0; });
  S.charts.push(new Chart($('#gpKat'), { type: 'bar', data: { labels: kat, datasets: [{ label: '% Terlaksana', data: pers, backgroundColor: '#1E5EFF', borderRadius: 4 }] }, options: { indexAxis: 'y', maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { min: 0, max: 100, grid: { color: grid } }, y: { grid: { display: false } } } } }));
}

/* Ekspor Excel */
ACTIONS.pmEkspor = async function (el) {
  var data = pmTersaring(); if (!data.length) { toast('Tidak ada data untuk diekspor.', 'warn'); return; }
  var kolom = ['ID_PM', 'ID_Aset', 'Nama_Aset', 'Kategori_Aset', 'Gedung', 'Lantai', 'Ruangan', 'Lokasi_Detail', 'Tanggal_Jadwal', 'Due_Date', 'Jenis_PM', 'Pelaksana', 'Nama_Teknisi_atau_Vendor', 'Aktivitas_Checklist', 'Prioritas_Perawatan', 'Frekuensi_PM', 'Status_PM', 'Waktu_Mulai', 'Waktu_Selesai', 'Hasil_Pemeriksaan', 'Ada_Temuan', 'Temuan', 'Tindakan', 'Verifikator'];
  var tabel = [kolom].concat(data.map(function (p) { return kolom.map(function (k) { return amanSel(p[k]); }); }));
  sibuk(el, true, 'Menyiapkan…');
  var nama = 'Preventive_Maintenance_' + new Date().toISOString().slice(0, 10);
  try {
    await muatSkrip(CDN_XLSX, function () { return typeof XLSX !== 'undefined'; });
    var ws = XLSX.utils.aoa_to_sheet(tabel), wb = XLSX.utils.book_new();
    ws['!cols'] = kolom.map(function (k) { return { wch: Math.min(40, Math.max(k.length + 2, 14)) }; });
    XLSX.utils.book_append_sheet(wb, ws, 'PM');
    unduhBlob(nama + '.xlsx', new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    toast('Berkas Excel diunduh (' + data.length + ' PM).', 'ok');
  } catch (e) {
    unduhBlob(nama + '.csv', new Blob([csvKeString(tabel)], { type: 'text/csv;charset=utf-8' }));
    toast('Pustaka Excel tidak tersedia — diunduh sebagai CSV (dapat dibuka di Excel).', 'warn', 6000);
  } finally { sibuk(el, false); }
};

/* Detail + progres */
ACTIONS.pmDetail = async function (el) {
  var id = el.dataset.id;
  bukaModal({ judul: 'Preventive Maintenance ' + id, size: 'lg', isi: '<div class="skeleton" style="height:240px"></div>', footer: '<button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>' });
  try {
    var d = await call('getPmDetail', id), p = d.pm, kelola = OPS.indexOf(S.user.peran) >= 0;
    $('#mdlContent .modal-body').innerHTML = detailPmHtml(d);
    $('#mdlContent .modal-footer').innerHTML =
      (kelola && p.Status_PM === 'Terjadwal' && !p.Waktu_Mulai ? '<button class="btn btn-light-x text-danger me-auto" data-act="pmHapus" data-id="' + esc(id) + '"><i class="bi bi-trash me-1"></i>Hapus</button>' : '') +
      (kelola && (p.Status_PM === 'Terjadwal' || p.Status_PM === 'Ditunda') ? '<button class="btn btn-light-x" data-act="pmUbah" data-id="' + esc(id) + '"><i class="bi bi-pencil me-1"></i>Ubah</button>' : '') +
      (kelola ? '<button class="btn btn-light-x" data-act="pmPdf" data-id="' + esc(id) + '"><i class="bi bi-file-earmark-pdf me-1"></i>Cetak PDF</button>' : '') +
      '<button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>';
  } catch (e) { tutupModal(); gagal(e); }
};
ACTIONS.riwayatPm = async function (el) {
  var id = el.dataset.id, idAset = el.dataset.aset;
  bukaModal({ judul: 'Preventive Maintenance ' + id, size: 'lg', isi: '<div class="skeleton" style="height:240px"></div>', footer: '<button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>' });
  try {
    var d = await call('getPmDetail', id);
    $('#mdlContent .modal-body').innerHTML = detailPmHtml(d);
    $('#mdlContent .modal-footer').innerHTML = '<button class="btn btn-light-x me-auto" data-act="lihatAset" data-id="' + esc(idAset) + '"><i class="bi bi-arrow-left me-1"></i>Kembali ke aset</button><button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>';
  } catch (e) { tutupModal(); gagal(e); }
};
ACTIONS.pmHapus = async function (el) {
  var id = el.dataset.id;
  if (!(await konfirm('Hapus jadwal PM?', id + ' belum dikerjakan dan akan dihapus permanen.', 'Hapus'))) return;
  try {
    await call('deletePm', id); PM.list = PM.list.filter(function (p) { return p.ID_PM !== id; });
    tutupModal(); toast('Jadwal PM dihapus.', 'ok'); if (S.page === 'pm') isiPm();
  } catch (e) { gagal(e); }
};

/* Formulir jadwal PM */
var PMF = null;
function formPmHtml(p) {
  var tahu = !!(PMF && PMF.mode === 'ubah');
  return '<form id="fPm" data-submit="simpanPm" novalidate><div class="row g-3">' +
    '<div class="col-12"><div class="fw-bold small text-muted text-uppercase">Aset</div>' + apHtml() + '<div class="form-text" id="pmFrek"></div></div>' +
    '<div class="col-md-6"><label class="form-label req">Tanggal Jadwal</label><input type="date" class="form-control" name="Tanggal_Jadwal" value="' + esc((p && p.Tanggal_Jadwal) || hariIniLokal()) + '" required></div>' +
    '<div class="col-md-6"><label class="form-label req">Due Date</label><input type="date" class="form-control" name="Due_Date" value="' + esc((p && p.Due_Date) || '') + '" required></div>' +
    '<div class="col-md-6"><label class="form-label req">Jenis PM</label><select class="form-select" name="Jenis_PM" required>' + opsi(listMaster('Jenis PM'), p && p.Jenis_PM, '— pilih —') + '</select></div>' +
    '<div class="col-md-6"><label class="form-label req">Pelaksana</label><select class="form-select" name="Pelaksana" required>' + opsi(listMaster('Pelaksana'), p && p.Pelaksana, '— pilih —') + '</select></div>' +
    vendorBoxHtml(p) +
    '<div class="col-12"><label class="form-label req">Aktivitas / Checklist PM</label><textarea class="form-control" name="Aktivitas_Checklist" rows="4" maxlength="2000" placeholder="Satu aktivitas per baris, mis.&#10;Cek tekanan oli&#10;Bersihkan filter">' + esc(p && p.Aktivitas_Checklist) + '</textarea></div>' +
    '<div class="col-md-6"><label class="form-label req">Prioritas Perawatan</label><select class="form-select" name="Prioritas_Perawatan" required>' + opsi(listMaster('Prioritas Perawatan'), p && p.Prioritas_Perawatan, '— pilih —') + '</select></div>' +
    '<div class="col-12"><div class="alert-x danger d-none mb-0" id="pmErr"></div></div></div></form>';
}
function bukaFormPm(p, ubah) {
  PMF = { mode: ubah ? 'ubah' : 'baru', id: p && p.ID_PM };
  apBaru(p, ''); AP.tdSaja = true; if (!(AP.aset)) { AP.mode = 'pilih'; AP.belum = false; }
  bukaModal({ judul: ubah ? 'Ubah ' + p.ID_PM : 'Jadwal PM Baru', size: 'lg', isi: formPmHtml(p),
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fPm">' + (ubah ? 'Simpan perubahan' : 'Simpan jadwal') + '</button>' });
}
ACTIONS.pmBaru = function () { bukaFormPm(null, false); };
ACTIONS.pmUbah = async function (el) { try { var d = await call('getPmDetail', el.dataset.id); bukaFormPm(d.pm, true); } catch (e) { gagal(e); } };
document.addEventListener('change', function (e) {
  var t = e.target; if (!t.closest || !t.closest('#fPm')) return;
  if (t.name === 'Pelaksana') $('#boxVendor').classList.toggle('d-none', t.value !== 'Vendor');
  if (t.name === 'VendorPilih') { var ni = $('#fPm [name=Nama_Vendor]'); ni.classList.toggle('d-none', t.value !== '__lain__'); if (t.value === '__lain__') ni.focus(); }
});
SUBMITS.simpanPm = async function (f) {
  var err = $('#pmErr'); err.classList.add('d-none');
  var salah = function (m) { err.textContent = m; err.classList.remove('d-none'); };
  var p = {};
  try { p = apNilai(); } catch (e) { salah(e.message); return; }
  if (!p.ID_Aset) { salah('Pilih aset yang sudah terdaftar di Data Aset.'); return; }
  ['Tanggal_Jadwal', 'Due_Date', 'Jenis_PM', 'Pelaksana', 'Aktivitas_Checklist', 'Prioritas_Perawatan'].forEach(function (k) { p[k] = f[k].value; });
  if (p.Pelaksana === 'Vendor') {
    var vp = f.VendorPilih.value, vn = vp === '__lain__' ? f.Nama_Vendor.value.trim() : vp;
    if (!vn) { salah('Nama vendor wajib dipilih atau diketik.'); return; }
    p.Nama_Vendor = vn;
  }
  if (PMF.mode === 'ubah') p.ID_PM = PMF.id;
  if (!p.Tanggal_Jadwal || !p.Due_Date || !p.Jenis_PM || !p.Pelaksana || !p.Aktivitas_Checklist.trim() || !p.Prioritas_Perawatan) { salah('Lengkapi semua kolom bertanda *.'); return; }
  if (p.Due_Date < p.Tanggal_Jadwal) { salah('Due Date tidak boleh sebelum Tanggal Jadwal.'); return; }
  var btn = $('button[type=submit][form=fPm]'); sibuk(btn, true, 'Menyimpan…');
  try {
    var d = await call('savePm', p, PMF.mode);
    gantiPmLokal(d.pm); tutupModal(); toast(PMF.mode === 'ubah' ? 'Jadwal PM diperbarui.' : 'Jadwal PM ' + d.pm.ID_PM + ' dibuat.', 'ok');
    if (S.page === 'pm') isiPm();
  } catch (e) { salah(e.message); } finally { sibuk(btn, false); }
};

/* Impor CSV */
var PM_KOLOM_CSV = ['ID_Aset', 'Tanggal_Jadwal', 'Due_Date', 'Jenis_PM', 'Pelaksana', 'Nama_Vendor', 'Aktivitas_Checklist', 'Prioritas_Perawatan'];
ACTIONS.pmImpor = function () {
  bukaModal({ judul: 'Impor Jadwal PM (CSV)', size: 'lg',
    isi: '<div class="alert-x info small">Kolom: <code>' + PM_KOLOM_CSV.join(', ') + '</code>. <b>ID_Aset</b> harus sudah terdaftar. Tanggal format <code>YYYY-MM-DD</code>. <code>Nama_Vendor</code> hanya diisi bila Pelaksana = Vendor. Baris yang salah dilaporkan, baris benar tetap diimpor. Maksimal 2.000 baris.</div>' +
      '<button class="btn btn-light-x btn-sm mb-3" data-act="pmTemplate"><i class="bi bi-download me-1"></i>Unduh template</button><input type="file" class="form-control" id="pmCsv" accept=".csv,text/csv"><div id="pmHasil" class="mt-3"></div>',
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Tutup</button>' });
};
ACTIONS.pmTemplate = function () { unduhBlob('template_jadwal_pm.csv', new Blob([csvKeString([PM_KOLOM_CSV, [(S.aset[0] && S.aset[0].ID_Aset) || 'AST-0001', hariIniLokal(), hariIniLokal(), listMaster('Jenis PM')[0] || 'Inspeksi', 'Internal', '', 'Cek kondisi umum', listMaster('Prioritas Perawatan')[0] || 'Tinggi']])], { type: 'text/csv;charset=utf-8' })); };
document.addEventListener('change', async function (e) {
  if (e.target.id !== 'pmCsv' || !e.target.files[0]) return;
  var box = $('#pmHasil'); box.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Mengimpor…';
  try {
    var t = parseCsv(await bacaFile(e.target.files[0], true)); if (t.length < 2) throw new Error('File kosong.'); if (t.length > 2001) throw new Error('Maksimal 2.000 baris.');
    var map = {}; t[0].forEach(function (h, i) { map[String(h).trim().toLowerCase().replace(/\s+/g, '_')] = i; });
    ['id_aset', 'tanggal_jadwal', 'due_date', 'jenis_pm', 'pelaksana', 'aktivitas_checklist', 'prioritas_perawatan'].forEach(function (k) { if (map[k] === undefined) throw new Error('Kolom ' + k + ' tidak ditemukan.'); });
    var g = function (r, k) { return map[k] === undefined ? '' : String(r[map[k]] || '').trim(); };
    var rows = t.slice(1).map(function (r, i) { var o = { _no: i + 2 }; PM_KOLOM_CSV.forEach(function (k) { o[k] = g(r, k.toLowerCase()); }); return o; });
    var tot = { berhasil: 0, gagal: [] };
    for (var i = 0; i < rows.length; i += 200) { var d = await call('importPm', rows.slice(i, i + 200)); tot.berhasil += d.berhasil; tot.gagal = tot.gagal.concat(d.gagal); }
    await muatPm(); if (S.page === 'pm') isiPm();
    box.innerHTML = '<div class="alert-x ok">' + tot.berhasil + ' jadwal PM ditambahkan.</div>' + (tot.gagal.length ? '<div class="alert-x danger"><b>' + tot.gagal.length + ' baris gagal:</b><br>' + tot.gagal.slice(0, 20).map(function (x) { return 'Baris ' + x.no + ': ' + esc(x.errors.join(' ')); }).join('<br>') + '</div>' : '');
  } catch (er) { box.innerHTML = '<div class="alert-x danger">' + esc(er.message) + '</div>'; }
});

/* ════════ Kartu: Preventive Maintenance (teknisi) ════════ */
function kpDraftSimpan() {
  var k = S.kp; if (!k.id || (!k.draftLapor && !k.draftMulai)) return;
  var o = muatDraft(k.id) || {};
  o.ts = Date.now();
  if (k.draftLapor) o.lapor = { form: k.form, foto: k.foto, docs: k.docs, sp: k.sp };
  if (k.draftMulai) o.mulai = k.mulai;
  try { LS.set(dkey(k.id), JSON.stringify(o)); } catch (e) {}
}
function kpCard(p, dari) {
  return '<button class="wo-card" data-act="kpBuka" data-id="' + esc(p.ID_PM) + '" data-dari="' + (dari || 'list') + '">' +
    '<div class="d-flex flex-wrap gap-1 align-items-center mb-1"><b>' + esc(p.ID_PM) + '</b>' + bStatusWo(p.Status_PM) + '<span class="badge-x">' + esc(p.Pelaksana) + '</span>' + (muatDraft(p.ID_PM) ? badge('Draft laporan', 'warn') : '') + '</div>' +
    '<div class="fw-bold">' + esc(p.Nama_Aset) + '</div><div class="small text-muted">' + esc(p.Jenis_PM) + '</div>' +
    '<div class="small mt-1"><i class="bi bi-geo-alt me-1"></i>' + esc(lokasiStr(p) || '—') + '</div>' +
    '<div class="small"><i class="bi bi-flag me-1"></i>Prioritas perawatan: ' + esc(p.Prioritas_Perawatan) + '</div>' +
    '<div class="small"><i class="bi bi-calendar-event me-1"></i>Due Date ' + esc(fmtTgl(p.Due_Date, true)) + '</div><div class="mt-2">' + bProgresPm(p) + '</div></button>';
}
async function kPmList(m) {
  if (PM.dimuat) { kPmGambar(m); muatKartu(false); return; }
  m.innerHTML = '<div class="skeleton" style="height:48px"></div><div class="skeleton mt-3" style="height:180px"></div>';
  try { await muatKartu(true); if (!PM.dimuat) await muatPm(); } catch (e) {}
  if (!PM.dimuat) {
    if (S.kPage === 'pm') m.innerHTML = '<div class="card-x k-card text-center"><i class="bi bi-exclamation-circle fs-1 text-danger"></i><p class="mt-2">Gagal memuat Preventive Maintenance.</p><button class="btn btn-primary k-btn" data-act="kNav" data-p="pm">Coba lagi</button></div>';
    return;
  }
  if (S.kPage !== 'pm') return;
  kPmGambar(m);
}
function kPmPerbarui() {
  $$('.k-tabs button[data-t]').forEach(function (b) { var t = b.dataset.t; if (KP_TAB[t]) b.textContent = ({ terjadwal: 'Terjadwal', ditunda: 'Ditunda', selesai: 'Selesai' })[t] + ' (' + PM.list.filter(function (p) { return KP_TAB[t].indexOf(p.Status_PM) >= 0; }).length + ')'; });
  kpIsi();
}
function kPmGambar(m) {
  var k = S.kp, jml = function (t) { return PM.list.filter(function (p) { return KP_TAB[t].indexOf(p.Status_PM) >= 0; }).length; };
  m.innerHTML = '<div class="d-flex align-items-center mb-3"><div class="h5 fw-bold mb-0">Preventive Maintenance</div>' + (adalahTeknisi() ? '<button class="btn btn-light-x btn-sm ms-auto" data-act="kpLaporBaru"><i class="bi bi-plus-lg me-1"></i>Lapor PM aset belum terdaftar</button>' : '') + '</div>' +
    '<div class="k-tabs">' + [['terjadwal', 'Terjadwal'], ['ditunda', 'Ditunda'], ['selesai', 'Selesai']].map(function (t) { return '<button data-act="kpTab" data-t="' + t[0] + '" class="' + (k.tab === t[0] ? 'active' : '') + '">' + t[1] + ' (' + jml(t[0]) + ')</button>'; }).join('') + '</div>' +
    '<div class="input-icon mb-2"><i class="bi bi-search"></i><input class="form-control" id="kpQ" placeholder="Cari ID PM, ID aset, nama aset…" value="' + esc(k.q) + '" style="min-height:48px" aria-label="Cari PM"></div>' +
    '<div class="row g-2 mb-3"><div class="col-6"><input type="date" class="form-control" data-kpfl="tgl" value="' + esc(k.tgl) + '" aria-label="Tanggal PM"></div>' +
    '<div class="col-6"><select class="form-select" data-kpfl="kat" aria-label="Kategori Aset">' + opsi(listMaster('Kategori Aset'), k.kat, 'Semua kategori') + '</select></div>' +
    '<div class="col-6"><select class="form-select" data-kpfl="lok" aria-label="Lokasi">' + opsi(uniq(PM.list.map(function (p) { return p.Ruangan; })).sort(), k.lok, 'Semua lokasi') + '</select></div>' +
    '<div class="col-6"><select class="form-select" data-kpfl="pel" aria-label="Pelaksana">' + opsi(listMaster('Pelaksana'), k.pel, 'Semua pelaksana') + '</select></div></div>' +
    '<div id="kpList"></div><div id="kpPager"></div>';
  kpIsi();
}
function kpIsi() {
  var k = S.kp, q = k.q.trim(), el = $('#kpList'); if (!el) return;
  var data = PM.list.filter(function (p) {
    if (KP_TAB[k.tab].indexOf(p.Status_PM) < 0) return false;
    if (k.tgl && p.Tanggal_Jadwal !== k.tgl) return false;
    if (k.kat && p.Kategori_Aset !== k.kat) return false;
    if (k.lok && p.Ruangan !== k.lok) return false;
    if (k.pel && p.Pelaksana !== k.pel) return false;
    return true;
  });
  if (q) data = cariUrut(data, q, function (p) { return p.ID_PM + ' ' + p.ID_Aset + ' ' + p.Nama_Aset + ' ' + p.Kategori_Aset + ' ' + p.Jenis_PM + ' ' + p.Lantai + ' ' + p.Ruangan + ' ' + p.Gedung; });
  else if (k.tab === 'selesai') data.sort(function (a, b) { return a.Waktu_Selesai < b.Waktu_Selesai ? 1 : -1; });
  else data.sort(function (a, b) { return a.Due_Date < b.Due_Date ? -1 : (a.Due_Date > b.Due_Date ? 1 : 0); });
  var per = 5, hal = Math.max(1, Math.ceil(data.length / per)); if (k.hal > hal) k.hal = hal;
  var mulai = (k.hal - 1) * per;
  el.innerHTML = data.length ? data.slice(mulai, mulai + per).map(function (p) { return kpCard(p); }).join('') : '<div class="empty"><i class="bi bi-calendar-check"></i>Tidak ada PM di tab ini.</div>';
  $('#kpPager').innerHTML = data.length > per ? '<div class="d-flex align-items-center justify-content-between mt-2"><button class="btn btn-light-x" data-act="kpHal" data-d="-1"' + (k.hal <= 1 ? ' disabled' : '') + '><i class="bi bi-chevron-left"></i></button><span class="small text-muted">Hal. ' + k.hal + ' / ' + hal + ' · ' + data.length + ' PM</span><button class="btn btn-light-x" data-act="kpHal" data-d="1"' + (k.hal >= hal ? ' disabled' : '') + '><i class="bi bi-chevron-right"></i></button></div>' : '';
}
ACTIONS.kpTab = function (el) { S.kp.tab = el.dataset.t; S.kp.hal = 1; $$('.k-tabs button').forEach(function (b) { b.classList.toggle('active', b.dataset.t === S.kp.tab); }); kpIsi(); };
ACTIONS.kpHal = function (el) { S.kp.hal += Number(el.dataset.d); kpIsi(); window.scrollTo(0, 0); };
document.addEventListener('input', debounce(function (e) { if (e.target.id === 'kpQ') { S.kp.q = e.target.value; S.kp.hal = 1; kpIsi(); } }, 150));
document.addEventListener('change', function (e) { var t = e.target; if (t.dataset && t.dataset.kpfl && S.kPage === 'pm') { S.kp[t.dataset.kpfl] = t.value; S.kp.hal = 1; kpIsi(); } });

ACTIONS.kpBuka = function (el) { S.kp.id = el.dataset.id; S.kp.dari = el.dataset.dari || 'list'; kpMuatDetail('info'); };
async function kpMuatDetail(langkah) {
  S.kPage = 'pmdetail'; S.kp.d = null; S.kp.err = ''; renderKartu();
  try {
    var pr = [call('getPmDetail', S.kp.id)];
    if (adalahTeknisi() && !WO.opsi) pr.push(call('getWoOpsi'));
    var hs = await Promise.all(pr); S.kp.d = hs[0];
    if (hs[1]) { WO.opsi = hs[1]; WO.verif = WO.opsi.verifikator; }
  } catch (e) { S.kp.err = e.message; }
  S.kp.langkah = langkah || 'info'; S.kp.draftLapor = false; S.kp.draftMulai = false;
  if (S.kp.d && langkah === 'lapor') kpInitLapor(S.kp.d.pm);
  if (S.kp.d && langkah === 'mulai') kpInitMulai(S.kp.d.pm);
  if (S.kPage === 'pmdetail') kPmDetail($('#kmain'));
}
ACTIONS.kpKembali = function () { if (S.kp.dari === 'aset' && S.kAset && S.kAset.aset) { S.kPage = 'aset'; renderKartu(); } else { S.kPage = 'pm'; renderKartu(); } };
ACTIONS.kpInfo = function () { S.kp.langkah = 'info'; kPmDetail($('#kmain')); };
ACTIONS.kpLangkahMulai = function () { S.kp.langkah = 'mulai'; kpInitMulai(S.kp.d.pm); kPmDetail($('#kmain')); window.scrollTo(0, 0); };
ACTIONS.kpLangkahLapor = function () { if (!kpBolehLapor(S.kp.d.pm)) { toast('Laporan hanya dapat diisi oleh teknisi yang menekan Mulai.', 'warn'); return; } S.kp.langkah = 'lapor'; kpInitLapor(S.kp.d.pm); kPmDetail($('#kmain')); window.scrollTo(0, 0); };
function kpBolehLapor(p) { return adalahTeknisi() && (!p.Dimulai_Oleh || p.Dimulai_Oleh === S.user.username); }
function kpRingkas(p) {
  return '<div class="card-x k-card"><div class="d-flex flex-wrap gap-1 align-items-center mb-1"><b>' + esc(p.ID_PM) + '</b>' + bStatusWo(p.Status_PM) + bPrioPm(p.Prioritas_Perawatan) + '</div><div class="fw-bold">' + esc(p.Nama_Aset) + '</div>' +
    '<div class="small text-muted">' + esc(lokasiLengkap(p)) + '</div><div class="small mt-1"><b>' + esc(p.Jenis_PM) + '</b> · Due ' + esc(fmtTgl(p.Due_Date, true)) + '</div><div class="small mt-1">' + teksPanjang(p.Aktivitas_Checklist) + '</div><div class="mt-2">' + bProgresPm(p) + ' <button class="btn-link-x" data-act="kpInfo">Lihat detail</button></div></div>';
}
function kPmDetail(m) {
  var k = S.kp, d = k.d;
  if (k.err) { m.innerHTML = '<button class="btn-link-x mb-2" data-act="kpKembali"><i class="bi bi-arrow-left me-1"></i>Kembali</button><div class="card-x k-card text-center"><i class="bi bi-exclamation-circle fs-1 text-danger"></i><p class="mt-2 mb-0">' + esc(k.err) + '</p></div>'; return; }
  if (!d) { var li = PM.list.filter(function (x) { return x.ID_PM === k.id; })[0]; m.innerHTML = '<button class="btn-link-x mb-2" data-act="kpKembali"><i class="bi bi-arrow-left me-1"></i>Kembali</button>' + (li ? kpRingkas(li).replace(/<button class="btn-link-x" data-act="kpInfo">Lihat detail<\/button>/, '') + '<div class="skeleton mt-3" style="height:140px"></div>' : '<div class="skeleton" style="height:60px"></div><div class="skeleton mt-3" style="height:260px"></div>'); return; }
  var p = d.pm, tek = adalahTeknisi(), balik = '<button class="btn-link-x mb-2" data-act="kpKembali"><i class="bi bi-arrow-left me-1"></i>Kembali</button>';
  if (k.langkah === 'mulai' && tek && (p.Status_PM === 'Terjadwal' || p.Status_PM === 'Ditunda')) { m.innerHTML = balik + kpRingkas(p) + kpMulaiHtml(p); return; }
  if (k.langkah === 'lapor' && kpBolehLapor(p) && p.Status_PM === 'Sedang Dikerjakan') { m.innerHTML = balik + kpRingkas(p) + '<div id="kpLaporBox">' + kpLaporIsi(p) + '</div>'; kpCekTombol(); return; }
  var aksi = '';
  if (tek && (p.Status_PM === 'Terjadwal' || p.Status_PM === 'Ditunda')) aksi = '<button class="btn btn-primary k-btn mt-2" data-act="kpLangkahMulai"><i class="bi bi-play-circle me-2"></i>Mulai Pengerjaan</button>';
  else if (tek && p.Status_PM === 'Sedang Dikerjakan' && kpBolehLapor(p)) aksi = (muatDraft(p.ID_PM) ? '<div class="alert-x warn small mt-2"><i class="bi bi-save me-1"></i>Ada draft laporan yang tersimpan di perangkat ini.</div>' : '') + '<button class="btn btn-primary k-btn mt-2" data-act="kpLangkahLapor"><i class="bi bi-clipboard2-check me-2"></i>' + (muatDraft(p.ID_PM) ? 'Lanjutkan Laporan' : 'Lapor Hasil PM') + '</button>';
  else if (tek && p.Status_PM === 'Sedang Dikerjakan') aksi = '<div class="alert-x info small mt-2"><i class="bi bi-lock me-1"></i>PM ini dimulai oleh <b>' + esc(p.Nama_Teknisi_atau_Vendor || p.Dimulai_Oleh) + '</b> (akun <b>' + esc(p.Dimulai_Oleh) + '</b>). Laporan hasil hanya dapat diisi oleh akun yang menekan Mulai.</div>';
  else if (PM_AKTIF.indexOf(p.Status_PM) >= 0) aksi = '<div class="alert-x info small mt-2"><i class="bi bi-info-circle me-1"></i>Hanya Teknisi yang dapat memulai dan melaporkan PM.</div>';
  m.innerHTML = balik + '<div class="h5 fw-bold mb-3">' + esc(p.Nama_Aset) + '</div>' + detailPmHtml(d) + aksi;
}

/* Langkah 1: pilih pelaksana lalu Mulai */
function kpInitMulai(p) {
  var dr = muatDraft(p.ID_PM), m = dr && dr.mulai;
  S.kp.mulai = { pel: p.Pelaksana || 'Internal', tek: (p.Pelaksana !== 'Vendor' && p.Nama_Teknisi_atau_Vendor) ? String(p.Nama_Teknisi_atau_Vendor).split(/\s*,\s*/) : [S.user.nama], vendor: p.Pelaksana === 'Vendor' ? (p.Nama_Teknisi_atau_Vendor || '') : '' };
  if (m) Object.keys(m).forEach(function (k) { S.kp.mulai[k] = m[k]; });
  S.kp.draftMulai = true;
}
function kpSimpanMulai() {
  var f = $('#fMulaiPm'); if (!f || !S.kp.draftMulai) return;
  var m = S.kp.mulai; m.pel = f.pel.value;
  if (m.pel === 'Vendor') { if (f.nama) m.vendor = f.nama.value; } else m.tek = $$('#fMulaiPm [name=tek]:checked').map(function (c) { return c.value; });
  kpDraftSimpan();
}
function kpNamaHtml(pel, m) {
  if (pel === 'Vendor') {
    var dv = uniq(listMaster('Nama Vendor').concat((WO.opsi && WO.opsi.vendor) || []));
    return '<label class="form-label req">Nama Vendor</label><input class="form-control" name="nama" list="dlVendorPm" maxlength="100" value="' + esc(m.vendor) + '" placeholder="Pilih atau ketik nama vendor"><datalist id="dlVendorPm">' + dv.map(function (x) { return '<option value="' + esc(x) + '">'; }).join('') + '</datalist>' +
      '<div class="form-text">' + (m.vendor ? 'Terisi otomatis dari jadwal PM. ' : '') + 'Vendor yang belum terdaftar dikirim ke Master Admin untuk didaftarkan.</div>';
  }
  var daftar = uniq(((WO.opsi && WO.opsi.teknisi) || []).concat(listMaster('Nama Teknisi')));
  if (!daftar.length) return '<label class="form-label req">Nama Teknisi</label><input class="form-control" name="tekBebas" maxlength="100" value="' + esc((m.tek || [])[0] || S.user.nama) + '">';
  return '<label class="form-label req">Nama Teknisi <span class="text-muted fw-normal">(boleh lebih dari satu)</span></label><div class="chk-list">' +
    daftar.map(function (n) { return '<label class="chk-item"><input type="checkbox" name="tek" value="' + esc(n) + '"' + ((m.tek || []).indexOf(n) >= 0 ? ' checked' : '') + '><span>' + esc(n) + '</span></label>'; }).join('') + '</div>';
}
function kpMulaiHtml(p) {
  var m = S.kp.mulai;
  return '<form id="fMulaiPm" class="k-form" data-submit="kpMulai" novalidate><div class="card-x k-card"><div class="fw-bold mb-2">Pilih pelaksana lalu tekan Mulai</div>' +
    (muatDraft(p.ID_PM) && muatDraft(p.ID_PM).mulai ? '<div class="alert-x warn small"><i class="bi bi-save me-1"></i>Isian sebelumnya dipulihkan dari draft.</div>' : '') +
    '<div class="row g-2"><div class="col-12"><label class="form-label req">Pelaksana</label><select class="form-select" name="pel" data-kppel="1">' + opsi(listMaster('Pelaksana'), m.pel, null) + '</select></div>' +
    '<div class="col-12" id="kpNama">' + kpNamaHtml(m.pel, m) + '</div></div>' +
    '<div class="alert-x danger d-none mt-3 mb-0" id="mulaiPmErr"></div></div>' +
    '<button class="btn btn-primary k-btn" type="submit"><i class="bi bi-play-circle me-2"></i>Mulai</button></form>';
}
document.addEventListener('change', function (e) {
  var t = e.target; if (!t.closest || !t.closest('#fMulaiPm')) return;
  if (t.dataset && t.dataset.kppel) { kpSimpanMulai(); S.kp.mulai.pel = t.value; $('#kpNama').innerHTML = kpNamaHtml(t.value, S.kp.mulai); }
  kpSimpanMulai();
});
document.addEventListener('input', function (e) { var t = e.target; if (t.closest && t.closest('#fMulaiPm')) kpSimpanMulai(); });
SUBMITS.kpMulai = async function (f) {
  var p0 = S.kp.d.pm, err = $('#mulaiPmErr'); err.classList.add('d-none');
  var salah = function (m) { err.textContent = m; err.classList.remove('d-none'); };
  var p = { Pelaksana: f.pel.value };
  if (p.Pelaksana === 'Vendor') { p.Nama_Vendor = f.nama.value.trim(); if (!p.Nama_Vendor) { salah('Nama vendor wajib diisi.'); return; } }
  else {
    p.Nama_Teknisi = f.tekBebas ? [f.tekBebas.value.trim()] : $$('#fMulaiPm [name=tek]:checked').map(function (c) { return c.value; });
    if (!p.Nama_Teknisi.length || !p.Nama_Teknisi[0]) { salah('Pilih minimal satu nama teknisi.'); return; }
  }
  var btn = $('#fMulaiPm button[type=submit]'); sibuk(btn, true, 'Memulai…');
  try { await call('mulaiPm', p0.ID_PM, p); var o = muatDraft(p0.ID_PM); if (o) { delete o.mulai; try { if (o.lapor) LS.set(dkey(p0.ID_PM), JSON.stringify(o)); else LS.del(dkey(p0.ID_PM)); } catch (e2) {} } S.kp.draftMulai = false; toast('Pekerjaan dimulai.', 'ok'); await kpMuatDetail('lapor'); }
  catch (e) { salah(e.message); } finally { sibuk(btn, false); }
};

/* Langkah 2: laporan hasil */
function kpAsetTerdaftar(p) { return !!(p.ID_Aset && p.ID_Aset !== 'Belum Tahu' && cariAset(p.ID_Aset)); }
function kpInitLapor(p) {
  var ib = infoAsetBaru(p);
  var f = { hasil: '', alasan: '', desk: '', periksa: p.Hasil_Pemeriksaan || '', temuanYa: p.Ada_Temuan || '', temuan: p.Temuan || '', tindakan: p.Tindakan || '', catatan: '', verif: '', tgl: hariIniLokal(), kondisi: p.Kondisi_Dilaporkan || '', lokDetail: p.Lokasi_Detail || '',
    iThn: ib.Tahun_Perolehan || '', iNup: ib.NUP || '', iMerk: ib.Merk_Type || '', iSeri: ib.Nomor_Seri || '', iStatus: ib.Status_Aset || '' };
  var fo = { b1: p.Foto_Before_1 || '', b2: p.Foto_Before_2 || '', a1: p.Foto_After_1 || '', a2: p.Foto_After_2 || '', ast: p.Foto_Aset_URL || '' };
  var docs = dokumenPm(p), sp = sparepartWo(p).map(function (x) { return { u: x.u, c: x.c }; }), dr = muatDraft(p.ID_PM);
  S.kp.dipulihkan = false;
  if (dr && dr.lapor) {
    Object.keys(dr.lapor.form || {}).forEach(function (k) { if (dr.lapor.form[k] !== undefined && dr.lapor.form[k] !== '') f[k] = dr.lapor.form[k]; });
    Object.keys(dr.lapor.foto || {}).forEach(function (k) { if (dr.lapor.foto[k]) fo[k] = dr.lapor.foto[k]; });
    if (dr.lapor.docs && dr.lapor.docs.length) docs = dr.lapor.docs;
    if (dr.lapor.sp && dr.lapor.sp.length) sp = dr.lapor.sp;
    S.kp.dipulihkan = true;
  }
  S.kp.form = f; S.kp.foto = fo; S.kp.docs = docs; S.kp.sp = sp; S.kp.draftLapor = true;
}
function kpSlot(key, label) { return slotHtml(key, label, S.kp.foto[key]); }
function kpAlasanBoleh() {
  var pernah = S.kp.d.log.some(function (x) { return x.Status_Baru === 'Ditunda'; }), master = listMaster('Alasan Penundaan');
  var dasar = ['Menunggu Sparepart', 'Kendala Akses, Cuaca, atau Jadwal', 'Dialihkan ke Pekerjaan Darurat', 'Menunggu Pihak Ketiga/Vendor'].filter(function (a) { return !master.length || master.indexOf(a) >= 0; });
  return pernah ? dasar.filter(function (a) { return ['Kendala Akses, Cuaca, atau Jadwal', 'Dialihkan ke Pekerjaan Darurat'].indexOf(a) >= 0; }) : dasar;
}
function kpSaranChecklist(p) {
  var baris = String(p.Aktivitas_Checklist || '').split(/\r?\n|;/).map(function (x) { return x.replace(/^[\s\-•*\d.)]+/, '').trim(); }).filter(Boolean);
  return baris.length ? baris.map(function (x) { return '- ' + x + ': '; }).join('\n') : '';
}
function kpSpHtml() {
  var a = S.kp.sp || [];
  return '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Foto sparepart yang diganti</div>' +
    '<div class="form-text mb-2">Minimal <b>' + SP_MIN + '</b> dan maksimal <b>' + SP_MAKS + '</b> foto, masing-masing wajib diberi keterangan. Terisi: <b id="kpSpJml">' + a.length + '</b>/' + SP_MAKS + '.</div>' +
    inputFotoHtml('kpSpIn', 'kpSpCam', a.length >= SP_MAKS) +
    '<div id="kpSpList" class="mt-2">' + a.map(function (x, i) {
      return '<div class="sp-item"><div class="foto-th"><img src="' + esc(imgU(x.u)) + '" alt="Sparepart ' + (i + 1) + '" referrerpolicy="no-referrer" data-fb="1"></div>' +
        '<div class="flex-grow-1"><label class="form-label small mb-1 req">Keterangan foto ' + (i + 1) + '</label><input class="form-control form-control-sm" maxlength="150" data-kpsp="' + i + '" value="' + esc(x.c) + '" placeholder="mis. Kapasitor kompresor lama yang diganti">' +
        '<button type="button" class="btn-link-x text-danger small mt-1" data-act="kpSpHapus" data-i="' + i + '">hapus foto</button></div></div>';
    }).join('') + '</div>';
}
ACTIONS.kpSpHapus = function (el) { S.kp.sp.splice(Number(el.dataset.i), 1); kpGambarLapor(); };
function kpLaporIsi(p) {
  var f = S.kp.form, h = f.hasil, sparepart = h === 'Ditunda' && f.alasan === 'Menunggu Sparepart', terdaftar = kpAsetTerdaftar(p), a0 = terdaftar ? cariAset(p.ID_Aset) : null;
  var seg = '<div class="seg">' + [['Terlaksana', 's-ok'], ['Ditunda', 's-warn'], ['Dibatalkan', 's-bad']].map(function (x) { return '<button type="button" class="' + x[1] + (h === x[0] ? ' on' : '') + '" data-act="kpHasil" data-v="' + x[0] + '">' + x[0] + '</button>'; }).join('') + '</div>';
  var periksa = '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Hasil pemeriksaan</div>' +
    (f.periksa.trim() ? '' : '<div class="mb-1"><button type="button" class="btn-link-x" data-act="kpSaran"><i class="bi bi-magic me-1"></i>Isi saran dari checklist PM</button></div>') +
    '<textarea class="form-control" rows="4" maxlength="2000" data-kpf="periksa">' + esc(f.periksa) + '</textarea>' +
    '<label class="form-label req mt-3">Temuan</label><div class="seg"><button type="button" class="s-warn' + (f.temuanYa === 'Ya' ? ' on' : '') + '" data-act="kpTemuan" data-v="Ya">Ada Temuan</button><button type="button" class="s-ok' + (f.temuanYa === 'Tidak' ? ' on' : '') + '" data-act="kpTemuan" data-v="Tidak">Tidak Ada Temuan</button></div>' +
    (f.temuanYa === 'Ya' ? '<label class="form-label req mt-2">Deskripsi temuan</label><textarea class="form-control" rows="3" maxlength="1000" data-kpf="temuan">' + esc(f.temuan) + '</textarea>' : '');
  var foto = '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Foto sebelum perawatan</div><div class="foto-grid">' + kpSlot('b1', 'Sebelum 1') + kpSlot('b2', 'Sebelum 2') + '</div>' +
    '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Foto sesudah perawatan</div><div class="foto-grid">' + kpSlot('a1', 'Sesudah 1') + kpSlot('a2', 'Sesudah 2') + '</div>';
  var cabang = '';
  if (h === 'Terlaksana') {
    cabang = periksa + (f.temuanYa === 'Ya' ? '<label class="form-label req mt-2">Tindakan yang dilakukan</label><textarea class="form-control" rows="3" maxlength="1000" data-kpf="tindakan">' + esc(f.tindakan) + '</textarea>' : '') + foto + (kpPerluSp() ? kpSpHtml() : '') +
      '<label class="form-label mt-3">Dokumen (opsional, JPG/PNG/PDF)</label><div class="mb-2"><label class="btn btn-light-x btn-sm"><i class="bi bi-camera-fill me-1"></i>Ambil foto dokumen<input type="file" id="kpDokCam" accept="image/*" capture="environment" hidden></label></div><input type="file" class="form-control" id="kpDok" multiple accept="image/*,application/pdf"><div class="form-text">Dokumen PDF maksimal <b>2 MB</b> per berkas. Foto dikompres otomatis.</div><div id="kpDokList" class="mt-2">' + kpDokHtml() + '</div>';
  } else if (h === 'Ditunda') {
    cabang = '<label class="form-label req mt-3">Alasan penundaan</label><select class="form-select" data-kpf="alasan">' + opsi(kpAlasanBoleh(), f.alasan, '— pilih alasan —') + '</select>' +
      '<label class="form-label req mt-2">Deskripsi penundaan</label><textarea class="form-control" rows="3" maxlength="1000" data-kpf="desk" placeholder="Jelaskan alasan penundaan">' + esc(f.desk) + '</textarea>' + (sparepart ? periksa : '');
  }
  var kondisi = '', asetBaru = '';
  if (h && terdaftar) kondisi = '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Kondisi aset</div><label class="form-label' + (h === 'Dibatalkan' ? '' : ' req') + '">Kondisi aset saat ini (menurut Anda)</label><select class="form-select" data-kpf="kondisi">' + opsi(listMaster('Kondisi Aset'), f.kondisi, '— pilih kondisi —') + '</select>' +
    '<div class="form-text">Sebelumnya tercatat: <b>' + esc(a0.Kondisi) + '</b>. Jika memilih <b>Rusak Ringan</b> atau <b>Rusak Berat</b>, Master Admin otomatis mendapat pemberitahuan untuk mengubah kondisi aset.</div>';
  if (h && !terdaftar) asetBaru = '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Aset belum terdaftar</div><div class="alert-x info small mb-2"><i class="bi bi-info-circle me-1"></i>Aset ini belum ada di Data Aset. ' + (h === 'Terlaksana' ? 'Unggah foto aset' : 'Foto aset') + ' dan lengkapi data di bawah; datanya dikirim ke Master Admin untuk didaftarkan.</div>' +
    '<div class="foto-grid">' + kpSlot('ast', 'Foto aset') + '</div><label class="form-label mt-2">Lokasi detail <span class="text-muted fw-normal">(ketik bebas)</span></label><input class="form-control" data-kpf="lokDetail" maxlength="150" value="' + esc(f.lokDetail) + '" placeholder="mis. dekat jendela barat">' +
    '<div class="row g-2 mt-1"><div class="col-6"><label class="form-label">Tahun perolehan <span class="text-muted fw-normal">(jika ada)</span></label><input class="form-control" inputmode="numeric" maxlength="4" data-kpf="iThn" value="' + esc(f.iThn) + '"></div>' +
    '<div class="col-6"><label class="form-label">NUP <span class="text-muted fw-normal">(jika ada)</span></label><input class="form-control" maxlength="60" data-kpf="iNup" value="' + esc(f.iNup) + '"></div>' +
    '<div class="col-6"><label class="form-label">Merk / Type <span class="text-muted fw-normal">(jika ada)</span></label><input class="form-control" maxlength="150" data-kpf="iMerk" value="' + esc(f.iMerk) + '"></div>' +
    '<div class="col-6"><label class="form-label">Nomor seri <span class="text-muted fw-normal">(jika ada)</span></label><input class="form-control" maxlength="100" data-kpf="iSeri" value="' + esc(f.iSeri) + '"></div>' +
    '<div class="col-6"><label class="form-label' + (h === 'Dibatalkan' ? '' : ' req') + '">Kondisi aset</label><select class="form-select" data-kpf="kondisi">' + opsi(listMaster('Kondisi Aset'), f.kondisi, '— pilih —') + '</select></div>' +
    '<div class="col-6"><label class="form-label' + (h === 'Dibatalkan' ? '' : ' req') + '">Status aset</label><select class="form-select" data-kpf="iStatus">' + opsi(listMaster('Status Aset').filter(function (x) { return x !== DIHAPUS; }), f.iStatus, '— pilih —') + '</select></div></div>' +
    '<div class="form-text">Jika Rusak Ringan atau Rusak Berat dipilih, Master Admin otomatis mendapat pemberitahuan.</div>';
  var verif = '';
  if (h) verif = '<div class="fw-bold small text-muted text-uppercase mt-3 mb-1">Verifikasi</div><div class="row g-2"><div class="col-12"><label class="form-label req">Verifikator (SPV)</label><select class="form-select" data-kpf="verif">' + opsi(WO.verif || [], f.verif, '— pilih verifikator —') + '</select></div>' +
    '<div class="col-6"><label class="form-label req">Tanggal verifikasi</label><input type="date" class="form-control" data-kpf="tgl" max="' + hariIniLokal() + '" value="' + esc(f.tgl) + '"><button type="button" class="btn-link-x small mt-1" data-act="kpHariIni">Pakai hari ini</button></div>' +
    '<div class="col-6"><label class="form-label">Status</label><select class="form-select" disabled><option>Terverifikasi</option></select></div></div>';
  var catatan = h ? '<label class="form-label mt-3">Catatan (opsional)</label><textarea class="form-control" rows="2" maxlength="1000" data-kpf="catatan" placeholder="Catatan tambahan untuk verifikator / Master Admin">' + esc(f.catatan) + '</textarea>' : '';
  return '<div class="card-x k-card k-form"><div class="fw-bold mb-2">Lapor hasil Preventive Maintenance</div>' + (S.kp.dipulihkan ? '<div class="alert-x warn small"><i class="bi bi-save me-1"></i>Draft laporan dipulihkan. Isian Anda tersimpan otomatis di perangkat ini.</div>' : '') +
    '<label class="form-label req">Status hasil</label>' + seg + cabang + kondisi + asetBaru + catatan + verif +
    '<div class="small text-warning mt-3 d-none" id="kpKurang"></div></div>' +
    '<button class="btn btn-primary k-btn" id="kpSimpan" data-act="kpSimpan" disabled><i class="bi bi-save me-2"></i>Simpan Laporan</button>';
}
function kpDokHtml() {
  return S.kp.docs.map(function (x, i) { return '<div class="d-flex align-items-center gap-2 small mb-1"><i class="bi bi-paperclip"></i><span class="text-truncate flex-grow-1">' + esc(x.n) + '</span><button type="button" class="icon-btn text-danger" data-act="kpDokHapus" data-i="' + i + '" title="Hapus dokumen" aria-label="Hapus dokumen ke-' + (i + 1) + '"><i class="bi bi-trash"></i></button></div>'; }).join('');
}
function kpGambarLapor() { var b = $('#kpLaporBox'); if (b) { b.innerHTML = kpLaporIsi(S.kp.d.pm); kpCekTombol(); } }
ACTIONS.kpHasil = function (el) { S.kp.form.hasil = el.dataset.v; if (el.dataset.v !== 'Ditunda') S.kp.form.alasan = ''; kpGambarLapor(); };
ACTIONS.kpTemuan = function (el) { S.kp.form.temuanYa = el.dataset.v; kpGambarLapor(); };
ACTIONS.kpSaran = function () { S.kp.form.periksa = kpSaranChecklist(S.kp.d.pm); kpGambarLapor(); };
ACTIONS.kpDokHapus = function (el) { S.kp.docs.splice(Number(el.dataset.i), 1); $('#kpDokList').innerHTML = kpDokHtml(); kpDraftSimpan(); };
ACTIONS.kpHariIni = function () { S.kp.form.tgl = hariIniLokal(); var i = $('[data-kpf=tgl]'); if (i) i.value = S.kp.form.tgl; kpCekTombol(); };
function kpKurang() {
  var f = S.kp.form, fo = S.kp.foto, k = [], p = S.kp.d.pm, td = kpAsetTerdaftar(p);
  if (!f.hasil) return ['Pilih status hasil'];
  var periksa = function () { if (!f.periksa.trim()) k.push('hasil pemeriksaan'); if (!f.temuanYa) k.push('pilihan ada/tidak ada temuan'); else if (f.temuanYa === 'Ya' && !f.temuan.trim()) k.push('deskripsi temuan'); };
  if (f.hasil === 'Terlaksana') {
    periksa(); if (f.temuanYa === 'Ya' && !f.tindakan.trim()) k.push('tindakan');
    if (!fo.b1 || !fo.b2) k.push('2 foto sebelum'); if (!fo.a1 || !fo.a2) k.push('2 foto sesudah');
    if (kpPerluSp()) {
      var sp = S.kp.sp || [];
      if (sp.length < SP_MIN) k.push('foto sparepart minimal ' + SP_MIN + ' (baru ' + sp.length + ')');
      if (sp.some(function (x) { return !String(x.c || '').trim(); })) k.push('keterangan tiap foto sparepart');
    }
  } else if (f.hasil === 'Ditunda') {
    if (!f.alasan) k.push('alasan penundaan'); if (!f.desk.trim()) k.push('deskripsi penundaan');
    if (f.alasan === 'Menunggu Sparepart') periksa();
  }
  if (f.hasil !== 'Dibatalkan' && !f.kondisi) k.push('kondisi aset');
  if (f.hasil !== 'Dibatalkan' && !td && !f.iStatus) k.push('status aset');
  if (!td && f.iThn && !/^\d{4}$/.test(f.iThn.trim())) k.push('tahun perolehan 4 digit');
  if (f.hasil === 'Terlaksana' && !td && !fo.ast) k.push('foto aset');
  if (!f.verif) k.push('verifikator'); if (!f.tgl) k.push('tanggal verifikasi');
  return k;
}
function kpCekTombol() {
  kpDraftSimpan();
  var b = $('#kpSimpan'); if (!b) return; var k = kpKurang();
  b.disabled = k.length > 0;
  var el = $('#kpKurang'); if (el) { el.classList.toggle('d-none', !k.length || !S.kp.form.hasil); el.textContent = k.length ? 'Belum lengkap: ' + k.join(', ') + '.' : ''; }
}
document.addEventListener('input', function (e) {
  var t = e.target;
  if (t && t.dataset && t.dataset.kpsp !== undefined && S.kPage === 'pmdetail') { var it = S.kp.sp[Number(t.dataset.kpsp)]; if (it) it.c = t.value; kpCekTombol(); return; }
  if (!t.dataset || t.dataset.kpf === undefined || S.kPage !== 'pmdetail') return;
  S.kp.form[t.dataset.kpf] = t.value; kpCekTombol();
});
document.addEventListener('change', async function (e) {
  var t = e.target; if (S.kPage !== 'pmdetail' || !t.dataset) return;
  if (t.dataset.kpf !== undefined && t.tagName === 'SELECT') { S.kp.form[t.dataset.kpf] = t.value; if (t.dataset.kpf === 'alasan') kpGambarLapor(); else kpCekTombol(); }
  else if (t.dataset.kpf === 'tgl') { S.kp.form.tgl = t.value; kpCekTombol(); }
  else if (t.dataset.slotIn && t.files && t.files[0]) {
    var key = t.dataset.slotIn, slot = t.closest('.foto-slot'); slot.classList.add('naik');
    try {
      var foto = await kompresFoto(t.files[0]);
      var u = await call('uploadFile', 'pm', S.kp.id, { name: foto.name, mime: foto.mime, base64: foto.base64 });
      catatFotoLokal(u.url, foto.preview); S.kp.foto[key] = u.url; var lab = slot.querySelector('.lbl-s').textContent; slot.outerHTML = kpSlot(key, lab); kpDraftSimpan();
    } catch (er) { slot.classList.remove('naik'); t.value = ''; toast(er.message, 'bad'); }
    kpCekTombol();
  } else if ((t.id === 'kpSpIn' || t.id === 'kpSpCam') && t.files && t.files.length) {
    var daftar = Array.prototype.slice.call(t.files), sisa = SP_MAKS - S.kp.sp.length; t.disabled = true;
    if (daftar.length > sisa) toast('Maksimal ' + SP_MAKS + ' foto sparepart; ' + (daftar.length - sisa) + ' foto terakhir diabaikan.', 'warn');
    for (var si = 0; si < Math.min(daftar.length, sisa); si++) {
      try {
        var cf = await kompresFoto(daftar[si]);
        var ur = await call('uploadFile', 'pm', S.kp.id, { name: cf.name, mime: cf.mime, base64: cf.base64 });
        catatFotoLokal(ur.url, cf.preview); S.kp.sp.push({ u: ur.url, c: '' });
      } catch (er3) { toast(er3.message, 'bad'); }
    }
    kpDraftSimpan(); kpGambarLapor();
  } else if ((t.id === 'kpDok' || t.id === 'kpDokCam') && t.files && t.files.length) {
    var list = Array.prototype.slice.call(t.files); t.disabled = true;
    for (var i = 0; i < list.length; i++) {
      var file = list[i];
      try {
        var body;
        if (file.type === 'application/pdf') {
          if (file.size > 2 * 1024 * 1024) throw new Error(file.name + ': ukuran dokumen PDF maksimal 2 MB (ukuran ' + (file.size / 1048576).toFixed(1).replace('.', ',') + ' MB).');
          body = { name: file.name, mime: 'application/pdf', base64: String(await bacaFile(file, false)).split(',')[1] };
        } else { var c = await kompresFoto(file); body = { name: c.name, mime: c.mime, base64: c.base64 }; }
        var r = await call('uploadFile', 'pm', S.kp.id, body); S.kp.docs.push({ n: file.name.slice(0, 80), u: r.unduh });
      } catch (er2) { toast(er2.message, 'bad'); }
    }
    t.value = ''; t.disabled = false; var dl = $('#kpDokList'); if (dl) dl.innerHTML = kpDokHtml(); kpDraftSimpan();
  }
});
ACTIONS.kpSimpan = async function (el) {
  var f = S.kp.form, fo = S.kp.foto, p0 = S.kp.d.pm, k = kpKurang(); if (k.length) { toast('Belum lengkap: ' + k.join(', ') + '.', 'warn'); return; }
  var p = { Status_Hasil: f.hasil, Catatan: f.catatan, Verifikator: f.verif, Tanggal_Verifikasi: f.tgl };
  if (f.hasil === 'Terlaksana' || (f.hasil === 'Ditunda' && f.alasan === 'Menunggu Sparepart')) { p.Hasil_Pemeriksaan = f.periksa; p.Ada_Temuan = f.temuanYa; p.Temuan = f.temuanYa === 'Ya' ? f.temuan : ''; }
  if (f.hasil === 'Terlaksana') { p.Foto_Before_1 = fo.b1; p.Foto_Before_2 = fo.b2; p.Foto_After_1 = fo.a1; p.Foto_After_2 = fo.a2; p.Tindakan = f.temuanYa === 'Ya' ? f.tindakan : ''; p.Dokumen = S.kp.docs; if (kpPerluSp()) p.Foto_Sparepart = S.kp.sp.map(function (x) { return { u: x.u, c: x.c.trim() }; }); }
  if (f.hasil === 'Ditunda') { p.Alasan_Tunda = f.alasan; p.Deskripsi_Tunda = f.desk; }
  if (f.kondisi) p.Kondisi_Dilaporkan = f.kondisi;
  if (!kpAsetTerdaftar(p0)) { if (fo.ast) p.Foto_Aset_URL = fo.ast; p.Lokasi_Detail = f.lokDetail; p.Info_Aset = { Tahun_Perolehan: f.iThn.trim(), NUP: f.iNup, Merk_Type: f.iMerk, Nomor_Seri: f.iSeri, Status_Aset: f.iStatus }; }
  if (!(await konfirm('Simpan laporan?', 'Status PM akan menjadi "' + f.hasil + '". Pastikan data sudah benar.', 'Simpan', false))) return;
  sibuk(el, true, 'Menyimpan…');
  try {
    await call('laporPm', p0.ID_PM, p); hapusDraft(p0.ID_PM); S.kp.draftLapor = false; toast('Laporan disimpan.', 'ok', 3500);
    S.kp.tab = f.hasil === 'Ditunda' ? 'ditunda' : 'selesai'; await kpMuatDetail('info');
  } catch (e) { gagal(e); sibuk(el, false); }
};

/* Lapor PM langsung: aset belum terdaftar (dari daftar PM) atau aset yang sedang dibuka (halaman aset) */
function kpFormLangsung(a) {
  return '<form id="fLP" class="k-form" data-submit="kpLaporLangsung" novalidate>' +
    (a ? '<div class="alert-x info small"><b>' + esc(a.Nama_Aset) + '</b><br>' + esc(a.ID_Aset) + ' · ' + esc(lokasiStr(a)) + '</div>' : '<div class="alert-x info small"><i class="bi bi-info-circle me-1"></i>Aset belum ada di Data Aset? Isi nama, kategori, dan lokasinya. Setelah Mulai, unggah foto aset di laporan; datanya dikirim ke Master Admin untuk didaftarkan.</div>' + apHtml()) +
    '<div class="my-3"><label class="form-label req">Jenis PM</label><select class="form-select" name="jenis">' + opsi(listMaster('Jenis PM'), '', '— pilih —') + '</select></div>' +
    '<div class="mb-3"><label class="form-label req">Aktivitas / checklist PM</label><textarea class="form-control" name="aktivitas" rows="3" maxlength="2000" placeholder="Satu aktivitas per baris"></textarea></div>' +
    '<div class="mb-2"><label class="form-label req">Prioritas perawatan</label><select class="form-select" name="prio">' + opsi(listMaster('Prioritas Perawatan'), '', '— pilih —') + '</select></div>' +
    '<div class="alert-x danger d-none mb-0" id="lpErr"></div></form>';
}
ACTIONS.kpLaporBaru = function () {
  apBaru(null, ''); AP.mode = 'lokasi'; AP.belum = true; AP.kategoriDulu = true; S.kp.asetLangsung = null;
  bukaModal({ judul: 'Lapor PM aset belum terdaftar', isi: kpFormLangsung(null), footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fLP">Buat &amp; lanjut</button>' });
};
ACTIONS.kpLapor = function () {
  var a = S.kAset && S.kAset.aset; if (!a) return; S.kp.asetLangsung = a;
  bukaModal({ judul: 'Lapor Preventive Maintenance', isi: kpFormLangsung(a), footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fLP">Buat &amp; lanjut</button>' });
};
SUBMITS.kpLaporLangsung = async function (f) {
  var err = $('#lpErr'); err.classList.add('d-none');
  var salah = function (m) { err.textContent = m; err.classList.remove('d-none'); };
  var p; try { p = S.kp.asetLangsung ? { ID_Aset: S.kp.asetLangsung.ID_Aset } : apNilai(); } catch (e) { salah(e.message); return; }
  if (!f.jenis.value || !f.aktivitas.value.trim() || !f.prio.value) { salah('Pilih jenis PM, isi aktivitas, dan pilih prioritas.'); return; }
  p.Jenis_PM = f.jenis.value; p.Aktivitas_Checklist = f.aktivitas.value; p.Prioritas_Perawatan = f.prio.value;
  var btn = $('button[type=submit][form=fLP]'); sibuk(btn, true, 'Membuat…');
  try {
    var d = await call('laporLangsungPm', p); gantiPmLokal(d.pm); var dari = S.kp.asetLangsung ? 'aset' : 'list'; tutupModal(); toast('PM ' + d.pm.ID_PM + ' dibuat.', 'ok');
    S.kp.id = d.pm.ID_PM; S.kp.dari = dari; await kpMuatDetail('mulai');
  } catch (e) { salah(e.message); } finally { sibuk(btn, false); }
};

/* Halaman aset: tab Preventive */
function kAsetPmHtml(a) {
  if (!PM.dimuat) return '<div class="skeleton" style="height:120px"></div>';
  var aktif = PM.list.filter(function (p) { return p.ID_Aset === a.ID_Aset && PM_AKTIF.indexOf(p.Status_PM) >= 0; }).sort(function (x, y) { return x.Due_Date < y.Due_Date ? -1 : 1; });
  var riw = PM.list.filter(function (p) { return p.ID_Aset === a.ID_Aset && PM_AKTIF.indexOf(p.Status_PM) < 0; }).sort(function (x, y) { return x.Tanggal_Jadwal < y.Tanggal_Jadwal ? 1 : -1; });
  var blok = function (jud, arr, kosong) { return '<div class="fw-bold small text-muted text-uppercase mb-2 mt-3">' + jud + ' (' + arr.length + ')</div>' + (arr.length ? arr.map(function (p) { return kpCard(p, 'aset'); }).join('') : '<div class="small text-muted mb-2">' + kosong + '</div>'); };
  return blok('PM terjadwal / berjalan untuk aset ini', aktif, 'Tidak ada PM aktif untuk aset ini.') + blok('Riwayat PM aset ini', riw.slice(0, 10), 'Belum ada riwayat PM.');
}

/* ════════════════════════════════════════════════════════
   RISK REGISTER — peta risiko 5×5, tabel, form, detail, grafik (Kabag/Kasubbag mengelola; SPV/PIC melihat)
   ════════════════════════════════════════════════════════ */
var RK = { list: [], level: [{ nama: 'Rendah', min: 1, max: 4 }, { nama: 'Sedang', min: 5, max: 9 }, { nama: 'Tinggi', min: 10, max: 16 }, { nama: 'Ekstrem', min: 17, max: 25 }], dimuat: false, f: { q: '', lv: '', kat: '', st: '', sel: '', hal: 1, per: 25 }, view: 'tabel' };
var RKF = null;
var RK_L = ['', 'Sangat jarang', 'Jarang', 'Mungkin terjadi', 'Sering', 'Hampir pasti'];
var RK_I = ['', 'Tidak signifikan', 'Kecil', 'Sedang', 'Besar', 'Katastropik'];
var RK_KLS = ['ok', 'info', 'warn', 'bad'];
function rkKelola() { return S.user && ['Kabag', 'Kasubbag', 'SPV', 'PIC'].indexOf(S.user.peran) >= 0; }
function rkIdx(skor) {
  var n = RK.level.length;
  for (var i = 0; i < n; i++) if (skor >= RK.level[i].min && skor <= RK.level[i].max) return n > 1 ? Math.round(i * 3 / (n - 1)) : 0;
  return 3;
}
function rkNamaLevel(skor) {
  for (var i = 0; i < RK.level.length; i++) if (skor >= RK.level[i].min && skor <= RK.level[i].max) return RK.level[i].nama;
  return RK.level[RK.level.length - 1].nama;
}
/** Level selalu dihitung dari skor memakai rentang terbaru (agar tetap benar bila rentang di Data Master diubah). */
function rkLv(r) { return rkNamaLevel(Number(r.Skor_Inherent)); }
function bLevelRk(r) { var s = Number(r.Skor_Inherent); return badge(rkNamaLevel(s) + ' · ' + s, RK_KLS[rkIdx(s)]); }
function rkTerlambat(r) { return r.Status_Pengendalian !== 'Terimplementasi' && r.Target_Penyelesaian && r.Target_Penyelesaian < hariIniLokal(); }
function bStatusRk(v) { return badge(v, { 'Belum Dimulai': '', 'Dalam Proses': 'warn', 'Terimplementasi': 'ok' }[v] || ''); }
async function muatRk() { var d = await call('getRiskList'); RK.list = d.risiko; RK.level = d.level; RK.dimuat = true; return d; }
function gantiRkLokal(r) {
  var ada = false; RK.list = RK.list.map(function (x) { if (x.ID_Risiko === r.ID_Risiko) { ada = true; return r; } return x; });
  if (!ada) RK.list.push(r);
}
function rkListStatus() { var l = listMaster('Status Risiko'); return l.length ? l : ['Belum Dimulai', 'Dalam Proses', 'Terimplementasi']; }

/* Peta risiko 5×5. Sumbu Y = Likelihood (5 di atas), sumbu X = Impact. kecil=true untuk Beranda. */
function rkPeta(data, kecil, aksi) {
  var h = '<div class="hm' + (kecil ? ' hm-kecil' : '') + '" role="grid" aria-label="Peta risiko 5 kali 5">';
  for (var l = 5; l >= 1; l--) {
    h += '<div class="hm-ax">' + l + '</div>';
    for (var i = 1; i <= 5; i++) {
      var n = data.filter(function (r) { return Number(r.Likelihood) === l && Number(r.Impact) === i; }).length, skor = l * i, kunci = l + '-' + i;
      h += '<button type="button" class="hm-c hm' + rkIdx(skor) + (RK.f.sel === kunci && !kecil ? ' pilih' : '') + (n ? '' : ' kosong') + '" data-act="' + aksi + '" data-s="' + kunci + '" aria-label="Likelihood ' + l + ', Impact ' + i + ', skor ' + skor + ', ' + n + ' risiko"><span class="hm-n">' + (n || '') + '</span>' + (kecil ? '' : '<span class="hm-s">' + skor + '</span>') + '</button>';
    }
  }
  h += '<div></div>' + [1, 2, 3, 4, 5].map(function (i) { return '<div class="hm-ax">' + i + '</div>'; }).join('') + '</div>' +
    '<div class="hm-leg"><span>↑ Likelihood · → Impact</span>' + RK.level.map(function (lv, i) { return '<span><i class="hm-k hm' + rkIdx(lv.min) + '"></i>' + esc(lv.nama) + ' (' + lv.min + '–' + lv.max + ')</span>'; }).join('') + '</div>';
  return h;
}

async function renderRk() {
  if (RK.view === 'grafik') { renderRkGrafik(); return; }
  $('#main').innerHTML = '<div class="skeleton" style="height:90px"></div><div class="skeleton mt-3" style="height:360px"></div>';
  try { await muatRk(); } catch (e) {
    if (S.page === 'risk') $('#main').innerHTML = '<div class="card-x"><div class="empty"><i class="bi bi-exclamation-circle"></i>' + esc(e.message) + '<div class="mt-3"><button class="btn btn-primary" data-act="rkMuat">Coba lagi</button></div></div></div>';
    return;
  }
  if (S.page !== 'risk' || RK.view === 'grafik') return;
  var f = RK.f;
  $('#main').innerHTML =
    '<div class="page-head"><h2>Risk Register</h2><div class="ms-auto d-flex flex-wrap gap-2">' +
    '<button class="btn btn-light-x" data-act="rkEkspor"><i class="bi bi-file-earmark-excel me-1"></i>Ekspor Excel</button>' +
    '<button class="btn btn-light-x" data-act="rkGrafik"><i class="bi bi-bar-chart-line me-1"></i>Lihat Grafik</button>' +
    (rkKelola() ? '<button class="btn btn-primary" data-act="rkBaru"><i class="bi bi-plus-lg me-1"></i>Tambah Risiko</button>' : '') + '</div></div>' +
    (rkKelola() ? '' : '<div class="alert-x info small"><i class="bi bi-eye me-1"></i>Anda dapat melihat Risk Register. Penambahan dan perubahan risiko dilakukan oleh Kabag, Kasubbag, SPV, dan PIC.</div>') +
    '<div class="kpi-grid" id="rkStat"></div>' +
    '<div class="grid-2"><div class="card-x"><div class="card-h"><i class="bi bi-grid-3x3-gap"></i>Peta Risiko<span class="badge-x ms-auto" id="rkSelInfo"></span></div><div class="card-b" id="rkPeta"></div></div>' +
    '<div class="card-x"><div class="card-h"><i class="bi bi-lightning-charge"></i>Risiko Teratas<span class="badge-x ms-auto">Skor tertinggi</span></div><div class="card-b" id="rkTop"></div></div></div>' +
    '<div class="card-x"><div class="toolbar">' +
    '<div class="grow input-icon"><i class="bi bi-search"></i><input class="form-control" data-rf="q" placeholder="Cari ID risiko, aset, deskripsi…" value="' + esc(f.q) + '" aria-label="Cari risiko"></div>' +
    '<select class="form-select fsel" data-rf="lv" aria-label="Level risiko">' + opsi(RK.level.map(function (x) { return x.nama; }), f.lv, 'Semua level') + '</select>' +
    '<select class="form-select fsel" data-rf="kat" aria-label="Kategori risiko">' + opsi(listMaster('Kategori Risiko'), f.kat, 'Semua kategori') + '</select>' +
    '<select class="form-select fsel" data-rf="st" aria-label="Status pengendalian">' + opsi(rkListStatus(), f.st, 'Semua status') + '</select>' +
    '<button class="btn btn-light-x btn-sm" data-act="rkReset">Reset</button></div>' +
    '<div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Risiko</th><th>Aset &amp; Lokasi</th><th>Skor (L × I)</th><th>Efektivitas</th><th>Status</th><th>Target</th><th></th></tr></thead><tbody id="rkBody"></tbody></table></div><div class="pager" id="rkPager"></div></div>';
  isiRk();
}
ACTIONS.rkMuat = function () { renderRk(); };
ACTIONS.rkReset = function () { RK.f = { q: '', lv: '', kat: '', st: '', sel: '', hal: 1, per: 25 }; renderRk(); };
function rkTersaring(abaikanSel) {
  var f = RK.f, q = f.q.trim();
  return RK.list.filter(function (r) {
    if (f.lv && rkLv(r) !== f.lv) return false;
    if (f.kat && r.Kategori_Risiko !== f.kat) return false;
    if (f.st && r.Status_Pengendalian !== f.st) return false;
    if (!abaikanSel && f.sel && (r.Likelihood + '-' + r.Impact) !== f.sel) return false;
    if (q && !cocokCari(q, r.ID_Risiko + ' ' + r.Nama_Aset + ' ' + r.ID_Aset + ' ' + r.Deskripsi_Risiko + ' ' + r.Kategori_Risiko)) return false;
    return true;
  });
}
function isiRk() {
  var f = RK.f, semua = RK.list, data = rkTersaring(false), total = data.length, hal = Math.max(1, Math.ceil(total / f.per));
  if (f.hal > hal) f.hal = hal; if (f.hal < 1) f.hal = 1;
  var mulai = (f.hal - 1) * f.per, tmpl = data.slice().sort(function (a, b) { return Number(b.Skor_Inherent) - Number(a.Skor_Inherent) || (a.ID_Risiko < b.ID_Risiko ? -1 : 1); }).slice(mulai, mulai + f.per);
  var n = function (fn) { return semua.filter(fn).length; };
  var idxTinggi = function (r) { return rkIdx(Number(r.Skor_Inherent)) >= 2; };
  $('#rkStat').innerHTML = kpiHtml('bi-exclamation-diamond', '', fmtAngka(semua.length), 'Total Risiko') +
    kpiHtml('bi-exclamation-triangle', 'bad', fmtAngka(n(idxTinggi)), 'Risiko Tinggi (Tinggi + Ekstrem)') +
    kpiHtml('bi-exclamation-circle', 'warn', fmtAngka(n(function (r) { return rkIdx(Number(r.Skor_Inherent)) === 1; })), 'Risiko Sedang') +
    kpiHtml('bi-shield-check', 'ok', fmtAngka(n(function (r) { return r.Status_Pengendalian === 'Terimplementasi'; })), 'Pengendalian Terimplementasi');
  $('#rkPeta').innerHTML = rkPeta(rkTersaring(true), false, 'rkSel');
  $('#rkSelInfo').innerHTML = f.sel ? 'L' + f.sel.split('-')[0] + ' × I' + f.sel.split('-')[1] + ' <button class="btn-link-x ms-1" data-act="rkSel" data-s="' + f.sel + '">hapus pilihan</button>' : 'Klik sel untuk memfilter';
  var top = semua.slice().sort(function (a, b) { return Number(b.Skor_Inherent) - Number(a.Skor_Inherent); }).slice(0, 5);
  $('#rkTop').innerHTML = top.length ? top.map(function (r) {
    return '<button type="button" class="riw-card" data-act="rkDetail" data-id="' + esc(r.ID_Risiko) + '"><div class="d-flex flex-wrap gap-1 align-items-center"><b>' + esc(r.ID_Risiko) + '</b>' + bLevelRk(r) + bStatusRk(r.Status_Pengendalian) + '</div><div class="riw-uraian">' + esc(r.Nama_Aset) + ' — ' + esc(r.Deskripsi_Risiko) + '</div></button>';
  }).join('') : '<div class="small text-muted">Belum ada risiko tercatat.</div>';
  $('#rkBody').innerHTML = tmpl.length ? tmpl.map(function (r) {
    return '<tr><td data-l="Risiko"><div class="nm">' + esc(r.ID_Risiko) + '</div><div class="sub">' + esc(r.Kategori_Risiko) + '</div><div class="sub">' + esc(String(r.Deskripsi_Risiko).slice(0, 90)) + (String(r.Deskripsi_Risiko).length > 90 ? '…' : '') + '</div></td>' +
      '<td data-l="Aset"><div class="nm">' + esc(r.Nama_Aset) + '</div><div class="sub">' + esc(r.ID_Aset) + ' · ' + esc(lokasiStr(r) || '—') + '</div></td>' +
      '<td data-l="Skor">' + bLevelRk(r) + '<div class="sub">L' + esc(r.Likelihood) + ' × I' + esc(r.Impact) + '</div></td>' +
      '<td data-l="Efektivitas"><div class="nm">' + esc(r.Efektivitas_Pengendalian || 0) + '%</div></td>' +
      '<td data-l="Status">' + bStatusRk(r.Status_Pengendalian) + '</td>' +
      '<td data-l="Target"><div class="nm">' + esc(fmtTgl(r.Target_Penyelesaian, true)) + '</div>' + (rkTerlambat(r) ? '<div class="sub">' + badge('Lewat target', 'bad') + '</div>' : '') + '</td>' +
      '<td class="acts"><button class="icon-btn" data-act="rkDetail" data-id="' + esc(r.ID_Risiko) + '" title="Detail" aria-label="Detail ' + esc(r.ID_Risiko) + '"><i class="bi bi-eye"></i></button>' +
      (rkKelola() ? '<button class="icon-btn" data-act="rkUbah" data-id="' + esc(r.ID_Risiko) + '" title="Ubah" aria-label="Ubah ' + esc(r.ID_Risiko) + '"><i class="bi bi-pencil"></i></button><button class="icon-btn text-danger" data-act="rkHapus" data-id="' + esc(r.ID_Risiko) + '" title="Hapus" aria-label="Hapus ' + esc(r.ID_Risiko) + '"><i class="bi bi-trash"></i></button>' : '') + '</td></tr>';
  }).join('') : '<tr><td colspan="7"><div class="empty"><i class="bi bi-shield-check"></i>' + (RK.list.length ? 'Tidak ada risiko yang cocok dengan filter.' : 'Belum ada risiko tercatat.') + '</div></td></tr>';
  $('#rkPager').innerHTML = '<div>' + (total ? (mulai + 1) + '–' + (mulai + tmpl.length) + ' dari ' + fmtAngka(total) + ' risiko' : '0 risiko') + '</div>' +
    '<div class="d-flex align-items-center gap-2"><button class="btn btn-light-x btn-sm" data-act="rkHal" data-d="-1"' + (f.hal <= 1 ? ' disabled' : '') + '><i class="bi bi-chevron-left"></i></button><span>Hal. ' + f.hal + ' / ' + hal + '</span><button class="btn btn-light-x btn-sm" data-act="rkHal" data-d="1"' + (f.hal >= hal ? ' disabled' : '') + '><i class="bi bi-chevron-right"></i></button></div>';
}
ACTIONS.rkHal = function (el) { RK.f.hal += Number(el.dataset.d); isiRk(); };
ACTIONS.rkSel = function (el) { RK.f.sel = RK.f.sel === el.dataset.s ? '' : el.dataset.s; RK.f.hal = 1; isiRk(); };
document.addEventListener('input', debounce(function (e) { var t = e.target; if (S.page === 'risk' && t.dataset && t.dataset.rf === 'q') { RK.f.q = t.value; RK.f.hal = 1; isiRk(); } }, 200));
document.addEventListener('change', function (e) { var t = e.target; if (S.page === 'risk' && t.dataset && t.dataset.rf && t.dataset.rf !== 'q') { RK.f[t.dataset.rf] = t.value; RK.f.hal = 1; isiRk(); } });

/* Grafik */
ACTIONS.rkGrafik = function () { RK.view = 'grafik'; hancurkanChart(); renderRkGrafik(); window.scrollTo(0, 0); };
ACTIONS.kembaliRk = function () { RK.view = 'tabel'; renderRk(); };
function renderRkGrafik() {
  hancurkanChart();
  var data = rkTersaring(false);
  var kotak = function (id, judul, ket) { return '<div class="card-x"><div class="card-h">' + judul + (ket ? '<span class="badge-x ms-auto">' + ket + '</span>' : '') + '</div><div class="card-b"><div class="chart-box"><canvas id="' + id + '" role="img" aria-label="' + esc(judul) + '"></canvas></div></div></div>'; };
  $('#main').innerHTML = '<div class="page-head"><button class="btn btn-light-x" data-act="kembaliRk"><i class="bi bi-arrow-left me-1"></i>Kembali ke Risk Register</button><h2>Grafik Risk Register</h2></div>' +
    '<div class="alert-x info small"><i class="bi bi-funnel me-1"></i>Grafik mengikuti filter di halaman Risk Register (' + fmtAngka(data.length) + ' risiko).</div>' +
    (data.length ? '<div class="grid-2">' + kotak('grKat', 'Risiko per Kategori') + kotak('grSt', 'Status Pengendalian') + '</div>' + kotak('grLv', 'Distribusi Level Risiko') : '<div class="card-x"><div class="empty"><i class="bi bi-bar-chart-line"></i>Belum ada risiko untuk digrafikkan.</div></div>');
  if (!data.length || typeof Chart === 'undefined') return;
  var gelap = S.theme === 'dark', grid = gelap ? '#22304D' : '#E2E8F0';
  Chart.defaults.color = gelap ? '#98A5BF' : '#6B7280'; Chart.defaults.font.family = "'Plus Jakarta Sans',sans-serif";
  var kat = hitung(data, 'Kategori_Risiko');
  S.charts.push(new Chart($('#grKat'), { type: 'bar', data: { labels: kat.map(function (x) { return x[0]; }), datasets: [{ label: 'Jumlah risiko', data: kat.map(function (x) { return x[1]; }), backgroundColor: '#1E5EFF', borderRadius: 4 }] },
    options: { indexAxis: 'y', maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { precision: 0 }, grid: { color: grid } }, y: { grid: { display: false } } } } }));
  var st = hitung(data, 'Status_Pengendalian'), wst = { 'Belum Dimulai': '#94A3B8', 'Dalam Proses': '#F59E0B', 'Terimplementasi': '#16A34A' };
  S.charts.push(new Chart($('#grSt'), { type: 'doughnut', data: { labels: st.map(function (x) { return x[0]; }), datasets: [{ data: st.map(function (x) { return x[1]; }), backgroundColor: st.map(function (x, i) { return wst[x[0]] || PALET[i % 8]; }), borderWidth: 0 }] }, options: { maintainAspectRatio: false, cutout: '62%', plugins: { legend: { position: 'bottom' } } } }));
  var wl = ['#16A34A', '#EAB308', '#F97316', '#DC2626'];
  S.charts.push(new Chart($('#grLv'), { type: 'bar', data: { labels: RK.level.map(function (x) { return x.nama; }), datasets: [{ label: 'Jumlah risiko', data: RK.level.map(function (x) { return data.filter(function (r) { return rkNamaLevel(Number(r.Skor_Inherent)) === x.nama; }).length; }), backgroundColor: RK.level.map(function (x) { return wl[rkIdx(x.min)]; }), borderRadius: 4 }] },
    options: { maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { ticks: { precision: 0 }, grid: { color: grid } }, x: { grid: { display: false } } } } }));
}

/* Ekspor Excel */
ACTIONS.rkEkspor = async function (el) {
  var data = rkTersaring(false); if (!data.length) { toast('Tidak ada data untuk diekspor.', 'warn'); return; }
  var kolom = ['ID_Risiko', 'ID_Aset', 'Nama_Aset', 'Gedung', 'Lantai', 'Ruangan', 'Lokasi_Detail', 'Kategori_Risiko', 'Deskripsi_Risiko', 'Penyebab_Risiko', 'Dampak_Risiko', 'Likelihood', 'Impact', 'Skor_Inherent', 'Level_Inherent', 'Pengendalian_Existing', 'Efektivitas_Pengendalian', 'Treatment_Mitigasi', 'Target_Penyelesaian', 'Status_Pengendalian', 'Dibuat_Oleh', 'Dibuat_Pada'];
  var tabel = [kolom].concat(data.map(function (r) { return kolom.map(function (k) { return amanSel(k === 'Level_Inherent' ? rkLv(r) : r[k]); }); }));
  sibuk(el, true, 'Menyiapkan…');
  var nama = 'Risk_Register_' + new Date().toISOString().slice(0, 10);
  try {
    await muatSkrip(CDN_XLSX, function () { return typeof XLSX !== 'undefined'; });
    var ws = XLSX.utils.aoa_to_sheet(tabel), wb = XLSX.utils.book_new();
    ws['!cols'] = kolom.map(function (k) { return { wch: Math.min(40, Math.max(k.length + 2, 14)) }; });
    XLSX.utils.book_append_sheet(wb, ws, 'Risk Register');
    unduhBlob(nama + '.xlsx', new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    toast('Berkas Excel diunduh (' + data.length + ' risiko).', 'ok');
  } catch (e) {
    unduhBlob(nama + '.csv', new Blob([csvKeString(tabel)], { type: 'text/csv;charset=utf-8' }));
    toast('Pustaka Excel tidak tersedia — diunduh sebagai CSV (dapat dibuka di Excel).', 'warn', 6000);
  } finally { sibuk(el, false); }
};

/* Detail */
function detailRkHtml(d) {
  var r = d.risiko, a = d.aset, s = Number(r.Skor_Inherent), ef = Number(r.Efektivitas_Pengendalian) || 0;
  var wo = d.wo.length ? d.wo.map(function (w) {
    return '<div class="riw-card riw-statis"><div class="d-flex flex-wrap gap-1 align-items-center"><b>' + esc(w.ID_WO) + '</b>' + bStatusWo(w.Status_WO) + bPrio(w.Prioritas) + '</div><div class="riw-uraian">' + esc(w.Uraian_Pekerjaan) + '</div><div class="riw-meta"><i class="bi bi-calendar-event me-1"></i>' + esc(fmtTgl(w.Tanggal_Pengerjaan, true)) + '</div></div>';
  }).join('') : '<div class="small text-muted">Belum ada Work Order dari risiko ini.</div>';
  var ubahSt = rkKelola() ? '<div class="d-flex flex-wrap gap-2 mt-2" role="group" aria-label="Ubah status pengendalian">' + rkListStatus().map(function (x) {
    return '<button type="button" class="btn btn-sm ' + (x === r.Status_Pengendalian ? 'btn-primary' : 'btn-light-x') + '" data-act="rkStatus" data-id="' + esc(r.ID_Risiko) + '" data-v="' + esc(x) + '"' + (x === r.Status_Pengendalian ? ' disabled' : '') + '>' + esc(x) + '</button>';
  }).join('') + '</div>' : '';
  return woSeksi('Identitas & Status', '<dl class="dl">' + dlRow('ID Risiko', '<b>' + esc(r.ID_Risiko) + '</b>') + dlRow('Kategori', esc(r.Kategori_Risiko)) + dlRow('Level (inherent)', bLevelRk(r)) + dlRow('Status Pengendalian', bStatusRk(r.Status_Pengendalian) + (rkTerlambat(r) ? ' ' + badge('Lewat target', 'bad') : '')) +
      dlRow('Dibuat', esc(r.Dibuat_Oleh) + ' · ' + esc(fmtTgl(r.Dibuat_Pada))) + dlRow('Diubah', esc(fmtTgl(r.Diubah_Pada))) + '</dl>' + ubahSt) +
    woSeksi('Aset & Lokasi', '<dl class="dl">' + dlRow('ID Aset', esc(r.ID_Aset)) + dlRow('Nama Aset', esc(r.Nama_Aset)) +
      (a ? dlRow('Kategori Aset', esc(a.Kategori_Aset)) + dlRow('Kritikalitas', bKrit(a.Kritikalitas)) + dlRow('Kondisi', bKondisi(a.Kondisi)) : '') +
      dlRow('Gedung', esc(r.Gedung)) + dlRow('Lantai', esc(r.Lantai)) + dlRow('Ruangan', esc(r.Ruangan)) + dlRow('Lokasi Detail', esc(r.Lokasi_Detail)) + '</dl>') +
    woSeksi('Penilaian Risiko', '<dl class="dl">' + dlRow('Likelihood', esc(r.Likelihood) + ' — ' + esc(RK_L[Number(r.Likelihood)] || '')) + dlRow('Impact', esc(r.Impact) + ' — ' + esc(RK_I[Number(r.Impact)] || '')) + dlRow('Skor Inherent', '<b>' + s + '</b> (L × I)') + '</dl>') +
    woSeksi('Uraian Risiko', '<dl class="dl">' + dlRow('Deskripsi', teksPanjang(r.Deskripsi_Risiko)) + dlRow('Penyebab', teksPanjang(r.Penyebab_Risiko)) + dlRow('Dampak', teksPanjang(r.Dampak_Risiko)) + '</dl>' +
      (r.Foto_Risiko_URL ? '<div class="foto-grid mt-2">' + fotoThumb(r.Foto_Risiko_URL, 'Foto risiko') + '</div>' : '')) +
    woSeksi('Pengendalian & Mitigasi', '<dl class="dl">' + dlRow('Pengendalian existing', teksPanjang(r.Pengendalian_Existing)) +
      dlRow('Efektivitas', '<div class="d-flex align-items-center gap-2"><div class="progress flex-grow-1" style="height:8px" role="progressbar" aria-valuenow="' + ef + '" aria-valuemin="0" aria-valuemax="100"><div class="progress-bar" style="width:' + ef + '%"></div></div><b>' + ef + '%</b></div>') +
      dlRow('Treatment / mitigasi', teksPanjang(r.Treatment_Mitigasi)) + dlRow('Target penyelesaian', esc(fmtTgl(r.Target_Penyelesaian, true))) + '</dl>') +
    woSeksi('Work Order terkait · ' + d.wo.length, wo);
}
ACTIONS.rkDetail = async function (el) {
  var id = el.dataset.id;
  bukaModal({ judul: 'Risiko ' + id, size: 'lg', isi: '<div class="skeleton" style="height:240px"></div>', footer: '<button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>' });
  try { await rkIsiDetail(id); } catch (e) { tutupModal(); gagal(e); }
};
async function rkIsiDetail(id) {
  var d = await call('getRiskDetail', id); RK.level = d.level; gantiRkLokal(d.risiko);
  $('#mdlContent .modal-body').innerHTML = detailRkHtml(d);
  $('#mdlContent .modal-footer').innerHTML = rkKelola() ?
    '<button class="btn btn-light-x text-danger me-auto" data-act="rkHapus" data-id="' + esc(id) + '" data-wo="' + d.wo.length + '"><i class="bi bi-trash me-1"></i>Hapus</button>' +
    '<button class="btn btn-light-x" data-act="rkBuatWo" data-id="' + esc(id) + '"><i class="bi bi-clipboard-plus me-1"></i>Buat Work Order</button>' +
    '<button class="btn btn-light-x" data-act="rkUbah" data-id="' + esc(id) + '"><i class="bi bi-pencil me-1"></i>Ubah</button><button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>' :
    '<button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>';
}
ACTIONS.rkStatus = async function (el) {
  try { var d = await call('updateStatusRisk', el.dataset.id, el.dataset.v); gantiRkLokal(d.risiko); toast('Status pengendalian diperbarui.', 'ok'); await rkIsiDetail(el.dataset.id); if (S.page === 'risk') isiRk(); } catch (e) { gagal(e); }
};
ACTIONS.rkHapus = async function (el) {
  var id = el.dataset.id;
  var jmlWo = Number(el.dataset.wo || 0);
  if (!(await konfirm('Hapus risiko?', id + ' akan dihapus permanen.' + (jmlWo ? ' ' + jmlWo + ' Work Order terkait tetap ada dan tidak ikut terhapus.' : ''), 'Hapus', true))) return;
  try { await call('deleteRisk', id); RK.list = RK.list.filter(function (x) { return x.ID_Risiko !== id; }); tutupModal(); toast('Risiko dihapus.', 'ok'); if (S.page === 'risk') isiRk(); } catch (e) { gagal(e); }
};
/* Buat WO dengan Sumber = Risk Register; ID Risiko & aset terisi otomatis (US-50 AC4). */
ACTIONS.rkBuatWo = function (el) {
  var r = RK.list.filter(function (x) { return x.ID_Risiko === el.dataset.id; })[0]; if (!r) return;
  tutupModal();
  setTimeout(function () {
    bukaFormWo({ Sumber_WO: 'Risk Register', ID_Risiko: r.ID_Risiko, ID_Aset: r.ID_Aset, Uraian_Pekerjaan: r.Treatment_Mitigasi, Temuan_Indikasi: r.Deskripsi_Risiko.slice(0, 500) }, null, false);
  }, 250);
};

/* Form tambah / ubah */
function formRkHtml(r) {
  var ubah = !!(RKF && RKF.mode === 'ubah');
  var opsiSkala = function (arr, sel) { var h = '<option value="">— pilih —</option>'; for (var i = 1; i <= 5; i++) h += '<option value="' + i + '"' + (String(sel) === String(i) ? ' selected' : '') + '>' + i + ' — ' + esc(arr[i]) + '</option>'; return h; };
  var asetBlok = ubah ? '<div class="ap-sel"><div class="min-w-0 flex-grow-1"><div class="fw-bold">' + esc(r.Nama_Aset) + '</div><div class="small text-muted">' + esc(r.ID_Aset) + ' · ' + esc(lokasiLengkap(r)) + '</div></div></div><div class="form-text">Aset tidak dapat diganti setelah risiko dibuat.</div>' : apHtml();
  return '<form id="fRk" data-submit="simpanRk" novalidate><div class="row g-3">' +
    '<div class="col-12"><div class="fw-bold small text-muted text-uppercase">Aset</div>' + asetBlok + '</div>' +
    '<div class="col-md-6"><label class="form-label req">Kategori Risiko</label><select class="form-select" name="Kategori_Risiko" required>' + opsi(listMaster('Kategori Risiko'), r && r.Kategori_Risiko, '— pilih —') + '</select></div>' +
    '<div class="col-md-6"><label class="form-label">Foto Risiko</label><input type="file" class="form-control" id="rkFotoIn" accept="image/*"><div class="form-text">Foto dikompres otomatis.</div></div>' +
    '<div class="col-12 d-flex align-items-center gap-3" id="rkFotoPrev">' + (r && r.Foto_Risiko_URL ? '<img class="thumb" style="width:72px;height:72px" src="' + esc(imgU(r.Foto_Risiko_URL)) + '" alt="Foto risiko" referrerpolicy="no-referrer"><button type="button" class="btn btn-light-x btn-sm" data-act="rkHapusFoto">Hapus foto</button>' : '') + '</div>' +
    '<div class="col-12"><label class="form-label req">Deskripsi Risiko</label><textarea class="form-control" name="Deskripsi_Risiko" rows="2" maxlength="1000" required>' + esc(r && r.Deskripsi_Risiko) + '</textarea></div>' +
    '<div class="col-md-6"><label class="form-label req">Penyebab Risiko</label><textarea class="form-control" name="Penyebab_Risiko" rows="2" maxlength="1000" required>' + esc(r && r.Penyebab_Risiko) + '</textarea></div>' +
    '<div class="col-md-6"><label class="form-label req">Dampak Risiko</label><textarea class="form-control" name="Dampak_Risiko" rows="2" maxlength="1000" required>' + esc(r && r.Dampak_Risiko) + '</textarea></div>' +
    '<div class="col-md-4"><label class="form-label req">Likelihood (1–5)</label><select class="form-select" name="Likelihood" required>' + opsiSkala(RK_L, r && r.Likelihood) + '</select></div>' +
    '<div class="col-md-4"><label class="form-label req">Impact (1–5)</label><select class="form-select" name="Impact" required>' + opsiSkala(RK_I, r && r.Impact) + '</select></div>' +
    '<div class="col-md-4 d-flex align-items-end"><div class="rk-skor w-100" id="rkSkor" aria-live="polite"></div></div>' +
    '<div class="col-12"><label class="form-label">Pengendalian Existing</label><textarea class="form-control" name="Pengendalian_Existing" rows="2" maxlength="1000">' + esc(r && r.Pengendalian_Existing) + '</textarea></div>' +
    '<div class="col-md-4"><label class="form-label">Efektivitas Pengendalian (%)</label><input type="number" class="form-control" name="Efektivitas_Pengendalian" min="0" max="100" step="1" value="' + esc(r ? r.Efektivitas_Pengendalian : '') + '" placeholder="0–100"></div>' +
    '<div class="col-md-4"><label class="form-label req">Target Penyelesaian</label><input type="date" class="form-control" name="Target_Penyelesaian" value="' + esc(r && r.Target_Penyelesaian) + '" required></div>' +
    '<div class="col-md-4"><label class="form-label req">Status Pengendalian</label><select class="form-select" name="Status_Pengendalian" required>' + opsi(rkListStatus(), (r && r.Status_Pengendalian) || 'Belum Dimulai', null) + '</select></div>' +
    '<div class="col-12"><label class="form-label req">Treatment / Rencana Mitigasi</label><textarea class="form-control" name="Treatment_Mitigasi" rows="2" maxlength="1000" required>' + esc(r && r.Treatment_Mitigasi) + '</textarea></div>' +
    '<div class="col-12"><div class="alert-x danger d-none mb-0" id="rkErr"></div></div></div></form>';
}
function rkHitungSkor() {
  var f = $('#fRk'), el = $('#rkSkor'); if (!f || !el) return;
  var l = Number(f.Likelihood.value), i = Number(f.Impact.value);
  if (!l || !i) { el.innerHTML = '<span class="text-muted small">Skor inherent = L × I</span>'; return; }
  el.innerHTML = '<span class="small text-muted">Skor inherent</span> <b>' + (l * i) + '</b> ' + badge(rkNamaLevel(l * i), RK_KLS[rkIdx(l * i)]);
}
function bukaFormRk(r, ubah) {
  RKF = { mode: ubah ? 'ubah' : 'baru', id: r && r.ID_Risiko, foto: null, hapusFoto: false };
  if (!ubah) { apBaru(null, ''); AP.tdSaja = true; AP.mode = 'pilih'; AP.belum = false; }
  bukaModal({ judul: ubah ? 'Ubah ' + r.ID_Risiko : 'Tambah Risiko', size: 'lg', isi: formRkHtml(r),
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fRk">' + (ubah ? 'Simpan perubahan' : 'Simpan risiko') + '</button>' });
  rkHitungSkor();
}
ACTIONS.rkBaru = function () { bukaFormRk(null, false); };
ACTIONS.rkUbah = async function (el) {
  var id = el.dataset.id;
  try { var d = await call('getRiskDetail', id); if (document.querySelector('#mdl.show')) tutupModal(); setTimeout(function () { bukaFormRk(d.risiko, true); }, 250); } catch (e) { gagal(e); }
};
ACTIONS.rkHapusFoto = function () { RKF.hapusFoto = true; RKF.foto = null; $('#rkFotoPrev').innerHTML = ''; $('#rkFotoIn').value = ''; };
document.addEventListener('change', async function (e) {
  var t = e.target; if (!t.closest || !t.closest('#fRk')) return;
  if (t.name === 'Likelihood' || t.name === 'Impact') rkHitungSkor();
  if (t.id === 'rkFotoIn' && t.files && t.files[0]) {
    try { RKF.foto = await kompresFoto(t.files[0]); RKF.hapusFoto = false; $('#rkFotoPrev').innerHTML = '<img class="thumb" style="width:72px;height:72px" src="' + RKF.foto.preview + '" alt="Pratinjau"><span class="small text-muted">Foto baru siap diunggah</span>'; }
    catch (er) { toast(er.message, 'bad'); t.value = ''; RKF.foto = null; }
  }
});
SUBMITS.simpanRk = async function (f) {
  var err = $('#rkErr'); err.classList.add('d-none');
  var salah = function (m) { err.textContent = m; err.classList.remove('d-none'); };
  var p = {};
  if (RKF.mode === 'baru') { try { p = apNilai(); } catch (e) { salah(e.message); return; } if (!p.ID_Aset) { salah('Pilih aset yang sudah terdaftar di Data Aset.'); return; } }
  ['Kategori_Risiko', 'Deskripsi_Risiko', 'Penyebab_Risiko', 'Dampak_Risiko', 'Likelihood', 'Impact', 'Pengendalian_Existing', 'Efektivitas_Pengendalian', 'Treatment_Mitigasi', 'Target_Penyelesaian', 'Status_Pengendalian'].forEach(function (k) { p[k] = f[k].value; });
  if (!p.Kategori_Risiko || !p.Deskripsi_Risiko.trim() || !p.Penyebab_Risiko.trim() || !p.Dampak_Risiko.trim() || !p.Likelihood || !p.Impact || !p.Treatment_Mitigasi.trim() || !p.Target_Penyelesaian || !p.Status_Pengendalian) { salah('Lengkapi semua kolom bertanda *.'); return; }
  var ef = String(p.Efektivitas_Pengendalian).trim();
  if (ef !== '' && (!/^\d{1,3}$/.test(ef) || Number(ef) > 100)) { salah('Efektivitas Pengendalian harus angka 0–100.'); return; }
  var lama = RKF.mode === 'ubah' ? RK.list.filter(function (x) { return x.ID_Risiko === RKF.id; })[0] : null;
  if (RKF.mode === 'ubah') { p.ID_Risiko = RKF.id; p.Foto_Risiko_URL = RKF.hapusFoto ? '' : (lama ? lama.Foto_Risiko_URL : ''); }
  var btn = $('button[type=submit][form=fRk]'); sibuk(btn, true, 'Menyimpan…');
  try {
    if (RKF.foto) { var u = await call('uploadFile', 'risiko', RKF.mode === 'ubah' ? RKF.id : 'baru', { name: RKF.foto.name, mime: RKF.foto.mime, base64: RKF.foto.base64 }); p.Foto_Risiko_URL = u.url; }
    var d = await call('saveRisk', p, RKF.mode);
    gantiRkLokal(d.risiko); tutupModal(); toast((RKF.mode === 'ubah' ? 'Risiko diperbarui.' : 'Risiko ' + d.risiko.ID_Risiko + ' ditambahkan.') + (d.woOtomatis ? ' Skor ≥ 10 atau kategori Listrik: Work Order ' + d.woOtomatis + ' dibuat otomatis.' : ''), d.woOtomatis ? 'warn' : 'ok'); if (d.woOtomatis) { try { WO.dimuat = false; } catch (e2) {} }
    if (S.page === 'risk') isiRk();
  } catch (e) { salah(e.message); } finally { sibuk(btn, false); }
};

/* Beranda: peta risiko ringkas (Kabag/Kasubbag/SPV/PIC) */
async function berandaRisiko() {
  var el = $('#rkBeranda'); if (!el) return;
  try {
    await muatRk(); el = $('#rkBeranda'); if (!el || S.page !== 'beranda') return;
    var tinggi = RK.list.filter(function (r) { return rkIdx(Number(r.Skor_Inherent)) >= 2 && r.Status_Pengendalian !== 'Terimplementasi'; }).length;
    el.innerHTML = '<div class="card-h"><i class="bi bi-grid-3x3-gap"></i>Peta Risiko<span class="badge-x ms-auto">' + fmtAngka(RK.list.length) + ' risiko · ' + fmtAngka(tinggi) + ' tinggi belum tertangani</span></div><div class="card-b">' + rkPeta(RK.list, true, 'rkBerandaSel') + '<div class="mt-2"><button class="btn-link-x" data-act="go" data-p="risk">Buka Risk Register <i class="bi bi-arrow-right"></i></button></div></div>';
  } catch (e) { if ($('#rkBeranda')) $('#rkBeranda').innerHTML = '<div class="card-b small text-muted">Peta risiko belum dapat dimuat.</div>'; }
}
ACTIONS.rkBerandaSel = function (el) { RK.f.sel = el.dataset.s; RK.f.hal = 1; go('risk'); };

/* ════════════════════════════════════════════════════════
   K3 & SAFETY — ringkasan, Inspeksi K3, Insiden, SOP, tindak lanjut (Kabag/Kasubbag/SPV/PIC)
   ════════════════════════════════════════════════════════ */
var K3 = { inspeksi: [], insiden: [], sop: [], dimuat: false, tab: 'ringkasan', per: 25,
  fi: { q: '', bahaya: '', lv: '', tl: '', hal: 1 }, fn: { q: '', jenis: '', kep: '', tl: '', hal: 1 }, fs: { q: '', kat: '', st: '', hal: 1 } };
var K3F = null;
var K3_STATUS_TL = ['Belum Dimulai', 'Dalam Proses', 'Selesai'];
var K3_STATUS_SOP = ['Berlaku', 'Dalam Revisi', 'Tidak Berlaku'];
function k3ListTl() { var l = listMaster('Status Temuan'); return l.length ? l : K3_STATUS_TL; }
async function muatK3() { var d = await call('getK3Data'); K3.inspeksi = d.inspeksi; K3.insiden = d.insiden; K3.sop = d.sop; RK.level = d.level; K3.dimuat = true; return d; }
function k3Ganti(jenis, r) {
  var arr = K3[jenis], kunci = { inspeksi: 'ID_Inspeksi', insiden: 'ID_Insiden', sop: 'ID_SOP' }[jenis], ada = false;
  K3[jenis] = arr.map(function (x) { if (x[kunci] === r[kunci]) { ada = true; return r; } return x; });
  if (!ada) K3[jenis].push(r);
}
function k3FotoArr(r) { try { return JSON.parse(r.Foto_URLs || '[]') || []; } catch (e) { return []; } }
function bTlK3(v) { return v ? badge(v, { 'Belum Dimulai': '', 'Dalam Proses': 'warn', 'Selesai': 'ok' }[v] || '') : '<span class="text-muted small">Tanpa tindak lanjut</span>'; }
function k3Telat(r) { return r.Status_Tindak_Lanjut && r.Status_Tindak_Lanjut !== 'Selesai' && r.Due_Date && r.Due_Date < hariIniLokal(); }
function bSkorK3(r) { var s = Number(r.Skor_Risiko); return badge(rkNamaLevel(s) + ' · ' + s, RK_KLS[rkIdx(s)]); }
function bYaTidak(v, buruk) { return v ? badge(v, v === 'Ya' ? (buruk ? 'bad' : 'ok') : '') : '<span class="text-muted">—</span>'; }
function k3Bulan(t) { return String(t || '').slice(0, 7); }
function k3LabelBulan(b) { return BLN[Number(b.slice(5)) - 1] + ' ' + b.slice(2, 4); }
function k3Sel(arr, nama, v, kosong) { return '<select class="form-select fsel" data-k3f="' + nama + '" aria-label="' + esc(kosong) + '">' + opsi(arr, v, kosong) + '</select>'; }
function k3TlSel() { return k3ListTl(); }

/* ── Halaman & tab ── */
async function renderK3() {
  $('#main').innerHTML = '<div class="skeleton" style="height:90px"></div><div class="skeleton mt-3" style="height:360px"></div>';
  try { await muatK3(); } catch (e) {
    if (S.page === 'k3') $('#main').innerHTML = '<div class="card-x"><div class="empty"><i class="bi bi-exclamation-circle"></i>' + esc(e.message) + '<div class="mt-3"><button class="btn btn-primary" data-act="k3Muat">Coba lagi</button></div></div></div>';
    return;
  }
  if (S.page !== 'k3') return;
  var tabs = [['ringkasan', 'Ringkasan'], ['inspeksi', 'Inspeksi K3'], ['insiden', 'Insiden'], ['sop', 'SOP K3']];
  $('#main').innerHTML = '<div class="page-head"><h2>K3 &amp; Safety</h2><div class="ms-auto d-flex flex-wrap gap-2" id="k3Aksi"></div></div>' +
    '<ul class="nav nav-tabs mb-3" style="overflow-x:auto;flex-wrap:nowrap">' + tabs.map(function (t) { return '<li class="nav-item"><button class="nav-link text-nowrap' + (K3.tab === t[0] ? ' active' : '') + '" data-act="k3Tab" data-t="' + t[0] + '">' + t[1] + '</button></li>'; }).join('') + '</ul><div id="k3Body"></div>';
  k3Isi();
}
ACTIONS.k3Muat = function () { renderK3(); };
ACTIONS.k3Tab = function (el) {
  K3.tab = el.dataset.t; hancurkanChart();
  $$('#main .nav-link').forEach(function (b) { b.classList.toggle('active', b.dataset.t === K3.tab); });
  k3Isi(); window.scrollTo(0, 0);
};
function k3Isi() {
  hancurkanChart();
  var a = { inspeksi: ['k3Baru', 'Inspeksi'], insiden: ['k3Baru', 'Insiden'], sop: ['k3Baru', 'SOP'] }[K3.tab];
  $('#k3Aksi').innerHTML = (K3.tab !== 'ringkasan' ? '<button class="btn btn-light-x" data-act="k3Ekspor"><i class="bi bi-file-earmark-excel me-1"></i>Ekspor Excel</button><button class="btn btn-primary" data-act="k3Baru"><i class="bi bi-plus-lg me-1"></i>' + (K3.tab === 'sop' ? 'Tambah SOP' : (K3.tab === 'insiden' ? 'Catat Insiden' : 'Catat Inspeksi')) + '</button>' : '<button class="btn btn-light-x" data-act="k3Muat"><i class="bi bi-arrow-clockwise me-1"></i>Muat ulang</button>');
  if (K3.tab === 'ringkasan') k3Ringkasan(); else if (K3.tab === 'inspeksi') k3TabInspeksi(); else if (K3.tab === 'insiden') k3TabInsiden(); else k3TabSop();
}

/* ── Ringkasan ── */
function k3Ringkasan() {
  var I = K3.inspeksi, N = K3.insiden, S0 = K3.sop;
  var apd = I.filter(function (r) { return r.Kepatuhan_APD !== ''; }), rata = apd.length ? Math.round(apd.reduce(function (s, r) { return s + Number(r.Kepatuhan_APD); }, 0) / apd.length) : null;
  var tl = I.concat(N).filter(function (r) { return r.Status_Tindak_Lanjut; }), tlSel = tl.filter(function (r) { return r.Status_Tindak_Lanjut === 'Selesai'; }).length;
  var telat = tl.filter(k3Telat).sort(function (a, b) { return a.Due_Date < b.Due_Date ? -1 : 1; });
  var kotak = function (id, judul, ket) { return '<div class="card-x"><div class="card-h">' + judul + (ket ? '<span class="badge-x ms-auto">' + ket + '</span>' : '') + '</div><div class="card-b"><div class="chart-box"><canvas id="' + id + '" role="img" aria-label="' + esc(judul) + '"></canvas></div></div></div>'; };
  $('#k3Body').innerHTML = '<div class="kpi-grid">' +
    kpiHtml('bi-bandaid', 'bad', fmtAngka(N.filter(function (r) { return r.Jenis_Insiden === 'Kecelakaan Kerja'; }).length), 'Total Kecelakaan Kerja') +
    kpiHtml('bi-exclamation-triangle', 'warn', fmtAngka(I.filter(function (r) { return r.Unsafe_Condition === 'Ya'; }).length), 'Total Unsafe Condition') +
    kpiHtml('bi-clipboard-check', '', fmtAngka(I.length), 'Total Inspeksi K3') +
    kpiHtml('bi-person-check', 'ok', rata === null ? '—' : rata + '%', 'Kepatuhan APD rata-rata') +
    kpiHtml('bi-journal-check', '', fmtAngka(S0.filter(function (r) { return r.Status === 'Berlaku'; }).length), 'Total SOP K3 (berlaku)') +
    kpiHtml('bi-check2-circle', 'ok', fmtAngka(tlSel) + ' / ' + fmtAngka(tl.length), 'Tindak Lanjut Selesai') + '</div>' +
    ((I.length + N.length) ? '<div class="grid-2">' + kotak('gk3Tren', 'Tren Insiden & Temuan K3', 'per bulan') + kotak('gk3Bahaya', 'Temuan per Kategori Bahaya') + '</div>' +
      '<div class="grid-2">' + kotak('gk3Tl', 'Status Tindak Lanjut') + kotak('gk3Apd', 'Kepatuhan APD', 'rata-rata % per bulan') + '</div>' : '<div class="card-x"><div class="empty"><i class="bi bi-shield-check"></i>Belum ada inspeksi atau insiden K3 tercatat.</div></div>') +
    '<div class="card-x"><div class="card-h"><i class="bi bi-alarm"></i>Tindak Lanjut Melewati Due Date<span class="badge-x ms-auto">' + telat.length + '</span></div><div class="card-b">' +
    (telat.length ? telat.slice(0, 8).map(function (r) {
      var insp = !!r.ID_Inspeksi;
      return '<button type="button" class="riw-card" data-act="' + (insp ? 'k3DetailI' : 'k3DetailN') + '" data-id="' + esc(insp ? r.ID_Inspeksi : r.ID_Insiden) + '"><div class="d-flex flex-wrap gap-1 align-items-center"><b>' + esc(insp ? r.ID_Inspeksi : r.ID_Insiden) + '</b>' + bTlK3(r.Status_Tindak_Lanjut) + badge('Due ' + fmtTgl(r.Due_Date, true), 'bad') + '</div><div class="riw-uraian">' + esc(insp ? r.Potensi_Bahaya : r.Jenis_Insiden + ' — ' + r.Kronologi.slice(0, 80)) + '</div><div class="riw-meta">PIC: ' + esc(r.PIC_Tindak_Lanjut) + '<i class="bi bi-chevron-right ms-auto"></i></div></button>';
    }).join('') : '<div class="small text-muted">Tidak ada tindak lanjut yang terlambat.</div>') + '</div></div>';
  if (!(I.length + N.length) || typeof Chart === 'undefined') return;
  var gelap = S.theme === 'dark', grid = gelap ? '#22304D' : '#E2E8F0';
  Chart.defaults.color = gelap ? '#98A5BF' : '#6B7280'; Chart.defaults.font.family = "'Plus Jakarta Sans',sans-serif";
  var temuan = I.filter(function (r) { return r.Temuan || r.Unsafe_Condition === 'Ya' || r.Unsafe_Action === 'Ya'; });
  var bln = uniq(N.map(function (r) { return k3Bulan(r.Tanggal); }).concat(temuan.map(function (r) { return k3Bulan(r.Tanggal_Inspeksi); }))).sort();
  S.charts.push(new Chart($('#gk3Tren'), { type: 'bar', data: { labels: bln.map(k3LabelBulan), datasets: [
    { label: 'Insiden', backgroundColor: '#DC2626', borderRadius: 4, data: bln.map(function (b) { return N.filter(function (r) { return k3Bulan(r.Tanggal) === b; }).length; }) },
    { label: 'Temuan inspeksi', backgroundColor: '#F59E0B', borderRadius: 4, data: bln.map(function (b) { return temuan.filter(function (r) { return k3Bulan(r.Tanggal_Inspeksi) === b; }).length; }) }] },
    options: { maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, scales: { y: { ticks: { precision: 0 }, grid: { color: grid } }, x: { grid: { display: false } } } } }));
  var kb = hitung(temuan, 'Kategori_Bahaya');
  S.charts.push(new Chart($('#gk3Bahaya'), { type: 'bar', data: { labels: kb.map(function (x) { return x[0]; }), datasets: [{ label: 'Temuan', data: kb.map(function (x) { return x[1]; }), backgroundColor: '#1E5EFF', borderRadius: 4 }] },
    options: { indexAxis: 'y', maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { precision: 0 }, grid: { color: grid } }, y: { grid: { display: false } } } } }));
  var st = hitung(tl, 'Status_Tindak_Lanjut'), wst = { 'Belum Dimulai': '#94A3B8', 'Dalam Proses': '#F59E0B', 'Selesai': '#16A34A' };
  S.charts.push(new Chart($('#gk3Tl'), { type: 'doughnut', data: { labels: st.map(function (x) { return x[0]; }), datasets: [{ data: st.map(function (x) { return x[1]; }), backgroundColor: st.map(function (x, i) { return wst[x[0]] || PALET[i % 8]; }), borderWidth: 0 }] }, options: { maintainAspectRatio: false, cutout: '62%', plugins: { legend: { position: 'bottom' } } } }));
  var ba = uniq(apd.map(function (r) { return k3Bulan(r.Tanggal_Inspeksi); })).sort();
  S.charts.push(new Chart($('#gk3Apd'), { type: 'line', data: { labels: ba.map(k3LabelBulan), datasets: [{ label: 'Kepatuhan APD (%)', borderColor: '#16A34A', backgroundColor: '#16A34A', tension: .25, data: ba.map(function (b) { var l = apd.filter(function (r) { return k3Bulan(r.Tanggal_Inspeksi) === b; }); return Math.round(l.reduce(function (s, r) { return s + Number(r.Kepatuhan_APD); }, 0) / l.length); }) }] },
    options: { maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { min: 0, max: 100, grid: { color: grid } }, x: { grid: { display: false } } } } }));
}

/* ── Tabel generik: pager ── */
function k3Pager(id, f, total, aksi) {
  var hal = Math.max(1, Math.ceil(total / K3.per)); if (f.hal > hal) f.hal = hal; if (f.hal < 1) f.hal = 1;
  var mulai = (f.hal - 1) * K3.per;
  $(id).innerHTML = '<div>' + (total ? (mulai + 1) + '–' + Math.min(total, mulai + K3.per) + ' dari ' + fmtAngka(total) + ' catatan' : '0 catatan') + '</div><div class="d-flex align-items-center gap-2"><button class="btn btn-light-x btn-sm" data-act="' + aksi + '" data-d="-1"' + (f.hal <= 1 ? ' disabled' : '') + '><i class="bi bi-chevron-left"></i></button><span>Hal. ' + f.hal + ' / ' + hal + '</span><button class="btn btn-light-x btn-sm" data-act="' + aksi + '" data-d="1"' + (f.hal >= hal ? ' disabled' : '') + '><i class="bi bi-chevron-right"></i></button></div>';
  return mulai;
}
ACTIONS.k3Hal = function (el) { var f = K3.tab === 'inspeksi' ? K3.fi : (K3.tab === 'insiden' ? K3.fn : K3.fs); f.hal += Number(el.dataset.d); k3Tabel(); };
ACTIONS.k3Reset = function () { if (K3.tab === 'inspeksi') K3.fi = { q: '', bahaya: '', lv: '', tl: '', hal: 1 }; else if (K3.tab === 'insiden') K3.fn = { q: '', jenis: '', kep: '', tl: '', hal: 1 }; else K3.fs = { q: '', kat: '', st: '', hal: 1 }; k3Isi(); };
function k3Tabel() { if (K3.tab === 'inspeksi') k3IsiInspeksi(); else if (K3.tab === 'insiden') k3IsiInsiden(); else k3IsiSop(); }
document.addEventListener('input', debounce(function (e) { var t = e.target; if (S.page === 'k3' && t.dataset && t.dataset.k3f === 'q') { var f = K3.tab === 'inspeksi' ? K3.fi : (K3.tab === 'insiden' ? K3.fn : K3.fs); f.q = t.value; f.hal = 1; k3Tabel(); } }, 200));
document.addEventListener('change', function (e) { var t = e.target; if (S.page === 'k3' && t.dataset && t.dataset.k3f && t.dataset.k3f !== 'q') { var f = K3.tab === 'inspeksi' ? K3.fi : (K3.tab === 'insiden' ? K3.fn : K3.fs); f[t.dataset.k3f] = t.value; f.hal = 1; k3Tabel(); } });
function k3Cari(f, kolom) { return function (r) { return !f.q.trim() || cocokCari(f.q, kolom.map(function (k) { return r[k]; }).join(' ')); }; }
function k3TlTeks(r) { return r.Status_Tindak_Lanjut ? bTlK3(r.Status_Tindak_Lanjut) + '<div class="sub">PIC ' + esc(r.PIC_Tindak_Lanjut) + ' · Due ' + esc(fmtTgl(r.Due_Date, true)) + '</div>' + (k3Telat(r) ? '<div class="sub">' + badge('Terlambat', 'bad') + '</div>' : '') : bTlK3(''); }

/* ── Tab Inspeksi ── */
function k3TabInspeksi() {
  var f = K3.fi;
  $('#k3Body').innerHTML = '<div class="card-x"><div class="toolbar"><div class="grow input-icon"><i class="bi bi-search"></i><input class="form-control" data-k3f="q" placeholder="Cari ID, aktivitas, bahaya, temuan…" value="' + esc(f.q) + '" aria-label="Cari inspeksi"></div>' +
    k3Sel(listMaster('Kategori Bahaya'), 'bahaya', f.bahaya, 'Semua kategori bahaya') + k3Sel(RK.level.map(function (x) { return x.nama; }), 'lv', f.lv, 'Semua level') + k3Sel(k3ListTl(), 'tl', f.tl, 'Semua status tindak lanjut') +
    '<button class="btn btn-light-x btn-sm" data-act="k3Reset">Reset</button></div><div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Inspeksi</th><th>Bahaya</th><th>Skor (L × I)</th><th>Unsafe</th><th>Kepatuhan APD</th><th>Tindak lanjut</th><th></th></tr></thead><tbody id="k3Rows"></tbody></table></div><div class="pager" id="k3Pager"></div></div>';
  k3IsiInspeksi();
}
function k3IsiInspeksi() {
  var f = K3.fi, cari = k3Cari(f, ['ID_Inspeksi', 'Aktivitas', 'Potensi_Bahaya', 'Temuan', 'Nama_Pelaksana']);
  var data = K3.inspeksi.filter(function (r) { return cari(r) && (!f.bahaya || r.Kategori_Bahaya === f.bahaya) && (!f.lv || rkNamaLevel(Number(r.Skor_Risiko)) === f.lv) && (!f.tl || r.Status_Tindak_Lanjut === f.tl); })
    .sort(function (a, b) { return (b.Tanggal_Inspeksi + b.Waktu_Inspeksi) < (a.Tanggal_Inspeksi + a.Waktu_Inspeksi) ? -1 : 1; });
  var mulai = k3Pager('#k3Pager', f, data.length, 'k3Hal');
  $('#k3Rows').innerHTML = data.length ? data.slice(mulai, mulai + K3.per).map(function (r) {
    return '<tr><td data-l="Inspeksi"><div class="nm">' + esc(r.ID_Inspeksi) + '</div><div class="sub">' + esc(fmtTgl(r.Tanggal_Inspeksi, true)) + ' ' + esc(r.Waktu_Inspeksi) + ' · ' + esc(r.Nama_Pelaksana) + '</div><div class="sub">' + esc(r.Aktivitas.slice(0, 70)) + '</div></td>' +
      '<td data-l="Bahaya"><div class="nm">' + esc(r.Kategori_Bahaya) + '</div><div class="sub">' + esc(r.Potensi_Bahaya.slice(0, 80)) + '</div></td>' +
      '<td data-l="Skor">' + bSkorK3(r) + '<div class="sub">L' + esc(r.Likelihood) + ' × I' + esc(r.Impact) + '</div></td>' +
      '<td data-l="Unsafe"><div class="sub">Condition ' + bYaTidak(r.Unsafe_Condition, true) + '</div><div class="sub">Action ' + bYaTidak(r.Unsafe_Action, true) + '</div></td>' +
      '<td data-l="APD">' + (r.Kepatuhan_APD === '' ? '<span class="text-muted">—</span>' : '<div class="nm">' + esc(r.Kepatuhan_APD) + '%</div>') + '</td>' +
      '<td data-l="Tindak lanjut">' + k3TlTeks(r) + '</td>' +
      '<td class="acts"><button class="icon-btn" data-act="k3DetailI" data-id="' + esc(r.ID_Inspeksi) + '" title="Detail" aria-label="Detail ' + esc(r.ID_Inspeksi) + '"><i class="bi bi-eye"></i></button><button class="icon-btn" data-act="k3UbahI" data-id="' + esc(r.ID_Inspeksi) + '" title="Ubah" aria-label="Ubah ' + esc(r.ID_Inspeksi) + '"><i class="bi bi-pencil"></i></button></td></tr>';
  }).join('') : '<tr><td colspan="7"><div class="empty"><i class="bi bi-clipboard-check"></i>' + (K3.inspeksi.length ? 'Tidak ada inspeksi yang cocok dengan filter.' : 'Belum ada inspeksi K3 tercatat.') + '</div></td></tr>';
}

/* ── Tab Insiden ── */
function k3TabInsiden() {
  var f = K3.fn;
  $('#k3Body').innerHTML = '<div class="card-x"><div class="toolbar"><div class="grow input-icon"><i class="bi bi-search"></i><input class="form-control" data-k3f="q" placeholder="Cari ID, kronologi, korban, lokasi…" value="' + esc(f.q) + '" aria-label="Cari insiden"></div>' +
    k3Sel(listMaster('Jenis Insiden'), 'jenis', f.jenis, 'Semua jenis') + k3Sel(listMaster('Tingkat Keparahan'), 'kep', f.kep, 'Semua keparahan') + k3Sel(k3ListTl(), 'tl', f.tl, 'Semua status tindak lanjut') +
    '<button class="btn btn-light-x btn-sm" data-act="k3Reset">Reset</button></div><div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Insiden</th><th>Lokasi</th><th>Jenis &amp; keparahan</th><th>Korban / pihak terlibat</th><th>Tindak lanjut</th><th></th></tr></thead><tbody id="k3Rows"></tbody></table></div><div class="pager" id="k3Pager"></div></div>';
  k3IsiInsiden();
}
function k3IsiInsiden() {
  var f = K3.fn, cari = k3Cari(f, ['ID_Insiden', 'Kronologi', 'Korban_Terlibat', 'Gedung', 'Ruangan', 'Penyebab']);
  var data = K3.insiden.filter(function (r) { return cari(r) && (!f.jenis || r.Jenis_Insiden === f.jenis) && (!f.kep || r.Tingkat_Keparahan === f.kep) && (!f.tl || r.Status_Tindak_Lanjut === f.tl); })
    .sort(function (a, b) { return (b.Tanggal + b.Waktu) < (a.Tanggal + a.Waktu) ? -1 : 1; });
  var mulai = k3Pager('#k3Pager', f, data.length, 'k3Hal');
  $('#k3Rows').innerHTML = data.length ? data.slice(mulai, mulai + K3.per).map(function (r) {
    return '<tr><td data-l="Insiden"><div class="nm">' + esc(r.ID_Insiden) + '</div><div class="sub">' + esc(fmtTgl(r.Tanggal, true)) + ' ' + esc(r.Waktu) + '</div><div class="sub">' + esc(r.Kronologi.slice(0, 80)) + '</div></td>' +
      '<td data-l="Lokasi"><div class="nm">' + esc(r.Ruangan) + '</div><div class="sub">' + esc([r.Gedung, r.Lantai].filter(Boolean).join(' · ')) + '</div></td>' +
      '<td data-l="Jenis">' + badge(r.Jenis_Insiden, r.Jenis_Insiden === 'Kecelakaan Kerja' ? 'bad' : 'warn') + '<div class="sub">' + badge(r.Tingkat_Keparahan, { 'Ringan': '', 'Sedang': 'warn', 'Berat': 'bad', 'Fatal': 'bad' }[r.Tingkat_Keparahan] || '') + '</div></td>' +
      '<td data-l="Korban">' + (esc(r.Korban_Terlibat) || '<span class="text-muted">—</span>') + '</td>' +
      '<td data-l="Tindak lanjut">' + k3TlTeks(r) + '</td>' +
      '<td class="acts"><button class="icon-btn" data-act="k3DetailN" data-id="' + esc(r.ID_Insiden) + '" title="Detail" aria-label="Detail ' + esc(r.ID_Insiden) + '"><i class="bi bi-eye"></i></button><button class="icon-btn" data-act="k3UbahN" data-id="' + esc(r.ID_Insiden) + '" title="Ubah" aria-label="Ubah ' + esc(r.ID_Insiden) + '"><i class="bi bi-pencil"></i></button></td></tr>';
  }).join('') : '<tr><td colspan="6"><div class="empty"><i class="bi bi-bandaid"></i>' + (K3.insiden.length ? 'Tidak ada insiden yang cocok dengan filter.' : 'Belum ada insiden tercatat.') + '</div></td></tr>';
}

/* ── Tab SOP ── */
function k3TabSop() {
  var f = K3.fs;
  $('#k3Body').innerHTML = '<div class="card-x"><div class="toolbar"><div class="grow input-icon"><i class="bi bi-search"></i><input class="form-control" data-k3f="q" placeholder="Cari nomor atau judul SOP…" value="' + esc(f.q) + '" aria-label="Cari SOP"></div>' +
    k3Sel(listMaster('Kategori SOP'), 'kat', f.kat, 'Semua kategori') + k3Sel(K3_STATUS_SOP, 'st', f.st, 'Semua status') +
    '<button class="btn btn-light-x btn-sm" data-act="k3Reset">Reset</button></div><div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>SOP</th><th>Kategori</th><th>Berlaku</th><th>Review</th><th>Status</th><th>File</th><th></th></tr></thead><tbody id="k3Rows"></tbody></table></div><div class="pager" id="k3Pager"></div></div>';
  k3IsiSop();
}
function sopPerluReview(r) { return r.Status === 'Berlaku' && r.Tanggal_Review && r.Tanggal_Review < hariIniLokal(); }
function k3IsiSop() {
  var f = K3.fs, cari = k3Cari(f, ['Nomor_SOP', 'Judul_SOP']);
  var data = K3.sop.filter(function (r) { return cari(r) && (!f.kat || r.Kategori === f.kat) && (!f.st || r.Status === f.st); }).sort(function (a, b) { return a.Nomor_SOP < b.Nomor_SOP ? -1 : 1; });
  var mulai = k3Pager('#k3Pager', f, data.length, 'k3Hal');
  $('#k3Rows').innerHTML = data.length ? data.slice(mulai, mulai + K3.per).map(function (r) {
    return '<tr><td data-l="SOP"><div class="nm">' + esc(r.Nomor_SOP) + '</div><div class="sub">' + esc(r.Judul_SOP) + '</div></td><td data-l="Kategori">' + esc(r.Kategori) + '</td>' +
      '<td data-l="Berlaku">' + esc(fmtTgl(r.Tanggal_Berlaku, true)) + '</td>' +
      '<td data-l="Review">' + esc(fmtTgl(r.Tanggal_Review, true)) + (sopPerluReview(r) ? '<div class="sub">' + badge('Perlu review', 'warn') + '</div>' : '') + '</td>' +
      '<td data-l="Status">' + badge(r.Status, { 'Berlaku': 'ok', 'Dalam Revisi': 'warn', 'Tidak Berlaku': '' }[r.Status] || '') + '</td>' +
      '<td data-l="File">' + (r.File_URL ? '<a href="' + esc(r.File_URL) + '" target="_blank" rel="noopener noreferrer"><i class="bi bi-file-earmark-pdf me-1"></i>Unduh</a>' : '<span class="text-muted">—</span>') + '</td>' +
      '<td class="acts"><button class="icon-btn" data-act="k3UbahS" data-id="' + esc(r.ID_SOP) + '" title="Ubah" aria-label="Ubah ' + esc(r.Nomor_SOP) + '"><i class="bi bi-pencil"></i></button><button class="icon-btn text-danger" data-act="k3HapusS" data-id="' + esc(r.ID_SOP) + '" title="Hapus" aria-label="Hapus ' + esc(r.Nomor_SOP) + '"><i class="bi bi-trash"></i></button></td></tr>';
  }).join('') : '<tr><td colspan="7"><div class="empty"><i class="bi bi-journal-text"></i>' + (K3.sop.length ? 'Tidak ada SOP yang cocok dengan filter.' : 'Belum ada SOP K3 tercatat.') + '</div></td></tr>';
}
ACTIONS.k3HapusS = async function (el) {
  var id = el.dataset.id, r = K3.sop.filter(function (x) { return x.ID_SOP === id; })[0]; if (!r) return;
  if (!(await konfirm('Hapus SOP?', r.Nomor_SOP + ' — ' + r.Judul_SOP + ' akan dihapus permanen.', 'Hapus', true))) return;
  try { await call('deleteSop', id); K3.sop = K3.sop.filter(function (x) { return x.ID_SOP !== id; }); toast('SOP dihapus.', 'ok'); k3IsiSop(); } catch (e) { gagal(e); }
};

/* ── Ekspor ── */
ACTIONS.k3Ekspor = async function (el) {
  var peta = { inspeksi: ['Inspeksi_K3', K3.inspeksi, 'K3_Inspeksi'], insiden: ['Insiden_K3', K3.insiden, 'K3_Insiden'], sop: ['SOP_K3', K3.sop, 'K3_SOP'] }[K3.tab];
  if (!peta[1].length) { toast('Tidak ada data untuk diekspor.', 'warn'); return; }
  var kolom = Object.keys(peta[1][0]);
  var tabel = [kolom].concat(peta[1].map(function (r) { return kolom.map(function (k) { return amanSel(r[k]); }); }));
  sibuk(el, true, 'Menyiapkan…');
  var nama = peta[0] + '_' + new Date().toISOString().slice(0, 10);
  try {
    await muatSkrip(CDN_XLSX, function () { return typeof XLSX !== 'undefined'; });
    var ws = XLSX.utils.aoa_to_sheet(tabel), wb = XLSX.utils.book_new();
    ws['!cols'] = kolom.map(function (k) { return { wch: Math.min(40, Math.max(k.length + 2, 14)) }; });
    XLSX.utils.book_append_sheet(wb, ws, peta[0].slice(0, 30));
    unduhBlob(nama + '.xlsx', new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    toast('Berkas Excel diunduh (' + peta[1].length + ' catatan).', 'ok');
  } catch (e) {
    unduhBlob(nama + '.csv', new Blob([csvKeString(tabel)], { type: 'text/csv;charset=utf-8' }));
    toast('Pustaka Excel tidak tersedia — diunduh sebagai CSV (dapat dibuka di Excel).', 'warn', 6000);
  } finally { sibuk(el, false); }
};

/* ── Detail & status tindak lanjut ── */
function k3FotoHtml(r) { var f = k3FotoArr(r); return f.length ? '<div class="foto-grid mt-2">' + f.map(function (u, i) { return fotoThumb(u, 'Foto ' + (i + 1)); }).join('') + '</div>' : ''; }
function k3TlBlok(r, jenis) {
  if (!r.Status_Tindak_Lanjut) return woSeksi('Tindak Lanjut', '<div class="small text-muted">Tidak ada temuan sehingga tidak ada tindak lanjut.</div>');
  var id = jenis === 'inspeksi' ? r.ID_Inspeksi : r.ID_Insiden;
  return woSeksi('Tindak Lanjut', '<dl class="dl">' + dlRow('Rekomendasi', teksPanjang(jenis === 'inspeksi' ? r.Rekomendasi_Tindakan : r.Rekomendasi)) + dlRow('PIC', esc(r.PIC_Tindak_Lanjut)) +
    dlRow('Due Date', esc(fmtTgl(r.Due_Date, true)) + (k3Telat(r) ? ' ' + badge('Terlambat', 'bad') : '')) + dlRow('Status', bTlK3(r.Status_Tindak_Lanjut)) + '</dl>' +
    '<div class="d-flex flex-wrap gap-2 mt-2" role="group" aria-label="Ubah status tindak lanjut">' + k3ListTl().map(function (x) {
      return '<button type="button" class="btn btn-sm ' + (x === r.Status_Tindak_Lanjut ? 'btn-primary' : 'btn-light-x') + '" data-act="k3Status" data-j="' + jenis + '" data-id="' + esc(id) + '" data-v="' + esc(x) + '"' + (x === r.Status_Tindak_Lanjut ? ' disabled' : '') + '>' + esc(x) + '</button>';
    }).join('') + '</div>');
}
function detailInspeksiHtml(r) {
  return woSeksi('Identitas', '<dl class="dl">' + dlRow('ID Inspeksi', '<b>' + esc(r.ID_Inspeksi) + '</b>') + dlRow('Tanggal & waktu', esc(fmtTgl(r.Tanggal_Inspeksi, true)) + ' ' + esc(r.Waktu_Inspeksi)) + dlRow('Pelaksana', esc(r.Pelaksana) + ' — ' + esc(r.Nama_Pelaksana)) + dlRow('Aktivitas', teksPanjang(r.Aktivitas)) + dlRow('Dicatat', esc(r.Dibuat_Oleh) + ' · ' + esc(fmtTgl(r.Dibuat_Pada))) + '</dl>') +
    woSeksi('Penilaian Bahaya', '<dl class="dl">' + dlRow('Kategori bahaya', esc(r.Kategori_Bahaya)) + dlRow('Potensi bahaya', teksPanjang(r.Potensi_Bahaya)) + dlRow('Likelihood', esc(r.Likelihood) + ' — ' + esc(RK_L[Number(r.Likelihood)] || '')) + dlRow('Impact', esc(r.Impact) + ' — ' + esc(RK_I[Number(r.Impact)] || '')) + dlRow('Skor risiko', bSkorK3(r)) + '</dl>') +
    woSeksi('Kondisi & APD', '<dl class="dl">' + dlRow('Unsafe Condition', bYaTidak(r.Unsafe_Condition, true)) + dlRow('Unsafe Action', bYaTidak(r.Unsafe_Action, true)) + dlRow('APD wajib', esc(r.APD_Wajib)) + dlRow('Kepatuhan APD', r.Kepatuhan_APD === '' ? '' : '<b>' + esc(r.Kepatuhan_APD) + '%</b>') + '</dl>') +
    woSeksi('Temuan', '<dl class="dl">' + dlRow('Temuan', teksPanjang(r.Temuan)) + '</dl>' + k3FotoHtml(r)) + k3TlBlok(r, 'inspeksi');
}
function detailInsidenHtml(r) {
  return woSeksi('Identitas', '<dl class="dl">' + dlRow('ID Insiden', '<b>' + esc(r.ID_Insiden) + '</b>') + dlRow('Tanggal & waktu', esc(fmtTgl(r.Tanggal, true)) + ' ' + esc(r.Waktu)) + dlRow('Jenis', badge(r.Jenis_Insiden, r.Jenis_Insiden === 'Kecelakaan Kerja' ? 'bad' : 'warn')) + dlRow('Tingkat keparahan', esc(r.Tingkat_Keparahan)) + dlRow('Dicatat', esc(r.Dibuat_Oleh) + ' · ' + esc(fmtTgl(r.Dibuat_Pada))) + '</dl>') +
    woSeksi('Lokasi', '<dl class="dl">' + dlRow('Gedung', esc(r.Gedung)) + dlRow('Lantai', esc(r.Lantai)) + dlRow('Ruangan', esc(r.Ruangan)) + dlRow('Lokasi Detail', esc(r.Lokasi_Detail)) + '</dl>') +
    woSeksi('Uraian Insiden', '<dl class="dl">' + dlRow('Kronologi', teksPanjang(r.Kronologi)) + dlRow('Korban / pihak terlibat', esc(r.Korban_Terlibat)) + dlRow('Penyebab', teksPanjang(r.Penyebab)) + dlRow('Tindakan segera', teksPanjang(r.Tindakan_Segera)) + '</dl>' + k3FotoHtml(r)) + k3TlBlok(r, 'insiden');
}
function k3BukaDetail(jenis, id) {
  var arr = K3[jenis === 'inspeksi' ? 'inspeksi' : 'insiden'], kunci = jenis === 'inspeksi' ? 'ID_Inspeksi' : 'ID_Insiden';
  var r = arr.filter(function (x) { return x[kunci] === id; })[0]; if (!r) return;
  bukaModal({ judul: (jenis === 'inspeksi' ? 'Inspeksi K3 ' : 'Insiden ') + id, size: 'lg', isi: jenis === 'inspeksi' ? detailInspeksiHtml(r) : detailInsidenHtml(r),
    footer: '<button class="btn btn-light-x text-danger me-auto" data-act="k3Hapus" data-j="' + jenis + '" data-id="' + esc(id) + '"><i class="bi bi-trash me-1"></i>Hapus</button><button class="btn btn-light-x" data-act="' + (jenis === 'inspeksi' ? 'k3UbahI' : 'k3UbahN') + '" data-id="' + esc(id) + '"><i class="bi bi-pencil me-1"></i>Ubah</button><button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>' });
}
ACTIONS.k3DetailI = function (el) { if (document.querySelector('#mdl.show')) tutupModal(); setTimeout(function () { k3BukaDetail('inspeksi', el.dataset.id); }, document.querySelector('#mdl.show') ? 250 : 0); };
ACTIONS.k3DetailN = function (el) { if (document.querySelector('#mdl.show')) tutupModal(); setTimeout(function () { k3BukaDetail('insiden', el.dataset.id); }, document.querySelector('#mdl.show') ? 250 : 0); };
ACTIONS.k3Status = async function (el) {
  var j = el.dataset.j, id = el.dataset.id;
  try {
    var d = await call('updateTindakLanjut', j, id, el.dataset.v); k3Ganti(j, d[j]); toast('Status tindak lanjut diperbarui.', 'ok');
    var r = d[j]; $('#mdlContent .modal-body').innerHTML = j === 'inspeksi' ? detailInspeksiHtml(r) : detailInsidenHtml(r);
    if (S.page === 'k3') k3Isi();
  } catch (e) { gagal(e); }
};
ACTIONS.k3Hapus = async function (el) {
  var j = el.dataset.j, id = el.dataset.id;
  if (!(await konfirm('Hapus catatan?', id + ' akan dihapus permanen.', 'Hapus', true))) return;
  try {
    await call(j === 'inspeksi' ? 'deleteInspeksi' : 'deleteInsiden', id);
    K3[j] = K3[j].filter(function (x) { return (x.ID_Inspeksi || x.ID_Insiden) !== id; }); tutupModal(); toast('Catatan dihapus.', 'ok'); if (S.page === 'k3') k3Isi();
  } catch (e) { gagal(e); }
};

/* ── Foto (maks 5) & PDF dalam form ── */
function k3FotoUi() {
  var el = $('#k3FotoList'); if (!el || !K3F) return;
  var lama = K3F.foto.map(function (u, i) { return '<div class="k3-th"><img src="' + esc(imgU(u)) + '" alt="Foto ' + (i + 1) + '" referrerpolicy="no-referrer"><button type="button" class="k3-x" data-act="k3FotoHapus" data-t="lama" data-i="' + i + '" aria-label="Hapus foto ' + (i + 1) + '"><i class="bi bi-x"></i></button></div>'; }).join('');
  var baru = K3F.baru.map(function (b, i) { return '<div class="k3-th"><img src="' + b.preview + '" alt="Foto baru ' + (i + 1) + '"><button type="button" class="k3-x" data-act="k3FotoHapus" data-t="baru" data-i="' + i + '" aria-label="Hapus foto baru ' + (i + 1) + '"><i class="bi bi-x"></i></button></div>'; }).join('');
  el.innerHTML = lama + baru; var n = K3F.foto.length + K3F.baru.length;
  var info = $('#k3FotoInfo'); if (info) info.textContent = n + ' / 5 foto';
}
ACTIONS.k3FotoHapus = function (el) { var i = Number(el.dataset.i); (el.dataset.t === 'lama' ? K3F.foto : K3F.baru).splice(i, 1); k3FotoUi(); };
function k3FotoBlok() { return '<div class="col-12"><label class="form-label">Foto <span class="text-muted fw-normal" id="k3FotoInfo"></span></label><input type="file" class="form-control" id="k3FotoIn" accept="image/*" multiple><div class="form-text">Maksimal 5 foto, Foto dikompres otomatis.</div><div class="k3-fotos mt-2" id="k3FotoList"></div></div>'; }
document.addEventListener('change', async function (e) {
  var t = e.target; if (!K3F || !t.closest || !t.closest('#fK3')) return;
  if (t.id === 'k3FotoIn' && t.files && t.files.length) {
    var files = Array.prototype.slice.call(t.files);
    for (var i = 0; i < files.length; i++) {
      if (K3F.foto.length + K3F.baru.length >= 5) { toast('Maksimal 5 foto.', 'warn'); break; }
      try { K3F.baru.push(await kompresFoto(files[i])); } catch (er) { toast(er.message, 'bad'); }
    }
    t.value = ''; k3FotoUi();
  }
  if (t.id === 'k3PdfIn' && t.files && t.files[0]) {
    var fl = t.files[0];
    if (fl.type !== 'application/pdf') { toast('File SOP harus berformat PDF.', 'bad'); t.value = ''; return; }
    if (fl.size > 2 * 1024 * 1024) { toast('Ukuran PDF maksimal 2 MB (ukuran ' + (fl.size / 1048576).toFixed(1).replace('.', ',') + ' MB).', 'bad'); t.value = ''; return; }
    K3F.pdf = fl; K3F.hapusFile = false; $('#k3PdfInfo').innerHTML = '<i class="bi bi-file-earmark-pdf me-1"></i>' + esc(fl.name) + ' (siap diunggah)';
  }
});
async function k3UnggahFoto(owner) {
  var urls = K3F.foto.slice();
  for (var i = 0; i < K3F.baru.length; i++) { var b = K3F.baru[i]; var u = await call('uploadFile', 'k3', owner, { name: b.name, mime: b.mime, base64: b.base64 }); urls.push(u.url); }
  return urls;
}

/* ── Form Inspeksi ── */
function k3Skala(arr, sel) { var h = '<option value="">— pilih —</option>'; for (var i = 1; i <= 5; i++) h += '<option value="' + i + '"' + (String(sel) === String(i) ? ' selected' : '') + '>' + i + ' — ' + esc(arr[i]) + '</option>'; return h; }
function k3YT(sel) { return '<option value="">— pilih —</option><option value="Ya"' + (sel === 'Ya' ? ' selected' : '') + '>Ya</option><option value="Tidak"' + (sel === 'Tidak' ? ' selected' : '') + '>Tidak</option>'; }
function k3TlForm(r, jenis, wajib) {
  var ins = jenis === 'insiden';
  return '<div id="k3TlBox" class="col-12' + (wajib ? '' : ' d-none') + '"><div class="fw-bold small text-muted text-uppercase mb-2">Tindak lanjut</div><div class="row g-3">' +
    '<div class="col-12"><label class="form-label req">Rekomendasi tindakan</label><textarea class="form-control" name="Rekomendasi" rows="2" maxlength="1000">' + esc(r && (ins ? r.Rekomendasi : r.Rekomendasi_Tindakan)) + '</textarea></div>' +
    '<div class="col-md-4"><label class="form-label req">PIC tindak lanjut</label><input class="form-control" name="PIC_Tindak_Lanjut" maxlength="100" value="' + esc(r && r.PIC_Tindak_Lanjut) + '"></div>' +
    '<div class="col-md-4"><label class="form-label req">Due Date</label><input type="date" class="form-control" name="Due_Date" value="' + esc(r && r.Due_Date) + '"></div>' +
    '<div class="col-md-4"><label class="form-label req">Status</label><select class="form-select" name="Status_Tindak_Lanjut">' + opsi(k3ListTl(), (r && r.Status_Tindak_Lanjut) || 'Belum Dimulai', null) + '</select></div></div></div>';
}
function formInspeksiHtml(r) {
  var apd = (r && r.APD_Wajib ? r.APD_Wajib.split(/\s*;\s*/) : []);
  var adaTemuan = !!(r && (r.Temuan || r.Unsafe_Condition === 'Ya' || r.Unsafe_Action === 'Ya'));
  return '<form id="fK3" data-submit="simpanInspeksi" novalidate><div class="row g-3">' +
    '<div class="col-md-4"><label class="form-label req">Tanggal Inspeksi</label><input type="date" class="form-control" name="Tanggal_Inspeksi" max="' + hariIniLokal() + '" value="' + esc((r && r.Tanggal_Inspeksi) || hariIniLokal()) + '"></div>' +
    '<div class="col-md-4"><label class="form-label req">Waktu Inspeksi</label><input type="time" class="form-control" name="Waktu_Inspeksi" value="' + esc((r && r.Waktu_Inspeksi) || jamLokal()) + '"></div>' +
    '<div class="col-md-4"><label class="form-label req">Pelaksana</label><select class="form-select" name="Pelaksana">' + opsi(listMaster('Pelaksana'), (r && r.Pelaksana) || 'Internal', null) + '</select></div>' +
    '<div class="col-md-6"><label class="form-label req">Nama Pelaksana</label><input class="form-control" name="Nama_Pelaksana" maxlength="150" list="k3Nama" value="' + esc(r && r.Nama_Pelaksana) + '"><datalist id="k3Nama"></datalist><div class="form-text" id="k3NamaInfo"></div></div>' +
    '<div class="col-md-6"><label class="form-label req">Kategori Bahaya</label><select class="form-select" name="Kategori_Bahaya">' + opsi(listMaster('Kategori Bahaya'), r && r.Kategori_Bahaya, '— pilih —') + '</select></div>' +
    '<div class="col-12"><label class="form-label req">Aktivitas yang diinspeksi</label><textarea class="form-control" name="Aktivitas" rows="2" maxlength="500">' + esc(r && r.Aktivitas) + '</textarea></div>' +
    '<div class="col-12"><label class="form-label req">Potensi Bahaya</label><textarea class="form-control" name="Potensi_Bahaya" rows="2" maxlength="500">' + esc(r && r.Potensi_Bahaya) + '</textarea></div>' +
    '<div class="col-md-4"><label class="form-label req">Likelihood (1–5)</label><select class="form-select" name="Likelihood">' + k3Skala(RK_L, r && r.Likelihood) + '</select></div>' +
    '<div class="col-md-4"><label class="form-label req">Impact (1–5)</label><select class="form-select" name="Impact">' + k3Skala(RK_I, r && r.Impact) + '</select></div>' +
    '<div class="col-md-4 d-flex align-items-end"><div class="rk-skor w-100" id="k3Skor" aria-live="polite"></div></div>' +
    '<div class="col-md-6"><label class="form-label req">Unsafe Condition</label><select class="form-select" name="Unsafe_Condition">' + k3YT(r && r.Unsafe_Condition) + '</select></div>' +
    '<div class="col-md-6"><label class="form-label req">Unsafe Action</label><select class="form-select" name="Unsafe_Action">' + k3YT(r && r.Unsafe_Action) + '</select></div>' +
    '<div class="col-12"><label class="form-label">APD wajib pada aktivitas ini</label><div class="chk-list">' + listMaster('APD Wajib').map(function (a) { return '<label class="chk-item"><input type="checkbox" name="apd" value="' + esc(a) + '"' + (apd.indexOf(a) >= 0 ? ' checked' : '') + '><span>' + esc(a) + '</span></label>'; }).join('') + '</div></div>' +
    '<div class="col-md-4"><label class="form-label" id="k3KpLbl">Kepatuhan APD (%)</label><input type="number" class="form-control" name="Kepatuhan_APD" min="0" max="100" step="1" value="' + esc(r && r.Kepatuhan_APD) + '" placeholder="0–100"><div class="form-text">Wajib diisi bila ada APD yang dipilih.</div></div>' +
    '<div class="col-12"><label class="form-label">Temuan</label><textarea class="form-control" name="Temuan" rows="2" maxlength="1000" placeholder="Wajib bila ada Unsafe Condition / Unsafe Action">' + esc(r && r.Temuan) + '</textarea></div>' +
    k3TlForm(r, 'inspeksi', adaTemuan) + k3FotoBlok() +
    '<div class="col-12"><div class="alert-x danger d-none mb-0" id="k3Err"></div></div></div></form>';
}
function k3HitungSkor() {
  var f = $('#fK3'), el = $('#k3Skor'); if (!f || !el || !f.Likelihood) return;
  var l = Number(f.Likelihood.value), i = Number(f.Impact.value);
  el.innerHTML = (!l || !i) ? '<span class="text-muted small">Skor = L × I</span>' : '<span class="small text-muted">Skor</span> <b>' + (l * i) + '</b> ' + badge(rkNamaLevel(l * i), RK_KLS[rkIdx(l * i)]);
}
function k3SesuaikanInspeksi() {
  var f = $('#fK3'); if (!f || !f.Unsafe_Condition) return;
  var ada = f.Temuan.value.trim() || f.Unsafe_Condition.value === 'Ya' || f.Unsafe_Action.value === 'Ya';
  $('#k3TlBox').classList.toggle('d-none', !ada);
  f.Kepatuhan_APD.disabled = !f.querySelector('[name=apd]:checked');
  var dl = $('#k3Nama'), vend = f.Pelaksana.value === 'Vendor';
  dl.innerHTML = listMaster(vend ? 'Nama Vendor' : 'Nama Teknisi').map(function (n) { return '<option value="' + esc(n) + '">'; }).join('');
  $('#k3NamaInfo').textContent = vend ? 'Vendor yang belum terdaftar akan diusulkan ke Master Admin.' : '';
}
function bukaFormK3(jenis, r, ubah) {
  K3F = { jenis: jenis, mode: ubah ? 'ubah' : 'baru', id: r && (r.ID_Inspeksi || r.ID_Insiden || r.ID_SOP), foto: r && r.Foto_URLs ? k3FotoArr(r) : [], baru: [], pdf: null, hapusFile: false };
  var judul = { inspeksi: 'Inspeksi K3', insiden: 'Insiden', sop: 'SOP K3' }[jenis];
  var isi = jenis === 'inspeksi' ? formInspeksiHtml(r) : (jenis === 'insiden' ? formInsidenHtml(r) : formSopHtml(r));
  bukaModal({ judul: (ubah ? 'Ubah ' + K3F.id : (jenis === 'sop' ? 'Tambah ' : 'Catat ') + judul), size: 'lg', isi: isi,
    footer: '<button class="btn btn-light-x" data-bs-dismiss="modal">Batal</button><button class="btn btn-primary" type="submit" form="fK3">' + (ubah ? 'Simpan perubahan' : 'Simpan') + '</button>' });
  if (jenis === 'inspeksi') { k3HitungSkor(); k3SesuaikanInspeksi(); }
  if (jenis !== 'sop') k3FotoUi();
}
ACTIONS.k3Baru = function () { bukaFormK3(K3.tab, null, false); };
function k3UbahKe(jenis, kunci, id) {
  var r = K3[jenis].filter(function (x) { return x[kunci] === id; })[0]; if (!r) return;
  var buka = document.querySelector('#mdl.show'); if (buka) tutupModal();
  setTimeout(function () { bukaFormK3(jenis, r, true); }, buka ? 250 : 0);
}
ACTIONS.k3UbahI = function (el) { k3UbahKe('inspeksi', 'ID_Inspeksi', el.dataset.id); };
ACTIONS.k3UbahN = function (el) { k3UbahKe('insiden', 'ID_Insiden', el.dataset.id); };
ACTIONS.k3UbahS = function (el) { k3UbahKe('sop', 'ID_SOP', el.dataset.id); };
document.addEventListener('change', function (e) {
  var t = e.target; if (!t.closest || !t.closest('#fK3') || !K3F) return;
  if (K3F.jenis === 'inspeksi') {
    if (t.name === 'Likelihood' || t.name === 'Impact') k3HitungSkor();
    k3SesuaikanInspeksi();
  }
  if (K3F.jenis === 'insiden' && t.dataset && t.dataset.ikLok !== undefined) {
    var f = $('#fK3'), i = Number(t.dataset.ikLok), nm = ['Jenis_Ruangan', 'Gedung', 'Lantai', 'Ruangan'];
    var pil = nm.map(function (n, k) { var v = f[n].value; if (k === 0) return ''; if (k === 2) return v === LANTAI_PILIH ? null : v; return v === '' ? null : v; });
    for (var j = i + 1; j < 4; j++) { pil[j] = null; f[nm[j]].innerHTML = lokOpsiHtml(j, pil, null); }
  }
});
document.addEventListener('input', function (e) {
  var t = e.target; if (!t.closest || !t.closest('#fK3') || !K3F || K3F.jenis !== 'inspeksi') return;
  if (t.name === 'Temuan') k3SesuaikanInspeksi();
});
SUBMITS.simpanInspeksi = async function (f) {
  var err = $('#k3Err'); err.classList.add('d-none');
  var salah = function (m) { err.textContent = m; err.classList.remove('d-none'); };
  var p = {};
  ['Tanggal_Inspeksi', 'Waktu_Inspeksi', 'Pelaksana', 'Nama_Pelaksana', 'Aktivitas', 'Kategori_Bahaya', 'Potensi_Bahaya', 'Likelihood', 'Impact', 'Unsafe_Condition', 'Unsafe_Action', 'Temuan'].forEach(function (k) { p[k] = f[k].value; });
  p.APD_Wajib = Array.prototype.map.call(f.querySelectorAll('[name=apd]:checked'), function (x) { return x.value; });
  p.Kepatuhan_APD = p.APD_Wajib.length ? f.Kepatuhan_APD.value : '';
  var ada = p.Temuan.trim() || p.Unsafe_Condition === 'Ya' || p.Unsafe_Action === 'Ya';
  if (ada) { p.Rekomendasi_Tindakan = f.Rekomendasi.value; p.PIC_Tindak_Lanjut = f.PIC_Tindak_Lanjut.value; p.Due_Date = f.Due_Date.value; p.Status_Tindak_Lanjut = f.Status_Tindak_Lanjut.value; }
  if (!p.Tanggal_Inspeksi || !p.Waktu_Inspeksi || !p.Nama_Pelaksana.trim() || !p.Aktivitas.trim() || !p.Kategori_Bahaya || !p.Potensi_Bahaya.trim() || !p.Likelihood || !p.Impact || !p.Unsafe_Condition || !p.Unsafe_Action) { salah('Lengkapi semua kolom bertanda *.'); return; }
  if (p.APD_Wajib.length && (!/^\d{1,3}$/.test(String(p.Kepatuhan_APD).trim()) || Number(p.Kepatuhan_APD) > 100)) { salah('Kepatuhan APD harus angka 0–100 bila APD dipilih.'); return; }
  if (ada && !p.Temuan.trim()) { salah('Uraikan temuan bila terdapat Unsafe Condition / Unsafe Action.'); return; }
  if (ada && (!p.Rekomendasi_Tindakan.trim() || !p.PIC_Tindak_Lanjut.trim() || !p.Due_Date)) { salah('Rekomendasi, PIC, dan Due Date tindak lanjut wajib diisi.'); return; }
  if (K3F.mode === 'ubah') p.ID_Inspeksi = K3F.id;
  var btn = $('button[type=submit][form=fK3]'); sibuk(btn, true, 'Menyimpan…');
  try {
    p.Foto_URLs = await k3UnggahFoto(K3F.mode === 'ubah' ? K3F.id : 'baru');
    var d = await call('saveInspeksi', p, K3F.mode); k3Ganti('inspeksi', d.inspeksi); tutupModal();
    toast(K3F.mode === 'ubah' ? 'Inspeksi diperbarui.' : 'Inspeksi ' + d.inspeksi.ID_Inspeksi + ' dicatat.', 'ok'); if (S.page === 'k3') k3Isi();
  } catch (e) { salah(e.message); } finally { sibuk(btn, false); }
};

/* ── Form Insiden ── */
function formInsidenHtml(r) {
  var pil = r ? ['', r.Gedung, r.Lantai, r.Ruangan] : ['', null, null, null];
  var sel = function (i, nm, lbl) { if (i === 0) return '<input type="hidden" name="' + nm + '" value="">'; return '<div class="col-md-6"><label class="form-label req">' + lbl + (i === 2 ? ' <span class="text-muted fw-normal">(boleh tanpa lantai)</span>' : '') + '</label><select class="form-select" name="' + nm + '" data-ik-lok="' + i + '">' + lokOpsiHtml(i, pil, pil[i]) + '</select></div>'; };
  return '<form id="fK3" data-submit="simpanInsiden" novalidate><div class="row g-3">' +
    '<div class="col-md-6"><label class="form-label req">Tanggal Insiden</label><input type="date" class="form-control" name="Tanggal" max="' + hariIniLokal() + '" value="' + esc((r && r.Tanggal) || hariIniLokal()) + '"></div>' +
    '<div class="col-md-6"><label class="form-label req">Waktu Insiden</label><input type="time" class="form-control" name="Waktu" value="' + esc((r && r.Waktu) || jamLokal()) + '"></div>' +
    sel(0, 'Jenis_Ruangan', '') + sel(1, 'Gedung', 'Gedung') + sel(2, 'Lantai', 'Lantai') + sel(3, 'Ruangan', 'Ruangan') +
    '<div class="col-12"><label class="form-label">Lokasi Detail <span class="text-muted fw-normal">(ketik bebas, boleh kosong)</span></label><input class="form-control" name="Lokasi_Detail" maxlength="150" value="' + esc(r && r.Lokasi_Detail) + '"></div>' +
    '<div class="col-md-6"><label class="form-label req">Jenis Insiden</label><select class="form-select" name="Jenis_Insiden">' + opsi(listMaster('Jenis Insiden'), r && r.Jenis_Insiden, '— pilih —') + '</select></div>' +
    '<div class="col-md-6"><label class="form-label req">Tingkat Keparahan</label><select class="form-select" name="Tingkat_Keparahan">' + opsi(listMaster('Tingkat Keparahan'), r && r.Tingkat_Keparahan, '— pilih —') + '</select></div>' +
    '<div class="col-12"><label class="form-label req">Kronologi</label><textarea class="form-control" name="Kronologi" rows="3" maxlength="2000">' + esc(r && r.Kronologi) + '</textarea></div>' +
    '<div class="col-12"><label class="form-label">Korban / pihak terlibat <span class="text-muted fw-normal">(wajib untuk Kecelakaan Kerja & Penyakit Akibat Kerja)</span></label><input class="form-control" name="Korban_Terlibat" maxlength="300" value="' + esc(r && r.Korban_Terlibat) + '"></div>' +
    '<div class="col-md-6"><label class="form-label req">Penyebab</label><textarea class="form-control" name="Penyebab" rows="2" maxlength="1000">' + esc(r && r.Penyebab) + '</textarea></div>' +
    '<div class="col-md-6"><label class="form-label req">Tindakan segera</label><textarea class="form-control" name="Tindakan_Segera" rows="2" maxlength="1000">' + esc(r && r.Tindakan_Segera) + '</textarea></div>' +
    k3TlForm(r, 'insiden', true) + k3FotoBlok() +
    '<div class="col-12"><div class="alert-x danger d-none mb-0" id="k3Err"></div></div></div></form>';
}
SUBMITS.simpanInsiden = async function (f) {
  var err = $('#k3Err'); err.classList.add('d-none');
  var salah = function (m) { err.textContent = m; err.classList.remove('d-none'); };
  var p = {};
  ['Tanggal', 'Waktu', 'Gedung', 'Lantai', 'Ruangan', 'Lokasi_Detail', 'Jenis_Insiden', 'Tingkat_Keparahan', 'Kronologi', 'Korban_Terlibat', 'Penyebab', 'Tindakan_Segera', 'Rekomendasi', 'PIC_Tindak_Lanjut', 'Due_Date', 'Status_Tindak_Lanjut'].forEach(function (k) { p[k] = f[k].value; });
  if (p.Lantai === LANTAI_PILIH) p.Lantai = '';
  var lokOk = p.Gedung && p.Ruangan && f.Lantai.value !== LANTAI_PILIH;
  if (!p.Tanggal || !p.Waktu || !lokOk || !p.Jenis_Insiden || !p.Tingkat_Keparahan || !p.Kronologi.trim() || !p.Penyebab.trim() || !p.Tindakan_Segera.trim() || !p.Rekomendasi.trim() || !p.PIC_Tindak_Lanjut.trim() || !p.Due_Date) { salah('Lengkapi semua kolom bertanda * (termasuk lokasi dan tindak lanjut).'); return; }
  if (!p.Korban_Terlibat.trim() && (p.Jenis_Insiden === 'Kecelakaan Kerja' || p.Jenis_Insiden === 'Penyakit Akibat Kerja')) { salah('Korban/pihak terlibat wajib diisi untuk jenis insiden ini.'); return; }
  if (K3F.mode === 'ubah') p.ID_Insiden = K3F.id;
  var btn = $('button[type=submit][form=fK3]'); sibuk(btn, true, 'Menyimpan…');
  try {
    p.Foto_URLs = await k3UnggahFoto(K3F.mode === 'ubah' ? K3F.id : 'baru');
    var d = await call('saveInsiden', p, K3F.mode); k3Ganti('insiden', d.insiden); tutupModal();
    toast(K3F.mode === 'ubah' ? 'Insiden diperbarui.' : 'Insiden ' + d.insiden.ID_Insiden + ' dicatat.', 'ok'); if (S.page === 'k3') k3Isi();
  } catch (e) { salah(e.message); } finally { sibuk(btn, false); }
};

/* ── Form SOP ── */
function formSopHtml(r) {
  return '<form id="fK3" data-submit="simpanSop" novalidate><div class="row g-3">' +
    '<div class="col-md-5"><label class="form-label req">Nomor SOP</label><input class="form-control" name="Nomor_SOP" maxlength="60" value="' + esc(r && r.Nomor_SOP) + '" placeholder="mis. SOP/K3/001"></div>' +
    '<div class="col-md-7"><label class="form-label req">Judul SOP</label><input class="form-control" name="Judul_SOP" maxlength="200" value="' + esc(r && r.Judul_SOP) + '"></div>' +
    '<div class="col-md-6"><label class="form-label req">Kategori</label><select class="form-select" name="Kategori">' + opsi(listMaster('Kategori SOP'), r && r.Kategori, '— pilih —') + '</select></div>' +
    '<div class="col-md-6"><label class="form-label req">Status</label><select class="form-select" name="Status">' + opsi(K3_STATUS_SOP, (r && r.Status) || 'Berlaku', null) + '</select></div>' +
    '<div class="col-md-6"><label class="form-label req">Tanggal Berlaku</label><input type="date" class="form-control" name="Tanggal_Berlaku" value="' + esc((r && r.Tanggal_Berlaku) || hariIniLokal()) + '"></div>' +
    '<div class="col-md-6"><label class="form-label req">Tanggal Review</label><input type="date" class="form-control" name="Tanggal_Review" value="' + esc(r && r.Tanggal_Review) + '"></div>' +
    '<div class="col-12"><label class="form-label">File SOP (PDF, maks. 2 MB)</label><input type="file" class="form-control" id="k3PdfIn" accept="application/pdf"><div class="form-text" id="k3PdfInfo">' + (r && r.File_URL ? '<a href="' + esc(r.File_URL) + '" target="_blank" rel="noopener noreferrer"><i class="bi bi-file-earmark-pdf me-1"></i>File saat ini</a> <button type="button" class="btn-link-x ms-2" data-act="k3PdfHapus">Hapus file</button>' : 'Belum ada file.') + '</div></div>' +
    '<div class="col-12"><div class="alert-x danger d-none mb-0" id="k3Err"></div></div></div></form>';
}
ACTIONS.k3PdfHapus = function () { K3F.hapusFile = true; K3F.pdf = null; $('#k3PdfIn').value = ''; $('#k3PdfInfo').textContent = 'File akan dihapus saat disimpan.'; };
SUBMITS.simpanSop = async function (f) {
  var err = $('#k3Err'); err.classList.add('d-none');
  var salah = function (m) { err.textContent = m; err.classList.remove('d-none'); };
  var p = {};
  ['Nomor_SOP', 'Judul_SOP', 'Kategori', 'Status', 'Tanggal_Berlaku', 'Tanggal_Review'].forEach(function (k) { p[k] = f[k].value; });
  if (!p.Nomor_SOP.trim() || !p.Judul_SOP.trim() || !p.Kategori || !p.Tanggal_Berlaku || !p.Tanggal_Review) { salah('Lengkapi semua kolom bertanda *.'); return; }
  if (p.Tanggal_Review < p.Tanggal_Berlaku) { salah('Tanggal Review tidak boleh sebelum Tanggal Berlaku.'); return; }
  var lama = K3F.mode === 'ubah' ? K3.sop.filter(function (x) { return x.ID_SOP === K3F.id; })[0] : null;
  if (K3F.mode === 'ubah') p.ID_SOP = K3F.id;
  p.File_URL = K3F.hapusFile ? '' : (lama ? lama.File_URL : '');
  var btn = $('button[type=submit][form=fK3]'); sibuk(btn, true, 'Menyimpan…');
  try {
    if (K3F.pdf) { var u = await call('uploadFile', 'k3', K3F.mode === 'ubah' ? K3F.id : 'baru', { name: K3F.pdf.name, mime: 'application/pdf', base64: String(await bacaFile(K3F.pdf, false)).split(',')[1] }); p.File_URL = u.unduh; }
    var d = await call('saveSop', p, K3F.mode); k3Ganti('sop', d.sop); tutupModal();
    toast(K3F.mode === 'ubah' ? 'SOP diperbarui.' : 'SOP ' + d.sop.Nomor_SOP + ' ditambahkan.', 'ok'); if (S.page === 'k3') k3Isi();
  } catch (e) { salah(e.message); } finally { sibuk(btn, false); }
};


/* ════════════════════════════════════════════════════════
   ANALITIK — Beranda lengkap (WO/PM/KPI/alert), SLA, KPI & Laporan
   ════════════════════════════════════════════════════════ */
var AN = { ber: null, sla: null, kpi: null, f: { sla: null, kpi: null }, dr: null };
var AN_KPI_PERAN = ['Master Admin', 'Kabag', 'Kasubbag'];
var AN_WST = { 'Open': '#1E5EFF', 'Sedang Dikerjakan': '#06B6D4', 'Ditunda': '#F59E0B', 'Selesai': '#16A34A', 'Dibatalkan': '#64748B' };
function anAngka(v, d) { return Number(v).toLocaleString('id-ID', { maximumFractionDigits: d === undefined ? 1 : d }); }
function anPct(v) { return v === null || v === undefined ? '—' : anAngka(v) + '%'; }
function anJam(v) { return v === null || v === undefined ? '—' : anAngka(v) + ' jam'; }
function anNilaiKpi(k) {
  if (k.nilai === null || k.nilai === undefined) return '—';
  if (k.kode === 'rasio') return esc(k.teks);
  return k.satuan === '%' ? anPct(k.nilai) : (k.satuan === 'jam' ? anJam(k.nilai) : anAngka(k.nilai, 0) + ' ' + esc(k.satuan));
}
function anLabelKpi(k) {
  var tren = k.target === null || k.target === undefined;
  var L = tren ? { hijau: 'Membaik', kuning: 'Stabil', merah: 'Memburuk', abu: 'Belum ada pembanding' }
    : (k.kode === 'rasio' ? { hijau: 'Sesuai target', kuning: 'Mendekati', merah: 'Menyimpang', abu: 'Belum ada data' } : { hijau: 'Tercapai', kuning: 'Mendekati', merah: 'Di bawah target', abu: 'Belum ada data' });
  return L[k.status] || '';
}
var AN_KLS = { hijau: 'ok', kuning: 'warn', merah: 'bad', abu: '' };
function anKartuKpi(k) {
  return '<div class="card-x kpi-k ' + k.status + '"><div class="kpi-k-n">' + esc(k.nama) + '</div><div class="kpi-k-v">' + anNilaiKpi(k) + '</div>' +
    '<div class="d-flex flex-wrap gap-2 align-items-center mt-1">' + badge(anLabelKpi(k), AN_KLS[k.status]) + '<span class="small text-muted">Target: ' + esc(k.targetTeks) + '</span></div>' +
    '<div class="small text-muted mt-1">' + esc(k.catatan) + '</div></div>';
}
function anTelat(jam) { return jam < 48 ? jam + ' jam' : Math.floor(jam / 24) + ' hari'; }
function anBlnLabel(b) { return BLN[Number(b.slice(5)) - 1] + ' ' + b.slice(0, 4); }
function anBlnOpsi(sel, kosong) {
  var skr = hariIniLokal().slice(0, 7), y = Number(skr.slice(0, 4)), m = Number(skr.slice(5)), h = kosong ? '<option value="">' + esc(kosong) + '</option>' : '';
  for (var i = 0; i < 36; i++) { var b = y + '-' + ('0' + m).slice(-2); h += '<option value="' + b + '"' + (b === sel ? ' selected' : '') + '>' + anBlnLabel(b) + '</option>'; m--; if (m === 0) { m = 12; y--; } }
  return h;
}
function anGedung() { return uniq((S.master.lokasi || []).map(function (x) { return x.G; })).sort(); }
function anChips(arr, jenis, kunci) {
  return '<div class="d-flex flex-wrap gap-2 mt-2">' + arr.map(function (x) { return '<button type="button" class="btn btn-sm btn-light-x" data-act="anDrillBtn" data-j="' + jenis + '" data-k="' + kunci + '" data-v="' + esc(x[0]) + '">' + esc(x[0]) + ' <b>' + fmtAngka(x[1]) + '</b></button>'; }).join('') + '</div>';
}
function anCursor(e, els) { if (e.native && e.native.target) e.native.target.style.cursor = els.length ? 'pointer' : 'default'; }

/* ── Beranda: bagian analitik (dimuat setelah kerangka Beranda tampil) ── */
async function berandaAn() {
  if (!$('#anKartu')) return;
  var d;
  try { d = await call('getBeranda'); } catch (e) {
    if (S.page === 'beranda' && $('#anKartu')) { $('#anKartu').innerHTML = '<div class="alert-x danger mb-0" style="grid-column:1/-1">' + esc(e.message) + ' <button class="btn-link-x" data-act="refreshAll">Coba lagi</button></div>'; if ($('#anAlert')) $('#anAlert').innerHTML = ''; }
    return;
  }
  if (S.page !== 'beranda' || !$('#anKartu')) return;
  AN.ber = d;
  var ops = OPS.indexOf(S.user.peran) >= 0;
  $('#anKartu').innerHTML = kpiHtml('bi-tools', 'warn', fmtAngka(d.woBerjalan), 'Pekerjaan Berjalan (WO)') +
    kpiHtml('bi-calendar-check', 'ok', anPct(d.pmTerselesaikan), 'PM Terselesaikan bulan ini (' + fmtAngka(d.pmTerlaksana) + '/' + fmtAngka(d.pmJadwal) + ')') +
    (d.risikoTinggi === null ? '' : kpiHtml('bi-exclamation-diamond', 'bad', fmtAngka(d.risikoTinggi), 'Total Risiko Tinggi (belum tertangani)'));
  // alert
  var al = d.alert;
  $('#anAlert').innerHTML = '<div class="card-h"><i class="bi bi-bell"></i>Alert &amp; Notifikasi<span class="badge-x ' + (d.alertTotal ? 'bad' : 'ok') + ' ms-auto">' + (d.alertTotal ? fmtAngka(d.alertTotal) + ' terlewat' : 'Tidak ada yang terlewat') + '</span></div>' +
    (al.length ? '<div class="an-alert">' + al.map(function (x) {
      var isi = '<span class="an-al-ic"><i class="bi ' + (x.jenis === 'WO' ? 'bi-tools' : (x.jenis === 'PM' ? 'bi-calendar-check' : 'bi-shield-check')) + '"></i></span><span class="an-al-t"><b>' + esc(x.jenis) + ' · ' + esc(x.id) + '</b><span class="d-block small text-muted">' + esc(x.judul) + '</span></span>' +
        '<span class="an-al-s">' + badge(x.status, x.status === 'Ditunda' ? 'warn' : '') + '<span class="d-block small text-danger fw-bold mt-1">Terlambat ' + anTelat(x.telatJam) + '</span><span class="d-block small text-muted">Batas ' + esc(fmtTgl(x.batas, !/\d:\d/.test(x.batas))) + '</span></span>';
      return ops ? '<button type="button" class="an-al" data-act="anAlert" data-j="' + esc(x.jenis) + '" data-id="' + esc(x.id) + '">' + isi + '</button>' : '<div class="an-al static">' + isi + '</div>';
    }).join('') + (d.alertTotal > al.length ? '<div class="small text-muted p-2">Menampilkan ' + al.length + ' terlama dari ' + fmtAngka(d.alertTotal) + '.</div>' : '') + '</div>'
      : '<div class="empty"><i class="bi bi-emoji-smile"></i>Semua WO, PM, dan tindak lanjut K3 masih dalam batas waktu.</div>');
  // grafik & daftar
  var gelap = S.theme === 'dark', grid = gelap ? '#22304D' : '#E2E8F0';
  if ($('#anJadwal')) {
    $('#anJadwal').innerHTML = d.jadwalPm.length ? '<div class="tbl-wrap"><table class="tbl"><tbody>' + d.jadwalPm.map(function (p) {
      return '<tr><td><div class="nm">' + esc(p.Nama_Aset) + '</div><div class="sub">' + esc(p.ID_PM) + ' · ' + esc(p.Gedung) + (p.Jenis_PM ? ' · ' + esc(p.Jenis_PM) : '') + '</div></td><td class="acts"><div class="small fw-bold">' + esc(fmtTgl(p.Tanggal_Jadwal, true)) + '</div>' + badge(p.Status_PM, p.Status_PM === 'Ditunda' ? 'warn' : (p.Status_PM === 'Sedang Dikerjakan' ? 'info' : '')) + '</td></tr>';
    }).join('') + '</tbody></table></div>' : '<div class="empty"><i class="bi bi-calendar-x"></i>Tidak ada PM terjadwal dalam 7 hari ke depan.</div>';
  }
  if ($('#anKpi')) $('#anKpi').innerHTML = '<div class="an-kpi-list">' + d.kpi.map(function (k) {
    return '<div class="an-kpi-r ' + k.status + '"><span class="an-dot" aria-hidden="true"></span><span class="an-kpi-n">' + esc(k.nama) + '</span><span class="an-kpi-v">' + anNilaiKpi(k) + '</span><span class="small text-muted an-kpi-l">' + esc(anLabelKpi(k)) + '</span></div>';
  }).join('') + '</div>' + (AN_KPI_PERAN.indexOf(S.user.peran) >= 0 ? '<div class="mt-2"><button class="btn-link-x" data-act="go" data-p="kpi">Buka KPI &amp; Laporan <i class="bi bi-arrow-right"></i></button></div>' : '');
  if (typeof Chart === 'undefined') return;
  Chart.defaults.color = gelap ? '#98A5BF' : '#6B7280'; Chart.defaults.font.family = "'Plus Jakarta Sans',sans-serif";
  if ($('#anTren')) {
    S.charts.push(new Chart($('#anTren'), { type: 'bar', data: { labels: d.tren.map(function (x) { return BLN[Number(x.bulan.slice(5)) - 1]; }), datasets: [
      { label: 'Preventive (PM terlaksana)', data: d.tren.map(function (x) { return x.preventive; }), backgroundColor: PALET[0], borderRadius: 4 },
      { label: 'Corrective (WO selesai)', data: d.tren.map(function (x) { return x.corrective; }), backgroundColor: PALET[2], borderRadius: 4 }] },
      options: { maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: grid } }, x: { grid: { display: false } } }, onHover: anCursor,
        onClick: function (e, els) { if (els.length) anDrill('tren', { bulan: d.tren[els[0].index].bulan, seri: els[0].datasetIndex === 0 ? 'Preventive' : 'Corrective' }); } } }));
  }
  var st = Object.keys(d.statusWo).map(function (k) { return [k, d.statusWo[k]]; });
  if ($('#anStatus') && st.length) {
    S.charts.push(new Chart($('#anStatus'), { type: 'doughnut', data: { labels: st.map(function (x) { return x[0]; }), datasets: [{ data: st.map(function (x) { return x[1]; }), backgroundColor: st.map(function (x, i) { return AN_WST[x[0]] || PALET[i % 8]; }), borderWidth: 0 }] },
      options: { maintainAspectRatio: false, cutout: '62%', plugins: { legend: { position: 'bottom' } }, onHover: anCursor, onClick: function (e, els) { if (els.length) anDrill('status', { status: st[els[0].index][0] }); } } }));
  }
  if ($('#anStatusChip')) $('#anStatusChip').innerHTML = st.length ? anChips(st, 'status', 'status') : '';
}
ACTIONS.anAlert = function (el) {
  var j = el.dataset.j, id = el.dataset.id;
  if (j === 'WO') ACTIONS.woDetail({ dataset: { id: id } });
  else if (j === 'PM') ACTIONS.pmDetail({ dataset: { id: id } });
  else if (j === 'K3 Inspeksi') { K3.tab = 'inspeksi'; K3.fi.q = id; K3.fi.hal = 1; go('k3'); }
  else if (j === 'K3 Insiden') { K3.tab = 'insiden'; K3.fn.q = id; K3.fn.hal = 1; go('k3'); }
};

/* ── Rincian dari klik grafik ── */
var AN_DR_JUDUL = { aset: 'Rincian Aset', tren: 'Rincian Pekerjaan Pemeliharaan', status: 'Rincian Work Order' };
ACTIONS.anDrillBtn = function (el) { var f = {}; f[el.dataset.k] = el.dataset.v; anDrill(el.dataset.j, f); };
async function anDrill(jenis, f) {
  AN.dr = { jenis: jenis, f: f || {} };
  bukaModal({ judul: AN_DR_JUDUL[jenis], size: 'lg', isi: '<div id="anDrBar"></div><div id="anDrIsi"><div class="skeleton" style="height:200px"></div></div>', footer: '<button class="btn btn-primary" data-bs-dismiss="modal">Tutup</button>' });
  anDrBar(); await anDrMuat();
}
function anDrBar() {
  var j = AN.dr.jenis, f = AN.dr.f, sel = function (k, lbl, h) { return '<div class="col-6 col-md-3"><label class="form-label small mb-1">' + lbl + '</label><select class="form-select form-select-sm" data-drf="' + k + '">' + h + '</select></div>'; };
  var h = '';
  if (j === 'aset') h += sel('kondisi', 'Kondisi', opsi(listMaster('Kondisi Aset'), f.kondisi, 'Semua kondisi'));
  if (j === 'tren') { h += sel('seri', 'Jenis', opsi(['Preventive', 'Corrective'], f.seri, null)); h += sel('bulan', 'Bulan', anBlnOpsi(f.bulan, 'Semua bulan')); }
  if (j === 'status') { h += sel('status', 'Status WO', opsi(Object.keys(AN_WST), f.status, 'Semua status')); h += sel('bulan', 'Bulan dibuat', anBlnOpsi(f.bulan, 'Semua bulan')); }
  h += sel('gedung', 'Gedung', opsi(anGedung(), f.gedung, 'Semua gedung')) + sel('kategori', 'Kategori aset', opsi(listMaster('Kategori Aset'), f.kategori, 'Semua kategori'));
  $('#anDrBar').innerHTML = '<div class="row g-2 mb-3">' + h + '</div>';
}
document.addEventListener('change', function (e) {
  var t = e.target; if (!t.dataset || t.dataset.drf === undefined || !AN.dr) return;
  AN.dr.f[t.dataset.drf] = t.value; anDrMuat();
});
async function anDrMuat() {
  var isi = $('#anDrIsi'); if (!isi) return;
  isi.innerHTML = '<div class="skeleton" style="height:200px"></div>';
  var jenis = AN.dr.jenis;
  try {
    var d = await call('getDrill', jenis, AN.dr.f); isi = $('#anDrIsi'); if (!isi || !AN.dr || AN.dr.jenis !== jenis) return;
    var ops = OPS.indexOf(S.user.peran) >= 0;
    if (!d.items.length) { isi.innerHTML = '<div class="empty"><i class="bi bi-inbox"></i>Tidak ada data untuk filter ini.</div>'; return; }
    isi.innerHTML = '<div class="small text-muted mb-2">' + fmtAngka(d.total) + ' data' + (d.total > d.items.length ? ' (menampilkan ' + fmtAngka(d.items.length) + ' terbaru)' : '') + '</div>' +
      '<div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>ID</th><th>Aset</th><th>Status</th><th>Tanggal</th><th></th></tr></thead><tbody>' + d.items.map(function (x) {
        var act = x.jenis === 'WO' && ops ? 'woDetail' : (x.jenis === 'PM' && ops ? 'pmDetail' : (x.jenis === 'Aset' ? 'lihatAset' : ''));
        return '<tr><td data-l="ID"><b>' + esc(x.id) + '</b></td><td data-l="Aset"><div class="nm">' + esc(x.judul) + '</div><div class="sub">' + esc(x.sub) + '</div></td><td data-l="Status">' + (x.jenis === 'Aset' ? bKondisi(x.status) : badge(x.status, AN_WST[x.status] ? { 'Selesai': 'ok', 'Terlaksana': 'ok', 'Ditunda': 'warn', 'Dibatalkan': '', 'Open': 'info', 'Sedang Dikerjakan': 'info' }[x.status] : '')) + '</td><td data-l="Tanggal">' + esc(x.tanggal ? fmtTgl(x.tanggal, true) : '—') + '</td>' +
          '<td class="acts">' + (act ? '<button class="icon-btn" data-act="' + act + '" data-id="' + esc(x.id) + '" title="Detail" aria-label="Detail ' + esc(x.id) + '"><i class="bi bi-eye"></i></button>' : '') + '</td></tr>';
      }).join('') + '</tbody></table></div>';
  } catch (e) { if ($('#anDrIsi')) $('#anDrIsi').innerHTML = '<div class="alert-x danger">' + esc(e.message) + '</div>'; }
}

/* ── Filter periode (SLA & KPI) ── */
function anFilterHtml(k) {
  var f = AN.f[k];
  return '<div class="card-x mb-3"><div class="card-b"><div class="row g-2 align-items-end">' +
    '<div class="col-6 col-md-2"><label class="form-label small mb-1">Dari bulan</label><select class="form-select form-select-sm" data-anf="dari">' + anBlnOpsi(f.dari) + '</select></div>' +
    '<div class="col-6 col-md-2"><label class="form-label small mb-1">Sampai bulan</label><select class="form-select form-select-sm" data-anf="sampai">' + anBlnOpsi(f.sampai) + '</select></div>' +
    '<div class="col-6 col-md-3"><label class="form-label small mb-1">Gedung</label><select class="form-select form-select-sm" data-anf="gedung">' + opsi(anGedung(), f.gedung, 'Semua gedung') + '</select></div>' +
    '<div class="col-6 col-md-3"><label class="form-label small mb-1">Kategori aset</label><select class="form-select form-select-sm" data-anf="kategori">' + opsi(listMaster('Kategori Aset'), f.kategori, 'Semua kategori') + '</select></div>' +
    '<div class="col-12 col-md-2"><button class="btn btn-primary btn-sm w-100" data-act="anTerapkan" data-k="' + k + '"><i class="bi bi-funnel me-1"></i>Terapkan</button></div></div>' +
    '<div class="small text-muted mt-2" id="anFilterErr" role="alert"></div></div></div>';
}
function anFilterBaru(k) { var skr = hariIniLokal().slice(0, 7), y = Number(skr.slice(0, 4)), m = Number(skr.slice(5)) - (k === 'sla' ? 5 : 2); while (m < 1) { m += 12; y--; } return { dari: y + '-' + ('0' + m).slice(-2), sampai: skr, gedung: '', kategori: '' }; }
ACTIONS.anTerapkan = function (el) {
  var k = el.dataset.k, f = AN.f[k];
  document.querySelectorAll('[data-anf]').forEach(function (s) { f[s.dataset.anf] = s.value; });
  if (f.dari > f.sampai) { $('#anFilterErr').textContent = 'Bulan awal tidak boleh setelah bulan akhir.'; return; }
  if ((Number(f.sampai.slice(0, 4)) - Number(f.dari.slice(0, 4))) * 12 + Number(f.sampai.slice(5)) - Number(f.dari.slice(5)) >= 24) { $('#anFilterErr').textContent = 'Rentang periode maksimal 24 bulan.'; return; }
  if (k === 'sla') renderSla(); else renderKpi();
};


/* ── Laporan PDF ── */
function unduhPdf(r) {
  var bin = atob(r.base64), u = new Uint8Array(bin.length);
  for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  unduhBlob(r.nama, new Blob([u], { type: 'application/pdf' }));
}
ACTIONS.woPdf = async function (el) {
  sibuk(el, true, 'Membuat PDF…');
  try { var r = await call('cetakPdfWo', el.dataset.id); unduhPdf(r); toast('PDF ' + el.dataset.id + ' diunduh.', 'ok'); } catch (e) { gagal(e); } finally { sibuk(el, false); }
};
ACTIONS.pmPdf = async function (el) {
  sibuk(el, true, 'Membuat PDF…');
  try { var r = await call('cetakPdfPm', el.dataset.id); unduhPdf(r); toast('PDF ' + el.dataset.id + ' diunduh.', 'ok'); } catch (e) { gagal(e); } finally { sibuk(el, false); }
};
function anRekapHtml() {
  var skr = hariIniLokal().slice(0, 7), y = Number(skr.slice(0, 4)), m = Number(skr.slice(5)) - 1; if (m < 1) { m = 12; y--; }
  var lalu = y + '-' + ('0' + m).slice(-2);
  return '<div class="card-x mt-3"><div class="card-h"><i class="bi bi-file-earmark-pdf"></i>Laporan PDF Rekap Bulanan</div><div class="card-b"><form id="fRekap" data-submit="anRekap" novalidate><div class="row g-2 align-items-end">' +
    '<div class="col-6 col-md-3"><label class="form-label small mb-1">Bulan</label><select class="form-select form-select-sm" name="bulan">' + anBlnOpsi(lalu) + '</select></div>' +
    '<div class="col-6 col-md-2"><label class="form-label small mb-1">Jenis</label><select class="form-select form-select-sm" name="jenis"><option value="Semua">WO &amp; PM</option><option value="WO">Work Order</option><option value="PM">Preventive Maintenance</option></select></div>' +
    '<div class="col-6 col-md-3"><label class="form-label small mb-1">Gedung</label><select class="form-select form-select-sm" name="gedung">' + opsi(anGedung(), '', 'Semua gedung') + '</select></div>' +
    '<div class="col-6 col-md-2"><label class="form-label small mb-1">Kategori aset</label><select class="form-select form-select-sm" name="kategori">' + opsi(listMaster('Kategori Aset'), '', 'Semua kategori') + '</select></div>' +
    '<div class="col-12 col-md-2"><button class="btn btn-primary btn-sm w-100" type="submit"><i class="bi bi-download me-1"></i>Buat PDF</button></div></div>' +
    '<div class="form-text mt-2">Memuat laporan WO Selesai dan PM Terlaksana pada bulan terpilih (menurut waktu selesai), satu halaman per laporan lengkap dengan foto. Maksimal 40 laporan per PDF — persempit dengan filter bila lebih. Salinan tersimpan di folder Drive &quot;Laporan PDF&quot;.</div>' +
    '<div class="alert-x danger d-none mt-2 mb-0" id="anRekapErr"></div></form></div></div>';
}
SUBMITS.anRekap = async function (f) {
  var err = $('#anRekapErr'); err.classList.add('d-none');
  var btn = f.querySelector('button[type=submit]'); sibuk(btn, true, 'Membuat PDF…');
  try {
    var r = await call('cetakPdfRekap', { bulan: f.bulan.value, jenis: f.jenis.value, gedung: f.gedung.value, kategori: f.kategori.value });
    unduhPdf(r); toast('PDF rekap diunduh (' + r.jumlah + ' laporan).', 'ok');
  } catch (e) { err.textContent = e.message; err.classList.remove('d-none'); } finally { sibuk(btn, false); }
};
ACTIONS.kpiEkspor = async function (el) {
  if (!AN.kpi) return;
  var t1 = [['Indikator', 'Nilai', 'Target', 'Status', 'Keterangan']].concat(AN.kpi.kpi.map(function (k) { return [k.nama, k.kode === 'rasio' ? k.teks : (k.nilai === null ? '' : k.nilai), k.targetTeks, anLabelKpi(k), k.catatan]; }));
  var t2 = [['Nama', 'Tipe', 'WO selesai', 'PM terlaksana', 'Total', 'Rata-rata waktu kerja (jam)']].concat(AN.kpi.teknisi.map(function (t) { return [amanSel(t.nama), t.tipe, t.wo, t.pm, t.total, t.rataJam === null ? '' : t.rataJam]; }));
  t1 = t1.map(function (r, i) { return i ? r.map(function (c) { return typeof c === 'number' ? c : amanSel(c); }) : r; });
  sibuk(el, true, 'Menyiapkan…');
  var nama = 'KPI_' + AN.kpi.filter.dari + '_sd_' + AN.kpi.filter.sampai;
  try {
    await muatSkrip(CDN_XLSX, function () { return typeof XLSX !== 'undefined'; });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(t1), 'KPI'); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(t2), 'Kinerja Teknisi');
    unduhBlob(nama + '.xlsx', new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    toast('Berkas Excel diunduh.', 'ok');
  } catch (e) {
    unduhBlob(nama + '.csv', new Blob([csvKeString(t1)], { type: 'text/csv;charset=utf-8' }));
    toast('Pustaka Excel tidak tersedia — KPI diunduh sebagai CSV.', 'warn', 6000);
  } finally { sibuk(el, false); }
};

/* ── Halaman SLA ── */
async function renderSla() {
  hancurkanChart();
  if (!AN.f.sla) AN.f.sla = anFilterBaru('sla');
  $('#main').innerHTML = '<div class="page-head"><div><h2>SLA</h2><div class="text-muted small">Tingkat layanan Work Order: ' + anBlnLabel(AN.f.sla.dari) + ' – ' + anBlnLabel(AN.f.sla.sampai) + '</div></div></div>' + anFilterHtml('sla') + '<div id="slaBody"><div class="skeleton" style="height:90px"></div><div class="skeleton mt-3" style="height:300px"></div></div>';
  var s;
  try { s = await call('getSla', AN.f.sla); } catch (e) { if (S.page === 'sla') $('#slaBody').innerHTML = '<div class="card-x"><div class="empty"><i class="bi bi-exclamation-circle"></i>' + esc(e.message) + '<div class="mt-3"><button class="btn btn-primary" data-act="anTerapkan" data-k="sla">Coba lagi</button></div></div></div>'; return; }
  if (S.page !== 'sla' || !$('#slaBody')) return;
  AN.sla = s;
  var st = s.compliance === null ? 'abu' : (s.compliance >= s.target ? 'hijau' : (s.compliance >= s.target * 0.9 ? 'kuning' : 'merah'));
  var kartu = anKartuKpi({ kode: 'sla', nama: 'SLA Compliance', satuan: '%', nilai: s.compliance, target: s.target, targetTeks: '≥ ' + s.target + '%', status: st, catatan: s.tepat + ' dari ' + s.selesai + ' WO selesai tepat waktu' });
  var kosong = !s.selesai;
  $('#slaBody').innerHTML = '<div class="kpi-grid">' + kartu +
    kpiHtml('bi-lightning-charge', 'warn', anJam(s.rataRespons), 'Rata-rata Waktu Respons') + kpiHtml('bi-hourglass-split', '', anJam(s.rataPenyelesaian), 'Rata-rata Waktu Penyelesaian') +
    kpiHtml('bi-x-octagon', 'bad', fmtAngka(s.dibatalkan), 'Total Permintaan Dibatalkan') + '</div>' +
    '<div class="small text-muted mb-3">SLA Compliance = WO Selesai tepat waktu ÷ total WO Selesai pada periode (menurut tanggal selesai). Waktu respons = sejak WO dibuat sampai teknisi menekan Mulai.</div>' +
    (kosong ? '<div class="card-x"><div class="empty"><i class="bi bi-inbox"></i>Belum ada Work Order selesai pada periode dan filter ini.</div></div>' :
      '<div class="grid-2"><div class="card-x" style="grid-column:1/-1"><div class="card-h">Tren SLA Compliance <span class="badge-x ms-auto">per bulan</span></div><div class="card-b"><div class="chart-box"><canvas id="gSlaTren" role="img" aria-label="Grafik tren SLA Compliance bulanan"></canvas></div></div></div>' +
      '<div class="card-x"><div class="card-h">Distribusi Status SLA</div><div class="card-b"><div class="chart-box"><canvas id="gSlaDis" role="img" aria-label="Distribusi status SLA"></canvas></div></div></div>' +
      '<div class="card-x"><div class="card-h">SLA Compliance per Kategori Aset</div><div class="card-b"><div class="chart-box"><canvas id="gSlaKat" role="img" aria-label="SLA Compliance per kategori aset"></canvas></div></div></div></div>' +
      '<div class="card-x mt-3"><div class="card-h">Rincian per Bulan</div><div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Bulan</th><th>WO selesai</th><th>Tepat waktu</th><th>SLA Compliance</th></tr></thead><tbody>' + s.tren.map(function (t) {
        return '<tr><td data-l="Bulan">' + anBlnLabel(t.bulan) + '</td><td data-l="WO selesai">' + fmtAngka(t.selesai) + '</td><td data-l="Tepat waktu">' + fmtAngka(t.tepat) + '</td><td data-l="SLA Compliance">' + anPct(t.compliance) + '</td></tr>'; }).join('') + '</tbody></table></div></div>');
  if (kosong || typeof Chart === 'undefined') return;
  var gelap = S.theme === 'dark', grid = gelap ? '#22304D' : '#E2E8F0';
  Chart.defaults.color = gelap ? '#98A5BF' : '#6B7280'; Chart.defaults.font.family = "'Plus Jakarta Sans',sans-serif";
  S.charts.push(new Chart($('#gSlaTren'), { type: 'line', data: { labels: s.tren.map(function (t) { return anBlnLabel(t.bulan); }), datasets: [
    { label: 'SLA Compliance (%)', data: s.tren.map(function (t) { return t.compliance; }), borderColor: PALET[0], backgroundColor: PALET[0], tension: .25, spanGaps: false },
    { label: 'Target (' + s.target + '%)', data: s.tren.map(function () { return s.target; }), borderColor: '#16A34A', borderDash: [6, 4], pointRadius: 0 }] },
    options: { maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, scales: { y: { min: 0, max: 100, grid: { color: grid }, ticks: { callback: function (v) { return v + '%'; } } }, x: { grid: { display: false } } } } }));
  S.charts.push(new Chart($('#gSlaDis'), { type: 'doughnut', data: { labels: ['Tepat Waktu', 'Terlambat'], datasets: [{ data: [s.tepat, s.terlambat], backgroundColor: ['#16A34A', '#DC2626'], borderWidth: 0 }] }, options: { maintainAspectRatio: false, cutout: '62%', plugins: { legend: { position: 'bottom' } } } }));
  S.charts.push(new Chart($('#gSlaKat'), { type: 'bar', data: { labels: s.perKategori.map(function (x) { return x.kategori + ' (' + x.selesai + ')'; }), datasets: [{ label: 'SLA Compliance (%)', data: s.perKategori.map(function (x) { return x.compliance; }), backgroundColor: s.perKategori.map(function (x) { return x.compliance >= s.target ? '#16A34A' : (x.compliance >= s.target * 0.9 ? '#F59E0B' : '#DC2626'); }), borderRadius: 4 }] },
    options: { indexAxis: 'y', maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { min: 0, max: 100, grid: { color: grid }, ticks: { callback: function (v) { return v + '%'; } } }, y: { grid: { display: false } } } } }));
}

/* ── Halaman KPI & Laporan ── */
async function renderKpi() {
  hancurkanChart();
  if (!AN.f.kpi) AN.f.kpi = anFilterBaru('kpi');
  $('#main').innerHTML = '<div class="page-head"><div><h2>KPI &amp; Laporan</h2><div class="text-muted small">Kinerja pemeliharaan: ' + anBlnLabel(AN.f.kpi.dari) + ' – ' + anBlnLabel(AN.f.kpi.sampai) + '</div></div>' +
    '<div class="ms-auto d-flex flex-wrap gap-2">' + '<button class="btn btn-light-x" data-act="kpiEkspor"><i class="bi bi-file-earmark-excel me-1"></i>Ekspor Excel</button>' + (bolehLabel() ? '<button class="btn btn-light-x" data-act="go" data-p="label"><i class="bi bi-qr-code me-1"></i>Cetak Label QR</button>' : '') + '</div></div>' + anFilterHtml('kpi') +
    '<div id="kpiBody"><div class="skeleton" style="height:90px"></div><div class="skeleton mt-3" style="height:300px"></div></div>';
  var d;
  try { d = await call('getKpi', AN.f.kpi); } catch (e) { if (S.page === 'kpi') $('#kpiBody').innerHTML = '<div class="card-x"><div class="empty"><i class="bi bi-exclamation-circle"></i>' + esc(e.message) + '<div class="mt-3"><button class="btn btn-primary" data-act="anTerapkan" data-k="kpi">Coba lagi</button></div></div></div>'; return; }
  if (S.page !== 'kpi' || !$('#kpiBody')) return;
  AN.kpi = d;
  $('#kpiBody').innerHTML = '<div class="an-legend small mb-3"><span class="badge-x ok">Hijau</span> mencapai target / membaik &nbsp; <span class="badge-x warn">Kuning</span> mendekati / stabil &nbsp; <span class="badge-x bad">Merah</span> di bawah target / memburuk &nbsp; <span class="badge-x">Abu-abu</span> belum ada data pembanding</div>' +
    '<div class="kpi-grid kpi-grid-k">' + d.kpi.map(anKartuKpi).join('') + '</div>' +
    '<div class="small text-muted mb-3">Target PM Compliance, SLA Compliance, dan Rasio Preventive dapat diubah Master Admin di Data Master → Target KPI. MTTR, Backlog, Aset Kondisi Baik, dan Tindak Lanjut K3 dinilai dari tren (backlog, aset, dan K3 dibandingkan dengan catatan bulan lalu).</div>' +
    '<div class="card-x"><div class="card-h"><i class="bi bi-people"></i>Kinerja per Teknisi <span class="badge-x ms-auto">periode terpilih</span></div>' +
    (d.teknisi.length ? '<div class="tbl-wrap"><table class="tbl stack"><thead><tr><th>Nama</th><th>Tipe</th><th>WO selesai</th><th>PM terlaksana</th><th>Total</th><th>Rata-rata waktu kerja</th></tr></thead><tbody>' + d.teknisi.map(function (t) {
      return '<tr><td data-l="Nama"><b>' + esc(t.nama) + '</b></td><td data-l="Tipe">' + badge(t.tipe, t.tipe === 'Vendor' ? 'info' : '') + '</td><td data-l="WO selesai">' + fmtAngka(t.wo) + '</td><td data-l="PM terlaksana">' + fmtAngka(t.pm) + '</td><td data-l="Total"><b>' + fmtAngka(t.total) + '</b></td><td data-l="Rata-rata waktu kerja">' + anJam(t.rataJam) + '</td></tr>'; }).join('') + '</tbody></table></div>'
      : '<div class="empty"><i class="bi bi-inbox"></i>Belum ada pekerjaan selesai pada periode dan filter ini.</div>') + '</div>' +
    (['Kabag', 'Kasubbag'].indexOf(S.user.peran) >= 0 ? anRekapHtml() : '');
}


/* ════════ Pengikat event global ════════ */
document.addEventListener('click', function (e) {
  var el = e.target.closest('[data-act]'); if (!el || el.disabled) return;
  var fn = ACTIONS[el.dataset.act]; if (fn) fn(el, e);
});
document.addEventListener('submit', function (e) {
  var f = e.target.closest('form[data-submit]'); if (!f) return;
  e.preventDefault(); var fn = SUBMITS[f.dataset.submit]; if (fn) fn(f);
});
document.addEventListener('visibilitychange', function () { if (document.hidden) matikanScan(); });

boot();
})();
