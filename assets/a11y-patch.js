/* =============================================================================
   assets/a11y-patch.js  ·  v1.3
   AIM+ AP Target Setting — Decision Workbook
   -----------------------------------------------------------------------------
   File terpisah, dimuat SETELAH assets/app.js di index.html. Tidak ada satu
   baris pun di app.js yang diubah atau di-monkey-patch.

   Cara kerjanya: patch ini tidak ikut campur ke boot(), go(), atau render sheet
   mana pun. Ia memasang satu MutationObserver, lalu menyapu ulang DOM setiap
   kali isinya berganti. Jadi tombol ikon yang dibuat app.js lewat string
   injection tetap ikut tertangani, tanpa perlu menyentuh 2000 baris itu.

   Untuk membatalkan seluruh patch: hapus satu <script> di index.html.
   ============================================================================= */

(function (w, d) {
  'use strict';

  /* Titik pasang aplikasi. Dashboard ini merender ke #app; versi workbook
     lama memakai #sheets. Keduanya didukung supaya patch tidak perlu diubah
     lagi kalau kerangkanya berganti. */
  function mount() {
    return d.getElementById('app') || d.getElementById('sheets') || null;
  }

  /* ---------------------------------------------------------------------------
     11.A · KAMUS ARIA-LABEL  (audit poin 3)
     ---------------------------------------------------------------------------
     Kunci = id tombol, atau glyph persis kalau tombolnya tak ber-id.
     Tambahkan baris baru di sini kalau ada ikon lain yang belum tertangani.
  --------------------------------------------------------------------------- */
  var LABEL_BY_ID = {
    impCheck : 'Periksa data yang ditempel',
    impGo    : 'Muat data ke dashboard'
    /* Tambahkan id tombol lain di sini kalau ada ikon yang belum tertangani. */
  };

  var LABEL_BY_GLYPH = {
    '✕': 'Tutup',
    '×': 'Tutup',
    '⤓': 'Impor',
    '⤒': 'Ekspor',
    '⎙': 'Cetak',
    '↻': 'Hitung ulang',
    '✎': 'Ubah',
    '⋯': 'Aksi lain',
    '▸': 'Buka rincian',
    '▾': 'Buka rincian',
    '▴': 'Tutup rincian',
    '↓': 'Unduh',
    '⤴': 'Unduh',
    '?': 'Bantuan',
    'ⓘ': 'Keterangan',
    '＋': 'Tambah baris',
    '+' : 'Tambah baris',
    '–' : 'Hapus baris',
    '−' : 'Hapus baris'
  };

  /* Sebuah tombol dianggap "icon-only" kalau teksnya <= 2 karakter
     dan tidak ada huruf/angka di dalamnya. */
  function isIconOnly(txt) {
    var t = (txt || '').trim();
    return t.length > 0 && t.length <= 2 && !/[0-9A-Za-z]/.test(t);
  }

  function labelButtons(root) {
    var btns = (root || d).querySelectorAll('button, [role="button"]');
    for (var i = 0; i < btns.length; i++) {
      var b = btns[i];
      if (b.getAttribute('aria-label')) continue;

      var txt = (b.textContent || '').trim();
      var lbl = LABEL_BY_ID[b.id] || null;

      if (!lbl && isIconOnly(txt)) lbl = LABEL_BY_GLYPH[txt] || null;
      if (!lbl && isIconOnly(txt) && b.hasAttribute('data-close')) lbl = 'Tutup';
      if (!lbl && isIconOnly(txt) && b.title) lbl = b.title;

      if (lbl) b.setAttribute('aria-label', lbl);

      /* Glyph dekoratif jangan ikut dibaca dua kali oleh screen reader
         kalau tombolnya sudah punya teks pendamping. */
      if (b.getAttribute('aria-label') && isIconOnly(txt)) {
        b.setAttribute('data-icon-only', '1');
      }
    }
  }

  /* ---------------------------------------------------------------------------
     11.B · MIGRASI title= → data-tip  (audit poin 4)
     ---------------------------------------------------------------------------
     title= dihapus supaya tooltip native browser tidak muncul bertumpuk
     dengan popover CSS. Isinya dipindah ke data-tip (mata) dan, kalau
     tombolnya belum punya nama aksesibel, ke aria-label juga (telinga).
  --------------------------------------------------------------------------- */
  function tipify(root) {
    var els = (root || d).querySelectorAll('[title]');
    for (var i = 0; i < els.length; i++) {
      var el = els[i];
      var t  = el.getAttribute('title');
      if (!t) { el.removeAttribute('title'); continue; }

      /* <abbr> dan <iframe> memang mengandalkan title — lewati. */
      var tag = el.tagName.toLowerCase();
      if (tag === 'abbr' || tag === 'iframe') continue;

      el.setAttribute('data-tip', t);
      if (!el.getAttribute('aria-label') && !(el.textContent || '').trim()) {
        el.setAttribute('aria-label', t);
      }
      el.removeAttribute('title');

      /* Tombol di chrome bar (paling atas) → balon diarahkan ke bawah. */
      if (el.closest && el.closest('.chrome')) {
        el.setAttribute('data-tip-pos', 'below');
      }
    }
  }

  /* ---------------------------------------------------------------------------
     11.C · FOCUS TRAP MODAL  (audit poin 3)
     ---------------------------------------------------------------------------
     Berlaku untuk #scrimImport, #scrimExport, #scrimLocked, dan .scrim mana pun
     yang ditambahkan nanti. Tidak memerlukan perubahan pada fungsi open/close
     yang sudah ada — dipicu oleh MutationObserver pada atribut class.

     Yang dijamin:
       · fokus pindah ke elemen pertama saat modal dibuka
       · Tab dan Shift+Tab berputar di dalam modal
       · Esc menutup modal
       · fokus kembali ke tombol pemicu saat modal ditutup
       · latar diberi aria-hidden supaya screen reader tidak menembusnya
  --------------------------------------------------------------------------- */
  var FOCUSABLE = [
    'a[href]', 'button:not([disabled])', 'input:not([disabled]):not([type="hidden"])',
    'select:not([disabled])', 'textarea:not([disabled])',
    '[tabindex]:not([tabindex="-1"])'
  ].join(',');

  var trap = {
    active   : null,   /* .scrim yang sedang terbuka */
    returnTo : null,   /* elemen yang harus menerima fokus kembali */
    onKey    : null
  };

  function focusablesIn(el) {
    var list = el.querySelectorAll(FOCUSABLE), out = [];
    for (var i = 0; i < list.length; i++) {
      var n = list[i];
      if (n.offsetWidth || n.offsetHeight || n.getClientRects().length) out.push(n);
    }
    return out;
  }

  function backdropAriaHidden(state) {
    var bg = [mount(), d.querySelector('.skiplink')];
    for (var i = 0; i < bg.length; i++) {
      if (!bg[i]) continue;
      if (state) bg[i].setAttribute('aria-hidden', 'true');
      else       bg[i].removeAttribute('aria-hidden');
    }
  }

  function openTrap(scrim) {
    if (trap.active === scrim) return;
    closeTrap(true);                                  /* jaga-jaga bertumpuk */

    trap.active   = scrim;
    trap.returnTo = d.activeElement;

    var modal = scrim.querySelector('.modal') || scrim;
    modal.setAttribute('role', 'dialog');
    modal.setAttribute('aria-modal', 'true');
    modal.setAttribute('tabindex', '-1');

    /* Judul modal jadi nama aksesibel dialog */
    var h = modal.querySelector('h3, h2');
    if (h) {
      if (!h.id) h.id = (scrim.id || 'modal') + '_title';
      modal.setAttribute('aria-labelledby', h.id);
    }

    d.body.classList.add('modal-open');
    backdropAriaHidden(true);

    /* Fokus ke kontrol pertama yang berguna, bukan ke tombol ✕ */
    var f = focusablesIn(modal);
    var first = null;
    for (var i = 0; i < f.length; i++) {
      if (!f[i].hasAttribute('data-close')) { first = f[i]; break; }
    }
    (first || f[0] || modal).focus();

    trap.onKey = function (e) {
      if (e.key === 'Escape') {
        e.preventDefault();
        var closer = modal.querySelector('[data-close]');
        if (closer) closer.click(); else scrim.classList.remove('on');
        return;
      }
      if (e.key !== 'Tab') return;

      var items = focusablesIn(modal);
      if (!items.length) { e.preventDefault(); return; }

      var firstEl = items[0], lastEl = items[items.length - 1];
      if (e.shiftKey && d.activeElement === firstEl) {
        e.preventDefault(); lastEl.focus();
      } else if (!e.shiftKey && d.activeElement === lastEl) {
        e.preventDefault(); firstEl.focus();
      } else if (!modal.contains(d.activeElement)) {
        e.preventDefault(); firstEl.focus();
      }
    };
    d.addEventListener('keydown', trap.onKey, true);
  }

  function closeTrap(silent) {
    if (!trap.active) return;
    d.removeEventListener('keydown', trap.onKey, true);
    d.body.classList.remove('modal-open');
    backdropAriaHidden(false);

    var back = trap.returnTo;
    trap.active = null; trap.returnTo = null; trap.onKey = null;

    if (!silent && back && d.contains(back) && typeof back.focus === 'function') {
      back.focus();
    }
  }

  /* Klik di area gelap juga menutup — perilaku yang diharapkan pengguna. */
  function wireScrimClick(scrim) {
    if (scrim.getAttribute('data-scrimwired')) return;
    scrim.setAttribute('data-scrimwired', '1');
    scrim.addEventListener('mousedown', function (e) {
      if (e.target === scrim) {
        var closer = scrim.querySelector('[data-close]');
        if (closer) closer.click(); else scrim.classList.remove('on');
      }
    });
  }

  /* ---------------------------------------------------------------------------
     11.D · SKELETON  (audit poin 5)
     ---------------------------------------------------------------------------
     Markup skeleton ditaruh statis di index.html (lihat catatan pemasangan),
     jadi ia tampil di paint pertama, sebelum app.js dieksekusi.
     Fungsi ini hanya mengumumkan statusnya ke screen reader dan
     membersihkan sisa kalau boot() ternyata tidak menimpanya.
  --------------------------------------------------------------------------- */
  function skeletonAnnounce() {
    var sk = d.querySelector('.skel-wrap');
    if (!sk) return;
    sk.setAttribute('aria-hidden', 'true');

    if (!d.getElementById('bootStatus')) {
      var live = d.createElement('p');
      live.id = 'bootStatus';
      live.setAttribute('role', 'status');
      live.setAttribute('aria-live', 'polite');
      live.style.cssText =
        'position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);';
      live.textContent = 'Memuat data workbook…';
      d.body.appendChild(live);
    }
  }

  /* v1.3 — PERBAIKAN BUG.
     v1.1 menghapus skeleton tanpa syarat. Kalau boot() gagal, halaman jadi
     benar-benar kosong dan tidak ada petunjuk apa pun. Sekarang skeleton hanya
     dihapus kalau #sheets sudah berisi sesuatu selain skeleton itu sendiri. */
  function skeletonDone() {
    var sheets = mount();
    if (!sheets) return;

    var sk = sheets.querySelector('.skel-wrap');
    if (!sk) return;

    var kids = sheets.children, other = 0;
    for (var i = 0; i < kids.length; i++) {
      if (kids[i] !== sk && !kids[i].classList.contains('skel-wrap')) other++;
    }
    if (!other) return;                       /* boot() belum merender apa pun */

    sk.parentNode.removeChild(sk);
    var live = d.getElementById('bootStatus');
    if (live) live.textContent = 'Workbook siap.';
  }

  /* Pengawas: kalau setelah 4 detik #sheets masih hanya berisi skeleton,
     boot() hampir pasti gagal. Tampilkan diagnosis, jangan diamkan. */
  function bootWatchdog() {
    w.setTimeout(function () {
      var sheets = mount();
      if (!sheets) return;

      var kids = sheets.children, other = 0;
      for (var i = 0; i < kids.length; i++) {
        if (!kids[i].classList.contains('skel-wrap')) other++;
      }
      if (other) return;                      /* aman, sudah terender */

      var globals = [];
      for (var k in w) { if (/^WVI/.test(k)) globals.push(k); }

      var box = d.createElement('div');
      box.className = 'bootfail noprint';
      box.innerHTML =
        '<b>boot() tidak selesai.</b>' +
        '<p>Kerangka halaman termuat, tapi app.js berhenti sebelum merender lembar. ' +
        'Buka DevTools \u2192 Console untuk melihat pesan error yang pertama.</p>' +
        '<p>Global data yang berhasil terbaca (' + globals.length + '): ' +
        (globals.length ? '<code>' + globals.join('</code> <code>') + '</code>'
                        : '<i>tidak ada satu pun</i> \u2014 periksa path di tag &lt;script&gt;') +
        '</p>';
      sheets.appendChild(box);
    }, 4000);
  }

  /* ---------------------------------------------------------------------------
     11.E · SAPUAN ULANG
  --------------------------------------------------------------------------- */
  var sweepQueued = false;

  function sweep() {
    sweepQueued = false;
    tipify(d);
    labelButtons(d);
    skeletonDone();

    var scrims = d.querySelectorAll('.scrim');
    for (var i = 0; i < scrims.length; i++) {
      var s = scrims[i];
      wireScrimClick(s);
      var open = s.classList.contains('on');
      if (open && trap.active !== s) openTrap(s);
      if (!open && trap.active === s) closeTrap(false);
    }
  }

  function queueSweep() {
    if (sweepQueued) return;
    sweepQueued = true;
    (w.requestAnimationFrame || w.setTimeout)(sweep, 0);
  }

  function start() {
    skeletonAnnounce();
    sweep();
    bootWatchdog();

    new MutationObserver(queueSweep).observe(d.body, {
      childList : true,
      subtree   : true,
      attributes: true,
      attributeFilter: ['class', 'title']
    });
  }

  if (d.readyState === 'loading') d.addEventListener('DOMContentLoaded', start);
  else start();

  /* Diekspos supaya bisa dipanggil manual dari konsol atau dari boot(). */
  w.WVIA11Y = {
    mount       : mount,
    sweep       : sweep,
    tipify      : tipify,
    labelButtons: labelButtons,
    skeletonDone: skeletonDone,
    bootWatchdog: bootWatchdog,
    LABEL_BY_ID : LABEL_BY_ID,
    LABEL_BY_GLYPH: LABEL_BY_GLYPH
  };

})(window, document);
