# AIM+ Target Setting — National Summary Dashboard

Situs statis, lima halaman. PEARL · Wahana Visi Indonesia · siklus FY27–FY30.
Data: `Master_Setting_Target (9).xlsx` — **506 baris · 17 kolom · 17 AP · 26 indikator · 4 Zonal**.
Kode akses: **`wvipearl`**.

| Halaman | Isi |
|---|---|
| **`NATIONAL SUMMARY`** *(landing)* | KPI · filter global · tabel weighted · Baseline vs Evaluation · Delta |
| `SUMMARY` | Tabel lengkap 15 kolom, bisa diurutkan |
| `ANALISIS AP` | Status per AP dan per Outcome, tabel analisis |
| `Asumsi Indikator` | Arah & Target Delta per indikator |
| `Pemetaan Indikator` | Matriks Yes/No indikator × AP |

---

## Metode: Weighted National (%)

```
Weighted National (%) = Σ Numerator ÷ Σ Denominator × 100
```

Numerator dijumlahkan dulu, denominator dijumlahkan dulu, baru dibagi. Dihitung per
indikator. Tidak ada rata-rata persentase antar AP atau antar zona.

**Setara dengan rumus di workbook Anda.** Di `IND160 Dashboard` Master (9), kolom
`HH Targeted` berisi angka yang **sama dengan Denominator (LOP)** pada seluruh 17 ADP.
Karena itu:

```
Σ(ADP Weight × ADP Rate) = Σ(den × num/den) ÷ Σden = Σ num ÷ Σ den
```

Kedua rumus identik secara aljabar. Terbukti pada angkanya: `OIOS 160` di sheet Anda
**18,34% evaluation** dan **24,72% baseline** — sama persis dengan halaman ini.
Dari 26 dashboard indikator, **24 cocok**.

Dua yang tidak cocok adalah `OIOS 2` dan `OIOS 3` — keduanya *Number of vulnerable
children reached*, jadi **bukan proporsi**. Membagi jumlah dengan jumlah menghasilkan
35.330% dan 26.438% di sheet Anda. Keduanya perlu diperlakukan sebagai angka absolut,
bukan persentase. `C5G.027599` juga berbeda sedikit pada baseline: 26,28% di sheet Anda
versus 24,38% dari Σnum ÷ Σden.

### Dua konsekuensi

**Baris tanpa denominator tidak bisa diberi bobot** dan dikeluarkan dari hitungan.
Cakupan Master (9): denominator baseline **199 dari 506 baris**, denominator evaluation
**201**. Kolom `AP` di tabel menunjukkan berapa Area Program yang benar-benar menyumbang
tiap angka.

**Tidak ada total lintas indikator.** Menjumlahkan numerator dua indikator dengan populasi
berbeda menghasilkan angka tanpa makna, jadi dashboard tidak menyediakan grand total.

---

## Otomatisasi

`data/indicators.csv` diperbarui otomatis oleh Power Automate. Langkah lengkapnya ada di
**`POWER_AUTOMATE.md`**.

Halaman mencoba `data/indicators.csv` lebih dulu, dan jatuh ke `data/indicators.js` kalau
CSV belum ada atau gagal dibaca — jadi flow bisa dinyalakan kapan saja tanpa risiko halaman
mati. Chip di kanan atas menunjukkan sumbernya: **CSV · 506 baris** atau nomor versi.

Pembaca CSV-nya mendeteksi sendiri pemisah kolom (Tab, titik koma, koma) dan pemisah
desimal (titik atau koma), memetakan kolom lewat **nama header** bukan posisi, dan
memperlakukan nilai error Excel seperti `#VALUE!` sebagai kosong.

---

## Peringatan data: 9 indikator menghasilkan weighted di atas 100%

Σ numerator melebihi Σ denominator, yang mustahil untuk sebuah proporsi.
**Angka-angka itu tidak boleh dikutip** sebelum sumbernya diperbaiki. Ada di
**28 baris**, dari dua sebab berbeda:

**Jakarta Utara — evaluation ~13× terlalu besar.** `IND-501` sampai `IND-503`:
Numerator LOP 2.957.373 dengan Denominator LOP 221.803, dan `% LOP` di sumbernya
memang tertulis **13,3333** (1333%). Tiga indikator berbeda memakai angka yang sama
persis, jadi kemungkinan besar satu angka populasi ter-paste ke beberapa indikator.
Ini yang menaikkan L1 OIOS 107, L1 OIOS 145, L1 OIOS 60, L1 OIOS 63, dan
`Hope rooted in Love — adults` sampai 355–550%.

**28 baris baseline dengan numerator tepat 2× denominator.** Misal `IND-013`
3.081 ÷ 1.540. Ini baris-baris yang `% Baseline`-nya tertulis **2** (200%), tersebar
di indikator `adolescent who married`, `self-efficacy`, dan `peaceful relations`.
Akibatnya weighted baseline keempat indikator itu 127–190%.

Metode weighted-lah yang memunculkan keduanya. Rata-rata persentase akan
menyembunyikannya, karena `% LOP` per baris tampak wajar sendiri-sendiri.

Setelah 9 indikator itu dikeluarkan, **16 indikator** punya delta yang sah:
peningkatan terbesar `L1 OIOS #119` guru dengan pedagogi berpusat anak **+84,2pp**,
penurunan terbesar `OIOS 167` Hope rooted in Love anak **−39,0pp**.

---

## Filter global

Satu panel mengendalikan KPI, tabel, dan kedua grafik sekaligus — tidak perlu refresh.

- **Periode**: tombol ON/OFF `Baseline` · `Evaluation` · `Both` (default). Memilih satu
  periode menyembunyikan kolom dan seri periode lainnya, dan grafik Delta ikut hilang
  karena delta butuh keduanya.
- **Zonal**, **Area Program**, **Outcome**, **Indicator** — pilihan berganda.
  Daftar AP mengikuti Zonal yang dipilih; daftar Indicator mengikuti Outcome.
- **Bersihkan semua** mengembalikan ke seluruh dataset.

---

## Kartu

**KPI**: Total AP · Total Zonal · Total Outcome · Total Indicators.
`Average Delta (pp)` tidak ada — rata-rata delta antar indikator dengan populasi
berbeda tidak punya arti.

**Ringkasan arah perubahan**: jumlah indikator Meningkat · Tetap · Menurun ·
Belum bisa dibandingkan. Batas "tetap" adalah |delta| ≤ 0,1pp.

Warna grafik Delta mengikuti **arah angka**, sesuai permintaan: hijau naik, abu-abu
tetap, merah turun. Indikator berarah Turun ditandai **↓** — pada indikator itu
merah berarti angkanya turun, dan itu hasil yang baik. Penandanya perlu, karena
stunting dan perkawinan anak yang menurun adalah keberhasilan.

---

## Update data

Tombol **⤓ Impor data** → tempel blok dari sheet `Indicators` → **Periksa** →
**Muat data** → **⤒ Simpan file data** → commit `data/indicators.js`.

Sudah menyesuaikan layout **15 kolom** Master (8): `AP Decision` dan
`AP AIM+ 2026 >= Threshold?` tidak ada lagi, jadi kolomnya juga hilang dari
halaman SUMMARY.

Angka boleh berkoma maupun bertitik — pemisah desimal dideteksi sekali untuk seluruh
blok. Sel berisi baris baru tetap terbaca. Nilai tidak pernah dikoreksi otomatis.

**Catatan Master (8):** 27 baris `Row ID`-nya bernilai `0`. Karena Row ID harus unik,
baris itu diberi ID sintetis `IND-X###` supaya tetap bisa dilacak. Sebaiknya
diperbaiki di sumbernya.

---

## Publikasi

Upload seluruh isi folder ke root repository → Settings → Pages →
Deploy from a branch → `main` / `(root)`.

```
index.html
assets/app.css · assets/app.js · assets/logo.svg (opsional, logo resmi WVI)
data/config.js · data/indicators.js · data/asumsi.js · data/pemetaan.js
```

Kode akses `wvipearl` adalah pagar sopan, bukan pengamanan: repository publik, jadi
folder `data/` bisa dibaca langsung di GitHub tanpa melewati layar itu.

---

Versi `v3.4` · 2 Agustus 2026 · data dari Master (9), pembaca CSV untuk pipeline Power Automate.

---

## Struktur berkas (v3.4)

```
index.html
.nojekyll
README.md

assets/
  app.js              engine, render seluruh sheet
  app.css             design token system
  a11y-patch.css      kontras WCAG AA, ukuran font minimum, tooltip, skeleton
  a11y-patch.js       aria-label, focus trap modal, pengawas boot

data/
  config.js           dibaca app.js
  indicators.js       dibaca app.js
  asumsi.js           dibaca app.js
  pemetaan.js         dibaca app.js
  master.js           ARSIP SUMBER, tidak dibaca app.js
  ap-register.js      ARSIP SUMBER, tidak dibaca app.js
  decisions.js        ARSIP SUMBER, tidak dibaca app.js
```

### Tentang tiga berkas arsip

`app.js` hanya merujuk empat global: `WVI_CONFIG`, `WVI_INDICATORS`, `WVI_ASUMSI`,
dan `WVI_PEMETAAN`. Tiga berkas lain tetap disimpan karena memuat jejak asal angka
— `master.js` adalah 394 baris submission per Area Programme, `decisions.js` adalah
log keputusan — tapi dashboard tidak membacanya.

**Mengedit ketiga berkas itu tidak akan mengubah tampilan dashboard.** Perubahan
harus diregenerasi ke `indicators.js` lewat sheet **Reports → Berkas untuk
repository**, lalu berkas hasilnya di-commit ke `data/`.

### Perubahan di v3.4

- kartu `Indicators Requiring Attention` dihapus dari Dashboard
- **perbaikan bug**: kotak sorot baris chart (`.chit`) tampil sebagai blok hitam.
  Penyebabnya `fill` hanya diatur di CSS, sementara nilai awal `fill` di SVG adalah
  hitam — jadi saat `app.css` belum ter-push atau masih kena cache, seluruh baris
  chart tertutup. Sekarang `fill` ditulis sebagai atribut di markup dan CSS hanya
  mengatur keadaan hover
- penanda versi `?v=3.4` pada semua asset di `index.html`, supaya browser dan CDN
  GitHub Pages tidak menyajikan berkas lama. **Naikkan `var V` di `index.html`
  setiap kali `app.js` atau `app.css` berubah.**

### Perubahan di v3.3

- sheet `Summary` diganti nama jadi **List Indicator AP**
- chart `Baseline vs Threshold per indikator` dihapus dari sheet itu
- `Outcome Summary` disejajarkan berdampingan dengan `Performance Summary`
- ketiga chart Dashboard jadi interaktif: sorot baris saat hover, tooltip kustom
  berisi nama indikator penuh dan angkanya, animasi masuk bertahap. Atribut
  `title=` bawaan tidak dipakai karena lambat dan tidak jalan di perangkat sentuh
- `app.css` §10 ditambahkan untuk gaya interaksi chart

### Perubahan di v3.2

- `analysis`, `mapping`, dan `assumptions` dihapus dari array `SHEETS`
- kartu `Executive Insights` dihapus dari Dashboard
- `Evaluation vs Threshold` dan `Delta per Indicator` melebar ke `c12`
- `Top Performing` dan `Lowest Performing` dipindah berdampingan setelah Delta
- helper `svgWrap()` baru: label chart rata kiri, wrap dua baris, diambil dari teks
  indikator penuh alih-alih `short` yang sudah dipangkas `…` di `data/config.js`.
  Versi lama memakai `text-anchor="end"` sehingga label panjang terpotong di tepi
  kiri viewBox
- `a11y-patch.css` dan `a11y-patch.js` ditambahkan
