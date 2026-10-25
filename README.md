# Bot CS WhatsApp + Katalog Google Sheets

Bot customer service untuk WhatsApp. Bot menjawab pertanyaan produk dari katalog
(Google Sheets atau file Excel di Google Drive), mencatat pesanan, mengirim gambar
QRIS, dan meneruskan chat ke admin.

Ini bukan website dan tidak perlu hosting web/domain. Bot berjalan sebagai program
Node.js di komputer atau VPS, lalu terhubung ke WhatsApp melalui scan QR.

## Fitur

- Menjawab harga, stok, foto produk, dan biaya kirim dari katalog.
- Pencarian produk lokal (cepat) + tool `cari_produk`, `cek_stok`, `cari_alternatif`.
- Katalog otomatis diperbarui saat file di Google Drive berubah.
- Sapaan otomatis, menu bantuan, dan info di luar jam operasional.
- Alur pesanan: form -> catatan -> QRIS -> pelanggan ketik `SUDAH` -> admin.
- Pesanan dicatat ke Google Sheet (status `Menunggu Verifikasi`) dan `orders.jsonl`.
- Notifikasi admin + bot pause 30 menit per nomor setelah pembayaran diklaim.
- Alih ke admin manusia dengan kata kunci (admin, komplain, refund, dll).
- Gambar QRIS diverifikasi SHA-256 agar tidak bisa diganti diam-diam.

## Yang perlu disiapkan

1. Komputer/VPS dengan **Node.js 24 atau lebih baru** (disarankan Node.js 22/24 LTS).
2. Akun Google untuk Google Drive dan Google Sheets.
3. Nomor WhatsApp khusus bot (disarankan nomor terpisah, bukan nomor pribadi).
4. **Gemini API Key** dari Google AI Studio.
5. File katalog (`.xlsx`) atau Google Sheets, dan satu Google Sheets untuk pesanan.

## Struktur file penting

```
gemini-wa-cs-bot/
  index.js              kode utama bot
  package.json          daftar dependency
  config.json           konfigurasi utama (RAHASIA, jangan di-commit)
  credentials.json      kunci Google Service Account (RAHASIA)
  .env                  QRIS dan variabel rahasia (RAHASIA)
  .env.example          contoh isi .env
  baileys_auth/         sesi login WhatsApp (RAHASIA, dibuat otomatis)
  auto_message_state.json  state sapaan/pause/draft (dibuat otomatis)
  orders.jsonl          cadangan pesanan lokal (dibuat otomatis)
  images/               cache gambar produk (dibuat otomatis)
```

File bertanda RAHASIA sudah masuk `.gitignore`. Jangan pernah mengunggahnya ke
GitHub atau membagikannya.

---

## Langkah 1 - Buat Google Service Account

Service Account adalah "akun robot" yang dipakai bot untuk membaca Drive/Sheets.

1. Buka <https://console.cloud.google.com/> dan login.
2. Buat project baru (nama bebas), lalu pilih project tersebut.
3. Buka **APIs & Services > Library**. Aktifkan dua API:
   - **Google Drive API**
   - **Google Sheets API**
4. Buka **IAM & Admin > Service Accounts > Create Service Account**.
   Nama bebas, misalnya `wa-cs-bot`. Lanjutkan sampai **Done**.
5. Klik service account tersebut > tab **Keys** > **Add key** >
   **Create new key** > pilih **JSON** > **Create**.
6. File JSON akan terunduh. Ganti namanya menjadi `credentials.json` dan letakkan
   di folder project (satu folder dengan `index.js`).
7. Catat email service account, bentuknya seperti:
   `nama@project-id.iam.gserviceaccount.com`
   Email ini dipakai untuk membagikan file Google Drive/Sheets ke bot.

---

## Langkah 2 - Siapkan katalog

Bot bisa membaca dua jenis file:

- Google Sheets asli, atau
- file `.xlsx` yang diunggah ke Google Drive.

### Nama tab dan kolom

Di dalam file harus ada tab (sheet) berikut:

**Tab `Katalog Barang`** - baris pertama adalah judul kolom:

| Kolom | Wajib | Keterangan |
|---|---|---|
| `nama_barang` | ya | nama produk |
| `kategori` | ya | contoh: Elektronik |
| `subkategori` | tidak | contoh: Aksesoris HP |
| `merek` | tidak | nama merek |
| `deskripsi` | tidak | penjelasan produk |
| `harga_idr` | ya | harga, boleh `Rp 243,000` |
| `stok` | ya | angka stok |
| `berat_gram` | tidak | berat dalam gram |
| `status` | tidak | `Aktif` / `Nonaktif` |
| `drive_file_id_gambar` | tidak | ID file gambar di Drive (opsional) |
| `sku` | tidak | kode produk, contoh `SKU-ELE-0001` |
| `id_barang` | tidak | ID internal |

**Tab `Biaya Kirim`** - baris pertama judul kolom:

| Kolom | Keterangan |
|---|---|
| `kurir` | nama kurir |
| `layanan` | jenis layanan |
| `tarif_1kg` ... `tarif_6kg` | tarif per berat |

### Bagikan file ke service account

1. Buka file katalog di Google Drive.
2. Klik **Share** / **Bagikan**.
3. Tambahkan email service account dari Langkah 1 (misalnya
   `nama@project-id.iam.gserviceaccount.com`) sebagai **Viewer**.
4. Simpan.

### Ambil ID file katalog

Lihat URL file di browser:

```
https://docs.google.com/spreadsheets/d/ID_PANJANG_INI/edit
https://drive.google.com/file/d/ID_PANJANG_INI/view
```

Bagian `ID_PANJANG_INI` itulah `spreadsheetId` untuk katalog.

---

## Langkah 3 - Siapkan Google Sheets untuk pesanan

1. Buat Google Sheets baru, beri nama bebas (misalnya `Pesanan Toko`).
2. Bagikan ke email service account sebagai **Editor** (bukan Viewer).
3. Ambil ID dari URL-nya, sama seperti Langkah 2.
4. Tab bernama `order` dan baris judul kolom akan dibuat otomatis oleh bot.
   Kolomnya: `waktu, nomor_wa, nama, produk, qty, alamat, kota, catatan, metode, status`.

Jika tab `order` sudah ada tetapi kolomnya berbeda, bot tetap menulis sesuai
urutan kolom di atas. Sebaiknya biarkan bot membuat tabnya sendiri.

---

## Langkah 4 - Ambil Gemini API Key

1. Buka <https://aistudio.google.com/apikey>.
2. Klik **Create API key**, pilih project Google Cloud yang sama (atau project baru).
3. Copy API key. Simpan sementara di catatan aman.

Catatan: kunci ini berbeda dari `credentials.json`. Jangan tertukar.

---

## Langkah 5 - Install di komputer

### 5.1 Install Node.js

- Windows/macOS: unduh installer LTS dari <https://nodejs.org/>, jalankan, klik Next
  sampai selesai.
- Linux: ikuti bagian VPS di bawah.

Buka **Terminal** (macOS/Linux) atau **PowerShell** (Windows), lalu cek:

```bash
node -v
npm -v
```

Harus muncul versi Node.js 20 atau lebih baru.

### 5.2 Ambil kode bot

Pilihan A - dengan Git:

```bash
git clone https://github.com/Laperrman123/gemini-wa-cs-bot.git
cd gemini-wa-cs-bot
```

Pilihan B - tanpa Git: buka halaman GitHub repo, klik **Code > Download ZIP**,
extract, lalu buka terminal di folder hasil extract.

### 5.3 Install dependency

```bash
npm install
```

Tunggu sampai selesai. Jika ada peringatan `npm warn`, biasanya aman.

### 5.4 Buat `config.json`

Buat file baru bernama `config.json` di folder project, isi minimal seperti ini
(ganti bagian dalam kurung siku):

```json
{
  "geminiApiKey": "ISI_GEMINI_API_KEY",
  "model": "gemini-3.5-flash-lite",
  "adminNumber": "62812xxxxxxx",
  "spreadsheetId": "ID_KATALOG",
  "sheetRange": "Katalog Barang, Biaya Kirim",
  "credentialsPath": "./credentials.json",
  "orders": {
    "enabled": true,
    "spreadsheetId": "ID_SHEET_PESANAN",
    "sheetName": "order",
    "confirmTimeoutMinutes": 120
  },
  "payment": {
    "enabled": true,
    "method": "QRIS"
  }
}
```

Keterangan:

- `adminNumber`: nomor WhatsApp admin format internasional tanpa `+`,
  contoh `6281223344556`. Jika diawali `0`, bot otomatis mengubahnya menjadi `62`.
- `model`: `gemini-3.5-flash-lite` direkomendasikan karena cepat. Bisa diganti
  `gemini-3.5-flash` atau `gemini-3.8-flash`, tetapi model besar bisa lambat.
  Jika `model` tidak diisi, bot akan menanyakannya dan memakai default
  `gemini-3.8-flash`; sebaiknya tetap isi `gemini-3.5-flash-lite` seperti contoh.
- `orders.enabled`: `true` untuk menulis pesanan ke Google Sheets. Pesanan selalu
  dicatat ke `orders.jsonl` sebagai cadangan.
- `orders.sheetName`: nama tab tujuan. Boleh diganti; jika tidak diisi, default-nya
  `Pesanan`. Contoh di atas memakai `order`.
- `handoff`, `autoMessage`, dan `payment` yang tidak ditulis akan memakai nilai
  default. Contoh lengkap ada di bagian **Konfigurasi lanjutan**.

Pastikan format JSON valid: tanda kutip lurus `"`, tidak ada koma di akhir,
tidak ada komentar `//`.

### 5.5 Buat `.env` untuk QRIS

- Windows (PowerShell): `copy .env.example .env`
- Linux/macOS: `cp .env.example .env`

Buka `.env` lalu isi:

```env
QRIS_IMAGE_URL=https://link-gambar-qris-anda
QRIS_IMAGE_SHA256=hash-dari-gambar-qris
```

Cara mengisi QRIS dibahas di **Langkah 6**. Jika belum siap, boleh dikosongkan
dulu; bot akan mengirim pesan bahwa QRIS akan dikirim admin secara manual.

### 5.6 Jalankan bot

```bash
npm start
```

Saat pertama kali berjalan, bot mungkin menanyakan beberapa konfigurasi yang
belum ada, lalu menyimpannya ke `config.json`.

Setelah itu akan muncul **QR code di terminal**:

```
[!] Scan QR code ini dengan aplikasi WhatsApp Anda
```

Di WhatsApp nomor bot: **Pengaturan > Perangkat tertaut > Tautkan perangkat**,
lalu scan QR di layar terminal.

Jika berhasil akan muncul:

```
[+] WhatsApp bot berhasil terhubung dan siap melayani customer!
```

Coba kirim chat ke nomor bot, misalnya: `ada earphone?`

---

## Langkah 6 - Mengatur gambar QRIS

Bot mengirim gambar QRIS setelah pelanggan mengisi catatan pesanan, lalu pelanggan
mengetik `SUDAH`.

Pilih salah satu sumber gambar:

### Opsi A - Link Google Drive (disarankan jika sudah punya di Drive)

1. Upload gambar QRIS ke Google Drive.
2. Share file ke email service account sebagai **Viewer**.
3. Copy link share-nya, contoh:
   `https://drive.google.com/file/d/ID_FILE/view?usp=sharing`
4. Isi `QRIS_IMAGE_URL` di `.env` dengan link tersebut.

### Opsi B - Link gambar langsung (direct link)

Gunakan URL yang langsung menampilkan gambar (biasanya berakhiran `.png` atau
`.jpg`), bukan halaman viewer. Contoh:

```env
QRIS_IMAGE_URL=https://contoh.com/qris.png
```

### Opsi C - File lokal

Simpan gambar di folder project, misalnya `assets/qris.png`, lalu:

```env
QRIS_IMAGE_URL=./assets/qris.png
```

### Wajib: isi hash QRIS

Hash membuat bot menolak mengirim QRIS jika gambarnya berubah. Hitung hash:

- Windows (PowerShell):
  ```powershell
  Get-FileHash -Algorithm SHA256 .\assets\qris.png
  ```
- Linux/macOS:
  ```bash
  sha256sum assets/qris.png
  ```

Copy nilai hash (64 karakter) ke `.env`:

```env
QRIS_IMAGE_SHA256=hasil_hash_tanpa_spasi
```

Jika memakai link, unduh dulu gambarnya, hitung hash file hasil unduhan, lalu
isi nilainya. Setelah mengganti gambar QRIS, **wajib** memperbarui hash dan
restart bot. Jika hash tidak cocok, bot tidak mengirim QRIS dan memberi tahu admin.

Catatan: gambar QRIS statis tidak bisa menampilkan nominal otomatis. Admin tetap
memverifikasi pembayaran secara manual.

---

## Langkah 7 - Install di VPS (Ubuntu/Debian)

VPS yang disarankan: minimal 1 CPU, RAM 1 GB, Ubuntu 22.04/24.04. Bot hanya
membuka koneksi keluar, jadi **tidak perlu membuka port** apa pun.

### 7.1 Masuk ke VPS

```bash
ssh user@IP_VPS
```

### 7.2 Install Node.js dan Git

```bash
sudo apt update && sudo apt upgrade -y
curl -fsSL https://deb.nodesource.com/setup_22.x | sudo -E bash -
sudo apt install -y nodejs git
node -v
```

### 7.3 Ambil kode dan install dependency

```bash
git clone https://github.com/Laperrman123/gemini-wa-cs-bot.git
cd gemini-wa-cs-bot
npm install
```

### 7.4 Siapkan file konfigurasi

```bash
nano config.json
# tempel isi config.json seperti Langkah 5.4, simpan dengan Ctrl+O, Enter, Ctrl+X

cp .env.example .env
nano .env
# isi QRIS_IMAGE_URL dan QRIS_IMAGE_SHA256
```

Letakkan `credentials.json` di folder ini. Caranya bisa dengan `scp` dari komputer:

```bash
# jalankan di komputer, bukan di VPS
scp credentials.json user@IP_VPS:/home/user/gemini-wa-cs-bot/
```

### 7.5 Uji jalan dulu

```bash
npm start
```

Scan QR yang muncul dengan WhatsApp. Jika sudah terhubung, tekan `Ctrl+C` untuk
berhenti sementara, lalu lanjut ke langkah berikut agar bot jalan 24 jam.

### 7.6 Jalankan permanen dengan PM2

PM2 membuat bot tetap hidup setelah terminal ditutup dan otomatis restart jika
crash atau VPS reboot.

```bash
sudo npm install -g pm2
pm2 start npm --name wa-cs-bot -- start
pm2 logs wa-cs-bot
```

Jika pada langkah 7.5 Anda sudah scan QR dan berhasil terhubung, bot akan langsung
jalan tanpa QR baru. Jika belum, scan QR yang muncul di `pm2 logs`. Setelah
terhubung, tekan `Ctrl+C` untuk keluar dari tampilan log (bot tetap berjalan).

```bash
pm2 save
pm2 startup
```

Perintah `pm2 startup` akan menampilkan satu perintah `sudo ...`; copy dan
jalankan perintah tersebut agar PM2 aktif otomatis saat VPS reboot.

### 7.7 Perintah PM2 sehari-hari

```bash
pm2 status                  # cek status
pm2 logs wa-cs-bot          # lihat log
pm2 restart wa-cs-bot       # restart setelah ubah config/.env
pm2 stop wa-cs-bot          # hentikan sementara
pm2 delete wa-cs-bot        # hapus dari PM2
```

### Alternatif systemd (opsional)

Jika tidak ingin PM2, buat file `/etc/systemd/system/wa-cs-bot.service`:

```ini
[Unit]
Description=Bot CS WhatsApp
After=network.target

[Service]
Type=simple
User=user
WorkingDirectory=/home/user/gemini-wa-cs-bot
ExecStart=/usr/bin/npm start
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

Lalu:

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now wa-cs-bot
sudo journalctl -u wa-cs-bot -f
```

---

## Alur kerja bot

1. Pelanggan bertanya produk. Bot mencari produk relevan dari katalog dan menjawab.
2. Jika pelanggan ingin memesan, bot mengumpulkan data: nama, produk, jumlah,
   alamat, kota.
3. Setelah data lengkap, bot menampilkan **Form Pesanan** dan bertanya:
   "Ada catatan tambahan? Ketik `-` jika tidak ada."
4. Pelanggan mengisi catatan (atau `-`). Bot mengirim ringkasan + **gambar QRIS**.
   Metode otomatis `QRIS`.
5. Pelanggan membayar lalu mengetik `SUDAH` (atau "sudah bayar/transfer").
6. Bot:
   - mencatat pesanan ke Google Sheets dengan status `Menunggu Verifikasi`,
   - mengirim pesan terima kasih,
   - memberi notifikasi ke admin,
   - mem-pause bot untuk nomor itu selama 30 menit.
7. Selama pause, pesan pelanggan diteruskan ke admin. Admin memverifikasi
   pembayaran dan memproses pesanan.

Aturan lain:

- Draft pesanan kedaluwarsa setelah 120 menit (`orders.confirmTimeoutMinutes`).
- Pelanggan bisa membatalkan dengan `batal` atau `tidak`.
- Kata kunci `admin`, `komplain`, `refund`, dan sejenisnya langsung mengalihkan
  chat ke admin (pause 30 menit).
- Bot **tidak** menandai pesanan lunas otomatis. Status `Menunggu Verifikasi`
  harus dicek admin.

---

## Konfigurasi lanjutan (`config.json`)

Contoh lengkap:

```json
{
  "geminiApiKey": "ISI_GEMINI_API_KEY",
  "model": "gemini-3.5-flash-lite",
  "adminNumber": "62812xxxxxxx",
  "spreadsheetId": "ID_KATALOG",
  "sheetRange": "Katalog Barang, Biaya Kirim",
  "credentialsPath": "./credentials.json",
  "autoMessage": {
    "greeting": "Halo kak! Selamat datang. Silakan tanya harga, stok, atau minta foto produk ya.",
    "menuKeywords": ["menu", "katalog", "bantuan", "help", "info"],
    "menuText": "Yang bisa saya bantu: tanya harga/stok, minta foto produk, atau tanya biaya kirim.",
    "businessHours": {
      "start": "08:00",
      "end": "21:00",
      "timezone": "Asia/Jakarta"
    },
    "offHoursText": "Mohon maaf kak, saat ini di luar jam operasional. Pesan Anda akan diteruskan ke admin.",
    "offHoursNotifyOncePerDay": true
  },
  "orders": {
    "enabled": true,
    "spreadsheetId": "ID_SHEET_PESANAN",
    "sheetName": "order",
    "confirmTimeoutMinutes": 120
  },
  "handoff": {
    "enabled": true,
    "pauseMinutes": 30,
    "keywords": ["admin", "manusia", "komplain", "refund", "kecewa"],
    "message": "Baik kak, saya sambungkan ke admin ya. Mohon tunggu sebentar."
  },
  "payment": {
    "enabled": true,
    "method": "QRIS",
    "qrisImageUrl": "",
    "qrisImageSha256": "",
    "qrisCaption": "Silakan scan QRIS berikut untuk pembayaran.",
    "thankYouText": "Terima kasih kak! Pembayaran akan diverifikasi admin."
  }
}
```

`QRIS_IMAGE_URL` dan `QRIS_IMAGE_SHA256` di `.env` akan menimpa nilai
`payment.qrisImageUrl` dan `payment.qrisImageSha256`.

---

## Perawatan

### Mengubah katalog

Edit Google Sheets atau file `.xlsx` di Drive. Bot mengecek perubahan file setiap
sekitar 1 detik saat ada pesan masuk, lalu memuat ulang katalog otomatis.

Jika memakai file `.xlsx`, unggah ulang file ke Drive (timpa file lama dengan ID
yang sama). Jika membuat file baru, ganti `spreadsheetId` di `config.json`.

### Update kode dari GitHub

```bash
git pull
npm install
pm2 restart wa-cs-bot      # atau: pm2 restart wa-cs-bot --update-env
```

### Backup

Yang perlu dibackup: `config.json`, `credentials.json`, `.env`, dan
`auto_message_state.json`. File `baileys_auth/` berisi sesi login WhatsApp: jangan
dibagikan, tetapi boleh dibackup agar tidak perlu scan QR ulang.

### Ganti nomor WhatsApp

1. Hentikan bot.
2. Hapus folder `baileys_auth/`.
3. Jalankan bot lagi dan scan QR dengan nomor baru.

---

## Troubleshooting

| Gejala | Penyebab | Solusi |
|---|---|---|
| `credentials tidak ditemukan` | file belum ada / path salah | letakkan `credentials.json` di folder project, cek `credentialsPath` |
| `403` atau `The caller does not have permission` | file belum dibagikan | share katalog ke service account (Viewer), sheet pesanan (Editor) |
| `File ... bukan Google Spreadsheet atau Excel` | ID salah / file bukan Sheets/xlsx | copy ulang ID dari URL file |
| `QRIS_IMAGE_URL belum diatur` | `.env` kosong | isi `QRIS_IMAGE_URL` lalu restart |
| `URL bukan gambar langsung (content-type: text/html)` | link share/viewer, bukan gambar | share file Drive ke service account, atau pakai direct image link |
| `Hash QRIS tidak cocok` | gambar berubah / hash lama | hitung ulang SHA-256, update `.env`, restart |
| `Gambar bukan PNG/JPEG/WEBP` | file bukan gambar | ganti dengan file gambar asli |
| `429` / `503 high demand` | kuota/model sedang sibuk | bot otomatis pindah model; tunggu, atau ganti `model` ke `gemini-3.5-flash-lite` |
| Balasan lambat | model besar sedang sibuk | gunakan `gemini-3.5-flash-lite` |
| QR tidak muncul di VPS | log tidak terlihat | `pm2 logs wa-cs-bot`, atau jalankan `npm start` dulu |
| WhatsApp logout / minta scan ulang | sesi tidak valid | hapus `baileys_auth/`, jalankan ulang, scan QR |
| Admin tidak menerima notifikasi | nomor salah format | gunakan format `62812...` (tanpa `+` dan tanpa `0` di depan) |
| Sapaan/emoji berantakan | teks di `config.json` rusak encoding | edit ulang teks di `config.json` |

---

## Keamanan

- Jangan unggah `config.json`, `credentials.json`, `.env`, `baileys_auth/`,
  `orders.jsonl`, atau `auto_message_state.json` ke GitHub.
- Di Linux, batasi izin file:
  ```bash
  chmod 600 .env config.json credentials.json
  ```
- Service account cukup diberi akses **Viewer** ke katalog dan file QRIS, dan
  **Editor** hanya ke sheet pesanan.
- Isi `QRIS_IMAGE_SHA256` agar gambar QRIS tidak bisa diganti tanpa terdeteksi.
- Bot tidak pernah menandai pembayaran lunas otomatis. Selalu verifikasi
  pembayaran sebelum memproses pesanan.
- Gunakan nomor WhatsApp khusus toko. Library WhatsApp yang dipakai bersifat
  tidak resmi; ada risiko nomor dibatasi WhatsApp jika dipakai untuk spam.
  Gunakan secara wajar dan patuhi kebijakan WhatsApp.

---

## Pertanyaan umum

**Apakah perlu website, domain, atau hosting?**
Tidak. Bot adalah program Node.js yang berjalan di komputer/VPS dan terhubung ke
WhatsApp. Tidak perlu membuka port atau memasang web server.

**Apakah bot harus selalu menyala?**
Ya, selama ingin melayani pelanggan. Di VPS gunakan PM2/systemd. Di komputer,
biarkan terminal menyala atau gunakan PM2/Task Scheduler.

**Di mana data pesanan disimpan?**
Selalu di `orders.jsonl` (cadangan lokal) dan, jika `orders.enabled` true, juga di
Google Sheets tab `order`.

**Apakah harga dan stok bisa berubah otomatis?**
Ya, selama file katalog di Drive diperbarui. Bot memuat ulang otomatis.

**Bagaimana kalau gambar produk tidak muncul?**
Isi kolom `drive_file_id_gambar` dengan ID file gambar di Drive dan pastikan file
gambar dibagikan ke service account (Viewer).

---

## Checklist sebelum dipakai

- [ ] Node.js 20+ terpasang.
- [ ] `npm install` selesai tanpa error.
- [ ] `credentials.json` ada dan valid.
- [ ] Katalog dibagikan ke service account (Viewer).
- [ ] Sheet pesanan dibagikan ke service account (Editor).
- [ ] `config.json` berisi API key, ID katalog, ID sheet pesanan, nomor admin.
- [ ] `.env` berisi QRIS URL dan hash yang benar.
- [ ] Bot terhubung ke WhatsApp (QR sudah discan).
- [ ] Tes chat: tanya produk, buat pesanan, isi catatan, cek QRIS terkirim.
- [ ] Tes ketik `SUDAH`: pesanan masuk sheet, admin dapat notifikasi, bot pause.
- [ ] PM2/systemd aktif agar bot jalan 24 jam.
