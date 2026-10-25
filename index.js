const fs = require('fs');
const path = require('path');
const readline = require('readline');
const https = require('https');
const crypto = require('crypto');
const pino = require('pino');
const qrcode = require('qrcode-terminal');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const { GoogleGenAI } = require('@google/genai');
const baileys = require('@whiskeysockets/baileys');
require('dotenv').config({ quiet: true });

const makeWASocket = baileys.default || baileys.makeWASocket;
const { useMultiFileAuthState, DisconnectReason } = baileys;

const CONFIG_PATH = path.join(__dirname, 'config.json');
const AUTO_STATE_PATH = path.join(__dirname, 'auto_message_state.json');

const DEFAULT_AUTO_MESSAGE = {
  greeting:
    'Halo kak 👋 Selamat datang! Saya asisten otomatis toko kami. Silakan tanya harga, stok, atau minta foto produk ya.',
  menuKeywords: ['menu', 'katalog', 'bantuan', 'help', 'info'],
  menuText:
    '📋 *Yang bisa saya bantu:*\n• Tanya *harga* / *stok* produk\n• Minta *foto* produk (contoh: "foto Earphone Bluetooth")\n• Tanya *biaya kirim* ke kota Anda\n\nLangsung ketik pertanyaan Anda ya kak 😊',
  businessHours: { start: '08:00', end: '21:00', timezone: 'Asia/Jakarta' },
  offHoursText:
    'Mohon maaf kak, saat ini kami di luar jam operasional toko. Asisten otomatis tetap siap membantu, dan pesan Anda akan diteruskan ke admin untuk ditindaklanjuti saat toko buka 🙏',
  offHoursNotifyOncePerDay: true,
};

function withAutoMessageDefaults(config) {
  const provided = config && typeof config.autoMessage === 'object' && config.autoMessage ? config.autoMessage : {};
  const merged = { ...DEFAULT_AUTO_MESSAGE, ...provided };
  merged.businessHours = {
    ...DEFAULT_AUTO_MESSAGE.businessHours,
    ...(provided.businessHours && typeof provided.businessHours === 'object' ? provided.businessHours : {}),
  };
  if (!Array.isArray(merged.menuKeywords)) merged.menuKeywords = DEFAULT_AUTO_MESSAGE.menuKeywords;
  return merged;
}

const ORDER_LOG_PATH = path.join(__dirname, 'orders.jsonl');const ORDER_HEADERS = ['waktu', 'nomor_wa', 'nama', 'produk', 'qty', 'alamat', 'kota', 'catatan', 'metode', 'status'];
const DEFAULT_ORDERS = { enabled: false, spreadsheetId: '', sheetName: 'Pesanan', confirmTimeoutMinutes: 120 };

function withOrdersDefaults(config) {
  const provided = config && typeof config.orders === 'object' && config.orders ? config.orders : {};
  const merged = { ...DEFAULT_ORDERS, ...provided };
  merged.sheetName = merged.sheetName || DEFAULT_ORDERS.sheetName;
  merged.enabled = Boolean(merged.enabled);
  merged.spreadsheetId = merged.spreadsheetId || '';
  const timeout = Number(merged.confirmTimeoutMinutes);
  merged.confirmTimeoutMinutes = Number.isFinite(timeout) && timeout > 0 ? timeout : DEFAULT_ORDERS.confirmTimeoutMinutes;
  return merged;
}

const DEFAULT_PAYMENT = {
  enabled: true,
  method: 'QRIS',
  qrisImageUrl: '',
  qrisImageSha256: '',
  qrisCaption: 'Silakan scan QRIS berikut untuk melakukan pembayaran.',
  thankYouText:
    'Terima kasih kak! Pembayaran Anda sudah kami terima dan akan diverifikasi admin. Pesanan diproses setelah pembayaran terkonfirmasi. 🙏',
};

function withPaymentDefaults(config) {
  const provided = config && typeof config.payment === 'object' && config.payment ? config.payment : {};
  const merged = { ...DEFAULT_PAYMENT, ...provided };
  merged.enabled = merged.enabled !== false;
  merged.method = String(merged.method || DEFAULT_PAYMENT.method).trim() || 'QRIS';
  merged.qrisImageUrl = String(process.env.QRIS_IMAGE_URL || merged.qrisImageUrl || '').trim();
  merged.qrisImageSha256 = String(process.env.QRIS_IMAGE_SHA256 || merged.qrisImageSha256 || '').trim();
  return merged;
}

const DEFAULT_HANDOFF = {
  enabled: true,
  pauseMinutes: 30,
  keywords: [
    'admin',
    'manusia',
    'orang asli',
    'cs manusia',
    'agen',
    'komplain',
    'keluhan',
    'refund',
    'pengembalian dana',
    'lapor',
    'penipuan',
    'penipu',
    'marah',
    'kecewa',
  ],
  message:
    'Baik kak, saya sambungkan ke admin kami ya 🙏 Mohon tunggu sebentar, admin akan membalas chat ini langsung dari nomor ini.',
};

function withHandoffDefaults(config) {
  const provided = config && typeof config.handoff === 'object' && config.handoff ? config.handoff : {};
  const merged = { ...DEFAULT_HANDOFF, ...provided };
  if (!Array.isArray(merged.keywords)) merged.keywords = DEFAULT_HANDOFF.keywords;
  const minutes = Number(merged.pauseMinutes);
  merged.pauseMinutes = Number.isFinite(minutes) && minutes > 0 ? minutes : DEFAULT_HANDOFF.pauseMinutes;
  merged.enabled = merged.enabled !== false;
  return merged;
}

function isHandoffRequest(text, keywords) {
  const normalized = String(text || '')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!normalized) return false;

  return (keywords || []).some((keyword) => {
    const k = String(keyword || '').toLowerCase().trim();
    if (!k) return false;
    if (k.includes(' ')) return normalized.includes(k);
    return new RegExp(`(^|\\s)${k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(\\s|$)`).test(normalized);
  });
}

function normalizeMessage(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function isOrderConfirmation(text) {
  const raw = String(text || '').trim();
  if (!raw || raw.length > 80 || /\d/.test(raw)) return false;
  if (
    /(ubah|ganti|tukar|tambah|kurang|pindah|jadi|salah|bukan|tapi|jangan|tunggu|belum|batal|nama|alamat|nomor|nomer|telepon|telp|hp|qty|jumlah|produk|barang|kirim ke)/i.test(
      raw
    )
  ) {
    return false;
  }

  const normalized = normalizeMessage(raw);
  if (!normalized) return false;
  if (/^(ya|iya|yup|yes|yoi|betul|benar|bener|ok|oke|okey|okay|konfirmasi|konfirm|setuju|lanjut|gas|sip|siap|deal|pesan|order)(\s|$)/.test(normalized)) {
    return true;
  }
  if (/^(sudah|udah|sdh)?\s*(benar|bener|sesuai|fix)(\s|$)/.test(normalized)) {
    return true;
  }
  return /\b(konfirmasi|confirm)\b/.test(normalized);
}

function isOrderCancellation(text) {
  const raw = String(text || '').trim();
  if (!raw || raw.length > 80) return false;
  const normalized = normalizeMessage(raw);
  return /^(batal|batalkan|cancel|tidak|nggak|enggak|gak|ga|jangan)(\s|$)/.test(normalized);
}

function getServiceAccountEmail(config) {
  try {
    const creds = JSON.parse(fs.readFileSync(path.resolve(__dirname, config.credentialsPath || './credentials.json'), 'utf-8'));
    return creds.client_email || '';
  } catch {
    return '';
  }
}

function formatOrderTime(timezone = 'Asia/Jakarta') {
  const now = new Date();
  const date = new Intl.DateTimeFormat('sv-SE', {
    timeZone: timezone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(now);
  const time = new Intl.DateTimeFormat('en-GB', {
    timeZone: timezone,
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23',
  }).format(now);
  return `${date} ${time}`;
}

async function ensureOrdersSheet(sheets, spreadsheetId, sheetName) {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets.properties(title)',
  });
  const exists = (meta.data.sheets || []).some((s) => s.properties?.title === sheetName);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] },
    });
  }

  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A1:J1`,
  });
  const header = headerRes.data.values?.[0] || [];
  if (header.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A1:J1`,
      valueInputOption: 'RAW',
      requestBody: { values: [ORDER_HEADERS] },
    });
  }
}

const recentOrders = new Map();
const ORDER_DEDUPE_MS = 10 * 60 * 1000;

function orderDedupeKey(jid, order) {
  return `${jid}|${order.nama || ''}|${order.produk || ''}|${order.qty || ''}|${order.alamat || ''}`.toLowerCase();
}

async function logOrder(config, jid, order, status = 'Baru') {
  const ordersConfig = withOrdersDefaults(config);
  const key = orderDedupeKey(jid, order);
  const last = recentOrders.get(key);
  if (last && Date.now() - last < ORDER_DEDUPE_MS) {
    console.log(`[Order] Duplikat diabaikan untuk ${jid}`);
    return { saved: false, duplicate: true };
  }

  const timezone = withAutoMessageDefaults(config).businessHours.timezone;
  const row = [
    formatOrderTime(timezone),
    jid.split('@')[0],
    order.nama || '-',
    order.produk || '-',
    order.qty || '-',
    order.alamat || '-',
    order.kota || '-',
    order.catatan || '-',
    order.metode || '-',
    status,
  ];

  fs.appendFileSync(ORDER_LOG_PATH, JSON.stringify({ ...order, waktu: row[0], nomor_wa: row[1], status }) + '\n');
  recentOrders.set(key, Date.now());
  console.log(`[Order] Pesanan dicatat ke orders.jsonl: ${order.produk} x${order.qty || '-'} (${jid})`);

  if (ordersConfig.enabled && ordersConfig.spreadsheetId) {
    try {
      const auth = new google.auth.GoogleAuth({
        keyFile: path.resolve(__dirname, config.credentialsPath),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });
      const sheets = google.sheets({ version: 'v4', auth });
      await ensureOrdersSheet(sheets, ordersConfig.spreadsheetId, ordersConfig.sheetName);
      await sheets.spreadsheets.values.append({
        spreadsheetId: ordersConfig.spreadsheetId,
        range: `${ordersConfig.sheetName}!A:J`,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [row] },
      });
      console.log(`[Order] Pesanan ditambahkan ke sheet "${ordersConfig.sheetName}"`);
    } catch (err) {
      console.error(`[Order] Gagal menulis ke Google Sheet: ${err.message}`);
    }
  }

  return { saved: true, duplicate: false };
}

const ask = (rl, question, defaultValue = '') =>
  new Promise((resolve) => {
    const prompt = defaultValue ? `${question} [${defaultValue}]: ` : `${question}: `;
    rl.question(prompt, (answer) => resolve((answer.trim() || defaultValue).trim()));
  });

async function ensureConfig() {
  let config = {};
  if (fs.existsSync(CONFIG_PATH)) {
    try {
      config = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf-8'));
    } catch {
      config = {};
    }
  }

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  let updated = false;

  if (!config.geminiApiKey) {
    console.log('\n--- Setup Bot Toko WhatsApp ---');
    config.geminiApiKey = await ask(rl, 'Masukkan Gemini API Key');
    updated = true;
  }

  if (!config.model) {
    config.model = await ask(rl, 'Pilih Model Gemini', 'gemini-3.8-flash');
    updated = true;
  }

  if (config.adminNumber === undefined) {
    const adminInput = await ask(rl, 'Nomor WhatsApp Admin (opsional, untuk notifikasi)', '');
    config.adminNumber = adminInput ? adminInput.replace(/\D/g, '') : '';
    updated = true;
  }

  if (!config.spreadsheetId) {
    config.spreadsheetId = await ask(rl, 'Masukkan Google Spreadsheet ID');
    updated = true;
  }

  if (!config.sheetRange) {
    config.sheetRange = await ask(rl, 'Masukkan Nama Sheet / Range', 'Sheet1!A:Z');
    updated = true;
  }

  if (!config.credentialsPath) {
    config.credentialsPath = await ask(rl, 'Path ke file credentials.json', './credentials.json');
    updated = true;
  }

  const autoMessage = withAutoMessageDefaults(config);
  if (JSON.stringify(config.autoMessage) !== JSON.stringify(autoMessage)) {
    config.autoMessage = autoMessage;
    updated = true;
  }

  const orders = withOrdersDefaults(config);
  if (JSON.stringify(config.orders) !== JSON.stringify(orders)) {
    config.orders = orders;
    updated = true;
  }

  const handoff = withHandoffDefaults(config);
  if (JSON.stringify(config.handoff) !== JSON.stringify(handoff)) {
    config.handoff = handoff;
    updated = true;
  }

  const credAbsPath = path.resolve(__dirname, config.credentialsPath);
  while (!fs.existsSync(credAbsPath)) {
    console.log(`\n[!] File credentials tidak ditemukan di: ${credAbsPath}`);
    console.log('Silakan copy file credentials.json (Google Service Account) ke lokasi tersebut.');
    await ask(rl, 'Tekan [Enter] jika file sudah diletakkan, atau ketik path baru', config.credentialsPath);
  }

  if (updated) {
    fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2));
    console.log('[+] Konfigurasi berhasil disimpan ke config.json\n');
  }

  rl.close();
  return config;
}

const SHEET_MIME = 'application/vnd.google-apps.spreadsheet';
const EXCEL_MIME = /(officedocument\.spreadsheetml|ms-excel|application\/vnd\.oasis\.opendocument\.spreadsheet|text\/csv)/i;

function buildSheetResult(rows, label) {
  if (!rows || rows.length === 0) return { text: `[${label}]: Kosong`, data: [], label };

  const headers = rows[0].map((h) => String(h ?? '').trim());
  const itemsData = rows
    .slice(1)
    .filter((row) => row.some((cell) => String(cell ?? '').trim() !== ''))
    .map((row) => {
      const item = {};
      headers.forEach((h, i) => {
        item[h] = String(row[i] ?? '').trim() || '-';
      });
      return item;
    });

  const text = `### Sheet: ${label} ###\n` + itemsData.map((item) =>
    Object.entries(item).map(([k, v]) => `${k}: ${v}`).join(' | ')
  ).join('\n');

  return { text, data: itemsData, label };
}

function parseRangeName(range) {
  const sheetPart = range.split('!')[0].trim();
  return sheetPart.replace(/^'|'$/g, '').trim();
}

async function loadFromGoogleSheets(sheets, config, ranges) {
  const results = await Promise.all(
    ranges.map(async (range) => {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: config.spreadsheetId,
        range,
      });
      return buildSheetResult(response.data.values || [], range);
    })
  );
  return results;
}

async function loadFromExcelFile(drive, config, ranges) {
  const res = await drive.files.get(
    { fileId: config.spreadsheetId, alt: 'media', supportsAllDrives: true },
    { responseType: 'arraybuffer', headers: { 'Cache-Control': 'no-cache' } }
  );
  const workbook = XLSX.read(Buffer.from(res.data), { type: 'buffer' });

  const wanted = ranges.map(parseRangeName).filter(Boolean);
  let sheetNames = workbook.SheetNames.filter((name) =>
    wanted.some((w) => w.toLowerCase() === name.toLowerCase())
  );
  if (sheetNames.length === 0) sheetNames = workbook.SheetNames;

  return sheetNames.map((name) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
      header: 1,
      raw: false,
      defval: '',
      blankrows: false,
    });
    return buildSheetResult(rows, name);
  });
}

async function loadSheetData(config) {
  const scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive.readonly'];
  const auth = new google.auth.GoogleAuth({
    keyFile: path.resolve(__dirname, config.credentialsPath),
    scopes,
  });

  const sheets = google.sheets({ version: 'v4', auth });
  const drive = google.drive({ version: 'v3', auth });
  const ranges = (config.sheetRange || 'Sheet1!A:Z')
    .split(',')
    .map((r) => r.trim())
    .filter(Boolean);

  const meta = await drive.files.get({
    fileId: config.spreadsheetId,
    fields: 'id,name,mimeType,modifiedTime,md5Checksum,version,size',
    supportsAllDrives: true,
  });

  const mimeType = meta.data.mimeType || '';
 
  let results;
  if (mimeType === SHEET_MIME) {
    results = await loadFromGoogleSheets(sheets, config, ranges);
  } else if (EXCEL_MIME.test(mimeType)) {
    console.log(`[i] File "${meta.data.name}" terdeteksi sebagai file Excel/Office, membaca via unduhan Drive...`);
    results = await loadFromExcelFile(drive, config, ranges);
  } else {
    throw new Error(
      `File dengan ID ${config.spreadsheetId} bukan Google Spreadsheet atau Excel/CSV yang didukung (mimeType: ${mimeType}).`
    );
  }

  const context = results.map(r => r.text).join('\n\n');

  return {
    context,
    sections: results,
    raw: results.flatMap(r => r.data),
    auth,
    mimeType,
    name: meta.data.name,
    modifiedTime: meta.data.modifiedTime,
    md5Checksum: meta.data.md5Checksum,
    signature: [meta.data.modifiedTime, meta.data.md5Checksum, meta.data.version, meta.data.size]
      .filter(Boolean)
      .join('|'),
    contentHash: crypto.createHash('sha256').update(context).digest('hex'),
  };
}

function normalizeCatalogKey(key) {
  return String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function pickCatalogValue(item, candidates) {
  const keys = Object.keys(item || {});
  for (const candidate of candidates) {
    const target = normalizeCatalogKey(candidate);
    const key = keys.find((k) => normalizeCatalogKey(k) === target);
    if (key === undefined) continue;
    const value = item[key];
    if (value !== undefined && value !== null && String(value).trim() !== '' && String(value).trim() !== '-') {
      return String(value).trim();
    }
  }
  return '';
}

function parseCatalogNumber(value) {
  const digits = String(value ?? '').replace(/[^\d]/g, '');
  return digits ? Number(digits) : NaN;
}

function normalizeProductLabel(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

const SIZE_PATTERN = /(\d+(?:[.,]\d+)?\s?(?:mm|cm|ml|gram|gr|kg|gb|tb|mah|watt|w|inch|in|m|l))\b/i;

function normalizeCatalogItem(item) {
  const nama = pickCatalogValue(item, ['nama_barang', 'nama_produk', 'nama', 'produk']);
  const varian = pickCatalogValue(item, ['varian', 'variasi', 'ukuran', 'size', 'tipe', 'type']);
  const sizeFromName = (nama.match(SIZE_PATTERN) || [])[0] || '';

  return {
    sku: pickCatalogValue(item, ['sku', 'kode_sku', 'id_barang', 'id_produk', 'id']),
    nama,
    kategori: pickCatalogValue(item, ['kategori']),
    subkategori: pickCatalogValue(item, ['subkategori', 'sub_kategori', 'subkategori_produk']),
    merek: pickCatalogValue(item, ['merek', 'brand']),
    harga: parseCatalogNumber(pickCatalogValue(item, ['harga_idr', 'harga', 'price'])),
    stok: parseCatalogNumber(pickCatalogValue(item, ['stok', 'stock', 'qty_stok'])),
    berat_gram: parseCatalogNumber(pickCatalogValue(item, ['berat_gram', 'berat', 'weight'])),
    status: pickCatalogValue(item, ['status']) || 'Aktif',
    varian: varian || sizeFromName,
    deskripsi: pickCatalogValue(item, ['deskripsi', 'description']),
    gambar: pickCatalogValue(item, ['drive_file_id_gambar', 'gambar', 'foto', 'image']),
  };
}

function createCatalogIndex(raw) {
  const items = (raw || []).map(normalizeCatalogItem).filter((item) => item.nama);
  const bySku = new Map();
  for (const item of items) {
    if (item.sku) bySku.set(item.sku.toLowerCase(), item);
  }
  return { items, bySku };
}

function isProductActive(item) {
  const status = String(item.status || '').trim();
  if (!status) return true;
  if (/non\s*aktif|tidak\s*aktif|inactive|discontinued|arsip/i.test(status)) return false;
  return /aktif|active|ready|tersedia/i.test(status);
}

function compactProduct(item) {
  return {
    sku: item.sku || '-',
    nama: item.nama,
    varian: item.varian || '-',
    kategori: item.kategori || '-',
    subkategori: item.subkategori || '-',
    merek: item.merek || '-',
    harga: item.harga,
    stok: item.stok,
    berat_gram: item.berat_gram,
    status: item.status || '-',
  };
}

const RETRIEVAL_STOPWORDS = new Set([
  'ada', 'adakah', 'apa', 'apakah', 'berapa', 'harga', 'harganya', 'stok', 'stoknya',
  'saya', 'aku', 'mau', 'ingin', 'pesan', 'beli', 'order', 'kak', 'min', 'gan', 'bang',
  'mas', 'mbak', 'tolong', 'bisa', 'gak', 'tidak', 'nggak', 'dong', 'ya', 'yang',
  'untuk', 'dengan', 'dan', 'atau', 'di', 'ke', 'dari', 'itu', 'ini', 'produk',
  'barang', 'toko', 'katalog', 'menu', 'info', 'tersedia', 'masih', 'punya', 'jual',
  'jualan', 'cari', 'minta', 'lihat', 'foto', 'gambar', 'ukurannya', 'ukuran', 'varian',
  'model', 'seri', 'tipe', 'halo', 'hai', 'hi', 'pagi', 'siang', 'sore', 'malam',
  'assalamualaikum', 'salam', 'permisi', 'kakak',
]);

function catalogSearchTokens(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\s]/gu, ' ')
    .split(/\s+/)
    .filter((token) => token.length >= 2 && !RETRIEVAL_STOPWORDS.has(token));
}

function searchCatalogProducts(index, query, options = {}) {
  const tokens = catalogSearchTokens(query);
  if (tokens.length === 0) return [];

  const phrase = tokens.join(' ');
  const kategori = String(options.kategori || '').toLowerCase().trim();
  const merek = String(options.merek || '').toLowerCase().trim();
  const limit = Math.min(Math.max(Number(options.limit) || 12, 1), 25);
  const minRaw = options.harga_min;
  const maxRaw = options.harga_max;
  const hargaMin = minRaw === undefined || minRaw === '' ? NaN : Number(minRaw);
  const hargaMax = maxRaw === undefined || maxRaw === '' ? NaN : Number(maxRaw);

  const scored = [];
  for (const item of index.items) {
    if (kategori && !`${item.kategori} ${item.subkategori}`.toLowerCase().includes(kategori)) continue;
    if (merek && !item.merek.toLowerCase().includes(merek)) continue;
    if (Number.isFinite(hargaMin) && !(item.harga >= hargaMin)) continue;
    if (Number.isFinite(hargaMax) && !(item.harga <= hargaMax)) continue;

    const nama = item.nama.toLowerCase();
    const haystack = `${nama} ${item.kategori} ${item.subkategori} ${item.merek} ${item.varian} ${item.deskripsi}`.toLowerCase();
    let score = 0;
    if (phrase.length >= 3 && nama.includes(phrase)) score += 12;
    for (const token of tokens) {
      if (nama.includes(token)) score += 5;
      else if (haystack.includes(token)) score += 1;
    }
    if (score === 0) continue;

    if (!isProductActive(item)) score -= 4;
    if (!(item.stok > 0)) score -= 2;
    scored.push({ item, score });
  }

  scored.sort(
    (a, b) =>
      b.score - a.score ||
      (b.item.stok || 0) - (a.item.stok || 0) ||
      (a.item.harga || 0) - (b.item.harga || 0)
  );

  return scored.slice(0, limit).map((entry) => entry.item);
}

function formatProductLine(item) {
  const harga = Number.isFinite(item.harga) ? `Rp${item.harga.toLocaleString('id-ID')}` : 'harga belum tersedia';
  const parts = [
    item.nama,
    item.sku ? `SKU ${item.sku}` : '',
    [item.kategori, item.subkategori].filter(Boolean).join('/'),
    item.merek ? `merek ${item.merek}` : '',
    harga,
    `stok ${Number.isFinite(item.stok) ? item.stok : 0}`,
    item.varian ? `varian ${item.varian}` : '',
    item.berat_gram ? `${item.berat_gram}g` : '',
    item.status && !/^aktif$/i.test(item.status) ? `status ${item.status}` : '',
    item.gambar ? 'foto tersedia' : 'foto belum tersedia',
  ].filter(Boolean);
  return `- ${parts.join(' | ')}`;
}

function buildProductContext(index, message, limit = 12) {
  const items = searchCatalogProducts(index, message, { limit });
  if (items.length === 0) return '';
  return items.map(formatProductLine).join('\n');
}

const CATALOG_TOOLS = [
  {
    type: 'function',
    name: 'cek_stok',
    description:
      'Cek ketersediaan dan sisa stok produk berdasarkan SKU atau nama produk (opsional dengan varian/ukuran dan jumlah yang diminta). Gunakan tool ini sebelum menyatakan stok/ketersediaan.',
    parameters: {
      type: 'object',
      properties: {
        sku: { type: 'string', description: 'SKU produk, misalnya SKU-ELE-0001' },
        nama: { type: 'string', description: 'Nama produk jika SKU tidak diketahui' },
        varian: { type: 'string', description: 'Varian/ukuran, misalnya 1m atau 128GB' },
        qty: { type: 'number', description: 'Jumlah yang ingin dipesan pelanggan' },
      },
      required: [],
    },
  },
  {
    type: 'function',
    name: 'cari_alternatif',
    description:
      'Cari produk alternatif/pengganti yang mirip berdasarkan kategori, subkategori, rentang harga, ukuran, atau kata kunci. Gunakan jika produk yang diminta tidak ada atau stok habis.',
    parameters: {
      type: 'object',
      properties: {
        kategori: { type: 'string', description: 'Kategori produk' },
        subkategori: { type: 'string', description: 'Subkategori produk' },
        harga_min: { type: 'number', description: 'Harga minimum dalam Rupiah' },
        harga_max: { type: 'number', description: 'Harga maksimum dalam Rupiah' },
        ukuran: { type: 'string', description: 'Ukuran/varian yang diinginkan, misalnya 1m atau XL' },
        kata_kunci: { type: 'string', description: 'Kata kunci dari nama/deskripsi produk' },
        limit: { type: 'number', description: 'Jumlah maksimal hasil (1-10, default 5)' },
      },
      required: [],
    },
  },
  {
    type: 'function',
    name: 'cari_produk',
    description:
      'Cari produk di katalog berdasarkan kata kunci nama/deskripsi, kategori, merek, atau rentang harga. Gunakan jika produk yang ditanyakan tidak ada di KONTEKS PRODUK RELEVAN atau pelanggan meminta daftar/katalog.',
    parameters: {
      type: 'object',
      properties: {
        kata_kunci: { type: 'string', description: 'Kata kunci produk, misalnya "cat", "earphone bluetooth", "kabel 1m"' },
        kategori: { type: 'string', description: 'Filter kategori (opsional)' },
        merek: { type: 'string', description: 'Filter merek (opsional)' },
        harga_min: { type: 'number', description: 'Harga minimum dalam Rupiah (opsional)' },
        harga_max: { type: 'number', description: 'Harga maksimum dalam Rupiah (opsional)' },
        limit: { type: 'number', description: 'Jumlah maksimal hasil (1-25, default 10)' },
      },
      required: ['kata_kunci'],
    },
  },
  {
    type: 'function',
    name: 'simpan_draft_pesanan',
    description:
      'Simpan form pesanan pelanggan setelah semua data wajib lengkap (nama, produk, jumlah, alamat/kota). Setelah tool ini dipanggil, sistem akan menampilkan form, menanyakan catatan, lalu mengirim QRIS. JANGAN menyatakan pesanan sudah diproses atau sudah dibayar.',
    parameters: {
      type: 'object',
      properties: {
        nama: { type: 'string', description: 'Nama pelanggan' },
        produk: { type: 'string', description: 'Nama produk yang dipesan' },
        qty: { type: 'string', description: 'Jumlah pesanan' },
        alamat: { type: 'string', description: 'Alamat pengiriman' },
        kota: { type: 'string', description: 'Kota/kabupaten pengiriman' },
        catatan: { type: 'string', description: 'Catatan tambahan, "-" jika tidak ada' },
        metode: { type: 'string', description: 'Metode pembayaran; toko hanya menerima QRIS (diisi otomatis oleh sistem)' },
      },
      required: ['nama', 'produk', 'qty', 'alamat', 'kota'],
    },
  },
];

function toolCekStok(index, args) {
  const qty = Math.max(1, Math.floor(Number(args.qty) || 1));
  const sku = String(args.sku || '').trim().toLowerCase();
  const nama = String(args.nama || '').trim().toLowerCase();
  const varian = String(args.varian || '').trim().toLowerCase();

  let item = sku ? index.bySku.get(sku) : null;
  if (!item && nama) {
    item = index.items.find((i) => {
      const matchName = i.nama.toLowerCase().includes(nama);
      const matchVarian = !varian || i.varian.toLowerCase().includes(varian) || i.nama.toLowerCase().includes(varian);
      return matchName && matchVarian;
    });
  }

  if (!item) {
    return JSON.stringify({
      ditemukan: false,
      pesan: 'Produk tidak ditemukan di katalog.',
      saran: 'Gunakan cari_alternatif dengan kategori/rentang harga/kata kunci.',
    });
  }

  const stok = Number.isFinite(item.stok) ? item.stok : 0;
  return JSON.stringify({
    ditemukan: true,
    produk: compactProduct(item),
    diminta: qty,
    stok_tersedia: stok,
    cukup: stok >= qty,
    sisa_setelah_pesan: Math.max(0, stok - qty),
    aktif: isProductActive(item),
  });
}

function toolCariAlternatif(index, args) {
  const limit = Math.min(Math.max(Number(args.limit) || 5, 1), 10);
  const kategori = String(args.kategori || '').trim().toLowerCase();
  const subkategori = String(args.subkategori || '').trim().toLowerCase();
  const ukuran = String(args.ukuran || '').trim().toLowerCase();
  const kataKunci = String(args.kata_kunci || '').trim().toLowerCase();
  const hargaMin = Number(args.harga_min);
  const hargaMax = Number(args.harga_max);
  const hasRange = Number.isFinite(hargaMin) && Number.isFinite(hargaMax);

  const base = index.items.filter((item) => {
    if (!isProductActive(item)) return false;
    if (Number.isFinite(item.stok) && item.stok <= 0) return false;
    if (kategori && !`${item.kategori} ${item.subkategori}`.toLowerCase().includes(kategori)) return false;
    if (subkategori && !item.subkategori.toLowerCase().includes(subkategori)) return false;
    if (ukuran && !`${item.varian} ${item.nama}`.toLowerCase().includes(ukuran)) return false;
    if (kataKunci) {
      const haystack = `${item.nama} ${item.deskripsi} ${item.merek}`.toLowerCase();
      const words = kataKunci.split(/\s+/).filter(Boolean);
      if (!words.every((word) => haystack.includes(word))) return false;
    }
    if (hasRange && !(item.harga >= hargaMin && item.harga <= hargaMax)) return false;
    return true;
  });

  let catatan = '';
  let results = base;
  if (hasRange && results.length === 0) {
    const pad = Math.max((hargaMax - hargaMin) * 0.25, 10000);
    const lo = hargaMin - pad;
    const hi = hargaMax + pad;
    results = index.items.filter((item) => {
      if (!isProductActive(item)) return false;
      if (Number.isFinite(item.stok) && item.stok <= 0) return false;
      if (kategori && !`${item.kategori} ${item.subkategori}`.toLowerCase().includes(kategori)) return false;
      if (ukuran && !`${item.varian} ${item.nama}`.toLowerCase().includes(ukuran)) return false;
      return item.harga >= lo && item.harga <= hi;
    });
    if (results.length > 0) catatan = 'Tidak ada yang persis dalam rentang harga; ini yang paling dekat.';
  }

  const target = hasRange ? (hargaMin + hargaMax) / 2 : Number.isFinite(hargaMax) ? hargaMax : Number.isFinite(hargaMin) ? hargaMin : null;
  results = results
    .slice()
    .sort((a, b) => {
      if (target !== null && Number.isFinite(a.harga) && Number.isFinite(b.harga)) {
        const diff = Math.abs(a.harga - target) - Math.abs(b.harga - target);
        if (diff !== 0) return diff;
      }
      return (b.stok || 0) - (a.stok || 0);
    })
    .slice(0, limit);

  return JSON.stringify({
    jumlah: results.length,
    catatan: catatan || undefined,
    hasil: results.map(compactProduct),
  });
}

function toolCariProduk(index, args) {
  const limit = Math.min(Math.max(Number(args.limit) || 10, 1), 25);
  const items = searchCatalogProducts(index, args.kata_kunci, { ...args, limit });

  return JSON.stringify({
    jumlah: items.length,
    hasil: items.map(compactProduct),
    catatan:
      items.length === 0
        ? 'Tidak ada produk yang cocok. Tawarkan bantuan admin atau minta pelanggan menyebutkan kategori/merek lain.'
        : undefined,
  });
}

function parseToolArguments(rawArgs) {
  if (typeof rawArgs === 'string') {
    try {
      return JSON.parse(rawArgs) || {};
    } catch {
      return {};
    }
  }
  return rawArgs && typeof rawArgs === 'object' ? rawArgs : {};
}

function executeCatalogTool(index, name, rawArgs) {
  const args = parseToolArguments(rawArgs);

  if (name === 'cek_stok') return toolCekStok(index, args);
  if (name === 'cari_alternatif') return toolCariAlternatif(index, args);
  if (name === 'cari_produk') return toolCariProduk(index, args);
  return JSON.stringify({ error: `Tool "${name}" tidak dikenal.` });
}

function buildSystemInstruction(storeInfo) {
  return `Anda adalah customer service bot ramah dan profesional untuk sebuah toko via WhatsApp.
ATURAN KRITIS (selalu patuhi):
- JANGAN pernah menyatakan pesanan sudah diproses/dikonfirmasi/dibayar. Setelah pelanggan memberi data pesanan lengkap, panggil tool "simpan_draft_pesanan"; sistem yang akan menampilkan form, menanyakan catatan, dan mengirim QRIS.
- JANGAN mengarang nama produk, harga, atau stok. Data produk hanya boleh diambil dari blok KONTEKS PRODUK RELEVAN atau hasil tool.
Informasi toko (bukan katalog produk):
---
${storeInfo || '-'}
---
Cara kerja data produk:
- Sistem sudah menyisipkan blok "[KONTEKS PRODUK RELEVAN]" di awal pesan pelanggan berisi produk yang paling cocok dengan pertanyaannya.
- Harga dan stok pada blok tersebut adalah data terbaru dan boleh langsung dipakai untuk menjawab.
- Jika produk yang ditanyakan SUDAH ADA di blok tersebut, JANGAN panggil cari_produk; langsung jawab dari data itu.
- Jika produk yang ditanyakan TIDAK ADA di blok tersebut, panggil tool "cari_produk" dengan kata kunci yang relevan sebelum menjawab. Jangan menebak.
- Panggil "cek_stok" hanya jika perlu memastikan ketersediaan untuk jumlah besar atau data stok di blok meragukan.
- Jika produk habis/tidak ada, panggil "cari_alternatif" lalu tawarkan alternatif terdekat.
Aturan:
1. Jawab secara jelas dan akurat (harga, stok, deskripsi, ketentuan) berdasarkan data yang tersedia.
2. Jika informasi tidak ada di data, katakan dengan sopan bahwa informasi belum tersedia dan tawarkan bantuan admin manusia.
3. Gunakan bahasa Indonesia yang ramah, sopan, solutif, dan format WhatsApp (gunakan *bold* untuk poin penting jika relevan).
4. JIKA pelanggan meminta foto/gambar suatu produk, periksa penanda "foto tersedia" atau "foto belum tersedia" pada data produk:
   - Jika "foto tersedia": sisipkan tag [GAMBAR: Nama_Produk] di akhir balasan (nama persis seperti data) DAN tetap tulis teks singkat untuk pelanggan.
   - Jika "foto belum tersedia": JANGAN menulis tag [GAMBAR: ...]. Jelaskan dengan sopan bahwa foto produk tersebut belum tersedia, lalu tawarkan info harga/stok atau bantuan admin.
5. JANGAN gunakan tag [GAMBAR: Nama_Produk] jika pelanggan hanya meminta list, daftar harga, atau katalog. JANGAN pernah mengirim balasan yang hanya berisi tag [GAMBAR: ...]; selalu ada teks untuk pelanggan.
6. JANGAN PERNAH menyertakan ID gambar, URL, atau isi kolom "Gambar" di dalam teks balasan Anda. Biarkan sistem yang mengirimkan gambarnya secara terpisah.
7. JIKA pelanggan sudah memberikan data pesanan (minimal nama, produk, jumlah, serta alamat/kota), WAJIB panggil tool "simpan_draft_pesanan" dengan data tersebut. Setelah tool dipanggil, sistem akan menampilkan form pesanan dan menanyakan catatan, lalu mengirim QRIS. JANGAN meminta konfirmasi YA, JANGAN menyatakan pesanan sudah diproses/dibayar. Jika pelanggan mengubah data, panggil tool lagi dengan data terbaru. JANGAN memanggil tool ini jika data pesanan belum lengkap.
8. Saat menampilkan daftar produk, tampilkan maksimal 10 produk paling relevan, jangan menampilkan seluruh katalog.`;
}

function buildStoreInfo(sheetData) {
  const sections = Array.isArray(sheetData.sections) ? sheetData.sections : [];
  const info = sections
    .filter((section) => !/katalog/i.test(section.label || ''))
    .map((section) => section.text)
    .join('\n\n');
  return info || 'Tidak ada informasi tambahan.';
}

async function createBot(config) {
  let sheetData = await loadSheetData(config);
  const sectionSummary = (sheetData.sections || []).map((s) => `${s.label} (${s.data.length} baris)`).join(', ');
  console.log(`[Data Sheet Terbaca] ${sheetData.name}: ${sectionSummary}`);
  let catalogIndex = createCatalogIndex(sheetData.raw);
  const ai = new GoogleGenAI({ apiKey: config.geminiApiKey });

  let systemInstruction = buildSystemInstruction(buildStoreInfo(sheetData));

  const modelCandidates = [
    'gemini-3.5-flash-lite',
    'gemini-3.1-flash-lite',
    'gemini-3.6-flash',
    'gemini-3.7-flash',
    'gemini-3.8-flash',
  ];

  const sessions = new Map();
  let activeModel = config.model || modelCandidates[0];
  let exhausted = new Set();

  const drive = google.drive({ version: 'v3', auth: sheetData.auth });
  const RELOAD_CHECK_MS = 1000;
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  let lastMetaCheck = 0;
  let reloadInFlight = null;

  async function refreshCatalogIfChanged() {
    if (Date.now() - lastMetaCheck < RELOAD_CHECK_MS) return;
    lastMetaCheck = Date.now();
    if (reloadInFlight) return reloadInFlight;

    reloadInFlight = (async () => {
      try {
        const meta = await drive.files.get({
          fileId: config.spreadsheetId,
          fields: 'modifiedTime,md5Checksum,version,size',
          supportsAllDrives: true,
        });
        const signature = [meta.data.modifiedTime, meta.data.md5Checksum, meta.data.version, meta.data.size]
          .filter(Boolean)
          .join('|');
        if (!signature || signature === sheetData.signature) return;

        let fresh = await loadSheetData(config);

        for (let i = 0; i < 3 && sheetData.mimeType !== SHEET_MIME && fresh.contentHash === sheetData.contentHash; i++) {
          console.log('[Catalog] Metadata berubah tapi isi masih sama (kemungkinan cache Drive), mencoba ulang...');
          await sleep(800);
          fresh = await loadSheetData(config);
        }

        if (fresh.contentHash === sheetData.contentHash) return;

        sheetData = fresh;
        catalogIndex = createCatalogIndex(fresh.raw);
        systemInstruction = buildSystemInstruction(buildStoreInfo(fresh));
        sessions.clear();
        console.log(`[Catalog] Katalog diperbarui (${signature}). Sesi pelanggan direset agar memakai data terbaru.`);
      } catch (err) {
        console.warn('[Catalog] Gagal memuat ulang katalog:', err.message);
      } finally {
        reloadInFlight = null;
      }
    })();

    return reloadInFlight;
  }

  return {
    driveAuth: sheetData.auth,

    getModel() {
      return activeModel;
    },

    async reply(senderJid, message, attempt = 0) {
      await refreshCatalogIfChanged();

      const lastId = sessions.get(senderJid);
      const productContext = buildProductContext(catalogIndex, message);
      const input = productContext
        ? `[KONTEKS PRODUK RELEVAN]\n${productContext}\n\n[PESAN PELANGGAN]\n${message}`
        : message;
      if (productContext) {
        console.log(`[Retrieval] ${productContext.split('\n').length} produk relevan disisipkan ke prompt.`);
      }

      const request = {
        model: activeModel,
        input,
        system_instruction: systemInstruction,
        tools: CATALOG_TOOLS,
      };
      if (lastId) request.previous_interaction_id = lastId;

      try {
        let interaction = await ai.interactions.create(request);

        let toolRounds = 0;
        let draftFromTool = null;
        while (toolRounds < 5) {
          const functionCalls = (interaction.steps || []).filter((step) => step.type === 'function_call');
          if (functionCalls.length === 0) break;

          const functionResults = functionCalls.map((call) => {
            const args = parseToolArguments(call.arguments);
            console.log(`[Tool] ${call.name}(${JSON.stringify(args)})`);

            let result;
            if (call.name === 'simpan_draft_pesanan') {
              draftFromTool = args;
              result = JSON.stringify({
                tersimpan: true,
                instruksi:
                  'Draft pesanan tersimpan. Sistem akan menampilkan form pesanan dan menanyakan catatan pelanggan, lalu mengirim QRIS. JANGAN meminta konfirmasi YA dan JANGAN menyatakan pesanan sudah diproses/dibayar.',
              });
            } else {
              result = executeCatalogTool(catalogIndex, call.name, args);
            }

            return {
              type: 'function_result',
              name: call.name,
              call_id: call.id,
              result,
            };
          });

          try {
            interaction = await ai.interactions.create({
              model: activeModel,
              previous_interaction_id: interaction.id,
              input: functionResults,
            });
          } catch (toolErr) {
            console.warn(`[Tool] Gagal melanjutkan setelah pemanggilan tool: ${toolErr.message}`);
            break;
          }
          toolRounds += 1;
        }

        sessions.set(senderJid, interaction.id);
        
        const responseText = interaction.output_text || '';
        console.log(`[AI Response]: ${responseText.substring(0, 100)}...`);
        const images = [];
        const imageProducts = [];

        // Deteksi produk di jawaban AI untuk kirim gambar
        const lowerResponse = responseText.toLowerCase();
        const userAskedImage = /\b(gambar|foto|lihat|spill|pic|picture)\b/i.test(message);
        const requestedImageNames = [
          ...new Set(
            [...responseText.matchAll(/\[GAMBAR:\s*([^\]]+)\]/gi)]
              .map((m) => m[1].trim())
              .filter(Boolean)
          ),
        ];
        const requestedLabels = new Set(requestedImageNames.map(normalizeProductLabel));

        for (const item of catalogIndex.items) {
          const valNama = item.nama;
          const valGambar = item.gambar;
          if (!valGambar || valGambar === '-' || !valNama) continue;

          const namaProduk = valNama.trim().toLowerCase();
          const label = normalizeProductLabel(valNama);
          const skuLabel = normalizeProductLabel(item.sku);
          const hasGambarTag = requestedLabels.has(label) || (skuLabel && requestedLabels.has(skuLabel));
          const match = hasGambarTag || (userAskedImage && lowerResponse.includes(namaProduk));
          if (!match) continue;

          console.log('[Image] Mengirim gambar untuk:', valNama);
          imageProducts.push(valNama);
          const list = valGambar.split(',').map((v) => v.trim()).filter(Boolean);
          for (const val of list) {
            let id = val;
            if (val.includes('/d/')) {
              id = val.match(/\/d\/([a-zA-Z0-9_-]+)/)?.[1] || val;
            }
            const ext = id.includes('.') ? '' : '.jpg';
            const localPath = path.resolve(__dirname, 'images', id + ext);

            if (fs.existsSync(localPath)) {
              images.push({ path: localPath });
            } else {
              const isFullUrl = val.startsWith('http://') || val.startsWith('https://');
              const isDrive = val.includes('drive.google.com') || val.includes('/d/') || !isFullUrl;
              if (isDrive) {
                images.push({ driveId: id, saveTo: localPath });
              } else {
                images.push({ url: val });
              }
            }
          }
        }
        const uniqueImages = [];
        const seen = new Set();
        for (const img of images) {
          const key = img.path || img.url || `drive:${img.driveId}`;
          if (!seen.has(key)) {
            seen.add(key);
            uniqueImages.push(img);
          }
        }
        const imageNames = [...new Set(imageProducts)];
        const missingImageNames = requestedImageNames.filter(
          (name) => !imageNames.some((product) => normalizeProductLabel(product) === normalizeProductLabel(name))
        );

        let orderDraft = draftFromTool;
        if (!orderDraft) {
          const orderMatch = responseText.match(/\[PESANAN_DRAFT:\s*(\{[\s\S]*?\})\s*\]/i);
          if (orderMatch) {
            try {
              orderDraft = JSON.parse(orderMatch[1]);
            } catch (parseErr) {
              console.warn('[Order] Gagal membaca tag draft pesanan:', parseErr.message);
            }
          }
        }

        let cleanText = responseText
          .replace(/\[PESANAN_DRAFT:[\s\S]*?\]/gi, '')
          .replace(/\[GAMBAR:[^\]]+\]/gi, '')
          .trim();

        if (!orderDraft && looksLikeOrderIntent(message)) {
          const extracted = await extractOrderDraft(ai, message, cleanText);
          if (extracted) {
            orderDraft = extracted;
            console.log('[Order] Draft dibuat via ekstraksi cadangan (tool tidak dipanggil model).');
          }
        }

        if (orderDraft) {
          orderDraft.metode = withPaymentDefaults(config).method;
          cleanText = formatOrderDraftForm(orderDraft);
          console.log('[Order] Form pesanan ditampilkan, menunggu catatan pelanggan.');
        } else if (!cleanText) {
          if (imageNames.length > 0) {
            cleanText = `Ini kak foto ${imageNames.join(', ')}.`;
          } else if (missingImageNames.length > 0) {
            cleanText = `Maaf kak, foto ${missingImageNames.join(', ')} belum tersedia saat ini. Saya bisa bantu info harga/stok, atau admin bisa mengirimkan fotonya secara manual.`;
          } else {
            cleanText =
              'Maaf kak, saya belum bisa memproses permintaan itu. Boleh kirim ulang pertanyaannya? 🙏';
          }
        }

        if (!orderDraft && userAskedImage && uniqueImages.length === 0 && requestedImageNames.length > 0) {
          const names = (missingImageNames.length ? missingImageNames : requestedImageNames).join(', ');
          const note = `foto ${names} belum tersedia saat ini. Saya bisa bantu info harga/stok, atau admin bisa mengirimkan fotonya secara manual.`;
          if (!/belum tersedia|tidak tersedia|belum ada|tidak ada foto/i.test(cleanText)) {
            cleanText = /foto|gambar/i.test(cleanText)
              ? `Maaf kak, ${note}`
              : `${cleanText}\n\nCatatan: ${note}`;
          }
        }

        return { text: cleanText, images: uniqueImages, orderDraft };
      } catch (err) {
        const message_ = err?.message || '';
        const status = err?.status || err?.code;
        const isModelError = status === 404 || /no longer available|not found|invalid model/i.test(message_);
        const isTransient =
          (typeof status === 'number' && status >= 500) ||
          /high demand|unavailable|overloaded|rate limit|quota|resource_exhausted/i.test(message_);
        const isRateLimit = status === 429 || /rate limit|quota|resource_exhausted/i.test(message_);

        if (!isModelError && !isTransient) throw err;
        if (attempt >= modelCandidates.length) throw err;

        exhausted.add(activeModel);
        const nextModel = modelCandidates.find((m) => !exhausted.has(m));
        if (!nextModel) throw err;

        console.warn(
          `[!] ${isRateLimit ? 'Limit' : 'Model tidak tersedia'}: ${activeModel} → beralih ke ${nextModel}`
        );
        activeModel = nextModel;
        sessions.delete(senderJid);
        return this.reply(senderJid, message, attempt + 1);
      }
    },
  };
}

function logAdminAlert(config, sock, jid, customerText, reason) {
  const line = `[${new Date().toISOString()}] JID: ${jid} | Pesan: ${customerText} | Penyebab: ${reason}`;
  fs.appendFileSync(path.join(__dirname, 'admin_queue.log'), line + '\n');

  if (!config.adminNumber) return;

  const adminJid = adminJidOf(config);
  sock
    .sendMessage(adminJid, {
      text: `⚠️ *Permintaan Admin*\n\nPelanggan: ${jid}\nPesan: ${customerText}\n\nAlasan: AI sedang limit.`,
    })
    .then(() => console.log(`[Notifikasi terkirim ke admin ${adminJid}]`))
    .catch((err) => console.error(`[Gagal notifikasi admin]: ${err.message}`));
}

function loadAutoState() {
  try {
    const parsed = JSON.parse(fs.readFileSync(AUTO_STATE_PATH, 'utf-8'));
    return {
      greeted: parsed.greeted && typeof parsed.greeted === 'object' ? parsed.greeted : {},
      offHoursNotified: parsed.offHoursNotified && typeof parsed.offHoursNotified === 'object' ? parsed.offHoursNotified : {},
      paused: parsed.paused && typeof parsed.paused === 'object' ? parsed.paused : {},
      pendingOrders: parsed.pendingOrders && typeof parsed.pendingOrders === 'object' ? parsed.pendingOrders : {},
    };
  } catch {
    return { greeted: {}, offHoursNotified: {}, paused: {}, pendingOrders: {} };
  }
}

function saveAutoState(state) {
  try {
    fs.writeFileSync(AUTO_STATE_PATH, JSON.stringify(state, null, 2));
  } catch (err) {
    console.warn('[AutoMessage] Gagal menyimpan state:', err.message);
  }
}

function getZonedParts(timezone) {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: timezone || 'Asia/Jakarta',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hourCycle: 'h23',
  }).formatToParts(new Date());
  const map = Object.fromEntries(parts.map((p) => [p.type, p.value]));
  return {
    dateKey: `${map.year}-${map.month}-${map.day}`,
    minutes: Number(map.hour) * 60 + Number(map.minute),
  };
}

function isWithinBusinessHours(hours) {
  const toMinutes = (value) => {
    const [h, m] = String(value).split(':').map(Number);
    return (Number.isFinite(h) ? h : 0) * 60 + (Number.isFinite(m) ? m : 0);
  };
  const start = toMinutes(hours.start);
  const end = toMinutes(hours.end);
  if (start === end) return true;

  const { minutes } = getZonedParts(hours.timezone);
  return start < end ? minutes >= start && minutes < end : minutes >= start || minutes < end;
}

function isMenuRequest(text, keywords) {
  const normalized = String(text || '')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!normalized) return false;

  const firstWord = normalized.split(' ')[0];
  return keywords.some((keyword) => {
    const k = String(keyword || '').toLowerCase().trim();
    return k && (normalized === k || firstWord === k);
  });
}

async function handleAutoMessages(config, sock, jid, text, state) {
  const auto = config.autoMessage || withAutoMessageDefaults(config);
  const isGroup = jid.endsWith('@g.us');

  if (!isGroup && auto.greeting && !state.greeted[jid]) {
    await sock.sendMessage(jid, { text: auto.greeting });
    state.greeted[jid] = new Date().toISOString();
    saveAutoState(state);
    console.log(`[AutoMessage] Sapaan pertama dikirim ke ${jid}`);
  }

  if (auto.menuText && isMenuRequest(text, auto.menuKeywords)) {
    await sock.sendMessage(jid, { text: auto.menuText });
    console.log(`[AutoMessage] Menu dikirim ke ${jid}`);
    return { handled: true };
  }

  const hours = auto.businessHours || {};
  if (!isGroup && auto.offHoursText && hours.start && hours.end && !isWithinBusinessHours(hours)) {
    const { dateKey } = getZonedParts(hours.timezone);
    const alreadyNotified = state.offHoursNotified[jid] === dateKey;
    if (!auto.offHoursNotifyOncePerDay || !alreadyNotified) {
      await sock.sendMessage(jid, { text: auto.offHoursText });
      if (auto.offHoursNotifyOncePerDay) {
        state.offHoursNotified[jid] = dateKey;
        saveAutoState(state);
      }
      console.log(`[AutoMessage] Info di luar jam operasional dikirim ke ${jid}`);
    }
  }

  return { handled: false };
}

function adminJidOf(config) {
  const digits = String(config.adminNumber || '').replace(/\D/g, '');
  if (!digits) return '';
  if (digits.startsWith('0')) return `62${digits.slice(1)}@s.whatsapp.net`;
  if (digits.startsWith('8')) return `62${digits}@s.whatsapp.net`;
  return `${digits}@s.whatsapp.net`;
}

function formatClock(isoString, timezone) {
  return new Intl.DateTimeFormat('en-GB', {
    timeZone: timezone || 'Asia/Jakarta',
    hour: '2-digit',
    minute: '2-digit',
    hourCycle: 'h23',
  }).format(new Date(isoString));
}

function getPauseRemaining(state, jid) {
  const untilIso = state.paused?.[jid];
  if (!untilIso) return 0;
  const remaining = new Date(untilIso).getTime() - Date.now();
  return Number.isFinite(remaining) && remaining > 0 ? remaining : 0;
}

async function notifyAdminHandoff(config, sock, jid, text, untilIso) {
  const adminJid = adminJidOf(config);
  if (!adminJid) return;

  const timezone = withAutoMessageDefaults(config).businessHours.timezone;
  const body = [
    '🙋 *Chat dialihkan ke admin*',
    '',
    `Pelanggan: ${jid.split('@')[0]}`,
    `Pesan: ${text}`,
    `Bot pause sampai: ${formatClock(untilIso, timezone)} (${withHandoffDefaults(config).pauseMinutes} menit)`,
    '',
    'Silakan balas langsung ke pelanggan. Bot tidak akan menjawab selama pause.',
  ].join('\n');

  try {
    await sock.sendMessage(adminJid, { text: body });
    console.log(`[Handoff] Notifikasi dialihkan ke admin ${adminJid}`);
  } catch (err) {
    console.error(`[Handoff] Gagal notifikasi admin: ${err.message}`);
  }
}

async function forwardToAdminPaused(config, sock, jid, text) {
  const adminJid = adminJidOf(config);
  if (!adminJid) return;

  try {
    await sock.sendMessage(adminJid, {
      text: `💬 *Pesan lanjutan (bot pause)*\nPelanggan: ${jid.split('@')[0]}\nPesan: ${text}`,
    });
  } catch (err) {
    console.error(`[Handoff] Gagal meneruskan pesan ke admin: ${err.message}`);
  }
}

function formatOrderSummary(order) {
  return [
    '✅ *Pesanan dikonfirmasi*',
    '',
    `Nama: ${order.nama || '-'}`,
    `Produk: ${order.produk || '-'}`,
    `Jumlah: ${order.qty || '-'}`,
    `Alamat: ${order.alamat || '-'}`,
    `Kota: ${order.kota || '-'}`,
    `Metode: ${order.metode || '-'}`,
    `Catatan: ${order.catatan || '-'}`,
    '',
    'Terima kasih kak 🙏 Admin kami akan segera memproses pesanan Anda.',
  ].join('\n');
}

function formatDraftSummary(order) {
  return [
    '📝 *Konfirmasi Pesanan*',
    '',
    `Nama: ${order.nama || '-'}`,
    `Produk: ${order.produk || '-'}`,
    `Jumlah: ${order.qty || '-'}`,
    `Alamat: ${order.alamat || '-'}`,
    `Kota: ${order.kota || '-'}`,
    `Metode: ${order.metode || '-'}`,
    `Catatan: ${order.catatan || '-'}`,
    '',
    'Ketik *YA* untuk konfirmasi, atau beri tahu jika ada yang ingin diubah.',
  ].join('\n');
}

function formatOrderDraftForm(order) {
  return [
    '📝 *Form Pesanan*',
    '',
    `Nama: ${order.nama || '-'}`,
    `Produk: ${order.produk || '-'}`,
    `Jumlah: ${order.qty || '-'}`,
    `Alamat: ${order.alamat || '-'}`,
    `Kota: ${order.kota || '-'}`,
    `Metode: ${order.metode || '-'}`,
    '',
    'Ada catatan tambahan untuk pesanan? (contoh: warna, ukuran, jam kirim)',
    'Ketik *-* jika tidak ada.',
  ].join('\n');
}

function formatPaymentRequest(order, payment) {
  return [
    '✅ *Form Pesanan Diterima*',
    '',
    `Nama: ${order.nama || '-'}`,
    `Produk: ${order.produk || '-'}`,
    `Jumlah: ${order.qty || '-'}`,
    `Alamat: ${order.alamat || '-'}`,
    `Kota: ${order.kota || '-'}`,
    `Metode: ${payment.method}`,
    `Catatan: ${order.catatan || '-'}`,
    '',
    'Silakan lakukan pembayaran melalui QRIS berikut.',
    'Setelah pembayaran berhasil, ketik *SUDAH* ya kak.',
  ].join('\n');
}

function normalizeNotes(text) {
  const raw = String(text || '').trim();
  if (!raw) return '-';
  if (/^(-|tidak|tdk|nggak|gak|ga|tidak ada|nggak ada|gak ada|none|no|tidak usah|gak usah)$/i.test(raw)) {
    return '-';
  }
  return raw.slice(0, 300);
}

function isPaymentConfirmation(text) {
  const raw = String(text || '').trim().toLowerCase();
  if (!raw || raw.length > 80 || /\?/.test(raw)) return false;
  if (/\b(belum|gagal|error|batal|tidak|nggak|gak)\b/.test(raw)) return false;
  return (
    /\b(sudah|udah|sdh)\s+(bayar|transfer|scan|qris|kirim|selesai|berhasil)\b/i.test(raw) ||
    /^(sudah|udah|sdh|done|paid|iya sudah|ya sudah|saya sudah|aku sudah)\b/i.test(raw)
  );
}

function extractDriveFileId(url) {
  const source = String(url || '');
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]{10,})/,
    /[?&]id=([a-zA-Z0-9_-]{10,})/,
    /\/d\/([a-zA-Z0-9_-]{10,})/,
  ];
  for (const pattern of patterns) {
    const match = source.match(pattern);
    if (match) return match[1];
  }
  return '';
}

const QRIS_MAX_BYTES = 5 * 1024 * 1024;

function normalizeSha256(value) {
  return String(value || '').toLowerCase().replace(/[^a-f0-9]/g, '');
}

function validateImageBuffer(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 100) throw new Error('File gambar kosong/terlalu kecil');
  if (buffer.length > QRIS_MAX_BYTES) throw new Error('File gambar terlalu besar (maks 5MB)');
  const isPng = buffer.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));
  const isJpg = buffer[0] === 0xff && buffer[1] === 0xd8 && buffer[2] === 0xff;
  const isWebp =
    buffer.subarray(0, 4).toString('ascii') === 'RIFF' && buffer.subarray(8, 12).toString('ascii') === 'WEBP';
  if (!isPng && !isJpg && !isWebp) throw new Error('File bukan gambar PNG/JPEG/WEBP yang valid');
}

function assertSafeRemoteUrl(url) {
  let parsed;
  try {
    parsed = new URL(url);
  } catch {
    throw new Error('URL QRIS tidak valid');
  }
  if (parsed.protocol !== 'https:') throw new Error('URL QRIS harus menggunakan https');
  const host = parsed.hostname.toLowerCase();
  const isPrivate =
    host === 'localhost' ||
    host.endsWith('.local') ||
    host === '[::1]' ||
    host === '::1' ||
    /^(127\.|10\.|0\.|169\.254\.|192\.168\.)/.test(host) ||
    /^172\.(1[6-9]|2\d|3[01])\./.test(host);
  if (isPrivate) throw new Error('Host URL QRIS tidak diizinkan');
}

async function fetchImageBuffer(url) {
  assertSafeRemoteUrl(url);
  const response = await fetch(url, { redirect: 'follow', signal: AbortSignal.timeout(15000) });
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const contentType = response.headers.get('content-type') || '';
  if (!/^image\//i.test(contentType)) {
    throw new Error(`bukan gambar langsung (content-type: ${contentType || 'unknown'})`);
  }
  const declaredLength = Number(response.headers.get('content-length'));
  if (Number.isFinite(declaredLength) && declaredLength > QRIS_MAX_BYTES) {
    throw new Error('File gambar terlalu besar (maks 5MB)');
  }
  const buffer = Buffer.from(await response.arrayBuffer());
  validateImageBuffer(buffer);
  return buffer;
}

async function loadQrisImage(source, driveAuth) {
  if (/^https?:\/\//i.test(source)) {
    const driveId = extractDriveFileId(source);
    if (driveId) {
      if (driveAuth) {
        const drive = google.drive({ version: 'v3', auth: driveAuth });
        const res = await drive.files.get(
          { fileId: driveId, alt: 'media', supportsAllDrives: true },
          { responseType: 'arraybuffer' }
        );
        const buffer = Buffer.from(res.data);
        validateImageBuffer(buffer);
        return buffer;
      }
      return fetchImageBuffer(`https://drive.google.com/uc?export=download&id=${driveId}`);
    }
    return fetchImageBuffer(source);
  }

  const filePath = path.isAbsolute(source) ? source : path.resolve(__dirname, source);
  if (!fs.existsSync(filePath)) throw new Error(`file tidak ditemukan: ${filePath}`);
  const buffer = fs.readFileSync(filePath);
  validateImageBuffer(buffer);
  return buffer;
}

function alertAdminSecurity(config, sock, message) {
  const adminJid = adminJidOf(config);
  if (!adminJid) return;
  sock
    .sendMessage(adminJid, { text: `⚠️ *Peringatan Keamanan QRIS*\n\n${message}` })
    .catch((err) => console.error('[Payment] Gagal kirim peringatan ke admin:', err.message));
}

async function sendQrisImage(config, sock, jid, driveAuth) {
  const payment = withPaymentDefaults(config);
  const source = payment.qrisImageUrl;
  if (!source) {
    console.warn('[Payment] QRIS_IMAGE_URL belum diatur di .env / config.payment.qrisImageUrl');
    await sock.sendMessage(jid, {
      text: 'Maaf kak, gambar QRIS belum tersedia. Admin akan mengirimkan QRIS secara manual ya.',
    });
    return false;
  }

  try {
    const buffer = await loadQrisImage(source, driveAuth);
    const expectedHash = normalizeSha256(payment.qrisImageSha256);
    if (expectedHash) {
      const actualHash = crypto.createHash('sha256').update(buffer).digest('hex');
      if (actualHash !== expectedHash) {
        throw new Error(
          `Hash QRIS tidak cocok (diharapkan ${expectedHash.slice(0, 12)}..., didapat ${actualHash.slice(0, 12)}...). Gambar TIDAK dikirim.`
        );
      }
    } else {
      console.warn(
        '[Payment] PERINGATAN: QRIS_IMAGE_SHA256 belum diset. Gambar QRIS tidak terverifikasi keasliannya.'
      );
    }

    await sock.sendMessage(jid, { image: buffer, caption: payment.qrisCaption });
    console.log('[Payment] Gambar QRIS terkirim.');
    return true;
  } catch (err) {
    console.error('[Payment] Gagal mengirim QRIS:', err.message);
    if (extractDriveFileId(source)) {
      console.error(
        `[Payment] File Drive harus dibagikan ke ${getServiceAccountEmail(config) || 'email service account'} (Viewer), atau gunakan direct image URL.`
      );
    }
    alertAdminSecurity(
      config,
      sock,
      `Pengiriman QRIS ke ${jid.split('@')[0]} dibatalkan.\nAlasan: ${err.message}\n\nPeriksa QRIS_IMAGE_URL / QRIS_IMAGE_SHA256.`
    );
    await sock.sendMessage(jid, {
      text: 'Maaf kak, gambar QRIS gagal dikirim. Admin akan mengirimkan QRIS secara manual ya.',
    });
    return false;
  }
}

function looksLikeOrderIntent(text) {
  return /(pesan|order|beli|checkout|cod|transfer|alamat|kirim\s*ke|atas\s*nama|nama\s*saya)/i.test(String(text || ''));
}

const EXTRACTOR_MODELS = ['gemini-3.5-flash-lite', 'gemini-3.1-flash-lite'];

async function extractOrderDraft(ai, message, replyText) {
  const prompt = [
    'Ekstrak data pesanan dari percakapan berikut. Balas HANYA JSON valid dengan format:',
    '{"is_order": boolean, "nama": string, "produk": string, "qty": string, "alamat": string, "kota": string, "catatan": string, "metode": string}',
    'Aturan: is_order=true HANYA jika nama, produk, qty, alamat, dan kota semuanya tersedia. Jika ada yang kosong/belum pasti, is_order=false. Jangan mengarang data.',
    '',
    'PESAN PELANGGAN: ' + message,
    'BALASAN ASISTEN: ' + replyText,
  ].join('\n');

  for (const model of EXTRACTOR_MODELS) {
    try {
      const interaction = await ai.interactions.create({
        model,
        response_mime_type: 'application/json',
        response_format: {
          type: 'text',
          mime_type: 'application/json',
          schema: {
            type: 'object',
            properties: {
              is_order: { type: 'boolean' },
              nama: { type: 'string' },
              produk: { type: 'string' },
              qty: { type: 'string' },
              alamat: { type: 'string' },
              kota: { type: 'string' },
              catatan: { type: 'string' },
              metode: { type: 'string' },
            },
            required: ['is_order', 'nama', 'produk', 'qty', 'alamat', 'kota'],
          },
        },
        input: prompt,
      });

      const parsed = JSON.parse(interaction.output_text || '{}');
      if (!parsed || parsed.is_order !== true) return null;

      const required = ['nama', 'produk', 'qty', 'alamat', 'kota'];
      const complete = required.every((key) => {
        const value = String(parsed[key] ?? '').trim();
        return value && value !== '-';
      });
      if (!complete) return null;

      return {
        nama: String(parsed.nama).trim(),
        produk: String(parsed.produk).trim(),
        qty: String(parsed.qty).trim(),
        alamat: String(parsed.alamat).trim(),
        kota: String(parsed.kota).trim(),
        catatan: String(parsed.catatan ?? '-').trim() || '-',
        metode: String(parsed.metode ?? '-').trim() || '-',
      };
    } catch (err) {
      console.warn('[Order] Ekstraksi draft via ' + model + ' gagal: ' + err.message);
    }
  }

  return null;
}
function notifyAdminOrder(config, sock, jid, order, note = '') {
  if (!config.adminNumber) return;

  const adminJid = adminJidOf(config);
  const lines = [
    '🛒 *Pesanan Baru dari WhatsApp*',
    '',
    `Nama: ${order.nama || '-'}`,
    `Produk: ${order.produk || '-'}`,
    `Jumlah: ${order.qty || '-'}`,
    `Alamat: ${order.alamat || '-'}`,
    `Kota: ${order.kota || '-'}`,
    `Metode: ${order.metode || '-'}`,
    `Catatan: ${order.catatan || '-'}`,
    `No. WA: ${jid.split('@')[0]}`,
  ];
  if (note) lines.push('', note);

  sock
    .sendMessage(adminJid, { text: lines.join('\n') })
    .then(() => console.log(`[Order] Notifikasi pesanan terkirim ke admin ${adminJid}`))
    .catch((err) => console.error(`[Order] Gagal notifikasi admin: ${err.message}`));
}

async function handlePendingOrderStage(config, sock, jid, text, msg, autoState, handoffConfig, driveAuth) {
  const pendingEntry = autoState.pendingOrders[jid];
  if (!pendingEntry) return false;

  const ordersConfig = withOrdersDefaults(config);
  const payment = withPaymentDefaults(config);
  const createdAt = new Date(pendingEntry.createdAt || 0).getTime();
  const timeoutMs = ordersConfig.confirmTimeoutMinutes * 60000;
  const stage = pendingEntry.stage || 'awaiting_notes';

  if (!pendingEntry.order || !Number.isFinite(createdAt) || Date.now() - createdAt > timeoutMs) {
    delete autoState.pendingOrders[jid];
    saveAutoState(autoState);
    console.log(`[Order] Draft pesanan ${jid} kedaluwarsa, dihapus.`);
    return false;
  }

  if (isOrderCancellation(text)) {
    delete autoState.pendingOrders[jid];
    saveAutoState(autoState);
    await sock.sendMessage(
      jid,
      { text: 'Baik kak, draft pesanan dibatalkan. Silakan chat lagi jika ingin memesan. 🙏' },
      { quoted: msg }
    );
    console.log(`[Order] Draft pesanan ${jid} dibatalkan pelanggan.`);
    return true;
  }

  if (stage === 'awaiting_notes') {
    const looksLikeQuestion =
      /\?/.test(text) ||
      /^(berapa|apa|apakah|bisa|harga|stok|ongkir|kirim|kapan|dimana|di mana)\b/i.test(text.trim());
    const looksLikeChange = /\b(ubah|ganti|tukar|salah|bukan|jadikan|pindah)\b/i.test(text);
    if (!looksLikeQuestion && !looksLikeChange) {
      pendingEntry.order.catatan = normalizeNotes(text);
      pendingEntry.stage = 'awaiting_payment';
      pendingEntry.updatedAt = new Date().toISOString();
      saveAutoState(autoState);

      await sock.sendMessage(
        jid,
        { text: formatPaymentRequest(pendingEntry.order, payment) },
        { quoted: msg }
      );
      if (payment.enabled) {
        await sendQrisImage(config, sock, jid, driveAuth);
      }
      console.log(`[Payment] Catatan disimpan dan QRIS dikirim ke ${jid}.`);
      return true;
    }
    return false;
  }

  if (stage === 'awaiting_payment' && isPaymentConfirmation(text)) {
    try {
      const orderResult = await logOrder(config, jid, pendingEntry.order, 'Menunggu Verifikasi');
      delete autoState.pendingOrders[jid];
      saveAutoState(autoState);

      await sock.sendMessage(jid, { text: payment.thankYouText }, { quoted: msg });

      const untilIso = new Date(Date.now() + handoffConfig.pauseMinutes * 60000).toISOString();
      autoState.paused[jid] = untilIso;
      saveAutoState(autoState);

      if (orderResult.saved) {
        await notifyAdminOrder(
          config,
          sock,
          jid,
          pendingEntry.order,
          `Status: pembayaran diklaim sudah (menunggu verifikasi)\nBot pause ${handoffConfig.pauseMinutes} menit untuk nomor ini.`
        );
      }
      fs.appendFileSync(
        path.join(__dirname, 'admin_queue.log'),
        `[${new Date().toISOString()}] JID: ${jid} | Pembayaran diklaim sudah | Produk: ${pendingEntry.order.produk} | Bot pause ${handoffConfig.pauseMinutes} menit\n`
      );
      console.log(
        `[Payment] Pembayaran diklaim sudah oleh ${jid}; pesanan dicatat, admin diberi tahu, bot pause ${handoffConfig.pauseMinutes} menit.`
      );
    } catch (orderErr) {
      console.error(`[Payment] Gagal memproses konfirmasi pembayaran: ${orderErr.message}`);
    }
    return true;
  }

  return false;
}
async function startWhatsAppBot(config, bot) {
  const authFolder = path.resolve(__dirname, 'baileys_auth');
  const autoState = loadAutoState();
  const { state, saveCreds } = await useMultiFileAuthState(authFolder);

  const sock = makeWASocket({
    auth: state,
    logger: pino({ level: 'silent' }),
    printQRInTerminal: false,
  });

  sock.ev.on('creds.update', saveCreds);

  sock.ev.on('connection.update', (update) => {
    const { connection, lastDisconnect, qr } = update;

    if (qr) {
      console.log('\n[!] Scan QR code ini dengan aplikasi WhatsApp Anda:\n');
      qrcode.generate(qr, { small: true });
    }

    if (connection === 'close') {
      const statusCode = lastDisconnect?.error?.output?.statusCode;
      const shouldReconnect = statusCode !== DisconnectReason.loggedOut;
      console.log(`[-] Koneksi terputus. Reconnect: ${shouldReconnect}`);
      if (shouldReconnect) {
        startWhatsAppBot(config, bot);
      }
    } else if (connection === 'open') {
      console.log('[+] WhatsApp bot berhasil terhubung dan siap melayani customer!\n');
    }
  });

  sock.ev.on('messages.upsert', async ({ messages, type }) => {
    if (type !== 'notify') return;

    for (const msg of messages) {
      const jid = msg.key.remoteJid;
      if (!jid || msg.key.fromMe || jid.endsWith('@broadcast')) continue;

      const text =
        msg.message?.conversation ||
        msg.message?.extendedTextMessage?.text ||
        msg.message?.imageMessage?.caption;

      if (!text) continue;

      console.log(`[Pesan Masuk dari ${jid}]: ${text}`);

      try {
        await sock.readMessages([msg.key]);

        const handoffConfig = withHandoffDefaults(config);
        const pauseUntilIso = autoState.paused[jid];
        if (pauseUntilIso) {
          if (getPauseRemaining(autoState, jid) > 0) {
            await forwardToAdminPaused(config, sock, jid, text);
            console.log(`[Handoff] Bot masih pause untuk ${jid}, pesan diteruskan ke admin.`);
            continue;
          }
          delete autoState.paused[jid];
          saveAutoState(autoState);
          console.log(`[Handoff] Pause untuk ${jid} berakhir, bot aktif kembali.`);
        }

        if (
          await handlePendingOrderStage(config, sock, jid, text, msg, autoState, handoffConfig, bot.driveAuth)
        )
          continue;
        let autoHandled = false;
        try {
          const autoResult = await handleAutoMessages(config, sock, jid, text, autoState);
          autoHandled = autoResult.handled;
        } catch (autoErr) {
          console.warn(`[AutoMessage] Gagal mengirim pesan otomatis ke ${jid}:`, autoErr.message);
        }
        if (autoHandled) continue;

        if (handoffConfig.enabled && isHandoffRequest(text, handoffConfig.keywords)) {
          const untilIso = new Date(Date.now() + handoffConfig.pauseMinutes * 60000).toISOString();
          autoState.paused[jid] = untilIso;
          saveAutoState(autoState);

          await sock.sendMessage(jid, { text: handoffConfig.message });
          await notifyAdminHandoff(config, sock, jid, text, untilIso);
          fs.appendFileSync(
            path.join(__dirname, 'admin_queue.log'),
            `[${new Date().toISOString()}] JID: ${jid} | Pesan: ${text} | Penyebab: permintaan alih ke admin (pause ${handoffConfig.pauseMinutes} menit)\n`
          );
          console.log(`[Handoff] ${jid} dialihkan ke admin, bot pause ${handoffConfig.pauseMinutes} menit.`);
          continue;
        }

        const response = await bot.reply(jid, text);
        
        await sock.sendMessage(jid, { text: response.text }, { quoted: msg });
        
        for (const img of response.images) {
          try {
            if (img.path && fs.existsSync(img.path)) {
              await sock.sendMessage(jid, { image: fs.readFileSync(img.path) });
            } else if (img.driveId) {
              let buffer = null;
              const saveTo = img.saveTo;
              console.log(`[Image] Mendownload gambar dari Drive (API): ${img.driveId}`);
              try {
                const drive = google.drive({ version: 'v3', auth: bot.driveAuth });
                const res = await drive.files.get({ fileId: img.driveId, alt: 'media' }, { responseType: 'arraybuffer' });
                buffer = Buffer.from(res.data);
                if (buffer && buffer.length > 1000) {
                  fs.writeFileSync(saveTo, buffer);
                  console.log(`[Image] Berhasil menyimpan: ${path.basename(saveTo)} (${buffer.length} bytes)`);
                } else {
                  console.error(`[Image] Gagal: file terlalu kecil (${buffer.length} bytes)`);
                }
              } catch (e) {
                console.error(`[Image] Gagal download ${img.driveId}:`, e.message);
              }
              if (buffer && buffer.length > 1000) {
                await sock.sendMessage(jid, { image: buffer });
              }
            } else if (img.url) {
              await sock.sendMessage(jid, { image: { url: img.url } });
            }
          } catch (imgErr) {
            console.error(`[Image] Gagal mengirim gambar ke ${jid}:`, imgErr.message);
          }
        }
        
        if (response.orderDraft) {
          autoState.pendingOrders[jid] = {
            order: response.orderDraft,
            stage: 'awaiting_notes',
            createdAt: new Date().toISOString(),
          };
          saveAutoState(autoState);
          console.log(`[Order] Form pesanan disimpan untuk ${jid}, menunggu catatan pelanggan.`);
        }

        console.log(`[Balasan Terkirim ke ${jid}]: ${response.text}\n`);
      } catch (err) {
        const isRateLimit = err?.status === 429 || /rate limit|quota|resource_exhausted/i.test(err?.message || '');
        if (isRateLimit) {
          const notice =
            'Mohon maaf, layanan asisten otomatis sedang sibuk dan kuota sementara habis. ' +
            'Pesan Anda sudah kami teruskan ke admin toko, silakan tunggu balasan dari admin ya. Terima kasih 🙏';
          try {
            await sock.sendMessage(jid, { text: notice }, { quoted: msg });
            console.warn(`[!] Rate limit AI, pelanggan ${jid} diberi pesan pengalihan ke admin.`);
            logAdminAlert(config, sock, jid, text, err.message);
          } catch (sendErr) {
            console.error(`[Gagal kirim pesan pengalihan ke ${jid}]:`, sendErr.message);
          }
        } else {
          console.error(`[Error balas pesan ke ${jid}]:`, err.message);
        }
      }
    }
  });
}

async function main() {
  try {
    const config = await ensureConfig();
    const ordersConfig = withOrdersDefaults(config);
    if (!ordersConfig.spreadsheetId) {
      console.log('[i] Pesanan selalu dicatat ke orders.jsonl.');
      console.log(`[i] Agar juga masuk Google Sheet: buat spreadsheet baru bernama "${ordersConfig.sheetName}",`);
      console.log(`    share ke ${getServiceAccountEmail(config) || 'email service account'} sebagai Editor,`);
      console.log('    lalu isi orders.spreadsheetId dan set orders.enabled=true di config.json.');
    }
    console.log('[*] Mengambil data katalog dari Google Sheet...');
    const bot = await createBot(config);
    console.log(`[*] Model AI aktif: ${bot.getModel()}`);
    console.log('[*] Menyiapkan koneksi WhatsApp Baileys...');
    await startWhatsAppBot(config, bot);
  } catch (err) {
    console.error('Fatal Error:', err.message);
  }
}

if (require.main === module) {
  main();
}

module.exports = {
  createBot,
  loadSheetData,
  ensureConfig,
  startWhatsAppBot,
  withAutoMessageDefaults,
  isMenuRequest,
  isWithinBusinessHours,
  getZonedParts,
  withOrdersDefaults,
  withPaymentDefaults,
  logOrder,
  formatOrderTime,
  formatOrderDraftForm,
  formatPaymentRequest,
  normalizeNotes,
  isPaymentConfirmation,
  handlePendingOrderStage,
  extractDriveFileId,
  withHandoffDefaults,
  isHandoffRequest,
  getPauseRemaining,
  createCatalogIndex,
  searchCatalogProducts,
  buildProductContext,
  buildSystemInstruction,
  executeCatalogTool,
  isOrderConfirmation,
  isOrderCancellation,
  formatOrderSummary,
  formatDraftSummary,
  extractOrderDraft,
  looksLikeOrderIntent,
};
