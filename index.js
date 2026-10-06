const fs = require('fs');
const path = require('path');
const readline = require('readline');
const https = require('https');
const pino = require('pino');
const qrcode = require('qrcode-terminal');
const { google } = require('googleapis');
const { GoogleGenAI } = require('@google/genai');
const baileys = require('@whiskeysockets/baileys');

const makeWASocket = baileys.default || baileys.makeWASocket;
const { useMultiFileAuthState, DisconnectReason } = baileys;

const CONFIG_PATH = path.join(__dirname, 'config.json');

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

async function loadSheetData(config) {
  const scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive.readonly'];
  const auth = new google.auth.GoogleAuth({
    keyFile: path.resolve(__dirname, config.credentialsPath),
    scopes,
  });

  const sheets = google.sheets({ version: 'v4', auth });
  const ranges = (config.sheetRange || 'Sheet1!A:Z')
    .split(',')
    .map((r) => r.trim())
    .filter(Boolean);

  const results = await Promise.all(
    ranges.map(async (range) => {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: config.spreadsheetId,
        range,
      });
      const rows = response.data.values || [];
      if (rows.length === 0) return { text: `[${range}]: Kosong`, data: [] };

      const headers = rows[0].map(h => h.trim());
      const itemsData = rows.slice(1).map((row) => {
        const item = {};
        headers.forEach((h, i) => {
          item[h] = (row[i] || '-').trim();
        });
        return item;
      });

      const text = `### Sheet: ${range} ###\n` + itemsData.map(item => 
        Object.entries(item).map(([k, v]) => `${k}: ${v}`).join(' | ')
      ).join('\n');

      return { text, data: itemsData };
    })
  );

  return {
    context: results.map(r => r.text).join('\n\n'),
    raw: results.flatMap(r => r.data),
    auth,
  };
}

async function createBot(config) {
  const sheetData = await loadSheetData(config);
  console.log('[Data Sheet Terbaca]:', JSON.stringify(sheetData.raw, null, 2));
  const ai = new GoogleGenAI({ apiKey: config.geminiApiKey });

  const systemInstruction = `Anda adalah customer service bot ramah dan profesional untuk sebuah toko via WhatsApp.
Jawab pertanyaan pelanggan berdasarkan data katalog & informasi toko berikut:
---
${sheetData.context}
---
Aturan:
1. Jika informasi ada di data, jawab secara jelas dan akurat (harga, stok, deskripsi, ketentuan).
2. Jika informasi tidak ada di data, katakan dengan sopan bahwa informasi belum tersedia dan tawarkan bantuan admin manusia.
3. Gunakan bahasa Indonesia yang ramah, sopan, solutif, dan format WhatsApp (gunakan *bold* untuk poin penting jika relevan).
4. JIKA PELANGGAN MEMINTA FOTO/GAMBAR dari suatu produk, maka Anda WAJIB menyisipkan tag berikut di akhir balasan Anda: [GAMBAR: Nama_Produk] (Pastikan Nama_Produk persis sesuai di data).
5. JANGAN gunakan tag [GAMBAR: Nama_Produk] jika pelanggan hanya meminta list, daftar harga, atau katalog. Gunakan tag tersebut HANYA JIKA pelanggan spesifik minta dilihatkan gambar.
6. JANGAN PERNAH menyertakan ID gambar, URL, atau isi kolom "Gambar" di dalam teks balasan Anda. Biarkan sistem yang mengirimkan gambarnya secara terpisah.`;

  const modelCandidates = [
    'gemini-3.1-flash-lite',
    'gemini-3.5-flash-lite',
    'gemini-3.6-flash',
    'gemini-3.7-flash',
    'gemini-3.8-flash',
    'gemini-3-flash',
  ];

  const sessions = new Map();
  let activeModel = modelCandidates[0];
  let exhausted = new Set();

  return {
    driveAuth: sheetData.auth,

    getModel() {
      return activeModel;
    },

    async reply(senderJid, message, attempt = 0) {
      const lastId = sessions.get(senderJid);
      const request = {
        model: activeModel,
        input: message,
        system_instruction: systemInstruction,
      };
      if (lastId) request.previous_interaction_id = lastId;

      try {
        const interaction = await ai.interactions.create(request);
        sessions.set(senderJid, interaction.id);
        
        const responseText = interaction.output_text || '';
        console.log(`[AI Response]: ${responseText.substring(0, 100)}...`);
        const images = [];
        
        // Deteksi produk di jawaban AI untuk kirim gambar
        console.log(`[Image Check] Total produk di sheet: ${sheetData.raw.length}`);
        if (sheetData.raw.length > 0) {
          console.log(`[Image Check] Keys kolom sheet pertama:`, Object.keys(sheetData.raw[0]));
          console.log(`[Image Check] Contoh item 1:`, JSON.stringify(sheetData.raw[0]));
        }
        for (const item of sheetData.raw) {
          // Cari key untuk Nama (case insensitive)
          const namaKey = Object.keys(item).find(k => k.toLowerCase().replace(/[_\s]/g, '').includes('nama'));
          const gambarKey = Object.keys(item).find(k => k.toLowerCase().replace(/[_\s]/g, '') === 'gambar' || k.toLowerCase().replace(/[_\s]/g, '').includes('foto'));
          
          const valNama = namaKey ? item[namaKey] : null;
          const valGambar = gambarKey ? item[gambarKey] : null;

          const hasGambar = valGambar && valGambar !== '-' && valGambar !== '';
          const hasNama = valNama && valNama !== '-' && valNama !== '';

          if (!hasGambar || !hasNama) {
            continue;
          }

          const namaProduk = valNama.trim().toLowerCase();
          // Cek apakah ada tag [GAMBAR: Nama_Produk] atau kata kunci pesan user minta gambar
          const hasGambarTag = responseText.toLowerCase().includes(`[gambar: ${namaProduk}]`) || responseText.toLowerCase().includes(`[gambar:${namaProduk}]`);
          const userAskedImage = /\b(gambar|foto|lihat|spill|pic|picture)\b/i.test(message);
          const match = hasGambarTag || (userAskedImage && responseText.toLowerCase().includes(namaProduk));
          
          console.log(`[Image Check] "${valNama}" → ${match ? 'Kirim Gambar ✓' : 'Lewati (tidak minta gambar)'}`);
          if (match) {
            const list = valGambar.split(',').map(v => v.trim()).filter(Boolean);
            list.forEach(val => {
              let id = val;
              if (val.includes('/d/')) {
                id = val.match(/\/d\/([a-zA-Z0-9_-]+)/)?.[1] || val;
              }
              const ext = id.includes('.') ? '' : '.jpg'; // Fallback ext
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
            });
          }
        } // akhir for loop

        const uniqueImages = [];
        const seen = new Set();
        for (const img of images) {
          const key = img.path || img.url;
          if (!seen.has(key)) {
            seen.add(key);
            uniqueImages.push(img);
          }
        }

        let cleanText = responseText.replace(/\[GAMBAR:[^\]]+\]/gi, '').trim();
        return { text: cleanText, images: uniqueImages };
      } catch (err) {
        const message_ = err?.message || '';
        const status = err?.status || err?.code;
        const isModelError = status === 404 || /no longer available|not found|invalid model/i.test(message_);
        const isRateLimit = status === 429 || /rate limit|quota|resource_exhausted/i.test(message_);

        if (!isModelError && !isRateLimit) throw err;
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

  const adminJid = `${config.adminNumber.replace(/\D/g, '')}@s.whatsapp.net`;
  sock
    .sendMessage(adminJid, {
      text: `⚠️ *Permintaan Admin*\n\nPelanggan: ${jid}\nPesan: ${customerText}\n\nAlasan: AI sedang limit.`,
    })
    .then(() => console.log(`[Notifikasi terkirim ke admin ${adminJid}]`))
    .catch((err) => console.error(`[Gagal notifikasi admin]: ${err.message}`));
}

async function startWhatsAppBot(config, bot) {
  const authFolder = path.resolve(__dirname, 'baileys_auth');
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
        const response = await bot.reply(jid, text);
        
        await sock.sendMessage(jid, { text: response.text }, { quoted: msg });
        
        for (const img of response.images) {
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

module.exports = { createBot, loadSheetData, ensureConfig, startWhatsAppBot };
