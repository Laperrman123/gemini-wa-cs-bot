const assert = require('assert');
const { createBot, loadSheetData, ensureConfig, startWhatsAppBot } = require('./index');

assert.strictEqual(typeof createBot, 'function');
assert.strictEqual(typeof loadSheetData, 'function');
assert.strictEqual(typeof ensureConfig, 'function');
assert.strictEqual(typeof startWhatsAppBot, 'function');

console.log('[✓] Self-check passed');
