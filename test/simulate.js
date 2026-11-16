/* Giả lập DriveApp / SpreadsheetApp / trigger để kiểm logic Code.gs mà không cần tài khoản Google.
   Chạy: node test/simulate.js
   Kịch bản: cây 3 cấp (320 file, 39 thư mục con, file trùng tên, lối tắt), một file lỗi lần đầu;
   chạy hết các lượt trigger → quét lại và chép bù → so cây đích với cây nguồn. */
const fs = require('fs'), vm = require('vm'), path = require('path'), assert = require('assert');
let NOW = 0;

// ---- Drive
let idc = 0; const nodes = {};
const mkFolder = (name, parent) => { const f = { id: 'F' + (++idc) + 'xxxxxxxxxx', name, kind: 'folder', children: [] }; nodes[f.id] = f; if (parent) parent.children.push(f); return f; };
const mkFile = (name, parent, size, mime) => { const f = { id: 'D' + (++idc) + 'xxxxxxxxxx', name, kind: 'file', size, mime: mime || 'application/pdf' }; nodes[f.id] = f; parent.children.push(f); return f; };
let copies = 0; const failOnce = new Set();
const iter = (arr) => { let i = 0; return { hasNext: () => i < arr.length, next: () => { NOW += 1000; return arr[i++]; }, getContinuationToken: () => JSON.stringify({ ids: arr.map(x => x.getId()), i }) }; };
const wrapFolder = (f) => ({
  getId: () => f.id, getName: () => f.name, getUrl: () => 'url/' + f.id,
  getFolders: () => iter(f.children.filter(c => c.kind === 'folder').map(wrapFolder)),
  getFiles: () => iter(f.children.filter(c => c.kind === 'file').map(wrapFile)),
  getFoldersByName: (n) => iter(f.children.filter(c => c.kind === 'folder' && c.name === n).map(wrapFolder)),
  createFolder: (n) => wrapFolder(mkFolder(n, f))
});
const wrapFile = (f) => ({
  getId: () => f.id, getName: () => f.name, getMimeType: () => f.mime, getSize: () => f.size, getDescription: () => '', setDescription() {}, getUrl: () => 'url/' + f.id,
  makeCopy: (n, d) => { if (failOnce.has(f.id)) { failOnce.delete(f.id); throw new Error('Access denied'); } NOW += 20000; copies++; return wrapFile(mkFile(n, nodes[d.getId()], f.size, f.mime)); }
});

// ---- Sheets
class Sheet {
  constructor(name) { this.name = name; this.rows = []; }
  getLastRow() { return this.rows.length; } getMaxRows() { return Math.max(this.rows.length, 1000); }
  setName(n) { ss.sheets[n] = this; delete ss.sheets[this.name]; this.name = n; }
  getRange(r, c, nr = 1, nc = 1) {
    const sh = this; const o = {
      setValues(v) { v.forEach((row, i) => { sh.rows[r - 1 + i] = sh.rows[r - 1 + i] || []; row.forEach((x, j) => { sh.rows[r - 1 + i][c - 1 + j] = x; }); }); return o; },
      setValue(x) { return o.setValues([[x]]); },
      getValues() { const out = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) row.push((sh.rows[r - 1 + i] || [])[c - 1 + j] ?? ''); out.push(row); } return out; },
      setFontWeight() { return o; }, setBackground() { return o; }, setFontColor() { return o; }
    }; return o;
  }
  deleteRow(r) { this.rows.splice(r - 1, 1); } deleteRows(r, n) { this.rows.splice(r - 1, n); } appendRow(v) { this.rows.push(v); }
  setFrozenRows() {} setColumnWidth() {} hideSheet() {} activate() {}
}
const ss = { sheets: {}, getSheetByName(n) { return this.sheets[n] || null; }, insertSheet(n) { return (this.sheets[n] = new Sheet(n)); }, deleteSheet(s) { delete this.sheets[s.name]; }, toast() {}, getUrl: () => 'ssurl' };
let prompts = [];
const ui = { ButtonSet: {}, Button: { OK: 1, YES: 2 }, alert: () => 2, prompt: () => { const t = prompts.shift(); return { getSelectedButton: () => 1, getResponseText: () => t }; } };
let triggers = []; const props = {}; const mails = [];
const root1 = mkFolder('Old', null), root2 = mkFolder('MyDrive', null);
const ctx = {
  console, JSON, Math, String, Object, Date: class extends Date { static now() { return NOW; } },
  Session: { getActiveUserLocale: () => 'vi', getScriptTimeZone: () => 'Asia/Ho_Chi_Minh', getEffectiveUser: () => ({ getEmail: () => 'me' }) },
  SpreadsheetApp: { getUi: () => ui, getActive: () => ss },
  DriveApp: {
    getFolderById: (id) => { if (!nodes[id]) throw new Error('not found'); return wrapFolder(nodes[id]); },
    getRootFolder: () => wrapFolder(root2),
    continueFileIterator: (t) => { const o = JSON.parse(t); const it = iter(o.ids.map(i => wrapFile(nodes[i]))); for (let k = 0; k < o.i; k++) it.next(); return it; }
  },
  PropertiesService: { getDocumentProperties: () => ({ getProperty: (k) => props[k], setProperty: (k, v) => { props[k] = v; } }) },
  ScriptApp: { getProjectTriggers: () => triggers, deleteTrigger: (t) => { triggers = triggers.filter(x => x !== t); }, newTrigger: (fn) => ({ timeBased: () => ({ after: () => ({ create: () => triggers.push({ getHandlerFunction: () => fn }) }) }) }) },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
  Utilities: { formatDate: () => '2026-11-16', sleep() {} },
  MailApp: { sendEmail: (to, subj, body) => mails.push(body) }
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8') + ';this.api={startTransfer,runBatch,rescanTransfer}', ctx);

// ---- Cây nguồn
let total = 0;
(function build(p, d) {
  for (let i = 0; i < 6; i++) { mkFile('f' + i + '.pdf', p, 1000); total++; }
  mkFile('dup.pdf', p, 5); mkFile('dup.pdf', p, 6); total += 2;                    // hai file cùng tên
  mkFile('shortcut', p, 0, 'application/vnd.google-apps.shortcut');
  if (d < 3) for (let i = 0; i < 3; i++) build(mkFolder('sub' + i, p), d + 1);
})(root1, 0);
failOnce.add(Object.values(nodes).find(n => n.kind === 'file' && n.name === 'f3.pdf').id);

const drain = () => { let rounds = 0; while (triggers.length && rounds < 500) { triggers = []; NOW += 60000; ctx.api.runBatch(); rounds++; } return rounds; };

prompts = ['https://drive.google.com/drive/folders/' + root1.id + '?usp=sharing', ''];
ctx.api.startTransfer();
const rounds = drain();
let st = JSON.parse(props.LM_DRIVE_TRANSFER);
console.log('Lượt 1:', rounds, 'lượt,', st.stats);
assert.strictEqual(st.status, 'done');
assert.ok(rounds > 1, 'phải chạy nhiều lượt (tự chạy tiếp)');
assert.strictEqual(st.stats.files, total - 1);
assert.strictEqual(st.stats.errors, 1);

ctx.api.rescanTransfer();
drain();
st = JSON.parse(props.LM_DRIVE_TRANSFER);
console.log('Sau chép bù:', st.stats, '· lệnh copy:', copies);
assert.strictEqual(copies, total, 'không được chép trùng');
assert.strictEqual(st.stats.errors, 0);

const sig = (f) => f.kind === 'file' ? f.name + ':' + f.size : '[' + f.children.filter(c => c.mime !== 'application/vnd.google-apps.shortcut').map(sig).sort().join(',') + ']';
assert.strictEqual(sig(root2.children[0]), sig(root1), 'cây đích phải giống cây nguồn');
assert.strictEqual(mails.length, 2);
console.log('✓ Đạt: ' + total + ' file, cây thư mục giống hệt, không chép trùng.');
