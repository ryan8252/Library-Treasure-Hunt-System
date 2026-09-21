const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.join(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const metrics = {reads: 0, writes: 0, formats: 0, wholeColumnFormats: 0, flushes: 0};
function resetMetrics() { for (const key of Object.keys(metrics)) metrics[key] = 0; }
let now = Date.now();
let failWrite = false;
let failCachePut = false;
const cacheEntries = new Map();
const cache = {
  get(key) {
    const entry = cacheEntries.get(key);
    return entry && entry.expires > now ? entry.value : null;
  },
  put(key, value, seconds) {
    if (failCachePut) throw new Error('cache unavailable');
    assert.ok(Buffer.byteLength(value, 'utf8') <= 100000);
    cacheEntries.set(key, {value, expires: now + seconds * 1000});
  },
  remove(key) { cacheEntries.delete(key); }
};

class Sheet {
  constructor(rows = []) { this.rows = rows; }
  getLastRow() { return this.rows.length; }
  getRange(row, col, height = 1, width = 1) {
    const sheet = this;
    return {
      setNumberFormat() {
        metrics.formats++;
        if (typeof row === 'string') metrics.wholeColumnFormats++;
        return this;
      },
      setValues(values) {
        if (failWrite) throw new Error('write failed');
        metrics.writes++;
        assert.equal(values.length, height, 'write height');
        values.forEach((cells, i) => {
          assert.equal(cells.length, width, 'write width');
          sheet.rows[row - 1 + i] ||= [];
          cells.forEach((value, j) => sheet.rows[row - 1 + i][col - 1 + j] = value);
        });
        return this;
      },
      setValue(value) { return this.setValues([[value]]); },
      getDisplayValues() {
        metrics.reads++;
        return Array.from({length: height}, (_, i) =>
          Array.from({length: width}, (_, j) => String(sheet.rows[row - 1 + i]?.[col - 1 + j] ?? '')));
      },
      getValue() { metrics.reads++; return sheet.rows[row - 1]?.[col - 1] ?? ''; }
    };
  }
  deleteRow(row) { this.rows.splice(row - 1, 1); }
}
const sheets = new Map([
  ['DemoTeams', new Sheet([['teacher data']])],
  ['DemoBorrowRecords', new Sheet([['teacher records']])],
  ['BorrowRecords', new Sheet([['library_card', 'barcode', 'found_at', 'status', 'class_id']])]
]);
const teacherSnapshot = JSON.stringify([...sheets].filter(([name]) => name.startsWith('Demo')));
const spreadsheet = {
  getSheetByName: name => sheets.get(name),
  insertSheet(name) { const sheet = new Sheet(); sheets.set(name, sheet); return sheet; },
  getSpreadsheetTimeZone: () => 'Asia/Taipei'
};
let locked = false;
const backend = vm.createContext({
  SpreadsheetApp: {getActiveSpreadsheet: () => spreadsheet, flush() { metrics.flushes++; }},
  CacheService: {getScriptCache: () => cache},
  Date: class extends Date { static now() { return now; } },
  LockService: {getScriptLock: () => ({
    waitLock() { assert.equal(locked, false); locked = true; },
    releaseLock() { locked = false; }
  })},
  Utilities: {formatDate: () => '12:00:00'},
  ScriptApp: {getService: () => ({getUrl: () => 'https://example.test/exec'})},
  HtmlService: {createHtmlOutputFromFile: file => ({file, setTitle() { return this; }})}
});
vm.runInContext(read('Backend.gs'), backend);

const card = 'A100000001';
assert.equal(backend.getLeaderboardData('603').teamCount, 0);
assert.equal(backend.getStudentProgress('603', card).registered, false);
assert.throws(() => backend.registerTeam('999', card, '隊名', '60301', '60302'));
assert.throws(() => backend.registerTeam('603', card, '超過五個字隊伍', '60301', '60302'));
assert.throws(() => backend.registerTeam('603', card, '隊名', '王老師', '60302'));
assert.throws(() => backend.registerTeam('603', card, '隊名', '60301', ''));
assert.throws(() => backend.registerTeam('603', card, '隊名', '60301', '60301'));
backend.registerTeam('603', card, '六三隊', '60301', '60302');
backend.registerTeam('604', card, '六四隊', '60401', '60402');
backend.registerTeam('603', 'A100000002', '第二隊', '60303', '60304');
assert.equal(backend.getTeamByCard('603', card).teamName, '六三隊');
assert.equal(backend.getTeamByCard('604', card).teamName, '六四隊');
assert.equal(backend.getTeamByCard('605', card), null);
for (let n = 1; n <= 11; n++) backend.addTreasureBook(String(n).padStart(3, '0'));
assert.throws(() => backend.saveBorrowRecord('605', card, '001'));
assert.throws(() => backend.saveBorrowRecord('603', card, '999'));
backend.saveBorrowRecord('603', card, '００１');
assert.equal(backend.saveBorrowRecord('604', card, '001').isDuplicate, false);
assert.equal(backend.saveBorrowRecord('603', card, '001').isDuplicate, true);
assert.equal(backend.getStudentProgress('603', card).foundCount, 1);
assert.equal(backend.getStudentProgress('604', card).foundCount, 1);
const board603 = backend.getLeaderboardData('603');
assert.equal(board603.teamCount, 2);
assert.ok(board603.teams.every(team => team.classId === '603'));
assert.equal(backend.getLeaderboardData('604').teamCount, 1);
assert.deepEqual(sheets.get('BorrowRecords').rows.slice(1).map(row => row[4]), ['603', '604']);

resetMetrics();
backend.getStudentProgress('603', card);
backend.getTeamByCard('604', card);
backend.getLeaderboardData('603');
assert.equal(metrics.reads, 0);
assert.equal(metrics.formats, 0);
assert.equal(metrics.writes, 0);
resetMetrics();
backend.saveBorrowRecord('603', card, '002');
assert.equal(metrics.reads, 0);
assert.equal(metrics.writes, 1);
assert.equal(metrics.wholeColumnFormats, 0);
assert.equal(metrics.flushes, 1);
console.log('Warm student scan:', JSON.stringify(metrics));
assert.equal(backend.getLeaderboardData('603').teams.find(team => team.libraryCard === card).foundCount, 2);
assert.equal(backend.getStudentProgress('604', card).foundCount, 1);
for (let n = 3; n <= 10; n++) backend.saveBorrowRecord('603', card, String(n).padStart(3, '0'));
assert.equal(backend.getStudentProgress('603', card).completed, true);
assert.equal(backend.saveBorrowRecord('603', card, '011').alreadyCompleted, true);
assert.equal(backend.saveBorrowRecord('604', card, '002').alreadyCompleted, false);
assert.equal(backend.getStudentProgress('604', card).foundCount, 2);
console.log('PASS: original student registration rules, class-isolated teams, duplicate checks, completion and rankings');

cacheEntries.clear();
resetMetrics();
assert.equal(backend.getStudentProgress('603', card).foundCount, 10);
assert.ok(metrics.reads > 0);
assert.equal(metrics.formats, 0);
assert.equal(backend.saveBorrowRecord('604', card, '001').isDuplicate, true);
now += 11000;
resetMetrics();
backend.getLeaderboardData('604');
assert.ok(metrics.reads > 0);
backend.deleteTreasureBook('002');
assert.equal(backend.getStudentProgress('603', card).foundCount, 9);
assert.equal(backend.getStudentProgress('604', card).foundCount, 1);
backend.addTreasureBook('002');
assert.equal(backend.getStudentProgress('603', card).foundCount, 10);
assert.equal(backend.registerTeam('603', card, '改名隊', '60301', '60302').updated, true);
assert.equal(backend.getTeamByCard('603', card).teamName, '改名隊');
assert.equal(backend.getTeamByCard('604', card).teamName, '六四隊');

failCachePut = true;
assert.equal(backend.saveBorrowRecord('604', card, '003').foundCount, 3);
assert.equal(backend.getStudentProgress('604', card).foundCount, 3);
failCachePut = false;
backend.getStudentProgress('604', card);
failWrite = true;
assert.throws(() => backend.saveBorrowRecord('604', card, '004'), /write failed/);
failWrite = false;
assert.equal(locked, false);
assert.equal(backend.getStudentProgress('604', card).foundCount, 3);
backend.saveBorrowRecord('604', card, '004');
const row604 = sheets.get('Teams').rows.find(row => row[0] === card && row[5] === '604');
row604[1] = '手動改名';
now += 8000;
backend.saveBorrowRecord('604', card, '005');
now += 3000;
assert.equal(backend.getTeamByCard('604', card).teamName, '手動改名');
cacheEntries.set('student-state-v1', {value: 'broken JSON', expires: now + 10000});
assert.equal(backend.getStudentProgress('604', card).foundCount, 5);
assert.equal(JSON.stringify([...sheets].filter(([name]) => name.startsWith('Demo'))), teacherSnapshot);
console.log('PASS: eviction, fixed expiry, mutation invalidation, failure recovery and teacher data isolation');

// Polling begins only after valid class resolution, and display text uses the same interval.
const elements = new Map(), pending = [], timers = [], calls = [];
const run = new Proxy({}, {get(target, method) {
  if (method === 'withSuccessHandler') return fn => { target.success = fn; return run; };
  if (method === 'withFailureHandler') return fn => { target.failure = fn; return run; };
  return (...args) => { calls.push({method, args}); pending.push({...target}); };
}});
const page = vm.createContext({
  document: {
    getElementById(id) { if (!elements.has(id)) elements.set(id, {}); return elements.get(id); },
    createElement() { return {set textContent(value) { this.innerHTML = String(value); }}; }
  },
  google: {script: {run, url: {getLocation: fn => fn({parameter: {class: '603'}})}}},
  window: {}, setInterval: (fn, ms) => timers.push({fn, ms})
});
vm.runInContext(read('Leaderboard.html').match(/<script>([\s\S]*?)<\/script>/)[1], page);
page.window.onload();
assert.equal(timers.length, 0);
assert.equal(calls[0].method, 'getClassInfo');
pending.shift().success({valid: true, classId: '603'});
assert.equal(timers[0].ms, 2000);
assert.equal(calls[1].method, 'getLeaderboardData');
assert.deepEqual(Array.from(calls[1].args), ['603']);
timers[0].fn();
assert.equal(pending.length, 1);
pending.shift().success(backend.getLeaderboardData('603'));
assert.ok(elements.get('updateInfo').textContent.includes('每 2 秒'));
timers[0].fn();
assert.equal(pending.length, 1);
pending.shift().failure({message: 'test error'});
timers[0].fn();
assert.equal(pending.length, 1, 'polling retries after error');
assert.doesNotMatch(read('Leaderboard.html'), /5 秒|5000/);
console.log('PASS: class-aware two-second polling, matching text, no overlapping requests and retries');

// Large history falls back to Sheets; class isolation and deduplication remain intact.
for (let n = 0; n < 2000; n++) {
  sheets.get('BorrowRecords').rows.push([card, '001', '', '尋寶成功', '603']);
}
cacheEntries.clear();
assert.equal(backend.getStudentProgress('603', card).foundCount, 10);
assert.equal(cache.get('student-state-v1'), null);
assert.equal(backend.getStudentProgress('604', card).foundCount, 5);
assert.equal(backend.saveBorrowRecord('604', card, '001').isDuplicate, true);
assert.equal(backend.saveBorrowRecord('604', card, '006').foundCount, 6);
assert.equal(locked, false);
console.log('PASS: oversized history fallback retains separate class progress and duplicate protection');
