const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.join(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');

class Sheet {
  constructor(rows = []) { this.rows = rows; }
  getLastRow() { return this.rows.length; }
  getRange(row, col, height = 1, width = 1) {
    const sheet = this;
    return {
      setNumberFormat() { return this; },
      setValues(values) {
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
        return Array.from({length: height}, (_, i) =>
          Array.from({length: width}, (_, j) => String(sheet.rows[row - 1 + i]?.[col - 1 + j] ?? '')));
      },
      getValue() { return sheet.rows[row - 1]?.[col - 1] ?? ''; }
    };
  }
  deleteRow(row) { this.rows.splice(row - 1, 1); }
}
const sheets = new Map([
  ['Teams', new Sheet([['student data']])],
  ['BorrowRecords', new Sheet([['student records']])]
]);
const studentSnapshot = JSON.stringify([...sheets]);
const spreadsheet = {
  getSheetByName: name => sheets.get(name),
  insertSheet(name) { const sheet = new Sheet(); sheets.set(name, sheet); return sheet; },
  getSpreadsheetTimeZone: () => 'Asia/Taipei'
};
let locked = false;
const backend = vm.createContext({
  SpreadsheetApp: {getActiveSpreadsheet: () => spreadsheet},
  LockService: {getScriptLock: () => ({
    waitLock() { assert.equal(locked, false); locked = true; },
    releaseLock() { locked = false; }
  })},
  Utilities: {formatDate: () => '12:00:00'},
  ScriptApp: {getService: () => ({getUrl: () => 'https://example.test/exec'})},
  HtmlService: {createHtmlOutputFromFile: file => ({file, setTitle() { return this; }})}
});
vm.runInContext(read('Backend.gs'), backend);
assert.equal(backend.getLeaderboardData().teamCount, 0);
assert.equal(backend.getStudentProgress('A100000001').registered, false);
const longName = '老師一起來圖書館找寶藏'.repeat(10);
const card = 'A100000001';
assert.equal(backend.registerTeam(card, longName, '王小明', '007').success, true);
assert.equal(backend.getTeamByCard(card).teamName, longName);
assert.equal(backend.getTeamByCard(card).member2Id, '007');
assert.equal(backend.registerTeam(card, longName, '同名老師', '同名老師').updated, true);
assert.equal(sheets.get('DemoTeams').getLastRow(), 2);
assert.throws(() => backend.registerTeam(card, '', '王', '陳'));
assert.throws(() => backend.registerTeam(card, '小隊', '', '陳'));
assert.throws(() => backend.saveBorrowRecord('A100000099', '001'));
for (let n = 1; n <= 11; n++) backend.addTreasureBook(String(n).padStart(3, '0'));
assert.throws(() => backend.saveBorrowRecord(card, '999'));
assert.equal(backend.saveBorrowRecord(card, '００１').foundCount, 1);
assert.equal(backend.saveBorrowRecord(card, '001').isDuplicate, true);
assert.equal(sheets.get('DemoBorrowRecords').getLastRow(), 2);
for (let n = 2; n <= 10; n++) backend.saveBorrowRecord(card, String(n).padStart(3, '0'));
assert.equal(backend.getStudentProgress(card).completed, true);
assert.equal(backend.saveBorrowRecord(card, '011').alreadyCompleted, true);
assert.equal(sheets.get('DemoBorrowRecords').getLastRow(), 11);
backend.registerTeam('A100000002', '第二隊', '中文姓名', '123456789');
backend.registerTeam('A100000003', '<script>第三隊</script>', '老師甲', '老師乙');
backend.saveBorrowRecord('A100000002', '001');
backend.saveBorrowRecord('A100000003', '001');
assert.equal(backend.getStudentProgress('A100000002').foundCount, 1);
const rankings = backend.getLeaderboardData();
assert.deepEqual(Array.from(rankings.teams, team => team.rank), [1, 2, 2]);
assert.deepEqual(Array.from(rankings.teams, team => team.foundCount), [10, 1, 1]);
assert.equal(JSON.stringify([...sheets].filter(([name]) => ['Teams', 'BorrowRecords'].includes(name))), studentSnapshot);
assert.equal(backend.doGet({parameter: {page: 'register'}}).file, 'Register');
assert.equal(backend.doGet({parameter: {page: 'leaderboard'}}).file, 'Leaderboard');
assert.equal(backend.doGet({parameter: {}}).file, 'Index');
console.log('PASS: registration, Unicode names, arbitrary IDs, updates, isolated storage, scan scoring and rankings');

function page(file, parameters = {}, deferred = false) {
  const elements = new Map();
  const calls = [];
  const timers = [];
  const opened = [];
  const pending = [];
  const element = () => ({value: '', innerHTML: '', textContent: '', disabled: false,
    focus() { this.focused = true; }, select() {}, addEventListener() {}});
  const run = new Proxy({}, {get(_, method) {
    if (method === 'withSuccessHandler') return fn => { run.success = fn; return run; };
    if (method === 'withFailureHandler') return fn => { run.failure = fn; return run; };
    if (method === 'success' || method === 'failure') return _[method];
    return (...args) => {
      calls.push({method, args});
      assert.equal(typeof backend[method], 'function', 'known backend method ' + method);
      const success = run.success, failure = run.failure;
      const respond = () => {
        let result;
        try { result = backend[method](...args); }
        catch (error) { if (failure) return failure(error); throw error; }
        if (success) success(result);
      };
      if (deferred) pending.push(respond);
      else respond();
    };
  }});
  const context = vm.createContext({
    document: {
      getElementById(id) { if (!elements.has(id)) elements.set(id, element()); return elements.get(id); },
      createElement() { return {set textContent(value) {
        this.innerHTML = String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      }}; }
    },
    window: {open: (...args) => opened.push(args)},
    google: {script: {run, url: {getLocation: fn => fn({parameter: parameters})}}},
    setInterval: (fn, ms) => timers.push({fn, ms})
  });
  const html = read(file);
  assert.doesNotMatch(html, /currentClassId|getClassInfo|parameter\.class/);
  vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], context);
  context.window.onload();
  return {context, elements, calls, timers, opened, pending};
}
const registration = page('Register.html');
assert.doesNotMatch(read('Register.html'), /maxlength|inputmode="numeric"/);
registration.elements.get('libraryCard').value = 'A100000004';
registration.elements.get('teamName').value = longName;
registration.elements.get('member1').value = '黃老師';
registration.elements.get('member2').value = '12';
registration.context.submitRegistration();
assert.equal(registration.calls[0].args.length, 4);
assert.equal(registration.opened[0][0], 'https://example.test/exec?registered=A100000004');
const scan = page('Index.html', {registered: 'A100000004'});
scan.context.processScan('001');
assert.equal(scan.calls.filter(call => call.method === 'saveBorrowRecord').length, 0);
scan.context.processScan('A100000004');
scan.context.processScan('001');
assert.equal(backend.getStudentProgress('A100000004').foundCount, 1);
assert.equal(scan.calls.find(call => call.method === 'saveBorrowRecord').args.length, 2);
function assertCardConsumed(view) {
  assert.equal(view.elements.get('studentId').value, '');
  assert.equal(view.elements.get('studentId').disabled, false);
  assert.equal(view.elements.get('studentId').focused, true);
  assert.equal(view.elements.get('bookBarcode').value, '');
  const count = view.calls.length;
  view.context.processScan('002');
  assert.equal(view.calls.length, count, 'another book must require a new card scan');
}
assertCardConsumed(scan);
for (const barcode of ['001', '999', '002']) {
  scan.context.processScan('A100000004');
  scan.context.processScan(barcode);
  assertCardConsumed(scan); // duplicate, rejected, and successful books
}
assert.equal(backend.getStudentProgress('A100000004').foundCount, 2);
const completedScan = page('Index.html');
completedScan.context.processScan(card);
completedScan.context.processScan('001');
assertCardConsumed(completedScan);
const slowScan = page('Index.html', {}, true);
slowScan.context.processScan('A100000004');
slowScan.context.processScan('A100000004');
slowScan.context.processScan('003');
assert.equal(slowScan.pending.length, 1, 'do not overlap card verification');
slowScan.pending.shift()();
slowScan.context.processScan('003');
assert.equal(slowScan.elements.get('studentId').value, '');
assert.equal(slowScan.elements.get('studentId').disabled, true);
slowScan.context.processScan('A100000004');
slowScan.context.processScan('004');
assert.equal(slowScan.pending.length, 1, 'do not scan while a book is saving');
slowScan.pending.shift()();
assertCardConsumed(slowScan);
console.log('PASS: each book consumes card verification, including duplicate/error/completed results and delayed responses');
const directScan = page('Index.html');
directScan.context.processScan('A100000004');
assert.equal(directScan.calls[0].method, 'getStudentProgress');
const board = page('Leaderboard.html');
assert.equal(board.calls[0].args.length, 0);
assert.equal(board.timers[0].ms, 5000);
assert.ok(board.elements.get('board').innerHTML.includes('&lt;script&gt;'));
assert.ok(board.elements.get('board').innerHTML.includes('黃老師'));
board.timers[0].fn();
assert.equal(board.calls.length, 2);
for (const file of ['Admin.html']) new vm.Script(read(file).match(/<script>([\s\S]*?)<\/script>/)[1]);
console.log('PASS: class-free registration redirect, required card scan, direct scan page, leaderboard refresh and escaped names');

// A one-person team must work through registration, scanning and display.
const solo = page('Register.html');
solo.elements.get('libraryCard').value = 'A100000005';
solo.elements.get('teamName').value = '單人小隊';
solo.elements.get('member1').value = '<王老師>';
solo.elements.get('member2').value = '   ';
solo.context.submitRegistration();
assert.equal(solo.opened[0][0], 'https://example.test/exec?registered=A100000005');
assert.equal(backend.getTeamByCard('A100000005').member2Id, '');
assert.ok(solo.elements.get('result').innerHTML.includes('隊員：<b>&lt;王老師&gt;</b>'));
const soloScan = page('Index.html', {registered: 'A100000005'});
assert.equal(soloScan.elements.get('teamMembersDisplay').textContent, '隊員：<王老師>');
soloScan.context.processScan('A100000005');
soloScan.context.processScan('001');
assert.equal(backend.getStudentProgress('A100000005').foundCount, 1);
assertCardConsumed(soloScan);
const soloBoard = page('Leaderboard.html');
assert.ok(soloBoard.elements.get('board').innerHTML.includes('<div class="members">&lt;王老師&gt;</div>'));
assert.ok(soloBoard.elements.get('board').innerHTML.includes('黃老師、12'));
assert.throws(() => backend.registerTeam('A100000006', '空隊', '', ''));
assert.throws(() => backend.registerTeam('A100000006', '空隊', '', '只有隊員二'));
const empty = page('Register.html');
empty.elements.get('libraryCard').value = 'A100000006';
empty.elements.get('teamName').value = '空隊';
empty.context.submitRegistration();
assert.equal(empty.calls.length, 0);
assert.equal(backend.registerTeam('A100000005', '單人小隊', '王老師', '陳老師').updated, true);
assert.equal(backend.getTeamByCard('A100000005').member2Id, '陳老師');
backend.registerTeam('A100000005', '單人小隊', '王老師');
assert.equal(backend.getTeamByCard('A100000005').member2Id, '');
assert.equal(backend.getStudentProgress('A100000005').foundCount, 1);
console.log('PASS: solo registration, optional second member, escaped solo display, two-person display, updates and required first member');
