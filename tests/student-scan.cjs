const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const html = fs.readFileSync(path.join(__dirname, '..', 'Index.html'), 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];

function page(parameters = {class: '603'}) {
  const elements = new Map(), pending = [], calls = [];
  let focused;
  const run = new Proxy({}, {get(target, method) {
    if (method === 'withSuccessHandler') return fn => { target.success = fn; return run; };
    if (method === 'withFailureHandler') return fn => { target.failure = fn; return run; };
    return (...args) => {
      calls.push({method, args});
      pending.push({method, success: target.success, failure: target.failure});
    };
  }});
  const context = vm.createContext({
    document: {
      getElementById(id) {
        if (!elements.has(id)) elements.set(id, {
          value: '', innerHTML: '', textContent: '', disabled: false,
          focus() { focused = id; }, addEventListener() {}
        });
        return elements.get(id);
      },
      createElement() { return {set textContent(value) { this.innerHTML = String(value); }}; }
    },
    window: {},
    google: {script: {run, url: {getLocation: fn => fn({parameter: parameters})}}}
  });
  vm.runInContext(script, context);
  context.window.onload();
  return {context, elements, pending, calls, focus: () => focused};
}
const card = 'A100000001';
const progress = {registered: true, teamName: '學生隊', member1Id: '60301', member2Id: '60302', foundCount: 1, totalCount: 10};
for (const outcome of ['success', 'duplicate', 'completed', 'failure']) {
  const view = page();
  assert.equal(view.pending[0].method, 'getClassInfo');
  assert.deepEqual(Array.from(view.calls[0].args), ['603']);
  view.pending.shift().success({valid: true, classId: '603'});
  view.context.processScan(card);
  assert.deepEqual(Array.from(view.calls[1].args), ['603', card]);
  view.context.processScan('001');
  view.context.processScan(card);
  assert.equal(view.pending.length, 1, 'verification must not overlap with more scans');
  view.pending.shift().success(progress);
  assert.equal(view.focus(), 'bookBarcode');
  view.context.processScan('001');
  assert.deepEqual(Array.from(view.calls[2].args), ['603', card, '001']);
  assert.equal(view.elements.get('studentId').value, '');
  assert.equal(view.elements.get('studentId').disabled, true);
  view.context.processScan(card);
  view.context.processScan('002');
  assert.equal(view.pending.length, 1);
  const request = view.pending.shift();
  if (outcome === 'failure') request.failure({message: '書籍錯誤'});
  else request.success({...progress, isDuplicate: outcome === 'duplicate', completed: outcome === 'completed'});
  assert.equal(view.elements.get('studentId').value, '');
  assert.equal(view.elements.get('bookBarcode').value, '');
  assert.equal(view.elements.get('studentId').disabled, false);
  assert.equal(view.elements.get('bookBarcode').disabled, false);
  assert.equal(view.focus(), 'studentId');
  view.context.processScan('002');
  assert.equal(view.calls.length, 3, 'a second book requires another card');
  view.context.processScan(card);
  view.pending.shift().success(progress);
  view.context.processScan('002');
  assert.deepEqual(Array.from(view.calls.at(-1).args), ['603', card, '002']);
}
const missing = page({});
assert.equal(missing.calls.length, 0);
assert.match(missing.elements.get('result').innerHTML, /網址沒有指定班級/);
const invalid = page({class: '999'});
invalid.pending.shift().success({valid: false});
assert.match(invalid.elements.get('result').innerHTML, /班級代碼不正確/);
const rejected = page();
rejected.pending.shift().success({valid: true, classId: '603'});
rejected.context.processScan(card);
rejected.pending.shift().failure({message: 'verification failed'});
rejected.context.processScan('001');
assert.equal(rejected.calls.length, 2);
rejected.context.processScan(card);
assert.equal(rejected.pending.length, 1, 'can retry card verification after failure');
console.log('PASS: class validation and class-scoped requests; one card per book for success, duplicate, completion, error and delayed responses');
