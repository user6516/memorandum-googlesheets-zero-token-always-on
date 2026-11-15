// 本地模拟验证；不连接 Google，也不会实际发送邮件。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const crypto = require('node:crypto');
const source = fs.readFileSync(__dirname + '/../src/Code.gs', 'utf8');
const HOUR = 3600000;
let now = Date.parse('2026-11-20T09:00:00+08:00');
const storage = {};
const emails = [];
let failMail = false;
let sheetReads = 0;
let metadataWrites = 0;
let timeWrites = 0;
let formulas = [];
let scriptLockAvailable = true;
let rows = [];
class TestDate extends Date {
  constructor(...args) { super(...(args.length ? args : [now])); }
  static now() { return now; }
}
const props = {
  getProperty: key => Object.hasOwn(storage, key) ? storage[key] : null,
  setProperty: (key, value) => { storage[key] = value; },
  setProperties: values => Object.assign(storage, values),
  getProperties: () => ({ ...storage }),
  deleteProperty: key => { delete storage[key]; },
};
const context = vm.createContext({
  Date: TestDate,
  console,
  PropertiesService: { getScriptProperties: () => props },
  LockService: {
    getScriptLock: () => ({ waitLock() {}, tryLock: () => scriptLockAvailable, releaseLock() {} }),
    getUserLock: () => ({ tryLock: () => true, releaseLock() {} }),
  },
  Utilities: {
    getUuid: () => crypto.randomUUID(),
    newBlob: value => ({ getBytes: () => Buffer.from(value, 'utf8') }),
    formatDate: value => value.toISOString(),
  },
  MailApp: {
    getRemainingDailyQuota: () => 100,
    sendEmail: mail => { if (failMail) throw new Error('模拟发信失败'); emails.push(mail); },
  },
  SpreadsheetApp: {
    openById() {
      sheetReads++;
      return { getSheetByName: () => ({
        getLastRow: () => rows.length + 1,
        getRange: (row, column, count, width) => ({
          getValues: () => rows.slice(row - 2, row - 2 + count).map(value => value.slice(column - 1, column - 1 + width)),
          getFormulas: () => rows.map((_, i) => formulas[i] || ['', '', '']),
          setValues: values => {
            if (column === 5) metadataWrites++;
            if (column === 1) timeWrites++;
            values.forEach((value, i) => value.forEach((cell, j) => { rows[row - 2 + i][column - 1 + j] = cell; }));
          },
        }),
      }) };
    },
  },
});
vm.runInContext(source + "\nCONFIG.recipient = 'test@example.com';", context);
const run = code => vm.runInContext(code, context);
const state = () => JSON.parse(JSON.stringify(run('readState_()')));

// 北京时间、日期校验、48小时严格边界、排序。
assert.equal(run("parseDue_('2026/11/20 09:00')"), now);
assert.equal(run("parseDue_('2026/02/30 09:00')"), null);
assert.equal(run("parseDue_('2026/11/20 25:00')"), null);
assert.equal(run("parseDue_('2026/11/20')"), null);
context.inputTasks = [
  { id: 'expired', due: now - 1 }, { id: 'now', due: now },
  { id: 'edge', due: now + 48 * HOUR }, { id: 'inside', due: now + 48 * HOUR - 1 },
  { id: 'soon', due: now + HOUR },
];
assert.deepEqual(Array.from(run('summaryTasks_(inputTasks, Date.now()).map(t => t.id)')), ['soon', 'inside']);
assert.equal(run('nextSummary_(Date.now())'), Date.parse('2026-11-20T14:20:00+08:00'));
now = Date.parse('2026-11-20T20:20:00+08:00');
assert.equal(run('nextSummary_(Date.now())'), Date.parse('2026-11-21T09:00:00+08:00'));

// 持久化分块能保存中文和 emoji，过大写入拒绝且保留原状态。
context.snapshot = { tasks: [{ id: 'long', title: '备忘🙂'.repeat(1000), due: now + 8 * HOUR }], sent: {}, nextSummary: now + HOUR };
run('saveState_(snapshot)');
assert.deepEqual(state(), context.snapshot);
const oldManifest = storage.AI_MEMO_V1_manifest;
assert.throws(() => run("saveState_({tasks:[],sent:{},text:'中'.repeat(60000)})"), /过大/);
assert.equal(storage.AI_MEMO_V1_manifest, oldManifest);

// 汇总与单项重合各发一封；定时执行不读表；重复 tick 不重复发。
context.snapshot = { tasks: [{ id: 'task', title: '论文', due: now + 5 * HOUR }], sent: {}, nextSummary: now };
run('saveState_(snapshot); tick(); tick();');
assert.equal(emails.length, 2);
assert.equal(sheetReads, 0);
assert.match(emails[0].body, /事务：论文/);
assert.match(emails[1].subject, /前5小时/);
assert.ok(state().sent['task:' + (now + 5 * HOUR)]);

// 修改标题不重发；修改 DDL 为另一已到提醒时间的时刻会重新提醒。
run("var s = readState_(); s.tasks[0].title = '改名'; saveState_(s); tick();");
assert.equal(emails.length, 2);
run('var s = readState_(); s.tasks[0].due -= 1000; saveState_(s); tick();');
assert.equal(emails.length, 3);

// 发信失败不能提前记作已发送，下次 tick 可以重试。
run('var s = readState_(); s.tasks[0].due -= 1000; saveState_(s);');
failMail = true;
assert.throws(() => run('tick()'), /模拟发信失败/);
assert.equal(Object.keys(state().sent).length, 2);
failMail = false;
run('tick()');
assert.equal(emails.length, 4);

// 编辑同步：自动生成编号、复制行重新编号、完成/删除后从快照消失、非法日期排除。
storage.AI_MEMO_V1_spreadsheet = 'mock-sheet';
rows = [['', '新任务', '2026/11/21 18:00', false, '', '']];
run('syncNow()');
const id = rows[0][4];
assert.equal(rows[0][0], '周六 18:00');
assert.match(id, /^[0-9a-f-]{36}$/);
assert.equal(state().tasks.length, 1);
const writesBefore = metadataWrites;
const timesBefore = timeWrites;
run('syncNow()');
assert.equal(metadataWrites, writesBefore, '任务和校验结果未变时不写回元数据');
assert.equal(timeWrites, timesBefore, '自动时间未变时不写回');
rows.push([...rows[0]]);
run('syncNow()');
assert.notEqual(rows[0][4], rows[1][4]);
rows[0][3] = true;
rows[1][2] = '2026/02/30 10:00';
run('syncNow()');
assert.equal(state().tasks.length, 0);
assert.match(rows[1][5], /DDL 无效/);
assert.equal(rows[1][0], '');
// 取消完成、改期、清空日期、日期公式与左侧自动列被误改。
rows[0][3] = false;
rows[0][2] = '2026/11/22 00:05';
run('syncNow()');
assert.equal(state().tasks.length, 1);
assert.equal(rows[0][0], '周日 00:05');
rows[0][0] = '误改';
run('syncNow()');
assert.equal(rows[0][0], '周日 00:05');
rows[0][2] = '';
run('syncNow()');
assert.equal(rows[0][0], '');
rows[0][2] = '2026/11/23 09:08';
formulas = [['', '=NOW()', '']];
run('syncNow()');
assert.equal(rows[0][0], '');
assert.equal(state().tasks.length, 0);
assert.match(rows[0][5], /不使用公式/);
formulas = [];
// D 列完成编辑也触发同步；自动元数据编辑不触发。
let editsSynced = 0;
context.countEdit = () => editsSynced++;
run('var realSync = syncNow; syncNow = countEdit;');
context.edit = { range: { getSheet: () => ({ getName: () => '我的备忘录' }), getColumn: () => 4, getLastRow: () => 2 } };
run('memoEdited(edit)');
assert.equal(editsSynced, 1);
context.edit.range.getColumn = () => 5;
run('memoEdited(edit)');
assert.equal(editsSynced, 1);
run('syncNow = realSync;');
rows = [];
run('syncNow()');
assert.equal(state().tasks.length, 0);

// 没有任务仍发送汇总；延迟跨过多个时间只补一封，并恢复未来时间。
const before = emails.length;
run('var s = readState_(); s.nextSummary = Date.now() - 24 * 3600000; saveState_(s); tick();');
assert.equal(emails.length, before + 1);
assert.match(emails.at(-1).body, /暂无/);
assert.ok(state().nextSummary > now);

// 10分钟调度配置；后台锁冲突不等待、不发邮件，释放后下次可补发。
assert.equal(run('CONFIG.checkMinutes'), 10);
const mailsBeforeLock = emails.length;
run('var s = readState_(); s.nextSummary = Date.now(); saveState_(s);');
scriptLockAvailable = false;
run('tick()');
assert.equal(emails.length, mailsBeforeLock);
scriptLockAvailable = true;
run('tick()');
assert.equal(emails.length, mailsBeforeLock + 1);

// 运行统计捕获正常/失败的实际耗时；超过保留窗口的统计清除。
context.advanceClock = ms => { now += ms; };
storage.AI_MEMO_V1_runtime = '[]';
run('measured_(function () { advanceClock(1200); });');
assert.throws(() => run("measured_(function () { advanceClock(300); throw new Error('test failure'); });"), /test failure/);
let buckets = JSON.parse(storage.AI_MEMO_V1_runtime);
assert.equal(buckets[0].count, 2);
assert.equal(buckets[0].ms, 1500);
assert.equal(buckets[0].failed, 1);
now += 26 * HOUR;
run('measured_(function () {});');
buckets = JSON.parse(storage.AI_MEMO_V1_runtime);
assert.equal(buckets.length, 1);
assert.equal(buckets[0].count, 1);
console.log('通过：原提醒逻辑、去重、失败重试、无变化不写回、10分钟配置、锁冲突不等待且可补发、耗时/失败统计及过期清理。');

// 新建、旧表升级、重复初始化保留完成状态/编号，以及条件格式的范围和去重。
function layoutSheet(data) {
  return {
    data, maxRows: 20, maxColumns: 26, rules: [], inserted: 0,
    getLastRow() { return this.data.length; },
    getMaxRows() { return this.maxRows; }, getMaxColumns() { return this.maxColumns; },
    insertRowsAfter(_, count) { this.maxRows += count; },
    insertColumnBefore(column) { assert.equal(column, 1); this.inserted++; this.maxColumns++; this.data.forEach(row => row.unshift('')); },
    getRange(row, column, count, width) {
      const sheet = this;
      const range = {
        row, column, count, width,
        getValues: () => Array.from({ length: count }, (_, i) => Array.from({ length: width }, (_, j) => sheet.data[row - 1 + i]?.[column - 1 + j] ?? '')),
        setValues(values) { values.forEach((cells, i) => { sheet.data[row - 1 + i] ||= []; cells.forEach((cell, j) => { sheet.data[row - 1 + i][column - 1 + j] = cell; }); }); return range; },
        setFontWeight() { return range; }, setNumberFormat() { return range; },
        insertCheckboxes() { for (let i = 0; i < count; i++) { sheet.data[row - 1 + i] ||= []; sheet.data[row - 1 + i][column - 1] = false; } return range; },
        setDataValidation() { return range; },
      };
      return range;
    },
    setFrozenRows() {}, setColumnWidth() {}, hideColumns(column) { this.hidden = column; },
    getConditionalFormatRules() { return this.rules; }, setConditionalFormatRules(rules) { this.rules = rules; },
  };
}
context.SpreadsheetApp.newDataValidation = () => ({ requireCheckbox() { return this; }, build() { return {}; } });
context.SpreadsheetApp.newConditionalFormatRule = () => ({
  whenFormulaSatisfied(value) { this.formula = value; return this; },
  setBackground(value) { this.background = value; return this; },
  setFontColor(value) { this.color = value; return this; },
  setStrikethrough(value) { this.strike = value; return this; },
  setRanges(value) { this.ranges = value; return this; },
  build() { return { ...this, getBooleanCondition: () => ({ getCriteriaValues: () => [this.formula] }) }; },
});
const oldHeaders = ['事务', 'DDL', '已完成', '任务编号（自动）', '校验结果（自动）'];
const existing = ['旧任务', '2026/11/20 18:00', true, id, '已完成，不再提醒'];
context.layout = layoutSheet([oldHeaders.slice(), existing.slice()]);
const userRule = { getBooleanCondition: () => null };
context.layout.rules = [userRule];
const sentBeforeMigration = JSON.stringify(state().sent);
run('prepareSheet_(layout); prepareSheet_(layout);');
assert.equal(context.layout.inserted, 1);
assert.deepEqual(context.layout.data[1], ['', ...existing]);
assert.equal(context.layout.hidden, 5);
assert.equal(context.layout.rules.length, 2);
assert.equal(context.layout.rules[1], userRule);
assert.equal(JSON.stringify(state().sent), sentBeforeMigration);
const completed = context.layout.rules[0];
assert.match(completed.formula, /\$D2=TRUE/);
assert.match(completed.formula, /\$D2="已完成"/);
assert.equal(completed.background, '#e0e0e0');
assert.equal(completed.strike, true);
assert.equal(completed.ranges[0].row, 2);
assert.equal(completed.ranges[0].count, context.layout.maxRows - 1);
assert.equal(completed.ranges[0].width, context.layout.maxColumns);
context.layout = layoutSheet([]);
run('prepareSheet_(layout)');
assert.equal(context.layout.data[0].length, 6);
assert.equal(context.layout.data[1][3], false);
context.layout = layoutSheet([['其他表头']]);
assert.throws(() => run('prepareSheet_(layout)'), /表头不同/);
assert.deepEqual(context.layout.data, [['其他表头']]);
assert.equal(run("formatShortDue_(parseDue_('2026/11/23 09:08'))"), '周一 09:08');
assert.equal(run("formatShortDue_(Date.parse('2026-11-21T16:05:00Z'))"), '周日 00:05');
console.log('通过：自动周几及24小时时间、改期/清空/无效日期、D列触发、旧表无损升级、重复初始化、灰底删除线规则及范围。');
