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
        getRange: (row, column) => ({
          getValues: () => rows.map(value => [...value]),
          getFormulas: () => rows.map(() => ['', '', '']),
          setValues: values => { metadataWrites++; values.forEach((value, i) => { rows[i][3] = value[0]; rows[i][4] = value[1]; }); },
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
rows = [['新任务', '2026/11/21 18:00', false, '', '']];
run('syncNow()');
const id = rows[0][3];
assert.match(id, /^[0-9a-f-]{36}$/);
assert.equal(state().tasks.length, 1);
const writesBefore = metadataWrites;
run('syncNow()');
assert.equal(metadataWrites, writesBefore, '任务和校验结果未变时不写回元数据');
rows.push([...rows[0]]);
run('syncNow()');
assert.notEqual(rows[0][3], rows[1][3]);
rows[0][2] = true;
rows[1][1] = '2026/02/30 10:00';
run('syncNow()');
assert.equal(state().tasks.length, 0);
assert.match(rows[1][4], /DDL 无效/);
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
