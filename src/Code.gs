// 修改此处为你的收件邮箱，然后在 Google Apps Script 中运行 setup。
const CONFIG = {
  recipient: 'YOUR_EMAIL_HERE',
  sheetName: '我的备忘录',
  timezone: 'Asia/Shanghai',
  summaryTimes: ['09:00', '14:20', '20:20'],
  windowHours: 48,
  advanceHours: 5,
  checkMinutes: 10, // 每 10 分钟检查一次；可改为 5。
};
const PREFIX = 'AI_MEMO_V1_';
const HOUR = 3600000;

function setup() {
  validateConfig_();
  withLock_(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('请从 Google 表格的“扩展程序 → Apps Script”打开本程序。');
    const props = PropertiesService.getScriptProperties();
    const previousId = props.getProperty(PREFIX + 'spreadsheet');
    if (previousId && previousId !== ss.getId()) throw new Error('本程序已绑定另一张表格。');
    let sheet = ss.getSheetByName(CONFIG.sheetName);
    if (!sheet) sheet = ss.insertSheet(CONFIG.sheetName);
    prepareSheet_(sheet);
    // Date 单元格由表格时区解释；统一为北京时间。
    ss.setSpreadsheetTimeZone(CONFIG.timezone);
    props.setProperty(PREFIX + 'spreadsheet', ss.getId());
    syncUnlocked_();
    const handlers = ['memoEdited', 'memoChanged', 'tick'];
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
      if (handlers.indexOf(trigger.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(trigger);
    });
    ScriptApp.newTrigger('memoEdited').forSpreadsheet(ss).onEdit().create();
    ScriptApp.newTrigger('memoChanged').forSpreadsheet(ss).onChange().create();
    ScriptApp.newTrigger('tick').timeBased().everyMinutes(CONFIG.checkMinutes).create();
    console.log('已启用：编辑后同步、结构变化后同步、每 ' + CONFIG.checkMinutes + ' 分钟检查提醒计划。');
  });
}

// 兼容原五列表格：只在最左侧插入自动时间列，保留原数据与任务编号。
function prepareSheet_(sheet) {
  const oldHeaders = ['事务', 'DDL', '已完成', '任务编号（自动）', '校验结果（自动）'];
  const headers = ['时间（自动）', '事务', 'DDL', '已完成', '任务编号（自动）', '校验结果（自动）'];
  const empty = sheet.getLastRow() === 0;
  if (!empty) {
    const current = sheet.getRange(1, 1, 1, 6).getValues()[0];
    if (JSON.stringify(current.slice(0, 5)) === JSON.stringify(oldHeaders)) {
      sheet.insertColumnBefore(1);
    } else if (JSON.stringify(current) !== JSON.stringify(headers)) {
      throw new Error('“备忘录”工作表已存在且表头不同。请先将原工作表改名，再运行 setup。');
    }
  }
  if (sheet.getMaxRows() < 2) sheet.insertRowsAfter(1, 1);
  sheet.getRange(1, 1, 1, 6).setValues([headers]).setFontWeight('bold');
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  sheet.getRange(2, 3, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  const checkboxes = sheet.getRange(2, 4, sheet.getMaxRows() - 1, 1);
  if (empty) checkboxes.insertCheckboxes();
  else checkboxes.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sheet.setFrozenRows(1);
  [130, 300, 180, 90, 180, 320].forEach(function (width, index) { sheet.setColumnWidth(index + 1, width); });
  sheet.hideColumns(5);
  // 条件格式即时响应勾选；取消后自动恢复。标记只用于替换本程序的规则。
  const marker = PREFIX + 'completed';
  const rules = sheet.getConditionalFormatRules().filter(function (rule) {
    const condition = rule.getBooleanCondition();
    return !condition || !condition.getCriteriaValues().some(function (value) {
      return String(value).indexOf(marker) >= 0;
    });
  });
  const completedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(N("' + marker + '")=0,OR($D2=TRUE,$D2="已完成"))')
    .setBackground('#e0e0e0').setFontColor('#757575').setStrikethrough(true)
    .setRanges([sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns())]).build();
  sheet.setConditionalFormatRules([completedRule].concat(rules));
}

function memoEdited(e) {
  if (e && e.range && e.range.getSheet().getName() !== CONFIG.sheetName) return;
  if (e && e.range && (e.range.getColumn() > 4 || e.range.getLastRow() < 2)) return;
  measured_(function () { syncNow(); });
}

function memoChanged(e) {
  // EDIT 已由编辑触发器处理。删除行、插入行、排序等变化另行更新快照。
  if (e && (e.changeType === 'EDIT' || e.changeType === 'FORMAT')) return;
  measured_(function () { syncNow(); });
}

function syncNow() {
  withLock_(syncUnlocked_);
}

function syncUnlocked_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PREFIX + 'spreadsheet');
  if (!id) throw new Error('请先运行 setup。');
  const sheet = SpreadsheetApp.openById(id).getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error('找不到“备忘录”工作表，请恢复名称并运行 syncNow。');
  const state = readState_() || { tasks: [], sent: {}, nextSummary: nextSummary_(Date.now()) };
  const count = sheet.getLastRow() - 1;
  const tasks = [];
  const seen = {};
  if (count > 0) {
    const rows = sheet.getRange(2, 1, count, 6).getValues();
    const formulas = sheet.getRange(2, 2, count, 3).getFormulas();
    const metadata = rows.map(function (row, index) {
      const title = String(row[1]).trim();
      if (!title && !row[2]) return ['', ''];
      let taskId = String(row[4]).trim();
      if (!/^[0-9a-f-]{36}$/i.test(taskId) || seen[taskId]) taskId = Utilities.getUuid();
      seen[taskId] = true;
      if (formulas[index].some(function (x) { return x !== ''; })) return [taskId, '请直接填写事务、日期、完成状态，不使用公式'];
      if (!title) return [taskId, '缺少事务名称，暂不提醒'];
      const due = parseDue_(row[2]);
      if (due === null) return [taskId, 'DDL 无效，请填写完整日期和时间，例如 2026/11/20 18:00'];
      const done = row[3] === true || String(row[3]).trim() === '已完成';
      if (!done && row[3] !== false && String(row[3]).trim() !== '' && String(row[3]).trim() !== '未完成') {
        return [taskId, '请使用复选框，或填写“已完成 / 未完成”'];
      }
      if (!done) tasks.push({ id: taskId, title: title, due: due });
      return [taskId, done ? '已完成，不再提醒' : '有效'];
    });
    if (tasks.length > 500) throw new Error('本轻量版本最多支持 500 条未完成任务。');
    if (metadata.some(function (value, i) { return value[0] !== rows[i][4] || value[1] !== rows[i][5]; })) {
      sheet.getRange(2, 5, count, 2).setValues(metadata);
    }
    const times = rows.map(function (row, index) {
      const due = formulas[index][1] ? null : parseDue_(row[2]);
      return [due === null ? '' : formatShortDue_(due)];
    });
    if (times.some(function (value, i) { return value[0] !== rows[i][0]; })) {
      sheet.getRange(2, 1, count, 1).setValues(times);
    }
  }
  state.tasks = tasks;
  const now = Date.now();
  Object.keys(state.sent).forEach(function (key) {
    if (state.sent[key] < now - 7 * 24 * HOUR) delete state.sent[key];
  });
  state.syncedAt = now;
  saveState_(state);
}

function tick() {
  return measured_(tickCore_);
}

function tickCore_() {
  const props = PropertiesService.getScriptProperties();
  const wake = props.getProperty(PREFIX + 'wake');
  if (wake && Date.now() < Number(wake)) return;
  // 有同步正在执行时立即退出，交给下次检查，避免原来最多 30 秒的等锁开销。
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return;
  try {
    const state = readState_();
    if (!state) throw new Error('请先运行 setup。');
    const now = Date.now();
    if (state.nextSummary <= now) {
      const tasks = summaryTasks_(state.tasks, now);
      const lines = tasks.length ? tasks.map(taskLine_).join('\n\n') : '未来 48 小时暂无未完成事项。';
      send_('【备忘录汇总】' + formatTime_(state.nextSummary),
        '计划发送时间：' + formatTime_(state.nextSummary) + '\n当前检查时间：' + formatTime_(now) + '\n\n' + lines);
      state.nextSummary = nextSummary_(now);
      saveState_(state);
    }
    const dueTasks = singleTasks_(state, now);
    dueTasks.forEach(function (task) {
      send_('【DDL 前5小时提醒】' + task.title.replace(/[\r\n]/g, ' ').slice(0, 150),
        taskLine_(task) + '\n\n计划提醒时间：' + formatTime_(task.due - CONFIG.advanceHours * HOUR)
        + '\n实际检查时间：' + formatTime_(now));
      state.sent[taskKey_(task)] = task.due;
      saveState_(state);
    });
    // 此函数只读已保存的数据，不调用 SpreadsheetApp、不读取表格。
    props.setProperty(PREFIX + 'wake', String(nextWake_(state, now)));
  } finally { lock.releaseLock(); }
}

function summaryTasks_(tasks, now) {
  return tasks.filter(function (task) {
    return task.due > now && task.due - now < CONFIG.windowHours * HOUR;
  }).sort(function (a, b) { return a.due - b.due; });
}

function singleTasks_(state, now) {
  return state.tasks.filter(function (task) {
    return task.due > now && task.due - CONFIG.advanceHours * HOUR <= now && !state.sent[taskKey_(task)];
  }).sort(function (a, b) { return a.due - b.due; });
}

function nextWake_(state, now) {
  let result = state.nextSummary;
  state.tasks.forEach(function (task) {
    if (task.due > now && !state.sent[taskKey_(task)]) {
      result = Math.min(result, Math.max(now, task.due - CONFIG.advanceHours * HOUR));
    }
  });
  return result;
}

function nextSummary_(now) {
  const china = new Date(now + 8 * HOUR);
  const midnight = Date.UTC(china.getUTCFullYear(), china.getUTCMonth(), china.getUTCDate()) - 8 * HOUR;
  const candidates = CONFIG.summaryTimes.map(function (value) {
    const parts = value.split(':').map(Number);
    return midnight + parts[0] * HOUR + parts[1] * 60000;
  });
  return candidates.find(function (time) { return time > now; }) || candidates[0] + 24 * HOUR;
}

function parseDue_(value) {
  if (value instanceof Date) return Number.isFinite(value.getTime()) ? value.getTime() : null;
  const match = /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(String(value).trim());
  if (!match) return null;
  const parts = match.slice(1).map(function (x) { return x === undefined ? 0 : Number(x); });
  const ms = Date.UTC(parts[0], parts[1] - 1, parts[2], parts[3], parts[4], parts[5]);
  const date = new Date(ms);
  if (date.getUTCFullYear() !== parts[0] || date.getUTCMonth() !== parts[1] - 1 || date.getUTCDate() !== parts[2]
    || date.getUTCHours() !== parts[3] || date.getUTCMinutes() !== parts[4] || date.getUTCSeconds() !== parts[5]) return null;
  return ms - 8 * HOUR;
}

function formatShortDue_(ms) {
  const china = new Date(ms + 8 * HOUR);
  return '周' + '日一二三四五六'[china.getUTCDay()] + ' '
    + String(china.getUTCHours()).padStart(2, '0') + ':' + String(china.getUTCMinutes()).padStart(2, '0');
}

function taskKey_(task) { return task.id + ':' + task.due; }
function formatTime_(ms) { return Utilities.formatDate(new Date(ms), CONFIG.timezone, 'yyyy/MM/dd HH:mm'); }
function taskLine_(task) { return '事务：' + task.title + '\nDDL：' + formatTime_(task.due) + '（北京时间）'; }

function send_(subject, body) {
  validateConfig_();
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('今日邮件额度已耗尽，请查看 Apps Script 执行记录。');
  MailApp.sendEmail({ to: CONFIG.recipient, subject: subject, body: body, name: '个人备忘录' });
}

function sendTestEmail() {
  send_('【备忘录】配置测试', '这是一封手动测试邮件。收到说明发信授权正常。\n北京时间：' + formatTime_(Date.now()));
}

function validateConfig_() {
  if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(CONFIG.recipient)) throw new Error('请先修改 CONFIG.recipient，填写一个有效收件邮箱。');
  if ([5, 10].indexOf(CONFIG.checkMinutes) < 0) throw new Error('checkMinutes 请填写 5 或 10。');
}

function withLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return callback(); } finally { lock.releaseLock(); }
}

// 用持久化 Properties 保存任务快照。分块避免单条属性的 9KB 限制。
// 先写新版本，再切换指针；中途写入失败仍保留原版本。
function readState_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PREFIX + 'manifest');
  if (!raw) return null;
  const manifest = JSON.parse(raw);
  let json = '';
  for (let i = 0; i < manifest.count; i++) {
    const part = props.getProperty(PREFIX + 'state_' + manifest.version + '_' + i);
    if (part === null) throw new Error('保存的任务数据不完整，请运行 syncNow 或检查执行记录。');
    json += part;
  }
  return JSON.parse(json);
}

function saveState_(state) {
  const props = PropertiesService.getScriptProperties();
  const raw = JSON.stringify(state);
  if (Utilities.newBlob(raw).getBytes().length > 150000) throw new Error('任务数据过大，请减少任务或缩短事务文本。');
  const version = Utilities.getUuid();
  const values = {};
  let count = 0;
  // Array.from 按 Unicode 字符分割，避免拆开 emoji 的代理对。
  const chars = Array.from(raw);
  for (let i = 0; i < chars.length; i += 1000) {
    values[PREFIX + 'state_' + version + '_' + count++] = chars.slice(i, i + 1000).join('');
  }
  props.setProperties(values, false);
  props.setProperty(PREFIX + 'manifest', JSON.stringify({ version: version, count: count }));
  // 将 wake 提前到当前时间，使下个 tick 重新判断更新后的计划。
  props.setProperty(PREFIX + 'wake', '0');
  const all = props.getProperties();
  Object.keys(all).forEach(function (key) {
    if (key.indexOf(PREFIX + 'state_') === 0 && !Object.prototype.hasOwnProperty.call(values, key)) props.deleteProperty(key);
  });
}

function stopReminders() {
  ['memoEdited', 'memoChanged', 'tick'].forEach(function (handler) {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
      if (trigger.getHandlerFunction() === handler) ScriptApp.deleteTrigger(trigger);
    });
  });
  console.log('已移除本程序的三个触发器，表格与任务数据保留。');
}

// 记录本项目触发器的主体耗时（不含这段统计自身的耗时），不代表 Google 官方剩余额度。
// 按整小时保留 24～25 小时；手动运行 tick 也计入；锁冲突时跳过记录并写日志。
function measured_(callback) {
  const started = Date.now();
  let failed = false;
  try { return callback(); } catch (error) { failed = true; throw error; }
  finally {
    const elapsed = Math.max(0, Date.now() - started);
    const lock = LockService.getUserLock();
    let acquired = false;
    try {
      acquired = lock.tryLock(0);
      if (!acquired) console.warn('耗时统计锁冲突，本次未记录：' + elapsed + ' ms');
      else {
        const props = PropertiesService.getScriptProperties();
        const hour = Math.floor(Date.now() / HOUR);
        const buckets = JSON.parse(props.getProperty(PREFIX + 'runtime') || '[]')
          .filter(function (b) { return b.hour >= hour - 24; });
        let bucket = buckets.find(function (b) { return b.hour === hour; });
        if (!bucket) { bucket = { hour: hour, count: 0, ms: 0, max: 0, failed: 0 }; buckets.push(bucket); }
        bucket.count++;
        bucket.ms += elapsed;
        bucket.max = Math.max(bucket.max, elapsed);
        if (failed) bucket.failed++;
        props.setProperty(PREFIX + 'runtime', JSON.stringify(buckets));
      }
    } catch (error) { console.warn('无法保存耗时统计：' + error.message); }
    finally { if (acquired) lock.releaseLock(); }
  }
}

function checkRuntime() {
  const props = PropertiesService.getScriptProperties();
  const hour = Math.floor(Date.now() / HOUR);
  const buckets = JSON.parse(props.getProperty(PREFIX + 'runtime') || '[]')
    .filter(function (b) { return b.hour >= hour - 24; });
  const total = buckets.reduce(function (result, b) {
    result.count += b.count; result.ms += b.ms; result.max = Math.max(result.max, b.max); result.failed += b.failed;
    return result;
  }, { count: 0, ms: 0, max: 0, failed: 0 });
  console.log('最近约24～25小时，本项目已记录检查/同步：' + total.count + ' 次；主体耗时合计：'
    + (total.ms / 60000).toFixed(2) + ' 分钟；平均：' + (total.count ? total.ms / total.count / 1000 : 0).toFixed(2)
    + ' 秒；最长：' + (total.max / 1000).toFixed(2) + ' 秒；失败：' + total.failed + ' 次。');
  console.log('此为程序内估算：不含统计自身、其他项目、setup等未记录操作；如有统计失败还会缺样。官方执行记录更适合核对实际运行时长。');
  console.log('邮件剩余收件人人次：' + MailApp.getRemainingDailyQuota());
}
