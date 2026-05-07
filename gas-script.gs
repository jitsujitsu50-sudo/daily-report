const SPREADSHEET_ID = '1GoQchevlrTbg8ZUm47BxpHcqAERYmS-UYormiLD4tdg';
const TZ = Session.getScriptTimeZone();

function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'syncDay')       return jsonOut(handleSyncDay(p));
    if (p.action === 'syncSchedules') return jsonOut(handleSyncSchedules(p));
    if (p.action === 'resetSheet')    return jsonOut(handleResetSheet());
    if (p.action === 'syncTasks')     return jsonOut(handleSyncTasks(p));
    if (p.month) return jsonOut(getMonthEvents(p.month));
    if (p.date)  return jsonOut(getDayEvents(p.date));
    return jsonOut({error: 'invalid request'});
  } catch (err) {
    return jsonOut({error: err.toString()});
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'bulkResults')   return jsonOut(handleBulkResults(body));
    if (body.action === 'bulkSchedules') return jsonOut(handleBulkSchedules(body));
    if (body.action === 'syncTasks')     return jsonOut(writeTasks(body.tasks || []));
    return jsonOut({error: 'invalid post action'});
  } catch (err) {
    return jsonOut({error: err.toString()});
  }
}

// 一括書き込み（全日付を1リクエストで） ────

function bulkWriteDay(sheet, dataByDate) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
    const now = nowStr();
    const rows = [];
    Object.keys(dataByDate).sort().forEach(function(date) {
      (dataByDate[date] || []).forEach(function(en) {
        if ((en.t && en.t.length) || (en.s && en.s.length)) {
          rows.push([date, en.s || '', en.e || '', en.t || '', en.n || '', now]);
        }
      });
    });
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    return {ok: true, rows: rows.length};
  } finally {
    lock.releaseLock();
  }
}

function handleBulkResults(body)   { return bulkWriteDay(getResultsSheet(),   body.data || {}); }
function handleBulkSchedules(body) { return bulkWriteDay(getSchedulesSheet(), body.data || {}); }

function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 日付セルを文字列に正規化（Sheetsの自動Date変換対策）
function normalizeDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  return String(v).trim();
}

function nowStr() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

// カレンダー取得 ────────────────────────────

function getDayEvents(dateStr) {
  const p = dateStr.split('-');
  const start = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]), 0, 0, 0);
  const end   = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]), 23, 59, 59);
  return CalendarApp.getDefaultCalendar().getEvents(start, end).map(formatEvent);
}

function getMonthEvents(monthStr) {
  const p = monthStr.split('-');
  const start = new Date(parseInt(p[0]), parseInt(p[1]) - 1, 1, 0, 0, 0);
  const end   = new Date(parseInt(p[0]), parseInt(p[1]),     0, 23, 59, 59);
  return CalendarApp.getDefaultCalendar().getEvents(start, end).map(formatEvent);
}

function formatEvent(ev) {
  return {
    id:        ev.getId(),
    summary:   ev.getTitle(),
    startTime: Utilities.formatDate(ev.getStartTime(), TZ, 'HH:mm'),
    endTime:   Utilities.formatDate(ev.getEndTime(),   TZ, 'HH:mm'),
    date:      Utilities.formatDate(ev.getStartTime(), TZ, 'yyyy-MM-dd'),
    allDay:    ev.isAllDayEvent()
  };
}

// スプレッドシート取得 ────────────────────────

function ssOpen() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

const DAY_HEADERS = ['日付', '開始', '終了', '内容', 'メモ', '更新日時'];

function ensureDayHeaders(sheet) {
  // ヘッダーが古い（5列）場合、6列目に更新日時を追加
  try {
    if (sheet.getRange(1, 6).getValue() !== '更新日時') {
      sheet.getRange(1, 6).setValue('更新日時');
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }
  } catch (e) {}
  // 注: 旧バージョンは setNumberFormat('@') を呼んでいたが、表形式（型付きカラム）と衝突するため削除。
  // 日付の型変換は normalizeDate() 側で吸収するため不要。
}

function getResultsSheet() {
  const ss = ssOpen();
  let sheet = ss.getSheetByName('実績');
  if (!sheet) {
    const sheets = ss.getSheets();
    if (sheets.length > 0 && sheets[0].getLastRow() > 0
        && sheets[0].getRange(1, 1).getValue() === '日付') {
      sheets[0].setName('実績');
      sheet = sheets[0];
    } else {
      sheet = ss.insertSheet('実績');
      sheet.appendRow(DAY_HEADERS);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }
  }
  ensureDayHeaders(sheet);
  return sheet;
}

function getSchedulesSheet() {
  const ss = ssOpen();
  let sheet = ss.getSheetByName('スケジュール');
  if (!sheet) {
    sheet = ss.insertSheet('スケジュール');
    sheet.appendRow(DAY_HEADERS);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  ensureDayHeaders(sheet);
  return sheet;
}

function getTasksSheet() {
  const ss = ssOpen();
  let sheet = ss.getSheetByName('タスク');
  if (!sheet) {
    sheet = ss.insertSheet('タスク');
    sheet.appendRow(['ID', 'タイトル', '優先度', '期日', 'カテゴリ', '対象月', '見積分', '実績分', '状態', 'メモ', '更新日時']);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
  } else if (sheet.getLastRow() > 0 && sheet.getRange(1, 8).getValue() !== '実績分') {
    sheet.clearContents();
    sheet.appendRow(['ID', 'タイトル', '優先度', '期日', 'カテゴリ', '対象月', '見積分', '実績分', '状態', 'メモ', '更新日時']);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
  }
  return sheet;
}

// 日付ごと書き込み（実績/スケジュール共通） ──

function syncDayToSheet(sheet, date, entries) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const lastRow = sheet.getLastRow();
    const allData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 6).getValues() : [];
    const filtered = [];
    allData.forEach(function(row) {
      const rd = normalizeDate(row[0]);
      if (rd !== date) {
        row[0] = rd;
        // 既存行は更新日時を保持（空なら空のまま）
        if (row.length < 6) row[5] = '';
        filtered.push(row);
      }
    });
    const now = nowStr();
    entries.forEach(function(en) {
      filtered.push([date, en.s || '', en.e || '', en.t || '', en.n || '', now]);
    });
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
    if (filtered.length > 0) sheet.getRange(2, 1, filtered.length, 6).setValues(filtered);
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

function handleSyncDay(params) {
  const entries = JSON.parse(params.entries || '[]');
  return syncDayToSheet(getResultsSheet(), params.date, entries);
}

function handleSyncSchedules(params) {
  const entries = JSON.parse(params.entries || '[]');
  return syncDayToSheet(getSchedulesSheet(), params.date, entries);
}

function handleResetSheet() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getResultsSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

// タスク同期 ──────────────────────────────

function handleSyncTasks(params) {
  return writeTasks(JSON.parse(params.tasks || '[]'));
}

function writeTasks(tasks) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getTasksSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 11).clearContent();
    if (tasks.length > 0) {
      const now = nowStr();
      const rows = tasks.map(function(t) {
        return [
          t.id || '',
          t.title || '',
          t.priority || '',
          t.dueDate || '',
          t.category || '',
          t.month || '',
          t.estMin || '',
          t.actualMin || '',
          t.done ? '完了' : '未完了',
          t.memo || '',
          now
        ];
      });
      sheet.getRange(2, 1, rows.length, 11).setValues(rows);
    }
    return {ok: true, count: tasks.length};
  } finally {
    lock.releaseLock();
  }
}
