// 얼굴 수평계 → 구글 시트 받기 코드 (v2: 로그인 + 기록 되찾기)
// - 사람(계정 이름)마다 탭 하나. 보낼 때마다 그 탭을 최신 기록으로 통째로 바꿔요(중복 없음).
// - "사용자" 탭: 이름과 암호화된 비밀번호(원래 비밀번호는 저장 안 함).
// - "받은기록" 탭: 언제 누가 몇 개를 보냈는지.
// - 비밀번호를 5번 틀리면 그 이름은 10분 동안 잠겨요.

var USERS = '사용자', LOG = '받은기록', RESERVED = [USERS, LOG];

// ▼ 분석용 읽기 문 비밀번호: 따옴표 안을 형만 아는 비밀번호로 바꾸세요 (이 코드는 시트 안에만 있고 공개되지 않아요)
var READ_KEY = '여기에-비밀번호';

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.v !== 1) return out({ ok: false, err: 'bad' });
    lock.waitLock(20000);
    if (d.act === 'register') return out(register_(d));
    if (d.act === 'login') return out(login_(d));
    if (d.act === 'save') return out(save_(d));
    return out(legacySave_(d));
  } catch (err) {
    return out({ ok: false, err: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// ---------- 계정 ----------
function cleanName_(n) { return String(n || '').replace(/[\[\]\*\?\/\\:']/g, '').trim().slice(0, 20); }
function okHash_(h) { return /^[0-9a-f]{64}$/.test(String(h || '')); }
function users_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(USERS);
  if (!sh) { sh = ss.insertSheet(USERS); sh.appendRow(['이름', '비밀번호(암호화)', '가입', '마지막 접속']); }
  return sh;
}
function findUser_(name) {
  var sh = users_(), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === name) return { row: i + 1, hash: String(v[i][1]) };
  return null;
}
function locked_(name) { return Number(CacheService.getScriptCache().get('fail_' + name) || 0) >= 5; }
function fail_(name) { var c = CacheService.getScriptCache(), k = 'fail_' + name; c.put(k, String(Number(c.get(k) || 0) + 1), 600); }
function check_(d) {
  var name = cleanName_(d.name);
  if (!name || RESERVED.indexOf(name) >= 0 || !okHash_(d.pin)) return { err: 'bad' };
  var u = findUser_(name);
  if (!u) return { err: 'nouser' };
  if (locked_(name)) return { err: 'locked' };
  if (u.hash !== d.pin) { fail_(name); return { err: 'wrong' }; }
  users_().getRange(u.row, 4).setValue(new Date());
  return { name: name };
}
function register_(d) {
  var name = cleanName_(d.name);
  if (!name || RESERVED.indexOf(name) >= 0 || !okHash_(d.pin)) return { ok: false, err: 'bad' };
  if (findUser_(name)) return { ok: false, err: 'exists' };
  users_().appendRow([name, d.pin, new Date(), new Date()]);
  return { ok: true, name: name, csv: tabCsv_(name) }; // 같은 이름의 탭이 이미 있으면(예: 옛 탭 이름을 바꿔 둔 경우) 그 기록을 돌려줌
}
function login_(d) {
  var c = check_(d); if (c.err) return { ok: false, err: c.err };
  return { ok: true, name: c.name, csv: tabCsv_(c.name) };
}
function save_(d) {
  var c = check_(d); if (c.err) return { ok: false, err: c.err };
  if (typeof d.csv !== 'string' || d.csv.length > 5000000) return { ok: false, err: 'bad csv' };
  var n = writeTab_(c.name, d.csv); logRow_(c.name, d.n || n - 1);
  return { ok: true, rows: n };
}
// 옛 버전 앱(계정 없음): "이름_폰번호" 탭에 저장. 계정 이름 탭은 덮어쓰지 못하게 막음
function legacySave_(d) {
  if (typeof d.csv !== 'string' || d.csv.length > 5000000) return { ok: false, err: 'bad' };
  var tab = cleanName_(String(d.tab || '무명').slice(0, 90)) || '무명';
  if (RESERVED.indexOf(tab) >= 0 || findUser_(tab)) return { ok: false, err: 'reserved' };
  var n = writeTab_(tab, d.csv); logRow_(tab, d.n || n - 1);
  return { ok: true, rows: n };
}

// ---------- 탭 쓰기/읽기 ----------
function writeTab_(tab, csv) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(tab) || ss.insertSheet(tab);
  var rows = Utilities.parseCsv(csv.replace(/^﻿/, ''));
  var w = 1; rows.forEach(function (r) { if (r.length > w) w = r.length; });
  rows = rows.map(function (r) {
    var o = r.map(function (v) {
      if (/^-?\d+(\.\d+)?$/.test(v)) return Number(v);   // 숫자는 숫자로
      if (/^[=+@]/.test(v)) return "'" + v;                // 수식 실행 막기
      return v;
    });
    while (o.length < w) o.push('');
    return o;
  });
  sh.clearContents();
  if (rows.length) {
    var rg = sh.getRange(1, 1, rows.length, w);
    rg.setNumberFormat('@');                              // 날짜처럼 생긴 글자가 날짜로 바뀌지 않게
    rg.setValues(rows.map(function (r) { return r.map(function (v) { return typeof v === 'number' ? String(v) : v; }); }));
  }
  return rows.length;
}
function tabCsv_(tab) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(tab);
  if (!sh || sh.getLastRow() === 0) return '';
  var tz = ss.getSpreadsheetTimeZone();
  return sh.getDataRange().getValues().map(function (r) {
    return r.map(function (c) {
      if (c instanceof Date) c = Utilities.formatDate(c, tz, 'yyyy-MM-dd');
      c = String(c);
      return /[",\n]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c;
    }).join(',');
  }).join('\n');
}
function logRow_(tab, n) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), log = ss.getSheetByName(LOG) || ss.insertSheet(LOG, 0);
  if (log.getLastRow() === 0) log.appendRow(['받은 시간', '탭', '기록 수']);
  log.appendRow([new Date(), tab, n]);
}

// ---------- 분석용 읽기 문 (비밀번호가 맞을 때만, "사용자" 탭은 빼고) ----------
function doGet(e) {
  var key = e && e.parameter && e.parameter.key;
  if (!key || READ_KEY === '여기에-비밀번호' || key !== READ_KEY) return out({ ok: true, hello: '얼굴 수평계 시트 연결 OK', v: 2 });
  var ss = SpreadsheetApp.getActiveSpreadsheet(), tabs = {};
  ss.getSheets().forEach(function (sh) { var n = sh.getName(); if (n !== USERS) tabs[n] = tabCsv_(n); });
  return out({ ok: true, tabs: tabs, t: new Date() });
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
