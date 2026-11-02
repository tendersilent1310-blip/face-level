// 얼굴 수평계 → 구글 시트 받기 코드
// 사람(별명)+폰마다 탭 하나. 보낼 때마다 그 탭을 최신 기록으로 통째로 바꿔요(중복 없음).
// "받은기록" 탭에는 언제 누가 몇 개를 보냈는지 한 줄씩 쌓여요.
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.v !== 1 || typeof d.csv !== 'string') return out({ ok: false, err: 'bad' });
    if (d.csv.length > 5000000) return out({ ok: false, err: 'too big' });
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tab = String(d.tab || '무명').replace(/[\[\]\*\?\/\\:']/g, '').slice(0, 90) || '무명';
    var sh = ss.getSheetByName(tab) || ss.insertSheet(tab);
    var rows = Utilities.parseCsv(d.csv.replace(/^﻿/, ''));
    var w = 1; rows.forEach(function (r) { if (r.length > w) w = r.length; });
    rows = rows.map(function (r) {
      var o = r.map(function (v) {
        if (/^-?\d+(\.\d+)?$/.test(v)) return Number(v);       // 숫자는 숫자로
        if (/^[=+@]/.test(v)) return "'" + v;                    // 수식 실행 막기
        return v;
      });
      while (o.length < w) o.push('');
      return o;
    });
    sh.clearContents();
    if (rows.length) sh.getRange(1, 1, rows.length, w).setValues(rows);
    var log = ss.getSheetByName('받은기록') || ss.insertSheet('받은기록', 0);
    if (log.getLastRow() === 0) log.appendRow(['받은 시간', '탭', '기록 수']);
    log.appendRow([new Date(), tab, d.n || rows.length - 1]);
    lock.releaseLock();
    return out({ ok: true, rows: rows.length });
  } catch (err) {
    return out({ ok: false, err: String(err) });
  }
}
function doGet() { return out({ ok: true, hello: '얼굴 수평계 시트 연결 OK' }); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
