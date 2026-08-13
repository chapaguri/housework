/**
 * 집안일 체크리스트 — 구글 시트를 저장소로 쓰는 동기화 API
 *
 * 설치 방법
 *  1) 데이터를 저장할 구글 시트를 만든다 (빈 시트로 충분)
 *  2) 시트 주소에서 ID를 복사해 아래 SHEET_ID 에 붙여넣는다
 *     (시트 안에서 확장 프로그램 → Apps Script 로 만든 경우엔 빈 값이어도 된다)
 *  3) 기본 코드(Code.gs)를 지우고 이 파일 전체를 붙여넣는다
 *  4) 편집기에서 readAll 함수를 한 번 실행해 권한을 승인한다
 *  5) 배포 → 새 배포 → 유형: 웹 앱
 *       - 실행 계정: 나
 *       - 액세스 권한: 모든 사용자   ← 배우자가 구글 로그인 없이 쓰려면 필수
 *  6) 배포 후 나오는 '웹 앱 URL'을 복사해서 앱 설정에 붙여넣는다
 *
 * 데이터는 '앱데이터' 탭에 key/value 로 쌓인다(숨김 탭, 자동 생성).
 * 저장 형태:
 *   d|2026-08-13|k1  = 1     (날짜|항목ID = 1 완료, 2 미이행)
 *   x|2026-08-16     = 1     (제외일 — 여행 등)
 *   m|2026-08|m3     = 1     (매달 항목)
 *   start            = 2026-08-13
 */

var TOKEN = 't36Bh5n_fJkFlqWam7vMIroC';   // 앱과 반드시 동일해야 함
var SHEET_NAME = '앱데이터';

/**
 * 데이터를 저장할 스프레드시트 ID.
 * 시트 주소에서 /d/ 와 /edit 사이의 긴 문자열이다.
 *   https://docs.google.com/spreadsheets/d/[여기가 ID]/edit
 * 시트에 붙여넣은 스크립트(확장 프로그램 → Apps Script)라면 ''(빈 값)으로 둬도 된다.
 */
var SHEET_ID = '1zRJjT09Nga9pDOFIVjj56UGlZD5zimRzyW8SWIk9gs4';

var VERSION = 'v3';   // 어떤 코드가 실제로 배포됐는지 확인하는 표식

/**
 * 읽기와 쓰기를 모두 GET 으로 처리한다.
 * Apps Script 의 POST 응답은 script.googleusercontent.com 으로 리다이렉트되는데
 * 이 응답 본문을 브라우저에서 안정적으로 읽지 못하는 경우가 있어(쓰기는 되지만
 * 결과를 못 받아 무한 재전송이 발생) GET 한 가지로 통일했다.
 *   읽기: ?token=...
 *   쓰기: ?token=...&ops=[{"k":"...","v":"1"},{"k":"...","v":null}]
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (String(p.token || '') !== TOKEN) return json({ error: 'unauthorized', ver: VERSION });

  var ops = null;
  if (p.ops) {
    try { ops = JSON.parse(p.ops); } catch (err) { return json({ error: 'bad ops', ver: VERSION }); }
  }
  if (!ops || !ops.length) return json({ ok: true, ver: VERSION, data: readAll() });
  return json({ ok: true, ver: VERSION, data: applyOps_(ops) });
}

/* POST 도 같은 동작으로 받아둔다 (호환용) */
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  if (String(body.token || '') !== TOKEN) return json({ error: 'unauthorized', ver: VERSION });
  return json({ ok: true, ver: VERSION, data: applyOps_(body.ops || []) });
}

function applyOps_(ops) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { throw new Error('busy'); }
  try {
    var data = readAll();
    for (var i = 0; i < ops.length; i++) {
      var op = ops[i];
      if (!op || !op.k) continue;
      if (op.v === null || op.v === undefined || op.v === '') delete data[op.k];
      else data[op.k] = String(op.v);
    }
    writeAll(data);
    return data;
  } finally {
    lock.releaseLock();
  }
}

/* ── 내부 ── */

function getSheet_() {
  var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('스프레드시트를 찾을 수 없습니다. SHEET_ID 를 설정하세요.');
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 2).setValues([['key', 'value']]).setFontWeight('bold');
    sh.hideSheet();
  }
  return sh;
}

function readAll() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return {};
  var vals = sh.getRange(2, 1, last - 1, 2).getDisplayValues();
  var out = {};
  for (var i = 0; i < vals.length; i++) {
    var k = String(vals[i][0]).trim();
    if (k) out[k] = String(vals[i][1]);
  }
  return out;
}

function writeAll(data) {
  var sh = getSheet_();
  var keys = Object.keys(data).sort();
  var maxRows = sh.getMaxRows();
  if (maxRows > 1) sh.getRange(2, 1, maxRows - 1, 2).clearContent();
  if (!keys.length) return;
  var rows = keys.map(function (k) { return [k, String(data[k])]; });
  // 텍스트 서식 고정 — '2026-08-13' 같은 값이 날짜로 자동 변환되는 것을 막는다
  sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
