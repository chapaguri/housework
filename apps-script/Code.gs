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
 * 데이터는 '앱데이터' 탭에 key/value 로 쌓인다(없으면 자동 생성, 항상 보이게 유지).
 * 저장 형태:
 *   d|2026-08-13|k1  = 1     (날짜|항목ID = 1 완료, 2 미이행)
 *   x|2026-08-16     = 1     (제외일 — 여행 등)
 *   m|2026-08|m3     = 1     (매달 항목)
 *   start            = 2026-08-13
 */

var TOKEN = 't36Bh5n_fJkFlqWam7vMIroC';   // 앱과 반드시 동일해야 함
var SHEET_NAME = '앱데이터';
var TASK_SHEET = '항목';   // 항목 정의 탭 (없으면 앱의 기본 목록으로 동작)

/**
 * 데이터를 저장할 스프레드시트 ID.
 * 시트 주소에서 /d/ 와 /edit 사이의 긴 문자열이다.
 *   https://docs.google.com/spreadsheets/d/[여기가 ID]/edit
 * 시트에 붙여넣은 스크립트(확장 프로그램 → Apps Script)라면 ''(빈 값)으로 둬도 된다.
 */
var SHEET_ID = '1zRJjT09Nga9pDOFIVjj56UGlZD5zimRzyW8SWIk9gs4';

var VERSION = 'v5';   // 어떤 코드가 실제로 배포됐는지 확인하는 표식

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
  if (!ops || !ops.length)
    return json({ ok: true, ver: VERSION, data: readAll(), tasks: readTasks() });
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
  }
  if (sh.isSheetHidden()) sh.showSheet();   // 탭을 항상 보이게 둔다
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


/* ══ 항목 정의 ══
   '항목' 탭에서 체크리스트 구성을 읽어온다. 탭이 없거나 비어 있으면 null 을
   돌려주고, 앱은 코드에 내장된 기본 목록으로 동작한다.

   열 구성
     A ID          바꾸지 말 것. 과거 체크 기록이 이 값으로 묶여 있다
     B 구분        화장실 / 주방 / 거실 / 세탁 / 청소
     C 항목명      자유롭게 변경 가능 (표시용)
     D 부가설명    작은 글씨로 함께 보임
     E 주기        매일 / 주말 / 매달 / 주3회 / 화,목,일
     F 목표        '주N회' 일 때 횟수
     G 보너스      Y 면 추가 청소 보너스 대상 (매달 항목에만)
     H 적용시작일  2026-08-18 형식. 이 날짜부터 항목이 나타난다 (비우면 전체 기간)
     I 사용        N 이면 앞으로 표시하지 않는다. 지난 기록은 보존된다
     J 짧은이름    정산 화면의 주간 목표 칩에 쓰는 축약 이름
*/
var TASK_COLS = 10;

function readTasks() {
  var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  var sh = ss.getSheetByName(TASK_SHEET);
  if (!sh) return null;

  var last = sh.getLastRow();
  if (last < 2) return null;

  var vals = sh.getRange(2, 1, last - 1, TASK_COLS).getDisplayValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var id = String(r[0]).trim();
    if (!id) continue;
    out.push({
      id:    id,
      cat:   String(r[1]).trim(),
      name:  String(r[2]).trim(),
      sub:   String(r[3]).trim(),
      freq:  String(r[4]).trim(),
      target: String(r[5]).trim(),
      bonus: String(r[6]).trim(),
      from:  String(r[7]).trim(),
      use:   String(r[8]).trim(),
      short: String(r[9]).trim()
    });
  }
  return out.length ? out : null;
}

/* 현재 항목을 '항목' 탭에 처음 채워넣는다. 편집기에서 한 번만 실행할 것.
   이미 내용이 있으면 아무것도 하지 않는다. */
function seedTasks() {
  var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TASK_SHEET);
  if (!sh) sh = ss.insertSheet(TASK_SHEET);
  if (sh.getLastRow() > 1) {
    throw new Error("'" + TASK_SHEET + "' 탭에 이미 내용이 있습니다. 비운 뒤 다시 실행하세요.");
  }

  var header = ['ID', '구분', '항목명', '부가설명', '주기', '목표', '보너스', '적용시작일', '사용', '짧은이름'];
  var rows = SEED_TASKS;

  sh.getRange(1, 1, 1, TASK_COLS).setValues([header])
    .setFontWeight('bold').setBackground('#eceff1');
  sh.getRange(2, 1, rows.length, TASK_COLS).setNumberFormat('@').setValues(rows);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 60);
  sh.setColumnWidth(3, 200);
  sh.setColumnWidth(4, 130);
  return rows.length + '개 항목을 채웠습니다.';
}

var SEED_TASKS = [
  ['b1', '화장실', '변기 커버 이물질 확인', '', '매일', '', '', '', 'Y', ''],
  ['b2', '화장실', '소모품 채우기', '', '매일', '', '', '', 'Y', ''],
  ['b3', '화장실', '변기 청소 · 하수구 청소', '', '주말', '', '', '', 'Y', ''],
  ['k1', '주방', '당일 설거지', '', '매일', '', '', '', 'Y', ''],
  ['k2', '주방', '전일 설거지 · 식기 정리', '', '매일', '', '', '', 'Y', ''],
  ['k5', '주방', '얼음 · 물 채워넣기', '', '매일', '', '', '', 'Y', ''],
  ['k3', '주방', '싱크대 및 하수구 청소', '', '주말', '', '', '', 'Y', ''],
  ['k4', '주방', '가스레인지 청소', '', '주말', '', '', '', 'Y', ''],
  ['l1', '거실', '쓰레기통 · 분리수거통 비우기', '', '매일', '', '', '', 'Y', ''],
  ['l2', '거실', '택배 정리 · 박스 버리기', '', '매일', '', '', '', 'Y', ''],
  ['l3', '거실', '테이블 및 쓰레기 정리', '', '매일', '', '', '', 'Y', ''],
  ['l4', '거실', '누누 물주기', '', '주말', '', '', '', 'Y', ''],
  ['w1', '세탁', '건조된 세탁물 정리', 'D+1일까지', '매일', '', '', '', 'Y', ''],
  ['w2', '세탁', '종량제봉투 내다놓기', '화 · 목 · 일', '화,목,일', '', '', '', 'Y', ''],
  ['w6', '세탁', '베개 커버 교체', '', '주말', '', '', '', 'Y', ''],
  ['w3', '세탁', '수건 빨래', '', '주1회', '1', '', '', 'Y', '수건'],
  ['w4', '세탁', '옷 빨래', '', '주1회', '1', '', '', 'Y', '옷'],
  ['w5', '세탁', '팬티 빨래', '', '주1회', '1', '', '', 'Y', '팬티'],
  ['c1', '청소', '로봇 청소기', '', '주3회', '3', '', '', 'Y', '로봇'],
  ['c2', '청소', '진공 청소기 · 스팀 청소기', '', '주말', '', '', '', 'Y', ''],
  ['m1', '화장실', '화장실 전체 청소', '', '매달', '', '', '', 'Y', ''],
  ['m2', '주방', '냉장고 · 음쓰 정리', '', '매달', '', '', '', 'Y', ''],
  ['m3', '세탁', '이불 세탁', '', '매달', '', 'Y', '', 'Y', ''],
  ['m4', '청소', '집안 전체 먼지 제거', '', '매달', '', '', '', 'Y', ''],
  ['m5', '청소', '창틀 청소', '', '매달', '', 'Y', '', 'Y', ''],
  ['m6', '청소', '공청기 필터 청소', '', '매달', '', 'Y', '', 'Y', '']
];
