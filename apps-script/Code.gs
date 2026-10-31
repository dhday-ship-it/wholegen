/**
 * 홀젠캠프 등록상태 — 구글폼 응답 시트에 붙이는 Apps Script
 *
 * 설치
 *  1. 구글폼 응답 시트 열기 → 확장 프로그램 → Apps Script → 이 파일 내용을 붙여넣고 저장
 *  2. 프로젝트 설정(톱니바퀴) → 스크립트 속성 → ADMIN_PASSWORD = 스태프 비밀번호
 *  3. 배포 → 새 배포 → 유형: 웹 앱 · 실행: 나 · 액세스 권한: 모든 사용자 → 배포
 *  4. 나온 웹 앱 URL(…/exec)을 홈페이지 reg-config.js 의 REG_API 에 넣기
 *  코드를 고친 뒤에는 배포 → 배포 관리 → 연필 → 버전: 새 버전 → 배포 (주소는 그대로)
 *  (선택) 스크립트 속성 CALENDAR_ID = 어드민 홈에 보여줄 구글 캘린더 ID (없으면 기본 캘린더)
 *
 * 시트 맨 오른쪽에 '등록상태' 열이 자동으로 생김. 비어 있으면 접수로 봄.
 * 공개(등록상태 페이지) 기본 리스트에는 교회 · 가린 담당자 이름 · 참석 일정만 나감. 교회 + 담당자로 조회하면 참가인원 · 상태까지.
 */

var STATUS_HEADER = '등록상태';
// 접수: 신청서 작성 · 가등록: 접수 및 가등록금 납부 · 완등록: 최종 등록 마무리
var STATUSES = ['접수', '가등록', '완등록'];
var LEGACY = { '가등록완료': '가등록' };

// 폼 응답이 들어오는 시트 (보통 첫 번째 시트)
function sheet_() {
  return SpreadsheetApp.getActive().getSheets()[0];
}

// 앞에 있는 키워드부터 찾음. skip 에 걸리는 제목은 건너뜀 (없으면 -1)
function findCol_(head, words, skip) {
  for (var j = 0; j < words.length; j++) {
    for (var i = 0; i < head.length; i++) {
      if (head[i].indexOf(words[j]) < 0) continue;
      if (skip && skip.test(head[i])) continue;
      return i;
    }
  }
  return -1;
}

// 연락처 열: 제목에 키워드가 있고 값이 전화번호처럼 생긴(숫자 8자리 이상) 열
function phoneCol_(head, values) {
  var words = ['연락처', '전화', '휴대폰', '핸드폰', '번호'];
  var looksPhone = function (c) {
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][c]).replace(/\D/g, '').length >= 8) return true;
    }
    return false;
  };
  for (var c = 0; c < head.length; c++) {
    var hit = words.some(function (w) { return head[c].indexOf(w) >= 0; });
    if (hit && looksPhone(c)) return c;
  }
  return -1;
}

function readRows_() {
  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  var head = values[0].map(function (h) { return String(h).trim(); });
  var cStatus = head.indexOf(STATUS_HEADER);
  if (cStatus < 0) {
    cStatus = head.length;
    sh.getRange(1, cStatus + 1).setValue(STATUS_HEADER);
  }
  var notName = /연락|전화|번호|메일|이메일/;
  var cTime = findCol_(head, ['타임스탬프', 'Timestamp']);
  var cChurch = findCol_(head, ['교회', '소속']);
  var cName = findCol_(head, ['신청인', '담당자 성함', '담당자 이름', '담당자명', '성함', '이름', '성명', '담당자'], notName);
  var cSchedule = findCol_(head, ['일정', '참석', '참가', '날짜'], /인원|명수|동의/);
  var cCount = findCol_(head, ['참가인원', '참가 인원', '참석인원', '참석 인원', '인원', '명수']);
  var cPhone = phoneCol_(head, values);
  var get = function (r, c) { return c >= 0 && c < r.length ? r[c] : ''; };
  // 어드민용: 시트의 모든 열 (등록상태 열 제외), 날짜는 보기 좋게
  var tz = Session.getScriptTimeZone();
  var cell = function (v) { return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy.MM.dd HH:mm') : String(v); };
  var headers = head.filter(function (h, c) { return c !== cStatus; });

  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    var church = String(get(r, cChurch)).trim();
    var name = String(get(r, cName)).trim();
    if (!church && !name) continue;
    var status = String(get(r, cStatus)).trim();
    status = LEGACY[status] || status;
    var time = get(r, cTime);
    rows.push({
      row: i + 1,
      time: time instanceof Date ? time.toISOString() : String(time),
      church: church,
      name: name,
      schedule: String(get(r, cSchedule)).trim(),
      count: String(get(r, cCount)).trim(),
      phone: String(get(r, cPhone)).trim(),
      status: STATUSES.indexOf(status) >= 0 ? status : STATUSES[0],
      cells: head.map(function (h, c) { return c === cStatus ? null : cell(get(r, c)); }).filter(function (v) { return v !== null; })
    });
  }
  return { rows: rows, headers: headers, statusCol: cStatus + 1 };
}

// '김홀젠 목사' → 이름 '김홀젠' + 직분 ' 목사'
var ROLES = /(목사|부목사|담임목사|전도사|강도사|간사|교사|선생님|선생|장로|집사|권사|사모|부장|리더|청년|학생)$/;
function splitName_(full) {
  full = String(full).trim();
  var m = full.match(/^(\S+)(\s.*)$/);
  if (m) return { name: m[1], rest: m[2] };
  var r = full.match(ROLES);
  if (r && full.length - r[0].length >= 2) return { name: full.slice(0, -r[0].length), rest: ' ' + r[0] };
  return { name: full, rest: '' };
}

function maskName_(n) {
  if (n.length <= 1) return n;
  if (n.length === 2) return n[0] + '*';
  return n[0] + '*'.repeat(n.length - 2) + n[n.length - 1];
}

// 이름 부분만 가림: '김홀젠 목사' → '김*젠 목사'
function mask_(full) {
  var p = splitName_(full);
  return maskName_(p.name) + p.rest;
}

// 조회: 이름만('김홀젠') 넣어도, 직분까지('김홀젠 목사') 넣어도 맞음
function nameMatch_(full, q) {
  var nq = q.replace(/\s/g, '');
  return full.replace(/\s/g, '') === nq || splitName_(full).name === nq;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 교회 이름 비교용: 띄어쓰기와 '교회'를 뺌 ('교회'만으로는 검색 안 되게)
function churchKey_(s) {
  return String(s).replace(/\s/g, '').replace(/교회/g, '');
}

// 회의록: 같은 스프레드시트의 '회의록' 탭 (없으면 자동으로 만듦)
var NOTE_HEADERS = ['id', '날짜', '제목', '참석자', '내용', '수정일'];
function notesSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('회의록');
  if (!sh) {
    sh = ss.insertSheet('회의록', ss.getNumSheets());
    sh.getRange(1, 1, 1, NOTE_HEADERS.length).setValues([NOTE_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readNotes_() {
  var sh = notesSheet_();
  var values = sh.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var d = function (v) { return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v); };
  var notes = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    notes.push({ id: String(r[0]), date: d(r[1]), title: String(r[2]), attendees: String(r[3]), body: String(r[4]),
      updated: r[5] instanceof Date ? r[5].toISOString() : String(r[5]), row: i + 1 });
  }
  notes.sort(function (a, b) { return b.date.localeCompare(a.date) || b.updated.localeCompare(a.updated); });
  return notes;
}

// ID / PW 목록: 같은 스프레드시트의 '계정' 탭 (없으면 자동으로 만듦)
var ACCOUNT_HEADERS = ['id', '서비스', '아이디', '비밀번호', '메모', '수정일'];
function accountsSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('계정');
  if (!sh) {
    sh = ss.insertSheet('계정', ss.getNumSheets());
    sh.getRange(1, 1, 1, ACCOUNT_HEADERS.length).setValues([ACCOUNT_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAccounts_() {
  var values = accountsSheet_().getDataRange().getDisplayValues();
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    list.push({ id: r[0], service: r[1], login: r[2], password: r[3], memo: r[4], updated: r[5], row: i + 1 });
  }
  list.sort(function (a, b) { return a.service.localeCompare(b.service, 'ko'); });
  return list;
}

// 전화부: 같은 스프레드시트의 '전화부' 탭 (없으면 자동으로 만듦)
var CONTACT_HEADERS = ['id', '이름', '소속', '역할', '전화번호', '이메일', '메모', '수정일'];
function contactsSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('전화부');
  if (!sh) {
    sh = ss.insertSheet('전화부', ss.getNumSheets());
    sh.getRange(1, 1, 1, CONTACT_HEADERS.length).setValues([CONTACT_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readContacts_() {
  var values = contactsSheet_().getDataRange().getDisplayValues();
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    list.push({ id: r[0], name: r[1], org: r[2], role: r[3], phone: r[4], email: r[5], memo: r[6], updated: r[7], row: i + 1 });
  }
  list.sort(function (a, b) { return a.name.localeCompare(b.name, 'ko'); });
  return list;
}

// 처음 한 번: 편집기에서 이 함수를 골라 '실행'하면 캘린더 권한 허용 창이 떠요
function authorize() {
  Logger.log(CalendarApp.getDefaultCalendar().getName());
}

// 공개: 등록상태 페이지
//  - 기본: 전체 명단, 담당자 이름은 가림
//  - ?church=교회이름&name=담당자이름: 둘 다 맞는 신청만, 담당자 이름 그대로
function doGet(e) {
  var rows = readRows_().rows;
  var p = (e && e.parameter) || {};
  var church = churchKey_(p.church || '');
  var name = String(p.name || '').replace(/\s/g, '');
  var search = church.length >= 1 && name.length >= 2;
  if (search) rows = rows.filter(function (r) {
    return churchKey_(r.church).indexOf(church) >= 0 && nameMatch_(r.name, name);
  });
  // 기본 리스트에는 참가인원 · 등록상태를 보내지 않음 (조회했을 때만)
  var items = rows.map(function (r) {
    return search
      ? { church: r.church, name: r.name, schedule: r.schedule, count: r.count, status: r.status }
      : { church: r.church, name: mask_(r.name), schedule: r.schedule };
  });
  return json_({ ok: true, search: search, updated: new Date().toISOString(), items: items });
}

// 어드민: { pw, action: 'list' } · { pw, action: 'set', row, status }
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  var pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!pw || body.pw !== pw) return json_({ ok: false, error: 'auth' });

  if (body.action === 'list') {
    var data = readRows_();
    return json_({ ok: true, updated: new Date().toISOString(), statuses: STATUSES, headers: data.headers, items: data.rows });
  }

  // 어드민 홈: 구글 캘린더 다가오는 일정 (기본 캘린더, 스크립트 속성 CALENDAR_ID 가 있으면 그 캘린더)
  if (body.action === 'events') {
    var calId = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');
    var cal = calId ? CalendarApp.getCalendarById(calId) : CalendarApp.getDefaultCalendar();
    if (!cal) return json_({ ok: false, error: 'calendar' });
    var from = new Date(); from.setHours(0, 0, 0, 0);
    var to = new Date(from.getTime() + 365 * 86400000);  // 앞으로 1년
    var events = cal.getEvents(from, to).slice(0, 100).map(function (ev) {
      return {
        title: ev.getTitle(),
        start: ev.getStartTime().toISOString(),
        end: ev.getEndTime().toISOString(),
        allDay: ev.isAllDayEvent(),
        location: ev.getLocation()
      };
    });
    return json_({ ok: true, calendar: cal.getName(), events: events });
  }

  // 회의록: 목록 · 저장(새로 쓰기/고치기) · 삭제
  if (body.action === 'notes') {
    return json_({ ok: true, notes: readNotes_() });
  }
  if (body.action === 'note_save' || body.action === 'note_delete') {
    var lockN = LockService.getScriptLock();
    lockN.waitLock(10000);
    try {
      var sh = notesSheet_();
      var n = body.note || {};
      var found = readNotes_().filter(function (x) { return x.id === String(n.id || body.id || ''); })[0];
      if (body.action === 'note_delete') {
        if (found) sh.deleteRow(found.row);
        return json_({ ok: true });
      }
      var rowVals = [found ? found.id : Utilities.getUuid(), String(n.date || ''), String(n.title || ''), String(n.attendees || ''), String(n.body || ''), new Date()];
      if (found) sh.getRange(found.row, 1, 1, rowVals.length).setValues([rowVals]);
      else sh.appendRow(rowVals);
      return json_({ ok: true, id: rowVals[0] });
    } finally {
      lockN.releaseLock();
    }
  }

  // ID / PW: 목록 · 저장 · 삭제
  if (body.action === 'accounts') {
    return json_({ ok: true, accounts: readAccounts_() });
  }
  if (body.action === 'account_save' || body.action === 'account_delete') {
    var lockA = LockService.getScriptLock();
    lockA.waitLock(10000);
    try {
      var shA = accountsSheet_();
      var ac = body.account || {};
      var hit = readAccounts_().filter(function (x) { return x.id === String(ac.id || body.id || ''); })[0];
      if (body.action === 'account_delete') {
        if (hit) shA.deleteRow(hit.row);
        return json_({ ok: true });
      }
      var vals = [hit ? hit.id : Utilities.getUuid(), String(ac.service || ''), String(ac.login || ''), String(ac.password || ''), String(ac.memo || ''), new Date()];
      var rng = hit ? shA.getRange(hit.row, 1, 1, vals.length) : shA.getRange(shA.getLastRow() + 1, 1, 1, vals.length);
      rng.setNumberFormat('@').setValues([vals.map(function (v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy.MM.dd HH:mm') : v; })]);
      return json_({ ok: true, id: vals[0] });
    } finally {
      lockA.releaseLock();
    }
  }

  // 전화부: 목록 · 저장 · 삭제
  if (body.action === 'contacts') {
    return json_({ ok: true, contacts: readContacts_() });
  }
  if (body.action === 'contact_save' || body.action === 'contact_delete') {
    var lockC = LockService.getScriptLock();
    lockC.waitLock(10000);
    try {
      var shC = contactsSheet_();
      var ct = body.contact || {};
      var hitC = readContacts_().filter(function (x) { return x.id === String(ct.id || body.id || ''); })[0];
      if (body.action === 'contact_delete') {
        if (hitC) shC.deleteRow(hitC.row);
        return json_({ ok: true });
      }
      var cvals = [hitC ? hitC.id : Utilities.getUuid(), String(ct.name || ''), String(ct.org || ''), String(ct.role || ''), String(ct.phone || ''), String(ct.email || ''), String(ct.memo || ''),
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy.MM.dd HH:mm')];
      var crng = hitC ? shC.getRange(hitC.row, 1, 1, cvals.length) : shC.getRange(shC.getLastRow() + 1, 1, 1, cvals.length);
      crng.setNumberFormat('@').setValues([cvals]);
      return json_({ ok: true, id: cvals[0] });
    } finally {
      lockC.releaseLock();
    }
  }

  if (body.action === 'set') {
    var row = Number(body.row);
    if (STATUSES.indexOf(body.status) < 0 || !(row >= 2)) return json_({ ok: false, error: 'bad' });
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var data = readRows_();
      if (row > sheet_().getLastRow()) return json_({ ok: false, error: 'bad' });
      sheet_().getRange(row, data.statusCol).setValue(body.status);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  }

  return json_({ ok: false, error: 'bad' });
}
