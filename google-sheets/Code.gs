/**
 * 실내자전거 트래커 ↔ 구글 시트 연동 스크립트
 *
 * 설치 방법(자세히): https://github.com/wbhongb01/indoor-bike/blob/main/google-sheets/README.md
 *  1) 구글 시트 → 확장 프로그램 → Apps Script 에 이 코드를 통째로 붙여넣기
 *  2) 바로 아래 FAMILY_KEY 를 가족끼리만 아는 비밀번호로 바꾸기
 *  3) 배포 → 새 배포 → 웹 앱 (다음 사용자로 실행: 나 / 액세스 권한: 모든 사용자)
 *  4) 나온 웹 앱 주소(…/exec)와 비밀번호를 앱의 설정 → 구글 시트 연동에 입력
 */

// ▼▼▼ 여기를 가족끼리만 아는 비밀번호로 바꾸세요 ▼▼▼
const FAMILY_KEY = '여기에-가족-비밀번호';
// ▲▲▲ 바꾼 뒤에는 저장하고 "새 버전"으로 다시 배포해야 적용됩니다 ▲▲▲

const SHEET = { users: '사용자', sessions: '운동기록', settings: '설정', summary: '월별 요약' };
const FMT = { t: '@', n0: '0', n1: '0.0', n2: '0.00', g: '0.##' };
// [앱에서 쓰는 이름, 시트 머리글, 셀 형식]
const USER_COLS = [
  ['name', '이름', 't'], ['weight', '체중(kg)', 'n1'], ['height', '키(cm)', 'n0'], ['age', '나이', 'n0'],
  ['sex', '성별', 't'], ['color', '색상', 't'], ['id', '사용자ID', 't']
];
const SESSION_COLS = [
  ['month', '월', 't'], ['local', '날짜', 't'], ['userName', '사용자', 't'], ['minutes', '운동시간(분)', 'n1'],
  ['distance', '거리(km)', 'n2'], ['kcal', '칼로리(kcal)', 'n0'], ['avgSpeed', '평균속도(km/h)', 'n1'],
  ['avgRpm', '평균RPM', 'n0'], ['avgPower', '평균파워(W)', 'n0'], ['maxSpeed', '최고속도(km/h)', 'n1'],
  ['level', '저항', 'n0'], ['sourceName', '입력방식', 't'],
  ['id', '기록ID', 't'], ['userId', '사용자ID', 't'], ['date', '시작시각(ISO)', 't'], ['duration', '운동시간(초)', 'n0'], ['source', 'source', 't']
];
const SETTING_COLS = [['key', '항목', 't'], ['value', '값', 'g'], ['desc', '설명', 't']];
const SETTING_KEYS = [['wheel', '바퀴 둘레(m)'], ['gear', '기어비(페달 1회전당 바퀴 회전수)'], ['goal', '하루 목표 운동 시간(분)']];
const SOURCE_NAME = { ble: '블루투스', tap: '페달 탭', manual: '수동', entry: '직접 입력', sheet: '시트 입력' };
const PALETTE = ['#ff6a2b', '#2f7cf6', '#10b981', '#a855f7', '#e5a50a', '#ec4899', '#14b8a6', '#64748b'];
const SUMMARY_FORMULA = `=IFERROR(QUERY('${SHEET.sessions}'!A2:Q, "select A, C, count(M), sum(D), sum(E), sum(F) where M <> '' group by A, C order by A desc, C label A '월', C '사용자', count(M) '운동 횟수', sum(D) '운동시간(분)', sum(E) '거리(km)', sum(F) '칼로리(kcal)'", 0), "아직 운동 기록이 없어요")`;

/* ---------- 웹 앱 입구 ---------- */
function doGet() {
  return json_({ ok: true, app: 'indoor-bike', msg: '연결 주소가 정상이에요. 이 주소를 앱의 설정 → 구글 시트 연동에 붙여넣으세요.' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'BAD_REQUEST', msg: '요청 내용을 읽을 수 없어요.' }); }

  if (!FAMILY_KEY || FAMILY_KEY === '여기에-가족-비밀번호')
    return json_({ ok: false, error: 'NO_KEY', msg: '스크립트의 FAMILY_KEY(가족 비밀번호)를 바꾼 뒤 새 버전으로 다시 배포하세요.' });
  if (String(req.key || '') !== FAMILY_KEY)
    return json_({ ok: false, error: 'BAD_KEY', msg: '가족 비밀번호가 맞지 않아요.' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return json_({ ok: false, error: 'BUSY', msg: '다른 기기가 저장 중이에요. 잠시 후 다시 시도하세요.' });
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const st = load_(ss);
    const ops = Array.isArray(req.ops) ? req.ops : [];
    ops.forEach(op => applyOp_(st, op));
    if (ops.length || st.dirty) save_(ss, st);
    return json_({
      ok: true,
      data: {
        users: st.users.map(publicUser_),
        sessions: st.sessions.map(publicSession_),
        settings: st.settings,
        sheetUrl: ss.getUrl()
      }
    });
  } catch (err) {
    return json_({ ok: false, error: 'SERVER', msg: '시트 처리 중 오류가 났어요: ' + ((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

/** (선택) 편집기에서 한 번 실행하면 시트 탭을 미리 만들고 권한을 승인합니다. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  save_(ss, load_(ss));
}

/* ---------- 시트 읽기 ---------- */
function load_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const st = { tz: tz, dirty: false, users: [], sessions: [], settings: {} };
  const uSh = sheet_(ss, SHEET.users, USER_COLS);
  const sSh = sheet_(ss, SHEET.sessions, SESSION_COLS);
  const cSh = sheet_(ss, SHEET.settings, SETTING_COLS);
  summary_(ss);

  rows_(uSh, USER_COLS.length).forEach(r => {
    const u = toObj_(r, USER_COLS, tz);
    if (!u.name && !u.id) return;
    if (!u.id) { u.id = newId_('u'); st.dirty = true; }
    st.users.push(cleanUser_(u));
  });
  const idByName = {};
  st.users.forEach(u => { idByName[u.name] = u.id; });

  rows_(sSh, SESSION_COLS.length).forEach(r => {
    const s = toObj_(r, SESSION_COLS, tz);
    if (!s.id && !s.local && !s.date && s.minutes == null) return;      // 빈 줄
    if (!s.id) { s.id = newId_('s'); st.dirty = true; }                   // 시트에 직접 적은 줄
    if (!s.userId || !st.users.some(u => u.id === s.userId)) {
      let id = idByName[s.userName];
      if (!id && s.userName) {
        const u = cleanUser_({ id: newId_('u'), name: s.userName });
        st.users.push(u); id = idByName[u.name] = u.id;
      }
      if (!id && st.users.length) id = st.users[0].id;
      if ((id || '') !== s.userId) { s.userId = id || ''; st.dirty = true; }
    }
    // 시트에서 '날짜'를 고쳤거나, 직접 적은 줄이면 시작시각을 날짜에 맞춤
    const fromLocal = s.local ? parseLocal_(s.local, tz) : null;
    if (!s.date) { s.date = (fromLocal || new Date()).toISOString(); st.dirty = true; }
    else if (fromLocal && Utilities.formatDate(new Date(s.date), tz, 'yyyy-MM-dd HH:mm') !== Utilities.formatDate(fromLocal, tz, 'yyyy-MM-dd HH:mm')) {
      s.date = fromLocal.toISOString(); st.dirty = true;
    }
    // 시트에서 '운동시간(분)'을 고쳤으면 초 단위 값도 맞춤
    if (s.minutes != null && Math.abs(s.minutes - (s.duration || 0) / 60) > 0.06) { s.duration = Math.round(s.minutes * 60); st.dirty = true; }
    if (!s.source) s.source = 'sheet';
    st.sessions.push(cleanSession_(s));
  });

  rows_(cSh, 2).forEach(r => {
    const k = String(r[0]).trim(), v = Number(r[1]);
    if (SETTING_KEYS.some(x => x[0] === k) && v > 0) st.settings[k] = v;
  });
  return st;
}

function sheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name, ss.getSheets().length);
    sh.getRange(1, 1, 1, cols.length).setValues([cols.map(c => c[1])]).setFontWeight('bold').setBackground('#ffe6da');
    sh.setFrozenRows(1);
    cleanupDefaultSheet_(ss);
  }
  return sh;
}

function summary_(ss) {
  if (ss.getSheetByName(SHEET.summary)) return;
  const sh = ss.insertSheet(SHEET.summary, ss.getSheets().length);
  sh.getRange('A1').setFormula(SUMMARY_FORMULA);
  sh.getRange('A1:F1').setFontWeight('bold').setBackground('#ffe6da');
  sh.setFrozenRows(1);
}

// 새 스프레드시트에 처음부터 있던 빈 '시트1' 정리
function cleanupDefaultSheet_(ss) {
  ss.getSheets().forEach(sh => {
    if (/^(시트|Sheet)\d+$/.test(sh.getName()) && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}

function rows_(sh, n) {
  const last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, n).getValues();
}

function toObj_(r, cols, tz) {
  const o = {};
  cols.forEach((c, i) => {
    let v = r[i];
    if (c[2] === 't') {
      if (Object.prototype.toString.call(v) === '[object Date]') {
        v = c[0] === 'month' ? Utilities.formatDate(v, tz, 'yyyy-MM')
          : c[0] === 'date' ? v.toISOString()
          : Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm');
      }
      o[c[0]] = v == null ? '' : String(v).trim();
    } else {
      o[c[0]] = (v === '' || v == null || isNaN(Number(v))) ? null : Number(v);
    }
  });
  return o;
}

function parseLocal_(str, tz) {
  str = String(str || '').trim().replace(/\s*[./]\s*/g, '-').replace(/-+$/, '').replace(/\s+/g, ' ');
  const fmts = ['yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd HH:mm', 'yyyy-MM-dd'];
  for (let i = 0; i < fmts.length; i++) {
    try { const d = Utilities.parseDate(str, tz, fmts[i]); if (d && !isNaN(d.getTime())) return d; } catch (e) { /* 다음 형식 */ }
  }
  return null;
}

/* ---------- 변경 적용 ---------- */
function applyOp_(st, op) {
  if (!op || typeof op !== 'object') return;
  const upsert = (list, item) => { const i = list.findIndex(x => x.id === item.id); if (i >= 0) list[i] = item; else list.push(item); };
  switch (op.op) {
    case 'putUser':
      if (op.user && op.user.id) upsert(st.users, cleanUser_(op.user));
      break;
    case 'delUser':
      st.users = st.users.filter(u => u.id !== op.id);
      st.sessions = st.sessions.filter(s => s.userId !== op.id);
      break;
    case 'putSession':
      if (op.session && op.session.id) upsert(st.sessions, cleanSession_(op.session));
      break;
    case 'delSession':
      st.sessions = st.sessions.filter(s => s.id !== op.id);
      break;
    case 'putSettings':
      SETTING_KEYS.forEach(k => { const v = Number(op.settings && op.settings[k[0]]); if (v > 0) st.settings[k[0]] = v; });
      break;
  }
}

function cleanUser_(u) {
  const sex = String(u.sex || '').trim(), id = String(u.id);
  let h = 0; for (let i = 0; i < id.length; i++) h = (h * 31 + id.charCodeAt(i)) >>> 0;
  return {
    id: id,
    name: String(u.name || '').trim().slice(0, 20) || '이름 없음',
    weight: Number(u.weight) > 0 ? Number(u.weight) : 70,
    height: Number(u.height) > 0 ? Number(u.height) : null,
    age: Number(u.age) > 0 ? Number(u.age) : null,
    sex: (sex === 'F' || sex === '여') ? 'F' : 'M',
    color: /^#[0-9a-f]{6}$/i.test(String(u.color || '')) ? String(u.color) : PALETTE[h % PALETTE.length]
  };
}

function cleanSession_(s) {
  const n = (v, d) => (v === '' || v == null || isNaN(Number(v))) ? (d === undefined ? 0 : d) : Number(v);
  return {
    id: String(s.id), userId: String(s.userId || ''), date: String(s.date || new Date().toISOString()),
    duration: Math.max(0, Math.round(n(s.duration))), distance: n(s.distance), avgRpm: n(s.avgRpm),
    avgSpeed: n(s.avgSpeed), maxSpeed: n(s.maxSpeed), avgPower: n(s.avgPower), kcal: n(s.kcal),
    level: n(s.level, 5), source: String(s.source || 'entry')
  };
}

/* ---------- 시트 쓰기 ---------- */
function save_(ss, st) {
  const names = {};
  st.users.forEach(u => { names[u.id] = u.name; });
  const users = st.users.map(u => Object.assign({}, u, { sex: u.sex === 'F' ? '여' : '남' }));
  const sessions = st.sessions.slice()
    .sort((a, b) => String(b.date).localeCompare(String(a.date)))     // 최근 기록이 위로
    .map(s => {
      const d = new Date(s.date);
      return Object.assign({}, s, {
        month: Utilities.formatDate(d, st.tz, 'yyyy-MM'),
        local: Utilities.formatDate(d, st.tz, 'yyyy-MM-dd HH:mm'),
        userName: names[s.userId] || '',
        minutes: Math.round(s.duration / 6) / 10,
        sourceName: SOURCE_NAME[s.source] || s.source
      });
    });
  const settings = SETTING_KEYS.filter(k => st.settings[k[0]] > 0).map(k => ({ key: k[0], value: st.settings[k[0]], desc: k[1] }));
  write_(ss.getSheetByName(SHEET.users), USER_COLS, users);
  write_(ss.getSheetByName(SHEET.sessions), SESSION_COLS, sessions);
  write_(ss.getSheetByName(SHEET.settings), SETTING_COLS, settings);
}

function write_(sh, cols, objs) {
  const n = cols.length, last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, n).clearContent();
  if (!objs.length) return;
  const rng = sh.getRange(2, 1, objs.length, n);
  const fmt = cols.map(c => FMT[c[2]] || '@');
  rng.setNumberFormats(objs.map(() => fmt));
  rng.setValues(objs.map(o => cols.map(c => (o[c[0]] == null ? '' : o[c[0]]))));
}

/* ---------- 응답 ---------- */
function publicUser_(u) {
  return { id: u.id, name: u.name, color: u.color, weight: u.weight, height: u.height, age: u.age, sex: u.sex };
}
function publicSession_(s) {
  return { id: s.id, userId: s.userId, date: s.date, duration: s.duration, distance: s.distance, avgRpm: s.avgRpm,
    avgSpeed: s.avgSpeed, maxSpeed: s.maxSpeed, avgPower: s.avgPower, kcal: s.kcal, level: s.level, source: s.source };
}
function newId_(p) { return p + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
