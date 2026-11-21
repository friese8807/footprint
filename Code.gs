/**
 * 발자취(Footprint) — Google Sheets 백엔드 (Google Apps Script)
 *
 * 설치: 구글 시트 > 확장 프로그램 > Apps Script 에 이 파일 전체를 붙여넣고
 *       1) CONFIG.TOKEN 을 나만의 긴 문자열로 바꾼 뒤
 *       2) 상단에서 setup 함수를 한 번 실행(권한 승인)
 *       3) 배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 모든 사용자)
 *
 * 시트 구성
 *   Log    : 위치 기록 (1시간마다 단축어/앱에서 append)
 *   Places : 등록 장소 (집, 회사 등) — 기록 시 자동으로 장소명이 태깅됨
 */

const CONFIG = {
  TOKEN: 'fp-d3kBM-J6fNQwyTKBoGzwUUgKeo798Cbg', // 토큰을 바꾸면 앱·단축어·위젯도 같이 바꿔야 함
  SHEET_ID: '1YUt8n2Hu4SRbJRcGRO1FcdNoN_Iwlw67Z_m4KdMvdcM', // 발자취 위치 기록 시트
  TZ: 'Asia/Seoul',
  LOG: 'Log',
  PLACES: 'Places',
  REVERSE_GEOCODE: true,   // Apps Script 내장 지오코더로 주소 자동 기입
  MIN_GAP_MINUTES: 5,      // 같은 위치에서 5분 이내 중복 기록은 무시
  TRIP: { CAL_ID: 'ha6sk6smfch90l4mlf24dm9c84@group.calendar.google.com', TITLE: '[출장] 부산', FROM: '14:00', TO: '14:59' },
};


const LOG_HEADERS = ['id', 'timestamp', 'date', 'time', 'lat', 'lng', 'accuracy', 'altitude', 'battery', 'source', 'place', 'address', 'memo'];
const PLACE_HEADERS = ['name', 'emoji', 'lat', 'lng', 'radius'];

/** 독립 실행형 스크립트에서도 동작하도록 시트를 ID로 연다 */
function ss_() {
  return CONFIG.SHEET_ID ? SpreadsheetApp.openById(CONFIG.SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

/* ───────────── 최초 1회 실행 ───────────── */
function setup() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone(CONFIG.TZ);
  const log = sheet_(CONFIG.LOG, LOG_HEADERS);
  const places = sheet_(CONFIG.PLACES, PLACE_HEADERS);
  // 날짜/시간/ID 가 자동 변환되지 않도록 텍스트 서식 고정
  log.getRange('A:D').setNumberFormat('@');
  log.getRange('J:M').setNumberFormat('@');
  places.getRange('A:B').setNumberFormat('@');
  log.setColumnWidth(12, 320);
  return 'ok';
}

function sheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ───────────── 엔드포인트 ───────────── */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const out = handle_(Object.assign({ action: 'summary' }, p));
  // 프레임 방식: 숨은 iframe 으로 열어 결과를 postMessage 로 앱에 전달 (fetch·JSONP 가 막힌 Safari 대비)
  const fid = String(p.frame || '').replace(/[^\w]/g, '').slice(0, 40);
  if (fid) {
    const data = out.getContent().replace(/</g, '\\u003c');
    return HtmlService.createHtmlOutput('<script>window.top.postMessage({fp:"' + fid + '",data:' + data + '},"*");</script>')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // JSONP: 아이폰 홈 화면 앱처럼 fetch(CORS)가 막히는 환경을 위해 callback=함수명 지원
  const cb = String(p.callback || '');
  if (/^[A-Za-z_$][\w$]{0,40}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + out.getContent() + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return out;
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (_) { body = {}; }
  return handle_(Object.assign({ action: 'log' }, (e && e.parameter) || {}, body));
}

function handle_(p) {
  try {
    if (String(p.token || '') !== CONFIG.TOKEN) return json_({ ok: false, error: 'unauthorized' });
    switch (p.action) {
      case 'log':         return json_(logPoint_(p));
      case 'day':         return json_({ ok: true, date: p.date, points: pointsOn_(p.date || today_()) });
      case 'days':        return json_({ ok: true, days: dayCounts_() });
      case 'summary':     return json_(summary_(p.date || today_()));
      case 'places':      return json_({ ok: true, places: places_() });
      case 'addPlace':    return json_(addPlace_(p));
      case 'deletePlace': return json_(deleteRow_(CONFIG.PLACES, 0, p.name));
      case 'deletePoint': return json_(deleteRow_(CONFIG.LOG, 0, p.id));
      case 'memo':        return json_(setMemo_(p.id, p.memo));
      case 'ping':        return json_({ ok: true, now: new Date().toISOString() });
      default:            return json_({ ok: false, error: 'unknown action: ' + p.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ───────────── 기록 ───────────── */
function logPoint_(p) {
  const lat = num_(p.lat), lng = num_(p.lng);
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) {
    return { ok: false, error: 'invalid lat/lng' };
  }
  const when = p.timestamp ? new Date(p.timestamp) : new Date();
  const ts = isNaN(when) ? new Date() : when;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_(CONFIG.LOG, LOG_HEADERS);

    // 중복 방지: 직전 기록과 MIN_GAP 이내 + 50m 이내면 skip
    const last = sh.getLastRow();
    if (last > 1) {
      const prev = sh.getRange(last, 1, 1, 6).getValues()[0];
      const gapMin = (ts - new Date(prev[1])) / 60000;
      if (gapMin >= 0 && gapMin < CONFIG.MIN_GAP_MINUTES && haversine_(lat, lng, +prev[4], +prev[5]) < 50) {
        return { ok: true, skipped: true, reason: 'duplicate' };
      }
    }

    const place = matchPlace_(lat, lng);
    const row = [
      Utilities.getUuid().slice(0, 8),
      ts.toISOString(),
      Utilities.formatDate(ts, CONFIG.TZ, 'yyyy-MM-dd'),
      Utilities.formatDate(ts, CONFIG.TZ, 'HH:mm'),
      round_(lat, 6), round_(lng, 6),
      num_(p.accuracy) ? round_(num_(p.accuracy), 0) : '',
      num_(p.altitude) ? round_(num_(p.altitude), 0) : '',
      battery_(p.battery),
      String(p.source || 'shortcut'),
      place ? place.name : '',
      reverseGeocode_(lat, lng),
      String(p.memo || ''),
    ];
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
        let trip = '';
    try { trip = autoTrip_(ts, row[3], lat, lng, row[11]); } catch (err) { trip = 'error: ' + err.message; }
    return { ok: true, point: toPoint_(row), trip: trip };
  } finally {
    lock.releaseLock();
  }
}

function reverseGeocode_(lat, lng) {
  if (!CONFIG.REVERSE_GEOCODE) return '';
  try {
    const r = Maps.newGeocoder().setLanguage('ko').setRegion('kr').reverseGeocode(lat, lng);
    if (r && r.results && r.results.length) {
      return String(r.results[0].formatted_address || '').replace(/^대한민국\s*/, '');
    }
  } catch (_) {}
  return '';
}

/* ───────────── 조회 ───────────── */
function readLog_() {
  const sh = sheet_(CONFIG.LOG, LOG_HEADERS);
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, LOG_HEADERS.length).getValues().map(toPoint_);
}

function toPoint_(r) {
  return {
    id: String(r[0]), ts: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
    date: r[2] instanceof Date ? Utilities.formatDate(r[2], CONFIG.TZ, 'yyyy-MM-dd') : String(r[2]),
    time: String(r[3]), lat: +r[4], lng: +r[5],
    accuracy: r[6] === '' ? null : +r[6], altitude: r[7] === '' ? null : +r[7],
    battery: r[8] === '' ? null : +r[8], source: String(r[9] || ''),
    place: String(r[10] || ''), address: String(r[11] || ''), memo: String(r[12] || ''),
  };
}

function pointsOn_(date) {
  return readLog_().filter(function (pt) { return pt.date === date; })
    .sort(function (a, b) { return a.ts < b.ts ? -1 : 1; });
}

function dayCounts_() {
  const out = {};
  readLog_().forEach(function (pt) { out[pt.date] = (out[pt.date] || 0) + 1; });
  return out;
}

/** 위젯용 요약 */
function summary_(date) {
  const pts = pointsOn_(date);
  let dist = 0;
  for (let i = 1; i < pts.length; i++) dist += haversine_(pts[i - 1].lat, pts[i - 1].lng, pts[i].lat, pts[i].lng);
  const placeSet = {};
  pts.forEach(function (pt) { placeSet[pt.place || pt.address.split(' ').slice(0, 3).join(' ')] = 1; });
  return {
    ok: true, date: date, count: pts.length,
    distanceKm: round_(dist / 1000, 1),
    places: Object.keys(placeSet).filter(String).length,
    last: pts.length ? pts[pts.length - 1] : null,
    path: pts.map(function (pt) {
      return [round_(pt.lat, 5), round_(pt.lng, 5), pt.time, pt.place || pt.address.split(' ').slice(-2).join(' ')];
    }),
  };
}

/* ───────────── 장소 ───────────── */
function places_() {
  const sh = sheet_(CONFIG.PLACES, PLACE_HEADERS);
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, PLACE_HEADERS.length).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) { return { name: String(r[0]), emoji: String(r[1] || '📍'), lat: +r[2], lng: +r[3], radius: +r[4] || 150 }; });
}

function matchPlace_(lat, lng) {
  let best = null, bestD = Infinity;
  places_().forEach(function (pl) {
    const d = haversine_(lat, lng, pl.lat, pl.lng);
    if (d <= pl.radius && d < bestD) { best = pl; bestD = d; }
  });
  return best;
}

function addPlace_(p) {
  const lat = num_(p.lat), lng = num_(p.lng);
  if (!p.name || !isFinite(lat) || !isFinite(lng)) return { ok: false, error: 'name/lat/lng required' };
  deleteRow_(CONFIG.PLACES, 0, p.name); // 같은 이름은 덮어쓰기
  const sh = sheet_(CONFIG.PLACES, PLACE_HEADERS);
  sh.appendRow([String(p.name), String(p.emoji || '📍'), round_(lat, 6), round_(lng, 6), num_(p.radius) || 150]);
  return { ok: true, places: places_() };
}

function deleteRow_(sheetName, col, key) {
  if (!key) return { ok: false, error: 'key required' };
  const sh = ss_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return { ok: true, deleted: 0 };
  const vals = sh.getRange(2, col + 1, sh.getLastRow() - 1, 1).getValues();
  let deleted = 0;
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]) === String(key)) { sh.deleteRow(i + 2); deleted++; }
  }
  return { ok: true, deleted: deleted };
}

function setMemo_(id, memo) {
  const sh = sheet_(CONFIG.LOG, LOG_HEADERS);
  if (sh.getLastRow() < 2) return { ok: false, error: 'not found' };
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sh.getRange(i + 2, 13).setValue(String(memo || ''));
      return { ok: true };
    }
  }
  return { ok: false, error: 'not found' };
}

/* ───────────── 유틸 ───────────── */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function today_() { return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd'); }
function num_(v) {
  if (v === null || v === undefined || v === '') return NaN;
  return parseFloat(String(v).replace(',', '.'));
}
function battery_(v) {
  const b = num_(v);
  if (!isFinite(b)) return '';
  return Math.round(b <= 1 ? b * 100 : b); // 0.82 또는 82 모두 허용
}
function round_(v, d) { const f = Math.pow(10, d); return Math.round(v * f) / f; }
function haversine_(lat1, lng1, lat2, lng2) {
  const R = 6371000, toR = Math.PI / 180;
  const dLat = (lat2 - lat1) * toR, dLng = (lng2 - lng1) * toR;
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1 * toR) * Math.cos(lat2 * toR) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

function isBusan_(lat, lng, address) {
  if (/부산/.test(address || '')) return true;
  return lat > 34.95 && lat < 35.40 && lng > 128.75 && lng < 129.32;
}
function autoTrip_(ts, hhmm, lat, lng, address) {
  const T = CONFIG.TRIP;
  if (!T || hhmm < T.FROM || hhmm > T.TO || !isBusan_(lat, lng, address)) return '';
  const cal = CalendarApp.getCalendarById(T.CAL_ID);
  if (!cal) return 'calendar not found';
  const ymd = Utilities.formatDate(ts, CONFIG.TZ, 'yyyy-MM-dd').split('-').map(Number);
  const day = new Date(ymd[0], ymd[1] - 1, ymd[2], 12);
  const dup = cal.getEventsForDay(day).some(function (ev) { const t = ev.getTitle(); return /출장/.test(t) && /부산/.test(t); });
  if (dup) return 'exists';
  cal.createAllDayEvent(T.TITLE, day);
  return 'created';
}
function authorizeCalendar() {
  const cal = CalendarApp.getCalendarById(CONFIG.TRIP.CAL_ID);
  Logger.log(cal ? '캘린더 연결 OK: ' + cal.getName() : '캘린더를 찾지 못했어요');
}

/** 편집기에서 실행해 보는 테스트 */
function testLog() {
  return logPoint_({ token: CONFIG.TOKEN, lat: 37.5665, lng: 126.9780, source: 'test' });
}