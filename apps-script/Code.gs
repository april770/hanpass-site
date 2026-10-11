/**
 * Passmakers - 한패스 부트캠프 · 참가자 신청 수신 (Apps Script 웹앱)
 *
 * 설치: 통합 관리 시트(소유자 계정)에서 [확장 프로그램 → Apps Script] → 이 코드 붙여넣기 → 저장
 *       → setup() 한 번 실행(권한 승인, 새 열 준비) → [배포 → 새 배포 → 웹 앱]
 *         실행 사용자: 나 / 액세스 권한: 모든 사용자 → 나온 …/exec 주소를 apply/index.html의 SCRIPT_URL에 붙여넣기
 * 코드를 고친 뒤에는 [배포 관리 → 연필 → 새 버전]. "새 배포"를 누르면 주소가 바뀜.
 *
 * 속도: 열 위치를 캐시하고, 붙어 있는 열은 한 번에 기록합니다. 시트 열 순서를 바꿨다면 setup()을 한 번 다시 실행하세요.
 *
 * 신청 내용은 '02_참가 신청 리스트' 탭에 "열 이름 기준"으로 들어갑니다(탭 이름이 바뀌어도 탭 ID로 찾음).
 * 구글폼이 만든 열 이름을 그대로 쓰므로 대시보드·다른 탭 수식은 그대로 동작합니다.
 *
 * 임시저장: 신청 페이지의 [임시저장하고 링크 복사하기] → 같은 탭에 '신청 상태 = 임시저장' 행으로 기록.
 *   같은 임시저장 ID로 다시 저장하면 그 행을 덮어쓰고, 최종 제출하면 그 행이 '제출 완료'로 바뀝니다(한 사람 한 행).
 *   이어쓰기 링크(?draft=ID)로 들어오면 doGet 이 저장된 내용을 돌려줍니다.
 *   신청 인원을 셀 때는 '신청 상태'가 '제출 완료'인 행만 세세요.
 */
const SHEET_NAME = '02_참가 신청 리스트';
const SHEET_ID = 267092303;   // 탭 ID(gid). 탭 이름을 바꿔도 이 ID로 먼저 찾음
const STATE_DRAFT = '임시저장';
const STATE_DONE = '제출 완료';

// 신청서 항목 → 시트 열. 열 이름이 조금 바뀌어도 찾을 수 있도록 '포함 문구'로 찾습니다.
const MAP = [
  ['timestamp',  ['타임스탬프']],
  ['state',      ['신청 상태']],
  ['name',       ['이름']],
  ['email',      ['이메일']],
  ['phone',      ['휴대전화']],
  ['nationality',['국적']],
  ['region',     ['사는 지역']],
  ['visa',       ['비자']],
  ['status',     ['현재 신분']],
  ['topik',      ['TOPIK']],
  ['korean',     ['한국어를 어느 정도']],
  ['field1',     ['1순위']],
  ['field2',     ['2순위']],
  ['attend',     ['1차 밋업']],
  ['motivation', ['지원한 이유']],
  ['video',      ['영상 링크']],
  ['hanwork',    ['HanWork 앱에 가입']],
  ['source',     ['어디에서 알게']],
  ['consent',    ['개인정보 수집·이용 안내를 확인']],
  ['media',      ['사진·영상 촬영 및 활용 안내']],
  ['news',       ['소식(다음 기수']],
  ['inflow',     ['유입 링크']],
  ['draftId',    ['임시저장 ID']],
];
// 웹 신청서에서 새로 생기는 열 (없으면 자동으로 만듦: [새 열 이름, 이 열 오른쪽에])
const NEW_COLUMNS = [
  ['유입 링크', '소식(다음 기수'],
  ['신청 상태', '타임스탬프'],
  ['임시저장 ID', '유입 링크'],
];
const REQUIRED = ['name','email','phone','nationality','region','status','topik','field1','field2','attend','motivation','video','hanwork','source','consent','media'];
const DRAFT_REQUIRED = ['name', 'nationality', 'email', 'phone'];      // 임시저장에 필요한 항목
const FORM_KEYS =MAP.map(m => m[0]).filter(k => ['timestamp', 'state', 'draftId'].indexOf(k) < 0);
const DRAFT_ID_RE = /^[A-Za-z0-9_-]{16,64}$/;

/** 최초 1회 실행: 권한 승인 + 새 열 준비 + 기존 신청 행 '제출 완료' 표시 + 임시저장 행 회색 표시 */
function setup() {
  const sh = getSheet_();
  ensureColumns_(sh);
  CacheService.getScriptCache().remove('colmap');                       // 열 위치 기억 초기화
  const cols = colMap_(sh);
  // 상태가 비어 있는 기존 신청 행은 모두 제출된 신청
  const last = lastDataRow_(sh, cols.timestamp);
  if (last >= 2) {
    const ts = sh.getRange(2, cols.timestamp, last - 1, 1).getValues();
    const st = sh.getRange(2, cols.state, last - 1, 1).getValues();
    sh.getRange(2, cols.state, last - 1, 1).setValues(st.map((r, i) => [r[0] || (String(ts[i][0]).trim() ? STATE_DONE : '')]));
  }
  // 임시저장 행은 회색으로
  const letter = sh.getRange(1, cols.state).getA1Notation().replace(/\d+/g, '');
  const formula = '=$' + letter + '2="' + STATE_DRAFT + '"';
  const rules = sh.getConditionalFormatRules();
  const has = rules.some(r => { const c = r.getBooleanCondition(); return c && String(c.getCriteriaValues()[0]) === formula; });
  if (!has) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula)
      .setBackground('#EEEEEE').setFontColor('#777777')
      .setRanges([sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns())]).build());
    sh.setConditionalFormatRules(rules);
  }
  Logger.log('준비 완료: ' + sh.getName() + ' / ' + JSON.stringify(cols));
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.website) return json_({ ok: true });                         // 봇 방지
    const isDraft = d.type === 'draft';                                  // 임시저장은 필수 항목 검사 없음
    const draftId = DRAFT_ID_RE.test(String(d.draftId || '')) ? String(d.draftId) : '';
    if (isDraft) {
      if (!draftId) return json_({ ok: false, error: 'invalid draftId' });
      const missing = DRAFT_REQUIRED.filter(k => !String(d[k] || '').trim());
      if (missing.length) return json_({ ok: false, error: 'missing: ' + missing.join(',') });
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) return json_({ ok: false, error: 'invalid email' });
    } else {
      const missing = REQUIRED.filter(k => !String(d[k] || '').trim());
      if (d.topik === '없음' && !String(d.korean || '').trim()) missing.push('korean');
      if (missing.length) return json_({ ok: false, error: 'missing: ' + missing.join(',') });
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) return json_({ ok: false, error: 'invalid email' });
    }

    // 저장할 값을 미리 만들어 둠 (잠금 시간을 줄이기 위해)
    const values = {};
    MAP.forEach(([key]) => {
      if (key === 'timestamp') { values[key] = new Date(); return; }
      if (key === 'state') { values[key] = isDraft ? STATE_DRAFT : STATE_DONE; return; }
      if (key === 'draftId') { values[key] = draftId; return; }
      let v = String(d[key] || '').slice(0, 5000);
      if (/^[=+\-@]/.test(v)) v = "'" + v;                               // 수식 삽입 방지
      else if (/^0\d+$/.test(v)) v = "'" + v;                            // 010… 앞자리 0 유지
      values[key] = v;
    });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = getSheet_();
      const cols = colMap_(sh);                                          // {key: 열 번호(1부터)}
      // 같은 임시저장 ID의 '임시저장' 행이 있으면 그 행을 덮어씀 (제출하면 그 행이 '제출 완료'가 됨)
      let target = draftId && cols.draftId ? findDraftRow_(sh, cols, draftId, true) : 0;
      if (isDraft && !target && draftId && findDraftRow_(sh, cols, draftId, false)) return json_({ ok: false, error: 'already submitted' });
      if (!target) {
        target = lastDataRow_(sh, cols.timestamp) + 1;
        if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
      }
      // 붙어 있는 열끼리 묶어서 한 번에 기록 → 시트 쓰기 횟수 최소화 (보통 1~3번)
      const list = Object.keys(cols).map(k => [cols[k], values[k]]).sort((a, b) => a[0] - b[0]);
      let start = 0;
      for (let i = 1; i <= list.length; i++) {
        if (i === list.length || list[i][0] !== list[i - 1][0] + 1) {
          const run = list.slice(start, i);
          sh.getRange(target, run[0][0], 1, run.length).setValues([run.map(x => x[1])]);
          start = i;
        }
      }
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 이어쓰기 링크로 불러오기: GET ?draft=ID → 작성 내용 */
function doGet(e) {
  try {
    const id = String((e && e.parameter && e.parameter.draft) || '');
    if (!DRAFT_ID_RE.test(id)) return json_({ ok: false, error: 'invalid draftId' });
    const sh = getSheet_();
    const cols = colMap_(sh);
    if (!cols.draftId) return json_({ ok: false, error: 'not found' });
    const row = findDraftRow_(sh, cols, id, true) || findDraftRow_(sh, cols, id, false);
    if (!row) return json_({ ok: false, error: 'not found' });
    const range = sh.getRange(row, 1, 1, sh.getLastColumn());
    const shown = range.getDisplayValues()[0];
    const data = {};
    FORM_KEYS.forEach(k => { data[k] = cols[k] ? shown[cols[k] - 1] : ''; });
    const ts = range.getValues()[0][cols.timestamp - 1];
    return json_({
      ok: true, data: data, state: shown[cols.state - 1],
      updatedAt: ts instanceof Date ? Utilities.formatDate(ts, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : '',
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 임시저장 ID로 행 찾기. draftOnly 면 '임시저장' 상태인 행만, 아니면 아무 상태나 (가장 아래 행) */
function findDraftRow_(sh, cols, id, draftOnly) {
  const n = sh.getLastRow();
  if (n < 2 || !cols.draftId) return 0;
  const cells = sh.getRange(2, cols.draftId, n - 1, 1).createTextFinder(id).matchEntireCell(true).findAll();
  for (let i = cells.length - 1; i >= 0; i--) {
    const r = cells[i].getRow();
    if (!draftOnly || String(sh.getRange(r, cols.state).getValue()) === STATE_DRAFT) return r;
  }
  return 0;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheets().find(s => s.getSheetId() === SHEET_ID) || ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('시트 탭을 찾을 수 없음: ' + SHEET_NAME);
  return sh;
}
/** 항목 → 열 번호 표. 한 번 찾으면 6시간 기억해 매번 헤더를 다시 읽지 않음 (열을 옮기면 setup() 다시 실행) */
function colMap_(sh) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('colmap');
  if (hit) return JSON.parse(hit);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const m = {};
  MAP.forEach(([key, keys]) => { const i = findCol_(headers, keys); if (i >= 0) m[key] = i + 1; });
  if (!m.timestamp) throw new Error('타임스탬프 열을 찾을 수 없음');
  cache.put('colmap', JSON.stringify(m), 21600);
  return m;
}
function findCol_(headers, keys) {
  for (const k of keys) { const i = headers.findIndex(h => h.indexOf(k) >= 0); if (i >= 0) return i; }
  return -1;
}
function ensureColumns_(sh) {
  NEW_COLUMNS.forEach(([title, after]) => {
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (headers.indexOf(title) >= 0) return;
    const a = findCol_(headers, [after]) + 1;                            // 1부터 시작하는 열 번호
    if (a > 0) sh.insertColumnAfter(a);
    const col = a > 0 ? a + 1 : sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue(title).setFontWeight('bold');
  });
}
/** 타임스탬프 열 기준 마지막 데이터 행 (중간 빈 행·수식 열에 영향받지 않게) */
function lastDataRow_(sh, tsCol) {
  const n = sh.getLastRow();
  if (n < 2 || tsCol < 1) return Math.max(n, 1);
  const vals = sh.getRange(2, tsCol, n - 1, 1).getValues();
  for (let r = vals.length - 1; r >= 0; r--) if (String(vals[r][0]).trim() !== '') return r + 2;
  return 1;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** 편집기에서 실행해 임시저장 → 불러오기 → 제출 흐름을 점검 (점검 행은 바로 지움) */
function testDraft() {
  const id = 'selftest' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  const post = o => JSON.parse(doPost({ postData: { contents: JSON.stringify(o) } }).getContent());
  const base = { type: 'draft', draftId: id, name: '[점검] 임시저장', nationality: '점검', email: 'test@example.com', phone: '01000000000' };
  Logger.log('draft without email (should fail): ' + JSON.stringify(post(Object.assign({}, base, { email: '' }))));
  Logger.log('draft 1: ' + JSON.stringify(post(base)));
  Logger.log('draft 2: ' + JSON.stringify(post(Object.assign({}, base, { region: '점검' }))));
  Logger.log('get: ' + doGet({ parameter: { draft: id } }).getContent());
  const full = { draftId: id, name: '[점검] 제출', email: 'test@example.com', phone: '01000000000', nationality: '점검', region: '점검',
    status: '구직 중', topik: '4급', field1: 'HR·조직문화', field2: '상품·서비스 기획', attend: '잘 모르겠어요', motivation: '점검',
    video: 'https://youtu.be/test', hanwork: '아직 가입하지 않았어요', source: 'Soo House', consent: '동의합니다', media: '동의합니다' };
  Logger.log('submit: ' + JSON.stringify(post(full)));
  const sh = getSheet_(), cols = colMap_(sh);
  const rows = sh.getRange(2, cols.draftId, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findAll().map(c => c.getRow());
  Logger.log('rows=' + rows.join(',') + ' state=' + rows.map(r => sh.getRange(r, cols.state).getValue()).join(',')
    + ' phone=' + rows.map(r => sh.getRange(r, cols.phone).getDisplayValue()).join(','));
  rows.reverse().forEach(r => sh.deleteRow(r));
}
