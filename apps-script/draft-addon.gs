/* ===== 임시저장 (신청서 작성 중 저장 → 이어쓰기 링크) =====
 * 신청 페이지의 [임시저장하고 링크 복사하기] → doPost(type: 'draft') → '02-1_참가 신청 임시저장' 탭에 임시저장 ID 기준으로 덮어씀
 * 이어쓰기 링크(?draft=ID)로 들어오면 doGet 이 저장된 내용을 돌려줌. 최종 제출하면 해당 행 상태가 '제출 완료'로 바뀜.
 * 탭이 없으면 처음 저장할 때 신청 탭 오른쪽에 자동으로 만듦. 탭 이름을 바꿔도 탭 ID로 찾음.
 */
const DRAFT_SHEET_NAME = '02-1_참가 신청 임시저장';
const APPLY_URL = 'https://hanpass.soohouse.app/apply/';
// [항목, 열 이름] — 열 순서를 바꿔도 열 이름으로 찾음
const DRAFT_COLS = [
  ['draftId', '임시저장 ID'],
  ['state', '상태'],
  ['createdAt', '최초 저장'],
  ['updatedAt', '최근 저장'],
  ['saves', '저장 횟수'],
  ['link', '이어쓰기 링크'],
  ['name', '이름'],
  ['email', '이메일'],
  ['phone', '휴대전화 번호'],
  ['nationality', '국적'],
  ['region', '사는 지역'],
  ['visa', '비자'],
  ['status', '현재 신분'],
  ['topik', 'TOPIK'],
  ['korean', '한국어 수준'],
  ['field1', '1순위 분야'],
  ['field2', '2순위 분야'],
  ['attend', '참여 가능 여부'],
  ['motivation', '지원 동기'],
  ['video', '영상 링크'],
  ['hanwork', 'HanWork 가입'],
  ['source', '알게 된 경로'],
  ['consent', '개인정보 동의'],
  ['media', '촬영·결과물 동의'],
  ['news', '소식 수신'],
  ['inflow', '유입 링크'],
];
const DRAFT_FIELDS = DRAFT_COLS.slice(6).map(c => c[0]);
const DRAFT_ID_RE = /^[A-Za-z0-9_-]{16,64}$/;

function saveDraft_(d) {
  const id = String(d.draftId || '');
  if (!DRAFT_ID_RE.test(id)) return json_({ ok: false, error: 'invalid draftId' });
  const values = {};
  DRAFT_FIELDS.forEach(k => {
    let v = String(d[k] || '').slice(0, 5000);
    if (/^[=+\-@]/.test(v)) v = "'" + v;                                  // 수식 삽입 방지
    values[k] = v;
  });
  if (!DRAFT_FIELDS.some(k => k !== 'inflow' && values[k])) return json_({ ok: false, error: 'empty' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getDraftSheet_();
    const cols = draftCols_(sh);
    const row = findDraftRow_(sh, cols, id);
    const now = nowText_();
    let createdAt = now, saves = 1;
    if (row) {
      createdAt = String(sh.getRange(row, cols.createdAt).getDisplayValue()) || now;
      saves = (Number(sh.getRange(row, cols.saves).getValue()) || 0) + 1;
    }
    Object.assign(values, {
      draftId: id, state: '작성 중', createdAt: createdAt, updatedAt: now, saves: saves,
      link: APPLY_URL + '?draft=' + id + '#apply',
    });
    const target = row || Math.max(sh.getLastRow(), 1) + 1;
    // 붙어 있는 열끼리 묶어서 한 번에 기록 (직원이 추가한 메모 열 등은 건드리지 않음)
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
}

/** 이어쓰기 링크로 불러오기: GET ?draft=ID */
function doGet(e) {
  try {
    const id = String((e && e.parameter && e.parameter.draft) || '');
    if (!DRAFT_ID_RE.test(id)) return json_({ ok: false, error: 'invalid draftId' });
    const sh = getDraftSheet_();
    const cols = draftCols_(sh);
    const row = findDraftRow_(sh, cols, id);
    if (!row) return json_({ ok: false, error: 'not found' });
    const r = sh.getRange(row, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    const data = {};
    DRAFT_FIELDS.forEach(k => { data[k] = r[cols[k] - 1] || ''; });
    return json_({ ok: true, data: data, state: r[cols.state - 1], updatedAt: r[cols.updatedAt - 1] });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 최종 제출 후 임시저장 행 표시 (실패해도 신청 접수에는 영향 없음) */
function markDraftSubmitted_(id) {
  try {
    if (!DRAFT_ID_RE.test(String(id))) return;
    const sh = getDraftSheet_();
    const cols = draftCols_(sh);
    const row = findDraftRow_(sh, cols, String(id));
    if (row) sh.getRange(row, cols.state).setValue('제출 완료');
  } catch (err) {
    console.error(err);
  }
}

function getDraftSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const sid = Number(props.getProperty('draftSheetId'));
  let sh = (sid && ss.getSheets().find(s => s.getSheetId() === sid)) || ss.getSheetByName(DRAFT_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(DRAFT_SHEET_NAME, getSheet_().getIndex());       // 신청 탭 바로 오른쪽
    sh.getRange(1, 1, 1, DRAFT_COLS.length).setValues([DRAFT_COLS.map(c => c[1])]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, sh.getMaxRows() - 1, DRAFT_COLS.length).setNumberFormat('@');   // 전화번호 앞 0 유지
  }
  if (String(sh.getSheetId()) !== props.getProperty('draftSheetId')) props.setProperty('draftSheetId', String(sh.getSheetId()));
  return sh;
}
/** 임시저장 탭 항목 → 열 번호(1부터). 없는 열은 맨 오른쪽에 만듦 */
function draftCols_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  let next = sh.getLastColumn() + 1;
  const m = {};
  DRAFT_COLS.forEach(([key, title]) => {
    const i = headers.indexOf(title);
    if (i >= 0) { m[key] = i + 1; return; }
    if (next > sh.getMaxColumns()) sh.insertColumnAfter(sh.getMaxColumns());
    sh.getRange(1, next).setValue(title).setFontWeight('bold');
    m[key] = next++;
  });
  return m;
}
function findDraftRow_(sh, cols, id) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const cell = sh.getRange(2, cols.draftId, n - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : 0;
}
function nowText_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');   // 한국 시간
}

/** 편집기에서 한 번 실행해 임시저장 저장·불러오기를 점검 (점검 행은 바로 지움) */
function testDraft() {
  const id = 'selftest_' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  Logger.log(saveDraft_({ draftId: id, name: '점검', phone: '010-0000-0000', inflow: '점검' }).getContent());
  Logger.log(doGet({ parameter: { draft: id } }).getContent());
  markDraftSubmitted_(id);
  const sh = getDraftSheet_();
  const row = findDraftRow_(sh, draftCols_(sh), id);
  Logger.log('state=' + sh.getRange(row, draftCols_(sh).state).getValue());
  sh.deleteRow(row);
}
