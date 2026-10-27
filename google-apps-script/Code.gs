/**
 * 효모 발효 탐구실 — 구글 시트 제출용 Apps Script
 *
 * 구글 시트에서 [확장 프로그램 → Apps Script]로 열어 이 파일(Code.gs)과
 * Index.html을 붙여 넣고 웹 앱으로 배포한다. 설정 방법은 README.md 참고.
 *
 * 학생 제출물은 '제출' 시트에 한 사람당 한 줄로 쌓인다. 다시 제출하면 같은 줄이 바뀐다.
 * 맨 오른쪽 두 칸(원본·채점 JSON)은 웹 앱이 읽는 값이므로 고치지 않는다.
 */

const SHEET_NAME = '제출';
const LETTERS = ['A', 'B', 'C', 'D'];
const IV_NAMES = { glucose: '포도당 농도', yeast: '효모액 농도', temp: '발효 온도' };
const LEVEL_SCORES = ['10', '9', '8', '7', '6', '5', '미응시'];
const AREA_LABEL = { '2': '잘함', '1': '보통', '0': '부족' };

const HEAD = [
  '제출 시각', '이메일', '학번', '이름', '독립 변인', '탐구 문제', '가설',
  '독립 변인(서술)', '종속 변인(서술)', '통제 변인(서술)',
  '발효관 A', '발효관 B', '발효관 C', '발효관 D', '대조군 설정 까닭',
  'A 20분(mL)', 'B 20분(mL)', 'C 20분(mL)', 'D 20분(mL)',
  'A KOH 후(mL)', 'B KOH 후(mL)', 'C KOH 후(mL)', 'D KOH 후(mL)',
  '결과 분석', '결론',
  '변인 설정', '실험 설계', '결과 분석 평가', '결론 도출', '점수', '메모',
  'ID', '원본(수정 금지)', '채점(수정 금지)',
];
const COL = name => HEAD.indexOf(name) + 1;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('효모 발효 탐구실')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- 내부 도우미 ----------
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideColumns(COL('ID'), 3);
  }
  return sh;
}
function email_() { return (Session.getActiveUser().getEmail() || '').toLowerCase(); }
function isTeacher_() {
  const me = email_();
  if (!me) return false;
  const owner = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const extra = (PropertiesService.getScriptProperties().getProperty('TEACHERS') || '')
    .toLowerCase().split(/[\s,]+/).filter(String);
  return me === owner || extra.indexOf(me) >= 0;
}
function requireTeacher_() { if (!isTeacher_()) throw new Error('선생님 계정만 쓸 수 있어요.'); }
function isClosed_() { return PropertiesService.getScriptProperties().getProperty('CLOSED') === '1'; }
function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, COL('ID'), last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}
const gLabel_ = g => g == null ? '' : g === 0 ? '증류수' : '포도당 ' + g + '%';
const yLabel_ = y => y == null ? '' : y === 0 ? '증류수' : '효모액 ' + y + '%';
const roleLabel_ = r => r === 'control' ? '대조군' : r === 'exp' ? '실험군' : '미정';
function tubeText_(t) {
  if (!t || (t.g == null && t.y == null && t.T == null)) return '';
  return gLabel_(t.g) + ' + ' + yLabel_(t.y) + ', ' + (t.T == null ? '' : t.T + '℃') + ', ' + roleLabel_(t.role);
}

// ---------- 학생 ----------
function init() {
  const out = { teacher: isTeacher_(), closed: isClosed_(), mine: null, sheetUrl: null };
  const me = email_();
  if (me) {
    const sh = sheet_();
    const row = findRow_(sh, me);
    if (row) out.mine = sh.getRange(row, COL('원본(수정 금지)')).getValue() || null;
  }
  if (out.teacher) out.sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  return JSON.stringify(out);
}

function submit(json) {
  if (isClosed_()) throw new Error('선생님이 제출을 마감했어요.');
  const p = JSON.parse(json);
  const student = p.student || {};
  if (!student.no || !student.name) throw new Error('학번과 이름이 필요해요.');
  const me = email_();
  // 학교 구글 계정이면 이메일로, 이메일을 알 수 없으면 학번+이름으로 한 사람을 구분한다.
  const id = me || ('학번:' + student.no + '-' + student.name);
  p.submittedAt = new Date().toISOString();

  const r = p.results || {};
  const data = r.data || {};
  const koh = r.koh || {};
  const values = {
    '제출 시각': new Date(p.submittedAt), '이메일': me, '학번': student.no, '이름': student.name,
    '독립 변인': IV_NAMES[p.iv] || '', '탐구 문제': p.problem, '가설': p.hypothesis,
    '독립 변인(서술)': p.vInd, '종속 변인(서술)': p.vDep, '통제 변인(서술)': p.vCtrl,
    '대조군 설정 까닭': p.controlReason, '결과 분석': p.analysis, '결론': p.conclusion,
    'ID': id, '원본(수정 금지)': JSON.stringify(p),
  };
  LETTERS.forEach(L => {
    values['발효관 ' + L] = tubeText_((p.tubes || {})[L]);
    values[L + ' 20분(mL)'] = data[L] ? data[L][data[L].length - 1] : '';
    values[L + ' KOH 후(mL)'] = koh[L] != null ? koh[L] : '';
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    let row = findRow_(sh, id);
    if (!row) row = sh.getLastRow() + 1;
    // 채점 칸(변인 설정 ~ 메모, 채점 JSON)은 건드리지 않는다.
    Object.keys(values).forEach(k => {
      const v = values[k];
      sh.getRange(row, COL(k)).setValue(typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v);
    });
  } finally {
    lock.releaseLock();
  }
  return JSON.stringify({ submittedAt: p.submittedAt });
}

// ---------- 선생님 ----------
function listAll() {
  requireTeacher_();
  const sh = sheet_();
  const subs = {}, grades = {};
  const last = sh.getLastRow();
  if (last >= 2) {
    const rows = sh.getRange(2, COL('ID'), last - 1, 3).getValues();
    rows.forEach(([id, raw, grade]) => {
      if (!id || !raw) return;
      try { subs[id] = JSON.parse(raw); } catch (e) { return; }
      if (grade) { try { grades[id] = JSON.parse(grade); } catch (e) {} }
    });
  }
  return JSON.stringify({ subs, grades, closed: isClosed_() });
}

function saveGrade(id, json) {
  requireTeacher_();
  const g = JSON.parse(json);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const row = findRow_(sh, id);
    if (!row) throw new Error('이 학생의 제출을 시트에서 찾지 못했어요.');
    const areas = g.areas || {};
    sh.getRange(row, COL('변인 설정')).setValue(AREA_LABEL[areas.vars] || '');
    sh.getRange(row, COL('실험 설계')).setValue(AREA_LABEL[areas.design] || '');
    sh.getRange(row, COL('결과 분석 평가')).setValue(AREA_LABEL[areas.analysis] || '');
    sh.getRange(row, COL('결론 도출')).setValue(AREA_LABEL[areas.conclusion] || '');
    sh.getRange(row, COL('점수')).setValue(g.level == null ? '' : LEVEL_SCORES[g.level]);
    sh.getRange(row, COL('메모')).setValue(g.memo || '');
    sh.getRange(row, COL('채점(수정 금지)')).setValue(json);
  } finally {
    lock.releaseLock();
  }
  return JSON.stringify({ ok: true });
}

function setClosed(closed) {
  requireTeacher_();
  PropertiesService.getScriptProperties().setProperty('CLOSED', closed ? '1' : '0');
  return JSON.stringify({ closed: !!closed });
}
