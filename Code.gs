/**
 * 한국어 읽기 진단 (1차 + 2차) — 구글 시트 저장용 Apps Script
 *
 * 1. 아래 TEACHER_KEY를 선생님만 아는 비밀번호로 바꾸세요.
 * 2. 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 3. 나온 웹 앱 주소(https://script.google.com/macros/s/.../exec)를
 *    index.html 맨 위 SCRIPT_URL에 붙여 넣으세요.
 *
 * 코드를 바꾼 뒤에는: 배포 > 배포 관리 > 연필(수정) > 버전: 새 버전 > 배포  (주소는 그대로예요)
 *
 * 아이들의 제출은 키 없이 저장되고, 목록 보기와 녹음 기록 저장은 TEACHER_KEY가 있어야 돼요.
 */
const TEACHER_KEY = '여기에-비밀번호';

const TABS = {
  results: {
    name: '결과',
    headers: ['제출 시각','ID','이름','판정','어휘 점수','문장 점수','전체 점수','듣기 순서','짧은 글 읽은 시간(초)','내용 질문(/2)','어휘(/5)','문장(/5)','다음 단계',
      '문장 사다리','듣기(/2)','읽기(/2)','확인 듣기(/2)','확인 읽기(/2)','이야기 들은 횟수','모르겠어요 개수','종합','듣기·읽기 해석','틀린 문항과 의미','코드']
  },
  answers: {
    name: '문항별 답',
    headers: ['제출 시각','ID','이름','R-1','R-2','B-1','B-2','B-3','B-4','B-5','B-6','B-7','B-8','B-9','B-10','L-1','L-2','L-3','L-4','H-1','H-2','H-3','H-4','H-5','H-6','H-7','H-8']
  },
  oral: {
    name: '녹음 기록',
    headers: ['기록 시각','ID','이름','틀린 낱말','낱말 (받침 없음 / 있음 / 소리 바뀜)','글 읽은 시간(초)','틀린 어절','읽는 모습','읽기 유형','종합 한 줄','메모','원자료']
  },
  results2: {
    name: '2차 결과',
    headers: ['제출 시각','ID','이름','2차 판정','교과 어휘','긴 문장','문단 읽기','교과서 지시문','전체 점수','학습 출발점','모르겠어요 개수','걸린 시간(분)','틀린 문항과 의미','코드']
  },
  answers2: {
    name: '2차 문항별 답',
    headers: ['제출 시각','ID','이름','V-1','V-2','V-3','V-4','V-5','S-1','S-2','S-3','S-4','S-5','P-1','P-2','P-3','P-4','P-5','D-1','D-2','D-3','D-4','D-5']
  }
};

function sheet_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = TABS[key];
  let sh = ss.getSheetByName(t.name);
  if (!sh) {
    sh = ss.insertSheet(t.name);
    sh.getRange(1, 1, 1, t.headers.length).setValues([t.headers]).setFontWeight('bold').setBackground('#E2EBF5');
    sh.setFrozenRows(1);
  } else {
    const cur = sh.getRange(1, 1, 1, t.headers.length).getValues()[0];
    if (cur.join('|') !== t.headers.join('|')) {
      sh.getRange(1, 1, 1, t.headers.length).setValues([t.headers]).setFontWeight('bold').setBackground('#E2EBF5');
    }
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    const now = new Date();

    if (d.action === 'submit') {
      sheet_('results').appendRow([now, d.id, d.name, d.verdict, d.vScore, d.sScore, d.total, d.version, d.readSec, d.rq, d.v, d.s, d.route,
        d.ladder, d.listen, d.read, d.listen2, d.read2, d.plays, d.skips, d.line, d.hl, d.wrong, d.code]);
      const a = d.answers || {};
      const ids = TABS.answers.headers.slice(3);
      sheet_('answers').appendRow([now, d.id, d.name].concat(ids.map(function (id) { return a[id] || ''; })));
      return json_({ ok: true });
    }

    if (d.action === 'submit2') {
      sheet_('results2').appendRow([now, d.id, d.name, d.verdict, d.V, d.S, d.P, d.D, d.total, d.start, d.skips, d.minutes, d.wrong, d.code]);
      const a2 = d.answers || {};
      const ids2 = TABS.answers2.headers.slice(3);
      sheet_('answers2').appendRow([now, d.id, d.name].concat(ids2.map(function (id) { return a2[id] || ''; })));
      return json_({ ok: true });
    }

    if (d.action === 'oral') {
      if (d.key !== TEACHER_KEY) return json_({ ok: false, error: 'key' });
      const sh = sheet_('oral');
      const row = [now, d.id, d.name, d.wrongWords, d.groups, d.sec, d.wrongE, d.style, d.type, d.line, d.note, d.raw];
      const ids = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
      const i = ids.indexOf(d.id);
      if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
      else sh.appendRow(row);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.action !== 'list') return json_({ ok: true, message: '한국어 읽기 진단 시트가 연결되어 있어요.' });
  if (p.key !== TEACHER_KEY) return json_({ ok: false, error: 'key' });

  const rs = sheet_('results');
  const results = [];
  if (rs.getLastRow() > 1) {
    rs.getRange(2, 1, rs.getLastRow() - 1, TABS.results.headers.length).getValues().forEach(function (r) {
      results.push({
        time: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
        id: r[1], name: r[2], verdict: r[3], v: r[10], s: r[11], code: r[23]
      });
    });
  }
  results.reverse();

  const os = sheet_('oral');
  const oral = {};
  if (os.getLastRow() > 1) {
    os.getRange(2, 1, os.getLastRow() - 1, TABS.oral.headers.length).getValues().forEach(function (r) {
      let raw = {};
      try { raw = JSON.parse(r[11] || '{}'); } catch (x) {}
      oral[r[1]] = { type: String(r[8] || '').charAt(0), note: r[10], raw: raw };
    });
  }
  return json_({ ok: true, results: results, oral: oral });
}
