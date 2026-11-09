/**
 * 본부 워크샵 일정 투표 — 응답 저장용 Google Apps Script
 * 이 코드를 Google 스프레드시트의 [확장 프로그램 > Apps Script]에 붙여 넣고
 * 웹 앱으로 배포하세요. 설정 방법은 README.md 참고.
 *
 * 2차 투표: 후보 날짜 3개 중에서만 선택. 1차 응답(`응답` 시트)은 그대로 보존합니다.
 */
const SHEET_NAME = '2차 응답';
const CANDIDATES = ['2026-12-03', '2026-12-17', '2026-12-18'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['이름', '참석 가능 날짜', '수정 시각']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 공개 응답에는 이름을 빼고 날짜만 내려줍니다
function publicVotes_() {
  const rows = sheet_().getDataRange().getDisplayValues().slice(1);
  return rows
    .filter(r => r[0])
    .map(r => ({
      dates: r[1] ? r[1].split(',').map(s => s.trim()).filter(d => CANDIDATES.indexOf(d) >= 0) : []
    }));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ok: true, votes: publicVotes_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const name = String(body.name || '').trim().slice(0, 20);
    if (!name) return json_({ ok: false, error: '이름이 비어 있습니다.' });

    const dates = (body.dates || []).filter(d => CANDIDATES.indexOf(d) >= 0);
    if (!dates.length) return json_({ ok: false, error: '날짜를 하나 이상 선택해 주세요.' });

    const sh = sheet_();
    const names = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getDisplayValues().map(r => r[0]);
    let row = names.indexOf(name) + 1;           // 같은 이름이면 그 행을 덮어씀
    if (row < 2) row = sh.getLastRow() + 1;

    const range = sh.getRange(row, 1, 1, 3);
    range.setNumberFormat('@');                   // 날짜 문자열이 자동 변환되지 않도록 텍스트로 저장
    range.setValues([[
      name,
      dates.join(', '),
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm')
    ]]);

    return json_({ ok: true, votes: publicVotes_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
