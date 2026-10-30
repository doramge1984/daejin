/**
 * 시트 연결 키트 — 내가 만든 화면(웹앱)에서 보낸 내용을 이 구글시트에 한 줄씩 쌓습니다.
 * 대진대학교 교수 역량 강화 연수 · 오후반 (강사 이애진 · FlowDX)
 *
 * 쓰는 법
 * 1) 데이터를 쌓을 구글시트에서: 확장 프로그램 → Apps Script → 이 코드 전체를 붙여넣고 저장
 * 2) 오른쪽 위 [배포] → [새 배포] → 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 * 3) [배포] → 권한 승인 → 나오는 "웹 앱 URL"(…/exec)을 복사해 내 화면에 넣습니다.
 *
 * 하는 일: 받은 항목 이름을 1행 제목으로 자동으로 만들고, 받은 시각과 함께 한 줄씩 추가합니다.
 * 하지 않는 일: 메일 발송, 다른 파일 접근, 시트 내용 밖으로 내보내기 — 없습니다.
 */
const SHEET_NAME = '응답';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let data = {};
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      data = (e && e.parameter) ? e.parameter : {};
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    let headers = [];
    if (sh.getLastColumn() > 0) {
      headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].filter(String);
    }
    if (headers.indexOf('받은 시각') === -1) headers.unshift('받은 시각');
    Object.keys(data).forEach(function (k) {
      if (headers.indexOf(k) === -1) headers.push(k);
    });
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);

    const row = headers.map(function (h) {
      if (h === '받은 시각') return new Date();
      const v = data[h];
      if (v === undefined || v === null) return '';
      return (typeof v === 'object') ? JSON.stringify(v) : v;
    });
    sh.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// 브라우저 주소창에 웹 앱 URL을 넣었을 때 연결 확인용 메시지
function doGet() {
  return ContentService.createTextOutput('연결 성공! 이 주소(…/exec)를 내 화면(웹앱)에 넣으세요.');
}
