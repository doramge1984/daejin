/**
 * 팀 주제 신청함 (교수용 템플릿)
 * 학생 팀장이 휴대폰으로 팀 번호, 팀원, 주제, 팀 사진을 등록하면
 * 이 시트의 '팀 명단' 탭에 한 사람이 한 줄로 쌓이고, 사진은 내 드라이브 폴더에 저장됩니다.
 * 주제는 먼저 신청한 팀이 가져가고, 정원이 찬 주제는 다른 팀 화면에서 '마감'으로 바뀝니다.
 *
 * 탭: 시작하기(B2 과목명, B3 팀 수, B4 주제당 팀 수, B5 사진 폴더, B6 등록 화면 주소)
 *     주제목록(A열 주제, B열 정원 — 비우면 B4 값)
 *     팀 명단(학생 등록이 쌓이는 곳)
 * 외부로 보내는 것은 없고, 학생은 이 시트를 볼 수 없습니다. 실행 계정은 '나'입니다.
 */

const TAB_SET = '시작하기';
const TAB_TOPIC = '주제목록';
const TAB_LIST = '팀 명단';
const HEAD = ['등록 시각', '팀 번호', '주제', '역할', '이름', '팀 사진', '팀 소개', '메모'];
const COLORS = ['#e8f0fe', '#fef7e0', '#e6f4ea', '#fce8e6', '#f3e8fd', '#e4f7fb'];

/* ---------- 학생 화면 ---------- */
function doGet() {
  처음준비();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(cfg_().course + ' 팀 주제 신청')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 화면이 열릴 때와 [새로고침]에서 부름: 과목명, 팀 수, 주제별 남은 자리 */
function getInfo() {
  const c = cfg_();
  return { course: c.course, teams: c.teams, topics: topicStatus_(), done: registeredTeams_() };
}

/** [등록하기]에서 부름 */
function submitTeam(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 여러 팀이 동시에 눌러도 한 팀씩 처리
  try {
    const team = Number(d.team);
    if (!team) throw new Error('팀 번호를 골라 주세요.');
    if (!d.agree) throw new Error('안내에 동의해 주세요.');
    const members = (d.members || []).filter(m => m && String(m.name || '').trim());
    if (!members.length) throw new Error('팀원을 한 명 이상 적어 주세요.');
    if (registeredTeams_().indexOf(team) > -1) {
      throw new Error(team + '팀은 이미 등록되어 있어요. 고칠 내용은 교수님께 말씀해 주세요.');
    }
    // 화면을 연 뒤 다른 팀이 먼저 가져갔을 수 있으니, 저장 직전에 한 번 더 확인
    const t = topicStatus_().filter(x => x.topic === d.topic)[0];
    if (!t) throw new Error('주제를 골라 주세요.');
    if (t.left <= 0) throw new Error('방금 다른 팀이 「' + d.topic + '」을(를) 먼저 신청했어요. [새로고침] 후 다른 주제를 골라 주세요.');

    let photo = '';
    if (d.photo && d.photo.data) {
      const blob = Utilities.newBlob(Utilities.base64Decode(d.photo.data), d.photo.type || 'image/jpeg',
        team + '팀_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd_HHmm') + '.jpg');
      const file = folder_().createFile(blob);
      photo = '=HYPERLINK("' + file.getUrl() + '","사진 보기")';
    }

    const sh = ss_().getSheetByName(TAB_LIST);
    const now = new Date();
    const rows = members.map((m, i) => [now, team, d.topic, String(m.role || ''), String(m.name).trim(),
      i === 0 ? photo : '', i === 0 ? String(d.intro || '') : '', '']);
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, HEAD.length).setValues(rows)
      .setBackground(COLORS[(team - 1) % COLORS.length]);
    return { ok: true, team: team, topic: d.topic };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 교수용: 시트 메뉴 ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('팀 주제 신청')
    .addItem('처음 준비 (탭 만들기)', '처음준비')
    .addItem('주제별 현황 보기', 'showStatus')
    .addToUi();
}

function showStatus() {
  const lines = topicStatus_().map(t => (t.left > 0 ? '○ ' : '● ') + t.topic + ' — ' +
    (t.teams.length ? t.teams.join(', ') + '팀' : '신청 없음') + (t.left > 0 ? ' (남은 자리 ' + t.left + ')' : ' (마감)'));
  SpreadsheetApp.getUi().alert('주제별 현황', lines.join('\n') || '주제목록 탭에 주제를 적어 주세요.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ---------- 내부 도우미 ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function cfg_() {
  const sh = ss_().getSheetByName(TAB_SET);
  const v = c => sh ? sh.getRange(c).getValue() : '';
  return {
    course: String(v('B2') || '우리 수업'),
    teams: Number(v('B3')) || 10,
    cap: Number(v('B4')) || 1
  };
}

function topicStatus_() {
  const cap = cfg_().cap;
  const tsh = ss_().getSheetByName(TAB_TOPIC);
  const n = tsh.getLastRow() - 1;
  const topics = n > 0 ? tsh.getRange(2, 1, n, 2).getValues().filter(r => String(r[0]).trim()) : [];
  const byTopic = {};
  listRows_().forEach(r => {
    const k = String(r[2]);
    (byTopic[k] = byTopic[k] || []);
    if (byTopic[k].indexOf(r[1]) < 0) byTopic[k].push(r[1]);
  });
  return topics.map(r => {
    const topic = String(r[0]).trim();
    const teams = byTopic[topic] || [];
    const max = Number(r[1]) || cap;
    return { topic: topic, max: max, teams: teams, left: max - teams.length };
  });
}

function registeredTeams_() {
  const s = [];
  listRows_().forEach(r => { const t = Number(r[1]); if (t && s.indexOf(t) < 0) s.push(t); });
  return s;
}

function listRows_() {
  const sh = ss_().getSheetByName(TAB_LIST);
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, HEAD.length).getValues() : [];
}

function folder_() {
  const props = PropertiesService.getDocumentProperties();
  const id = props.getProperty('FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(cfg_().course + ' 팀 사진');
  props.setProperty('FOLDER', f.getId());
  ss_().getSheetByName(TAB_SET).getRange('B5').setValue(f.getUrl());
  return f;
}

/** 탭이 없으면 만들기 (템플릿에는 이미 있음) */
function 처음준비() {
  const ss = ss_();
  if (!ss.getSheetByName(TAB_SET)) {
    const s = ss.insertSheet(TAB_SET, 0);
    s.getRange('A1:C6').setValues([
      ['항목', '값', '설명'],
      ['과목명', '[가상] 캡스톤디자인', '학생 화면 맨 위에 나오는 이름'],
      ['팀 수', 8, '팀 번호를 1번부터 몇 번까지 고르게 할지'],
      ['주제당 팀 수', 1, '한 주제를 몇 팀까지 가져갈 수 있는지 (주제목록 B열이 우선)'],
      ['사진 폴더', '(첫 사진이 올라오면 자동으로 생김)', '팀 사진이 모이는 내 드라이브 폴더'],
      ['등록 화면 주소', '(8단계에서 붙여 넣기)', '학생에게 줄 주소. 크롬에서 열고 QR로 만들기']
    ]);
    s.getRange('A1:C1').setFontWeight('bold').setBackground('#23303f').setFontColor('#ffffff');
    s.getRange('B2:B6').setBackground('#fff8e1');
    s.setColumnWidth(1, 120); s.setColumnWidth(2, 320); s.setColumnWidth(3, 360);
  }
  if (!ss.getSheetByName(TAB_TOPIC)) {
    const s = ss.insertSheet(TAB_TOPIC, 1);
    s.getRange('A1:B7').setValues([
      ['주제', '정원(팀 수, 비우면 시작하기 B4)'],
      ['캠퍼스 셔틀 대기 시간 줄이기', ''], ['학과 행사 참가 신청 자동화', ''],
      ['도서관 좌석 찾기 도우미', ''], ['지역 소상공인 홍보 페이지', ''],
      ['1인 가구 식단 기록 앱', ''], ['자유 주제', 3]
    ]);
    s.getRange('A1:B1').setFontWeight('bold').setBackground('#23303f').setFontColor('#ffffff');
    s.setColumnWidth(1, 300); s.setColumnWidth(2, 260);
  }
  if (!ss.getSheetByName(TAB_LIST)) {
    const s = ss.insertSheet(TAB_LIST, 2);
    s.getRange(1, 1, 1, HEAD.length).setValues([HEAD])
      .setFontWeight('bold').setBackground('#23303f').setFontColor('#ffffff');
    s.setFrozenRows(1);
    s.setColumnWidth(3, 220); s.setColumnWidth(7, 260); s.setColumnWidth(8, 200);
  }
}
