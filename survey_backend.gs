// ──────────────────────────────────────────────
//  강사 섭외 리서치 설문 — Apps Script 백엔드
//  배포: 확장 프로그램 > Apps Script > 배포 > 웹 앱
//  실행 계정: 나, 액세스: 모든 사용자
// ──────────────────────────────────────────────
const SHEET_NAME = '설문응답';
const ADMIN_PW   = 'day1admin';   // ← 원하는 비밀번호로 변경

// POST: 설문 응답 저장
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    // 시트가 없으면 생성 + 헤더
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        '제출시각','이름','섭외빈도','단계별시간',
        '섭외여정','후보수','페인포인트','페인포인트_기타','자유의견'
      ]);
      sheet.setFrozenRows(1);
    }

    const freq = data.requestFreq
      ? `${data.requestFreq.count}${data.requestFreq.unit === 'week' ? '건/주' : '건/월'}`
      : '';

    sheet.appendRow([
      new Date().toLocaleString('ko-KR'),
      data.name              || '',
      freq,
      JSON.stringify(data.timeBreakdown || {}),
      (data.journey         || []).join(' → '),
      data.candidateCount   || '',
      (data.painPoints      || []).join(', '),
      data.painPoints_기타  || '',
      data.freeText         || '',
    ]);

    return ok({saved: true});
  } catch (err) {
    return ok({error: err.message});
  }
}

// GET: 응답 목록 조회 (비밀번호 필요)
function doGet(e) {
  if (e.parameter.pw !== ADMIN_PW) {
    return ok({error: 'unauthorized'});
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return ok([]);

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return ok([]);

  const records = rows.slice(1).map(r => ({
    submittedAt     : r[0],
    name            : r[1],
    requestFreq     : r[2],
    timeBreakdown   : safeJson(r[3]),
    journey         : r[4] ? r[4].split(' → ') : [],
    candidateCount  : r[5],
    painPoints      : r[6] ? r[6].split(', ') : [],
    'painPoints_기타': r[7],
    freeText        : r[8],
  }));

  return ok(records);
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeJson(s) {
  try { return JSON.parse(s); } catch(_) { return {}; }
}
