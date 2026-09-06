/**
 * 여행이력 기록 앱 - 백엔드 (Google Apps Script)
 * -------------------------------------------------
 * 조경수조사앱과 동일한 구조입니다:
 *  - index.html은 이 스크립트와 분리되어 별도(GitHub Pages 등)에 호스팅되고,
 *    fetch()로 이 웹앱 URL을 호출해서 데이터를 주고받습니다 (google.script.run 아님).
 *  - doGet(?action=...): 조회 전용
 *  - doPost(type: ...): 생성/수정/삭제
 *  - 사진은 base64로 전달받아 구글 드라이브에 저장 후 링크만 시트에 기록합니다.
 *
 * v2.20 - 공용 백엔드 + 그룹 로그인 지원
 * -------------------------------------------------
 * 이 스프레드시트 하나를 여러 가족/그룹이 함께 쓸 수 있도록, "Groups" 시트에 그룹을
 * 등록해두면 그룹ID+멤버명으로 로그인하고, 그룹별로 여행/기록이 분리되어 보입니다.
 * Groups 시트에 아무 행도 없으면(=혼자 쓰는 경우) 로그인 없이 예전처럼 그대로 동작합니다.
 *
 * Groups 시트 구성 (직접 한 줄씩 추가):
 *   ID(그룹ID) | 그룹명 | 멤버목록(콤마로 구분) | 등록일시
 * 예) hong-family | 홍씨네 가족여행 | 아빠,엄마,첫째 |
 *
 * Code.gs를 수정한 뒤에는 반드시
 * [배포 → 배포 관리 → 편집(연필) → 새 버전으로 배포] 를 다시 실행해야 반영됩니다.
 */

// ===== 설정 =====
const SCRIPT_VERSION = '2.22.0'; // 프론트엔드 index.html의 APP_VERSION과 비교해 설정 탭에 표시됨

const PHOTO_FOLDER_NAME = '여행이력_사진';
const TRIPS_SHEET = 'Trips';
const LEGS_SHEET = 'Legs';
const GROUPS_SHEET = 'Groups';

const TRIP_HEADERS = ['ID', '그룹ID', '제목', '시작일', '종료일', '동행자', '예산', '만족도', '전체메모', '작성자', '등록일시'];
const LEG_HEADERS = ['ID', '그룹ID', 'TripID', '날짜', '출발시간', '도착시간', '출발지', '도착지', '교통수단', '숙소유형', '숙소명', '음식유형', '음식명', '실지출', '이동비', '숙박비', '식대', '커피', '기타비', '메모', '사진링크', '작성자', '위도', '경도', '등록일시'];
const GROUP_HEADERS = ['ID', '그룹명', '멤버목록', '등록일시'];

// ===== 진입점 =====
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : null;
  const groupId = e && e.parameter ? (e.parameter.groupId || '') : '';
  try {
    if (action === 'data') {
      return jsonOut(Object.assign({ ok: true, version: SCRIPT_VERSION }, getAllData_(groupId)));
    }
    if (action === 'groups') {
      return jsonOut({ ok: true, groups: listGroups_(), version: SCRIPT_VERSION });
    }
    return jsonOut({ ok: true, message: '여행이력 API 정상 작동 중', version: SCRIPT_VERSION, multiTenant: hasGroups_() });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const type = body.type;
    const data = body.data || {};
    const photos = body.photos || []; // [{name, mime, base64}]

    if (type === 'group_login') return jsonOut(groupLogin_(data.groupId, data.memberName));
    if (type === 'group_register') return jsonOut(registerGroup_(data.groupId, data.groupName, data.members));

    if (type === 'trip_add') return jsonOut(addTrip_(data));
    if (type === 'trip_update') return jsonOut(updateTrip_(data));
    if (type === 'trip_delete') return jsonOut(deleteTrip_(data.id, data.groupId));

    if (type === 'leg_add') return jsonOut(addLeg_(data, photos));
    if (type === 'leg_update') return jsonOut(updateLeg_(data, photos));
    if (type === 'leg_delete') return jsonOut(deleteLeg_(data.id, data.groupId));
    if (type === 'photo_delete') return jsonOut(deletePhotos_(data.legId, data.urls || [], data.groupId));

    if (type === 'migrate') return jsonOut(migrateAll_());
    if (type === 'claim_ungrouped') return jsonOut(claimUngroupedData_(data.groupId));

    if (type === 'photo_upload') return jsonOut(uploadSinglePhoto_(data.name, data.mime, data.base64));

    throw new Error('알 수 없는 요청 유형입니다: ' + type);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 시트/스프레드시트 확보 =====
function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('SPREADSHEET_ID');
  if (savedId) {
    try { return SpreadsheetApp.openById(savedId); } catch (e) { /* 재생성 */ }
  }
  const ss = SpreadsheetApp.create('여행이력 데이터');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function getSheet_(name, headers) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatDate_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value;
}

/**
 * 출발/도착 시간 포맷 - 구글시트가 "HH:MM" 형태의 문자열을 시간(Date) 타입으로
 * 자동 인식해버리는 경우가 있어, 그 값을 다시 순수 "HH:mm" 문자열로 되돌림
 */
function formatTime_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  if (typeof value === 'string') {
    // 예전 버그로 "1899-12-29T22:02:08.000Z" 같은 전체 날짜 문자열이 그대로
    // 저장돼버린 경우도 시:분만 추출해서 보여줌
    const m = value.match(/T(\d{2}):(\d{2})/);
    if (m) return m[1] + ':' + m[2];
  }
  return value;
}

function geocodeLocation_(place) {
  if (!place) return { lat: '', lng: '' };
  try {
    const geocoder = Maps.newGeocoder();
    const result = geocoder.geocode(place);
    if (result.status === 'OK' && result.results && result.results.length > 0) {
      const loc = result.results[0].geometry.location;
      return { lat: loc.lat, lng: loc.lng };
    }
  } catch (e) { /* 좌표 없이 저장 */ }
  return { lat: '', lng: '' };
}

// ===== 그룹(다중 사용자) =====

/** Groups 시트에 등록된 행이 하나라도 있으면 "여러 그룹이 함께 쓰는 모드"로 동작 */
function hasGroups_() {
  const sheet = getSheet_(GROUPS_SHEET, GROUP_HEADERS);
  return sheet.getLastRow() >= 2;
}

function listGroups_() {
  const sheet = getSheet_(GROUPS_SHEET, GROUP_HEADERS);
  const lastRow = sheet.getLastRow();
  const result = [];
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, GROUP_HEADERS.length).getValues().forEach(function (row) {
      if (!row[0]) return;
      result.push({ id: row[0], name: row[1] });
    });
  }
  return result;
}

/** 그룹ID + 멤버명으로 로그인 (비밀번호 없음 - 가벼운 구분용) */
function groupLogin_(groupId, memberName) {
  if (!groupId) throw new Error('그룹ID를 입력해주세요.');
  const sheet = getSheet_(GROUPS_SHEET, GROUP_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('등록된 그룹이 없습니다.');
  const rows = sheet.getRange(2, 1, lastRow - 1, GROUP_HEADERS.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(groupId)) {
      const members = String(rows[i][2] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
      if (members.length > 0 && members.indexOf(memberName) === -1) {
        throw new Error('그룹에 등록되지 않은 멤버명이에요.');
      }
      if (members.length === 0 && !memberName) {
        throw new Error('멤버명을 입력해주세요.');
      }
      return { ok: true, groupId: rows[i][0], groupName: rows[i][1], members: members, version: SCRIPT_VERSION };
    }
  }
  throw new Error('그룹ID를 찾을 수 없어요.');
}

/** 새 그룹 등록 - 앱 안에서 바로 그룹을 만들 수 있게 함 (구글시트를 직접 열 필요 없음) */
function registerGroup_(groupId, groupName, membersStr) {
  const cleanId = String(groupId || '').trim();
  const cleanName = String(groupName || '').trim();
  if (!cleanId) throw new Error('그룹ID를 입력해주세요.');
  if (!cleanName) throw new Error('그룹명을 입력해주세요.');
  if (!/^[a-zA-Z0-9_-]+$/.test(cleanId)) throw new Error('그룹ID는 영문/숫자/하이픈만 사용할 수 있어요.');

  const sheet = getSheet_(GROUPS_SHEET, GROUP_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === cleanId) {
        throw new Error('이미 사용 중인 그룹ID예요. 다른 ID를 써주세요.');
      }
    }
  }

  sheet.appendRow([cleanId, cleanName, membersStr || '', new Date()]);
  return { ok: true, groupId: cleanId, groupName: cleanName, version: SCRIPT_VERSION };
}

/**
 * 그룹 기능을 쓰기 전에 만들어둔, 그룹ID가 비어있는 예전 여행/기록들을
 * 현재 로그인한 그룹으로 한 번에 연결함 (구글시트를 직접 열어 수정할 필요 없게)
 */
function claimUngroupedData_(groupId) {
  const cleanId = String(groupId || '').trim();
  if (!cleanId) throw new Error('그룹ID가 없습니다.');

  const GROUP_COL_TRIP = TRIP_HEADERS.indexOf('그룹ID') + 1;
  const GROUP_COL_LEG = LEG_HEADERS.indexOf('그룹ID') + 1;

  let tripCount = 0;
  const tripSheet = getSheet_(TRIPS_SHEET, TRIP_HEADERS);
  const tLast = tripSheet.getLastRow();
  if (tLast >= 2) {
    const range = tripSheet.getRange(2, GROUP_COL_TRIP, tLast - 1, 1);
    const values = range.getValues();
    for (let i = 0; i < values.length; i++) {
      if (!values[i][0]) { values[i][0] = cleanId; tripCount++; }
    }
    range.setValues(values);
  }

  let legCount = 0;
  const legSheet = getSheet_(LEGS_SHEET, LEG_HEADERS);
  const lLast = legSheet.getLastRow();
  if (lLast >= 2) {
    const range = legSheet.getRange(2, GROUP_COL_LEG, lLast - 1, 1);
    const values = range.getValues();
    for (let i = 0; i < values.length; i++) {
      if (!values[i][0]) { values[i][0] = cleanId; legCount++; }
    }
    range.setValues(values);
  }

  return { ok: true, tripCount: tripCount, legCount: legCount, version: SCRIPT_VERSION };
}

/**
 * 데이터 구조 점검/복구
 * -------------------------------------------------
 * 이 앱은 버전이 올라가면서 시트에 컬럼(출발시간/도착시간, 음식유형/음식명, 그룹ID 등)을
 * 추가해왔는데, 이미 저장된 예전 행들은 새 컬럼 위치에 맞춰 자동으로 밀려나지 않습니다.
 * 그 결과 예전 데이터의 값이 엉뚱한 항목(예: 위도 값이 음식명 칸에 표시)으로 보일 수 있습니다.
 *
 * 이 함수는 시트의 실제 헤더 행(1행)에 적힌 "컬럼 이름"을 기준으로 각 데이터를 찾아
 * 현재 코드가 기대하는 컬럼 순서로 안전하게 재배치합니다. 이미 최신 구조이면 아무 것도
 * 바꾸지 않고 그대로 둡니다. 여러 번 실행해도 안전합니다.
 */
function migrateAll_() {
  const tripResult = migrateSheetToHeaders_(TRIPS_SHEET, TRIP_HEADERS);
  const legResult = migrateSheetToHeaders_(LEGS_SHEET, LEG_HEADERS);
  return {
    ok: true,
    version: SCRIPT_VERSION,
    trips: tripResult,
    legs: legResult
  };
}

function migrateSheetToHeaders_(sheetName, targetHeaders) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { changed: false, message: sheetName + ' 시트 없음(정상 - 아직 데이터가 없습니다)' };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { changed: false, message: sheetName + ' 빈 시트' };

  const oldHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const isSame = oldHeaders.length === targetHeaders.length &&
    oldHeaders.every(function (h, i) { return h === targetHeaders[i]; });
  if (isSame) return { changed: false, message: sheetName + ' 이미 최신 구조입니다' };

  const oldIndexByHeader = {};
  oldHeaders.forEach(function (h, i) { if (h) oldIndexByHeader[h] = i; });

  const dataRowCount = lastRow - 1;
  let newData = [];
  if (dataRowCount > 0) {
    const oldData = sheet.getRange(2, 1, dataRowCount, oldHeaders.length).getValues();
    newData = oldData.map(function (oldRow) {
      return targetHeaders.map(function (h) {
        const idx = oldIndexByHeader[h];
        return (idx === undefined) ? '' : oldRow[idx];
      });
    });
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
  sheet.getRange(1, 1, 1, targetHeaders.length).setFontWeight('bold').setBackground('#f0f0f0');
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, targetHeaders.length).setValues(newData);
  }
  sheet.setFrozenRows(1);

  return { changed: true, message: sheetName + ' ' + dataRowCount + '행을 최신 구조로 재정렬했습니다', movedRows: dataRowCount };
}

// ===== 전체 데이터 조회 =====
function getAllData_(groupId) {
  const multiTenant = hasGroups_();
  const tripSheet = getSheet_(TRIPS_SHEET, TRIP_HEADERS);
  const legSheet = getSheet_(LEGS_SHEET, LEG_HEADERS);

  const trips = [];
  const tLast = tripSheet.getLastRow();
  if (tLast >= 2) {
    tripSheet.getRange(2, 1, tLast - 1, TRIP_HEADERS.length).getValues().forEach(function (row) {
      if (!row[0]) return;
      trips.push({
        id: row[0], groupId: row[1], title: row[2],
        startDate: formatDate_(row[3]), endDate: formatDate_(row[4]),
        companions: row[5], budget: row[6], rating: row[7],
        memo: row[8], author: row[9], createdAt: formatDate_(row[10])
      });
    });
  }

  const legs = [];
  const lLast = legSheet.getLastRow();
  if (lLast >= 2) {
    legSheet.getRange(2, 1, lLast - 1, LEG_HEADERS.length).getValues().forEach(function (row) {
      if (!row[0]) return;
      legs.push({
        id: row[0], groupId: row[1], tripId: row[2], date: formatDate_(row[3]),
        departTime: formatTime_(row[4]), arriveTime: formatTime_(row[5]),
        fromPlace: row[6], toPlace: row[7], transport: row[8],
        lodgingType: row[9], lodgingName: row[10],
        foodType: row[11], foodName: row[12],
        actualSpend: row[13],
        costTransport: row[14], costLodging: row[15], costFood: row[16], costCoffee: row[17], costEtc: row[18],
        memo: row[19], photoUrl: row[20], author: row[21],
        lat: row[22], lng: row[23], createdAt: formatDate_(row[24])
      });
    });
  }

  let filteredTrips = trips;
  let filteredLegs = legs;
  if (multiTenant) {
    filteredTrips = trips.filter(function (t) { return t.groupId === groupId; });
    filteredLegs = legs.filter(function (l) { return l.groupId === groupId; });
  }

  filteredTrips.sort(function (a, b) { return new Date(b.startDate) - new Date(a.startDate); });
  filteredLegs.sort(function (a, b) { return new Date(a.date) - new Date(b.date); });
  return { trips: filteredTrips, legs: filteredLegs };
}

// ===== 여행(Trip) CRUD =====
function addTrip_(data) {
  const sheet = getSheet_(TRIPS_SHEET, TRIP_HEADERS);
  const id = Utilities.getUuid();
  sheet.appendRow([
    id, data.groupId || '', data.title || '', data.startDate || '', data.endDate || '',
    data.companions || '', data.budget || '', data.rating || '',
    data.memo || '', data.author || '', new Date()
  ]);
  return { ok: true, id: id, version: SCRIPT_VERSION };
}

function updateTrip_(data) {
  const sheet = getSheet_(TRIPS_SHEET, TRIP_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('여행이 없습니다.');
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // ID, 그룹ID
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      assertGroupMatch_(rows[i][1], data.groupId);
      sheet.getRange(i + 2, 3, 1, 8).setValues([[
        data.title || '', data.startDate || '', data.endDate || '',
        data.companions || '', data.budget || '', data.rating || '', data.memo || '', data.author || ''
      ]]);
      return { ok: true, version: SCRIPT_VERSION };
    }
  }
  throw new Error('해당 여행을 찾을 수 없습니다.');
}

function deleteTrip_(id, groupId) {
  const tripSheet = getSheet_(TRIPS_SHEET, TRIP_HEADERS);
  const tLast = tripSheet.getLastRow();
  if (tLast >= 2) {
    const rows = tripSheet.getRange(2, 1, tLast - 1, 2).getValues(); // ID, 그룹ID
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === id) {
        assertGroupMatch_(rows[i][1], groupId);
        tripSheet.deleteRow(i + 2);
        break;
      }
    }
  }
  const legSheet = getSheet_(LEGS_SHEET, LEG_HEADERS);
  const lLast = legSheet.getLastRow();
  if (lLast >= 2) {
    const tripIds = legSheet.getRange(2, 3, lLast - 1, 1).getValues(); // TripID 컬럼
    for (let i = tripIds.length - 1; i >= 0; i--) {
      if (tripIds[i][0] === id) legSheet.deleteRow(i + 2);
    }
  }
  return { ok: true, version: SCRIPT_VERSION };
}

/** 그룹 모드일 때, 대상 행의 그룹ID와 요청자의 그룹ID가 다르면 접근을 막음 */
function assertGroupMatch_(rowGroupId, requestGroupId) {
  if (!hasGroups_()) return; // 개인 사용(그룹 미설정) 모드에서는 검사하지 않음
  if (String(rowGroupId || '') !== String(requestGroupId || '')) {
    throw new Error('다른 그룹의 기록은 수정/삭제할 수 없습니다.');
  }
}

// ===== 일자 기록(Leg) CRUD =====
function addLeg_(data, photos) {
  if (!data.tripId) throw new Error('여행을 먼저 선택하거나 만들어주세요.');
  const sheet = getSheet_(LEGS_SHEET, LEG_HEADERS);
  const id = Utilities.getUuid();
  const geo = geocodeLocation_(data.toPlace || data.fromPlace);
  const photoUrl = combinePhotoUrls_(data.photoUrl, photos);

  sheet.appendRow([
    id, data.groupId || '', data.tripId, data.date || '',
    data.departTime || '', data.arriveTime || '',
    data.fromPlace || '', data.toPlace || '', data.transport || '',
    data.lodgingType || '', data.lodgingName || '',
    data.foodType || '', data.foodName || '',
    data.actualSpend || '',
    data.costTransport || '', data.costLodging || '', data.costFood || '', data.costCoffee || '', data.costEtc || '',
    data.memo || '', photoUrl, data.author || '', geo.lat, geo.lng, new Date()
  ]);
  return { ok: true, id: id, photoUrl: photoUrl, version: SCRIPT_VERSION };
}

function updateLeg_(data, photos) {
  const sheet = getSheet_(LEGS_SHEET, LEG_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('기록이 없습니다.');
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // ID, 그룹ID
  const geo = geocodeLocation_(data.toPlace || data.fromPlace);
  const photoUrl = combinePhotoUrls_(data.photoUrl, photos);

  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      assertGroupMatch_(rows[i][1], data.groupId);
      sheet.getRange(i + 2, 4, 1, 20).setValues([[
        data.date || '', data.departTime || '', data.arriveTime || '',
        data.fromPlace || '', data.toPlace || '', data.transport || '',
        data.lodgingType || '', data.lodgingName || '',
        data.foodType || '', data.foodName || '',
        data.actualSpend || '',
        data.costTransport || '', data.costLodging || '', data.costFood || '', data.costCoffee || '', data.costEtc || '',
        data.memo || '', photoUrl, data.author || '', geo.lat, geo.lng
      ]]);
      return { ok: true, photoUrl: photoUrl, version: SCRIPT_VERSION };
    }
  }
  throw new Error('해당 기록을 찾을 수 없습니다.');
}

/**
 * 사진 업로드 - 사진 1장이 실패해도 나머지 저장은 계속 진행되도록
 * 개별 사진마다 try/catch로 감싸서 "간혹 저장 실패"의 원인 중 하나(사진 1장 오류가
 * 전체 저장을 막는 문제)를 줄임
 */
function combinePhotoUrls_(existingUrl, photos) {
  let photoUrl = existingUrl || '';
  if (photos && photos.length > 0) {
    const links = [];
    photos.forEach(function (p) {
      try {
        links.push(savePhotoToDrive_(p.name, p.mime, p.base64));
      } catch (e) {
        // 사진 1장 업로드 실패는 건너뛰고 나머지 기록 저장은 계속 진행
      }
    });
    if (links.length > 0) {
      photoUrl = (photoUrl ? photoUrl + ',' : '') + links.join(',');
    }
  }
  return photoUrl;
}

function deleteLeg_(id, groupId) {
  const sheet = getSheet_(LEGS_SHEET, LEG_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('기록이 없습니다.');
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // ID, 그룹ID
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === id) {
      assertGroupMatch_(rows[i][1], groupId);
      sheet.deleteRow(i + 2);
      return { ok: true, version: SCRIPT_VERSION };
    }
  }
  throw new Error('해당 기록을 찾을 수 없습니다.');
}

// 사진첩에서 사진 일부만 골라 삭제 - 해당 기록의 사진링크 목록에서 지정된 URL만 제거함
// (구글드라이브의 실제 원본 파일은 삭제하지 않고 그대로 둠)
function deletePhotos_(legId, urlsToRemove, groupId) {
  if (!legId) throw new Error('기록을 찾을 수 없습니다.');
  const GROUP_COL = LEG_HEADERS.indexOf('그룹ID') + 1;
  const PHOTO_COL = LEG_HEADERS.indexOf('사진링크') + 1; // 헤더 배열에서 동적으로 계산 (컬럼 추가돼도 안전)
  const sheet = getSheet_(LEGS_SHEET, LEG_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('기록이 없습니다.');
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === legId) {
      const rowIndex = i + 2;
      assertGroupMatch_(sheet.getRange(rowIndex, GROUP_COL).getValue(), groupId);
      const cell = sheet.getRange(rowIndex, PHOTO_COL);
      const current = (cell.getValue() || '').split(',').map(s => s.trim()).filter(Boolean);
      const removeSet = {};
      (urlsToRemove || []).forEach(function (u) { removeSet[u] = true; });
      const remaining = current.filter(function (u) { return !removeSet[u]; });
      cell.setValue(remaining.join(','));

      return { ok: true, photoUrl: remaining.join(','), version: SCRIPT_VERSION };
    }
  }
  throw new Error('해당 기록을 찾을 수 없습니다.');
}

// ===== 사진 저장 =====
function savePhotoToDrive_(name, mime, base64) {
  const folder = getPhotoFolder_();
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mime || 'image/jpeg', name || ('travel_' + new Date().getTime() + '.jpg'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// 사진 한 장만 독립적으로 업로드 (기록 저장과 분리 - 하나가 실패해도 다른 사진/기록에 영향 없음)
function uploadSinglePhoto_(name, mime, base64) {
  if (!base64) throw new Error('사진 데이터가 없습니다.');
  const url = savePhotoToDrive_(name, mime, base64);
  return { ok: true, url: url, version: SCRIPT_VERSION };
}

function getPhotoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('PHOTO_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) { /* 재생성 */ }
  }
  const folder = DriveApp.createFolder(PHOTO_FOLDER_NAME);
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}
