/* 상하이 가족여행 — Google Sheets 양방향 연결 (v2.1)
 * 원본 ‘시트1’ 탭은 절대 건드리지 않고 ‘대시보드 일정’·‘대시보드 숙소’·‘대시보드 좌석’ 탭만 사용합니다.
 * 스크립트 속성 FAMILY_KEY에 20자 이상의 긴 임의 암호를 먼저 설정하세요.
 *
 * v2 변경점
 *  - 시트에 직접 추가한 행에 ID가 없거나 중복되면 고유 ID를 자동 생성해 시트에 기록합니다.
 *  - 날짜·시간을 표준 형식(2026-12-06 / 09:30)으로 정규화하고 텍스트(@) 서식으로 고정해
 *    스프레드시트가 값을 자동 변형하지 않도록 합니다.
 *  - ping 작업은 연결 암호 없이 배포·권한 상태만 확인합니다(일정·숙소 데이터는 반환하지 않음).
 *
 * v2.1 변경점
 *  - 항공 좌석 배정(가는/오는 편 · 좌석 · 이름)을 ‘대시보드 좌석’ 탭과 양방향 동기화합니다.
 *
 * v2.2 변경점
 *  - 일정 항목에 주소(간체)·추천 음식·분류(관광/식당/이동/휴식/기타)를 J·K·L 열에 추가했습니다.
 *  - 기존 9열 시트를 읽을 때 헤더를 12열로 자동 이전합니다(데이터 행은 그대로 보존).
 */

const SPREADSHEET_ID = '1lpoAC1WBGzQ2WhT4crZ6sYi7cuOEuuTo9Ms03fD2RhM';
const SCHEDULE_TAB = '대시보드 일정';
const HOTEL_TAB = '대시보드 숙소';
const SEATS_TAB = '대시보드 좌석';
const HEADERS = ['ID','날짜','시간','일정','중국어 장소(간체)','이동수단','메모','구분','완료','주소(간체)','추천 음식','분류'];
const HOTEL_HEADERS = ['숙소명','중국어 숙소명','중국어 주소','메모'];
const SEAT_HEADERS = ['방향','좌석','이름'];
const TRIP_DAYS = ['2026-12-05','2026-12-06','2026-12-07','2026-12-08'];
const SERVICE = 'shanghai-family-trip-sync';
const VERSION = 'v2.2';
const FAMILY_KEY_FALLBACK = ''; // 비워 두면 스크립트 속성 FAMILY_KEY만 사용. 저장소에는 항상 비워 둡니다.

function doGet(){
  // 배포·권한 상태 확인용 ping. 일정·숙소 데이터는 포함하지 않습니다.
  return output(pingResult());
}

function output(value){
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try{
    const req = e && e.postData && e.postData.contents ? JSON.parse(String(e.postData.contents).trim()) : null;
    if(!req || typeof req.action !== 'string') throw new Error('요청 형식이 올바르지 않습니다.');
    if(req.action === 'ping') return output(pingResult());
    const secret = PropertiesService.getScriptProperties().getProperty('FAMILY_KEY') || FAMILY_KEY_FALLBACK;
    if(!secret || String(secret).length < 20 || req.key !== secret){
      throw new Error('가족 연결 암호를 확인하세요. Apps Script 스크립트 속성의 FAMILY_KEY(20자 이상) 또는 코드의 FAMILY_KEY_FALLBACK 설정이 필요합니다.');
    }
    if(['read','write','seed'].indexOf(req.action) < 0) throw new Error('지원하지 않는 작업입니다.');
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try{
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      if(req.action === 'read') return output({ok:true, ...readAll(ss)});
      const current = readAll(ss);
      if(req.action === 'seed' && current.data.items.length){
        throw new Error('시트에 이미 일정이 있습니다. 초기화하지 않고 중단했습니다. 연결 확인을 눌러 불러오세요.');
      }
      if(req.action === 'write' && req.revision !== current.revision){
        throw new Error('다른 가족이 시트를 수정했습니다. 대시보드에서 최신 일정을 불러온 뒤 다시 저장해주세요.');
      }
      const items = req.data && req.data.items;
      if(!Array.isArray(items) || items.length > 1000) throw new Error('일정 형식 또는 행 수를 확인하세요.');
      if(req.data && req.data.hotel && (!req.data.hotel.name || !req.data.hotel.zh || !req.data.hotel.address)){
        throw new Error('숙소명과 중국어 주소를 확인하세요.');
      }
      if(req.data && req.data.flights !== undefined && !validSeats(req.data.flights)){
        throw new Error('좌석 형식을 확인하세요. 가는/오는 편(outbound/inbound)의 좌석·이름 목록이 필요합니다.');
      }
      const ids = {};
      const rows = items.map(item=>{
        if(!item.id || ids[item.id] || !/^2026-12-0[5-8]$/.test(item.day) || !String(item.title||'').trim()){
          throw new Error('ID·날짜·일정 이름이 잘못되었습니다.');
        }
        ids[item.id] = true;
        const address = String(item.address||'').trim();
        const foods = String(item.foods||'').trim();
        const kind = String(item.kind||'').trim();
        if(address.length > 300 || foods.length > 500 || kind.length > 20){
          throw new Error('주소(300자)·추천 음식(500자)·분류(20자) 길이를 확인하세요.');
        }
        return [item.id,item.day,item.time,item.title,item.zh,item.transport,item.note,item.status,item.done?'TRUE':'FALSE',address,foods,kind]
          .map(v=>escapeCell(v));
      });
      writeSchedule(ss, rows, current.data.items.length);
      if(req.data.hotel) writeHotel(ss, req.data.hotel);
      if(req.data.flights) writeSeats(ss, req.data.flights);
      SpreadsheetApp.flush();
      return output({ok:true, ...readAll(ss)});
    }finally{
      if(lock.hasLock()) lock.releaseLock();
    }
  }catch(err){
    return output({ok:false, error:String((err && err.message) || err)});
  }
}

/* 배포·권한 상태만 확인합니다. 시트 데이터는 반환하지 않습니다. */
function pingResult(){
  const out = {ok:true, service:SERVICE, version:VERSION};
  try{
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    out.spreadsheet = String(ss.getName());
    out.scheduleTab = !!ss.getSheetByName(SCHEDULE_TAB);
    out.hotelTab = !!ss.getSheetByName(HOTEL_TAB);
    out.seatsTab = !!ss.getSheetByName(SEATS_TAB);
  }catch(err){
    out.ok = false;
    out.error = '스프레드시트 접근 권한을 확인해주세요: ' + String((err && err.message) || err);
  }
  return out;
}

function escapeCell(v){
  v = String(v == null ? '' : v);
  return /^[=+@]/.test(v) ? "'" + v : v; // 수식·명령 해석 방지
}

function scheduleSheet(ss){
  let sheet = ss.getSheetByName(SCHEDULE_TAB);
  if(!sheet) sheet = ss.insertSheet(SCHEDULE_TAB);
  // 기존 9열 헤더(v2.1 이전)를 12열로 자동 이전합니다. 데이터 행은 건드리지 않습니다.
  const header = sheet.getRange(1,1,1,HEADERS.length).getDisplayValues()[0].map(v=>String(v||'').trim());
  if(HEADERS.some((h,i)=>header[i] !== h)){
    sheet.getRange(1,1,1,HEADERS.length).setNumberFormat('@').setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,HEADERS.length).setBackground('#173e42').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(4,210); sheet.setColumnWidth(5,230); sheet.setColumnWidth(7,480);
    sheet.setColumnWidth(10,300); sheet.setColumnWidth(11,340); sheet.setColumnWidth(12,80);
  }
  return sheet;
}

/* 시트 → 대시보드 데이터. 읽으면서 ID·날짜·시간을 스스로 고쳐(자동 ID, 형식 정규화) 시트에 기록합니다. */
function readSheet(ss, sheet){
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return {data:{version:1,items:[]}};
  const numRows = lastRow - 1;
  const range = sheet.getRange(2,1,numRows,12);
  const display = range.getDisplayValues();
  const raw = range.getValues();
  const tz = ss.getSpreadsheetTimeZone();
  const items = [];
  const repairs = []; // {row,col,value} — 시트에 다시 써야 할 정규화 값
  const seenIds = {};
  for(let r=0; r<numRows; r++){
    const rowNo = r + 2;
    const d = display[r];
    if(!d.some(v=>String(v).trim() !== '')) continue; // 완전히 빈 행은 건너뜁니다
    // ID: 비어 있거나 중복이면 고유 ID를 만들어 시트에 기록합니다(시트에서 직접 행 추가 지원).
    let id = String(d[0]||'').trim();
    if(!id || seenIds[id]){
      id = 'sheet-' + Utilities.getUuid();
      repairs.push({row:rowNo, col:1, value:id});
    }
    seenIds[id] = true;
    // 날짜(B열): 스프레드시트가 바꿔 놓은 값이나 다양한 표기를 2026-12-06 형태로 통일합니다.
    const dayCell = String(d[1]||'').trim();
    const day = normalizeDay(d[1], raw[r][1], tz);
    if(!day){
      throw new Error('시트 ' + rowNo + '행: 날짜를 확인해주세요. 2026-12-05 ~ 2026-12-08 형식으로 입력하세요. (현재 값: "' + dayCell.slice(0,40) + '")');
    }
    if(TRIP_DAYS.indexOf(day) < 0){
      throw new Error('시트 ' + rowNo + '행: 날짜 ' + day + '은(는) 여행 기간(2026-12-05~08)이 아닙니다. 기간 안의 날짜로 수정해주세요.');
    }
    if(raw[r][1] instanceof Date || day !== dayCell) repairs.push({row:rowNo, col:2, value:day});
    // 시간(C열): 스프레드시트가 시각으로 바꾼 셀은 09:30 형태의 텍스트로 고정합니다.
    let time = String(d[2]||'').trim();
    if(raw[r][2] instanceof Date){
      time = normalizeTime(raw[r][2], tz);
      repairs.push({row:rowNo, col:3, value:time});
    }
    // 나머지 열에 시각·날짜로 변형된 셀이 있으면 화면에 보이는 텍스트 그대로 고정합니다.
    for(let c=3; c<12; c++){
      if(raw[r][c] instanceof Date) repairs.push({row:rowNo, col:c+1, value:String(d[c]||'').trim()});
    }
    items.push({
      id: id,
      day: day,
      time: time,
      title: String(d[3]||'').trim(),
      zh: String(d[4]||'').trim(),
      transport: String(d[5]||'').trim(),
      note: String(d[6]||'').trim(),
      status: String(d[7]||'').trim() || '확인 필요',
      done: isDone(d[8]),
      address: String(d[9]||'').trim(),
      foods: String(d[10]||'').trim(),
      kind: String(d[11]||'').trim()
    });
  }
  applyRepairs(sheet, numRows, repairs);
  return {data:{version:1, items: items}};
}

/* 읽은 범위를 텍스트(@) 서식으로 고정하고 정규화한 값을 다시 써서 이후 입력이 변형되지 않게 합니다. */
function applyRepairs(sheet, numRows, repairs){
  if(numRows > 0) sheet.getRange(2,1,numRows,12).setNumberFormat('@');
  for(let i=0; i<repairs.length; i++){
    sheet.getRange(repairs[i].row, repairs[i].col).setValue(repairs[i].value);
  }
  if(repairs.length) SpreadsheetApp.flush();
}

function isDone(v){
  return /^(true|1|완료|체크|done|yes|y|v|✓|✔|o|ㅇ|○)$/i.test(String(v||'').trim());
}

/* 날짜 표기를 2026-12-06 형태로 통일. 스프레드시트가 날짜로 바꾼 셀은 원본 값을 기준으로 변환합니다. */
function normalizeDay(displayValue, rawValue, timeZone){
  if(rawValue instanceof Date) return Utilities.formatDate(rawValue, timeZone, 'yyyy-MM-dd');
  const s = String(displayValue||'').trim().replace(/\s+/g,'');
  if(!s) return null;
  const m = s.match(/^2026(?:년|-|\.|\/)?12(?:월|-|\.|\/)([0-9]{1,2})(?:일)?\.?$/)
        || s.match(/^12(?:월|-|\.|\/)([0-9]{1,2})(?:일)?\.?$/)
        || s.match(/^12(?:월|-|\.|\/)([0-9]{1,2})(?:-|\.|\/)2026$/)
        || s.match(/^202612([0-9]{1,2})$/);
  if(!m) return null;
  const n = parseInt(m[1], 10);
  if(n < 1 || n > 31) return null;
  return '2026-12-' + (n < 10 ? '0' + n : String(n));
}

/* 시각 값(Date)을 09:30 또는 09:30:15 형태로 변환합니다. */
function normalizeTime(dateValue, timeZone){
  const withSeconds = Utilities.formatDate(dateValue, timeZone, 'HH:mm:ss');
  return withSeconds.endsWith(':00') ? withSeconds.slice(0,5) : withSeconds;
}

function hotelSheet(ss){
  let s = ss.getSheetByName(HOTEL_TAB);
  if(!s){
    s = ss.insertSheet(HOTEL_TAB);
    s.getRange(1,1,1,4).setValues([HOTEL_HEADERS]);
    s.setFrozenRows(1);
    s.getRange('A:D').setNumberFormat('@');
    s.getRange(1,1,1,4).setBackground('#173e42').setFontColor('#ffffff').setFontWeight('bold');
  }
  return s;
}

function readHotel(ss){
  const values = hotelSheet(ss).getRange(2,1,1,4).getDisplayValues()[0];
  const row = values.map(v=>String(v||'').trim());
  if(!row.some(Boolean)) return null;
  return {name:row[0], zh:row[1], address:row[2], note:row[3]};
}

function readAll(ss){
  const base = readSheet(ss, scheduleSheet(ss));
  const hotel = readHotel(ss);
  if(hotel) base.data.hotel = hotel;
  const seats = readSeats(ss);
  if(seats) base.data.flights = seats;
  base.revision = hash(base.data);
  return base;
}

function hash(data){
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(data), Utilities.Charset.UTF_8)
  );
}

/* 대시보드 → 시트. 사용한 범위를 지우고 텍스트 서식으로 채워 넣습니다. */
function writeSchedule(ss, rows, previousCount){
  const sheet = scheduleSheet(ss);
  const count = Math.max(rows.length, previousCount);
  if(count > 0) sheet.getRange(2,1,count,12).clearContent().setNumberFormat('@');
  if(rows.length) sheet.getRange(2,1,rows.length,12).setValues(rows).setWrap(true);
  SpreadsheetApp.flush();
}

function writeHotel(ss, hotel){
  if(!hotel || !hotel.name || !hotel.zh || !hotel.address) throw new Error('숙소명과 중국어 주소를 확인하세요.');
  const values = [hotel.name, hotel.zh, hotel.address, hotel.note || ''].map(escapeCell);
  hotelSheet(ss).getRange(2,1,1,4).setNumberFormat('@').setValues([values]).setWrap(true);
}

/* ---------- 항공 좌석 (‘대시보드 좌석’ 탭) ---------- */

function seatsSheet(ss){
  let s = ss.getSheetByName(SEATS_TAB);
  if(!s){
    s = ss.insertSheet(SEATS_TAB);
    s.getRange(1,1,1,3).setValues([SEAT_HEADERS]);
    s.setFrozenRows(1);
    s.getRange('A:C').setNumberFormat('@');
    s.getRange(1,1,1,3).setBackground('#173e42').setFontColor('#ffffff').setFontWeight('bold');
    s.setColumnWidth(1,70); s.setColumnWidth(2,90); s.setColumnWidth(3,140);
  }
  return s;
}

/* 대시보드 → 시트 저장 직전 형식 검증. 방향별 좌석 중복, 빈 값, 행 수 초과를 확인합니다. */
function validSeats(flights){
  if(!flights || typeof flights !== 'object' || Array.isArray(flights)) return false;
  for(const dir of ['outbound','inbound']){
    const seats = flights[dir] && flights[dir].seats;
    if(seats === undefined) continue;
    if(!Array.isArray(seats) || seats.length > 100) return false;
    const seen = {};
    for(const s of seats){
      const seat = String((s && s.seat) || '').trim();
      const name = String((s && s.name) || '').trim();
      if(!seat || !name || seat.length > 8 || name.length > 60 || seen[seat]) return false;
      seen[seat] = true;
    }
  }
  return true;
}

/* 시트 → 대시보드 좌석. 비어 있으면 null을 돌려주고, 잘못된 행은 행 번호와 함께 안내합니다. */
function readSeats(ss){
  const sheet = seatsSheet(ss);
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return null;
  const rows = sheet.getRange(2,1,lastRow-1,3).getDisplayValues();
  const outbound = [];
  const inbound = [];
  const seen = {};
  for(let r=0; r<rows.length; r++){
    const rowNo = r + 2;
    const row = rows[r].map(v=>String(v||'').trim());
    if(!row.some(Boolean)) continue; // 완전히 빈 행은 건너뜁니다
    const dir = row[0];
    if(dir !== '가는' && dir !== '오는'){
      throw new Error('좌석 ' + rowNo + '행: 방향을 "가는" 또는 "오는"으로 입력하세요. (현재 값: "' + dir.slice(0,20) + '")');
    }
    const seat = row[1];
    const name = row[2];
    if(!seat || !name) throw new Error('좌석 ' + rowNo + '행: 좌석과 이름을 모두 입력하세요.');
    const key = dir + ' ' + seat;
    if(seen[key]) throw new Error('좌석 ' + rowNo + '행: ' + dir + ' 편에 좌석 ' + seat + '이(가) 중복됩니다.');
    seen[key] = true;
    (dir === '가는' ? outbound : inbound).push({seat:seat, name:name});
  }
  if(!outbound.length && !inbound.length) return null;
  return {outbound:{seats:outbound}, inbound:{seats:inbound}};
}

/* 대시보드 → 시트 좌석. 사용한 범위를 지우고 텍스트 서식으로 채워 넣습니다. */
function writeSeats(ss, flights){
  if(!validSeats(flights)) throw new Error('좌석 형식을 확인하세요.');
  const rows = [];
  [['가는','outbound'],['오는','inbound']].forEach(pair=>{
    const seats = flights[pair[1]] && flights[pair[1]].seats;
    (Array.isArray(seats) ? seats : []).forEach(s=>{
      rows.push([escapeCell(pair[0]), escapeCell(String(s.seat||'').trim()), escapeCell(String(s.name||'').trim())]);
    });
  });
  const sheet = seatsSheet(ss);
  const previous = Math.max(sheet.getLastRow() - 1, 0);
  const count = Math.max(rows.length, previous);
  if(count > 0) sheet.getRange(2,1,count,3).clearContent().setNumberFormat('@');
  if(rows.length) sheet.getRange(2,1,rows.length,3).setValues(rows).setWrap(true);
  SpreadsheetApp.flush();
}
