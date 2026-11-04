/* 상하이 가족여행 — Google Sheets 연결
 * 원본 탭은 건드리지 않고 ‘대시보드 일정’ 탭만 사용합니다.
 * 스크립트 속성 FAMILY_KEY에 긴 임의 암호를 먼저 설정하세요.
 */
const SPREADSHEET_ID = '1lpoAC1WBGzQ2WhT4crZ6sYi7cuOEuuTo9Ms03fD2RhM';
const TAB = '대시보드 일정';
const HEADERS = ['ID','날짜','시간','일정','중국어 장소(간체)','이동수단','메모','구분','완료'];
function doGet(){return output({ok:false,error:'대시보드 연결은 인증된 POST 요청만 받습니다.'});}
function output(value){return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);}
function doPost(e){
 const lock=LockService.getScriptLock();
 try{
  const req=JSON.parse(e.postData.contents);
  const secret=PropertiesService.getScriptProperties().getProperty('FAMILY_KEY');
  if(!secret||secret.length<20||req.key!==secret)throw new Error('가족 연결 암호를 확인하세요. FAMILY_KEY는 20자 이상이어야 합니다.');
  if(!['read','write','seed'].includes(req.action))throw new Error('지원하지 않는 작업입니다.');
  lock.waitLock(20000);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet=ss.getSheetByName(TAB);
  if(!sheet){sheet=ss.insertSheet(TAB);sheet.getRange(1,1,1,9).setValues([HEADERS]);sheet.setFrozenRows(1);sheet.getRange('A:I').setNumberFormat('@');sheet.getRange('A1:I1').setBackground('#173e42').setFontColor('#ffffff').setFontWeight('bold');sheet.setColumnWidth(4,210);sheet.setColumnWidth(5,230);sheet.setColumnWidth(7,480);}
  const current=readAll(ss,sheet);
  if(req.action==='read')return output({ok:true,...current});
  if(req.action==='seed'&&current.data.items.length)throw new Error('시트에 이미 일정이 있습니다. 초기화하지 않고 중단했습니다. 연결 확인을 눌러 불러오세요.');
  if(req.action==='write'&&req.revision!==current.revision)throw new Error('다른 가족이 시트를 수정했습니다. 최신 일정을 불러와주세요.');
  if(req.data&&req.data.hotel&&(!req.data.hotel.name||!req.data.hotel.zh||!req.data.hotel.address))throw new Error('숙소명과 중국어 주소를 확인하세요.');
  const items=req.data&&req.data.items;
  if(!Array.isArray(items)||items.length>1000)throw new Error('일정 형식 또는 행 수를 확인하세요.');
  const ids={};
  const rows=items.map(i=>{
   if(!i.id||ids[i.id]||!/^2026-12-0[5-8]$/.test(i.day)||!String(i.title||'').trim())throw new Error('ID·날짜·일정 이름이 잘못되었습니다.');
   ids[i.id]=true;
   return [i.id,i.day,i.time,i.title,i.zh,i.transport,i.note,i.status,i.done?'TRUE':'FALSE'].map(v=>String(v==null?'':v));
  });
  // Text formatting prevents times/dates from being silently converted; leading '=' is escaped.
  const safe=rows.map(row=>row.map(v=>/^[=+@]/.test(v)?"'"+v:v));
  const count=Math.max(rows.length,current.data.items.length);
  if(count){sheet.getRange(2,1,count,9).clearContent().setNumberFormat('@');}
  if(safe.length)sheet.getRange(2,1,safe.length,9).setValues(safe).setWrap(true);
  SpreadsheetApp.flush();
  if(req.data.hotel){writeHotel(ss,req.data.hotel);SpreadsheetApp.flush();}
  return output({ok:true,...readAll(ss,sheet)});
 }catch(err){return output({ok:false,error:String(err.message||err)});}
 finally{if(lock.hasLock())lock.releaseLock();}
}
function readSheet(sheet){
 const values=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,9).getDisplayValues():[];
 const items=values.filter(r=>r.some(Boolean)).map((r,index)=>{
  if(!r[0])throw new Error('시트 '+(index+2)+'행에 ID가 없습니다. 새 일정은 대시보드에서 추가하거나 고유 ID를 입력하세요.');
  return {id:r[0],day:r[1],time:r[2],title:r[3],zh:r[4],transport:r[5],note:r[6],status:r[7]||'확인 필요',done:/^(true|1|완료)$/i.test(r[8])};
 });
 const revision=Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(items),Utilities.Charset.UTF_8));
 return {data:{version:1,items},revision};
}

function hotelSheet(ss){let s=ss.getSheetByName('대시보드 숙소');if(!s){s=ss.insertSheet('대시보드 숙소');s.getRange(1,1,1,4).setValues([['숙소명','중국어 숙소명','중국어 주소','메모']]);s.setFrozenRows(1);}return s;}
function readAll(ss,sheet){const base=readSheet(sheet);const h=hotelSheet(ss).getRange(2,1,1,4).getDisplayValues()[0];if(h.some(Boolean))base.data.hotel={name:h[0],zh:h[1],address:h[2],note:h[3]};base.revision=Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(base.data),Utilities.Charset.UTF_8));return base;}
function writeHotel(ss,h){if(!h||!h.name||!h.zh||!h.address)throw new Error('숙소명과 중국어 주소를 확인하세요.');const values=[h.name,h.zh,h.address,h.note||''].map(v=>{v=String(v);return /^[=+@]/.test(v)?"'"+v:v;});hotelSheet(ss).getRange(2,1,1,4).setNumberFormat('@').setValues([values]).setWrap(true);}
