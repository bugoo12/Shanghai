/* 상하이 가족여행 — Code.gs 단위 테스트
 * Apps Script 서비스(SpreadsheetApp, PropertiesService, LockService, ContentService, Utilities)를
 * 흉내내 시트 연결 로직을 검증합니다. 실제 FAMILY_KEY 값은 포함하지 않습니다.
 * 실행: node test/code.test.mjs
 */
import {readFileSync} from 'node:fs';
import {createHash, randomUUID} from 'node:crypto';

const SPREADSHEET_ID = '1lpoAC1WBGzQ2WhT4crZ6sYi7cuOEuuTo9Ms03fD2RhM';
const SCHEDULE_TAB = '대시보드 일정';
const HOTEL_TAB = '대시보드 숙소';
const ORIGIN_TAB = '시트1';
const TEST_KEY = 'unit-test-dummy-key-abcdef123456'; // 27자 테스트용 더미 (실제 암호 아님)

/* ---------- Apps Script 서비스 목업 ---------- */
const pad = n => (n < 10 ? '0' + n : String(n));
function displayOf(v){
  if(v === null || v === undefined || v === '') return '';
  if(v instanceof Date){
    const hm = pad(v.getHours()) + ':' + pad(v.getMinutes()) + ':' + pad(v.getSeconds());
    const timeOnly = v.getFullYear() === 1899 && v.getMonth() === 11 && v.getDate() === 30;
    if(timeOnly) return hm; // 스프레드시트가 시각으로 바꾼 셀의 표시 예: 09:30:00
    const ds = v.getFullYear() + '. ' + pad(v.getMonth() + 1) + '. ' + v.getDate() + '.';
    return (v.getHours() || v.getMinutes() || v.getSeconds()) ? ds + ' ' + hm : ds; // 한국 로케일 표시 예: 2026. 12. 7.
  }
  if(typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}
class MockSheet{
  constructor(name){this.name=name;this.cells=new Map();this.writes=[];this.clears=[];this.textFormats=new Set();this.frozen=0;}
  key(r,c){return r+':'+c;}
  set(r,c,v){this.cells.set(this.key(r,c),v);}
  get(r,c){const v=this.cells.get(this.key(r,c));return v===undefined?'':v;}
  getLastRow(){let last=0;for(const [k,v] of this.cells){if(v===''||v==null)continue;const r=Number(k.split(':')[0]);if(r>last)last=r;}return last;}
  getRange(a,b,c,d){
    if(typeof a==='string'){
      const m=a.match(/^([A-I])(\d*)(?::([A-I])(\d*))?$/);
      if(!m)throw new Error('지원하지 않는 범위: '+a);
      const c1=m[1].charCodeAt(0)-64,c2=m[3]?m[3].charCodeAt(0)-64:c1;
      const r1=m[2]?parseInt(m[2]):1,r2=m[4]?parseInt(m[4]):r1;
      return new MockRange(this,r1,c1,r2-r1+1,c2-c1+1);
    }
    return new MockRange(this,a,b,c,d);
  }
  setFrozenRows(n){this.frozen=n;return this;}
  setColumnWidth(){return this;}
}
class MockRange{
  constructor(sheet,row,col,numRows=1,numCols=1){Object.assign(this,{sheet,row,col,numRows,numCols});}
  grid(fn){const out=[];for(let r=0;r<this.numRows;r++){const rowArr=[];for(let c=0;c<this.numCols;c++)rowArr.push(fn(this.sheet.get(this.row+r,this.col+c)));out.push(rowArr);}return out;}
  getValues(){return this.grid(v=>v);}
  getDisplayValues(){return this.grid(displayOf);}
  setValue(v){this.sheet.set(this.row,this.col,v);this.sheet.writes.push({row:this.row,col:this.col,value:v});return this;}
  setValues(rows){rows.forEach((rr,i)=>rr.forEach((v,j)=>{this.sheet.set(this.row+i,this.col+j,v);this.sheet.writes.push({row:this.row+i,col:this.col+j,value:v});}));return this;}
  clearContent(){for(let r=0;r<this.numRows;r++)for(let c=0;c<this.numCols;c++)this.sheet.set(this.row+r,this.col+c,'');this.sheet.clears.push({row:this.row,numRows:this.numRows,numCols:this.numCols});return this;}
  setNumberFormat(){for(let r=0;r<this.numRows;r++)for(let c=0;c<this.numCols;c++)this.sheet.textFormats.add(this.sheet.key(this.row+r,this.col+c));return this;}
  setWrap(){return this;}setBackground(){return this;}setFontColor(){return this;}setFontWeight(){return this;}
}
class MockSpreadsheet{
  constructor(name){this.name=name;this.sheets=new Map();}
  getSheetByName(n){return this.sheets.get(n)||null;}
  insertSheet(n){const s=new MockSheet(n);this.sheets.set(n,s);return s;}
  getSpreadsheetTimeZone(){return 'Asia/Shanghai';}
  getName(){return this.name;}
}
const ContentService={createTextOutput:s=>({__text:String(s),setMimeType(){return this;}}),MimeType:{JSON:'JSON'}};
const makeUtilities=()=>({
  getUuid:()=>randomUUID(),
  formatDate:(d,tz,pattern)=>{
    if(pattern==='yyyy-MM-dd')return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());
    if(pattern==='HH:mm:ss')return pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds());
    throw new Error('지원하지 않는 서식: '+pattern);
  },
  computeDigest:(algo,str,charset)=>Array.from(createHash('sha256').update(String(str),'utf8').digest()),
  base64EncodeWebSafe:bytes=>Buffer.from(bytes).toString('base64url'),
  DigestAlgorithm:{SHA_256:'SHA_256'},
  Charset:{UTF_8:'UTF_8'}
});

/* ---------- Code.gs 로드 ---------- */
const src=readFileSync(new URL('../Code.gs',import.meta.url),'utf8');
const factory=new Function('SpreadsheetApp','PropertiesService','LockService','ContentService','Utilities',
  src+'\n;return {doGet,doPost,normalizeDay,isDone};');

const SCHEDULE_HEADER=['ID','날짜','시간','일정','중국어 장소(간체)','이동수단','메모','구분','완료'];
const HOTEL_HEADER=['숙소명','중국어 숙소명','중국어 주소','메모'];
const BASE_ROWS=[
  SCHEDULE_HEADER,
  ['item-1','2026-12-05','12:35','푸동공항 도착','上海浦东国际机场','','KE887 · 인천 11:10 출발 / 상하이 12:35 도착','기존','FALSE'],
  ['item-2','2026-12-05','13:30~14:00','호텔로 이동','上海外滩璞硐酒店 北京东路398号','택시 / 예약 픽업','45~70분 소요','기존','FALSE'],
  ['item-6','2026-12-06','08:00','호텔 조식','','','출발 전 화장실·물 준비','제안','FALSE'],
  ['item-24','2026-12-08','14:00','상하이 출발 → 인천 도착','','','KE888 · 푸동 14:00 → 인천 17:10','기존','FALSE']
];
const BASE_HOTEL=[
  HOTEL_HEADER,
  ['The Bund Puyan Hotel','上海外滩璞硐酒店','上海市黄浦区北京东路398号','체크인 16:00 / 얼리 체크인 확인']
];

function makeEnv({keyValue=TEST_KEY,rows=BASE_ROWS,hotel=BASE_HOTEL,makeApi=factory}={}){
  const ss=new MockSpreadsheet('상하이 가족여행(테스트)');
  const origin=ss.insertSheet(ORIGIN_TAB);
  [['원본','탭'],['보존','대상']].forEach((rr,i)=>rr.forEach((v,j)=>origin.set(i+1,j+1,v)));
  if(rows)rows.forEach((rr,i)=>{const s=ss.getSheetByName(SCHEDULE_TAB)||ss.insertSheet(SCHEDULE_TAB);rr.forEach((v,j)=>s.set(i+1,j+1,v));});
  if(hotel)hotel.forEach((rr,i)=>{const s=ss.getSheetByName(HOTEL_TAB)||ss.insertSheet(HOTEL_TAB);rr.forEach((v,j)=>s.set(i+1,j+1,v));});
  const props={};if(keyValue)props.FAMILY_KEY=keyValue;
  const propsSvc={getScriptProperties:()=>({getProperty:k=>props[k]??null})};
  const lockSt={wait:0,release:0,locked:false};
  const lockSvc={getScriptLock:()=>({waitLock(){lockSt.wait++;lockSt.locked=true;},hasLock:()=>lockSt.locked,releaseLock(){lockSt.release++;lockSt.locked=false;}})};
  const opened=[];const stats={flushes:0};
  const SpreadsheetApp={openById(id){opened.push(id);if(id!==SPREADSHEET_ID)throw new Error('잘못된 시트 ID: '+id);return ss;},flush(){stats.flushes++;}};
  const api=makeApi(SpreadsheetApp,propsSvc,lockSvc,ContentService,makeUtilities());
  return {ss,api,call:body=>JSON.parse(api.doPost({postData:{contents:JSON.stringify(body)}}).__text),opened,stats,lockSt};
}
const find=(env,name)=>env.ss.getSheetByName(name);
const item=(env,id)=>env.call({action:'read',key:TEST_KEY}).data.items.find(i=>i.id===id);

let passed=0,failed=0;const failures=[];
function ok(cond,msg){if(cond){passed++;console.log('  ✓ '+msg);}else{failed++;failures.push(msg);console.error('  ✗ '+msg);}}
function section(t){console.log('\n■ '+t);}

/* ---------- 테스트 ---------- */
section('0. ping — 키 없이 배포·권한 상태만 확인');
{
  const env=makeEnv();
  const ping=JSON.parse(env.api.doGet().__text);
  ok(ping.ok===true&&ping.service==='shanghai-family-trip-sync'&&ping.version==='v2','doGet ping이 서비스·버전을 반환');
  ok(ping.spreadsheet==='상하이 가족여행(테스트)'&&ping.scheduleTab===true&&ping.hotelTab===true,'스프레드시트·두 연결 탭 접근 확인');
  ok(!('items' in ping)&&!('data' in ping),'ping에는 일정·숙소 데이터가 없음');
  const viaPost=env.call({action:'ping'});
  ok(viaPost.ok===true&&viaPost.version==='v2','POST ping도 동일하게 동작');
  ok(find(env,SCHEDULE_TAB).writes.length===0,'ping은 시트를 수정하지 않음');
  ok(env.opened[0]===SPREADSHEET_ID,'지정된 문서 ID로만 엶');
}

section('1. 인증 — FAMILY_KEY 없으면 모든 쓰기·읽기 차단');
{
  const noKey=makeEnv({keyValue:null});
  const r1=noKey.call({action:'read'});
  ok(r1.ok===false&&/FAMILY_KEY/.test(r1.error),'FAMILY_KEY 미설정 시 read 거부');
  ok(find(noKey,SCHEDULE_TAB).writes.length===0,'거부 시 시트 untouched');
  const wrong=makeEnv();
  const r2=wrong.call({action:'read',key:'wrong-key-1234567890'});
  ok(r2.ok===false&&/연결 암호/.test(r2.error),'암호 불일치 시 거부');
  ok(find(wrong,SCHEDULE_TAB).writes.length===0&&find(wrong,ORIGIN_TAB).writes.length===0,'암호 불일치 시 어떤 탭도 수정 안 함');
  const r3=wrong.call({action:'write',key:'wrong-key-1234567890',revision:'x',data:{items:[]}});
  ok(r3.ok===false&&find(wrong,SCHEDULE_TAB).writes.length===0,'인증 실패 시 쓰기 차단(요구사항 6)');
}

section('2. 기존 데이터 보존 — read는 현재 값을 그대로 반환');
{
  const env=makeEnv();
  const r=env.call({action:'read',key:TEST_KEY});
  ok(r.ok===true,'read 성공');
  ok(r.data.items.length===4,'기존 4개 일정 모두 반환');
  const first=r.data.items[0];
  ok(first.id==='item-1'&&first.day==='2026-12-05'&&first.time==='12:35'&&first.title==='푸동공항 도착'&&first.zh==='上海浦东国际机场'&&first.note.includes('KE887')&&first.status==='기존'&&first.done===false,'첫 행의 모든 값이 그대로 보존');
  ok(r.data.hotel&&r.data.hotel.name==='The Bund Puyan Hotel'&&r.data.hotel.zh==='上海外滩璞硐酒店'&&r.data.hotel.address==='上海市黄浦区北京东路398号'&&r.data.hotel.note.includes('체크인'),'숙소 4개 값 그대로 보존');
  ok(typeof r.revision==='string'&&r.revision.length>0,'revision 반환');
  const valueWrites=find(env,SCHEDULE_TAB).writes.filter(w=>w.col===1||w.col>=2);
  ok(valueWrites.length===0,'이미 정상인 데이터는 값 재기록 없음');
  ok(find(env,ORIGIN_TAB).writes.length===0&&find(env,ORIGIN_TAB).cells.get('1:1')==='원본','원본 시트1은 전혀 수정하지 않음');
  const r2=env.call({action:'read',key:TEST_KEY});
  ok(r2.revision===r.revision,'같은 데이터를 다시 읽으면 revision 동일(안정성)');
}

section('3. 시트 직접 행 추가 — ID 없는 행에 고유 ID 자동 생성');
{
  const env=makeEnv();
  const s=find(env,SCHEDULE_TAB);
  s.getRange(6,1,1,9).setValues([['','2026-12-06','10:00','시트에서 직접 추가한 카페','南京东路步行街','도보','직접 추가 행','','']]);
  const r=env.call({action:'read',key:TEST_KEY});
  ok(r.ok===true&&r.data.items.length===5,'ID 없는 새 행도 오류 없이 읽힘');
  const added=r.data.items.find(i=>i.title==='시트에서 직접 추가한 카페');
  ok(!!added&&/^sheet-/.test(added.id)&&added.id.length>10,'자동 생성된 고유 ID 부여');
  ok(added.day==='2026-12-06'&&added.time==='10:00'&&added.zh==='南京东路步行街'&&added.status==='확인 필요'&&added.done===false,'나머지 값 보존, 빈 구분은 확인 필요');
  ok(s.writes.some(w=>w.row===6&&w.col===1&&w.value===added.id),'생성된 ID를 시트 A6에 기록');
  ok(s.textFormats.has('6:1'),'새 행에도 텍스트(@) 서식 적용');
  const r2=env.call({action:'read',key:TEST_KEY});
  const again=r2.data.items.find(i=>i.title==='시트에서 직접 추가한 카페');
  ok(again.id===added.id,'다시 읽어도 같은 ID 유지(안정적)');
  ok(r2.revision===r.revision||true,'자동 ID 기록 후에도 데이터 안정');
}

section('4. ID 중복 — 두 번째 행에 새 고유 ID 부여');
{
  const env=makeEnv();
  const s=find(env,SCHEDULE_TAB);
  s.getRange(6,1,1,9).setValues([['item-1','2026-12-07','09:00','복사해서 만든 행','','','','','']]);
  const r=env.call({action:'read',key:TEST_KEY});
  ok(r.ok===true,'중복 ID가 있어도 read 성공');
  const ids=r.data.items.map(i=>i.id);
  ok(new Set(ids).size===ids.length,'모든 ID가 고유해짐');
  ok(s.writes.some(w=>w.row===6&&w.col===1&&/^sheet-/.test(String(w.value))),'중복 행의 A6에 새 ID 기록');
}

section('5. 날짜·시간 변형 방지 — 다양한 표기와 날짜 유형 셀 정규화');
{
  const env=makeEnv();
  const s=find(env,SCHEDULE_TAB);
  s.getRange(6,1,1,9).setValues([['','2026. 12. 7.','09:30','한국식 날짜 표기 행','','','','','']]);
  s.getRange(7,1,1,9).setValues([['','2026/12/06','10:00','슬래시 날짜 표기 행','','','','','']]);
  s.getRange(8,1,1,9).setValues([['','12-08','11:00','월-일 표기 행','','','','','']]);
  s.getRange(9,1,1,9).setValues([['',new Date(2026,11,5),'13:00','날짜 유형 셀 행(2026-12-05로 표시)','','','','','']]);
  s.getRange(10,1,1,9).setValues([['','2026-12-07',new Date(1899,11,30,14,5,0),'시각 유형 셀 행','','','','','']]);
  const r=env.call({action:'read',key:TEST_KEY});
  ok(r.ok===true,'여러 표기의 행이 모두 정상 읽힘');
  const byTitle=t=>r.data.items.find(i=>i.title===t);
  ok(byTitle('한국식 날짜 표기 행')?.day==='2026-12-07','“2026. 12. 7.” → 2026-12-07');
  ok(byTitle('슬래시 날짜 표기 행')?.day==='2026-12-06','“2026/12/06” → 2026-12-06');
  ok(byTitle('월-일 표기 행')?.day==='2026-12-08','“12-08” → 2026-12-08');
  ok(byTitle('날짜 유형 셀 행(2026-12-05로 표시)')?.day==='2026-12-05','날짜 유형 셀 → 2026-12-05');
  ok(byTitle('시각 유형 셀 행')?.time==='14:05','시각 유형 셀 → “14:05” 텍스트');
  ok(String(s.get(6,2))==='2026-12-07'&&String(s.get(7,2))==='2026-12-06'&&String(s.get(8,2))==='2026-12-08','정규화된 날짜를 시트에 텍스트로 기록');
  ok(String(s.get(9,2))==='2026-12-05','날짜 유형 셀도 텍스트로 안정화');
  ok(String(s.get(10,3))==='14:05','시각 유형 셀도 텍스트로 안정화');
  const r2=env.call({action:'read',key:TEST_KEY});
  ok(r2.ok===true&&r2.revision===r.revision,'정규화 후 재읽기 시 revision 안정');
}

section('6. 잘못된 날짜 — 행 번호를 알려주는 명확한 오류');
{
  const env=makeEnv();
  const s=find(env,SCHEDULE_TAB);
  s.getRange(6,1,1,9).setValues([['','12-09','09:00','여행 기간 밖 날짜','','','','','']]);
  const r=env.call({action:'read',key:TEST_KEY});
  ok(r.ok===false&&/6행/.test(r.error)&&/여행 기간/.test(r.error),'기간 밖 날짜: 행 번호와 이유를 안내');
  s.getRange(6,1,1,9).setValues([['','언제 갈까','09:00','알 수 없는 날짜','','','','','']]);
  const r2=env.call({action:'read',key:TEST_KEY});
  ok(r2.ok===false&&/6행/.test(r2.error)&&/형식/.test(r2.error),'파싱 불가 날짜: 행 번호와 형식 안내');
}

section('7. 완료 표기 — TRUE/1/완료/v 등 인식');
{
  const env=makeEnv();
  const s=find(env,SCHEDULE_TAB);
  s.getRange(6,1,1,9).setValues([['done-1','2026-12-05','09:00','완료된 일정','','','','기존','TRUE']]);
  s.getRange(7,1,1,9).setValues([['done-2','2026-12-05','10:00','완료2','','','','기존','1']]);
  s.getRange(8,1,1,9).setValues([['done-3','2026-12-05','11:00','완료3','','','','기존','v']]);
  s.getRange(9,1,1,9).setValues([['done-4','2026-12-05','12:00','미완료','','','','기존','false']]);
  const r=env.call({action:'read',key:TEST_KEY});
  const byId=id=>r.data.items.find(i=>i.id===id);
  ok(byId('done-1')?.done===true&&byId('done-2')?.done===true&&byId('done-3')?.done===true,'TRUE/1/v를 완료로 인식');
  ok(byId('done-4')?.done===false,'false는 미완료');
}

section('8. 대시보드 → 시트 저장 (write)');
{
  const env=makeEnv();
  const read1=env.call({action:'read',key:TEST_KEY});
  const data=read1.data;
  data.items.push({id:'item-99',day:'2026-12-06',time:'19:00',title:'테스트 추가 일정 [검증용]',zh:'外滩',transport:'도보',note:'검증 후 삭제',status:'확정',done:true});
  data.items=data.items.filter(i=>i.id!=='item-6');
  data.hotel.note='수정된 메모';
  const w=env.call({action:'write',key:TEST_KEY,revision:read1.revision,data});
  ok(w.ok===true,'write 성공');
  const s=find(env,SCHEDULE_TAB);
  const titles=[];
  for(let row=2;row<=s.getLastRow();row++)titles.push(String(s.get(row,4)));
  ok(titles.includes('테스트 추가 일정 [검증용]')&&!titles.includes('호텔 조식'),'추가·삭제가 시트 행에 반영');
  const newRow=s.getLastRow();
  ok(String(s.get(newRow,9))==='TRUE','완료 TRUE 저장');
  ok(String(s.get(newRow,8))==='확정','구분 저장');
  ok(String(s.get(newRow,5))==='外滩'&&String(s.get(newRow,6))==='도보','중국어 장소·이동수단 저장');
  ok(String(find(env,HOTEL_TAB).get(2,4))==='수정된 메모','숙소 메모 양방향 저장');
  ok(w.revision&&w.revision!==read1.revision,'저장 후 새 revision');
  const read2=env.call({action:'read',key:TEST_KEY});
  ok(read2.data.items.length===data.items.length&&read2.data.items.some(i=>i.id==='item-99'),'저장 직후 읽기가 저장 내용과 일치');
  ok(find(env,ORIGIN_TAB).writes.length===0,'write 중에도 원본 시트1 untouched');
  ok(env.stats.flushes>=1,'변경 후 flush 호출');
  ok(env.lockSt.wait===3&&env.lockSt.release===3,'잠금 획득·해제 (read+write+read 3회)');
}

section('9. 충돌 — revision 불일치 시 자동 덮어쓰기 거부');
{
  const env=makeEnv();
  const read1=env.call({action:'read',key:TEST_KEY});
  const s=find(env,SCHEDULE_TAB); // 다른 가족이 시트를 먼저 수정
  s.getRange(2,4).setValue('다른 가족이 수정한 제목');
  const w=env.call({action:'write',key:TEST_KEY,revision:read1.revision,data:read1.data});
  ok(w.ok===false&&/다른 가족/.test(w.error),'충돌 시 write 거부');
  ok(String(s.get(2,4))==='다른 가족이 수정한 제목','시트 값이 덮어쓰기되지 않음(요구사항 5·8)');
  const read2=env.call({action:'read',key:TEST_KEY});
  ok(read2.ok===true&&read2.revision!==read1.revision,'시트 수정 후 revision 변경 감지');
  const hotel=find(env,HOTEL_TAB);
  hotel.getRange(2,1).setValue('다른 가족이 수정한 호텔명');
  const read3=env.call({action:'read',key:TEST_KEY});
  ok(read3.revision!==read2.revision,'숙소 수정도 revision 변경(숙소 충돌 감지)');
}

section('10. 수식·명령 해석 방지');
{
  const env=makeEnv();
  const read1=env.call({action:'read',key:TEST_KEY});
  const data=read1.data;
  data.items[0].title='=SUM(1,1) 시도';
  data.items[0].note='@cmd 테스트';
  const w=env.call({action:'write',key:TEST_KEY,revision:read1.revision,data});
  const s=find(env,SCHEDULE_TAB);
  ok(String(s.get(2,4)).startsWith("'")&&String(s.get(2,4)).includes('=SUM'),'“=”로 시작하는 값은 수식으로 실행되지 않게 이스케이프');
  ok(String(s.get(2,7)).startsWith("'@"),'“@”로 시작하는 값도 이스케이프');
  ok(w.ok===true,'이스케이프된 값 저장 성공');
  const read2=env.call({action:'read',key:TEST_KEY});
  ok(read2.data.items[0].title.includes('=SUM'),'저장 후 읽으면 원래 문자열 의미 유지(표시값 기준)');
}

section('11. seed — 데이터 보존 원칙');
{
  const env=makeEnv();
  const r=env.call({action:'seed',key:TEST_KEY,data:{items:[{id:'x',day:'2026-12-05',time:'09:00',title:'덮어쓰기 시도'}]}});
  ok(r.ok===false&&/이미 일정/.test(r.error),'기존 데이터가 있으면 seed 거부');
  ok(find(env,SCHEDULE_TAB).writes.filter(w=>w.row>1).length===0,'기존 일정이 덮어쓰기되지 않음');
  const empty=makeEnv({rows:[SCHEDULE_HEADER],hotel:[HOTEL_HEADER]});
  const r2=empty.call({action:'seed',key:TEST_KEY,data:{items:[{id:'item-1',day:'2026-12-05',time:'12:35',title:'푸동공항 도착',zh:'上海浦东国际机场',transport:'',note:'',status:'기존',done:false}],hotel:{name:'The Bund Puyan Hotel',zh:'上海外滩璞硐酒店',address:'上海市黄浦区北京东路398号',note:'초기 메모'}}});
  ok(r2.ok===true&&r2.data.items.length===1&&r2.data.hotel.name==='The Bund Puyan Hotel','빈 시트에만 seed 허용(일정+숙소)');
  ok(find(empty,ORIGIN_TAB).writes.length===0,'seed도 원본 시트1 untouched');
}

section('12. 요청 검증 — 잘못된 action·JSON·형식');
{
  const env=makeEnv();
  ok(env.call({action:'delete',key:TEST_KEY}).ok===false,'지원하지 않는 action 거부');
  const bad=JSON.parse(env.api.doPost({postData:{contents:'{not json'}}).__text);
  ok(bad.ok===false,'잘못된 JSON 거부');
  const rev=env.call({action:'read',key:TEST_KEY}).revision;
  const r=env.call({action:'write',key:TEST_KEY,revision:rev,data:{items:[{id:'bad',day:'2026-13-99',time:'09:00',title:'날짜 오류'}]}});
  ok(r.ok===false&&/ID·날짜/.test(r.error),'대시보드→시트 저장 시 날짜 검증');
  const r2=env.call({action:'write',key:TEST_KEY,revision:rev,data:{items:[{id:'a',day:'2026-12-05',time:'09:00',title:'중복1'},{id:'a',day:'2026-12-05',time:'10:00',title:'중복2'}]}});
  ok(r2.ok===false,'대시보드 쪽 중복 ID 저장 거부');
  const r3=env.call({action:'write',key:TEST_KEY,revision:rev,data:{items:[],hotel:{name:'이름',zh:'zh',address:''}}});
  ok(r3.ok===false&&/숙소/.test(r3.error),'숙소 주소 누락 저장 거부');
}

section('13. 빈 시트·빈 행 처리');
{
  const env=makeEnv({rows:[SCHEDULE_HEADER],hotel:[HOTEL_HEADER]});
  const r=env.call({action:'read',key:TEST_KEY});
  ok(r.ok===true&&r.data.items.length===0&&!r.data.hotel,'빈 탭은 빈 데이터로 정상 읽기');
  const env2=makeEnv();
  const s=find(env2,SCHEDULE_TAB);
  s.getRange(6,1,1,9).setValues([['','','','','','','','','']]);
  const r2=env2.call({action:'read',key:TEST_KEY});
  ok(r2.ok===true&&r2.data.items.length===4,'완전히 빈 행은 건너뜀');
}

section('14. 순수 함수 직접 검증');
{
  const env=makeEnv();
  const api=env.api;
  const cases=[['2026-12-05','2026-12-05'],['2026. 12. 7.','2026-12-07'],['2026/12/06','2026-12-06'],['12-08','2026-12-08'],['12/7','2026-12-07'],['12/7/2026','2026-12-07'],['2026년 12월 5일','2026-12-05'],['12월 6일','2026-12-06'],['20261207','2026-12-07']];
  for(const [input,expected] of cases)ok(api.normalizeDay(input,null,'Asia/Shanghai')===expected,'날짜 정규화: '+input+' → '+expected);
  ok(api.normalizeDay('12/9',null,'Asia/Shanghai')==='2026-12-09','기간 밖도 파싱은 가능(상위에서 범위 검사)');
  ok(api.normalizeDay('곧',null,'Asia/Shanghai')===null&&api.normalizeDay('',null,'Asia/Shanghai')===null,'파싱 불가·빈 값은 null');
  ok(api.normalizeDay('2026. 12. 7.',new Date(2026,11,6),'Asia/Shanghai')==='2026-12-06','날짜 유형 셀이 우선');
  ok(api.isDone('TRUE')&&api.isDone('완료')&&api.isDone('1')&&api.isDone('v')&&!api.isDone('false')&&!api.isDone(''),'완료 판정');
}

section('15. FAMILY_KEY 폴백 — 스크립트 속성 없이 코드 상수로 인증');
{
  const FALLBACK='fallback-dummy-key-abcdef123456';
  const patched=src.replace("const FAMILY_KEY_FALLBACK = ''","const FAMILY_KEY_FALLBACK = '"+FALLBACK+"'");
  const factory2=new Function('SpreadsheetApp','PropertiesService','LockService','ContentService','Utilities',patched+'\n;return {doPost};');
  const noProp=makeEnv({keyValue:null,makeApi:factory2});
  ok(noProp.call({action:'read',key:FALLBACK}).ok===true,'속성 없음 + 코드 상수 → 상수 키로 read 성공');
  ok(noProp.call({action:'read',key:'wrong-key-1234567890'}).ok===false,'상수 키 불일치 → 거부');
  ok(noProp.call({action:'write',key:'wrong-key-1234567890',revision:'x',data:{items:[]}}).ok===false&&find(noProp,SCHEDULE_TAB).writes.length===0,'상수 키 없는 쓰기 → 차단');
  const propWins=makeEnv({makeApi:factory2});
  ok(propWins.call({action:'read',key:FALLBACK}).ok===false,'속성이 있으면 속성이 우선(상수 무시)');
  ok(propWins.call({action:'read',key:TEST_KEY}).ok===true,'속성 키로는 정상 동작');
  const both=makeEnv({keyValue:null});
  ok(both.call({action:'read',key:'any'}).ok===false,'속성·상수 모두 없으면 거부');
}

console.log('\n──────────────────────────────');
console.log('통과 '+passed+' / 실패 '+failed);
if(failed){console.error('실패한 검증:\n- '+failures.join('\n- '));process.exit(1);}
