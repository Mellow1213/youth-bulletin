const fs=require('node:fs/promises');
const path=require('node:path');
const os=require('node:os');
const {randomUUID}=require('node:crypto');
const ZIP=require('jszip');
const MAX_BYTES=10*1024*1024;
function validateDraft(data){
 const fail=()=>{throw new Error('주보 자료의 형식이 올바르지 않습니다. 현재 자료는 변경하지 않았습니다.');};
 const string=(v,max=100000)=>typeof v==='string'&&v.length<=max;
 const strings=(v,max=1000,size=10000)=>Array.isArray(v)&&v.length<=max&&v.every(x=>string(x,size));
 if(!data||typeof data!=='object'||Array.isArray(data)||data.schemaVersion!==2||Buffer.byteLength(JSON.stringify(data))>2000000)fail();
 if(!/^\d{4}-\d{2}-\d{2}$/.test(data.date||'')||!Number.isFinite(Date.parse(data.date+'T00:00:00Z'))||new Date(data.date+'T00:00:00Z').toISOString().slice(0,10)!==data.date)fail();
 for(const key of ['title','reference','body','cellTitle','cells'])if(!string(data[key]))fail();
 if(!strings(data.ads,100)||!strings(data.militaryMembers,100,120))fail();
 if(data.lastWeekAds!==undefined&&!strings(data.lastWeekAds,100))fail();
 if(!Array.isArray(data.birthdayRoster)||data.birthdayRoster.length>1000||data.birthdayRoster.some(b=>!b||!string(b.name,120)||!['string','number'].includes(typeof b.month)||!['string','number'].includes(typeof b.day)||String(b.month).length>8||String(b.day).length>8))fail();
 if(!Array.isArray(data.cellGroups)||data.cellGroups.length!==6||data.cellGroups.some((g,i)=>!g||g.id!=='cell-'+(i+1)||!string(g.leader,120)||!strings(g.members,100,120)))fail();
 if(!Array.isArray(data.prayers)||data.prayers.length>5||data.prayers.some(p=>!p||!/^\d{4}-\d{2}-\d{2}$/.test(p.date||'')||typeof p.enabled!=='boolean'||!string(p.name,120)))fail();
 if(!['number','circle','bullet'].includes(data.adStyle)||!['both','pptx','pdf'].includes(data.exportFormat))fail();
 return data;
}
function createBackup(data,version){
 validateDraft(data);
 const content=JSON.stringify({format:'youth-bulletin-backup',formatVersion:1,appVersion:version,createdAt:new Date().toISOString(),data},null,2);
 if(Buffer.byteLength(content)>MAX_BYTES)throw new Error('백업 자료가 너무 큽니다.');
 return content;
}
function parseBackup(content){
 if(Buffer.byteLength(content)>MAX_BYTES)throw new Error('백업 파일은 10MB 이하만 불러올 수 있습니다.');
 let value;try{value=JSON.parse(content.replace(/^\uFEFF/,''));}catch{throw new Error('백업 JSON 파일을 읽을 수 없습니다. 현재 자료는 변경하지 않았습니다.');}
 if(value?.format==='youth-bulletin-backup'){
  if(value.formatVersion!==1)throw new Error('지원하지 않는 백업 버전입니다. 현재 자료는 변경하지 않았습니다.');
  return validateDraft(value.data);
 }
 // Also support copying a draft.json from the existing portable package.
 return validateDraft(value);
}
async function readBackup(file){
 if((await fs.stat(file)).size>MAX_BYTES)throw new Error('백업 파일은 10MB 이하만 불러올 수 있습니다.');
 return parseBackup(await fs.readFile(file,'utf8'));
}
async function writeAtomic(file,content){
 await fs.mkdir(path.dirname(file),{recursive:true});
 const temp=file+'.'+randomUUID()+'.tmp';
 try{await fs.writeFile(temp,content);await fs.rename(temp,file);}finally{await fs.unlink(temp).catch(()=>{});}
}
async function environmentCheck({templatePath,helperPath,userDataPath,execute,lookup}){
 const items=[{name:'운영체제',status:process.platform==='win32'?'ok':'fail',detail:process.platform==='win32'?'Windows '+os.release():'Windows에서 파일 생성을 지원합니다.'}];
 const probe=path.join(userDataPath,'.write-check-'+randomUUID());
 try{await fs.mkdir(userDataPath,{recursive:true});await fs.writeFile(probe,'ok');items.push({name:'자료 저장 폴더',status:'ok',detail:'명단과 작성 내용을 저장할 수 있습니다.'});}catch{items.push({name:'자료 저장 폴더',status:'fail',detail:'앱 폴더 또는 자료 폴더에 쓰기 권한이 필요합니다.'});}finally{await fs.unlink(probe).catch(()=>{});}
 let templateReady=false,helperReady=false;
 try{const z=await ZIP.loadAsync(await fs.readFile(templatePath));templateReady=!!z.file('ppt/slides/slide2.xml');if(!templateReady)throw Error();const fonts=Object.keys(z.files).filter(n=>n.startsWith('ppt/fonts/')&&!z.files[n].dir).length;items.push({name:'주보 양식·글꼴',status:'ok',detail:'양식 확인 완료 · 포함 글꼴 '+fonts+'개'});}catch{items.push({name:'주보 양식·글꼴',status:'fail',detail:'양식 파일을 읽을 수 없습니다. 앱을 다시 압축 해제하거나 업데이트하세요.'});}
 try{helperReady=(await fs.stat(helperPath)).isFile();}catch{}
 if(!helperReady)items.push({name:'출력 도구',status:'fail',detail:'출력 도구가 없습니다. 실행 파일만 옮기지 말고 앱 폴더 전체를 복사하세요.'});
 await Promise.all([
  (async()=>{if(!templateReady||!helperReady||process.platform!=='win32'){items.push({name:'PowerPoint·PDF 생성',status:'fail',detail:'위 항목을 먼저 확인한 뒤 다시 점검하세요.'});return;}
   const folder=await fs.mkdtemp(path.join(os.tmpdir(),'youth-bulletin-check-'));const ppt=path.join(folder,'check.pptx'),report=path.join(folder,'check.json'),pdf=path.join(folder,'check.pdf');
   try{await fs.copyFile(templatePath,ppt);await execute(helperPath,[ppt,report,pdf],{windowsHide:true,timeout:120000});const r=JSON.parse(await fs.readFile(report,'utf8'));if(!r.ok)throw Error(r.issues.join(' '));const bytes=await fs.readFile(pdf);if(bytes.subarray(0,4).toString()!=='%PDF')throw Error('PDF 확인 실패');items.push({name:'PowerPoint·PDF 생성',status:'ok',detail:'설치된 PowerPoint로 양식을 열고 PDF 생성까지 확인했습니다.'});}
   catch{items.push({name:'PowerPoint·PDF 생성',status:'fail',detail:'PowerPoint로 파일을 생성하지 못했습니다. PC용 Microsoft PowerPoint를 설치하고 처음 실행 시 나오는 설정을 완료한 뒤 다시 점검하세요.'});}
   finally{for(const file of [ppt,report,pdf])await fs.unlink(file).catch(()=>{});await fs.rmdir(folder).catch(()=>{});}
  })(),
  (async()=>{try{await lookup('요한복음 1:1');items.push({name:'새번역 본문 조회',status:'ok',detail:'대한성서공회에서 본문을 정상적으로 가져왔습니다.'});}catch{items.push({name:'새번역 본문 조회',status:'warning',detail:'인터넷 또는 조회 사이트 연결을 확인하세요. 본문 직접 입력으로 사용할 수 있습니다.'});}})()
 ]);
 return {ready:items.every(i=>i.status!=='fail'),checkedAt:new Date().toISOString(),items};
}
module.exports={validateDraft,createBackup,parseBackup,readBackup,writeAtomic,environmentCheck};
