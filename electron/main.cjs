const { app, BrowserWindow, ipcMain, dialog, shell, session } = require('electron');
const path = require('node:path');
const fs = require('node:fs/promises');
const engine = require('./engine.cjs');
const birthdaySource=require('../assets/birthdays-2026.json');
const {migrateBirthdayRoster,birthdaysForDate}=require('../shared/birthdays.mjs');
const {publishFiles}=require('./export-files.cjs');
const maintenance=require('./maintenance.cjs');
let maintenanceBusy=false,restoring=false;
let exporting=false;
const recentOutputs=new Set();
const {execFile}=require('node:child_process');
const {promisify}=require('node:util');
const execute=promisify(execFile);
let window;
const template = () => path.join(app.getAppPath(), 'assets', 'template.pptx');
const draftFile = () => path.join(app.getPath('userData'), 'draft.json');
let draftQueue = Promise.resolve();
function trusted(event) {
  if (event.sender !== window.webContents || event.senderFrame !== window.webContents.mainFrame) throw new Error('허용되지 않은 요청입니다.');
}
function handle(name, fn) { ipcMain.handle(name, async (event, ...args) => { trusted(event); try { return { ok:true, value:await fn(...args) }; } catch (e) { return { ok:false, error:e.message }; } }); }
app.setName('YouthBulletinReact');
// Portable bundles keep their profile next to the executable, independent of this PC's installed copy.
if(app.isPackaged&&require('node:fs').existsSync(path.join(path.dirname(process.execPath),'portable.json'))){
  const portableData=path.join(path.dirname(process.execPath),'data');
  require('node:fs').mkdirSync(portableData,{recursive:true});
  app.setPath('userData',portableData);
}
if (!app.isPackaged && process.env.BULLETIN_TEST_DATA) app.setPath('userData',process.env.BULLETIN_TEST_DATA);
if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
app.on('second-instance',()=>{if(window){if(window.isMinimized())window.restore();window.show();window.focus();}});
app.whenReady().then(async () => {
  session.defaultSession.setPermissionRequestHandler((_,__,callback)=>callback(false));
  async function loadInitial() {
    const initial = await engine.defaults(template());
    if(birthdaySource.mode==='external-only'){initial.title='';initial.body='';initial.ads=[];initial.birthdays=[];initial.preacher='';initial.team='';}
    let draft = null;
    try { draft=JSON.parse(await fs.readFile(draftFile(),'utf8')); } catch {}
    if(draft && draft.schemaVersion!==2){
      const legacy=JSON.parse(await fs.readFile(path.join(app.getAppPath(),'assets','legacy-defaults.json'),'utf8'));
      for(const key of Object.keys(legacy))if(JSON.stringify(draft[key])===JSON.stringify(legacy[key]))delete draft[key];
    }
    if(draft && typeof draft==='object') {
      for(const key of ['title','reference','body','cellTitle','preacher','team']) if(typeof draft[key]==='string') initial[key]=draft[key];
      if(Array.isArray(draft.ads) && draft.ads.every(x=>typeof x==='string')) initial.ads=draft.ads;
      if(Array.isArray(draft.birthdays)) initial.birthdays=draft.birthdays.filter(x=>x && typeof x.name==='string').map(x=>({day:String(x.day),name:x.name}));
      if(['number','circle','bullet'].includes(draft.adStyle))initial.adStyle=draft.adStyle;
      if(Array.isArray(draft.prayers))initial.prayers=initial.prayers.map(p=>{const old=draft.prayers.find(x=>x.date===p.date);return old?{...p,enabled:!!old.enabled,name:typeof old.name==='string'?old.name:''}:p;});
    }
    if(Array.isArray(draft?.cellNames)&&draft.cellNames.length===initial.cellNames.length&&draft.cellNames.every(v=>typeof v==='string'))initial.cellNames=draft.cellNames;
    initial.cellGroups=Array.isArray(draft?.cellGroups)?draft.cellGroups:engine.migrateCellGroups(initial.cellNames);
    if(birthdaySource.mode==='external-only'&&!draft?.cellGroups)initial.cellGroups=initial.cellGroups.map(g=>({...g,leader:'',members:[]}));
    initial.militaryMembers=Array.isArray(draft?.militaryMembers)?draft.militaryMembers:(birthdaySource.mode==='external-only'?[]:initial.cellNames.slice(51,54).filter(v=>v.trim()));
    const month=Number(initial.date.slice(5,7));
    const oldMonth=draft?.date?Number(draft.date.slice(5,7)):10;
    const previousRoster=Array.isArray(draft?.birthdayRoster)?draft.birthdayRoster.filter(b=>b&&typeof b.name==='string').map(b=>({...b,month:String(b.month??''),day:String(b.day??'')})):initial.birthdays.map(b=>({...b,month:String(oldMonth)}));
    const migrated=migrateBirthdayRoster(previousRoster,draft?.birthdaySourceVersion,birthdaySource);
    initial.birthdayRoster=migrated.roster;initial.birthdaySourceVersion=migrated.version;
    initial.birthdayExpectedNames=Array.isArray(draft?.birthdayExpectedNames)?draft.birthdayExpectedNames:initial.birthdayRoster.map(r=>r.name);
    initial.birthdays=birthdaysForDate(initial.birthdayRoster,initial.date);
    initial.bodyReference=typeof draft?.bodyReference==='string'?draft.bodyReference:initial.body.split(/\r?\n/)[0].trim();
    initial.bibleSource=typeof draft?.bibleSource==='string'?draft.bibleSource:'이전에 저장한 본문';
    const original=await engine.defaults(template());original.bodyReference=original.reference;original.bibleSource='원본 템플릿 본문';
    initial.exportFormat=['both','pptx','pdf'].includes(draft?.exportFormat)?draft.exportFormat:'both';
    const newWeek=!!draft?.date&&draft.date!==initial.date;
    initial.lastWeekAds=newWeek?(Array.isArray(draft.ads)?draft.ads:[]):(Array.isArray(draft?.lastWeekAds)?draft.lastWeekAds:[]);
    if(newWeek){await fs.mkdir(path.join(app.getPath('userData'),'history'),{recursive:true});await fs.writeFile(path.join(app.getPath('userData'),'history',String(draft.date).replace(/[^0-9-]/g,'')+'.json'),JSON.stringify(draft));}
    if(newWeek)initial.ads=[];
    return {data:initial,original,restored:!!draft,newWeek,books:engine.books,birthdaySource};
  }
  const updates=require('./updates.cjs').createUpdates(require('electron-updater').autoUpdater,{
    version:app.getVersion(),enabled:app.isPackaged&&!require('node:fs').existsSync(path.join(path.dirname(process.execPath),'portable.json')),
    beforeInstall:async()=>{
      if(exporting||maintenanceBusy||restoring)throw new Error('진행 중인 작업을 마친 뒤 업데이트하세요.');
      await draftQueue;
      const data=JSON.parse(await fs.readFile(draftFile(),'utf8'));
      await maintenance.writeAtomic(path.join(app.getPath('userData'),'backups','before-update-'+Date.now()+'.json'),maintenance.createBackup(data,app.getVersion()));
    }
  });
  updates.events.on('state',state=>{if(window&&!window.isDestroyed())window.webContents.send('update-state',state);});
  handle('update-state',()=>updates.get());handle('update-check',()=>updates.check());handle('update-download',()=>updates.download());handle('update-install',()=>updates.install());
  handle('init',loadInitial);
  handle('draft', async data => {
    if(restoring)throw new Error('자료를 복원하고 있습니다. 잠시 기다려 주세요.');
    const serialized=JSON.stringify({...data,schemaVersion:2});if(serialized.length>2000000)throw new Error('입력 내용이 너무 큽니다.');
    draftQueue=draftQueue.catch(()=>{}).then(async()=>{await fs.mkdir(app.getPath('userData'),{recursive:true});const temp=draftFile()+'.tmp';await fs.writeFile(temp,serialized);await fs.rename(temp,draftFile());});
    await draftQueue; return true;
  });
  handle('backup',async()=>{
    if(maintenanceBusy||exporting)throw new Error('다른 작업을 마친 뒤 다시 시도하세요.');maintenanceBusy=true;
    try{
      await draftQueue;
      const data=JSON.parse(await fs.readFile(draftFile(),'utf8'));
      const content=maintenance.createBackup(data,app.getVersion());
      const day=new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Seoul'}).format(new Date());
      const result=await dialog.showSaveDialog(window,{title:'주보 자료 백업',defaultPath:path.join(app.getPath('documents'),day+' 주보 자료 백업.json'),filters:[{name:'주보 자료 백업',extensions:['json']}]});
      if(result.canceled)return null;
      const file=result.filePath.toLowerCase().endsWith('.json')?result.filePath:result.filePath+'.json';
      if(path.resolve(file).toLowerCase()===path.resolve(draftFile()).toLowerCase())throw new Error('현재 작성 파일에는 덮어쓸 수 없습니다. 다른 위치를 선택하세요.');
      await maintenance.writeAtomic(file,content);recentOutputs.add(path.resolve(file));return {path:file,people:data.birthdayRoster.length};
    }finally{maintenanceBusy=false;}
  });
  handle('restore',async()=>{
    if(maintenanceBusy||exporting)throw new Error('다른 작업을 마친 뒤 다시 시도하세요.');maintenanceBusy=true;
    try{
      const result=await dialog.showOpenDialog(window,{title:'주보 자료 복원',properties:['openFile'],filters:[{name:'주보 자료 백업 / draft.json',extensions:['json']}]});if(result.canceled)return null;
      const data=await maintenance.readBackup(result.filePaths[0]);
      const answer=await dialog.showMessageBox(window,{type:'question',title:'주보 자료 복원',message:'선택한 자료를 불러올까요?',detail:'생일 명단 '+data.birthdayRoster.length+'명 · 셀 '+data.cellGroups.length+'개\n현재 자료는 먼저 자동 백업합니다. 날짜는 돌아오는 일요일에 맞춰집니다.',buttons:['취소','복원하기'],defaultId:0,cancelId:0});if(answer.response!==1)return null;
      restoring=true;await draftQueue;
      const previous=JSON.parse(await fs.readFile(draftFile(),'utf8'));
      const safety=path.join(app.getPath('userData'),'backups','before-restore-'+Date.now()+'.json');
      await maintenance.writeAtomic(safety,maintenance.createBackup(previous,app.getVersion()));
      await maintenance.writeAtomic(draftFile(),JSON.stringify(data));
      let state;try{state=await loadInitial();}catch{
        try{await maintenance.writeAtomic(draftFile(),JSON.stringify(previous));}catch{throw new Error('복원을 마치지 못했습니다. 이전 자료 자동 백업: '+safety);}
        throw new Error('복원을 마치지 못해 이전 자료로 되돌렸습니다.');
      }
      recentOutputs.add(path.resolve(safety));return {...state,previousBackup:safety};
    }finally{restoring=false;maintenanceBusy=false;}
  });
  handle('environment',async()=>{
    if(maintenanceBusy||exporting)throw new Error('다른 작업을 마친 뒤 다시 시도하세요.');maintenanceBusy=true;
    try{return await maintenance.environmentCheck({templatePath:template(),helperPath:app.isPackaged?path.join(process.resourcesPath,'ValidatePpt.exe'):path.join(app.getAppPath(),'assets','ValidatePpt.exe'),userDataPath:app.getPath('userData'),execute,lookup:engine.bible});}finally{maintenanceBusy=false;}
  });
  handle('bible', engine.bible);
  handle('export', async data => {
    if(maintenanceBusy)throw new Error('자료 관리 작업을 마친 뒤 저장하세요.');
    if(exporting)throw new Error('주보를 저장하고 있습니다. 잠시 기다려 주세요.');exporting=true;
    const cleanup=[];
    try {
      if(data.bodyReference!==data.reference||!data.body?.trim())throw new Error('선택한 구절의 본문을 먼저 불러오거나 직접 입력해 주세요.');
      if(!data.title?.trim())throw new Error('설교 제목을 입력해 주세요.');
      const buffer=await engine.exportPpt(template(),data);
      const format=['both','pptx','pdf'].includes(data.exportFormat)?data.exportFormat:'both';
      let prefs={};try{prefs=JSON.parse(await fs.readFile(path.join(app.getPath('userData'),'export-preferences.json'),'utf8'));}catch{}
      const ext=format==='pdf'?'pdf':'pptx';
      const result=await dialog.showSaveDialog(window,{title:format==='both'?'PPT와 PDF 함께 저장':'주보 저장',defaultPath:path.join(prefs.directory||app.getPath('documents'),data.date.slice(2).replaceAll('-','')+' 청년부 주보.'+ext),filters:[{name:ext==='pdf'?'PDF':'PowerPoint',extensions:[ext]}]});
      if(result.canceled)return null;
      const target=result.filePath.toLowerCase().endsWith('.'+ext)?result.filePath:result.filePath+'.'+ext;
      const pptPath=ext==='pptx'?target:target.slice(0,-4)+'.pptx',pdfPath=ext==='pdf'?target:target.slice(0,-5)+'.pdf';
      if(path.resolve(pptPath)===path.resolve(template()))throw new Error('원본 템플릿에는 덮어쓸 수 없습니다.');
      if(format==='both'&&await fs.stat(pdfPath).catch(()=>null)){
        const answer=await dialog.showMessageBox(window,{type:'question',buttons:['취소','PPT·PDF 덮어쓰기'],defaultId:0,cancelId:0,message:'같은 이름의 PDF도 있습니다.',detail:'두 파일을 새 주보로 바꿉니다.\n'+pdfPath});if(answer.response!==1)return null;
      }
      const token=Date.now()+'-'+Math.random().toString(16).slice(2),tmp=path.join(path.dirname(target),'.bulletin-'+token+'.pptx'),report=tmp+'.json',pdfTemp=tmp+'.pdf';cleanup.push(tmp,report,pdfTemp);
      await fs.writeFile(tmp,buffer);
      const helper=app.isPackaged?path.join(process.resourcesPath,'ValidatePpt.exe'):path.join(app.getAppPath(),'assets','ValidatePpt.exe');
      try{await execute(helper,format==='pptx'?[tmp,report]:[tmp,report,pdfTemp],{windowsHide:true,timeout:120000});}catch(e){if(!await fs.stat(report).catch(()=>null))throw new Error('PowerPoint에서 주보를 만들지 못했습니다. PowerPoint 실행 상태를 확인하고 다시 시도해 주세요.');}
      const check=JSON.parse(await fs.readFile(report,'utf8'));if(!check.ok)throw new Error(check.issues.join('\n'));
      const files=[];if(format!=='pdf')files.push({temp:tmp,target:pptPath});if(format!=='pptx'){if((await fs.stat(pdfTemp)).size===0)throw new Error('PDF가 생성되지 않았습니다.');files.push({temp:pdfTemp,target:pdfPath});}
      await publishFiles(files,token);for(const file of files)recentOutputs.add(path.resolve(file.target));
      await fs.writeFile(path.join(app.getPath('userData'),'export-preferences.json'),JSON.stringify({directory:path.dirname(target)})).catch(()=>{});
      return {path:files[0].target,pptPath:format==='pdf'?null:pptPath,pdfPath:format==='pptx'?null:pdfPath,fitting:check.fitting};
    }finally{for(const file of cleanup)await fs.unlink(file).catch(()=>{});exporting=false;}
  });
  async function checkOutput(file){if(typeof file!=='string'||!recentOutputs.has(path.resolve(file)))throw new Error('이 앱에서 방금 저장한 파일만 열 수 있습니다.');await fs.access(file);}
  handle('reveal',async file=>{await checkOutput(file);shell.showItemInFolder(file);return true;});
  handle('open-output',async file=>{await checkOutput(file);const error=await shell.openPath(file);if(error)throw new Error('파일을 열 수 없습니다: '+error);return true;});
  window=new BrowserWindow({width:1240,height:880,minWidth:940,minHeight:680,backgroundColor:'#fafbf7',title:'청년부 주보',autoHideMenuBar:true,webPreferences:{preload:path.join(__dirname,'preload.cjs'),nodeIntegration:false,contextIsolation:true,sandbox:true}});
  window.webContents.setWindowOpenHandler(()=>({action:'deny'}));
  window.webContents.on('will-navigate',(event)=>event.preventDefault());
  await window.loadFile(path.join(app.getAppPath(),'dist','index.html'));
  updates.check();
});
app.on('window-all-closed',()=>app.quit());
}
