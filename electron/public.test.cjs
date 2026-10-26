const test=require('node:test'),assert=require('node:assert/strict'),path=require('node:path'),{EventEmitter}=require('node:events');
const engine=require('./engine.cjs'),maintenance=require('./maintenance.cjs'),{createUpdates}=require('./updates.cjs');
const template=path.resolve(__dirname,'../assets/template.pptx');
test('public seed is empty and generic template exports without losing font resources',async()=>{
 const source=require('../assets/birthdays-2026.json');assert.equal(source.mode,'external-only');assert.equal(source.records.length,0);
 const data=await engine.defaults(template);assert.equal(data.cellNames.length,54);assert.equal(data.birthdays.length,0);
 const original=await engine.readTemplate(template);data.cellGroups=engine.migrateCellGroups(data.cellNames);data.militaryMembers=[];data.birthdays=[];data.ads=['안내 사항'];data.body='1 말씀 본문';data.prayers=data.prayers.map(p=>({...p,name:'담당자'}));
 const ZIP=require('jszip'),z=await ZIP.loadAsync(await engine.exportPpt(template,data));
 assert.equal(Object.keys(z.files).filter(k=>k.startsWith('ppt/fonts/')).length,Object.keys(original.zip.files).filter(k=>k.startsWith('ppt/fonts/')).length);
 assert.equal(Object.keys(z.files).some(k=>/thumbnail/i.test(k)),false);
});
test('backup round-trip preserves user records; invalid restore is rejected',async()=>{
 const d=await engine.defaults(template);Object.assign(d,{schemaVersion:2,cellGroups:engine.migrateCellGroups(d.cellNames),militaryMembers:[],birthdayRoster:[{name:'테스트인',month:'2',day:'29'}],exportFormat:'both'});
 assert.deepEqual(maintenance.parseBackup(maintenance.createBackup(d,'1.5.0')),d);assert.throws(()=>maintenance.parseBackup('{"schemaVersion":2}'));
});
test('update flow waits for download and successful backup before installation',async()=>{
 const fake=new EventEmitter();const calls=[];fake.checkForUpdates=async()=>{fake.emit('checking-for-update');fake.emit('update-available',{version:'1.6.0'});};fake.downloadUpdate=async()=>{fake.emit('download-progress',{percent:43.5});fake.emit('update-downloaded',{version:'1.6.0'});};fake.quitAndInstall=(...a)=>calls.push(['install',...a]);
 let fail=true;const u=createUpdates(fake,{version:'1.5.0',enabled:true,beforeInstall:async()=>{calls.push('backup');if(fail)throw Error('backup failed');}});
 assert.equal(fake.autoDownload,false);assert.equal(fake.autoInstallOnAppQuit,false);await assert.rejects(()=>u.install());
 await u.check();assert.equal(u.get().status,'available');await u.download();assert.equal(u.get().status,'downloaded');await assert.rejects(()=>u.install());assert.deepEqual(calls,['backup']);fail=false;await u.install();assert.deepEqual(calls,['backup','backup',['install',true,true]]);
});
test('offline and development mode leave editing available',async()=>{
 const f=new EventEmitter();f.checkForUpdates=async()=>{throw Error('private network detail');};const u=createUpdates(f,{version:'1.5.0',enabled:true,beforeInstall:async()=>{}});await u.check();assert.equal(u.get().status,'error');assert.equal(u.get().message.includes('private'),false);
 const disabled=createUpdates(new EventEmitter(),{version:'1.5.0',enabled:false});await disabled.check();assert.equal(disabled.get().status,'disabled');
});
