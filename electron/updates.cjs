const {EventEmitter}=require('node:events');
function createUpdates(updater,{version,enabled,beforeInstall}){
 const events=new EventEmitter();let state={status:enabled?'idle':'disabled',version,latestVersion:null,percent:0,message:enabled?'시작할 때 새 버전을 확인합니다.':'설치한 앱에서 업데이트를 사용할 수 있습니다.'};let busy=false;
 const set=patch=>{state={...state,...patch};events.emit('state',{...state});};
 updater.autoDownload=false;updater.autoInstallOnAppQuit=false;updater.allowPrerelease=false;updater.allowDowngrade=false;updater.disableWebInstaller=true;
 updater.on('checking-for-update',()=>set({status:'checking',message:'새 버전을 확인하고 있습니다.'}));
 updater.on('update-available',info=>set({status:'available',latestVersion:info.version,message:'새 버전을 받을 수 있습니다.'}));
 updater.on('update-not-available',()=>set({status:'current',message:'최신 버전입니다.'}));
 updater.on('download-progress',p=>set({status:'downloading',percent:Math.round(p.percent),message:'업데이트를 내려받고 있습니다.'}));
 updater.on('update-downloaded',info=>set({status:'downloaded',latestVersion:info.version,percent:100,message:'저장 후 다시 시작하면 업데이트됩니다.'}));
 const fail=()=>set({status:'error',message:'업데이트 서버에 연결하지 못했습니다. 잠시 후 다시 확인하세요.'});updater.on('error',fail);
 return {events,get:()=>({...state}),async check(){if(!enabled||busy||['available','downloading','downloaded','installing'].includes(state.status))return {...state};busy=true;try{await updater.checkForUpdates();}catch{fail();}finally{busy=false;}return {...state};},async download(){if(state.status!=='available'||busy)return {...state};busy=true;set({status:'downloading',percent:0});try{await updater.downloadUpdate();}catch{fail();}finally{busy=false;}return {...state};},async install(){if(state.status!=='downloaded'||busy)throw new Error('업데이트 다운로드를 먼저 완료하세요.');busy=true;try{await beforeInstall();set({status:'installing',message:'자료를 백업했습니다. 앱을 다시 시작합니다.'});updater.quitAndInstall(true,true);}finally{busy=false;}return {...state};}};
}
module.exports={createUpdates};
