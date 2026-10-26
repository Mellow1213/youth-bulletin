const fs=require('node:fs/promises');
// Publish the complete set together. Restore previous files if a destination is locked or a move fails.
async function publishFiles(files, token) {
 const backups=[],published=[];
 try {
  for(const file of files){const existing=await fs.stat(file.target).catch(()=>null);if(existing&&!existing.isFile())throw new Error('파일 경로를 확인하세요.');if(existing){const backup=file.target+'.'+token+'.backup';await fs.rename(file.target,backup);backups.push({target:file.target,backup});}}
  for(const file of files){await fs.rename(file.temp,file.target);published.push(file.target);}
 } catch(error) {
  for(const target of published)await fs.unlink(target).catch(()=>{});
  const recovery=[];
  for(const item of backups){try{await fs.rename(item.backup,item.target);}catch{recovery.push(item.backup);}}
  throw new Error('파일을 저장하지 못했습니다. 같은 이름의 PPT 또는 PDF가 열려 있다면 닫고 다시 저장하세요.'+(recovery.length?' 이전 파일 보관 위치: '+recovery.join(', '):''));
 }
 for(const item of backups)await fs.unlink(item.backup).catch(()=>{});
}
module.exports={publishFiles};
