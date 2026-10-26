export const normalizeName = value => String(value || '').normalize('NFKC').replace(/\s+/g,'');
export function validBirthday(record) {
 const month=String(record?.month??''),day=String(record?.day??'');
 return /^\d+$/.test(month)&&/^\d+$/.test(day)&&+month>=1&&+month<=12&&+day>=1&&+day<=new Date(Date.UTC(2000,+month,0)).getUTCDate();
}
export const birthdaysForDate=(roster,date)=>roster.filter(b=>b.name?.trim()&&validBirthday(b)&&+b.month===+date.slice(5,7)).sort((a,b)=>+a.day-+b.day||a.name.localeCompare(b.name,'ko'));
export function migrateBirthdayRoster(roster,version,source) {
 if(version===source.version)return {roster:roster.map(r=>({...r})),version};
 const sourceKeys=new Set(source.records.map(r=>normalizeName(r.name)));
 // The supplied workbook is authoritative on first import; blank dates remain unknown.
 return {roster:[...source.records.map(r=>({...r})),...roster.filter(r=>r.name?.trim()&&!sourceKeys.has(normalizeName(r.name))).map(r=>({...r}))],version:source.version};
}
export function birthdayIssues(roster,baseline,groups=[],military=[]) {
 const byName=new Map(),issues=[];
 for(let index=0;index<roster.length;index++){
  const r=roster[index],key=normalizeName(r.name);if(!key)continue;
  if(!byName.has(key))byName.set(key,[]);byName.get(key).push(index);
  if(!validBirthday(r))issues.push({kind:!r.month&&!r.day?'unknown':'invalid',name:r.name,index,message:!r.month&&!r.day?'생일 미등록':'생일 날짜 확인 필요'});
 }
 const expected=new Map();for(const r of baseline)expected.set(normalizeName(r.name),r.name);
 for(const g of groups)for(const name of [g.leader,...g.members])if(normalizeName(name))expected.set(normalizeName(name),name);
 for(const name of military)if(normalizeName(name))expected.set(normalizeName(name),name);
 for(const [key,name] of expected)if(!byName.has(key))issues.push({kind:'missing',name,index:null,message:'생일 명단에 없음'});
 for(const [key,indexes] of byName)if(indexes.length>1)issues.push({kind:'duplicate',name:roster[indexes[0]].name,index:indexes[0],message:'같은 이름의 기록 중복 · 동명이인 여부 확인'});
 return issues;
}
