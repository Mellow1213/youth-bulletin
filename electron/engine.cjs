const fs = require('node:fs/promises');
const JSZip = require('jszip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const cheerio = require('cheerio');
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const descendants = (n, ns, name) => Array.from(n.getElementsByTagNameNS(ns, name));
const children = (n, name) => Array.from(n.childNodes).filter(x => x.nodeType === 1 && (!name || x.localName === name));
const shape = (doc, id) => descendants(doc, P, 'sp').find(s => descendants(s, P, 'cNvPr')[0]?.getAttribute('id') === String(id));
const text = s => children(descendants(s, P, 'txBody')[0], 'p').map(p => descendants(p, A, 't').map(t => t.textContent).join('')).join('\n');
const layout = require('./layout.cjs');
const setText = layout.setText;
function koreanDay(now = new Date()) {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Seoul', year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(now);
  const p = Object.fromEntries(parts.map(p => [p.type, p.value]));
  return `${p.year}-${p.month}-${p.day}`;
}
const dateOf = str => new Date(str + 'T00:00:00Z');
const iso = d => d.toISOString().slice(0, 10);
function upcomingSunday(now = new Date()) {
  const d = dateOf(koreanDay(now)); d.setUTCDate(d.getUTCDate() + (7 - d.getUTCDay()) % 7); return iso(d);
}
function monthSundays(day) {
  const d = dateOf(day); d.setUTCDate(1); const month = d.getUTCMonth();
  d.setUTCDate(1 + (7 - d.getUTCDay()) % 7); const list = [];
  while (d.getUTCMonth() === month) { list.push(iso(d)); d.setUTCDate(d.getUTCDate() + 7); } return list;
}
const shortDate = str => { const d = dateOf(str); return `${d.getUTCMonth() + 1}/${d.getUTCDate()}`; };
function formatAds(items, style = 'number') {
  let number = 0; const output = [];
  for (const item of items) {
    const value = String(item).trimEnd(); if (!value.trim()) continue;
    const lines = value.replace(/\r\n?/g, '\n').split('\n'); number++;
    const label = style === 'bullet' ? '•' : style === 'circle' && number <= 20 ? String.fromCharCode(0x245f + number) : `${number}.`;
    output.push(`${label} ${lines[0].trimStart()}`);
    for (const line of lines.slice(1)) {
      const depth = (line.match(/^\t*/) || [''])[0].length;
      output.push(depth ? '    '.repeat(depth) + '- ' + line.trimStart() : '    ' + line);
    }
  }
  return output.join('\n');
}
async function readTemplate(template) {
  const zip = await JSZip.loadAsync(await fs.readFile(template));
  const slides = {};
  for (const i of [1, 2]) slides[i] = new DOMParser().parseFromString(await zip.file(`ppt/slides/slide${i}.xml`).async('string'), 'application/xml');
  return { zip, slides };
}

function cellFields(s) {
 const fields=[];
 children(descendants(s,P,'txBody')[0],'p').forEach((p,paragraph)=>{
  const line=descendants(p,A,'t').map(t=>t.textContent).join('');
  for(const m of line.matchAll(/[^,\n]+/g)) {
   let offset=m.index;
   for(const part of m[0].split(/( {8,})/)) {
    const value=part.trim();if(value)fields.push({paragraph,start:offset+part.indexOf(value),length:value.length,value});offset+=part.length;
   }
  }
 });return fields;
}
function applyCellFields(s,values) {
 const fields=cellFields(s),ps=children(descendants(s,P,'txBody')[0],'p');
 if(!Array.isArray(values)||values.length!==fields.length||values.some(v=>typeof v!=='string'||v.length>120||/[\r\n]/.test(v)))throw new Error('셀 편성 기준 명단을 확인하세요.');
 for(let i=fields.length-1;i>=0;i--){const f=fields[i];if(values[i]===f.value)continue;let offset=0;const nodes=descendants(ps[f.paragraph],A,'t').map(t=>{const n={t,start:offset,end:offset+t.textContent.length};offset=n.end;return n;});let inserted=false;for(const n of nodes){const from=Math.max(f.start,n.start),to=Math.min(f.start+f.length,n.end);if(from>=to)continue;const old=n.t.textContent;n.t.textContent=old.slice(0,from-n.start)+(inserted?'':values[i])+old.slice(to-n.start);inserted=true;}}
}


// Original template: two rows of three cells; the service-member list is separate.
const cellSlots=[{head:0,rows:[[3,4,5,6],[15,16,17,18]]},{head:1,rows:[[7,8,9,10],[19,20,21,22]]},{head:2,rows:[[11,12,13,14]]},{head:23,rows:[[26,27,28,29],[38,39,40,41],[48]]},{head:24,rows:[[30,31,32,33],[42,43,44,45],[49]]},{head:25,rows:[[34,35,36,37],[46,47]]}];
function migrateCellGroups(names){return cellSlots.map((slot,i)=>({id:'cell-'+(i+1),leader:names[slot.head].replace(/\s*셀$/,''),members:slot.rows.flat().map(index=>names[index]).filter(name=>name.trim())}));}
function applyCellGroups(s,groups,military){
 if(!Array.isArray(groups)||groups.length!==6||groups.some((g,i)=>!g||g.id!=='cell-'+(i+1)||typeof g.leader!=='string'||!Array.isArray(g.members)||g.members.some(v=>typeof v!=='string'||/[\r\n]/.test(v))||g.members.length>100)||!Array.isArray(military)||military.some(v=>typeof v!=='string'||/[\r\n]/.test(v)))throw new Error('셀 편성 명단을 확인하세요.');
 const fields=cellFields(s),ps=children(descendants(s,P,'txBody')[0],'p'),replacements=[];
 const add=(indexes,value)=>{const first=fields[indexes[0]],last=fields[indexes.at(-1)];replacements.push({...first,length:last.start+last.length-first.start,value});};
 groups.forEach((g,i)=>{const slot=cellSlots[i];add([slot.head],g.leader.trim()?g.leader.trim()+' 셀':'');const members=g.members.filter(v=>v.trim());let used=0;slot.rows.forEach((row,j)=>{const count=j===slot.rows.length-1?members.length-used:row.length;add(row,members.slice(used,used+count).join(', '));used+=count;});});
 add([51,52,53],military.filter(v=>v.trim()).join(', '));
 replacements.sort((a,b)=>b.paragraph-a.paragraph||b.start-a.start);const changedParagraphs=new Set();
 for(const f of replacements){let offset=0;const nodes=descendants(ps[f.paragraph],A,'t').map(t=>{const n={t,start:offset,end:offset+t.textContent.length};offset=n.end;return n;});const current=nodes.map(n=>n.t.textContent).join('').slice(f.start,f.start+f.length);if(current===f.value)continue;changedParagraphs.add(f.paragraph);let inserted=false;for(const n of nodes){const from=Math.max(f.start,n.start),to=Math.min(f.start+f.length,n.end);if(from>=to)continue;const old=n.t.textContent;n.t.textContent=old.slice(0,from-n.start)+(inserted?'':f.value)+old.slice(to-n.start);inserted=true;}}
 // Anchor the three existing columns with tab stops when their content changes.
 for(const index of changedParagraphs){const p=ps[index],nodes=descendants(p,A,'t');if(!nodes.some(t=>/ {8,}/.test(t.textContent)))continue;
  for(const t of nodes)t.textContent=t.textContent.replace(/ {8,}/g,'\t');
  let pr=children(p,'pPr')[0];if(!pr){pr=p.ownerDocument.createElementNS(A,'a:pPr');p.insertBefore(pr,p.firstChild);}for(const old of children(pr,'tabLst'))pr.removeChild(old);
  const tabs=p.ownerDocument.createElementNS(A,'a:tabLst');for(const position of [99.36,198.72]){const tab=p.ownerDocument.createElementNS(A,'a:tab');tab.setAttribute('pos',String(Math.round(position*12700)));tab.setAttribute('algn','l');tabs.appendChild(tab);}pr.appendChild(tabs);
 }
}

async function defaults(template, now) {
  const { slides } = await readTemplate(template);
  const get = (i, id) => text(shape(slides[i], id));
  const date = upcomingSunday(now);
  const ads = get(1, 128).split(/^\s*\d+\.\s*/m).filter(x => x.trim()).map(x => x.trim());
  const birthdays = Array.from(get(1, 61).matchAll(/(\d+)\s*\/\s*([^\d/]+)/g), m => ({ day: m[1], name: m[2].trim() }));
  return { cellNames:cellFields(shape(slides[1],8)).map(f=>f.value), date, title: get(2, 97), preacher: get(2, 94), reference: get(2, 11), body: get(2, 11) + '\n' + get(2, 3), team: get(2, 91), cellTitle: get(1, 157), cells: get(1, 8), ads, adStyle: 'number', birthdays, prayers: monthSundays(date).map(date => ({ date, enabled: true, name: '' })) };
}
function validate(data, allowHistorical = false) {
  if (!data || !/^\d{4}-\d{2}-\d{2}$/.test(data.date)) throw new Error('주보 날짜를 확인하세요.');
  if (!allowHistorical && data.date !== upcomingSunday()) throw new Error('날짜가 변경됐습니다. 앱을 다시 열어 돌아오는 일요일을 확인하세요.');
  for (const key of ['title', 'reference', 'body', 'cellTitle', 'cells']) if (typeof data[key] !== 'string' || data[key].length > 100000) throw new Error('입력 내용을 확인하세요.');
  if (!Array.isArray(data.ads) || data.ads.length > 100 || data.ads.some(s => typeof s !== 'string' || s.length > 10000)) throw new Error('광고 내용을 확인하세요.');
  if(data.birthdayRoster?.some(b=>b.name?.trim()&&(String(b.month??'').trim()||String(b.day??'').trim())&&(!Number.isInteger(+b.month)||+b.month<1||+b.month>12||!Number.isInteger(+b.day)||+b.day<1||+b.day>new Date(Date.UTC(2000,+b.month,0)).getUTCDate())))throw new Error('기준 자료에서 생일의 월과 일을 확인하세요.');
  if (!Array.isArray(data.birthdays) || data.birthdays.length > 500) throw new Error('생일자 목록을 확인하세요.');
  for (const b of data.birthdays) if (b.name.trim() && (!/^\d+$/.test(String(b.day)) || +b.day < 1 || +b.day > 31)) throw new Error('생일은 1~31 사이의 일자로 입력하세요.');
  const allowed = monthSundays(data.date);
  if (!Array.isArray(data.prayers) || data.prayers.length !== allowed.length || data.prayers.some((p, i) => p.date !== allowed[i])) throw new Error('해당 월의 주일 목록을 확인하세요.');
  for (const p of data.prayers) if (p.enabled && !p.name.trim()) throw new Error(`${shortDate(p.date)} 기도 담당자 이름을 입력하거나 진행 체크를 해제하세요.`);
}
async function exportPpt(template, data, options = {}) {
  validate(data, options.allowHistorical === true);
  const { zip, slides } = await readTemplate(template);
  const originalReference=text(shape(slides[2],11));
  const d = dateOf(data.date);
  const ads = formatAds(data.ads, data.adStyle);
  const values = [[2,94,data.preacher??''],[2,91,data.team??''],[1,178,data.date], [1,62,`${d.getUTCFullYear()}년 ${d.getUTCMonth()+1}월 ${d.getUTCDate()}일`], [2,97,data.title], [2,11,data.reference],    [1,61,data.birthdays.filter(b=>b.name.trim()).map((b,i)=>`${+b.day}${i?'':' '}/ ${b.name.trim()}`).join('   ')], [2,108,`${d.getUTCMonth()+1}월 공동 기도 위원`]];
  for (const [i,id,value] of values) { const s=shape(slides[i],id); if(text(s)!==value) setText(s,value); }
  if(data.reference!==originalReference){const pr=descendants(shape(slides[2],11),A,'bodyPr')[0];pr.setAttribute('wrap','none');children(pr).filter(n=>['spAutoFit','normAutofit','noAutofit'].includes(n.localName)).forEach(n=>pr.removeChild(n));pr.appendChild(slides[2].createElementNS(A,'a:noAutofit'));}
  if(data.cellGroups)applyCellGroups(shape(slides[1],8),data.cellGroups,data.militaryMembers||[]);else if(data.cellNames)applyCellFields(shape(slides[1],8),data.cellNames);
  if(typeof data.cellTitle==='string'&&data.cellTitle!==text(shape(slides[1],157)))setText(shape(slides[1],157),data.cellTitle);
  layout.ads(shape(slides[1],128), data.ads, data.adStyle);
  layout.bible(slides[2], data.reference, data.body);
  layout.prayers(slides[2], data.prayers.filter(p=>p.enabled));
  for (const i of [1,2]) zip.file(`ppt/slides/slide${i}.xml`, new XMLSerializer().serializeToString(slides[i]),{createFolders:false});
  return zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});
}
const books = '창세기|출애굽기|레위기|민수기|신명기|여호수아|사사기|룻기|사무엘상|사무엘하|열왕기상|열왕기하|역대상|역대하|에스라|느헤미야|에스더|욥기|시편|잠언|전도서|아가|이사야|예레미야|예레미야애가|에스겔|다니엘|호세아|요엘|아모스|오바댜|요나|미가|나훔|하박국|스바냐|학개|스가랴|말라기|마태복음|마가복음|누가복음|요한복음|사도행전|로마서|고린도전서|고린도후서|갈라디아서|에베소서|빌립보서|골로새서|데살로니가전서|데살로니가후서|디모데전서|디모데후서|디도서|빌레몬서|히브리서|야고보서|베드로전서|베드로후서|요한일서|요한이서|요한삼서|유다서|요한계시록'.split('|');
const codes = 'gen|exo|lev|num|deu|jos|jdg|rut|1sa|2sa|1ki|2ki|1ch|2ch|ezr|neh|est|job|psa|pro|ecc|sng|isa|jer|lam|ezk|dan|hos|jol|amo|oba|jon|mic|nam|hab|zep|hag|zec|mal|mat|mrk|luk|jhn|act|rom|1co|2co|gal|eph|php|col|1th|2th|1ti|2ti|tit|phm|heb|jas|1pe|2pe|1jn|2jn|3jn|jud|rev'.split('|');
async function bible(reference) {
  const match = String(reference).trim().match(/^(.+?)\s*(\d+)\s*:\s*(\d+)(?:\s*[-~–]\s*(\d+))?$/);
  if (!match) throw new Error('요한복음 21:1-14처럼 한 장 안의 구절 범위를 입력하세요.');
  const book = match[1].replace(/\s/g,'').replace(/복음서$/,'복음');
  const index = books.indexOf(book); if(index<0) throw new Error('성경 이름을 전체 이름으로 입력하세요. 예: 요한복음, 시편');
  const chapter=+match[2], first=+match[3], last=+(match[4]||match[3]);
  if(chapter<1||chapter>150||first<1||last<first||last>176) throw new Error('장과 절 범위를 확인하세요.');
  const source=`https://www.bskorea.or.kr/bible/korbibReadpage.php?version=SAENEW&book=${codes[index]}&chap=${chapter}`;
  let response;
  try { response=await fetch(source,{signal:AbortSignal.timeout(20000)}); } catch { throw new Error('본문을 가져오지 못했습니다. 인터넷 연결을 확인하거나 직접 붙여넣어 주세요.'); }
  if(!response.ok) throw new Error('성경 조회 사이트에 연결하지 못했습니다. 본문을 직접 붙여넣어 주세요.');
  const html=await response.text();
  const body=html.match(/<div\s+id="tdBible1"[^>]*>([\s\S]*)/i)?.[1]; if(!body) throw new Error('사이트에서 본문을 찾지 못했습니다. 직접 붙여넣어 주세요.');
  const marks=Array.from(body.matchAll(/<span\s+class="number">\s*(\d+)[\s\S]*?<\/span>/gi)); const verses=new Map();
  marks.forEach((m,i)=>{
    const start=m.index+m[0].length;const end=marks[i+1]?.index??body.indexOf('</div>',start);
    if(end<start)return;
    const $=cheerio.load(body.slice(start,end));$('font.smallTitle,sup,script,style,a.comment,div.D2,[id^="D_"]').remove();
    verses.set(+m[1],$.root().text().replace(/\s+/g,' ').trim());
  });
  const lines=[];for(let v=first;v<=last;v++){if(!verses.get(v))throw new Error(`${chapter}장 ${v}절을 찾지 못했습니다.`);lines.push(`${v}. ${verses.get(v)}`);}
  return {reference:`${book} ${chapter}:${first}${last!==first?'-'+last:''}`,text:lines.join('\n'),source};
}
module.exports={migrateCellGroups,applyCellGroups,books,cellFields,applyCellFields,defaults,exportPpt,bible,upcomingSunday,monthSundays,formatAds,readTemplate,shape,text};
