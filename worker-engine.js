/* GTMDJD background processing engine. Heavy work stays off the UI thread. */
importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');

const MAX_SCAN_COLUMNS=128;
const PROGRESS_ROW_INTERVAL=1000;
const clean=v=>String(v??'').trim();
const norm=v=>clean(v).toUpperCase().normalize('NFKD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9]+/g,' ').replace(/\s+/g,' ').trim();
const digits=v=>(clean(v).replace(/\D/g,'').replace(/^0+/,'')||'');
const money=v=>{if(v==null||v==='')return null;if(typeof v==='number'&&Number.isFinite(v))return Math.round(v*100)/100;let s=clean(v),neg=/^\(.*\)$/.test(s);s=s.replace(/[$,()\s]/g,'');const n=Number(s);return Number.isFinite(n)?Math.round((neg?-n:n)*100)/100:null;};
function date(v){
  if(v instanceof Date&&!isNaN(v))return `${String(v.getMonth()+1).padStart(2,'0')}/${String(v.getDate()).padStart(2,'0')}/${v.getFullYear()}`;
  if(typeof v==='number'&&Number.isFinite(v)){const d=new Date(Date.UTC(1899,11,30)+Math.round(v)*86400000);return `${String(d.getUTCMonth()+1).padStart(2,'0')}/${String(d.getUTCDate()).padStart(2,'0')}/${d.getUTCFullYear()}`;}
  const m=clean(v).match(/\b(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})\b/);if(!m)return'';let y=+m[3];if(y<100)y+=2000;const d=new Date(y,+m[1]-1,+m[2]);if(isNaN(d)||d.getMonth()!=+m[1]-1||d.getDate()!=+m[2])return'';return `${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}/${d.getFullYear()}`;
}
const dt=v=>{const d=date(v);return d?new Date(+d.slice(6),+d.slice(0,2)-1,+d.slice(3,5)).getTime():NaN};
const progress=(pct,text,detail={})=>postMessage({type:'progress',pct,text,detail,at:performance.now()});
const pause=()=>new Promise(r=>setTimeout(r,0));
const timed=async(name,fn)=>{const start=performance.now();progress(0,`START ${name}`,{stage:name,event:'start'});const value=await fn();const duration=Math.round((performance.now()-start)*10)/10;progress(0,`END ${name} (${duration.toLocaleString()} ms)`,{stage:name,event:'end',duration});return value;};

function headerMap(h){
  const m={};h.forEach((value,i)=>{const x=norm(value);
    if(m.clientName==null&&(/(CLIENT|PATIENT|CLAIMANT).*NAME/.test(x)||x==='NAME'))m.clientName=i;
    else if(m.clientId==null&&(/(CLIENT|PATIENT|CLAIMANT).*ID/.test(x)||x==='ID'))m.clientId=i;
    else if(m.billId==null&&/(BILL|INVOICE|CLAIM).*ID/.test(x))m.billId=i;
    else if(m.checkNumber==null&&(/(CHECK|CK).*(NO|NUM|NUMBER)/.test(x)||x==='CHECK'))m.checkNumber=i;
    else if(m.checkDate==null&&(/(CHECK|PAYMENT|PAID).*(DATE)/.test(x)||x==='DATE'))m.checkDate=i;
    else if(m.billed==null&&/(BILLED|CHARGE|BILL AMOUNT)/.test(x))m.billed=i;
    else if(m.principal==null&&/(PRINCIPAL|PRINCIPLE|RECEIVED|PAID AMOUNT|PAYMENT AMOUNT|CHECK AMOUNT)/.test(x))m.principal=i;
    else if(m.interest==null&&/(INTEREST|INT AMT|INT AMOUNT)/.test(x))m.interest=i;
  });return m;
}

function cellValue(ws,r,c){const dense=ws['!data'];const cell=dense?.[r]?.[c]??ws[XLSX.utils.encode_cell({r,c})];return cell?.v??'';}
function rowValues(ws,r,startCol,endCol){const out=[];for(let c=startCol;c<=endCol;c++)out.push(cellValue(ws,r,c));return out;}
function headerStrength(row){return row.map(norm).filter(x=>/CLIENT|PATIENT|CLAIMANT|CHECK|DATE|BILL|PRINCIP|INTEREST|AMOUNT/.test(x)).length;}
function safeColumnBounds(ws,range){
  const start=range.s.c;
  const declaredEnd=range.e.c;
  const cappedEnd=Math.min(declaredEnd,start+MAX_SCAN_COLUMNS-1);
  let observedEnd=start;
  const scanRows=Math.min(range.e.r,range.s.r+100);
  for(let r=range.s.r;r<=scanRows;r++)for(let c=start;c<=cappedEnd;c++)if(clean(cellValue(ws,r,c))!=='')observedEnd=Math.max(observedEnd,c);
  return{start,end:Math.min(cappedEnd,Math.max(observedEnd,start+15)),declaredEnd,capped:declaredEnd>cappedEnd};
}

async function parseWorksheet(ws,src,type,onStep){
  if(!ws['!ref'])return[];
  const range=XLSX.utils.decode_range(ws['!ref']);
  const bounds=safeColumnBounds(ws,range);
  progress(0,`Worksheet bounds ${src}: rows ${range.s.r+1}-${range.e.r+1}, columns ${bounds.start+1}-${bounds.end+1}${bounds.capped?` (declared through ${bounds.declaredEnd+1}; safely capped)`:''}`,{stage:'sheet-bounds',source:src,rows:range.e.r-range.s.r+1,columns:bounds.end-bounds.start+1,declaredColumns:bounds.declaredEnd-bounds.start+1});
  let headerRow=-1,header=[];
  const scanEnd=Math.min(range.e.r,range.s.r+100);
  for(let r=range.s.r;r<=scanEnd;r++){
    const row=rowValues(ws,r,bounds.start,bounds.end);
    if(headerStrength(row)>=2){headerRow=r;header=row;break;}
  }
  if(headerRow<0)return[];
  const map=headerMap(header),out=[];
  const at=(row,key)=>map[key]==null?'':row[map[key]];
  const total=Math.max(range.e.r-headerRow,1);
  for(let r=headerRow+1;r<=range.e.r;r++){
    const row=rowValues(ws,r,bounds.start,bounds.end);
    if(row.some(v=>clean(v)!=='')){
      const o={sourceFile:src,sourceType:type,sourceRow:r+1,clientName:clean(at(row,'clientName')),clientId:clean(at(row,'clientId')),billId:clean(at(row,'billId')),checkNumber:clean(at(row,'checkNumber')),checkDate:date(at(row,'checkDate')),billed:money(at(row,'billed')),principal:money(at(row,'principal')),interest:money(at(row,'interest'))??0};
      const paymentLike=digits(o.checkNumber)&&o.checkDate&&o.principal!=null;
      if(type==='MyCases'?(paymentLike||(o.clientName&&o.checkDate&&o.principal!=null)):paymentLike)out.push(o);
    }
    const done=r-headerRow;
    if(done%PROGRESS_ROW_INTERVAL===0){onStep?.(done,total);await pause();}
  }
  onStep?.(total,total);return out;
}

async function parseWorkbook(file,type,fileIndex,totalFiles){
  progress(type==='MyCases'?12:38,`Opening ${file.name}...`);await pause();
  const wb=await timed(`workbook parsing: ${file.name}`,async()=>XLSX.read(new Uint8Array(file.buffer),{type:'array',cellDates:true,dense:true}));
  const out=[],sheetTotal=wb.SheetNames.length;
  for(let index=0;index<sheetTotal;index++){
    const sheetName=wb.SheetNames[index],ws=wb.Sheets[sheetName];
    const base=type==='MyCases'?15:40,span=type==='MyCases'?20:25;
    const parsed=await timed(`sheet parsing: ${file.name} / ${sheetName}`,()=>parseWorksheet(ws,`${file.name} / ${sheetName}`,type,(done,total)=>{
      const fileFraction=(fileIndex+(index+done/Math.max(total,1))/sheetTotal)/Math.max(totalFiles,1);
      progress(base+Math.round(fileFraction*span),`Reading ${file.name}: ${sheetName} — ${done.toLocaleString()} of ${total.toLocaleString()} rows`);
    }));
    out.push(...parsed);progress(base+Math.round(((fileIndex+(index+1)/sheetTotal)/Math.max(totalFiles,1))*span),`Finished ${sheetName}: ${parsed.length.toLocaleString()} payment rows detected.`);await pause();
  }
  return out;
}

async function parseFile(file,type,fileIndex,totalFiles){
  const n=file.name.toLowerCase();
  if(n.endsWith('.csv')){
    const text=new TextDecoder().decode(file.buffer),wb=XLSX.read(text,{type:'string',dense:true});
    return parseWorksheet(wb.Sheets[wb.SheetNames[0]],file.name,type,(done,total)=>progress(type==='MyCases'?25:55,`Reading ${file.name}: ${done.toLocaleString()} of ${total.toLocaleString()} rows`));
  }
  if(/\.xlsx?$/.test(n))return parseWorkbook(file,type,fileIndex,totalFiles);
  throw new Error(`Unsupported file: ${file.name}. Use XLS, XLSX, or CSV for this verified run.`);
}

const key=r=>[digits(r.checkNumber),date(r.checkDate),money(r.principal),money(r.interest)??0].join('|');
function score(t,m){let s=0,e=[];if(digits(t.checkNumber)&&digits(t.checkNumber)===digits(m.checkNumber)){s+=55;e.push('Check number');}if(date(t.checkDate)&&date(t.checkDate)===date(m.checkDate)){s+=15;e.push('Check date');}if(money(t.principal)!=null&&money(t.principal)===money(m.principal)){s+=20;e.push('Principal amount');}if((money(t.interest)??0)===(money(m.interest)??0)){s+=3;e.push('Interest amount');}if(norm(t.clientName)&&norm(t.clientName)===norm(m.clientName)){s+=7;e.push('Client name');}if(t.billId&&clean(t.billId)===clean(m.billId)){s+=7;e.push('Bill ID');}return{s,e,m,k:key(m)};}

async function reconcile(myRows,tocRows){
  const grouped=new Map();for(const r of tocRows){const k=key(r);if(!grouped.has(k))grouped.set(k,[]);grouped.get(k).push(r);}
  const byCheck=new Map(),byDateAmount=new Map();for(const m of myRows){const ck=digits(m.checkNumber);if(ck){if(!byCheck.has(ck))byCheck.set(ck,[]);byCheck.get(ck).push(m);}const da=[date(m.checkDate),money(m.principal),money(m.interest)??0].join('|');if(!byDateAmount.has(da))byDateAmount.set(da,[]);byDateAmount.get(da).push(m);}
  const used=new Set(),master=[],groups=[...grouped.values()];
  for(let i=0;i<groups.length;i++){
    const occ=groups[i],t=occ[0],ck=digits(t.checkNumber),da=[date(t.checkDate),money(t.principal),money(t.interest)??0].join('|');let candidates=ck?(byCheck.get(ck)||[]):[];if(!candidates.length)candidates=byDateAmount.get(da)||[];
    const ranked=candidates.map(m=>score(t,m)).filter(x=>!used.has(x.k)).sort((a,b)=>b.s-a.s),b=ranked[0],second=ranked[1];let status='Not Matched',reason='No qualifying MyCases payment found.',evidence='',m={};
    if(b&&b.s>=55){m=b.m;evidence=b.e.join(', ');if(second&&second.s===b.s){status='Needs Review';reason='Two or more MyCases candidates have the same score.';}else if(b.s>=90){status='Matched';reason='';used.add(b.k);}else{status='Needs Review';reason='Partial match requires human confirmation.';}}
    const p=money(m.principal)??money(t.principal)??0,ii=money(m.interest)??money(t.interest)??0;
    master.push({'Payment Year':(date(m.checkDate)||date(t.checkDate)).slice(-4),'Check Date':date(m.checkDate)||date(t.checkDate),'Check Number':clean(m.checkNumber)||clean(t.checkNumber),'Client Name':clean(m.clientName)||clean(t.clientName),'Client ID':clean(m.clientId)||clean(t.clientId),'Bill ID':clean(m.billId)||clean(t.billId),'Billed Amount':money(m.billed)??money(t.billed)??'','Principal Amount':p,'Interest Amount':ii,'Payment Total':Math.round((p+ii)*100)/100,'Match Status':status,'Match Evidence':evidence,'Exception Reason':reason,'Source Occurrences':occ.length,'Source References':occ.slice(0,20).map(o=>`${o.sourceFile} row ${o.sourceRow}`).join(' | '),'MyCases Source':clean(m.sourceFile),'Verification Method':status==='Matched'?'Automated strict match':'Automated comparison; human review required','Import Recommendation':status==='Matched'?'Ready after owner approval':'HOLD'});
    if((i+1)%250===0){progress(68+Math.round(((i+1)/Math.max(groups.length,1))*22),`Reconciling ${(i+1).toLocaleString()} of ${groups.length.toLocaleString()} unique payments...`);await pause();}
  }
  master.sort((a,b)=>dt(a['Check Date'])-dt(b['Check Date']));return master;
}

onmessage=async e=>{
  const runStart=performance.now();
  try{
    const {myFiles,tocFiles}=e.data;progress(5,'Starting isolated background engine...');let myRows=[],tocRows=[];
    myRows=await timed('all MyCases parsing',async()=>{const rows=[];for(let i=0;i<myFiles.length;i++)rows.push(...await parseFile(myFiles[i],'MyCases',i,myFiles.length));return rows;});
    postMessage({type:'coverage-source',rows:myRows});await pause();
    tocRows=await timed('all GTMDJD parsing',async()=>{const rows=[];for(let i=0;i<tocFiles.length;i++)rows.push(...await parseFile(tocFiles[i],'TOC',i,tocFiles.length));return rows;});
    progress(68,'Building indexed comparison...');const master=await timed('index construction and matching',()=>reconcile(myRows,tocRows)),exceptions=master.filter(r=>r['Match Status']!=='Matched'),matched=master.length-exceptions.length;
    const totalDuration=Math.round((performance.now()-runStart)*10)/10;
    postMessage({type:'done',myRows,tocRowsCount:tocRows.length,master,exceptions,matched,performance:{totalDuration,rootCause:'Worksheet declared ranges included excessive empty/formatted columns; parser now caps and detects actual used columns.',maxColumns:MAX_SCAN_COLUMNS}});
  }catch(error){postMessage({type:'error',message:error?.stack||String(error)});}
};