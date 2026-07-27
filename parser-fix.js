// Parser patch for real GTMDJD Office workbooks.
// Loaded after app.js so these functions replace the original implementations.
const yieldToBrowser=()=>new Promise(resolve=>setTimeout(resolve,0));

function date(v){
  if(v instanceof Date&&!isNaN(v))return `${String(v.getMonth()+1).padStart(2,'0')}/${String(v.getDate()).padStart(2,'0')}/${v.getFullYear()}`;
  if(typeof v==='number'&&Number.isFinite(v)){
    const d=new Date(Date.UTC(1899,11,30)+Math.round(v)*86400000);
    return `${String(d.getUTCMonth()+1).padStart(2,'0')}/${String(d.getUTCDate()).padStart(2,'0')}/${d.getUTCFullYear()}`;
  }
  const m=clean(v).match(/\b(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})\b/);
  if(!m)return'';
  let y=+m[3];if(y<100)y+=2000;
  const d=new Date(y,+m[1]-1,+m[2]);
  if(isNaN(d)||d.getMonth()!=+m[1]-1||d.getDate()!=+m[2])return'';
  return `${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}/${d.getFullYear()}`;
}

function headerMap(h){
  const m={};
  h.forEach((x,i)=>{
    x=norm(x);
    if(m.clientName==null&&(/(CLIENT|PATIENT|CLAIMANT).*NAME/.test(x)||x==='NAME'))m.clientName=i;
    else if(m.clientId==null&&(/(CLIENT|PATIENT|CLAIMANT).*ID/.test(x)||x==='ID'))m.clientId=i;
    else if(m.billId==null&&/(BILL|INVOICE|CLAIM).*ID/.test(x))m.billId=i;
    else if(m.checkNumber==null&&(/(CHECK|CK).*(NO|NUM|NUMBER)/.test(x)||x==='CHECK'))m.checkNumber=i;
    else if(m.checkDate==null&&(/(CHECK|PAYMENT|PAID).*(DATE)/.test(x)||x==='DATE'))m.checkDate=i;
    else if(m.billed==null&&/(BILLED|CHARGE|BILL AMOUNT)/.test(x))m.billed=i;
    else if(m.principal==null&&/(PRINCIPAL|PRINCIPLE|RECEIVED|PAID AMOUNT|PAYMENT AMOUNT|CHECK AMOUNT)/.test(x))m.principal=i;
    else if(m.interest==null&&/(INTEREST|INT AMT|INT AMOUNT)/.test(x))m.interest=i;
  });
  return m;
}

function normalizeRows(rows,src,type){
  let hi=-1;
  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    if(!Array.isArray(r)||!r.some(v=>clean(v)!==''))continue;
    if(r.map(norm).filter(x=>/CLIENT|PATIENT|CLAIMANT|CHECK|DATE|BILL|PRINCIP|INTEREST|AMOUNT/.test(x)).length>=2){hi=i;break;}
  }
  if(hi<0)return[];
  const m=headerMap(rows[hi]||[]),out=[];
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i]||[];
    if(!Array.isArray(r)||!r.some(v=>clean(v)!==''))continue;
    const at=k=>m[k]==null?'':r[m[k]];
    const o={sourceFile:src,sourceType:type,sourceRow:i+1,clientName:clean(at('clientName')),clientId:clean(at('clientId')),billId:clean(at('billId')),checkNumber:clean(at('checkNumber')),checkDate:date(at('checkDate')),billed:money(at('billed')),principal:money(at('principal')),interest:money(at('interest'))??0,raw:r.map(clean).join(' | ')};
    const paymentLike=digits(o.checkNumber)&&o.checkDate&&o.principal!=null;
    if(type==='MyCases'?(paymentLike||(o.clientName&&o.checkDate&&o.principal!=null)):paymentLike)out.push(o);
  }
  return out;
}

async function normalizeRowsBatched(rows,src,type){
  let hi=-1;
  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    if(Array.isArray(r)&&r.some(v=>clean(v)!=='')&&r.map(norm).filter(x=>/CLIENT|PATIENT|CLAIMANT|CHECK|DATE|BILL|PRINCIP|INTEREST|AMOUNT/.test(x)).length>=2){hi=i;break;}
    if(i>0&&i%400===0)await yieldToBrowser();
  }
  if(hi<0)return[];
  const m=headerMap(rows[hi]||[]),out=[];
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i]||[];
    if(Array.isArray(r)&&r.some(v=>clean(v)!=='')){
      const at=k=>m[k]==null?'':r[m[k]];
      const o={sourceFile:src,sourceType:type,sourceRow:i+1,clientName:clean(at('clientName')),clientId:clean(at('clientId')),billId:clean(at('billId')),checkNumber:clean(at('checkNumber')),checkDate:date(at('checkDate')),billed:money(at('billed')),principal:money(at('principal')),interest:money(at('interest'))??0,raw:r.map(clean).join(' | ')};
      const paymentLike=digits(o.checkNumber)&&o.checkDate&&o.principal!=null;
      if(type==='MyCases'?(paymentLike||(o.clientName&&o.checkDate&&o.principal!=null)):paymentLike)out.push(o);
    }
    if(i>hi&&i%250===0)await yieldToBrowser();
  }
  return out;
}

async function parseSheet(f,type){
  progress(type==='TOC'?43:18,`Opening ${f.name}...`);
  await yieldToBrowser();
  const bytes=new Uint8Array(await f.arrayBuffer());
  await yieldToBrowser();
  const book=XLSX.read(bytes,{type:'array',cellDates:true,dense:true});
  const out=[],total=book.SheetNames.length;
  for(let index=0;index<total;index++){
    const s=book.SheetNames[index];
    const pct=type==='TOC'?45+Math.round(((index+1)/total)*24):20+Math.round(((index+1)/total)*20);
    progress(pct,`Reading ${f.name}: sheet ${index+1} of ${total} — ${s}`);
    log(`Reading ${f.name} sheet ${index+1}/${total}: ${s}`);
    await yieldToBrowser();
    const rows=XLSX.utils.sheet_to_json(book.Sheets[s],{header:1,defval:'',raw:true,blankrows:false});
    await yieldToBrowser();
    const parsed=await normalizeRowsBatched(rows,`${f.name} / ${s}`,type);
    out.push(...parsed);
    log(`Finished ${s}: ${parsed.length} payment rows detected.`);
    await yieldToBrowser();
  }
  return out;
}
