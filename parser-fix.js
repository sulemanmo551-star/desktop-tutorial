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

async function worksheetRowsBatched(ws,onProgress){
  const ref=ws['!ref'];
  if(!ref)return[];
  const range=XLSX.utils.decode_range(ref),rows=[];
  const dense=ws['!data'];
  for(let r=range.s.r;r<=range.e.r;r++){
    const row=[];
    for(let c=range.s.c;c<=range.e.c;c++){
      const cell=dense?.[r]?.[c]??ws[XLSX.utils.encode_cell({r,c})];
      row.push(cell?.v??'');
    }
    rows.push(row);
    if((r-range.s.r+1)%200===0){
      if(onProgress)onProgress(r-range.s.r+1,range.e.r-range.s.r+1);
      await yieldToBrowser();
    }
  }
  return rows;
}

async function normalizeRowsBatched(rows,src,type){
  let hi=-1;
  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    if(Array.isArray(r)&&r.some(v=>clean(v)!=='')&&r.map(norm).filter(x=>/CLIENT|PATIENT|CLAIMANT|CHECK|DATE|BILL|PRINCIP|INTEREST|AMOUNT/.test(x)).length>=2){hi=i;break;}
    if(i>0&&i%300===0)await yieldToBrowser();
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
    if(i>hi&&i%200===0)await yieldToBrowser();
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
    const s=book.SheetNames[index],ws=book.Sheets[s];
    const base=type==='TOC'?45:20,span=type==='TOC'?24:20;
    progress(base+Math.round((index/total)*span),`Reading ${f.name}: sheet ${index+1} of ${total} — ${s}`);
    log(`Reading ${f.name} sheet ${index+1}/${total}: ${s}`);
    await yieldToBrowser();
    const rows=await worksheetRowsBatched(ws,(done,totalRows)=>{
      progress(base+Math.round(((index+Math.min(done/Math.max(totalRows,1),0.99))/total)*span),`Reading ${s}: ${done.toLocaleString()} of ${totalRows.toLocaleString()} rows`);
    });
    const parsed=await normalizeRowsBatched(rows,`${f.name} / ${s}`,type);
    out.push(...parsed);
    log(`Finished ${s}: ${parsed.length} payment rows detected.`);
    await yieldToBrowser();
  }
  return out;
}
