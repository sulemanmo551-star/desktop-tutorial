// Fast indexed reconciliation patch.
// Replaces the O(TOC × MyCases) full scan with indexed candidate lookup.
(function(){
  function candidateScore(toc,m){
    let s=0,e=[];
    if(digits(toc.checkNumber)&&digits(toc.checkNumber)===digits(m.checkNumber)){s+=55;e.push('Check number');}
    if(date(toc.checkDate)&&date(toc.checkDate)===date(m.checkDate)){s+=15;e.push('Check date');}
    if(money(toc.principal)!=null&&money(toc.principal)===money(m.principal)){s+=20;e.push('Principal amount');}
    if((money(toc.interest)??0)===(money(m.interest)??0)){s+=3;e.push('Interest amount');}
    if(norm(toc.clientName)&&norm(toc.clientName)===norm(m.clientName)){s+=7;e.push('Client name');}
    if(toc.billId&&clean(toc.billId)===clean(m.billId)){s+=7;e.push('Bill ID');}
    return {m,k:key(m),s,e};
  }

  function makeIndexes(){
    const byCheck=new Map(),byDateAmount=new Map();
    for(const m of S.myRows){
      const ck=digits(m.checkNumber);
      if(ck){if(!byCheck.has(ck))byCheck.set(ck,[]);byCheck.get(ck).push(m);}
      const da=[date(m.checkDate),money(m.principal),money(m.interest)??0].join('|');
      if(!byDateAmount.has(da))byDateAmount.set(da,[]);
      byDateAmount.get(da).push(m);
    }
    return {byCheck,byDateAmount};
  }

  compare=function(toc,used,indexes){
    const ck=digits(toc.checkNumber);
    const da=[date(toc.checkDate),money(toc.principal),money(toc.interest)??0].join('|');
    let candidates=ck?(indexes.byCheck.get(ck)||[]):[];
    if(!candidates.length)candidates=indexes.byDateAmount.get(da)||[];
    const scored=candidates.map(m=>candidateScore(toc,m)).filter(x=>!used.has(x.k)).sort((a,b)=>b.s-a.s);
    const b=scored[0],second=scored[1];
    if(!b||b.s<55)return{status:'Not Matched',reason:'No qualifying MyCases payment found.',evidence:''};
    if(second&&second.s===b.s)return{status:'Needs Review',match:b.m,reason:'Two or more MyCases candidates have the same score.',evidence:b.e.join(', ')};
    if(b.s>=90){used.add(b.k);return{status:'Matched',match:b.m,reason:'',evidence:b.e.join(', ')}};
    return{status:'Needs Review',match:b.m,reason:'Partial match requires human confirmation.',evidence:b.e.join(', ')};
  };

  build=function(){
    const grouped=new Map();
    for(const r of S.tocRows){
      const k=key(r);
      if(!grouped.has(k))grouped.set(k,[]);
      grouped.get(k).push(r);
    }
    const indexes=makeIndexes(),used=new Set(),out=[];
    let done=0,total=grouped.size;
    for(const occ of grouped.values()){
      const t=occ[0],x=compare(t,used,indexes),m=x.match||{};
      const p=money(m.principal)??money(t.principal)??0;
      const ii=money(m.interest)??money(t.interest)??0;
      out.push({'Payment Year':(date(m.checkDate)||date(t.checkDate)).slice(-4),'Check Date':date(m.checkDate)||date(t.checkDate),'Check Number':clean(m.checkNumber)||clean(t.checkNumber),'Client Name':clean(m.clientName)||clean(t.clientName),'Client ID':clean(m.clientId)||clean(t.clientId),'Bill ID':clean(m.billId)||clean(t.billId),'Billed Amount':money(m.billed)??money(t.billed)??'','Principal Amount':p,'Interest Amount':ii,'Payment Total':Math.round((p+ii)*100)/100,'Match Status':x.status,'Match Evidence':x.evidence,'Exception Reason':x.reason,'Source Occurrences':occ.length,'Source References':occ.map(o=>`${o.sourceFile} row ${o.sourceRow}`).join(' | '),'Raw Source Evidence':occ.map(o=>o.raw).join(' || '),'MyCases Source':clean(m.sourceFile),'Verification Method':x.status==='Matched'?'Automated strict match':'Automated comparison; human review required','Import Recommendation':x.status==='Matched'?'Ready after owner approval':'HOLD'});
      done++;
      if(done%250===0)progress(75+Math.round((done/Math.max(total,1))*15),`Reconciling ${done.toLocaleString()} of ${total.toLocaleString()} unique payments...`);
    }
    log(`Indexed reconciliation completed for ${total.toLocaleString()} unique TOC payments.`);
    return out.sort((a,b)=>dt(a['Check Date'])-dt(b['Check Date']));
  };
})();
