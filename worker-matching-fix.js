/* Matching rules validated against the verified GTMDJD master. */
(function(){
  const nameTokens=v=>norm(v).split(' ').filter(Boolean);
  const editAffinity=(a,b)=>{
    const aa=norm(a).replace(/\s/g,''),bb=norm(b).replace(/\s/g,'');
    if(!aa||!bb)return 0;
    const previous=Array.from({length:bb.length+1},(_,i)=>i);
    for(let i=1;i<=aa.length;i++){
      const current=[i];
      for(let j=1;j<=bb.length;j++)current[j]=Math.min(current[j-1]+1,previous[j]+1,previous[j-1]+(aa[i-1]===bb[j-1]?0:1));
      for(let j=0;j<current.length;j++)previous[j]=current[j];
    }
    return 1-previous[bb.length]/Math.max(aa.length,bb.length);
  };
  const nameAffinity=(a,b)=>{
    const aa=nameTokens(a),bb=nameTokens(b);
    if(!aa.length||!bb.length)return 0;
    const aj=aa.join(''),bj=bb.join('');
    if(aj===bj)return 1;
    let score=editAffinity(a,b);
    if(aa[0]===bb[0])score=Math.max(score,0.75);
    if(aa[0]===bb[0]&&aa[1]&&bb[1]){
      let prefix=0,limit=Math.min(aa[1].length,bb[1].length);
      while(prefix<limit&&aa[1][prefix]===bb[1][prefix])prefix++;
      if(prefix>=3)score=Math.max(score,prefix>=5?0.92:0.84);
    }
    const aset=new Set(aa),bset=new Set(bb),intersection=[...aset].filter(x=>bset.has(x)).length;
    score=Math.max(score,(2*intersection)/(aset.size+bset.size));
    return score;
  };
  const sameMoney=(a,b)=>money(a)!=null&&money(a)===money(b);

  reconcile=async function(myRows,tocRows){
    const grouped=new Map();
    for(const r of tocRows){const k=key(r);if(!grouped.has(k))grouped.set(k,[]);grouped.get(k).push(r);}
    const byCheck=new Map(),byDateAmount=new Map();
    for(const m of myRows){
      const ck=digits(m.checkNumber);
      if(ck){if(!byCheck.has(ck))byCheck.set(ck,[]);byCheck.get(ck).push(m);}
      const da=[date(m.checkDate),money(m.principal),money(m.interest)??0].join('|');
      if(!byDateAmount.has(da))byDateAmount.set(da,[]);
      byDateAmount.get(da).push(m);
    }
    const used=new Set(),master=[],groups=[...grouped.values()];
    for(let i=0;i<groups.length;i++){
      const occ=groups[i],t=occ[0],ck=digits(t.checkNumber),da=[date(t.checkDate),money(t.principal),money(t.interest)??0].join('|');
      const checkCandidates=(ck?(byCheck.get(ck)||[]):[]).filter(m=>!used.has(key(m)));
      let ranked=checkCandidates.map(m=>score(t,m)).sort((a,b)=>b.s-a.s),b=ranked[0],second=ranked[1];
      let status='Not Matched',reason='No corresponding payment exists in the selected MyCases reference.',evidence='',m={};

      if(b){
        const affinity=nameAffinity(t.clientName,b.m.clientName);
        const uniqueCheck=checkCandidates.length===1;
        const strongSecondary=date(t.checkDate)===date(b.m.checkDate)||sameMoney(t.principal,b.m.principal)||affinity>=0.84;
        if(second&&second.s===b.s){
          status='Needs Review';m=b.m;evidence=b.e.join(', ');reason='Two or more MyCases candidates have the same score.';
        }else if(b.s>=90||(uniqueCheck&&strongSecondary)){
          status='Matched';m=b.m;used.add(b.k);
          const parts=[...b.e];
          if(uniqueCheck&&!parts.includes('Unique check number'))parts.push('Unique check number');
          if(affinity>=0.84&&!parts.includes('Client name'))parts.push('Client name similarity');
          evidence=parts.join(', ');reason='';
        }else if(b.s>=55){
          status='Needs Review';m=b.m;evidence=b.e.join(', ');reason='Partial match requires human confirmation.';
        }
      }else{
        const fallback=(byDateAmount.get(da)||[]).filter(x=>!used.has(key(x))).map(x=>({m:x,k:key(x),affinity:nameAffinity(t.clientName,x.clientName)})).filter(x=>x.affinity>=0.78).sort((a,b)=>b.affinity-a.affinity);
        const best=fallback[0],next=fallback[1];
        if(best&&(!next||best.affinity-next.affinity>=0.08)){
          status='Matched';m=best.m;used.add(best.k);reason='';
          evidence='Check date, Principal amount, Interest amount, Client name similarity; check number discrepancy';
        }else if(best){
          status='Needs Review';m=best.m;
          evidence='Check date, Principal amount, Interest amount, Client name similarity';
          reason='Multiple MyCases payments fit the same date, amount, and client; check number requires confirmation.';
        }
      }

      const p=money(m.principal)??money(t.principal)??0,ii=money(m.interest)??money(t.interest)??0;
      master.push({'Payment Year':(date(m.checkDate)||date(t.checkDate)).slice(-4),'Check Date':date(m.checkDate)||date(t.checkDate),'Check Number':clean(m.checkNumber)||clean(t.checkNumber),'Client Name':clean(m.clientName)||clean(t.clientName),'Client ID':clean(m.clientId)||clean(t.clientId),'Bill ID':clean(m.billId)||clean(t.billId),'Billed Amount':money(m.billed)??money(t.billed)??'','Principal Amount':p,'Interest Amount':ii,'Payment Total':Math.round((p+ii)*100)/100,'Match Status':status,'Match Evidence':evidence,'Exception Reason':reason,'Source Occurrences':occ.length,'Source References':occ.slice(0,20).map(o=>`${o.sourceFile} row ${o.sourceRow}`).join(' | '),'MyCases Source':clean(m.sourceFile),'Verification Method':status==='Matched'?'Automated verified match':'Automated comparison; human review required','Import Recommendation':status==='Matched'?'Ready after owner approval':'HOLD'});
      if((i+1)%250===0){progress(68+Math.round(((i+1)/Math.max(groups.length,1))*22),`Reconciling ${(i+1).toLocaleString()} of ${groups.length.toLocaleString()} unique payments...`);await pause();}
    }
    master.sort((a,b)=>dt(a['Check Date'])-dt(b['Check Date']));
    return master;
  };
})();