/* Final verified reconciliation override. */
(function(){
  const tokens=v=>norm(v).split(' ').filter(Boolean);
  const affinity=(a,b)=>{const A=tokens(a),B=tokens(b);if(!A.length||!B.length)return 0;if(A.join('')===B.join(''))return 1;let n=[...new Set(A)].filter(x=>new Set(B).has(x)).length,s=(2*n)/(new Set(A).size+new Set(B).size);if(A[0]===B[0])s=Math.max(s,.75);if(A[0]===B[0]&&A[1]&&B[1]){let p=0,l=Math.min(A[1].length,B[1].length);while(p<l&&A[1][p]===B[1][p])p++;if(p>=3)s=Math.max(s,p>=5?.92:.84);}return s;};
  const total=r=>Math.round(((money(r.principal)??0)+(money(r.interest)??0))*100)/100;
  const mismatch=(t,m)=>clean(t.clientId)&&clean(m.clientId)&&clean(t.clientId)!==clean(m.clientId);
  function row(occ,t,m,status,evidence,reason){const mapped=!!m.checkNumber,p=money(m.principal)??money(t.principal)??0,i=money(m.interest)??money(t.interest)??0;return{
    'Payment Year':(date(m.checkDate)||date(t.checkDate)).slice(-4),'Check Date':date(m.checkDate)||date(t.checkDate),'Check Number':clean(m.checkNumber)||clean(t.checkNumber),'Client Name':clean(m.clientName)||clean(t.clientName),'Client ID':clean(m.clientId)||clean(t.clientId),'Bill ID':clean(m.billId)||clean(t.billId),'Billed Amount':money(m.billed)??money(t.billed)??'','Principal Amount':p,'Interest Amount':i,'Payment Total':Math.round((p+i)*100)/100,
    'Match Status':status,'Match Evidence':evidence,'Exception Reason':reason,'Import Recommendation':status==='Matched'?'Ready after owner approval':'HOLD',
    'TOC Client Name':clean(t.clientName),'TOC Client ID':clean(t.clientId),'TOC Bill ID':clean(t.billId),'TOC Check Number':clean(t.checkNumber),'TOC Check Date':date(t.checkDate),'TOC Billed Amount':money(t.billed)??'','TOC Principal Amount':money(t.principal)??'','TOC Interest Amount':money(t.interest)??0,'TOC Payment Total':total(t),
    'MyCases Client Name':mapped?clean(m.clientName):'','MyCases Client ID':mapped?clean(m.clientId):'','MyCases Bill ID':mapped?clean(m.billId):'','MyCases Check Number':mapped?clean(m.checkNumber):'','MyCases Check Date':mapped?date(m.checkDate):'','MyCases Billed Amount':mapped?(money(m.billed)??''):'','MyCases Principal Amount':mapped?(money(m.principal)??''):'','MyCases Interest Amount':mapped?(money(m.interest)??0):'','MyCases Payment Total':mapped?total(m):'',
    'Source Occurrences':occ.length,'Source References':occ.slice(0,20).map(o=>`${o.sourceFile} row ${o.sourceRow}`).join(' | '),'MyCases Source':mapped?clean(m.sourceFile):'','Verification Method':status==='Matched'?'Automated verified match':'Automated comparison; human review required'};}
  reconcile=async function(myRows,tocRows){
    const gm=new Map();for(const r of tocRows){const k=key(r);if(!gm.has(k))gm.set(k,[]);gm.get(k).push(r);}const groups=[...gm.values()],byCheck=new Map(),byDA=new Map();
    for(const m of myRows){const c=digits(m.checkNumber);if(c){if(!byCheck.has(c))byCheck.set(c,[]);byCheck.get(c).push(m);}const d=[date(m.checkDate),money(m.principal),money(m.interest)??0].join('|');if(!byDA.has(d))byDA.set(d,[]);byDA.get(d).push(m);}
    const used=new Set(),out=new Array(groups.length),later=[];
    for(let x=0;x<groups.length;x++){
      const occ=groups[x],t=occ[0],c=digits(t.checkNumber),cand=(c?(byCheck.get(c)||[]):[]).filter(m=>!used.has(key(m)));
      if(!cand.length){later.push(x);continue;}const ranked=cand.map(m=>score(t,m)).sort((a,b)=>b.s-a.s),b=ranked[0],second=ranked[1];if(!b||b.s<55){later.push(x);continue;}
      const unique=cand.length===1,strong=date(t.checkDate)===date(b.m.checkDate)||money(t.principal)===money(b.m.principal)||affinity(t.clientName,b.m.clientName)>=.84;let status='Needs Review',reason='Partial match requires human confirmation.',ev=b.e.join(', ');
      if(second&&second.s===b.s)reason='Two or more MyCases candidates have the same score.';else if(b.s>=90||(unique&&strong)){used.add(b.k);if(mismatch(t,b.m))reason=`Client ID discrepancy: TOC ${clean(t.clientId)} vs MyCases ${clean(b.m.clientId)}.`;else{status='Matched';reason='';}ev=[...b.e,'Unique check number'].filter((v,i,a)=>a.indexOf(v)===i).join(', ');}out[x]=row(occ,t,b.m,status,ev,reason);
    }
    for(let z=0;z<later.length;z++){
      const x=later[z],occ=groups[x],t=occ[0],d=[date(t.checkDate),money(t.principal),money(t.interest)??0].join('|'),cand=(byDA.get(d)||[]).filter(m=>!used.has(key(m))).map(m=>({m,k:key(m),a:affinity(t.clientName,m.clientName)})).filter(v=>v.a>=.78).sort((a,b)=>b.a-a.a),b=cand[0],n=cand[1];let status='Not Matched',reason='No corresponding payment exists in the selected MyCases reference.',ev='',m={};
      if(b&&(!n||b.a-n.a>=.08)){m=b.m;used.add(b.k);ev='Check date, Principal amount, Interest amount, Client name similarity; check number discrepancy';if(mismatch(t,m)){status='Needs Review';reason=`Client ID discrepancy: TOC ${clean(t.clientId)} vs MyCases ${clean(m.clientId)}.`;}else{status='Matched';reason='';}}
      else if(b){m=b.m;status='Needs Review';ev='Check date, Principal amount, Interest amount, Client name similarity';reason='Multiple MyCases payments fit the same date, amount, and client; check number requires confirmation.';}
      out[x]=row(occ,t,m,status,ev,reason);if((z+1)%250===0)await pause();
    }
    out.sort((a,b)=>dt(a['Check Date'])-dt(b['Check Date']));return out;
  };
})();