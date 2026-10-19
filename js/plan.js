// Excel plan / liste şablonları (bölümlü satırlar, aylık X işaretleri): iskelet çözümleme, parça parça içerik üretimi, biçim korunarak yazma
const AYL=['ocak','şubat','mart','nisan','mayıs','haziran','temmuz','ağustos','eylül','ekim','kasım','aralık'],mIdx=t=>typeof t=='string'?AYL.indexOf(t.toLocaleLowerCase('tr').trim()):-1;
const slv2=c=>c.isMerged&&c.master&&c.master.address!=c.address,cT=c=>{const v=cv(c);return v===null||v===undefined?'':String(v).replace(/\r/g,'').trim()};
const cnt0=(ws,r)=>{let n=0;if(r>ws.rowCount)return 0;ws.getRow(r).eachCell(c=>{if(!slv2(c)&&cT(c))n++});return n};
const mrg=ws=>((ws.model&&ws.model.merges)||[]).map(m=>{const x=String(m).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);return x?{s:m,l:ws.getColumn(x[1]).number,t:+x[2],r:ws.getColumn(x[3]).number,b:+x[4]}:null}).filter(Boolean);
// Başlık satırı: kendisi dolu, hemen altındaki 3 satır boş değil (etiket/değer bloğu ve boş ayırıcılar elenir)
function xpScan(ws){const M=mrg(ws);let hr0=0;for(let r=1;r<=Math.min(ws.rowCount,40);r++){if(cnt0(ws,r)>=4&&[1,2,3].every(k=>r+k<=ws.rowCount&&cnt0(ws,r+k)>0)){hr0=r;break}}if(!hr0)return null;
 let hr1=hr0;M.forEach(m=>{if(m.t==hr0&&m.b>hr1&&m.b<=hr0+3)hr1=m.b});
 if(hr1==hr0){let a=0,b=0;ws.getRow(hr0).eachCell(c=>{if(mIdx(cT(c))>=0)a++});ws.getRow(hr0+1).eachCell(c=>{if(mIdx(cT(c))>=0)b++});if(!a&&b>=6)hr1=hr0+1}
 const cols=[],mon={};let lastCol=0;for(let n=1;n<=Math.min(ws.columnCount,60);n++){let t=cT(ws.getCell(hr1,n));if(!t||slv2(ws.getCell(hr1,n))&&hr1==hr0)t=t||'';if(!t&&hr1>hr0)t=cT(ws.getCell(hr0,n));
  const mi=mIdx(t);if(mi>=0){mon[mi+1]=n;lastCol=n;continue}if(t){cols.push({n,L:ws.getColumn(n).letter,t:t.replace(/\s+/g,' ')});lastCol=n}}
 if(cols.length<2)return null;
 let last=hr1;for(let r=hr1+1;r<=ws.rowCount;r++)if(cnt0(ws,r))last=r;let fs=last+1;while(fs-1>hr1&&cnt0(ws,fs-1)==1)fs--;
 const mset=new Set(Object.values(mon)),items=[],foot=[],sp=r=>{const m=M.find(x=>x.t==r&&x.r>x.l);return m?{l:m.l,rr:m.r}:{l:1,rr:lastCol}};
 for(let r=hr1+1;r<=last;r++){const k=cnt0(ws,r);if(!k)continue;
  if(r>=fs){const f=ws.getRow(r),c1=(()=>{let x=0;f.eachCell(c=>{if(!x&&!slv2(c)&&cT(c))x=c.col});return x})(),s=sp(r);foot.push({r,t:cT(ws.getCell(r,c1)),l:s.l,rr:s.rr});continue}
  if(k==1){let t='',c1=0;ws.getRow(r).eachCell(c=>{if(!slv2(c)&&cT(c)){t=cT(c);c1=c.col}});const s=sp(r);items.push({r,sec:t,l:s.l,rr:s.rr});continue}
  const c={},m=[];cols.forEach(x=>{const v=cT(ws.getCell(r,x.n));if(v)c[x.L]=v});Object.entries(mon).forEach(([i,n])=>{if(cT(ws.getCell(r,n)))m.push(+i)});items.push({r,c,m,sig:Object.keys(c).join('')})}
 return{name:ws.name,hr0,hr1,lastCol,cols,mon,items,foot,title:(()=>{const o=[];for(let r=1;r<hr0;r++)ws.getRow(r).eachCell(c=>{if(!slv2(c)&&cT(c))o.push({a:c.address,t:cT(c)})});return o})()}}
async function xpAnalyze(buf){const wb=new ExcelJS.Workbook();await wb.xlsx.load(buf);const l=wb.worksheets.map(w=>({w,p:xpScan(w)})).filter(x=>x.p&&x.p.items.length>=3).sort((a,b)=>b.p.items.length-a.p.items.length)[0];return l?l.p:null}
const xpLine=i=>i.sec!==undefined?`[BÖLÜM] ${i.sec}`:`[SATIR] ${JSON.stringify(Object.fromEntries(Object.entries(i.c).map(([k,v])=>[k,v.slice(0,300)])))}${i.m.length?' AYLAR='+i.m.join(','):''}`;
const xpSummary=p=>({name:p.name,cols:p.cols.map(c=>({L:c.L,t:c.t})),months:Object.keys(p.mon).length,n:p.items.length,secs:(()=>{const o=[];let k=null;p.items.forEach(i=>{if(i.sec!==undefined){k={name:i.sec,n:0};o.push(k)}else if(k)k.n++});return o})()});
const xpFilter=(p,skip)=>{if(!skip||!skip.length)return p;let off=false;return{...p,items:p.items.filter(i=>{if(i.sec!==undefined)off=skip.includes(i.sec);return!off})}};
// Parçalar: bölüm sınırlarında bölünür; her parça önceki bölüm adını taşır
function xpChunks(p,n=16){const ch=[];let cur=[],cn=0,prev='',last='';const push=()=>{if(cur.length)ch.push({items:cur,prev});cur=[];cn=0;prev=last};
 p.items.forEach(it=>{if(it.sec!==undefined){if(cn>=Math.floor(n/2))push();last=it.sec}cur.push(it);if(it.sec===undefined){cn++;if(cn>=n)push()}});push();return ch}
const hEst=(ws,cells,w0)=>{let m=1;cells.forEach(([n,t])=>{if(!t)return;const w=(w0||(ws.getColumn(n).width||10))*1.05,ln=String(t).split('\n').reduce((a,l)=>a+Math.max(1,Math.ceil(l.length/w)),0);m=Math.max(m,ln)});return Math.min(409,Math.max(27.6,m*13+8))};
async function xpFill(buf,D){xlSt='plan şablonu okuma';const wb=new ExcelJS.Workbook();await wb.xlsx.load(buf);
 const sh=wb.worksheets.map(w=>({w,p:xpScan(w)})).filter(x=>x.p&&x.p.items.length>=3).sort((a,b)=>b.p.items.length-a.p.items.length)[0];if(!sh)throw new Error('Excel dosyasında plan tablosu bulunamadı');
 const{w:ws,p}=sh,F=D.firma||{},xp=D.fill&&D.fill.xp,LC=p.lastCol;xlSt='firma bilgileri';xlHead(ws,p.hr0,F,{...D,sheetName:ws.name.replace(/^[A-ZÇĞİÖŞÜ]+\s/,'')||ws.name});
 const sty=r=>{const o={};for(let n=1;n<=LC;n++)o[n]=JSON.parse(JSON.stringify(ws.getCell(r,n).style||{}));return o};
 const rp=p.items.filter(i=>i.sec===undefined).map(i=>({sig:i.sig,st:sty(i.r)})),sp=p.items.find(i=>i.sec!==undefined),sP=sp&&{st:sty(sp.r),l:sp.l,rr:sp.rr,h:ws.getRow(sp.r).height||21.75},
  fP=p.foot.map(f=>({st:sty(f.r),l:f.l,rr:f.rr,t:f.t})),mk={},um={};
 p.items.filter(i=>i.sec===undefined).forEach(i=>Object.entries(p.mon).forEach(([k,n])=>{const s=JSON.parse(JSON.stringify(ws.getCell(i.r,n).style||{}));if(i.m.includes(+k)){if(!mk[n])mk[n]=s}else if(!um[n])um[n]=s}));
 const anyMk=Object.values(mk)[0],anyUm=Object.values(um)[0];
 xlSt='satır yazma';mrg(ws).filter(m=>m.t>p.hr1).forEach(m=>ws.unMergeCells(m.s));
 for(let r=p.hr1+1;r<=ws.rowCount;r++){const rw=ws.getRow(r);for(let n=1;n<=Math.max(LC,ws.columnCount);n++){const c=rw.getCell(n);c.value=null;c.style={}}rw.height=undefined}
 const rows=xp&&Array.isArray(xp.rows)?xp.rows:p.items.map(i=>i.sec!==undefined?{b:i.sec}:{h:i.c,m:i.m}),sw=x=>typeof x=='string'?swapFirm(x,F,{}):x;let R=p.hr1+1;
 rows.forEach(it=>{if(it.b!==undefined){if(sP){for(let n=sP.l;n<=sP.rr;n++)ws.getCell(R,n).style=JSON.parse(JSON.stringify(sP.st[n]||{}));ws.getCell(R,sP.l).value=String(it.b);if(sP.rr>sP.l)ws.mergeCells(R,sP.l,R,sP.rr);ws.getRow(R).height=sP.h}else ws.getCell(R,1).value=String(it.b);R++;return}
  const h=it.h||{},sig=Object.keys(h).filter(k=>String(h[k]??'').trim()).join(''),pr=rp.find(x=>x.sig==sig)||rp[0],cells=[];
  for(let n=1;n<=LC;n++)ws.getCell(R,n).style=JSON.parse(JSON.stringify(pr.st[n]||{}));
  p.cols.forEach(c=>{let v=h[c.L];if(v===undefined||v===null||v==='')return;v=sw(String(v));cells.push([c.n,v]);ws.getCell(R,c.n).value=/^\d+$/.test(v)?+v:v});
  const ms=new Set((it.m||[]).map(Number));Object.entries(p.mon).forEach(([k,n])=>{const on=ms.has(+k),c=ws.getCell(R,n);c.style=JSON.parse(JSON.stringify(on?(mk[n]||anyMk||{}):(um[n]||anyUm||{})));if(on)c.value='X'});
  ws.getRow(R).height=hEst(ws,cells);R++});
 const last=R-1,fl=(xp&&Array.isArray(xp.foot)?xp.foot:p.foot.map(f=>f.t));
 fl.forEach((t,i)=>{const f=fP[Math.min(i,fP.length-1)]||{st:{},l:1,rr:LC};for(let n=f.l;n<=f.rr;n++)ws.getCell(R,n).style=JSON.parse(JSON.stringify(f.st[n]||{}));ws.getCell(R,f.l).value=sw(String(t));if(f.rr>f.l)ws.mergeCells(R,f.l,R,f.rr);let w=0;for(let n=f.l;n<=f.rr;n++)w+=ws.getColumn(n).width||10;ws.getRow(R).height=hEst(ws,[[f.l,t]],w);R++});
 const ps=ws.pageSetup;ps.printArea=`A1:${ws.getColumn(LC).letter}${Math.max(R-1,last)}`;if(ps.fitToPage&&ps.fitToHeight==1&&rows.length>30)ps.fitToHeight=0;if(Array.isArray(ws.rowBreaks))ws.rowBreaks=ws.rowBreaks.filter(b=>b.id<last);
 xlSt='kaydetme';wb.calcProperties.fullCalcOnLoad=true;return new Blob([await xlWrite(wb)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
