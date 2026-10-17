// Şablon iskeleti: Word (.docx) ve Excel (.xlsx) dosyalarının biçimini koruyarak doldurma
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',wq=(n,t)=>[...n.getElementsByTagNameNS(W,t)],wk=(n,t)=>[...n.childNodes].filter(c=>c.localName==t);
const ptx=p=>wq(p,'t').map(t=>t.textContent).join('');
const LBL=[['ad',/(firma|iş\s?yeri|işyeri|kurum|şirket).{0,14}(ad|unvan|ünvan)|^firma|unvan|ünvan/i],['sgk',/sgk|sicil/i],['adres',/adres/i],['th',/tehlike sınıf/i],['isveren',/işveren/i],['uzman',/uzman|hazırlayan/i],['hekim',/hekim/i],['tarih',/tarih/i]];
function lbl(t,f,date){t=String(t||'').replace(/[:：_\s.]+$/,'').trim();if(!t||t.length>45)return'';for(const[k,r]of LBL)if(r.test(t))return k=='th'?(TH[f.th]||''):k=='tarih'?new Date(date).toLocaleDateString('tr'):(f[k]||'');return''}
const hk=t=>{t=String(t||'').toLocaleLowerCase('tr');return/sıra|^no$|^s\.?\s?no/.test(t)?'no':/önlem|kontrol|düzeltici/.test(t)?'t':/skor|puan|derece|risk değeri/.test(t)?'skor':/seviye|sınıf|düzey|kategori/.test(t)?'lv':/olasılık/.test(t)?'o':/frekans|sıklık/.test(t)?'f':/şiddet/.test(t)?'s':/tehlike/.test(t)?'h':/risk|sonuç|zarar/.test(t)?'r':/alan|bölüm|faaliyet|yer/.test(t)?'a':''};
const riskCell=(k,r,i,mt,ht)=>k?({no:i+1,a:r.a,h:r.h,r:r.r,o:r.o,f:r.f,s:r.s,skor:sk(r,mt),lv:lv(sk(r,mt),mt),t:r.t})[k]??'':(r.x&&r.x[ht])||'';
// ---------- WORD ----------
async function dxLoad(buf){const z=await JSZip.loadAsync(buf),doc=new DOMParser().parseFromString(await z.file('word/document.xml').async('string'),'application/xml'),body=doc.getElementsByTagNameNS(W,'body')[0],TBL=wk(body,'tbl');
 const T=TBL.map((t,i)=>{const rows=wk(t,'tr'),tx=rows.map(r=>wk(r,'tc').map(c=>wq(c,'p').map(ptx).join(' ').trim())),hd=tx[0]||[],bd=tx.slice(1);
  const rk=hd.filter(x=>hk(x)).length>=3&&hd.some(x=>/tehlike|risk/i.test(x)),list=rows.length>=2&&hd.length>=3&&((hd.every(x=>x)&&bd.every(r=>!r[0]||/^\d+[.)]?$/.test(r[0])))||rk);let b=t.previousSibling;while(b&&!(b.localName=='p'&&ptx(b).trim()))b=b.previousSibling;
  return{i,list,hd,n:bd.length,before:b?ptx(b).trim().slice(0,80):''}});
 const sl=[];wq(body,'p').forEach((p,k)=>{let ti=-1,r=-1,c=-1;const tc=p.parentNode.localName=='tc'?p.parentNode:null;
  if(tc){const tr=tc.parentNode,tl=tr.parentNode;ti=TBL.indexOf(tl);r=wk(tl,'tr').indexOf(tr);c=wk(tr,'tc').indexOf(tc);if(wq(tc,'p')[0]!==p)return}
  if(ti>=0&&T[ti].list)return;const t=ptx(p).trim();if(ti<0&&!t)return;sl.push({id:'s'+k,k,t,ti,r,c})});
 const rt=T.find(t=>t.list&&t.hd.filter(hk).length>=3),method=rt?(rt.hd.some(h=>hk(h)=='f')?'fk':'mx'):'';return{z,doc,body,TBL,pub:{kind:'docx',T,sl,method}}}
const dxAnalyze=async b=>(await dxLoad(b)).pub;
const dxText=p=>[...p.T.filter(t=>t.list).map(t=>`[TABLO ${t.i}] başlıklar: ${t.hd.join(' | ')} (mevcut satır: ${t.n}; öncesindeki başlık: ${t.before})`),...p.sl.map(s=>`[${s.id}]${s.ti>=0?` (tablo ${s.ti} satır ${s.r} sütun ${s.c})`:''} ${s.t||'(boş)'}`)].join('\n').slice(0,14000);
function setP(p,text){const d=p.ownerDocument,rs=wq(p,'r');let r=rs[0];if(!r){r=d.createElementNS(W,'w:r');const rp=wq(p,'rPr')[0];if(rp&&rp.parentNode.localName=='pPr')r.appendChild(rp.cloneNode(true));p.appendChild(r)}
 [...r.childNodes].forEach(n=>{if(['t','tab','br'].includes(n.localName))r.removeChild(n)});const t=d.createElementNS(W,'w:t');t.setAttribute('xml:space','preserve');t.textContent=text;r.appendChild(t);rs.slice(1).forEach(x=>x.parentNode&&x.parentNode.removeChild(x))}
// D: {firma,date,fill:{slots,rows},risk:{rows,method},teams:(hd,before)=>rows|null}
async function dxFill(buf,D){const{z,doc,body,TBL,pub}=await dxLoad(buf),P=wq(body,'p'),F=D.firma||{},fl=D.fill||{slots:{},rows:{}},by={};pub.sl.forEach(s=>by[s.ti+'.'+s.r+'.'+s.c]=s);
 pub.sl.forEach(s=>{const p=P[s.k];let v=fl.slots[s.id];
  if(v===undefined){ // etiket tabanlı otomatik doldurma
   if(s.ti>=0&&s.c>0){const l=by[s.ti+'.'+s.r+'.'+(s.c-1)];const x=l&&lbl(l.t,F,D.date);if(x)v=x}
   else{const q=s.t.match(/^(.{2,40}?)\s*[:：]\s*[_\s.…]*$/);if(q){const x=lbl(q[1],F,D.date);if(x)v=q[1]+': '+x}}}
  if(v!==undefined)setP(p,String(v))});
 pub.T.forEach(t=>{if(!t.list)return;let rows=fl.rows[t.i];
  if(D.teams){const x=D.teams(t.hd,t.before);if(x)rows=x}
  if(D.risk&&t.hd.filter(h=>hk(h)).length>=3){const ks=t.hd.map(hk);rows=D.risk.rows.map((r,i)=>ks.map((k,j)=>riskCell(k,r,i,D.risk.method,t.hd[j])))}
  if(!rows)return;const tb=TBL[t.i],trs=wk(tb,'tr'),pr=trs[trs.length-1].cloneNode(true);trs.slice(1).forEach(x=>tb.removeChild(x));
  rows.forEach(rw=>{const tr=pr.cloneNode(true);wk(tr,'tc').forEach((tc,ci)=>{const ps=wq(tc,'p');setP(ps[0],String(rw[ci]??''));ps.slice(1).forEach(x=>tc.removeChild(x))});tb.appendChild(tr)})});
 let x=new XMLSerializer().serializeToString(doc);if(!x.startsWith('<?xml'))x='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+x;z.file('word/document.xml',x);return z.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'})}
// ---------- EXCEL ----------
const slave=c=>c.isMerged&&c.master&&c.master.address!=c.address&&c.master.row==c.row; // yalnızca yatay birleşik hücre kalıntısı; dikey birleşik başlıklar gerçek başlıktır
const cv=c=>{const v=c.value;return v&&v.richText?v.richText.map(x=>x.text).join(''):v&&typeof v=='object'&&v.result!==undefined?v.result:v&&v.text?v.text:v};
function xlHdr(ws,risk){let best=0,row=0;for(let r=1;r<=Math.min(ws.rowCount,40);r++){let n=0;ws.getRow(r).eachCell(c=>{if(slave(c))return;const v=cv(c);if(typeof v=='string'&&v.trim()&&v.length<=40&&(!risk||hk(v)))n++});if(n>best){best=n;row=r}}return best>=(risk?2:3)?row:0}
const xlSheet=(wb,risk)=>wb.worksheets.map(w=>({w,hr:xlHdr(w,risk)})).filter(x=>x.hr).sort((a,b)=>b.w.rowCount-a.w.rowCount)[0];
const hdrLike=(ws,r)=>{let n=0;ws.getRow(r).eachCell(c=>{if(slave(c))return;const v=cv(c);if(typeof v=='string'&&hk(v))n++});return n>=2};
// Sütunlar: iki seviyeli başlıkta boş alt başlık üst satırdan tamamlanır; tekrar eden anlamlar (ör. "önlem sonrası" grubu) ilk gruptan sonra "dup" sayılır
const xlCols=(ws,hr,ovr={})=>{const up=hr>1&&hdrLike(ws,hr-1),cols=[],seen={},cn={},ho=(o,k)=>Object.prototype.hasOwnProperty.call(o,k);
 for(let n=1;n<=Math.min(ws.columnCount,80);n++){let c=ws.getCell(hr,n),t=slave(c)?null:cv(c);if((t===null||t===undefined||t==='')&&up){const u=ws.getCell(hr-1,n);t=slave(u)?null:cv(u)}
  if(t===null||t===undefined||typeof t=='object')continue;const s=String(t).trim();if(!s)continue;let k=ho(ovr,s)?ovr[s]:hk(s),dup=false,tt=s;
  if(k){if(seen[k]){cn[k]=(cn[k]||1)+1;tt=s+'#'+cn[k];k=ho(ovr,tt)?ovr[tt]:'';dup=true}else seen[k]=1}cols.push({n,t:tt,k,dup})}
 return cols};
// Tablo mantığı: skor sütunu çarpım formülüyse girdi sütunlarının Olasılık/Frekans/Şiddet olduğu, IF/arama formülü skora bağlıysa seviye sütunu olduğu çıkarılır
function xlLogic(ws,cols,pr){const by={};cols.forEach(c=>by[c.n]=c);
 cols.forEach(c=>{if(c.dup||(c.k&&c.k!='skor'))return;const f=ws.getCell(pr,c.n).formula;if(typeof f!='string'||!f.includes('*'))return;
  const refs=[...f.matchAll(/\$?([A-Z]{1,3})\$?(\d+)/g)].filter(x=>+x[2]==pr).map(x=>ws.getColumn(x[1]).number);if(refs.length<2||refs.length>3||refs.includes(c.n))return;
  c.k='skor';(refs.length==3?['o','f','s']:['o','s']).forEach((k,i)=>{const d=by[refs[i]];if(d&&!d.k&&!d.dup)d.k=k})});
 const sc=cols.find(c=>c.k=='skor'&&!c.dup);if(sc){const L=ws.getColumn(sc.n).letter;cols.forEach(c=>{if(c.k||c.dup)return;const f=ws.getCell(pr,c.n).formula;if(typeof f=='string'&&/^(IF|IFS|VLOOKUP|LOOKUP|INDEX)/i.test(f)&&new RegExp('\\$?'+L+'\\$?'+pr+'(?!\\d)').test(f))c.k='lv'})}}
const xlProto=(ws,hr,cols)=>{const hc=cols.find(c=>c.k=='h');for(let r=hr+1;r<=Math.min(ws.rowCount,hr+80);r++){const v=hc&&cv(ws.getCell(r,hc.n));if(typeof v=='string'&&v.length>=8&&!isNum(v))return r}return hr+1};
async function xlWrite(wb){try{return await wb.xlsx.writeBuffer()}catch(e){wb.worksheets.forEach(w=>{try{w.dataValidations.model={}}catch(x){}w.conditionalFormattings=[];w.headerFooter={}});return await wb.xlsx.writeBuffer()}}
let xlSt='';
async function xlAnalyze(buf){const wb=new ExcelJS.Workbook();await wb.xlsx.load(buf);const rs=xlSheet(wb,true),sh=rs||xlSheet(wb,false);if(!sh)return{kind:'xlsx',sheet:'',hr:0,hd:[],risk:false,method:''};const cols=xlCols(sh.w,sh.hr);xlLogic(sh.w,cols,sh.hr+1);const ks=cols.map(c=>c.k);
 return{kind:'xlsx',sheet:sh.w.name,hr:sh.hr,hd:cols.map(c=>c.t),cols:cols.map(c=>({t:c.t,k:c.k,dup:!!c.dup,f:typeof sh.w.getCell(sh.hr+1,c.n).formula=='string'})),risk:!!rs,method:rs?(ks.includes('f')?'fk':ks.includes('o')&&ks.includes('s')?'mx':''):''}}
const tr1=(f,a,b)=>f.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/g,(x,c,d,n)=>d||+n!=a?x:c+b);
// D: {firma,date,risk:{rows,method},grid}
async function xlFill(buf,D){try{return await xlFill0(buf,D)}catch(e){throw new Error('Excel doldurma hatası ['+xlSt+']: '+(e.message||e))}}
async function xlFill0(buf,D){xlSt='şablon okuma';const wb=new ExcelJS.Workbook();await wb.xlsx.load(buf);const sh=(D.risk&&xlSheet(wb,true))||xlSheet(wb,false);if(!sh)throw new Error('Excel şablonunda başlık satırı bulunamadı');const ws=sh.w,F=D.firma||{},hr=sh.hr,cols=xlCols(ws,hr);xlLogic(ws,cols,hr+1);xlSt='firma bilgileri';
 const isL=t=>{t=String(t||'').replace(/[:：_\s.]+$/,'').trim();return t.length>0&&t.length<=45&&LBL.some(([k,r])=>r.test(t))};
 const td=new Date(D.date||Date.now()),fd=d=>d.toLocaleDateString('tr'),en=new Date(td);en.setFullYear(td.getFullYear()+[6,4,2][F.th??1]),DT=/\d{1,2}\.\d{1,2}\.\d{4}/;
 for(let r=1;r<hr;r++)ws.getRow(r).eachCell(c=>{const t=cv(c);if(typeof t!='string'||c.isMerged&&c.master&&c.master.address!=c.address)return;
  if(/geçerlilik\s*tarihi/i.test(t))c.value=t.replace(/(\d{1,2}\.\d{1,2}\.\d{4})\s*([-–])\s*(\d{1,2}\.\d{1,2}\.\d{4})/,`${fd(td)} $2 ${fd(en)}`);else if(/(revizyon|yayın|hazırlanma|düzenleme)\s*tarihi/i.test(t)&&DT.test(t))c.value=t.replace(DT,fd(td))});
 for(let r=1;r<hr;r++)ws.getRow(r).eachCell((c,n)=>{const t=cv(c);if(typeof t!='string'||!isL(t)||DT.test(t))return;const x=lbl(t,F,D.date);
  for(let k=n+1;k<=n+8;k++){let y=ws.getCell(r,k);if(y.isMerged&&y.master&&y.master.address==c.address)continue;if(y.isMerged&&y.master)y=y.master;if(typeof cv(y)=='string'&&isL(cv(y)))break;y.value=x;break}});
 const rows=D.risk?D.risk.rows.map((r,i)=>cols.map(c=>riskCell(c.k,r,i,D.risk.method,c.t))):D.grid||[],P=hr+1,pc={};
 cols.forEach(c=>{const x=ws.getCell(P,c.n);pc[c.n]={st:JSON.parse(JSON.stringify(x.style||{})),f:x.formula}});
 xlSt='satır yazma';const last=P+rows.length-1,cfc=new Set();
 (ws.conditionalFormattings||[]).forEach(cf=>{if(typeof cf.ref!='string')return;cf.ref=cf.ref.replace(/([A-Z]+)(\d+):([A-Z]+)(\d+)/g,(m,a,b,c2,d)=>{if(+b>P||+d<P)return m;for(let k=ws.getColumn(a).number;k<=ws.getColumn(c2).number;k++)cfc.add(k);return rows.length?`${a}${b}:${c2}${Math.max(+d,last)}`:m})});
 cols.forEach(c=>{if(cfc.has(c.n)&&pc[c.n].st.fill)delete pc[c.n].st.fill});
 rows.forEach((rw,i)=>{const R=P+i;ws.getRow(R).height=ws.getRow(P).height;cols.forEach((c,j)=>{const x=ws.getCell(R,c.n),p=pc[c.n],v=rw[j]===undefined||rw[j]===null?'':rw[j];x.style=JSON.parse(JSON.stringify(p.st));
  if(typeof p.f=='string'&&p.f)x.value=c.dup&&v===''?null:{formula:tr1(p.f,P,R),result:v};else x.value=v})});
 for(let R=P+rows.length;R<=ws.rowCount;R++)cols.forEach(c=>{const x=ws.getCell(R,c.n);if(!x.formula&&!x.isMerged)x.value=null});
 xlSt='özet sayfası';if(D.risk){const o=wb.addWorksheet('Özet'),s=summ({rows:D.risk.rows,method:D.risk.method});o.addRow(['Alana göre dağılım','Adet']).font={bold:true};Object.entries(s.alan).forEach(a=>o.addRow(a));o.addRow([]);o.addRow(['Risk seviyesine göre dağılım','Adet']).font={bold:true};Object.entries(s.seviye).forEach(a=>o.addRow(a));o.getColumn(1).width=40}
 xlSt='kaydetme';wb.calcProperties.fullCalcOnLoad=true;return new Blob([await xlWrite(wb)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
// Hazır Excel risk şablonu (formüllü): skorlar Excel'de otomatik hesaplanır
async function mkXlRisk(mt){const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('Risk Analizi'),fk=mt=='fk';
 const H=['Sıra No','Alan / Bölüm','Tehlike','Risk',fk?'Olasılık (Ö)':'Olasılık (O)',...(fk?['Frekans (F)']:[]),'Şiddet (Ş)','Risk Skoru','Risk Seviyesi','Alınacak Önlemler'],n=H.length;
 ws.mergeCells(1,1,1,n);ws.getCell(1,1).value='RİSK DEĞERLENDİRMESİ FORMU';ws.getCell(1,1).font={bold:true,size:16};ws.getCell(1,1).alignment={horizontal:'center'};
 ['Firma Unvanı:','Adres:','SGK Sicil No:','Tehlike Sınıfı:','Hazırlayan:','Tarih:'].forEach((t,i)=>{ws.getCell(3+i,1).value=t;ws.getCell(3+i,1).font={bold:true};ws.mergeCells(3+i,2,3+i,4)});
 const hr=10;H.forEach((t,i)=>{const c=ws.getCell(hr,i+1);c.value=t;c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F3864'}};c.alignment={horizontal:'center',vertical:'middle',wrapText:true}});
 const sc=fk?8:7,L=x=>String.fromCharCode(64+x),R=hr+1;
 for(let r=R;r<R+3;r++){H.forEach((t,i)=>{const c=ws.getCell(r,i+1);c.alignment={vertical:'top',wrapText:true};c.border={top:{style:'thin'},left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}}});
  ws.getCell(r,sc).value={formula:fk?`E${r}*F${r}*G${r}`:`E${r}*F${r}`};const v=`${L(sc)}${r}`;
  ws.getCell(r,sc+1).value={formula:fk?`IF(${v}>400,"Çok Yüksek",IF(${v}>200,"Yüksek",IF(${v}>70,"Önemli",IF(${v}>20,"Olası","Düşük"))))`:`IF(${v}>=15,"Çok Yüksek",IF(${v}>=10,"Yüksek",IF(${v}>=5,"Orta","Düşük")))`}}
 [7,18,34,34,11,...(fk?[11]:[]),11,11,14,50].forEach((w,i)=>ws.getColumn(i+1).width=w);ws.views=[{state:'frozen',ySplit:hr}];return wb.xlsx.writeBuffer()}

// ---------- Kütüphane için okuma (yapay zekâsız) ----------
const num=v=>parseFloat(String(v).replace(',','.'))||0,isNum=v=>/^[\d.,\s]+$/.test(v);
// ovr: {başlık metni: anlam}. Anlamı olmayan sütunlar (sorumlu, termin vb.) madde ile birlikte "x" içinde saklanır.
async function xlRead(buf,ovr={}){const wb=new ExcelJS.Workbook();await wb.xlsx.load(buf);const out=[];let fk=false;
 wb.worksheets.forEach(ws=>{const hr=xlHdr(ws,true);if(!hr)return;const cols=xlCols(ws,hr,ovr);xlLogic(ws,cols,xlProto(ws,hr,cols));if(!cols.some(c=>c.k=='h'))return;if(cols.some(c=>c.k=='f'))fk=true;let sec='';
  for(let r=hr+1;r<=ws.rowCount;r++){const o={a:'',x:{}},tx=[];cols.forEach(c=>{let v=cv(ws.getCell(r,c.n));if(v===null||v===undefined||v==='')return;if(typeof v=='object')v=v.text||v.result||'';const nv=typeof v=='number'?v:null;v=String(v).trim();if(!v)return;tx.push(v);
    if(['o','f','s'].includes(c.k))o[c.k]=num(v);else if(['a','h','r','t'].includes(c.k))o[c.k]=v.slice(0,600);else if(!c.k)o.x[c.t]=nv!==null?nv:v.slice(0,300)});
   // bölüm başlığı / özet / boş skorlu satırlar madde sayılmaz; bölüm başlığı sonraki satırların alanı olur
   if(!o.h||isNum(o.h)||o.h.length<6||(!o.r&&!o.t&&!o.o&&!o.s)){const t=!o.h||isNum(o.h)?tx.find(x=>!isNum(x)&&x.length<80):o.h;if(t&&t.length<45&&!o.r&&!o.t&&o.o===undefined)sec=t;continue}
   if(!o.a)o.a=sec||ws.name;out.push(o)}});
 return{rows:out,method:fk?'fk':'mx'}}
async function dxRead(buf,ovr={}){const{TBL}=await dxLoad(buf),out=[];let fk=false;
 TBL.forEach(t=>{const trs=wk(t,'tr'),tx=trs.map(r=>wk(r,'tc').map(c=>wq(c,'p').map(ptx).join(' ').trim())),hd=tx[0]||[],ks=hd.map(h=>Object.prototype.hasOwnProperty.call(ovr,h)?ovr[h]:hk(h));if(ks.filter(Boolean).length<3||!ks.includes('h'))return;if(ks.includes('f'))fk=true;
  tx.slice(1).forEach(rw=>{const o={a:'',x:{}};ks.forEach((k,i)=>{const v=rw[i];if(!v)return;if(['o','f','s'].includes(k))o[k]=num(v);else if(['a','h','r','t'].includes(k))o[k]=v.slice(0,600);else if(!k&&hd[i])o.x[hd[i]]=/^-?\d+([.,]\d+)?$/.test(v)?num(v):v.slice(0,300)});if(o.h&&!isNum(o.h)&&o.h.length>=6)out.push(o)})});
 return{rows:out,method:fk?'fk':'mx'}}
async function xlProbe(buf){const wb=new ExcelJS.Workbook();await wb.xlsx.load(buf);const sh=xlSheet(wb,true);if(!sh)throw new Error('Risk tablosu başlığı bulunamadı (Tehlike, Risk, Önlem gibi başlıklar aranıyor).');const{w,hr}=sh,cols=xlCols(w,hr);xlLogic(w,cols,xlProto(w,hr,cols));
 cols.forEach(c=>{c.s=[];for(let r=hr+1;r<=Math.min(w.rowCount,hr+80)&&c.s.length<3;r++){let v=cv(w.getCell(r,c.n));if(v&&typeof v=='object')v=v.text||v.result;if(v!==null&&v!==undefined&&String(v).trim())c.s.push(String(v).trim().slice(0,40))}});
 return{sheet:w.name,sheets:wb.worksheets.filter(x=>xlHdr(x,true)).length,rows:w.rowCount-hr,cols}}
async function dxProbe(buf){const{TBL}=await dxLoad(buf);for(const t of TBL){const trs=wk(t,'tr'),tx=trs.map(r=>wk(r,'tc').map(c=>wq(c,'p').map(ptx).join(' ').trim())),hd=tx[0]||[];if(hd.filter(h=>hk(h)).length>=3)return{sheet:'Word tablosu',sheets:1,rows:tx.length-1,cols:hd.map((h,i)=>({n:i,t:h,k:hk(h),s:tx.slice(1,40).map(r=>r[i]).filter(Boolean).slice(0,3).map(x=>x.slice(0,40))}))}}throw new Error('Word dosyasında risk tablosu bulunamadı.')}
// Kütüphane dosyasından boş iskelet: başlıklar, biçimler, genişlikler, formüller kalır; veri satırları silinir
async function xlSkeleton(buf){const wb=new ExcelJS.Workbook();await wb.xlsx.load(buf);const cand=wb.worksheets.map(w=>({w,hr:xlHdr(w,true)})).filter(x=>x.hr).sort((a,b)=>b.w.rowCount-a.w.rowCount);if(!cand.length)throw new Error('iskelet çıkarılamadı');
 const{w:ws,hr}=cand[0];cand.slice(1).forEach(x=>wb.removeWorksheet(x.w.id));const cols=xlCols(ws,hr),pr=xlProto(ws,hr,cols);
 const st={},fm={};cols.forEach(c=>{const x=ws.getCell(pr,c.n);st[c.n]=JSON.parse(JSON.stringify(x.style||{}));fm[c.n]=x.formula});const ht=ws.getRow(pr).height;
 for(let r=hr+1;r<=ws.rowCount;r++){const rw=ws.getRow(r);rw.eachCell({includeEmpty:true},c=>{c.value=null;c.style={}});rw.height=undefined}cols.forEach(c=>{const x=ws.getCell(hr+1,c.n);x.style=st[c.n];if(fm[c.n])x.value={formula:tr1(fm[c.n],pr,hr+1)}});ws.getRow(hr+1).height=ht;return xlWrite(wb)}
