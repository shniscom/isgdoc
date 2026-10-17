// Şablon iskeleti: Word (.docx) ve Excel (.xlsx) dosyalarının biçimini koruyarak doldurma
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',wq=(n,t)=>[...n.getElementsByTagNameNS(W,t)],wk=(n,t)=>[...n.childNodes].filter(c=>c.localName==t);
const ptx=p=>wq(p,'t').map(t=>t.textContent).join('');
const LBL=[['ad',/(firma|iş\s?yeri|işyeri|kurum|şirket).{0,14}(ad|unvan|ünvan)|^firma|unvan|ünvan/i],['sgk',/sgk|sicil/i],['adres',/adres/i],['th',/tehlike sınıf/i],['isveren',/işveren/i],['uzman',/uzman|hazırlayan/i],['hekim',/hekim/i],['tarih',/tarih/i]];
function lbl(t,f,date){t=String(t||'').replace(/[:：_\s.]+$/,'').trim();if(!t||t.length>45)return'';for(const[k,r]of LBL)if(r.test(t))return k=='th'?(TH[f.th]||''):k=='tarih'?new Date(date).toLocaleDateString('tr'):(f[k]||'');return''}
const hk=t=>{t=String(t||'').toLocaleLowerCase('tr');return/sıra|^no$|^s\.?\s?no/.test(t)?'no':/önlem|kontrol|düzeltici/.test(t)?'t':/skor|puan|derece|risk değeri/.test(t)?'skor':/seviye|sınıf|düzey|kategori/.test(t)?'lv':/olasılık/.test(t)?'o':/frekans|sıklık/.test(t)?'f':/şiddet/.test(t)?'s':/tehlike/.test(t)?'h':/risk|sonuç|zarar/.test(t)?'r':/alan|bölüm|birim|faaliyet|yer/.test(t)?'a':''};
const riskCell=(k,r,i,mt)=>({no:i+1,a:r.a,h:r.h,r:r.r,o:r.o,f:r.f,s:r.s,skor:sk(r,mt),lv:lv(sk(r,mt),mt),t:r.t})[k]??'';
// ---------- WORD ----------
async function dxLoad(buf){const z=await JSZip.loadAsync(buf),doc=new DOMParser().parseFromString(await z.file('word/document.xml').async('string'),'application/xml'),body=doc.getElementsByTagNameNS(W,'body')[0],TBL=wk(body,'tbl');
 const T=TBL.map((t,i)=>{const rows=wk(t,'tr'),tx=rows.map(r=>wk(r,'tc').map(c=>wq(c,'p').map(ptx).join(' ').trim())),hd=tx[0]||[],bd=tx.slice(1);
  const list=rows.length>=2&&hd.length>=3&&hd.every(x=>x)&&bd.every(r=>!r[0]||/^\d+[.)]?$/.test(r[0]));let b=t.previousSibling;while(b&&!(b.localName=='p'&&ptx(b).trim()))b=b.previousSibling;
  return{i,list,hd,n:bd.length,before:b?ptx(b).trim().slice(0,80):''}});
 const sl=[];wq(body,'p').forEach((p,k)=>{let ti=-1,r=-1,c=-1;const tc=p.parentNode.localName=='tc'?p.parentNode:null;
  if(tc){const tr=tc.parentNode,tl=tr.parentNode;ti=TBL.indexOf(tl);r=wk(tl,'tr').indexOf(tr);c=wk(tr,'tc').indexOf(tc);if(wq(tc,'p')[0]!==p)return}
  if(ti>=0&&T[ti].list)return;const t=ptx(p).trim();if(ti<0&&!t)return;sl.push({id:'s'+k,k,t,ti,r,c})});
 return{z,doc,body,TBL,pub:{kind:'docx',T,sl}}}
const dxAnalyze=async b=>(await dxLoad(b)).pub;
const dxText=p=>[...p.T.filter(t=>t.list).map(t=>`[TABLO ${t.i}] başlıklar: ${t.hd.join(' | ')} (mevcut satır: ${t.n}; öncesindeki başlık: ${t.before})`),...p.sl.map(s=>`[${s.id}]${s.ti>=0?` (tablo ${s.ti} satır ${s.r} sütun ${s.c})`:''} ${s.t||'(boş)'}`)].join('\n').slice(0,14000);
function setP(p,text){const d=p.ownerDocument,rs=wq(p,'r');let r=rs[0];if(!r){r=d.createElementNS(W,'w:r');const rp=wq(p,'rPr')[0];if(rp&&rp.parentNode.localName=='pPr')r.appendChild(rp.cloneNode(true));p.appendChild(r)}
 [...r.childNodes].forEach(n=>{if(['t','tab','br'].includes(n.localName))r.removeChild(n)});const t=d.createElementNS(W,'w:t');t.setAttribute('xml:space','preserve');t.textContent=text;r.appendChild(t);rs.slice(1).forEach(x=>x.parentNode&&x.parentNode.removeChild(x))}
// D: {firma,date,fill:{slots,rows},risk:{rows,method},teams:(hd,before)=>rows|null}
async function dxFill(buf,D){const{z,doc,body,TBL,pub}=await dxLoad(buf),P=wq(body,'p'),F=D.firma||{},fl=D.fill||{slots:{},rows:{}},by={};pub.sl.forEach(s=>by[s.ti+'.'+s.r+'.'+s.c]=s);
 pub.sl.forEach(s=>{const p=P[s.k];let v=fl.slots[s.id];
  if(v===undefined){ // etiket tabanlı otomatik doldurma
   if(s.ti>=0&&!s.t&&s.c>0){const l=by[s.ti+'.'+s.r+'.'+(s.c-1)];const x=l&&lbl(l.t,F,D.date);if(x)v=x}
   else{const q=s.t.match(/^(.{2,40}?)\s*[:：]\s*[_\s.…]*$/);if(q){const x=lbl(q[1],F,D.date);if(x)v=q[1]+': '+x}}}
  if(v!==undefined)setP(p,String(v))});
 pub.T.forEach(t=>{if(!t.list)return;let rows=fl.rows[t.i];
  if(D.teams){const x=D.teams(t.hd,t.before);if(x)rows=x}
  if(D.risk&&t.hd.filter(h=>hk(h)).length>=3){const ks=t.hd.map(hk);rows=D.risk.rows.map((r,i)=>ks.map(k=>riskCell(k,r,i,D.risk.method)))}
  if(!rows)return;const tb=TBL[t.i],trs=wk(tb,'tr'),pr=trs[trs.length-1].cloneNode(true);trs.slice(1).forEach(x=>tb.removeChild(x));
  rows.forEach(rw=>{const tr=pr.cloneNode(true);wk(tr,'tc').forEach((tc,ci)=>{const ps=wq(tc,'p');setP(ps[0],String(rw[ci]??''));ps.slice(1).forEach(x=>tc.removeChild(x))});tb.appendChild(tr)})});
 let x=new XMLSerializer().serializeToString(doc);if(!x.startsWith('<?xml'))x='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+x;z.file('word/document.xml',x);return z.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'})}
// ---------- EXCEL ----------
const slave=c=>c.isMerged&&c.master&&c.master.address!=c.address;
const cv=c=>{const v=c.value;return v&&v.richText?v.richText.map(x=>x.text).join(''):v&&typeof v=='object'&&v.result!==undefined?v.result:v&&v.text?v.text:v};
function xlHdr(ws,risk){let best=0,row=0;for(let r=1;r<=Math.min(ws.rowCount,40);r++){let n=0;ws.getRow(r).eachCell(c=>{if(slave(c))return;const v=cv(c);if(typeof v=='string'&&v.trim()&&v.length<=40&&(!risk||hk(v)))n++});if(n>best){best=n;row=r}}return best>=(risk?2:3)?row:0}
async function xlAnalyze(buf){const wb=new ExcelJS.Workbook();await wb.xlsx.load(buf);const ws=wb.worksheets[0],rk=xlHdr(ws,true),hr=rk||xlHdr(ws,false),hd=[];if(hr)ws.getRow(hr).eachCell(c=>{if(slave(c))return;const v=cv(c);if(v)hd.push(String(v))});return{kind:'xlsx',sheet:ws.name,hr,hd,risk:!!rk}}
const tr1=(f,a,b)=>f.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/g,(x,c,d,n)=>d||+n!=a?x:c+b);
// D: {firma,date,risk:{rows,method},grid}
async function xlFill(buf,D){const wb=new ExcelJS.Workbook();await wb.xlsx.load(buf);const ws=wb.worksheets[0],F=D.firma||{},hr=xlHdr(ws,!!D.risk)||xlHdr(ws,false);if(!hr)throw new Error('Excel şablonunda başlık satırı bulunamadı');
 const cols=[];ws.getRow(hr).eachCell((c,n)=>{if(slave(c))return;const t=cv(c);if(t)cols.push({n,k:hk(t)})});
 for(let r=1;r<hr;r++)ws.getRow(r).eachCell((c,n)=>{const t=cv(c);if(typeof t!='string')return;const x=lbl(t,F,D.date);if(!x)return;
  for(let k=n+1;k<=n+8;k++){let y=ws.getCell(r,k);if(y.isMerged&&y.master&&y.master.address==c.address)continue;if(y.isMerged&&y.master)y=y.master;if(cv(y)===null||cv(y)===undefined||cv(y)===''){y.value=x}break}});
 const rows=D.risk?D.risk.rows.map((r,i)=>cols.map(c=>riskCell(c.k,r,i,D.risk.method))):D.grid||[],P=hr+1,pc={};
 cols.forEach(c=>{const x=ws.getCell(P,c.n);pc[c.n]={st:JSON.parse(JSON.stringify(x.style||{})),f:x.formula}});
 rows.forEach((rw,i)=>{const R=P+i;ws.getRow(R).height=ws.getRow(P).height;cols.forEach((c,j)=>{const x=ws.getCell(R,c.n),p=pc[c.n];x.style=JSON.parse(JSON.stringify(p.st));x.value=p.f?{formula:tr1(p.f,P,R),result:rw[j]}:rw[j]})});
 if(D.risk){const o=wb.addWorksheet('Özet'),s=summ({rows:D.risk.rows,method:D.risk.method});o.addRow(['Alana göre dağılım','Adet']).font={bold:true};Object.entries(s.alan).forEach(a=>o.addRow(a));o.addRow([]);o.addRow(['Risk seviyesine göre dağılım','Adet']).font={bold:true};Object.entries(s.seviye).forEach(a=>o.addRow(a));o.getColumn(1).width=40}
 wb.calcProperties.fullCalcOnLoad=true;return new Blob([await wb.xlsx.writeBuffer()],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
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
