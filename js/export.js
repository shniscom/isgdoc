// Word/Excel çıktıları: şablon varsa iskelet korunur, yoksa varsayılan düzen kullanılır
const X=()=>docx,P=(t,o={})=>new (X().Paragraph)({heading:o.h,alignment:o.a,spacing:{before:o.sp||0,after:120},children:[new (X().TextRun)({text:t,bold:o.b,size:o.s})]});
const H1=t=>P(t,{h:X().HeadingLevel.HEADING_1,b:1}),H2=t=>P(t,{h:X().HeadingLevel.HEADING_2,b:1});
function bl(t){return t.split('\n').filter(l=>l.trim()).map(l=>l.startsWith('## ')?H2(l.slice(3)):l.startsWith('# ')?H1(l.slice(2)):/^[-*] /.test(l)?new (X().Paragraph)({text:l.slice(2),bullet:{level:0}}):P(l.replace(/\*\*/g,'')))}
function TB(rows,w){const x=X();return new x.Table({width:{size:100,type:x.WidthType.PERCENTAGE},rows:rows.map((r,i)=>new x.TableRow({tableHeader:!i,children:r.map((c,j)=>new x.TableCell({shading:!i?{fill:'D9E2F3'}:undefined,children:[new x.Paragraph({children:[new x.TextRun({text:String(c??''),bold:!i,size:18})]})]}))}))})}
function cover(d,t){const f=d.firma,x=X(),C=x.AlignmentType.CENTER;return [P('',{sp:2400}),P(f.ad||'',{a:C,b:1,s:40}),P('',{sp:600}),P(t.toUpperCase(),{a:C,b:1,s:36}),P('',{sp:1200}),...[['SGK Sicil No',f.sgk],['Adres',f.adres],['Tehlike Sınıfı',TH[f.th]],['İş Kolu',d.sektor],['Hazırlayan',f.uzman],['Hazırlanma Tarihi',new Date(d.date).toLocaleDateString('tr')]].map(([k,v])=>P(`${k}: ${v||''}`,{a:C,s:24})),new x.Paragraph({children:[new x.PageBreak()]})]}
const dlBlob=async(b,name)=>{const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=name.replace(/[\\/:*?"<>|]/g,'');a.click();await new Promise(r=>setTimeout(r,800))};
const pack=async(name,secs)=>dlBlob(await X().Packer.toBlob(new (X().Document)({sections:secs.map(c=>({children:c}))})),name+'.docx');
// Acil durum ekipleri: gerekli kişi sayısı çalışan sayısı ve tehlike sınıfına göre
const need=(d,i)=>Math.max(1,Math.ceil((+d.firma.calisan||0)/(i==3?ILK:DIG)[d.firma.th??1]));
function teamRows(d,hd,before,tpl){const bf=String(before||'').toLocaleLowerCase('tr'),i=EKR.findIndex(r=>r.test(bf));if(i<0)return null;const ppl=(d.ekip[EK[i]]||'').split('\n').filter(s=>s.trim()),tp=tpl||[],lab=tp.filter(r=>r.some(x=>x)).length,len=Math.max(ppl.length,need(d,i)+3,lab);
 const txt=hd.every(h=>typeof h=='string');
 return Array.from({length:len},(_,j)=>{const p=(ppl[j]||'').split(';').map(s=>s.trim());
  if(!txt){const r=[j+1,p[0]||'',p[1]||'',p[2]||'',''];while(r.length<hd.length)r.push('');return r.slice(0,hd.length)}
  const b=tp[j]||[];return hd.map((h,c)=>{const t=String(h).toLocaleLowerCase('tr'),bc=b[c]||'';
   if(/^(no|sıra|s\.?\s?no)/.test(t))return bc||(ppl[j]||bc?j+1:'');
   if(/ad[ıi]\s*soyad|isim/.test(t))return p[0]||'';
   if(/görev|unvan|ünvan|rol/.test(t))return bc||p[1]||'';
   if(/telefon|dahili|iletişim|gsm/.test(t))return p[2]||'';
   if(/imza|paraf/.test(t))return '';return bc})})}
function ekipTb(d){const out=[];EK.forEach((k,i)=>{const n=need(d,i),rows=teamRows(d,[1,2,3,4,5],k);out.push(P(`${k} (en az ${n} kişi; liste ${rows.length} satır)`,{b:1,sp:200}),TB([['No','Adı Soyadı','Görevi','İletişim','İmza'],...rows]))});return out}
const tplOf=d=>d.tpl&&TP.find(t=>t.id==d.tpl);
async function indir(d){try{await indir0(d)}catch(x){alert('İndirme hatası: '+(x.message||x))}}
async function indir0(d){const tt=TY[d.type][0],f=d.firma,nm=`${f.ad} - `,t=tplOf(d),ext=t&&t.kind=='xlsx'?'.xlsx':'.docx',base={firma:f,date:d.date,year:(String(d.ek||'').match(/yılı[^\d]{0,5}(20\d\d)/i)||[])[1]};
 if(d.type=='risk'){const mt=d.method,rk={rows:d.rows,method:mt};const og=d.out||{a:1,k:1,p:1};if(!og.a&&!og.k&&!og.p)throw new Error('İndirilecek dosya seçilmedi');if(og.k)await pack(nm+'Kapak',[cover(d,'Risk Değerlendirmesi')]);
  if(!og.a){}else if(t)await dlBlob(t.kind=='xlsx'?await xlFill(t.buf,{...base,risk:rk}):await dxFill(t.buf,{...base,risk:rk}),nm+'Risk Analizi'+ext);
  else{const hd=mt=='fk'?['Sıra','Alan','Tehlike','Risk','Ö','F','Ş','Skor','Seviye','Önlemler']:['Sıra','Alan','Tehlike','Risk','O','Ş','Skor','Seviye','Önlemler'],s=summ(d),sm=q=>Object.entries(q),x=X();
   const rows=d.rows.map((r,i)=>(mt=='fk'?[i+1,r.a,r.h,r.r,r.o,r.f,r.s]:[i+1,r.a,r.h,r.r,r.o,r.s]).concat([sk(r,mt),lv(sk(r,mt),mt),r.t]));
   await pack(nm+'Risk Analizi',[[H1(`${f.ad} – RİSK ANALİZİ`),P(`Yöntem: ${mt=='fk'?'Fine-Kinney':'5×5 Matris'} | Toplam risk: ${d.rows.length}`),TB([hd,...rows]),P('Alanlara Göre Dağılım',{b:1,sp:300}),TB([['Alan','Adet'],...sm(s.alan)]),P('Risk Seviyesine Göre Dağılım',{b:1,sp:300}),TB([['Seviye','Adet'],...sm(s.seviye)]),P('Risk Değerlendirme Ekibi',{b:1,sp:300}),TB([['Ünvan','Adı Soyadı','İmza'],['İşveren / Vekili',f.isveren,''],['İSG Uzmanı',f.uzman,''],['İşyeri Hekimi',f.hekim,''],['Çalışan Temsilcisi','',''],['Destek Elemanı','',''],['Bilgi Sahibi Çalışan','','']])]])}
  if(og.p){const pt=procTpl();if(pt&&d.pfill)await dlBlob(await dxFill(pt.buf,{...base,fill:d.pfill}),nm+'Risk Değerlendirme Prosedürü.docx');else await pack(nm+'Risk Değerlendirme Prosedürü',[d.proc?bl(d.proc):[P('Prosedür henüz hazırlanmadı.')]])}return}
 if(t)await planOf(t);if(t){const D={...base,fill:d.fill,teams:d.type=='acil'?(hd,b,tp)=>teamRows(d,hd,b,tp):null,grid:d.fill&&d.fill.grid};
  if(d.type=='acil')await pack(nm+'Kapak',[cover(d,tt)]);
  return dlBlob(t.kind=='xlsx'?await (t.pub&&t.pub.plan?xpFill:D.fill&&D.fill.xs?xsFill:xlFill)(t.buf,D):await dxFill(t.buf,D),nm+tt+ext)}
 const body=d.type=='acil'?[...bl(d.text||''),H1('ACİL DURUM EKİPLERİ'),...ekipTb(d),H1('EKLER'),...['EK-1 Acil durum krokisi','EK-2 Kaçış yolları ve toplanma yeri krokisi','EK-3 İlk yardım dolabı yerleri ve sayısı','EK-4 Yangın söndürücü yerleri ve sayısı','EK-5 Acil durum ekipman listesi','EK-6 Tatbikat formu'].map(s=>new (X().Paragraph)({text:s,bullet:{level:0}}))]:bl(d.text||'');
 await pack(nm+tt,d.type=='acil'?[cover(d,tt),body]:[body])}
// Yapay zekânın hazırladığı biçimli Word şablonu (üst/alt bilgi, tablolar, sütun genişlikleri)
async function specDocx(sp){const x=X(),W=9000,B=x.BorderStyle.SINGLE,bd={style:B,size:4,color:'808080'},bds={top:bd,bottom:bd,left:bd,right:bd};
 const mk=b=>{if(b.t=='h1')return[H1(b.x)];if(b.t=='h2')return[H2(b.x)];if(b.t=='tablo'||b.t=='form'){const hd=b.t=='form'?null:b.basliklar||[],w=b.w||(hd||b.satirlar[0]).map(()=>1),sm=w.reduce((a,c)=>a+c,0),cw=w.map(v=>Math.round(v/sm*W));
   const cell=(t,i,h)=>new x.TableCell({width:{size:cw[i],type:x.WidthType.DXA},borders:bds,shading:h||(b.t=='form'&&i==0)?{fill:'D9E2F3'}:undefined,children:[new x.Paragraph({children:[new x.TextRun({text:String(t??''),bold:h||(b.t=='form'&&i==0),size:20})]})]});
   const rows=b.t=='form'?b.satirlar.map(r=>new x.TableRow({children:r.map((c,i)=>cell(c,i))})):[new x.TableRow({tableHeader:true,children:hd.map((c,i)=>cell(c,i,1))}),...Array.from({length:b.n||3},(_,j)=>new x.TableRow({children:hd.map((c,i)=>cell(i==0&&/sıra|no/i.test(c)?j+1:'',i))}))];
   return[new x.Table({width:{size:W,type:x.WidthType.DXA},columnWidths:cw,rows}),P('')]}return[P(b.x||'')]};
 const wide=(sp.bloklar||[]).some(b=>(b.basliklar||[]).length>6);
 const doc=new x.Document({sections:[{properties:{page:{size:{orientation:wide?x.PageOrientation.LANDSCAPE:x.PageOrientation.PORTRAIT}}},headers:{default:new x.Header({children:[P(sp.ustBilgi||'',{a:x.AlignmentType.CENTER,s:18})]})},footers:{default:new x.Footer({children:[new x.Paragraph({alignment:x.AlignmentType.CENTER,children:[new x.TextRun({text:(sp.altBilgi||'')+'  Sayfa ',size:16}),new x.TextRun({children:[x.PageNumber.CURRENT],size:16})]})]})},children:[P(sp.baslik||'',{a:x.AlignmentType.CENTER,b:1,s:30}),...(sp.bloklar||[]).flatMap(mk)]}]});
 return x.Packer.toBlob(doc)}
