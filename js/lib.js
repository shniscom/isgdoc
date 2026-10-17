// Bilgi kütüphanesi: uzman risk analizlerinden madde havuzu (kodla ayrıştırma, iş kolu etiketi, benzer madde birleştirme)
let LIB=[],UC=0,libRep=[];
const ETK=[...SEC,'Yemekhane / Mutfak','Atölye / Bakım','Elektrik / Enerji','Araç / Taşımacılık','Laboratuvar','Genel'];
const STOP=new Set(['ile','için','olan','olarak','veya','gibi','daha','çok','nedeniyle','sırasında','kaynaklı','riski','risk','yetersiz','uygun','olması','olmaması']);
const toks=s=>[...new Set(String(s||'').toLocaleLowerCase('tr').replace(/[^a-zçğıöşü0-9 ]/g,' ').split(/\s+/).filter(w=>w.length>2&&!STOP.has(w)).map(w=>w.slice(0,6)))];
const jac=(a,b)=>{const B=new Set(b);let i=0;a.forEach(x=>B.has(x)&&i++);return i/(a.length+b.length-i||1)};
// rows: [{a,h,r,o,f,s,t,m}], labs(row)->[etiketler]. Benzerlik >= %75 ise yeni madde eklenmez, mevcut maddeye birleştirilir.
function libMerge(rows,labs,src,lib=LIB){const idx=new Map(),ch=new Set(),ad=(i)=>i.k.forEach(t=>{if(!idx.has(t))idx.set(t,[]);idx.get(t).push(i)});lib.forEach(ad);let nw=0,dup=0;
 rows.forEach(r=>{const k=toks(r.h+' '+r.r),L=labs(r),c=new Map();let best=null,bs=0;k.forEach(t=>(idx.get(t)||[]).forEach(i=>c.set(i,(c.get(i)||0)+1)));
  c.forEach((n,i)=>{if(n>=Math.min(2,k.length)){const j=jac(k,i.k);if(j>bs){bs=j;best=i}}});
  if(best&&bs>=0.75){dup++;if(!best.src.includes(src)){best.src.push(src);best.n=best.src.length}L.forEach(l=>{if(!best.l.includes(l))best.l.push(l)});if(!best.t&&r.t)best.t=r.t;ch.add(best);return}
  const it={id:uid()+(UC++),a:r.a||'',h:r.h||'',r:r.r||'',o:r.o||0,f:r.f||0,s:r.s||0,t:r.t||'',m:r.m,l:[...L],src:[src],n:1,k};nw++;lib.push(it);ch.add(it);ad(it)});
 return{nw,dup,ch:[...ch]}}
const libLabels=()=>{const c={};LIB.forEach(i=>i.l.forEach(l=>c[l]=(c[l]||0)+1));return Object.entries(c).sort((a,b)=>b[1]-a[1])};
const libBest=s=>{s=String(s||'').toLocaleLowerCase('tr').split(/[\s/]+/)[0];return s?(libLabels().find(([l])=>l.toLocaleLowerCase('tr').includes(s))||[''])[0]:''};
const O5=[0.5,1,3,6,10],S5=[1,3,7,15,40],near=(v,a)=>{let b=0;a.forEach((x,k)=>{if(Math.abs(x-v)<Math.abs(a[b]-v))b=k});return b};
function conv(i,to){const r={a:i.a,h:i.h,r:i.r,t:i.t,lib:1,o:i.o,f:i.f,s:i.s};if(i.m==to)return r;
 if(to=='fk'){r.o=O5[Math.min(4,Math.max(0,Math.round(i.o)-1))]||1;r.s=S5[Math.min(4,Math.max(0,Math.round(i.s)-1))]||3;r.f=3}else{r.o=near(i.o,O5)+1;r.s=near(i.s,S5)+1;r.f=1}return r}
// Seçim: etikete uyanlar, alanlara dengeli dağılım; birden çok kaynakta geçen (n yüksek) maddeler önce
function libPick(label,n,alanTxt,to){let c=LIB.filter(i=>i.l.includes(label));const q=String(alanTxt||'').toLocaleLowerCase('tr').split(',').map(x=>x.trim()).filter(Boolean);
 if(q.length){const f=c.filter(i=>q.some(x=>i.a.toLocaleLowerCase('tr').includes(x)));if(f.length)c=f}
 const g={};c.sort((a,b)=>b.n-a.n).forEach(i=>(g[i.a||'-']=g[i.a||'-']||[]).push(i));const gs=Object.values(g),out=[];
 for(let k=0;out.length<n&&gs.some(x=>k<x.length);k++)gs.forEach(x=>{if(k<x.length&&out.length<n)out.push(x[k])});return out.map(i=>conv(i,to))}
async function libUp(){const fs=[...$('#lf').files];if(!fs.length)return alert('Dosya seçin');const main=$('#lm').value.trim();libRep=[];
 for(let q=0;q<fs.length;q++){const f=fs[q],P=(x,t)=>pg(Math.round((q+x/100)/fs.length*100),`${f.name}: ${t}`),isx=/\.xlsx$/i.test(f.name);if(!isx&&!/\.docx$/i.test(f.name)){libRep.push({name:f.name,err:'Yalnızca .xlsx / .docx'});continue}
  P(5,'okunuyor');const buf=await f.arrayBuffer(),rd=isx?await xlRead(buf):await dxRead(buf);if(!rd.rows.length){libRep.push({name:f.name,err:'Risk tablosu bulunamadı'});continue}
  P(30,'iş kolları belirleniyor');let ana=main,mp={};const al=[...new Set(rd.rows.map(r=>r.a).filter(Boolean))].slice(0,80),st=Math.ceil(rd.rows.length/12);
  if(S.k){try{const o=jp(await ai(`Bir risk analizi dosyasındaki alan/bölüm adları: ${JSON.stringify(al)}. Örnek tehlikeler: ${JSON.stringify(rd.rows.filter((_,i)=>i%st==0).slice(0,12).map(r=>r.h))}. Dosya adı: ${f.name}. Etiket listesi: ${ETK.join(', ')}.\nHer alan adı için listeden en uygun iş kolu/faaliyet etiketini seç (uymuyorsa kısa yeni etiket yaz; genel alanlar için "Genel"). Belgenin ana iş kolunu da belirle. SADECE JSON: {"ana":"","alanlar":{"alan adı":"etiket"}}`,2500));if(!ana)ana=o.ana;mp=o.alanlar||{}}catch(x){}}
  ana=ana||'Genel';P(60,'benzer maddeler ayıklanıyor');const sid=uid(),labs=r=>{const l=[ana],x=mp[r.a];if(x&&x!='Genel'&&!l.includes(x))l.push(x);return l};rd.rows.forEach(r=>r.m=rd.method);
  const g=libMerge(rd.rows,labs,sid);P(90,'kaydediliyor');await kPut(g.ch);const rp={id:sid,name:f.name,date:new Date().toISOString(),okunan:rd.rows.length,nw:g.nw,dup:g.dup,ana,method:rd.method,alanlar:al.length};S.ls.push(rp);libRep.push(rp);sv()}
 pg(100,'Tamamlandı');libView()}
async function libDel(id){if(!confirm('Bu kaynağa ait maddeler kütüphaneden çıkarılsın mı? (Başka kaynakta da geçen maddeler kalır)'))return;const up=[],dl=[];LIB.forEach(i=>{if(!i.src.includes(id))return;i.src=i.src.filter(x=>x!=id);i.n=i.src.length;i.src.length?up.push(i):dl.push(i.id)});LIB=LIB.filter(i=>i.src.length);await kDel(dl);await kPut(up);S.ls=S.ls.filter(x=>x.id!=id);sv();libView()}
async function libClear(){if(!confirm('Tüm kütüphane silinsin mi?'))return;await kDel(LIB.map(i=>i.id));LIB=[];S.ls=[];sv();libView()}
function libView(){const L=libLabels();m.innerHTML=`<h2>Kütüphane</h2><div class="card"><p style="font-size:13px;color:var(--m)">Uzmanlarca hazırlanmış risk analizlerini (Excel/Word) yükleyin. Tablolar yapay zekâsız okunur; yalnızca alan/bölüm adları iş kolu etiketlemesi için tek bir kısa yapay zekâ çağrısıyla sınıflandırılır. Firma adı, adres, kişi bilgileri alınmaz; yalnızca tehlike/risk/skor/önlem maddeleri saklanır. Benzer maddeler birleştirilir.</p>
 <div class="g"><label>Dosyalar (.xlsx / .docx, birden fazla seçilebilir)<input type="file" id="lf" multiple accept=".xlsx,.docx"></label><label>Ana iş kolu (biliyorsanız; boşsa yapay zekâ belirler)<input id="lm" list="sl2" placeholder="Örn: İnşaat"><datalist id="sl2">${SEC.map(x=>`<option>${x}`).join('')}</datalist></label></div>
 <button class="b" onclick="run(this,libUp)">Yükle ve analiz et</button><div id="pg"></div></div>
 ${libRep.length?`<div class="card"><h3>Son yükleme</h3>${libRep.map(r=>r.err?`<p>❌ ${e(r.name)}: ${e(r.err)}</p>`:`<p>✅ <b>${e(r.name)}</b> — ${r.okunan} satır okundu, <b>${r.nw} yeni madde</b>, ${r.dup} benzer/tekrar madde birleştirildi. Ana iş kolu: ${e(r.ana)}; ${r.alanlar} farklı alan/bölüm.</p>`).join('')}</div>`:''}
 <div class="card"><h3>Kütüphane durumu: ${LIB.length} madde</h3>${L.length?`<div class="w"><table><tr><th>İş kolu / faaliyet<th>Madde sayısı<th>Birden çok kaynakta geçen</tr>${L.map(([l,n])=>`<tr><td>${e(l)}<td>${n}<td>${LIB.filter(i=>i.l.includes(l)&&i.n>1).length}</tr>`).join('')}</table></div><p style="font-size:12px;color:var(--m)">Bir madde birden çok iş koluna etiketli olabilir (ör. inşaat belgesindeki yemekhane riskleri hem İnşaat hem Yemekhane / Mutfak sayılır).</p>`:'<p>Kütüphane boş.</p>'}</div>
 ${S.ls.length?`<div class="card"><h3>Kaynak belgeler</h3>${S.ls.map(r=>`<p>${e(r.name)} <span class="tag">${e(r.ana)}</span><span class="tag">${r.nw} yeni</span><span class="tag">${r.method=='fk'?'Fine-Kinney':'5×5'}</span> <button class="b r" onclick="run(this,()=>libDel('${r.id}'))">Çıkar</button></p>`).join('')}<button class="b r" onclick="run(this,libClear)">Kütüphaneyi temizle</button></div>`:''}`}
