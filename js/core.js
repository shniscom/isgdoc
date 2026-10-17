// Çekirdek: durum, depolama (localStorage + IndexedDB), yapay zekâ, hesaplar
const $=s=>document.querySelector(s),m=$('#m'),e=x=>String(x??'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;'),uid=()=>Date.now().toString(36)+Math.random().toString(36).slice(2,5);
const TY={risk:['Risk Analizi','y'],acil:['Acil Durum Eylem Planı','y'],cal:['Yıllık Çalışma Planı','y'],egp:['Yıllık Eğitim Planı','y'],deg:['Yıllık Değerlendirme Raporu','y'],
at_ct:['Çalışan Temsilcisi Atama Yazısı','a'],at_de:['Destek Elemanı Atama Yazısı','a'],at_bs:['Bilgi Sahibi Çalışan Atama Yazısı','a'],at_iv:['İşveren Vekili Atama Yazısı','a'],at_re:['Risk Değerlendirme Ekibi Atama Yazısı','a'],at_ek:['Acil Durum Ekibi Atama Yazısı','a'],
eg_or:['Oryantasyon Eğitimi','e'],eg_ts:['Temel İSG Eğitimi','e'],eg_re:['Risk Değerlendirme Ekibi Eğitimi','e'],eg_de:['Destek Elemanı Eğitimi','e'],eg_ct:['Çalışan Temsilcisi Eğitimi','e'],eg_bs:['Bilgi Sahibi Çalışan Eğitimi','e'],eg_iy:['İlk Yardım Eğitimi','e'],eg_yg:['Yangın ve Tahliye Eğitimi','e']};
const TABS=[['y','Yasal Evrak Hazırla'],['a','Personel Atamaları'],['e','Eğitim Dokümanları'],['k','Kaydedilenler'],['l','Kütüphane'],['t','Şablonlar'],['s','Ayarlar']];
const TH=['Az Tehlikeli','Tehlikeli','Çok Tehlikeli'],SEC=['İnşaat','Metal / Demir-Çelik','Gıda','Tekstil','Mobilya / Ahşap','Ofis / Hizmet','Depolama / Lojistik','Sağlık','Eğitim','Otomotiv Servis','Kimya','Madencilik'];
const ILK=[20,15,10],DIG=[50,40,30]; // Oranlar güncel yönetmelikten teyit edilmelidir
const EK=['Söndürme Ekibi','Koruma Ekibi','Arama-Kurtarma Ekibi','İlk Yardım Ekibi'],EKR=[/söndürme|yangın/i,/koruma/i,/arama|kurtarma/i,/ilk\s?yard[ıi]m/i];
const SYS='Sen Türkiye mevzuatına (6331 sayılı İSG Kanunu, Risk Değerlendirmesi Yönetmeliği, Acil Durumlar Hakkında Yönetmelik, İlkyardım Yönetmeliği, Çalışanların İSG Eğitimlerinin Usul ve Esasları Hakkında Yönetmelik, İSG Hizmetleri Yönetmeliği) hâkim, deneyimli bir A sınıfı iş güvenliği uzmanısın. Türkçe, resmî ve uygulanabilir çıktı üret.';
let S=JSON.parse(localStorage.getItem('isg2')||'null');
if(!S){const o=JSON.parse(localStorage.getItem('isg')||'{}');S={d:o.d||[],c:[],k:o.k||'',m:o.m||'claude-sonnet-5-5'}}
S.ls=S.ls||[];
const sv=()=>localStorage.setItem('isg2',JSON.stringify(S));
let TP=[],cur=null,tab='y',busy=0;
// Şablonlar ikili dosya olduğu için IndexedDB'de tutulur
const DB=new Promise(r=>{const q=indexedDB.open('isgdb',2);q.onupgradeneeded=()=>{const d=q.result;['t','k'].forEach(n=>{if(!d.objectStoreNames.contains(n))d.createObjectStore(n,{keyPath:'id'})})};q.onsuccess=()=>r(q.result)});
const tx=async(mode,fn,st='t')=>{const db=await DB;return new Promise((res,rej)=>{const t=db.transaction(st,mode),q=fn(t.objectStore(st));t.oncomplete=()=>res(q&&q.result);t.onerror=()=>rej(t.error)})};
const kAll=()=>tx('readonly',s=>s.getAll(),'k'),kPut=a=>tx('readwrite',s=>{a.forEach(i=>s.put(i))},'k'),kDel=a=>tx('readwrite',s=>{a.forEach(i=>s.delete(i))},'k');
const tAll=()=>tx('readonly',s=>s.getAll()),tPut=o=>tx('readwrite',s=>s.put(o)),tDel=id=>tx('readwrite',s=>s.delete(id));
async function ai(p,mx=8000){if(!S.k){alert('Önce Ayarlar bölümüne Anthropic API anahtarınızı girin.');throw 0}
 for(let t=0;t<2;t++){const r=await fetch('https://api.anthropic.com/v1/messages',{method:'POST',headers:{'content-type':'application/json','x-api-key':S.k,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},body:JSON.stringify({model:S.m,max_tokens:mx,system:SYS,messages:[{role:'user',content:p}]})});
  const j=await r.json();if(j.error){if(!t&&/overload|rate/i.test(j.error.message)){await new Promise(z=>setTimeout(z,3000));continue}throw new Error(j.error.message)}return j.content.map(c=>c.text||'').join('')}}
const jp=s=>{s=s.slice(Math.max(0,s.search(/[\[{]/)));const a=s[0]=='[';try{return JSON.parse(s.slice(0,s.lastIndexOf(a?']':'}')+1))}catch(x){if(a)return JSON.parse(s.slice(0,s.lastIndexOf('}')+1)+']');throw x}};
async function run(btn,fn){if(busy)return;busy=1;const o=btn.textContent;btn.textContent='Hazırlanıyor…';try{await fn()}catch(x){if(x)alert('Hata: '+(x.message||x))}busy=0;btn.textContent=o}
const pg=(p,t)=>{const b=$('#pg');if(b)b.innerHTML=`<div class="pb"><i style="width:${p}%"></i></div><small>%${p} — ${e(t)}</small>`};
// Risk hesapları (yöntem her zaman açıkça verilir)
const sk=(r,m)=>(+r.o)*(m=='fk'?+r.f:1)*(+r.s);
const lv=(v,m)=>m=='fk'?(v>400?'Çok Yüksek':v>200?'Yüksek':v>70?'Önemli':v>20?'Olası':'Düşük'):(v>=15?'Çok Yüksek':v>=10?'Yüksek':v>=5?'Orta':'Düşük');
const cnt=(a,f)=>a.reduce((o,x)=>(o[f(x)]=(o[f(x)]||0)+1,o),{});
const summ=d=>({alan:cnt(d.rows,r=>r.a||'-'),seviye:cnt(d.rows,r=>lv(sk(r,d.method),d.method))});
