# İSG Evrak Asistanı

Tarayıcıda çalışan, sunucusuz (statik) İSG evrak hazırlama sistemi. Veriler kullanıcının tarayıcısında saklanır.

## GitHub Pages ile yayınlama
1. Bu klasörün **içindeki** tüm dosyaları (index.html, css/, js/, .nojekyll) yeni bir GitHub deposunun ana dizinine yükleyin.
2. Depo → Settings → Pages → Source: `Deploy from a branch`, Branch: `main` / `(root)`.
3. Kendi alan adınız için Pages ayarında *Custom domain* alanına alan adını yazın; alan adı sağlayıcınızda CNAME kaydını `KULLANICIADI.github.io` adresine yönlendirin.
4. Siteyi açıp **Ayarlar** sekmesine Anthropic API anahtarınızı girin.

## Dosyalar
- `index.html` – iskelet, `css/style.css` – görünüm
- `js/core.js` – durum, depolama, yapay zekâ çağrısı, risk hesapları
- `js/skeleton.js` – Word/Excel şablon iskeleti çözümleme ve doldurma
- `js/export.js` – Word/Excel çıktıları, acil durum ekip tabloları
- `js/app.js` – ekranlar

## Şablon mantığı
Yüklenen .docx/.xlsx şablonun tabloları, başlıkları, sütun genişlikleri, renkleri ve Excel formülleri aynen korunur. Firma bilgileri etiketlere (Firma Adı, Adres, SGK…) yazılır; liste tabloları ve Excel veri satırları şablondaki son satırın biçimiyle çoğaltılır.

## Kütüphane
Kütüphane sekmesine yüklenen uzman risk analizleri (Excel/Word) kodla okunur, iş kolu etiketlenir, benzer maddeler birleştirilir. Risk analizi hazırlarken önce kütüphane maddeleri kullanılır; eksik kalan madde sayısı yapay zekâyla tamamlanır. Veriler tarayıcıda (IndexedDB) tutulur.
