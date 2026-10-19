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

### Sütun eşleştirme ve iskelet
Kütüphane yüklemesinde önce "Dosyayı incele" ile sütunlar ve örnek değerleri gösterilir; anlamlar (Alan, Tehlike, Risk, Olasılık, Frekans, Şiddet, Önlem, Diğer) onaylanır/düzeltilir. Excel yüklemelerinde dosyanın boş iskeleti (başlıklar, biçim, formüller) Şablonlar listesine otomatik eklenir. Risk analizi hazırlarken yönteme (Fine-Kinney / 5×5) uygun iskelet otomatik seçilir; yoksa uyarı verilir.

### Sihirbaz ve etiketler
Yeni belge oluştururken adım adım sihirbaz çalışır (risk analizi: firma → kapsam → yöntem/şablon → madde sayısı → özet; diğer belgeler: firma → sorular → şablon → özet). İnşaat için aşama/iş kalemi etiketleri (kalıp, demir, iskele, elektrik…) kütüphaneye otomatik atanır; kalem başına benzersiz madde sayısı görünür. Etiketsizler isteğe bağlı yapay zekâyla sınıflanır.

### Risk Değerlendirme Prosedürü şablonu
Şablonlar sekmesinde "Risk Değerlendirme Prosedürü" türünde örnek Word yüklenir. Risk analizi sonuç ekranındaki "Prosedür iskeleti" seçiminden seçilince prosedür bu iskeletin başlık/biçimi korunarak, firma, yöntem ve risk sonuçlarına göre yeniden yazılır. İskelet yoksa metin iki parça halinde (tam bölümlerle) sıfırdan yazılır.

### Firma bilgisi temizliği
Şablondan kalan firma unvanı, proje adı, SGK/sicil numarası; Excel başlık bloğunda, Word üst/alt bilgisinde ve metinde seçilen firmayla değiştirilir (firma kaydındaki "Proje / şantiye adı" alanı kullanılır, boşsa temizlenir). Önlem sonrası O/F/Ş, sorumlu, termin, etkilenenler ve açıklama hücreleri boşsa kural tabanlı uzman varsayılanlarıyla doldurulur.

### Plan / liste tipi Excel şablonları (yıllık çalışma, eğitim planı vb.)
Excel şablonunun yapısı otomatik çözülür: başlık bloğu (etiket altındaki değer hücreleri), başlık satırları (iki seviyeli olabilir), ay sütunları (X işaretleri), bölüm satırları ve dipnotlar. Yapay zekâ satırları bölüm bölüm (parça parça) üretir, iskeletin bölüm/sütun düzenine sadık kalır; çıktı iskeletin biçimiyle (bölüm rengi, X hücre rengi, birleşik hücreler, yazdırma alanı) yazılır. Üretilen satırlar "Şablon alanları" bölümünde düzenlenebilir. Eski yüklenmiş şablonlar ilk kullanımda otomatik çözümlenir.

### Şablon türüne göre motor seçimi
Risk türündeki Excel şablonları risk motoruyla, diğer Excel şablonları (plan, liste, form) önce plan motoruyla (bölüm/ay yapısı), yapı bulunamazsa hücre adresli form motoruyla doldurulur. Word şablonları alan/tablo motoruyla doldurulur. Yapay zekâ, parça başına satır sayısı çok düşerse yeniden dener; yine düşükse şablonun kendi satırlarını kullanır.
