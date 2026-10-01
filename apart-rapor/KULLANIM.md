# Apart Günlük Rapor – Kullanım Kılavuzu

Dosyalar
- `Apart_Gunluk_Rapor_SABLON.xlsx` – boş, kullanıma hazır dosya (resepsiyon bunu kullanır)
- `Apart_Gunluk_Rapor_ORNEK_30.09.2026.xlsx` – 30.09.2026 verisiyle dolu örnek (06-07-08 giden satırları örnektir)
- `olustur.py` – dosyaları üreten betik (`pip install openpyxl`, sonra `python3 olustur.py`)

Gereksinim: Excel 2019 / Microsoft 365 (TEXTJOIN fonksiyonu). Makro yoktur.

## Sayfalar
| Sayfa | Kim kullanır | Ne yapılır |
|---|---|---|
| **Rapor** | Resepsiyon | Sadece **tarihi** (A4) ve **kirli boş oda sayısını** (I6) girer. Geri kalan her şey otomatik; yazdır / PDF al. |
| **Konaklama** | Resepsiyon | Her konaklama için bir satır: oda, isim, kişi, giriş, çıkış, durum, tutar, kaynak. |
| **Ödemeler** | Resepsiyon | Alınan her ödeme için bir satır. |
| **Ayarlar** | Yönetici | Oda listesi, kaynak listesi (ÖĞRENCİ, MÜNFERİT…), ödeme türleri, azami kişi. |

## Günlük akış
1. Yeni misafir/rezervasyon → **Konaklama**'ya yeni satır (gelecek olanlar için Giriş Durumu = *Bekleniyor*).
2. Misafir gelince durumu *Geldi* yapın; gelmezse *İptal*.
3. Ödeme alınca **Ödemeler**'e satır ekleyin.
4. Gün sonunda **Rapor**'da tarihi seçin, kirli boş oda sayısını girin. Üstteki bant yeşil (✔) olmalı.
5. Giden misafirin satırını **silmeyin**; çıkış tarihi gelince raporda otomatik "Giden" olur.

## Sistemin engellediği hatalar
- Geçersiz/olmayan oda no, metin olarak yazılmış tarih (ör. `' 30.09.2026'`), çıkış ≤ giriş
- Aynı odaya çakışan tarihlerde ikinci misafir (çıkış günü başka misafirin girişi serbesttir)
- Kişi sayısı 1–6 dışında (Ayarlar'dan değişir), negatif tutar, isimde baş/son boşluk
- Listede olmayan kaynak/ödeme türü (`MÜNFERİT ` gibi yazım farkları)
- Misafiri olmayan odaya/tarihe ödeme girilmesi
- Kirli oda sayısı boş oda sayısından fazla olamaz

Girişte engellenemeyen durumlar (yapıştırma, sonradan tarih değiştirme) için **Kontrol** sütunu satırı kırmızı yapar ve Rapor'un üstündeki bant "⚠ DİKKAT" der.

## Hesaplar (orijinal dosyadaki mantık)
- Gün Başı = giriş < tarih ≤ çıkış · Gelen = girişi bugün ve *Geldi* · Gelecek = girişi bugün ve *Bekleniyor* · Giden = çıkışı bugün
- Gün Sonu = Gün Başı − Giden + Gelen + Gelecek (kayıtlarla ayrıca çapraz kontrol edilir)
- Boş Oda = 30 − Gün Sonu · Günlük Fiyat = Toplam Tutar ÷ Süre · Süre = Çıkış − Giriş

## Notlar
- Sayfalar parolasız korumalıdır (formüller yanlışlıkla silinmesin diye). Yönetici: Gözden Geçir → Sayfa Korumasını Kaldır.
- Kapasite: 1000 konaklama, 500 ödeme satırı; raporda günlük en fazla 20 ödeme satırı yer alır (aşılırsa uyarı verir).
- Odalar 03–32 (30 oda) kabul edildi (orijinaldeki `30 − Gün Sonu` ve boş oda listesine göre).

## Görsel rapor uygulaması (`rapor.html`)
Excel'de makro olmadan "Rapor Oluştur" butonu yapılamadığı için aynı mantığın web uygulaması da eklendi.
Tarayıcıda açın (bilgisayar/telefon), verileri girin, **📄 Rapor Oluştur** ile boş satırsız, tek parça PNG rapor görseli alın;
**PNG indir**, **Paylaş** (WhatsApp vb.) veya **Yazdır/PDF** ile gönderin. Doğrulama kuralları Excel'dekiyle aynıdır
ve hata varken rapor oluşturulmaz. Veriler tarayıcıda saklanır; Ayarlar'dan JSON yedeği alın. "Örnek veriyi yükle" ile 30.09.2026 örneği denenebilir.

### İki rapor formatı
- **Geliştirilmiş (varsayılan):** doluluk %, gün başı/gelen/gelecek/giden/gün sonu/boş kutuları, isimli hareket listeleri, konuk tablosu, ödemeler ve tahsilat toplamı, yarın özeti. Telefonda okunacak dikey formattır.
- **Klasik:** mevcut Excel raporunuzun birebir görünümü. Rapor sekmesindeki "Format" menüsünden seçilir.

- **Oda haritası** (renkli oda rack'i) varsayılan olarak **kapalı** gelir; Rapor sekmesindeki **🗺 Oda haritası** butonuyla açılıp kapatılır (görsel hazırsa anında yenilenir).
- **Gelmedi / İptal:** gelmesi beklenen misafir gelmezse "Gelmedi / İptal" yapılır; misafir rapordan düşer ve gelmesi gereken günün ertesi günü de dahil olmak üzere raporda kırmızı **GELMEDİ / İPTAL** satırında görünür (ör. 25 numaralı oda).

### Resepsiyon için kolaylıklar
- Yeni misafir formunda giriş tarihi rapor tarihi gelir; **gece sayısı** (veya 1/2/3/7/30 hızlı butonları) yazılınca çıkış tarihi otomatik hesaplanır.
- Seçilen tarihlerde dolu olan odalar listede "— dolu" görünür ve seçilemez.
- Kaynak, en son kullanılan değerle önceden seçili gelir. İsimler otomatik BÜYÜK HARF yapılır.
- Rapor sekmesinde bugün gelecek misafirler için tek dokunuşla **Geldi / İptal**.
- Boş odalardan kirli olanlar dokunarak işaretlenir; bugün çıkış yapan odalar kirli işaretli gelir, temizlenenler dokunup kaldırılır. Rapor oluşturmadan önce **Onayla** düğmesine basmak zorunludur (varsayılan yanlış olabileceği için sessizce kabul edilmez).

### Günlük kullanım (dosya mantığı)
- `ORNEK_Apart_Gunluk_Rapor_30.09.2026.html`: 30.09.2026 verileri girili hazır dosya. Çift tıklayıp tarayıcıda açın.
- Her gün **son kaydedilen dosyayı** açın. Yeni güne geçince **📅 Yeni güne geç** butonuna basın:
  1. Hata veya "Geldi/İptal" bekleyen misafir varsa sistem geçişe izin vermez; ekrandan tek dokunuşla çözülür.
  2. Dünkü tüm konaklamalar yeni güne devreder, çıkış yapanlar kayıttan düşer ve kirli boş oda işaretlenir.
  3. Resepsiyon yalnızca değişiklikleri girer: gelenler (Geldi), yeni misafirler, uzama/iptal, ödemeler.
- Gün sonunda **📄 Rapor Oluştur** ile görseli alın, ardından **💾 Günün dosyasını kaydet** ile `Apart_Gunluk_Rapor_gg.aa.yyyy.html` dosyasını indirin. Ertesi gün bu dosya açılır.
- Veri hem dosyanın içinde hem tarayıcıda saklanır; aynı bilgisayarda dosya yeniden açılırsa tarayıcıdaki daha yeni kayıt kullanılır. Farklı bilgisayara geçerken son günün dosyasını taşıyın. Eski günlerin dosyalarını arşiv olarak saklayın.
