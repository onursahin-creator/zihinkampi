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
