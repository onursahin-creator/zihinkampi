# Club Müskebi – Günlük Doluluk Raporu (`muskebi-rapor.html`)

Mevcut günlük rapor (9 günlük doluluk grafiği + "Şimdiki durum" ve "Doluluk" tabloları) aynı içerikle, otomatik hesaplanan ve hata kontrolü yapan bir uygulamaya çevrildi. Tarayıcıda açılır (bilgisayar/telefon), kurulum gerekmez.

Dosyalar: `muskebi-rapor.html` (boş uygulama), `ORNEK_Muskebi_Gunluk_Rapor_30.09.2026.html` (30.09.2026 verisi girili), `ORNEK_Muskebi_Rapor_30.09.2026.png` (örnek çıktı).

## Her gün girilenler
1. **Sabahki durum** (oda/kişi): dünkü gün sonundan otomatik gelir.
2. **Gerçekleşen gelişler** ve **bugün ayrılmış olanlar** (oda/kişi).
3. **Beklenen gelişler/çıkışlar** (yoksa 0).
4. **Günlük gelir** (₺); isteğe bağlı kişi geceleme (boşsa gün sonu kişi sayısı kullanılır).
5. **Önümüzdeki 8 günün dolu oda sayısı** (rezervasyon durumuna göre).

## Otomatik hesaplananlar
- Şimdiki durum = sabah + gerçekleşen gelen − ayrılan · Gün sonu = şimdiki durum + beklenen gelen − beklenen çıkan
- Oda % = oda ÷ toplam oda · Yatak % = kişi geceleme ÷ toplam yatak
- Ort. kişi fiyatı = gelir ÷ kişi geceleme · Ort. oda fiyatı = gelir ÷ gün sonu oda
- Grafikte ilk sütun bugünün gün sonu, sonraki 8 sütun girilen tahminlerdir (kapasite çizgisi ve gün bazlı % ile).

## Hatayı engelleyen kontroller
Eksik alan, tam sayı olmayan veya negatif değer, kapasiteyi (166 oda) aşan oda, oda başına 1–4 kişi dışındaki kişi sayısı, sabahtan fazla ayrılan, şimdiki durumdan fazla beklenen çıkış, gün sonu odası varken gelir 0, tahmin eksikliği. Ortalama oda fiyatı olağan aralığın dışındaysa veya sabahki durum dünkü gün sonuyla farklıysa uyarı çıkar. Hata yoksa **"Rakamları kontrol ettim, doğru"** kutusu işaretlenmeden rapor oluşturulmaz (herhangi bir alan değişince kutu sıfırlanır).

## Yeni gün
**📅 Yeni güne geç**: gün sonu → yarının sabahki durumu olur, tahmin tablosu bir gün kayar (yalnızca yeni 8. gün girilir). Gün sonunda **Rapor Oluştur** → PNG indir/paylaş, ardından **Günün dosyasını kaydet**; ertesi gün o dosya açılır.

## Varsayımlar (lütfen doğrulayın)
- 166 oda / 332 yatak: örnek raporun yüzdelerinden çıkarıldı (11/166 = %6,63; 18/332 = %5,42). Ayarlar'dan değiştirilir.
- Örnekte gün sonu 19 kişi, kişi geceleme 18 görünüyor; kişi geceleme ayrı alan olarak bırakıldı (ör. ücretsiz çocuk).
- Olağan ort. oda fiyatı aralığı 1.500–30.000 ₺ varsayıldı.
- Tahminler elle girilir. Rezervasyon sisteminden (PMS) otomatik alınması için örnek bir dışa aktarım dosyası gerekir.
