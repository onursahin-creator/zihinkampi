# Club Müskebi – Günlük Doluluk Raporu (`muskebi-rapor.html`)

Mevcut günlük rapor (9 günlük doluluk grafiği + "Şimdiki durum" ve "Doluluk" tabloları) aynı içerikle, otomatik hesaplanan ve hata kontrolü yapan bir uygulamaya çevrildi. Tarayıcıda açılır (bilgisayar/telefon), kurulum gerekmez.

Dosyalar: `muskebi-rapor.html` (boş uygulama), `ORNEK_Muskebi_Gunluk_Rapor_30.09.2026.html` (30.09.2026 verisi girili), `ORNEK_Muskebi_Rapor_30.09.2026.png` (örnek çıktı).

## En kolay yol: forecast dosyasını yükleyin
PMS'ten alınan **"Müsaitliğe göre forecast" (.xlsx)** dosyasını sayfanın üstündeki **📂 yükleme alanına sürükleyin** (veya "Dosya seç"). Dosyadan otomatik okunanlar:
- Rapor tarihi (dosyadaki ilk gün) ve **8 günlük dolu oda tahmini** (grafik)
- **Gün sonu**: Dolu Oda ve Toplam Kişi · **Kişi geceleme**: Yetişkin + Çocuk (bebek hariç) · **Günlük gelir**: Oda Geliri
- Dosyadaki "Gelen Oda" değeri, "Gerçekleşen gelişler — oda" alanına önerilir

Resepsiyonun elle girdikleri (en fazla 2 alan):
- **Sabahki durum** (oda/kişi): dünkü kayıt varsa otomatik gelir; ilk gün bir kez elle girilir.
- **Gerçekleşen gelişler — kişi sayısı** (dosyada kişi sayısı yok).
- Gün içi rapor gönderiliyorsa **beklenen gelişler/çıkışlar** (yoksa 0).
**Bugün ayrılmış olanlar otomatik hesaplanır**: sabah + gelen − gün sonu. Sonuç mantıksız çıkarsa (ör. negatif ayrılan) sistem hata verir.

Dosya kontrolleri: .xlsx olmayan dosya, forecast olmayan dosya (sütun adları), 9 günden az satır, ardışık olmayan tarih, kapasiteyi aşan dolu oda, oda varken gelir 0 reddedilir. Dosya adındaki tarih ile ilk gün farklıysa, toplam kişi ≠ yetişkin+çocuk+bebek ise veya kapasite Ayarlar'dan farklıysa uyarı verir. Aynı gün için dosya yeniden yüklenirse (ör. sabah ve akşam) dosya kaynaklı değerler güncellenir, elle girilenler korunur.

Ardından **"Rakamları kontrol ettim, doğru"** → **📄 Rapor Oluştur** → PNG indir/paylaş → **💾 Günün dosyasını kaydet**. Ertesi gün yeni dosyayı yüklemek yeterlidir; sabahki durum dünkü kayıttan gelir.

## Elle giriş (dosya yoksa)
### Her gün girilenler
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
- Örnek raporda gün sonu 19 kişi, kişi geceleme 18 idi. Forecast dosyasında Yetişkin+Çocuk ve Bebek ayrı sütunlar olduğundan, kişi geceleme = Yetişkin + Çocuk (bebek hariç) varsayıldı. 30.09 dosyasıyla doğrulanmadı.
- Dosyadaki "Gelen Oda" bugünkü toplam gelen oda sayısı varsayıldı (gün sonu alınan dosyada tamamı gerçekleşmiş sayılır). Dosya sabah alınıyorsa gelmemiş olanlar "Beklenen gelişler" alanına taşınmalıdır.
- Olağan ort. oda fiyatı aralığı 1.500–30.000 ₺ varsayıldı.
- Tahminler elle girilir. Rezervasyon sisteminden (PMS) otomatik alınması için örnek bir dışa aktarım dosyası gerekir.
