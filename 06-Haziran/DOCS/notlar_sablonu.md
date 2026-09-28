# SKF Notlar Şablonu

Bu şablon, `SKF Başlıkları.xlsx` dosyasındaki "Notlar" sayfalarına yeni kurallar
eklerken kullanılacak standart formattır.

---

## Şablon Yapısı

Her not maddesi aşağıdaki 3 bölümü içermelidir:

### KURAL (Ne yapılacak?)
> Kuralın kısa ve net açıklaması. Belirsizliğe yer bırakmayın.

### GİRDİ (Hangi veri?)
> Kuralın uygulanacağı veri kaynağı, bölge, sütun, koşul.

### BEKLENEN ÇIKTI (Sonuç ne olmalı?)
> Kuralın uygulanması sonucunda çıktıda ne görülmesi gerektiği.

---

## Örnek: İyi Yazılmış Not

```
NOT 12 #1

KURAL:
  Sakarya EDAŞ için Aktif Enerji Tüketim (kWh) değeri, kaynak
  dosyadaki "Dağıtım Miktarı" başlığından ETSO bazlı olacak
  şekilde getirilmelidir. Başka bir sütunla herhangi bir işlem
  yapılmamalıdır.

GİRDİ:
  - Bölge: Sakarya EDAŞ
  - Kaynak sütun: Dağıtım Miktarı (Sütun: AR)
  - Koşul: Tüm ETSO kodları için geçerli

BEKLENEN ÇIKTI:
  - Çıktıdaki "Aktif Enerji Tüketim (kWh)" sütunu yalnız
    "Dağıtım Miktarı" sütunundaki değeri içerir
  - Trafo kaybı veya başka sütunlarla toplama yapılmaz
  - Sayısal format: Türkçe virgül → ondalık dönüşümü uygulanır
```

---

## Örnek: Kötü Yazılmış Not (YAPMAYIN)

```
NOT X #Y

Sakarya'da aktif enerji değeri yanlış geliyor, düzeltin.
```

> ❌ Hangi sütun? Hangi koşul? Beklenen sonuç ne?

---

## Kontrol Listesi

Yeni not eklerken aşağıdakileri kontrol edin:

- [ ] Hangi bölge(ler) etkileniyor? (Tümü mü, belirli bölge mi?)
- [ ] Kaynak dosyadaki hangi sütun/başlık kullanılacak?
- [ ] Çıktıdaki hangi sütun etkilenecek?
- [ ] Sayısal bir hesaplama var mı? (Toplama, çıkarma, bölme?)
- [ ] Özel koşul var mı? (Tek terimli, çift terimli, format farkı?)
- [ ] Diğer bölgeleri etkiler mi?
- [ ] Mevcut bir not ile çelişiyor mu?

---

## Notlar Sayfa İsimlendirme

| Sayfa | İçerik |
|---|---|
| Notlar 1-10 | Temel kurallar |
| Notlar 11 | KDV Matrahı, Dicle dual format, ENO |
| Notlar 12 | Sakarya aktif enerji, Dicle dinamik format, standart dışı koruma |
| Notlar 13+ | Yeni kurallar |
