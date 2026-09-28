# SKF Sistemi — AI Ajanı Dar Kapsamlı Kontrol Sözleşmesi

**Versiyon:** 2.0  
**Tarih:** 24 Eylül 2026  
**Amaç:** Yapay zeka ajanının token ekonomisini korumak, bağlam çöküşünü (context collapse) engellemek, yeni modüler mimari sınırlarına uymasını sağlamak ve regresyon riskini sıfıra indirmek

---

## 1. Proje Bağlamı (System Prompt Context)

### 1.1 Tek Paragraf Özet

SKF, 21 farklı Türk elektrik dağıtım şirketinin (EDAŞ) aylık fatura verilerini (xlsx/xls/html/xml) tek bir standart Excel çıktısına birleştiren Python+Streamlit uygulamasıdır. Sistem; `SKF Başlıkları.xlsx` dosyasındaki yapılandırma-odaklı mapping ve `config/header_overrides.json` dinamik yapılandırması üzerinden her EDAŞ'ın farklı sütun isimlerini ve formatlarını standart 18 sütunlu çıktıya dönüştürür. Mimari; `config/` ve `core/` altında modüler katmanlara ayrılmıştır, `extract_and_compare.py` ise geriye dönük uyumlu bir cephe (facade) olarak görev yapar. Kanonik çıktı: **2593 kayıt, 21 bölge, 0 hata**.

### 1.2 Dosya Haritası ve Modüler Mimari

```
SKF/
├── config/
│   ├── constants.py            → Sabitler (BASE_PATH, STANDARD_COLUMNS, SUM_FIELDS, MAX_FIELDS vb.)
│   ├── mappings.py             → Mapping sözlükleri (header_overrides.json'dan yüklenir, fallback içerir)
│   └── header_overrides.json   → Başlık eşleştirmeleri ve alan sinonimleri için düzenlenebilir JSON yapılandırması
├── core/
│   ├── normalizer.py           → Bölge adı ve ETSO normalizasyonu
│   ├── number_cleaner.py       → Türkçe sayı temizleme (clean_turkish_number), başlık normalizasyonu
│   ├── column_resolver.py      → Dinamik sütun çözümleme (resolve_column_indices)
│   ├── reactive_calculator.py  → Reaktif ve İlk Reaktif toplamlar (DRY, tek kaynak)
│   ├── eno_enricher.py         → ENO dosyası zenginleştirmesi (enrich_from_eno_file)
│   └── excel_writer.py         → Excel çıktısı oluşturma ve biçimlendirme (save_to_excel)
├── extract_and_compare.py      → Facade (Geriye dönük uyumlu çekirdek arayüz)
├── app.py                      → Sunum: Streamlit UI, sidebar, pipeline tetikleme
├── new_mapping_parser.py       → Yapılandırma: SKF Başlıkları.xlsx → mapping dict
├── tests/                      → Test paketi: 61 test (Golden Core: 2593 kayıt, 21 bölge doğrulaması)
├── SKF Başlıkları.xlsx         → Veri-odaklı config: Başlıklar (21+1 bölge), Düzeltme, Notlar
└── XX ENO AyAdı.xlsx           → Lookup: Sayaç ID → EIC Kod, EIC → Abone Ad-Soyad
```

### 1.3 Veri Akışı

```
Kullanıcı Klasör/Parametre Seçer (app.py)
                │
                ▼
new_mapping_parser.py + config/mappings.py (header_overrides.json)
                │ (Mapping sözlükleri & şemalar yüklenir)
                ▼
extract_and_compare.py (Facade / process_all_regions)
                │
                ├─► Format Okuyucular (xlsx, xls, html, xml)
                │     ├── core/column_resolver.py     (Sütun indekslerini tespit eder)
                │     ├── core/number_cleaner.py      (Türkçe sayı formatlarını temizler)
                │     └── core/reactive_calculator.py (Reaktif/İlk Reaktif toplamları hesaplar - DRY)
                │
                ├─► core/normalizer.py   (Bölge adlarını ve ETSO kodlarını standartlaştırır)
                ├─► core/eno_enricher.py (ENO dosyasından Sayaç/EIC/Abone zenginleştirmesi)
                │
                ▼
core/excel_writer.py (save_to_excel -> Formüllü 18 sütunlu çıktı)
                │
                ▼
Streamlit UI Sonuç Gösterimi & İndirme (app.py)
```

---

## 2. Kodlama Standartları

### 2.1 İsimlendirme ve Dosya Organizasyonu

| Tür | Kural | Örnek |
|---|---|---|
| Modül / Dosya | snake_case | `reactive_calculator.py`, `column_resolver.py` |
| Fonksiyon | snake_case, fiil ile başla | `resolve_column_indices()`, `calculate_reactive_totals()` |
| Sabit | UPPER_SNAKE_CASE | `STANDARD_COLUMNS`, `REGION_NAME_MAPPING` |
| İç fonksiyon | `_` prefix | `_find_eno_file()`, `_format_guc_kw()` |
| Değişken | snake_case, anlamlı | `etso_idx`, `reactive_total`, `region_mapping` |
| Dictionary key | Orijinal Türkçe başlık | `'Dağıtım Bölgesi'`, `'Etso Kodu'`, `'Güç kW'` |

### 2.2 Modüler Kod Yerleşimi Kuralları

1. **ASLA `extract_and_compare.py` içine yeni kod YAZILMAZ**:
   - `extract_and_compare.py` sadece geriye dönük uyumluluk sağlayan bir facade katmanıdır.
   - Tüm yeni kodlar ve iş mantığı doğrudan ilgili modüle (`core/` veya `config/`) eklenmelidir.
2. **Yeni Reaktif Mantık → `core/reactive_calculator.py`**:
   - 4 format okuyucunun (xlsx, xls, html, xml) ortak reaktif ve ilk reaktif hesaplama mantığı tek kaynaktadır (DRY).
   - Reaktif hesaplama ile ilgili her türlü değişiklik yalnızca `core/reactive_calculator.py` üzerinde yapılmalıdır.
3. **Yeni Başlık Eşleştirmeleri → `config/header_overrides.json`**:
   - EDAŞ'lara özel başlık eşleştirmeleri koda ASLA hardcode edilmez.
   - Yeni eşleştirmeler doğrudan `config/header_overrides.json` dosyasındaki `SPECIAL_HEADER_MAPPING` altına eklenir.
4. **Yeni Alan Sinonimleri (Synonyms) → `config/header_overrides.json`**:
   - Sütun başlık varyasyonları koda gömülmez.
   - Yeni başlık sinonimleri `config/header_overrides.json` dosyasındaki `FIELD_SYNONYMS` bloğuna eklenir.
5. **Dinamik Sütun Çözümleme**:
   - Sütun indeksleri için her zaman `core/column_resolver.py` içindeki `resolve_column_indices()` kullanılır.
6. **Sayısal Değer Temizleme**:
   - Her zaman `core/number_cleaner.py` içindeki `clean_turkish_number()` kullanılır. Asla yalın `float()` kullanılmaz.

### 2.3 Hata Yönetimi

```python
# DOĞRU: Spesifik exception, kullanıcı dostu mesaj, alternatif yol
try:
    wb = openpyxl.load_workbook(fpath)
except PermissionError:
    print(f"[WARNING] {fpath} dosyası Excel'de açık!")
    alt_path = fpath.replace('.xlsx', '_Guncel.xlsx')
    wb.save(alt_path)

# YANLIŞ: Sessiz geçiş, veri kaybı riski
try:
    value = float(cell)
except:
    pass
```

### 2.4 Sayı İşleme

```python
# DOĞRU: core/number_cleaner üzerinden temizleme
from core.number_cleaner import clean_turkish_number

value = clean_turkish_number(raw_value, region_name=region, field_name='aktif_enerji')

# YANLIŞ: Direkt float() çevirme
value = float(raw_value)  # ❌ Türkçe "1.234,56" veya "(100)" formatı bozulur/hata verir
```

### 2.5 Sütun Çözümleme

```python
# DOĞRU: core/column_resolver üzerinden dinamik çözümleme
from core.column_resolver import resolve_column_indices

indices = resolve_column_indices(headers, region_mapping)

# YANLIŞ: Sabit sütun indeksi varsayımı
etso = row[9]  # ❌ Sütun sırası değiştiğinde veya farklı EDAŞ formatında bozulur
```

### 2.6 Test Zorunluluğu

- **Testler ZORUNLUDUR**: Herhangi bir commit veya değişiklik öncesinde ve sonrasında `python -m pytest tests/` komutu çalıştırılmalı ve istisnasız TÜM testler geçmelidir.
- **Mevcut Test Paketi**: 61 adet otomatikleştirilmiş test bulunmaktadır:
  - `tests/test_golden_core.py` (Kanonik çıktı doğrulaması: **2593 kayıt, 21 bölge, 0 hata**)
  - `tests/test_clean_number.py` (Sayı temizleme ve bölgesel edge case'ler)
  - `tests/test_normalizer.py` (Bölge adı ve ETSO normalizasyon testleri)
  - `tests/test_reactive_calc.py` (Reaktif ve İlk Reaktif tek kaynak hesaplama testleri)

---

## 3. Davranış Kuralları (Kırmızı Çizgiler)

### 3.1 ASLA Yapılmayacaklar (Kırmızı Çizgiler / Red Lines)

| # | Kırmızı Çizgi | Gerekçe / Detay |
|---|---|---|
| 🔴 1 | **Web upload / file picker bileşeni EKLEME** | **KVKK & Gizlilik**: ETSO ve müşteri tüketim verileri gizlidir; veriler kesinlikle yerel ortamda kalmalıdır. |
| 🔴 2 | **Mapping dosyası gömmeyi (embedding) KALDIRMA** | Sistem `SKF Başlıkları.xlsx` dosyasındaki dinamik yapılandırma ve gömülü kurallarla çalışır. |
| 🔴 3 | **UI içine AI asistanı EKLEME** | Arayüz temiz, deterministik ve operasyon odaklı kalmalıdır; chat/AI widget'ı eklenemez. |
| 🔴 4 | **`extract_and_compare.py` import geriye dönük uyumluluğunu BOZMA** | Harici script'ler ve UI bileşenleri bu modüldeki fonksiyonları doğrudan çağırır (Facade deseni). |
| 🔴 5 | **`core/reactive_calculator.py` dosyasını 61 testin tümünü koşturmadan DEĞİŞTİRME** | 4 format okuyucunun tek kaynağıdır; yapılacak en ufak hata 21 bölgenin faturasına etki eder. |
| 🔴 6 | **Testler geçmeden `main` branch'ine COMMIT/MERGE YAPMA** | Canlı kod bütünlüğü Golden Core (2593 kayıt, 21 bölge) ile garanti altındadır. |
| 🔴 7 | **`extract_and_compare.py` içine yeni kod/iş mantığı EKLEME** | Monolitik yapı terk edilmiştir; yeni kodlar ilgili modüle (`core/`, `config/`) gitmelidir. |
| 🔴 8 | **Başlık eşleştirmelerini koda HARDCODE ETME** | Yeni eşleştirmeler ve sinonimler `config/header_overrides.json` üzerinden dinamik yönetilmelidir. |
| 🔴 9 | **Mevcut çıkarma/birleştirme mantığını bozarak Golden Core'dan sapma** | Çıktı kesinlikle **2593 kayıt, 21 bölge, 0 hata** olmalıdır. |
| 🔴 10 | **`clean_turkish_number` sayı temizleme mantığını bozma** | 21 EDAŞ'ın tüm edge-case ve istisna formatları bu mantıkla çözülmüştür. |
| 🔴 11 | **10 Milyar TL filtresini kaldırma** | Aykırı ve hatalı verileri engelleyen kritik bütünlük filtresidir. |
| 🔴 12 | **Standart dışı dosyada tahmin yürütme** | Notlar 12 #3 kuralı: Format okunamıyorsa uydurma veri üretilmez, okunamadığı belirtilir. |

### 3.2 HER ZAMAN Yapılacaklar

| # | Kural | Gerekçe |
|---|---|---|
| 🟢 1 | **Her değişiklik öncesi ve sonrası `python -m pytest tests/` çalıştır** | 61 testin tamamı (Golden Core dahil) yeşil olmalıdır. |
| 🟢 2 | **Yeni başlık veya alan sinonimi ihtiyacında `config/header_overrides.json` kullan** | Kod değişikliği yapmadan yapılandırma ile çözüm ilkesi. |
| 🟢 3 | **Reaktif mantık güncellemelerini `core/reactive_calculator.py` içinde yap** | Tek kaynak (Single Source of Truth) ve DRY prensibi. |
| 🟢 4 | **Yeni sütun eklerken `STANDARD_COLUMNS` (`config/constants.py`) ve Excel formüllerini (`core/excel_writer.py`) senkronize güncelle** | Excel çıktısında sütun ve formül kayması riskini engelleme. |
| 🟢 5 | **Her değişikliği `PROJE_DOKUMANTASYONU.md` ve teknik dokümanlara tarih damgasıyla işle** | Tam izlenebilirlik ve denetim uyumluluğu. |
| 🟢 6 | **`extract_and_compare.py` modülünü geriye dönük uyumlu facade olarak koru** | Eski fonksiyon imzalarını ve içe aktarımları kırmama garantisi. |

### 3.3 Modül Sınırları ve Sorumluluk Matrisi (Context Boundaries)

```
┌────────────────────────────────────────────────────────────────────────┐
│ CONFIG KATMANI                                                         │
│ ├─ config/constants.py         → SADECE global sabitler ve kolon listeleri│
│ ├─ config/header_overrides.json→ SADECE başlık/sinonim JSON yapılandırması│
│ └─ config/mappings.py          → SADECE bölge eşleştirmeleri ve lookup'lar│
└──────────────────────────────────┬─────────────────────────────────────┘
                                   │
┌──────────────────────────────────▼─────────────────────────────────────┐
│ CORE KATMANI                                                           │
│ ├─ core/normalizer.py          → SADECE bölge ve ETSO normalizasyonu   │
│ ├─ core/number_cleaner.py      → SADECE sayısal veri ve başlık temizliği│
│ ├─ core/column_resolver.py     → SADECE dinamik sütun indeksi tespiti  │
│ ├─ core/reactive_calculator.py → SADECE reaktif & ilk reaktif hesabı (DRY)│
│ ├─ core/eno_enricher.py        → SADECE ENO lookup ve veri zenginleştirme│
│ └─ core/excel_writer.py        → SADECE Excel dosya yazımı ve biçimlendirme│
└──────────────────────────────────┬─────────────────────────────────────┘
                                   │
┌──────────────────────────────────▼─────────────────────────────────────┐
│ ENTEGRASYON VE ÇEKİRDEK CEPHE (FACADE)                                │
│ ├─ new_mapping_parser.py       → SADECE SKF Başlıkları.xlsx ayrıştırma  │
│ └─ extract_and_compare.py      → Geriye dönük uyumlu facade & orkestrasyon│
└──────────────────────────────────┬─────────────────────────────────────┘
                                   │
┌──────────────────────────────────▼─────────────────────────────────────┐
│ SUNUM KATMANI                                                          │
│ └─ app.py                      → SADECE Streamlit UI & kullanıcı etkileşimi│
└────────────────────────────────────────────────────────────────────────┘
```

---

## 4. Git İş Akışı (Git Workflow)

Geliştirme sürecinde kararlılık ve kod kalitesini korumak için aşağıdaki Git akışı zorunludur:

```
[dev branch] ───► [feature/fix branch] ───► [Testler: 61/61 PASS] ───► [dev] ───► [main]
```

1. **Ana Geliştirme `dev` Branch'indedir**:
   - Tüm aktif geliştirmeler `dev` branch'inde yürütülür.
   - Doğrudan `main` branch'i üzerinde geliştirme YAPILMAZ.
2. **Feature Branch Stratejisi**:
   - Yeni özellikler ve düzeltmeler için `dev` branch'inden dallanılır (`feature/<ozellik-adi>` veya `fix/<hata-adi>`).
3. **Commit Öncesi Test Zorunluluğu**:
   - Commit atılmadan önce yerel ortamda `python -m pytest tests/` komutu çalıştırılmalı ve 61 testin tümü başarıyla tamamlanmalıdır.
4. **`main` Branch'ine Birleştirme (Merge)**:
   - `dev` branch'i, yalnızca tüm testler (özellikle Golden Core 2593 kayıt, 21 bölge doğrulaması) eksiksiz geçtiğinde `main` branch'ine merge edilebilir.
   - Testleri geçmeyen hiçbir kod `main` branch'ine commit veya merge EDİLEMEZ.

---

## 5. Ajan Görev Şablonu (Task Template)

Yapay zeka ajanı kendisine verilen her görevde aşağıdaki adımları sırayla işletmelidir:

```markdown
## Görev: [Kısa açıklama]

### 1. Analiz ve Hazırlık
- [ ] Git branch kontrolü yap (`dev` veya `feature/*` üzerinde çalış)
- [ ] İlgili modülü belirle (`config/`, `core/`, `app.py`, `new_mapping_parser.py`)
- [ ] Etkilenen fonksiyonları listele
- [ ] Golden Core etkisini değerlendir (Kayıt sayısı: 2593, Bölge sayısı: 21)

### 2. Uygulama (Modüler İzolasyon)
- [ ] Yeni kodu SADECE ilgili modüle yaz (`extract_and_compare.py` içine yeni kod ekleme)
- [ ] Başlık eşleştirmesi veya sinonim gerekiyorsa `config/header_overrides.json`'a ekle
- [ ] Reaktif hesaplama ile ilgiliyse `core/reactive_calculator.py`'yi güncelle
- [ ] `extract_and_compare.py` facade geriye dönük uyumluluğunu koru
- [ ] Mevcut yorum satırlarını ve docstring'leri koru

### 3. Doğrulama (Kalite Kapısı)
- [ ] Syntax kontrolü: `python -m py_compile <degistirilen_dosyalar>`
- [ ] Test paketi koştur: `python -m pytest tests/` (61 testin tamamı geçmeli)
- [ ] Golden Core doğrulaması: 2593 kayıt, 21 bölge, 0 hata kontrolü

### 4. Dokümantasyon ve Git
- [ ] Değişiklikleri ilgili dokümantasyona tarih damgası ile kaydet
- [ ] Git commit (açıklayıcı commit mesajı ile)
- [ ] Yalnızca testler başarılı olduğunda merge sürecini tamamla
```

---

*Bu sözleşme, SKF sisteminde çalışan tüm AI ajanlarının uymakla yükümlü olduğu temel teknik sözleşmedir (v2.0). Sözleşme dışı davranış üretildiğinde ajan durdurulmalı ve bağlam yeniden daraltılmalıdır.*
