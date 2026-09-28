# SKF Fatura Birleştirme Sistemi — Yazılım Tasarım Spesifikasyonu

**Versiyon:** 2.0  
**Tarih:** 24 Eylül 2026  
**Hazırlayan:** Principal Software Engineer (AI-Assisted Architecture)  
**Durum:** Canlı (Production) — Modüler Cephe (Progressive Facade) + Streamlit Cloud + Yerel Çalışma

---

## 1. Proje Özeti ve İş Amacı

### 1.1 Çözülen Problem

CK Enerji bünyesindeki **21 Elektrik Dağıtım Şirketi (EDAŞ)**, her ay farklı formatlarda (xlsx, xls, html, xml) fatura verileri üretmektedir. Her şirketin:
- Farklı sütun isimleri (ör: "DUY Kodu" vs "ETSO Kodu" vs "Etso Kod")
- Farklı sütun sıralamaları (ör: Dicle EDAŞ'ta ETSO col J'de, Sakarya'da col K'da)
- Farklı dosya formatları (xlsx, html-as-xls, xml-as-xls)
- Farklı sayı formatları (Türkçe virgül vs US dot, parantezli negatif)

olması, **manuel birleştirmeyi** son derece hata eğilimli ve zaman alıcı kılmaktadır.

### 1.2 Operasyonel Fayda

| Metrik | Manuel Süreç | SKF Sistemi |
|---|---|---|
| Birleştirme süresi | ~2 iş günü | <30 saniye |
| Hata oranı | Yüksek (copy-paste) | 0 (deterministic) |
| Format tutarlılığı | Manuel kontrol | Otomatik standartlaştırma |
| ETSO zenginleştirme | Ayrı süreç | Entegre (ENO dosyasından) |

### 1.3 Kanonik Çıktı (Golden Core)

Sistem, 21 farklı kaynaktan **2593 kayıt** üretir. Bu sayı, çekirdek regresyon metriğidir. Her değişiklikte bu sayı, bölge dağılımı ve mali alan toplamları doğrulanır.

---

## 2. Sistem Mimarisi ve Teknoloji Yığını

### 2.1 Teknoloji Yığını

| Katman | Teknoloji | Amaç |
|---|---|---|
| **Sunum** | Streamlit 1.x | Web UI, veri tablosu, filtreler |
| **İş Mantığı & Orkestrasyon** | Python 3.14 | Çıkarma, dönüşüm, agregasyon, facade mimarisi |
| **Modüler Çekirdek (Core)** | Python modülleri (`core/`) | Sayı temizleme, sütun çözümleme, reaktif hesaplama, normalizasyon, ENO, Excel |
| **Veri Okuma** | openpyxl, xlrd, lxml, BeautifulSoup | Çoklu format desteği (xlsx, xls, html, xml) |
| **Veri İşleme** | pandas, collections | Agregasyon, normalizasyon |
| **Çıktı** | openpyxl (write) | Formüllü Excel çıktısı |
| **Dinamik Yapılandırma** | SKF Başlıkları.xlsx + JSON | Kodsuz kural yönetimi, `config/header_overrides.json` |
| **Test & Kalite Güvencesi** | pytest 9.x | 61 birim ve golden pipeline regresyon testi |
| **Dağıtım & Çalışma** | Streamlit Cloud + Yerel (LAN) | 3 güvenlik modlu başlatıcı (`run.py`) |
| **Sürüm Kontrolü** | Git (main + dev) | Çift dallı sürüm ve değişiklik yönetimi |

### 2.2 Genel Mimari Diyagramı

```
┌──────────────────────────────────────────────────────────────────┐
│                         app.py (UI Katmanı)                     │
│  ┌─────────────┐  ┌──────────────┐  ┌────────────────────────┐  │
│  │ render_     │  │ run_analysis │  │ build_processing_logs  │  │
│  │ sidebar()   │→ │ _pipeline()  │→ │ audit_extracted_data   │  │
│  └─────────────┘  └──────┬───────┘  └────────────────────────┘  │
│                          │                                       │
├──────────────────────────┼───────────────────────────────────────┤
│            extract_and_compare.py (Facade & Orkestratör)         │
│                          │                                       │
│  ┌───────────────────────▼──────────────────────────────────┐   │
│  │              process_all_regions()                        │   │
│  │  ┌──────────────┐  ┌──────────────┐  ┌───────────────┐  │   │
│  │  │ discover_    │→ │ read_excel_  │→ │ aggregate_    │  │   │
│  │  │ supported_   │  │ content()    │  │ etso_data()   │  │   │
│  │  │ files()      │  │ read_html_   │  └───────────────┘  │   │
│  │  └──────────────┘  │ content()    │                      │   │
│  │                     │ read_xml_    │                      │   │
│  │                     │ content()    │                      │   │
│  │                     └──────────────┘                      │   │
│  └──────────────────────────────────────────────────────────┘   │
│                          │                                       │
├──────────────────────────┼───────────────────────────────────────┤
│                       core/ Katmanı                              │
│  ┌────────────────────┐ ┌──────────────────┐ ┌────────────────┐  │
│  │ column_resolver.py │ │ normalizer.py    │ │ number_cleaner │  │
│  │ (resolve_indices,  │ │ (region name,    │ │ (clean_turkish_│  │
│  │  extract_value)    │ │  etso_kodu)      │ │  number, guc)  │  │
│  └────────────────────┘ └──────────────────┘ └────────────────┘  │
│  ┌────────────────────┐ ┌──────────────────┐ ┌────────────────┐  │
│  │reactive_calculator │ │ eno_enricher.py  │ │ excel_writer.py│  │
│  │(calculate_reactive_│ │ (enrich_from_    │ │ (save_to_excel,│  │
│  │ totals — DRY)      │ │  eno_file)       │ │  styling)      │  │
│  └────────────────────┘ └──────────────────┘ └────────────────┘  │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │ readers/ (xlsx, xls, html, xml — Gelecek İzolasyon Fazı) │   │
│  └──────────────────────────────────────────────────────────┘   │
├──────────────────────────────────────────────────────────────────┤
│                      config/ Katmanı                             │
│  ┌──────────────────┐ ┌──────────────────┐ ┌─────────────────┐  │
│  │ constants.py     │ │ mappings.py      │ │header_overrides │  │
│  │ (STANDARD_COLS,  │ │ (REGION_MAPPING, │ │.json (kodsuz    │  │
│  │  SUM/MAX fields) │ │  SYNONYMS loader)│ │ override config)│  │
│  └──────────────────┘ └──────────────────┘ └─────────────────┘  │
├──────────────────────────────────────────────────────────────────┤
│           new_mapping_parser.py (Yapılandırma Ayrıştırıcı)       │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │  parse_mapping_file()                                     │   │
│  │  ├─ Başlıklar sayfası → regions{} + alt_regions{}         │   │
│  │  ├─ Düzeltme sayfası → corrections{} + distribution_names │   │
│  │  └─ Notlar sayfası → rules{}                              │   │
│  └──────────────────────────────────────────────────────────┘   │
├──────────────────────────────────────────────────────────────────┤
│                    Veri Kaynakları                                │
│  ┌────────────────┐  ┌──────────────────┐  ┌─────────────────┐  │
│  │ SKF Başlıkları │  │ 21 EDAŞ Klasörü  │  │ XX ENO Ay.xlsx  │  │
│  │ .xlsx (Config) │  │ (Ham Faturalar)  │  │ (ETSO Lookup)   │  │
│  └────────────────┘  └──────────────────┘  └─────────────────┘  │
└──────────────────────────────────────────────────────────────────┘
```

---

## 3. Modüler İzolasyon Stratejisi (ICM Analizi)

### 3.1 Mevcut Modül Haritası

Monolitik çekirdek kısmi olarak ayrıştırılmış (partially decomposed) ve modüler bir mimariye kavuşturulmuştur. `extract_and_compare.py` artık bir **FACADE** olarak görev yapmaktadır.

| Modül | Satır | Sorumluluk | Bağımlılıklar |
|---|---|---|---|
| `app.py` | ~1050 | Streamlit UI, sidebar yönetimi, pipeline tetikleme | `extract_and_compare` |
| `extract_and_compare.py` | ~2670 | **FACADE & Orkestratör**: Format okuyucuları, agregasyon, pipeline koordinasyonu (2740'tan indirildi) | `core.*`, `config.*`, `new_mapping_parser` |
| `new_mapping_parser.py` | ~515 | `SKF Başlıkları.xlsx` okuma, dinamik mapping ve düzeltme kuralları üretimi | `openpyxl`, `pandas` |
| `config/constants.py` | ~38 | Sabitler: `BASE_PATH`, `MAPPING_FILE`, `STANDARD_COLUMNS`, `SUM_FIELDS`, `MAX_FIELDS`, `TEXT_FIELDS_AGG` | Bağımsız (standalone) |
| `config/mappings.py` | ~180 | Bölge eşleştirmeleri (`REGION_NAME_MAPPING`), referans çevirileri, JSON yükleyici | `json`, `os` |
| `config/header_overrides.json` | ~152 | `SPECIAL_HEADER_MAPPING` ve `FIELD_SYNONYMS` verilerini kod değişikliği gerektirmeden yapılandırma | JSON yapılandırma dosyası |
| `core/normalizer.py` | ~75 | Bölge adı ve ETSO normalizasyonu (`normalize_region_name`, `normalize_etso_kodu`) | `config.mappings` |
| `core/number_cleaner.py` | ~100 | Sayı temizleme (`clean_turkish_number`), başlık sadeleştirme (`normalize_header`), güç formatlama (`_format_guc_kw`) | Bağımsız (standalone) |
| `core/column_resolver.py` | ~196 | Dinamik sütun eşleştirme (`resolve_column_indices`), grup çözümleme (`resolve_reactive_field_groups`), değer çıkarma (`extract_value_from_row`) | `core.number_cleaner`, `config.mappings`, `core.normalizer` |
| `core/reactive_calculator.py` | ~70 | Reaktif ve İlk Reaktif bedel hesaplama (`calculate_reactive_totals`, DRY: 4 kopyadan 1'e indirildi) | `core.number_cleaner` |
| `core/eno_enricher.py` | ~150 | ENO dosyasından Sayaç ID → EIC ve Abone bilgisi zenginleştirme (`enrich_from_eno_file`) | Bağımsız (`openpyxl`, `os`, `re`) |
| `core/excel_writer.py` | ~150 | Formüllü ve muhasebe biçimlendirmeli Excel çıktısı üretme (`save_to_excel`) | `config.constants`, `core.number_cleaner`, `openpyxl` |
| `core/readers/__init__.py` | ~2 | Format okuyucuları paketi (gelecek okuyucu ayrıştırması için yer tutucu) | Bağımsız |

### 3.2 Modül Bağımlılıkları ve Import Zinciri

Modüller arasındaki bağımlılık zinciri döngüsüz (DAG) ve katı hiyerarşik prensiplerle kurgulanmıştır:

```
config.constants           (Standalone - hiçbir iç bağımlılığı yok)
       │
config.mappings ──────────► json (config/header_overrides.json okur)
       │
core.normalizer ──────────► config.mappings
       │
core.number_cleaner        (Standalone - re harici bağımlılığı yok)
       │
core.column_resolver ─────► core.number_cleaner + config.mappings + core.normalizer
       │
core.reactive_calculator ─► core.number_cleaner
       │
core.eno_enricher          (Standalone - openpyxl, os, re)
       │
core.excel_writer ────────► config.constants + core.number_cleaner
       │
extract_and_compare (FACADE) ──► config.* + core.* + new_mapping_parser
```

### 3.3 ICM Değerlendirmesi: Modular Facade (Progressive Isolation)

> **Mevcut Durum: Modular Facade (Progressive Isolation)**  
> *(Eski Durum: Monolithic Core — Aşama 1 Başarıyla Tamamlandı)*

Önceki 2740+ satırlık monolitik yapı yerine sistem artık **kademeli olarak izole edilmiş bir modüler mimari** ile çalışmaktadır:
1. **Facade Prensibi:** `extract_and_compare.py` dosyası ~2670 satıra indirilmiş olup, `app.py` ve diğer tüketici katmanlar için geriye dönük tam uyumlu bir cephe (facade) görevi görür.
2. **DRY İhlallerinin Giderilmesi:** Reaktif hesaplama mantığı daha önce 4 ayrı format okuyucusunda birebir tekrarlanırken, `core/reactive_calculator.py` içerisindeki tek fonksiyona (`calculate_reactive_totals`) konsolide edilmiştir.
3. **Kodsuz Başlık Yönetimi:** Başlık eş anlamlıları ve bölgeye özel istisnalar `config/header_overrides.json` dosyasına alınarak, Python kodu değiştirmeden yeni fatura şablonlarına adapte olma kabiliyeti kazanılmıştır.
4. **Sıradaki Adım (Aşama 2):** `extract_and_compare.py` içerisinde kalan 4 format okuyucunun (xlsx, xls, html, xml) `core/readers/` dizini altına bağımsız modüller olarak taşınması planlanmaktadır.

### 3.4 Dizin Yapısı

```
06-Haziran/
├── config/
│   ├── __init__.py
│   ├── constants.py              ← Sabitler, standart sütunlar, SUM/MAX listeleri
│   ├── mappings.py               ← Bölge isimleri, referans sözlükleri, JSON yükleyici
│   └── header_overrides.json     ← SPECIAL_HEADER_MAPPING ve FIELD_SYNONYMS (JSON)
├── core/
│   ├── __init__.py
│   ├── column_resolver.py        ← Sütun indeksi çözümleme ve satır çıkarma
│   ├── eno_enricher.py           ← ENO dosyasından ETSO/Abone zenginleştirme
│   ├── excel_writer.py           ← Formüllü Excel yazıcı ve stil şablonları
│   ├── normalizer.py             ← Bölge adı ve ETSO normalizasyon fonksiyonları
│   ├── number_cleaner.py         ← Türkçe/US sayı ayrıştırma ve başlık temizleme
│   ├── reactive_calculator.py    ← Reaktif/İlk Reaktif toplam hesaplayıcı (DRY)
│   └── readers/
│       └── __init__.py           ← Format okuyucuları için yer tutucu (Faz 2)
├── tests/
│   ├── conftest.py               ← Test fikstürleri ve ortak yollar
│   ├── test_clean_number.py      ← Sayı temizleme birim testleri (24 test)
│   ├── test_normalizer.py        ← Normalizasyon birim testleri (17 test)
│   ├── test_golden_core.py       ← 2593 kayıt tam regresyon testleri (8 test)
│   └── test_reactive_calc.py     ← Reaktif hesaplama testleri (6 test)
├── extract_and_compare.py        ← Facade & Pipeline Orkestratörü (~2670 satır)
├── new_mapping_parser.py         ← SKF Başlıkları.xlsx dinamik ayrıştırıcı
├── app.py                        ← Streamlit Web Arayüzü (~1050 satır)
├── run.py                        ← 3 Güvenlik modlu port-otomasyonlu başlatıcı
├── Uygulamayi_Baslat.bat         ← Tek tıkla yerel çalıştırma betiği
└── DOCS/
    ├── software_design_spec.md   ← Bu tasarım spesifikasyonu
    ├── engineering_health_audit.md
    ├── security_report.md
    └── agents_spec.md
```

> **Önemli Not:** Yapılan tüm modülerleştirme ve yeniden yapılandırma çalışmaları kanonik çıktıyı (2593 kayıt, 21 bölge, birebir mali hash) mutlak surette korumuştur.

---

## 4. Veri Akışı (Data Flow)

### 4.1 End-to-End Pipeline

```
1. KULLANICI
   │
   ▼
2. app.py → render_sidebar()
   │  Kullanıcı veri klasörünü seçer (tkinter.filedialog veya web)
   │  Mapping dosyası: BASE_DIR / "SKF Başlıkları.xlsx" (gömülü)
   │
   ▼
3. run_analysis_pipeline(root_folder, mapping_path)
   │
   ├─ 3a. parse_mapping_file(mapping_path) [new_mapping_parser.py]
   │       → Başlıklar sayfası: 21+1 bölge mapping'i (Dicle çift format)
   │       → Düzeltme sayfası: Tarife/AG OG/Terim düzeltmeleri + dağıtım adı varyasyonları
   │       → Notlar sayfası: Bölge-özel kurallar (rules{})
   │
   ├─ 3b. process_all_regions(base_path, mapping) [extract_and_compare.py Facade]
   │       │
   │       ├─ Dinamik klasör eşleştirme (merged_folder_lookup)
   │       ├─ discover_supported_files(region_path) — xlsx/xls/html/xml
   │       ├─ detect_file_format() → format-specific reader seçimi
   │       ├─ resolve_column_indices() [core/column_resolver.py]
   │       │   → Çift format algılama (Dicle: skor bazlı)
   │       │   → Standart dışı format koruması (Notlar 12 #3)
   │       ├─ clean_turkish_number() [core/number_cleaner.py]
   │       │   → Satır-satır veri çıkarma (16+ alan)
   │       ├─ calculate_reactive_totals() [core/reactive_calculator.py]
   │       │   → Reaktif/İlk Reaktif hesaplama (DRY)
   │       ├─ aggregate_etso_data() — ETSO bazlı birleştirme (SUM/MAX/TEXT)
   │       └─ apply_note_rules() — Düzeltme tablosu uygulaması
   │
   ├─ 3c. normalize_extracted_regions(records) [core/normalizer.py]
   │       → Bölge isimlerini referans formatına çevir
   │
   ├─ 3d. enrich_from_eno_file(records, root_folder) [core/eno_enricher.py]
   │       → Sayaç ID → EIC Kod dönüşümü (40Z formatı)
   │       → Boş Müşteri → Abone Ad-Soyad doldurma
   │
   └─ 3e. save_to_excel(records, output_buffer) [core/excel_writer.py]
           → Formüllü Excel: KDV Matrahı, KDV, Toplam
           → Başlık biçimlendirme, sayı formatları, otomatik filtre
```

### 4.2 Sayı Temizleme Zinciri

```
Kaynak Değer          → clean_turkish_number()      → Çıktı
──────────────────────────────────────────────────────────────
"1.234.567,89"        → Türkçe format algılama      → 1234567.89
"1,234,567.89"        → US format algılama           → 1234567.89
"(1.234,56)"          → Parantezli negatif           → -1234.56
"-1234.56"            → Direkt negatif               → -1234.56
"1234.56-"            → Sondan eksi (Yeşilırmak)    → -1234.56
"X"                   → Sakarya İlk Reaktif          → "X" (olduğu gibi)
None / "" / "Boş"     → Boş değer                   → 0
```

---

## 5. Bileşen (Component) Dağılımı

### 5.1 Çekirdek Fonksiyonlar ve Modül Konumları

| Fonksiyon | Modül | Satır | Görev | Kritiklik |
|---|---|---|---|---|
| `process_all_regions()` | `extract_and_compare.py` | ~200 | Orkestratör: tüm bölgeleri dolaş, dosyaları oku, birleştir | 🔴 Kritik |
| `resolve_column_indices()` | `core/column_resolver.py` | ~196 | Başlık→sütun indeksi eşleştirme (3 katmanlı fallback, synonyms) | 🔴 Kritik |
| `calculate_reactive_totals()` | `core/reactive_calculator.py` | ~70 | Reaktif ve İlk Reaktif toplama (DRY tek kaynak) | 🔴 Kritik |
| `clean_turkish_number()` | `core/number_cleaner.py` | ~100 | Çoklu sayı formatı ayrıştırma (Türkçe/US/negatifler) | 🔴 Kritik |
| `normalize_header()` | `core/number_cleaner.py` | ~30 | Türkçe karakter ve boşluk arındırılmış başlık temizliği | 🟡 Önemli |
| `normalize_region_name()` | `core/normalizer.py` | ~40 | Extraction <-> Reference çift yönlü bölge ismi çevirisi | 🟡 Önemli |
| `normalize_etso_kodu()` | `core/normalizer.py` | ~35 | Çarpışmasız ETSO kodu standartlaştırması | 🟡 Önemli |
| `enrich_from_eno_file()` | `core/eno_enricher.py` | ~150 | ENO dosyasından ETSO+Müşteri zenginleştirme | 🟡 Önemli |
| `save_to_excel()` | `core/excel_writer.py` | ~150 | Formüllü ve stilli Excel çıktısı üretme | 🟡 Önemli |
| `read_excel_content()` | `extract_and_compare.py` | ~180 | xlsx/xls dosyalarından veri çıkarma + çift format | 🔴 Kritik |
| `aggregate_etso_data()` | `extract_and_compare.py` | ~100 | ETSO bazlı SUM/MAX/TEXT birleştirme | 🔴 Kritik |
| `parse_mapping_file()` | `new_mapping_parser.py` | ~300 | SKF Başlıkları.xlsx → mapping dict üretimi | 🔴 Kritik |
| `apply_note_rules()` | `new_mapping_parser.py` | ~60 | Düzeltme tablosu kurallarının uygulanması | 🟡 Önemli |

### 5.2 Yapılandırma ve Veri Sözlükleri

| Sözlük / Yapı | Bulunduğu Yer | Amaç |
|---|---|---|
| `REGION_NAME_MAPPING` | `config/mappings.py` | Extraction → Reference bölge ismi eşleştirme tablosu |
| `REFERENCE_TO_EXTRACTION_MAPPING` | `config/mappings.py` | Reference → Extraction ters çeviri tablosu |
| `REGION_NORMALIZATION_MAPPING` | `config/mappings.py` | Klasör adı normalizasyon eşleşmeleri |
| `SPECIAL_HEADER_MAPPING` | `config/header_overrides.json` + `config/mappings.py` | Bölgeye özel başlık override'ları (kodsuz JSON) |
| `FIELD_SYNONYMS` | `config/header_overrides.json` + `config/mappings.py` | Sütun eş anlamlı kelime havuzu (kodsuz JSON) |
| `STANDARD_COLUMNS` | `config/constants.py` | Çıktı Excel standart 18 sütun sırası |
| `SUM_FIELDS` / `MAX_FIELDS` | `config/constants.py` | Agregasyon sırasında toplanacak ve max alınacak alanlar |

---

## 6. Çalışma ve Güvenlik Modları (run.py & Dağıtım)

Uygulama, hem yerel masaüstü ortamında hem de kurumsal intranet (LAN) senaryolarında güvenli ve çakışmasız çalışabilmesi için `run.py` başlatıcısı ile donatılmıştır.

### 6.1 Güvenlik Modları (`run.py` Başlatıcı)

`run.py` komut satırı arayüzü 3 farklı güvenlik modu sunar:

| Parametre | Dinleme Adresi | XSRF Koruması | Kullanım Senaryosu |
|---|---|---|---|
| `--local` *(Varsayılan)* | `127.0.0.1` | Açık | **En güvenli mod.** Yalnızca yerel makineden erişilebilir; ağa kapalıdır. |
| `--lan` | `0.0.0.0` | **Açık** | **Güvenli LAN modu.** Kurumsal ağdaki diğer kullanıcıların erişimine açar; XSRF aktiftir. |
| `--lan-open` | `0.0.0.0` | **Kapalı** | **Geliştirme modu.** Ağ üzerinden testlerde tarayıcı güvenlik kısıtlamalarını esnetir. |

#### Dinamik Port Algılama (`find_free_port`)
`run.py`, makinede çalışan diğer yerel projelerle çakışmayı önlemek için varsayılan olarak `8505` portundan başlayarak ilk boş TCP portunu otomatik olarak bulur (`8505`–`8600` aralığı). Sabit bir port istenirse `--port <NO>` parametresi kullanılabilir.

### 6.2 Yerel Mod (Masaüstü İş Akışı - `Uygulamayi_Baslat.bat`)

```
Kullanıcı → Uygulamayi_Baslat.bat → run.py (Varsayılan: --local) → streamlit run app.py
          → Port otomatik bulunur (ör: 8505)
          → tkinter.filedialog.askdirectory() ile yerel klasör seçimi
          → İşlem tamamen yerel makinede, sıfır ağ riski
```

### 6.3 LAN Paylaşım Modu

```
Host (run.py --lan) → 0.0.0.0:8505 dinler (XSRF aktif)
LAN kullanıcıları   → http://<host-ip>:8505
                    → Sadece sonuçları inceler / rapor indirir (read-only)
                    → Klasör seçimi host makinede yürütülür
```

### 6.4 Streamlit Cloud Modu

```
GitHub push (main dalı) → Otomatik deploy tetiklenir
                        → https://ck-enerji-teklif-uretici-ve-karsilastirma.streamlit.app/
                        → Dağıtım ve canlı sunum
```

---

## 7. Test ve Regresyon Altyapısı (pytest)

Sistemin modülerleştirilmesi ve refactoring süreçlerinin sıfır regresyonla yürütülebilmesi amacıyla **pytest tabanlı kapsamlı bir test süiti** inşa edilmiştir.

### 7.1 Test Süiti Yapısı (61 Test / 4 Dosya)

```
tests/
├── test_clean_number.py      → 24 test
├── test_normalizer.py        → 17 test
├── test_golden_core.py       → 8 test (Tam pipeline regresyonu)
└── test_reactive_calc.py     → 6 test
```

| Test Modülü | Test Sayısı | Kapsanan İşlevler |
|---|---|---|
| `tests/test_clean_number.py` | 24 | Türkçe/US sayı ayrıştırma, parantezli negatifler, sondan eksi (Yeşilırmak), Sakarya 'X' değeri, boş/None toleransı, başlık normalizasyonu ve güç kW formatlama |
| `tests/test_normalizer.py` | 17 | `normalize_region_name` çift yönlü çevirileri (Ref<->Ext), Türkçe İ/ı case-insensitivity doğrulaması, `normalize_etso_kodu` ile çarpışmasız EIC/sayaç ID standartlaştırması |
| `tests/test_golden_core.py` | 8 | **Golden Core Regresyon Kilidi:** 21 EDAŞ bölgesinin eksiksiz işlenmesi, toplam 2.593 kayıt çıkması, sıfır boş ETSO garantisi, Sakarya aktif enerji kontrolü ve standart kolon şeması doğrulaması |
| `tests/test_reactive_calc.py` | 6 | DRY reaktif hesaplayıcısı: Reaktif/İlk Reaktif ayrımı, Sakarya 'X' bayrağı tespiti, negatif değerler ve geçersiz indeks koruması |

### 7.2 Test Çalıştırma Komutu

Tüm test süiti aşağıdaki komutla çalıştırılır:

```bash
python -m pytest tests/ -v
```

> **Regresyon Prensibi:** `test_golden_core.py` testleri çalıştırıldığında gerçek 21 EDAŞ ham verisi üzerinde tam pipeline yürütülür ve 2.593 kayıt sayısı doğrulanmadan hiçbir kod değişikliği kabul edilmez.

---

## 8. Sürüm Kontrolü ve Git Stratejisi

Proje, kurumsal yazılım standartlarına uygun çift dallı (`main` + `dev`) bir Git iş akışı ile yönetilmektedir.

```
[dev]  ──► [Modüler Refactoring / Özellikler] ──► [pytest (61/61 Yeşil)] ──┐
                                                                           │ (PR / Merge)
[main] ───────────────────────────────────────────────────────────────────►┴──► [Canlı / Streamlit Cloud]
```

### 8.1 Dallar ve Sorumluluklar

- **`main` Dalı (Canlı / Production):**  
  Doğrulanmış, stabil ve Streamlit Cloud üzerinde canlı çalışan sürümdür. Bu dala doğrudan commit atılmaz; yalnızca `dev` dalından kapsamlı testleri geçmiş kodlar birleştirilir.
- **`dev` Dalı (Geliştirme / Active Development):**  
  Modüler izolasyon çalışmaları, yeni EDAŞ format adaptasyonları ve kod iyileştirmelerinin yürütüldüğü ana çalışma dalıdır.

### 8.2 Birleştirme (Merge) Ön Koşulları

`dev` dalından `main` dalına geçiş yapılabilmesi için aşağıdaki koşulların eksiksiz sağlanması zorunludur:
1. Tüm pytest süitinin (`61/61`) hatasız geçmesi (`pytest tests/ -v`).
2. Kanonik Golden Core çıktısının (**2.593 kayıt, 21 bölge**) bozulmadığının doğrulanması.
3. Mali alan toplamlarının (Dağıtım Bedeli, Tazminat, Standart Dışı) kanonik hash değerleriyle örtüşmesi.

---

## 9. Evrim Tarihi (FAZ Kronolojisi)

| FAZ | Tarih | Kapsam |
|---|---|---|
| FAZ 1-9 | Ağustos 2026 | Bölge standardizasyonu, reaktif hesaplama, düzeltme tablosu, dark mode |
| FAZ 10 | 28 Ağu 2026 | Gömülü mapping, dinamik klasör normalizasyonu |
| Notlar 11 | 16 Eyl 2026 | KDV Matrahı sütunu, Dicle çift format, ENO zenginleştirme |
| Notlar 12 | 22 Eyl 2026 | Sakarya aktif enerji, dinamik eşleştirme, standart dışı koruma |
| **FAZ 13 (v2.0)** | **24 Eylül 2026** | **Modüler İzolasyon (Progressive Facade):** `extract_and_compare.py` facade yapısına dönüştürüldü. 8 yeni modül (`config/`, `core/`), JSON tabanlı kodsuz başlık yapılandırması (`header_overrides.json`), DRY reaktif hesaplayıcı, 61 testli `pytest` süiti, 3 güvenlik modlu `run.py` ve `main`/`dev` Git stratejisi entegre edildi. |

---

*Bu doküman, SKF Fatura Birleştirme Sistemi'nin 24 Eylül 2026 tarihi itibarıyla güncel modüler mimari spesifikasyonudur.*
