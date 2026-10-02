#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
CK ENERJİ FATURA VERİ ÇIKARMA MOTORU - ÜRETİM TEKNİK SÖZLEŞMESİ
================================================================

AMAÇ VE AKTİF MİMARİ
--------------------
Bu modül, mapping dosyasındaki dağıtım bölgelerini dinamik olarak okur; her
bölgenin klasöründeki XLS, XLSX, XML ve HTML kaynaklarını işler; aynı ETSO ve
bölgeye ait kayıtları birleştirir; bölgesel iş kurallarını ve ``Düzeltme``
eşleştirmelerini uygular; standart Excel çıktısını üretir.

Streamlit üretim zinciri doğrudan bu modülü kullanır::

    app.py -> extract_and_compare.py -> new_mapping_parser.py

``core/``, ``config/`` ve ``analytics/`` altında tarihsel/modüler kopyalar
bulunabilir; veri çıkarma için doğrulanmış kaynak bu dosyadır. Çalışan çekirdek
mantık tarihsel olarak doğrulanmış "golden" davranıştır. Aşağıdaki alan adları,
işlem sırası, formüller ve bölgesel istisnalar regresyon testi yapılmadan
değiştirilmemelidir.

VARSAYILAN DOSYALAR VE TAŞINABİLİRLİK
-------------------------------------
* ``BASE_PATH`` bu dosyanın bulunduğu klasördür; sabit kullanıcı/Masaüstü yolu
  kullanılmaz.
* Varsayılan mapping ``SKF Başlıkları.xlsx`` dosyasıdır.
* Karşılaştırma referansı ``Dağıtımın Kestiği Faturalar Özet.xlsx`` dosyasıdır.
* Varsayılan çıktı ``Çıkarılan_Veriler.xlsx`` dosyasıdır.
* ``process_all_regions`` çağrısındaki ``base_path`` ve
  ``mapping_file_path`` parametreleri Streamlit'in dinamik klasör seçimini
  sağlar. Ana klasör veya mapping yoksa ``FileNotFoundError`` üretilir.

MAPPING DOSYASI SÖZLEŞMESİ
--------------------------
Mapping ``new_mapping_parser.parse_mapping_file`` ile okunur.

1. ``Başlıklar`` zorunludur. Bölge listesi hardcoded değildir; A sütunundaki
   dolu bölge satırlarından, dosya sırası korunarak oluşur. Üretim referansında
   21 EDAŞ bölgesi vardır.
2. Not sayfası önceliği: Notlar 7, 6, 5, 4, 3, 2, 1, Notlar.
3. ``Düzeltme`` varsa üç alan lookup'u ve varsa I:L Sayax tarife tablosu okunur.
4. Ana 16 kaynak alanı: ETSO, Müşteri, Tarife, AG/OG, Terim, Güç, Kurulu Güç,
   Aktif Enerji, Trafo Kaybı, Dağıtım Bedeli, Güç Bedeli, Güç Aşım Bedeli,
   iki reaktif ve iki reaktif-tenzil alanıdır.
5. Mapping 7 bunlara üç fiziksel tazminat ve dokuz fiziksel standart-dışı
   başlık/sütun çifti ekler. Güncel referansın mantıksal başlıkları
   ``Tazminat Bedeli-1..3`` ve ``Standart Dışı-1..9`` şeklindedir.
6. Boş, NaN ve ``-`` mapping hücreleri tanımsızdır. Excel sütun harfleri
   sıfır tabanlı indekse çevrilir: A=0, Z=25, AA=26; geçersiz değer -1'dir.
7. Şema özellikleri ``notlar_7``, ``corrections`` ve
   ``extended_financial_fields`` bayraklarıyla taşınır.
8. Mapping 6 gibi eski bir şema genişletilmiş alan taşımıyorsa eski kayıt
   anahtarları ve 18 sütunlu Excel aynen korunur.
9. ``load_mapping`` tarihsel eski fonksiyondur; üretim akışı
   ``load_mapping_new`` üzerinden parserı kullanır.

DOSYA KEŞFİ, SEÇİMİ VE FORMAT TESPİTİ
-------------------------------------
* Her bölge klasörü ``os.walk`` ile bütün alt klasörleri dahil taranır. Bu,
  özellikle Vangölü EDAŞ alt klasörleri için zorunludur.
* Desteklenen uzantılar case-insensitive ``.xls/.xlsx/.xml/.html``; ``~$``
  geçici Excel dosyaları atlanır. Dizin ve dosya sırası casefold ile
  deterministiktir.
* AKEDAŞ'ta yalnız adında birebir ``TL_Raporu`` geçen dosyalar işlenir.
* Uludağ'da adında ``4008`` ve ``TL`` bulunmalı; büyük harfe çevrilmiş adda
  ``KWH`` bulunmamalıdır.
* Uzantıya tek başına güvenilmez: XML, HTML, ZIP/XLSX ve OLE2/XLS imzaları ilk
  baytlardan belirlenir. ``.xls`` uzantılı gerçek XML/HTML dosyaları ilgili
  yedek okuyucuya yönlendirilir.
* XLSX'te bütün sayfalar ve ilk satır başlığı; XLS'te bütün sayfalar ve gerekirse
  ikinci satır başlığı; XML/HTML'de banner sonrası ilk yeterince dolu başlık
  satırı işlenir. HTML'de en çok satırlı tablo seçilir.
* XLSX okuyucusundaki ``FATURA_TURU`` ve ``FATURA_NO`` yardımcı alanları final
  Excel'e yazılmaz ve mevcut üretimde eleme anahtarı değildir.

DİNAMİK BAŞLIK ÇÖZÜMLEME ÖNCELİĞİ
---------------------------------
``resolve_column_indices`` her gerçek kaynak başlığında şu sırayı kullanır:

1. Mapping 7 Çamlıbel zorunlu sabit indeks kuralı.
2. ``SPECIAL_HEADER_MAPPING`` bölgesel karşılığı.
3. Mappingdeki gerçek kaynak başlığının normalize edilmiş eşleşmesi.
4. Mappingdeki sabit sütun indeksi.
5. ``FIELD_SYNONYMS`` eş anlamlı listesi.
6. Bulunamazsa -1.

Başlık normalizasyonu Türkçe karakterleri ASCII benzerlerine çevirir; küçük
harf, boşluk/alt çizgi ve özel karakter temizliği uygular. Önemli istisnalar:
AKEDAŞ ETSO ``Id`` ve müşteri kapalı; Yeşilırmak'ın ETSO/müşteri/güç/reaktif
başlıkları; Toroslar AG/OG ``Gerilim Seviyesi``; Meram ve Sakarya reaktif
başlıkları.

Sütun/başlık denetimi veri çıkarma sırasını değiştirmez. Çamlıbel, Sakarya ve
Kayseri için kabul edilen Haziran 2026 başlık düzeni
``config/source_header_layouts.json`` içinde tutulur; mevcut düzen korunuyorsa
uyarı üretilmez, başlık satırı değişirse kaynak formatı uyarısı verilir.

ETSO, BÖLGE VE SAYI NORMALİZASYONU
---------------------------------
* ETSO: None/boş/null/``-`` -> None; ``123.0`` -> ``123``; EIC biçiminde 00Z
  sonrası anlamlı sayısal kök alınır; diğer biçimlerde rakamlar birleştirilip
  baştaki sıfırlar kaldırılır. Agregasyon anahtarı
  ``(normalize ETSO, ham dağıtım bölgesi)`` çiftidir; aynı ETSO farklı
  bölgelerde birleşmez.
* Bölge: çift yönlü casefold tabanlı extraction/reference eşleştirmesi vardır.
  Streamlit/terminal Excel'e yazmadan önce referans adlarını kullanır.
* ``clean_turkish_number`` Türkçe ``1.234.567,89``, US
  ``1,234,567.89``, parantezli negatif, baştan eksi ve Yeşilırmak'ta görülen
  sondan eksi biçimlerini okur. Boş/çözülemeyen değer 0 olur. Kaynak değere
  hiçbir 1.000 çarpma veya bölme ölçeği uygulanmaz.

AKTİF ENERJİ KURALLARI
----------------------
* Boğaziçi, Uludağ ve Yeşilırmak'ta, trafo sütunu çözülebiliyorsa
  aktif enerji ``aktif_enerji + trafo_kaybı``; aksi halde yalnız aktiftir.
* Notlar 12 #1: Sakarya EDAŞ aktif enerji her zaman ``Dağıtım Miktarı``
  (Sütun AR) başlığından gelir. Fallback kaldırıldı.
* Agregasyon sonrası Yeşilırmak aktif toplamı yaklaşık sıfırsa ve grupta
  pozitif aktifler varsa yalnız pozitiflerin toplamı alınır.

REAKTİF VE İLK REAKTİF KURALLARI
--------------------------------
* Mapping 7'de ADM, Gediz, Trakya ve Uludağ için Reaktif Bedel yalnız
  ``reaktif + reaktif2``; tenzil alanları bu toplama girmez.
* Meram EDAŞ'ta Reaktif Bedel, tenzil/ihlal alanları eklenmeden hesaplanır;
  Reaktif Bedel İhlal yalnız İlk Reaktif alanına yazılır.
* Golden eski kural gereği Aras, Çoruh, Dicle, Fırat ve Vangölü de yalnız
  ``reaktif + reaktif2`` kullanır.
* Diğer bölgeler ``reaktif + reaktif2 + reaktif_tenzil +
  reaktif_tenzil2`` kullanır.
* İlk Reaktif Trakya'da yalnız ``reaktif_tenzil2``; diğerlerinde iki tenzil
  alanının toplamıdır.
* Kaynak reaktif alanındaki ``X`` sayısal toplama girmez; tekil kaydın İlk
  Reaktif değeri ``X`` olur. Mükerrer SUM sırasında X atlanır.

MAPPING 7 YENİ İŞ KURALLARI
---------------------------
1. **Çamlıbel AV:** Notlar 7 aktifken Tarife Grubu, AG OG ve Terim başlık adıyla
   ilk eşleşmeden değil, zorunlu AV / zero-based index 47 kaynağından okunur.
   Bu kural Mapping 6'ya taşmaz.
2. **Tazminat:** Bölge için tanımlı ``Tazminat Bedeli-1..3`` kaynaklarının
   tamamı ham satırda toplanır; ETSO+bölge agregasyonunda tekrar SUM edilir;
   nihai ``Tazminat Bedeli`` her zaman en son sütundur.
3. **Standart dışı:** Fiziksel ``Standart Dışı-1..9`` çiftlerinden bölgede
   tanımlı olanların tamamı ham satırda ve ardından ETSO+bölge grubunda
   toplanır. Nihai sözleşmedeki bilinçli yazım tam olarak
   ``Standart Dışı Tutar (TL)`` şeklindedir.
4. **Düzeltme:** A:B Tarife Grubu, C:D AG OG, E:F TERİM lookup bloklarıdır;
   I:L bloğu bu üç alanın birleşimini Sayax tarifesine çevirir.
   Eşleşme Unicode NFKC + boşluk sadeleştirme + casefold ile yapılır. Yalnız
   lookup'ta bulunan değer değişir; bulunmayan mevcut sonucu aynen korur.
   Düzeltme, ETSO agregasyonu ve bütün legacy bölgesel post-process
   işlemlerinden SONRA uygulanır. Tekrarlı kaynakta ilk hedef korunur;
   çelişkili hedef mapping uyarısı üretir.

BÖLGESEL GOLDEN POST-PROCESS
----------------------------
* Boğaziçi hariç Tarife sonundaki ``(Ek)/(ek)/ ek/(EK)`` kaldırılır.
* ADM/Gediz: TICARET/TİCARET -> ``TICARETHAN``; SANAYI/SANAYİ -> ``SANAYI``.
* Çamlıbel: Ticarethane/``12-`` -> ``TİCARETHANE``; Sanayi -> ``SANAYİ``;
  müşteri nihai kayıtta bilerek boşaltılır.
* AG/OG, Başkent/Toroslar/Gediz/ADM/Trakya dışındaki bölgelerde AG/DAG -> AG,
  OG/DOG -> OG olarak sadeleştirilir; sayılan beş bölgede ham değer korunur.
* Yeşilırmak TERİM boştur.
* Boğaziçi TERİM: Tarife Boş veya aktif+dağıtım yaklaşık sıfırsa boş; tek ->
  Tek Terimli; cift/çift -> Çift Terimli. Mükerrer metin seçiminde anlamlı
  tarife yoksa Tarife ``Boş``; tarife varken geçerli terim yoksa
  ``Tek Terimli`` kullanılır.
* Osmangazi: ``Çift/Tek Terim-Tek Zamanlı    `` değerlerindeki tarihsel son
  boşluklar golden çıktının parçasıdır.
* Çamlıbel Terim: Tek/12-/TekTerim -> Tek Terimli; Çift -> Çift Terimli.
* Diğer bölgeler Tek/Çift Terim -> Tek/Çift Terimli olur.
* XLSX Akdeniz'de ``Enerji_Acma_Kesme_Bedeli`` bulunursa Dağıtım Bedeline
  eklenir. Parantezli ve diğer negatif değerler bütün bölgelerde korunur.

SATIR FİLTRELEME VE EŞİKLER
---------------------------
1. ETSO veya müşteri hücresinde bilinen başlık metni taşıyan tekrarlı alt
   başlık satırları atılır.
2. Yalnız **pozitif** ``Dağıtım Bedeli(TL) > 10_000_000_000`` bozuk/absürt
   sayılır ve atılır. ``abs`` kullanılmaz.
3. Negatif dağıtım bedelleri kural gereği KORUNUR ve ``negative_count``
   sayacına girer; Akdeniz negatifleri de korunur.
4. Bedel 0 olsa bile ETSO veya müşteri doluysa kayıt korunur. Yalnız bedel 0,
   ETSO boş ve müşteri boş olan tamamen anlamsız satır atılır.
5. Yeşilırmak'ın 1 milyon TL üzeri/negatif değerleri ayrıca loglanır; yalnız
   log nedeniyle elenmez.

ÖNEMLİ: Çekirdekteki 10 **milyar** TL bozuk veri filtresi ile Streamlit
anomali panelindeki 10 **milyon** TL kesin gösterim eşiği farklıdır. Aktif UI
kuralı ``app.py`` içindedir; ``analytics/anomaly_detector.py`` yalnız pasif
tarihsel/modüler kopyadır. Bu iki eşik birbirine dönüştürülmemelidir.

ETSO AGREGASYON SÖZLEŞMESİ
--------------------------
Filtre sonrası sıralı gruplamada:

* SUM: Aktif Enerji, Dağıtım Bedeli, Güç Bedeli, Güç Aşım Bedeli, Reaktif,
  İlk Reaktif; Mapping 7'de ayrıca Standar Dışı ve Tazminat.
* MAX: Güç kW ve Kurulu Güç.
* TEXT: Müşteri, Tarife, AG OG, TERİM için ilk geçerli değer; Boğaziçi'nin
  yukarıdaki özel kuralı saklıdır.
* ``merged_duplicates`` her grupta kayıt sayısı eksi bir toplamıdır.
* Boş ETSO fakat dolu müşterili satırlar filtrede kalabildiğinden aynı
  bölgedeki boş-ETSO satırları tek grupta birleşebilir; bu golden davranıştır.

EXCEL ÇIKTI SÖZLEŞMESİ
----------------------
Standart 18 sütun: Dağıtım Bölgesi, ETSO Kodu, Müşteri, Tarife Grubu, AG OG,
Terim, Güç, Kurulu Güç, Aktif Enerji, Dağıtım Bedeli, Güç Bedeli, Güç Aşım,
Reaktif, KDV Matrahı, KDV, Toplam, İlk Reaktif, Sayax'a Atılacak Tarife.

Mapping 7 kaydı varsa S=``Standart Dışı Tutar (TL)`` ve son
T=``Tazminat Bedeli`` eklenir. Formüller her veri satırında aynen::

    N = J + K + L + M
    O = N * 0.2
    P = N + O + Q

İlk Reaktif Q sütununda, Sayax tarife R sütunundadır. S/T alanları KDV/Toplam
formüllerine dahil edilmez. Sayfa adı ``Çıkarılan Veriler``; ETSO biçimi ``0``;
sayısal alanlar mevcut muhasebe formatındadır. Kaydetmeden hemen önce her
sütun maksimum string uzunluğu + 2 genişliğe getirilir. PermissionError'da
``_Guncel.xlsx`` alternatifi yazılır.

Excel her zaman orijinal/maskesiz ETSO içerir. ETSO maskeleme yalnız Streamlit
görüntü kopyasında yapılır; ``save_to_excel`` öncesi ham kayıt değiştirilmez.

PROCESS_ALL_REGIONS VE STATS SÖZLEŞMESİ
---------------------------------------
Varsayılan dönüş ``list[dict]``; ``return_stats=True`` dönüşü
``(list[dict], stats)`` olur. Stats içinde en az şu alanlar korunur:

``expected_region_count``, ``read_region_count``, ``successful_regions``,
``missing_regions``, ``processed_regions``, ``total_extracted``,
``total_records``, ``zero_filtered``, ``negative_count``,
``akdeniz_filtered`` (tarihsel; normalde 0), ``absurd_filtered``,
``merged_duplicates``, ``discovered_files``, ``processed_files``,
``file_errors``, ``per_region_records``, ``notes_sheet``,
``mapping_features``, ``mapping_warnings``, ``source_header_warnings``,
``correction_duplicates``,
``correction_counts``, ``correction_total``, ``correction_examples``,
``standart_disi_total`` ve ``tazminat_total``.

Bir klasörün bulunması başarı sayılmaz; bölge yalnız filtre/agregasyon sonrası
nihai veride en az bir kaydı varsa ``read_region_count`` içine girer.
``missing_regions`` eksik adları taşır. ``status_callback`` mapping sonrası,
her bölge öncesi ve tamamlanınca ilerleme bildirir.

REFERANS KARŞILAŞTIRMA VE TERMINAL MAIN
---------------------------------------
``compare_with_reference`` referansın ``Dağıtımın Kestiği`` sayfasındaki A:M
alanlarını okur; anahtar normalize ETSO + iki ondalık Dağıtım Bedeli +
normalize Bölgedir. Müşteri asıl anahtara girmez; dolu/boş müşteri önceliği ve
duplicate adetleri korunur. Fonksiyon sonuçları konsola raporlar.

``main`` veriyi çıkarır, bölge adlarını referans biçimine çevirir,
``Çıkarılan_Veriler.xlsx`` ve ``Nihai_Birlestirilmis_Faturalar.xlsx``
dosyalarını yazar ve referans karşılaştırmasını çalıştırır.

REGRESYON KİLİTLERİ VE DEĞİŞİKLİK DİSİPLİNİ
-------------------------------------------
Kaynak veri değişmediği sürece beklenen golden ölçüler:

* Mapping 6: 2.593 kayıt, 21/21 bölge, Dağıtım Bedeli
  353.348.178,79 TL; genişletilmiş alan yok; kanonik SHA-256
  ``239a50d6ab0f5d9152d9b1fced7a9536788d1ccf1a384760dbe00bef5d21955e``.
* Güncel Mapping 7: 2.593 kayıt, 21/21, aynı dağıtım toplamı;
  Standar Dışı 140,00 TL; Tazminat 489,16 TL; düzeltme sayıları
  Tarife 1.779, AG OG 463, TERİM 2.211; kanonik SHA-256
  ``e6da5dbabc265e173b08d4a1ba90ff3a9546f8f6276f1533db5db5acea4fc9d3``.

Şunlar regresyonsuz değiştirilmemelidir: negatiflerin korunması; ETSO+bölge
grubu; SUM/MAX/TEXT ayrımı; Çamlıbel AV'nin yalnız Mapping 7'de olması; dört
bölgenin base-only reaktif kuralı; post-process'in Düzeltme'den önce olması;
18/20 kolon koşulu; N/O/P formülleri; ``Standar`` yazımı ve maskesiz Excel.

``apply_note_rules`` bazı tarihsel metadata bayrakları üretir; gerçek çalışma
davranışının otoritesi yalnız bayrak adları değil ``filter_extracted_data``,
``resolve_reactive_field_groups``, ``clean_turkish_number`` ve
``process_all_regions`` fonksiyon akışıdır. Çekirdek yalnız bayraklara bakarak
yeniden yazılmamalıdır.
"""

import os
import re
import unicodedata
import openpyxl
import xlrd
import sys
import pandas as pd
from xml.etree import ElementTree as ET
from html.parser import HTMLParser

# Add parent directory to path for importing new_mapping_parser
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from new_mapping_parser import (
    parse_mapping_file,
    apply_note_rules,
    excel_column_letter_to_index,
    index_to_excel_column,
    extract_file_filter,
)
from core.source_header_validator import (
    find_effective_header_row,
    is_known_source_header_layout,
    known_layout_change_warning,
)

# ── Modüler yapı importları ──────────────────────────────────────────────
# Aşağıdaki modüller, bu monolitik dosyadaki fonksiyonların bağımsız
# kopyalarını barındırır. Dış modüller (app.py gibi) bu dosyadan import
# etmeye devam edebilir (geriye dönük uyumluluk / facade pattern).
# Yeni kod ise doğrudan core.* ve config.* modüllerinden import etmelidir.
# ─────────────────────────────────────────────────────────────────────────

# Sabitler
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
MAPPING_FILE = os.path.join(BASE_PATH, 'SKF Başlıkları.xlsx')
REFERENCE_FILE = os.path.join(BASE_PATH, 'Dağıtımın Kestiği Faturalar Özet.xlsx')
OUTPUT_FILE = os.path.join(BASE_PATH, 'Çıkarılan_Veriler.xlsx')
SUPPORTED_INPUT_EXTENSIONS = ('.xls', '.xlsx', '.xml', '.html')

# Klasör isimleri ile mapping dosyası isimleri arasında eşleştirme
REGION_NAME_MAPPING = {
    "Akdeniz EDAŞ": "AKDENİZ EDAŞ",
    "Boğaziçi EDAŞ": "BEDAŞ",
    "Gediz EDAŞ": "GDZ EDAŞ",
    "Kayseri ve Civarı": "KCETAŞ",
    "Meram EDAŞ": "MERAM",
    "Osmangazi EDAŞ": "OEDAŞ",
    "Sakarya EDAŞ": "SEDAŞ",
    "Trakya EDAŞ": "TREDAŞ",
    "Uludağ EDAŞ": "UEDAŞ",
    "Çamlıbel EDAŞ": "ÇAMLIBEL EDAŞ",
    "Vangölü EDAŞ": "VANGÖLÜ EDAŞ",
    "Yeşilırmak EDAŞ": "YEDAŞ",
    "Başkent EDAŞ": "BAŞKENT EDAŞ",
    "Toroslar EDAŞ": "TOROSLAR EDAŞ",
    "Fırat EDAŞ": "FIRAT EDAŞ",
    "Dicle EDAŞ": "DİCLE EDAŞ",
    "Çoruh EDAŞ": "ÇORUH EDAŞ",
    "Aras EDAŞ": "ARAS EDAŞ",
    "AYEDAŞ": "AYEDAŞ",
    "ADM EDAŞ": "ADM EDAŞ",
    "AKEDAŞ": "AKEDAŞ"
}

# Referans dosyası ile extraction arasındaki region name mismatch'leri düzeltmek için
# Extraction'da: 'ADM EDAŞ', 'AKEDAŞ', 'AYEDAŞ', 'Dicle EDAŞ', 'Fırat EDAŞ', 'Kayseri ve Civarı', 'Vangölü EDAŞ'
# Referansda: 'AKEDAŞ (Göksu EDAŞ)', 'Aydem EDAŞ', 'AYEDAŞ', 'Dicle EDAŞ', 'Fırat EDAŞ', 'Kayseri EDAŞ', 'Vangölü EDAŞ'

# REFERANS formatından EXTRACTION formatına çevirme (reference'dan extraction'a)
REFERENCE_TO_EXTRACTION_MAPPING = {
    "AKEDAŞ (Göksu EDAŞ)": "AKEDAŞ",
    "Akdeniz EDAŞ": "Akdeniz EDAŞ",
    "Aras EDAŞ": "Aras EDAŞ",
    "Aydem EDAŞ": "ADM EDAŞ",
    "AYEDAŞ": "AYEDAŞ",
    "Başkent EDAŞ": "Başkent EDAŞ",
    "Boğaziçi EDAŞ": "Boğaziçi EDAŞ",
    "Çamlıbel EDAŞ": "Çamlıbel EDAŞ",
    "Çoruh EDAŞ": "Çoruh EDAŞ",
    "Dicle EDAŞ": "DİCLE EDAŞ",
    "Fırat EDAŞ": "FIRAT EDAŞ",
    "Gediz EDAŞ": "GDZ EDAŞ",
    "Kayseri EDAŞ": "KCETAŞ",
    "Meram EDAŞ": "MERAM",
    "Osmangazi EDAŞ": "OEDAŞ",
    "Sakarya EDAŞ": "SEDAŞ",
    "Toroslar EDAŞ": "TOROSLAR EDAŞ",
    "Trakya EDAŞ": "TREDAŞ",
    "Uludağ EDAŞ": "UEDAŞ",
    "Vangölü EDAŞ": "VANGÖLÜ EDAŞ",
    "Yeşilırmak EDAŞ": "YEŞİLIRMAK EDAŞ",
}

# Folder name'den reference format'a çevirme
# Bu mapping, process_all_regions()'dan gelen region isimleri (=folder isimleri) ile
# referans dosyasındaki region isimlerini eşleştirir
REGION_NORMALIZATION_MAPPING = {
    "AKEDAŞ": "AKEDAŞ (Göksu EDAŞ)",
    "ADM EDAŞ": "Aydem EDAŞ",
    "AYEDAŞ": "AYEDAŞ",
    "DİCLE EDAŞ": "Dicle EDAŞ",
    "Dicle EDAŞ": "Dicle EDAŞ",        # Folder name format
    "FIRAT EDAŞ": "Fırat EDAŞ",
    "Fırat EDAŞ": "Fırat EDAŞ",        # Folder name format
    "KAYSERI VE CIVARI": "Kayseri EDAŞ",
    "Kayseri ve Civarı": "Kayseri EDAŞ", # Folder name format
    "KCETAŞ": "Kayseri EDAŞ",
    "VANGÖLÜ EDAŞ": "Vangölü EDAŞ",
    "Vangölü EDAŞ": "Vangölü EDAŞ",    # Folder name format
    "BOĞAZİÇİ EDAŞ": "Boğaziçi EDAŞ",
    "BAŞKENT EDAŞ": "Başkent EDAŞ",
    "SEDAŞ": "Sakarya EDAŞ",
    "TREDAŞ": "Trakya EDAŞ",
    "UEDAŞ": "Uludağ EDAŞ",
    "ÇAMLIBEL EDAŞ": "Çamlıbel EDAŞ",
    "YEŞİLIRMAK EDAŞ": "Yeşilırmak EDAŞ",
    "Yeşilırmak EDAŞ": "Yeşilırmak EDAŞ", # Folder name -> reference
    "YEDAŞ": "Yeşilırmak EDAŞ",
    "GDZ EDAŞ": "Gediz EDAŞ",
    "MERAM": "Meram EDAŞ",
    "OEDAŞ": "Osmangazi EDAŞ",
    "TOROSLAR EDAŞ": "Toroslar EDAŞ",
    "ÇORUH EDAŞ": "Çoruh EDAŞ",
    "ARAS EDAŞ": "Aras EDAŞ",
    "AKDENİZ EDAŞ": "Akdeniz EDAŞ",
}

# casefold() tabanlı lookup dictionary oluştur (Türkçe İ/ı/Ğ/ğ safe)
_REGION_NORM_CASEFOLDED = {k.casefold(): v for k, v in REGION_NORMALIZATION_MAPPING.items()}

# REFERANS formatından EXTRACTION formatına çevirme
_REF_TO_EXT_CASEFOLDED = {k.casefold(): v for k, v in REFERENCE_TO_EXTRACTION_MAPPING.items()}


def normalize_region_name(region_name, to_format='extraction'):
    """
    Region ismini standartlaştırmak için fonksiyon
    to_format: 'extraction' -> reference formatından extraction formatına çevir
              'reference' -> extraction formatından reference formatına çevir
    
    casefold() kullanarak Türkçe İ/ı karakterlerini düzgün handle eder
    """
    if region_name is None:
        return None

    region_name = str(region_name).strip()

    if to_format == 'extraction':
        # Reference formatından extraction formatına çevir
        return _REF_TO_EXT_CASEFOLDED.get(region_name.casefold(), region_name)
    else:
        # Extraction formatından reference formatına çevir
        return _REGION_NORM_CASEFOLDED.get(region_name.casefold(), region_name)


def normalize_etso_kodu(etso_kodu):
    """
    Etso Kodu'yu standartlaştırmak için fonksiyon
    Farklı formatları (örn. "40Z000000123456T", "123456", 123456, "12345678") aynı forma çevirir.
    Kök numaraları 100% benzersiz şekilde tutar (collision-free).
    """
    if etso_kodu is None:
        return None

    etso_str = str(etso_kodu).strip()
    if not etso_str or etso_str.lower() in ('none', 'null', '-'):
        return None

    if '.' in etso_str:
        try:
            fval = float(etso_str)
            if fval == int(fval):
                etso_str = str(int(fval))
        except ValueError:
            pass

    eic_match = re.search(r'\d{2}Z0*([1-9]\d*)', etso_str, re.IGNORECASE)
    if eic_match:
        digits = eic_match.group(1)
        if digits[-1].isalpha():
            digits = digits[:-1]
        return digits

    digits = ''.join(re.findall(r'\d+', etso_str)).lstrip('0')
    return digits if digits else etso_str


# Standart sütunlar
STANDARD_COLUMNS = [
    'Dağıtım Bölgesi', 'ETSO Kodu', 'Müşteri', 'Tarife Grubu', 'AG OG',
    'Terim', 'Güç (kW)', 'Kurulu Güç', 'Aktif Enerji Tüketim (kWh)',
    'Dağıtım Bedeli (TL)', 'Güç Bedeli (TL)', 'Güç Aşım Bedeli (TL)',
    'Reaktif Bedel (TL)', 'KDV Matrahı (TL)', 'KDV', 'Toplam (TL)',
    'İlk Reaktif Bedeli (TL)', 'Sayax\'a Atılacak Tarife'
]

# Mapping 7 alanları yalnız ilgili mapping şeması aktifken kayıtlara/çıktıya eklenir.
# Bu koşullu yapı Mapping 6 golden kayıt sözlüklerini ve kanonik hash'ini korur.
EXTENDED_OUTPUT_COLUMNS = ['Standart Dışı Tutar (TL)', 'Tazminat Bedeli']
EXTENDED_SUM_FIELDS = ['Standart Dışı Tutar (TL)', 'Tazminat Bedeli']

# Aggregation rules per Notlar 3
# SUM: numeric fields that should be summed across duplicate ETSOs
SUM_FIELDS = ['Aktif Enerji Tüketim (kWh)', 'Dağıtım Bedeli(TL)', 'Güç Bedeli(TL)',
              'Güç Aşım Bedeli (TL)', 'Reaktif Bedel (TL)', 'İlk Reaktif']
# MAX: Güç kW and Kurulu Güç - never summed, take max (Notlar 3 #3)
MAX_FIELDS = ['Güç kW', 'KURULU GÜÇ']
# TEXT: take first non-empty value
TEXT_FIELDS_AGG = ['Müşteri', 'Tarife Grubu', 'AG OG', 'TERİM']


def detect_file_format(filepath):
    """Dosya formatını içeriğine göre tespit et"""
    try:
        with open(filepath, 'rb') as f:
            header = f.read(20)
        
        if header.startswith(b'<?xml'):
            return 'xml'
        elif b'<html xm' in header.lower() or header.startswith(b'<html'):
            return 'html'
        elif header.startswith(b'\x50\x4B\x03\x04'):
            return 'xlsx'  # ZIP format (xlsx)
        elif header.startswith(b'\xD0\xCF\x11\xE0'):
            return 'xls'   # OLE2 format (xls)
        else:
            # Excel olarak denemeyi devam et
            return 'excel'
    except Exception as e:
        print(f"  Dosya formatı tespiti hatası: {e}")
        return 'unknown'


def parse_excel_file(filepath):
    """Excel dosyasını oku ve sayfa/content döndür - format tespiti ile"""
    try:
        with open(filepath, 'rb') as f:
            header = f.read(20)
        
        # Formatı içeriğine göre tespit et (uzantıya güvenme)
        if header.startswith(b'<?xml'):
            # XML formatı - parse_xml_file kullan
            return None  # Caller should use parse_xml_file
        elif b'<html xm' in header.lower() or header.startswith(b'<html'):
            # HTML formatı - parse_html_file kullan
            return None  # Caller should use parse_html_file
        elif filepath.endswith(('.xlsx', '.XLSX')):
            wb = openpyxl.load_workbook(filepath)
            return {'type': 'xlsx', 'workbook': wb}
        elif filepath.endswith(('.xls', '.XLS')):
            try:
                wb = xlrd.open_workbook(filepath)
                return {'type': 'xls', 'workbook': wb}
            except Exception as xlrd_e:
                # xlrd başarısız oldu, HTML/XML olarak tekrar dene
                with open(filepath, 'rb') as f:
                    content = f.read(500)
                if b'<html xm' in content.lower() or content.startswith(b'<html'):
                    return None  # Caller should use parse_html_file
                elif content.startswith(b'<?xml') or b'<html' in content:
                    return None  # Caller should use parse_xml_file
                raise xlrd_e
    except Exception as e:
        return {'type': 'error', 'error': str(e)}
    return None


def parse_xml_file(filepath):
    """XML dosyasını oku - Excel 2003 XML formatı desteği"""
    try:
        tree = ET.parse(filepath)
        root = tree.getroot()
        return {'type': 'xml', 'root': root}
    except Exception as e:
        return {'type': 'error', 'error': str(e)}


def parse_html_file(filepath):
    """HTML dosyasını oku"""
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        return {'type': 'html', 'content': content}
    except Exception as e:
        return {'type': 'error', 'error': str(e)}


def _xml_row_values(row, namespace):
    """Excel XML satırındaki ss:Index boşluklarını koruyarak hücreleri aç."""
    values = []
    next_column = 1
    index_attribute = '{urn:schemas-microsoft-com:office:spreadsheet}Index'

    for cell in row.findall('ss:Cell', namespace):
        explicit_index = cell.get(index_attribute)
        if explicit_index:
            target_column = int(explicit_index)
            while next_column < target_column:
                values.append(None)
                next_column += 1

        data_elem = cell.find('ss:Data', namespace)
        values.append(data_elem.text if data_elem is not None else None)
        next_column += 1

    return values


def read_xml_content(
    file_info,
    region_mapping,
    region_name=None,
    source_name=None,
    warning_collector=None,
):
    """XML dosyasından veri çıkar (Excel 2003 XML formatı)"""
    data_rows = []

    if file_info['type'] != 'xml':
        return data_rows

    root = file_info['root']

    # Excel XML namespace
    ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}

    # Workbook içindeki tüm sayfaları bul
    worksheets = root.findall('.//ss:Worksheet', ns)

    for worksheet in worksheets:
        worksheet_name = worksheet.get('{urn:schemas-microsoft-com:office:spreadsheet}Name')
        table = worksheet.find('ss:Table', ns)
        if table is None:
            continue

        rows = list(table.findall('ss:Row', ns))
        if not rows:
            continue

        # Header satırını çıkar - dinamik header detection
        headers = None
        header_row_index = 0

        # Önce rows[0] kontrol et
        headers_row = rows[0]
        headers = _xml_row_values(headers_row, ns)

        # Eğer headers tek bir başlık banner satırı ise (<=2 dolu hücre), gerçek header satırını ara
        row0_vals = [h for h in headers if h is not None and str(h).strip()]
        if len(row0_vals) <= 2:
            # Tüm satırları tara, ilk gerçek header satırını (en az 3 sütunlu) bul
            for i, row in enumerate(rows):
                row_headers = _xml_row_values(row, ns)

                # En az 3 header değeri varsa, bu gerçek header satırı
                if len([h for h in row_headers if h is not None and str(h).strip()]) > 2:
                    headers = row_headers
                    header_row_index = i
                    break
        else:
            header_row_index = 0

        if not headers or not any(h is not None and str(h).strip() for h in headers):
            continue

        # Sütun indekslerini dinamik resolver ile çözümler
        indices = resolve_column_indices(headers, region_mapping)
        etso_idx = indices['etso']
        # Notlar 12 #3: Standart dışı format kontrolü
        if etso_idx < 0:
            if warning_collector is not None and is_known_source_header_layout(region_name):
                report_source_header_mismatches(
                    headers,
                    region_mapping,
                    region_name,
                    f"{source_name or 'Kaynak'} / {worksheet_name or 'Sayfa'}",
                    warning_collector,
                    header_row_number=header_row_index + 1,
                )
            print(f"  [UYARI] {region_name}: ETSO sütunu bulunamadı – standart dışı format, atlanıyor")
            continue
        if warning_collector is not None:
            report_source_header_mismatches(
                headers,
                region_mapping,
                region_name,
                f"{source_name or 'Kaynak'} / {worksheet_name or 'Sayfa'}",
                warning_collector,
                header_row_number=header_row_index + 1,
            )
        musteri_idx = indices['musteri']
        tarife_idx = indices['tarife']
        ag_og_idx = indices['ag_og']
        term_idx = indices['terim']
        guc_kw_idx = indices['güç_kw']
        kurulu_guc_idx = indices['kurulu_güç']
        aktif_idx = indices['aktif_enerji']
        dagitim_idx = indices['dagitim_bedeli']
        guc_bedeli_idx = indices['güç_bedeli']
        guc_asim_idx = indices['güç_aşım']
        trafo_kaybi_idx = indices['trafo_kaybı']

        # P+R aktif enerji kuralı kontrolü
        pr_sum_rule = region_mapping.get('rules', {}).get('active_energy_rule') == 'P + R sum'

        reaktif_fields, ilk_reaktif_fields = resolve_reactive_field_groups(
            region_mapping, region_name
        )

        # Veri satırlarını işle (header_row_index'den sonraki satırlar)
        for row in rows[header_row_index + 1:]:
            cell_values = _xml_row_values(row, ns)

            if not cell_values or all(v is None for v in cell_values):
                continue

            # Reaktif ve İlk Reaktif hesaplama — DRY: core.reactive_calculator
            from core.reactive_calculator import calculate_reactive_totals
            reaktif_toplam, ilk_reaktif_value, sakarya_x = calculate_reactive_totals(
                cell_values, indices, reaktif_fields, ilk_reaktif_fields, region_name
            )

            guc_val = extract_value_from_row(cell_values, guc_kw_idx, field_name='Güç kW', region_name=region_name)
            kurulu_val = extract_value_from_row(cell_values, kurulu_guc_idx, field_name='kurulu_güç', region_name=region_name)
            aktif_val = extract_value_from_row(cell_values, aktif_idx, field_name='aktif_enerji', region_name=region_name)
            trafo_val = extract_value_from_row(cell_values, trafo_kaybi_idx, field_name='trafo_kaybı', region_name=region_name)
            dagitim_val = extract_value_from_row(cell_values, dagitim_idx, field_name='dagitim_bedeli', region_name=region_name)
            guc_bedeli_val = extract_value_from_row(cell_values, guc_bedeli_idx, field_name='güç_bedeli', region_name=region_name)
            guc_asim_val = extract_value_from_row(cell_values, guc_asim_idx, field_name='güç_aşım', region_name=region_name)

            data_row = {
                'Dağıtım Bölgesi': region_mapping.get('region_name', ''),
                'Etso Kodu': extract_value_from_row(cell_values, etso_idx, field_name='etso', region_name=region_name),
                'Müşteri': cell_values[musteri_idx] if musteri_idx >= 0 and musteri_idx < len(cell_values) else None,
                'Tarife Grubu': cell_values[tarife_idx] if tarife_idx >= 0 and tarife_idx < len(cell_values) else None,
                'AG OG': cell_values[ag_og_idx] if ag_og_idx >= 0 and ag_og_idx < len(cell_values) else None,
                'TERİM': cell_values[term_idx] if term_idx >= 0 and term_idx < len(cell_values) else None,
                'Güç kW': guc_val,
                'KURULU GÜÇ': kurulu_val,
                'Aktif Enerji Tüketim (kWh)': (aktif_val + trafo_val) if pr_sum_rule and trafo_kaybi_idx >= 0 else aktif_val,
                'Dağıtım Bedeli(TL)': dagitim_val,
                'Güç Bedeli(TL)': guc_bedeli_val,
                'Güç Aşım Bedeli (TL)': guc_asim_val,
                'Reaktif Bedel (TL)': reaktif_toplam,
                'İlk Reaktif': ilk_reaktif_value
            }
            add_extended_financial_fields(
                data_row, cell_values, indices, region_mapping, region_name
            )
            data_rows.append(data_row)

    return data_rows


def read_html_content(
    file_info,
    region_mapping,
    region_name=None,
    source_name=None,
    warning_collector=None,
):
    """HTML dosyasından veri çıkar (Excel HTML formatı)"""
    data_rows = []

    if file_info['type'] != 'html':
        return data_rows

    content = file_info['content']

    # HTML parser kullanarak tabloyu bul
    parser = ExcelHTMLParser()
    parser.feed(content)

    if not parser.tables:
        return data_rows

    # Pick table with most rows (main data table)
    table = max(parser.tables, key=len)
    if not table or len(table) < 2:
        return data_rows

    # Dinamik header detection - HTML formatı için
    headers = table[0]
    header_row_index = 0

    # İlk satırın header olabileceğini kontrol et (başlık banner satırlarını atla)
    row0_vals = [h for h in headers if h is not None and str(h).strip()]
    if len(row0_vals) <= 2:
        # İlk satır banner, gerçek header satırını ara (en az 3 sütunlu)
        for i, row in enumerate(table):
            r_vals = [h for h in row if h is not None and str(h).strip()]
            if len(r_vals) > 2:
                headers = row
                header_row_index = i
                break

    if not any(h is not None and str(h).strip() for h in headers):
        return data_rows

    # Sütun indekslerini dinamik resolver ile çözümler
    indices = resolve_column_indices(headers, region_mapping)
    etso_idx = indices['etso']
    # Notlar 12 #3: Standart dışı format kontrolü
    if etso_idx < 0:
        if warning_collector is not None and is_known_source_header_layout(region_name):
            report_source_header_mismatches(
                headers,
                region_mapping,
                region_name,
                source_name or 'HTML kaynağı',
                warning_collector,
                header_row_number=header_row_index + 1,
            )
        print(f"  [UYARI] {region_name}: ETSO sütunu bulunamadı – standart dışı format, atlanıyor")
        return []
    if warning_collector is not None:
        report_source_header_mismatches(
            headers,
            region_mapping,
            region_name,
            source_name or 'HTML kaynağı',
            warning_collector,
            header_row_number=header_row_index + 1,
        )
    musteri_idx = indices['musteri']
    tarife_idx = indices['tarife']
    ag_og_idx = indices['ag_og']
    term_idx = indices['terim']
    guc_kw_idx = indices['güç_kw']
    kurulu_guc_idx = indices['kurulu_güç']
    aktif_idx = indices['aktif_enerji']
    dagitim_idx = indices['dagitim_bedeli']
    guc_bedeli_idx = indices['güç_bedeli']
    guc_asim_idx = indices['güç_aşım']
    trafo_kaybi_idx = indices['trafo_kaybı']

    # P+R aktif enerji kuralı kontrolü
    pr_sum_rule = region_mapping.get('rules', {}).get('active_energy_rule') == 'P + R sum'

    reaktif_fields, ilk_reaktif_fields = resolve_reactive_field_groups(
        region_mapping, region_name
    )

    # Veri satırlarını işle (header_row_index'den sonraki satırlar)
    for row in table[header_row_index + 1:]:
        if not row or all(v is None for v in row):
            continue

        # Reaktif ve İlk Reaktif hesaplama — DRY: core.reactive_calculator
        from core.reactive_calculator import calculate_reactive_totals
        reaktif_toplam, ilk_reaktif_value, sakarya_x = calculate_reactive_totals(
            row, indices, reaktif_fields, ilk_reaktif_fields, region_name
        )

        guc_val = extract_value_from_row(row, guc_kw_idx, field_name='Güç kW', region_name=region_name)
        kurulu_val = extract_value_from_row(row, kurulu_guc_idx, field_name='kurulu_güç', region_name=region_name)
        aktif_val = extract_value_from_row(row, aktif_idx, field_name='aktif_enerji', region_name=region_name)
        trafo_val = extract_value_from_row(row, trafo_kaybi_idx, field_name='trafo_kaybı', region_name=region_name)
        dagitim_val = extract_value_from_row(row, dagitim_idx, field_name='dagitim_bedeli', region_name=region_name)
        guc_bedeli_val = extract_value_from_row(row, guc_bedeli_idx, field_name='güç_bedeli', region_name=region_name)
        guc_asim_val = extract_value_from_row(row, guc_asim_idx, field_name='güç_aşım', region_name=region_name)

        data_row = {
            'Dağıtım Bölgesi': region_mapping.get('region_name', ''),
            'Etso Kodu': extract_value_from_row(row, etso_idx, field_name='etso', region_name=region_name),
            'Müşteri': row[musteri_idx] if musteri_idx >= 0 and musteri_idx < len(row) else None,
            'Tarife Grubu': row[tarife_idx] if tarife_idx >= 0 and tarife_idx < len(row) else None,
            'AG OG': row[ag_og_idx] if ag_og_idx >= 0 and ag_og_idx < len(row) else None,
            'TERİM': row[term_idx] if term_idx >= 0 and term_idx < len(row) else None,
            'Güç kW': guc_val,
            'KURULU GÜÇ': kurulu_val,
            'Aktif Enerji Tüketim (kWh)': (aktif_val + trafo_val) if pr_sum_rule and trafo_kaybi_idx >= 0 else aktif_val,
            'Dağıtım Bedeli(TL)': dagitim_val,
            'Güç Bedeli(TL)': guc_bedeli_val,
            'Güç Aşım Bedeli (TL)': guc_asim_val,
            'Reaktif Bedel (TL)': reaktif_toplam,
            'İlk Reaktif': ilk_reaktif_value
        }
        add_extended_financial_fields(
            data_row, row, indices, region_mapping, region_name
        )
        data_rows.append(data_row)

    return data_rows


class ExcelHTMLParser(HTMLParser):
    """Excel HTML formatındaki tabloları parse etmek için özel parser"""
    def __init__(self):
        super().__init__()
        self.tables = []  # Table[rows][cells]
        self.current_table = None
        self.current_row = None
        self.current_cell = None
        self.in_cell = False
        self.cell_data = []
    
    def handle_starttag(self, tag, attrs):
        attrs_dict = {k.lower(): v for k, v in attrs}
        
        if tag == 'table':
            self.current_table = []
            self.current_row = None
        elif tag == 'tr':
            self.current_row = []
        elif tag == 'td' or tag == 'th':
            self.in_cell = True
            self.cell_data = []
    
    def handle_endtag(self, tag):
        if tag == 'table':
            if self.current_table:
                self.tables.append(self.current_table)
            self.current_table = None
            self.current_row = None
        elif tag == 'tr':
            if self.current_row is not None and self.current_table is not None:
                self.current_table.append(self.current_row)
                self.current_row = None
        elif tag == 'td' or tag == 'th':
            self.in_cell = False
            cell_text = ''.join(self.cell_data).strip()
            if self.current_row is not None:
                self.current_row.append(cell_text if cell_text else None)
    
    def handle_data(self, data):
        if self.in_cell:
            self.cell_data.append(data)


def normalize_header(header):
    """Header'ı temizle ve normalleştir"""
    if header is None:
        return ''
    header = str(header).strip()
    # Türkçe karakterleri temizle ve küçük harfe çevir
    header = header.lower()
    replacements = {
        'ı': 'i', 'İ': 'i', 'ğ': 'g', 'Ğ': 'g', 'ü': 'u', 'Ü': 'u',
        'ö': 'o', 'Ö': 'o', 'ç': 'c', 'Ç': 'c', 'ş': 's', 'Ş': 's'
    }
    for tr, en in replacements.items():
        header = header.replace(tr, en)
    # Alt çizgiler ve boşlukları temizle
    header = re.sub(r'[\s_]+', '', header)
    # Parantezleri ve özel karakterleri temizle
    header = re.sub(r'[^\w]', '', header)
    return header


def collect_source_header_mismatches(
    headers,
    region_mapping,
    region_name,
    source_name,
    resolved_header_shifts=None,
):
    """Mapping'deki sütun harfi ile gerçek kaynak başlığını karşılaştır."""
    fields = [
        'etso', 'musteri', 'tarife', 'ag_og', 'terim', 'güç_kw', 'kurulu_güç',
        'aktif_enerji', 'trafo_kaybı', 'dagitim_bedeli', 'güç_bedeli',
        'güç_aşım', 'reaktif', 'reaktif2', 'reaktif_tenzil', 'reaktif_tenzil2',
        *region_mapping.get('tazminat_fields', []),
        *region_mapping.get('standart_disi_fields', []),
    ]
    is_dicle_region = normalize_header(region_name) == normalize_header('Dicle EDAŞ')
    resolved_indices = resolve_column_indices(headers, region_mapping)
    mismatches = []

    for field in fields:
        expected_header = region_mapping.get(field)
        expected_normalized = normalize_header(expected_header)
        column_index = region_mapping.get(f'{field}_index', -1)
        if not expected_normalized or not isinstance(column_index, int) or column_index < 0:
            continue

        actual_header = headers[column_index] if column_index < len(headers) else None
        if normalize_header(actual_header) == expected_normalized:
            continue

        matching_indices = [
            index
            for index, header in enumerate(headers)
            if normalize_header(header) == expected_normalized
        ]
        resolved_index = resolved_indices.get(field, -1)
        if resolved_index in matching_indices:
            if resolved_header_shifts is not None and resolved_index != column_index:
                resolved_header_shifts.append(field)
            if is_dicle_region:
                continue

        matching_columns = [index_to_excel_column(index) for index in matching_indices]
        resolved_header = (
            headers[resolved_index]
            if isinstance(resolved_index, int) and 0 <= resolved_index < len(headers)
            else None
        )
        resolved_location = '-'
        if isinstance(resolved_index, int) and 0 <= resolved_index < len(headers):
            resolved_name = (
                str(resolved_header).strip()
                if resolved_header is not None and str(resolved_header).strip()
                else '(başlıksız)'
            )
            resolved_location = f"{index_to_excel_column(resolved_index)} - {resolved_name}"
        mismatches.append({
            'Dağıtım Bölgesi': region_name or region_mapping.get('region_name', ''),
            'Kaynak': source_name,
            'Alan': field,
            'Eşlenen sütun': index_to_excel_column(column_index),
            'Beklenen başlık': str(expected_header),
            'Sütundaki başlık': str(actual_header) if actual_header is not None else '(boş)',
            'Başlığın bulunduğu sütun': ', '.join(matching_columns) if matching_columns else '(bulunamadı)',
            'Bulunduğu Başlık/Sütun': resolved_location,
        })

    return mismatches


def report_source_header_mismatches(
    headers,
    region_mapping,
    region_name,
    source_name,
    warning_collector,
    header_row_number=None,
):
    """Uyuşmazlıkları hem konsol çıktısına hem işlem istatistiklerine ekle."""
    format_warning = None
    if is_known_source_header_layout(region_name):
        format_warning = known_layout_change_warning(
            headers,
            region_name,
            source_name,
            header_row_number=header_row_number,
        )
        if not format_warning:
            return

    resolved_header_shifts = []
    mismatches = collect_source_header_mismatches(
        headers,
        region_mapping,
        region_name,
        source_name,
        resolved_header_shifts=resolved_header_shifts,
    )
    if not mismatches and format_warning:
        if resolved_header_shifts and normalize_header(region_name) == normalize_header('Dicle EDAŞ'):
            return
        warning_collector.append(format_warning)
        return

    warning_collector.extend(mismatches)


# Special header mappings for regions where expected header differs from actual
SPECIAL_HEADER_MAPPING = {
    'AKEDAŞ': {
        'etso': 'Id',  # Mapping file says 'PMUM ID' but actual is 'Id'
        'musteri': None,  # Not used for AKEDAŞ
    },
    'Çamlıbel EDAŞ': {
        # Notlar 7 #1 & Notlar 9 #1: Tarife Grubu, AG OG ve Terim AV sütunundan (index 47) okunur
        'tarife': 47,
        'ag_og': 47,
        'terim': 47,
    },
    'ÇAMLIBEL EDAŞ': {
        'tarife': 47,
        'ag_og': 47,
        'terim': 47,
    },
    'YEŞİLIRMAK EDAŞ': {
        'etso': 'Etso/Kullanıcı',
        'musteri': 'Müşteri Grubu',
        # Tarife, AG OG, Terim hepsi 'Müşteri Grubu' sütunundan doldurulur (Başlıklar sayfası)
        # Düzeltme tablosu ile dönüştürülecek
        'güç_bedeli': 'Güç Bedeli (TL)',
        'reaktif': 'Reaktif Bedeli (TL)',
        'reaktif_tenzil': 'İlk Reaktif İade Tutar',
    },
    'Toroslar EDAŞ': {
        'ag_og': 'Gerilim Seviyesi',
    },
    'Meram EDAŞ': {
        'reaktif': 'REAKTİF TÜKETİM',
        'reaktif_tenzil': 'Reaktif Bedel İhlal',
    },
    'MERAM EDAŞ': {
        'reaktif': 'REAKTİF TÜKETİM',
        'reaktif_tenzil': 'Reaktif Bedel İhlal',
    },
    'MERAM': {
        'reaktif': 'REAKTİF TÜKETİM',
        'reaktif_tenzil': 'Reaktif Bedel İhlal',
    },
    'Sakarya EDAŞ': {
        'reaktif': 'Reaktif Bedel (TL)',
    },
    'SAKARYA EDAŞ': {
        'reaktif': 'Reaktif Bedel (TL)',
    },
}


def find_column_index(headers, target_headers, region=None, field_name=None):
    """Header listesinden hedef header'ı bul"""
    # Special handling for region-specific mappings
    if region and field_name and region in SPECIAL_HEADER_MAPPING:
        special_field = SPECIAL_HEADER_MAPPING[region].get(field_name)
        if special_field:
            normalized_headers = [normalize_header(h) for h in headers]
            normalized_target = normalize_header(special_field)
            if normalized_target in normalized_headers:
                return normalized_headers.index(normalized_target)
    
    normalized_headers = [normalize_header(h) for h in headers]
    for target in target_headers:
        normalized_target = normalize_header(target)
        if normalized_target in normalized_headers:
            return normalized_headers.index(normalized_target)
    return -1


def clean_turkish_number(val, region_name=None, field_name=None):
    """Sayısal veriyi Türkçe/US sayı formatlarına göre akıllı temizle"""
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        val_float = float(val)
    else:
        val_str = str(val).strip()
        if not val_str or val_str.lower() in ('boş', 'none', 'null', '-'):
            return 0

        is_negative = False
        if val_str.startswith('(') and val_str.endswith(')'):
            is_negative = True
            val_str = val_str[1:-1].strip()
        elif val_str.startswith('-'):
            is_negative = True
            val_str = val_str[1:].strip()
        elif val_str.endswith('-'):  # Note #10: Yeşilırmak eksi değerlerin eksisi sonda!
            is_negative = True
            val_str = val_str[:-1].strip()

        # Handle comma and dot
        if ',' in val_str and '.' in val_str:
            first_comma = val_str.find(',')
            first_dot = val_str.find('.')
            if first_comma < first_dot:
                # US format: 1,688,498.61 or 876,697.41 -> remove commas
                val_str = val_str.replace(',', '')
            else:
                # TR format: 1.688.498,61 -> remove dots, replace comma with dot
                val_str = val_str.replace('.', '').replace(',', '.')
        elif ',' in val_str and '.' not in val_str:
            if re.search(r',\d{3}$', val_str) or len(val_str.split(',')) > 2:
                val_str = val_str.replace(',', '')
            else:
                val_str = val_str.replace(',', '.')
        elif '.' in val_str and ',' not in val_str:
            parts = val_str.split('.')
            if len(parts) > 2:
                val_str = val_str.replace('.', '')

        try:
            val_float = float(val_str)
            if is_negative:
                val_float = -val_float
        except ValueError:
            return 0

    # Not 4: Tüm dağıtım şirketlerinde kaynak veriler hiçbir çarpma veya bölme işlemi uygulanmadan aktarılır.
    # 1000-scaling rules kaldırıldı.

    return val_float


FIELD_SYNONYMS = {
    'etso': ['etsokodu', 'sayacid', 'sayacno', 'duykodu', 'pmumid', 'tesisatno', 'tesisatnumarasi', 'id', 'pmumcd'],
    'musteri': ['musteri', 'musteriadi', 'muhatapadi', 'aboneadi', 'unvan', 'firmadi', 'kisiadi', 'adsoyad', 'isim'],
    'tarife': ['trfnominaldrgr', 'trfnominal', 'nominaldrgr', 'tarifegrubu', 'abonetarifetipi', 'musterigrubu', 'tarife', 'tarifetipi'],
    'ag_og': ['agog', 'tarifetipi', 'gerilimtipi', 'ag/og', 'faturaagogadi'],
    'terim': ['terim', 'terimtipi', 'faturaterim'],
    'güç_kw': ['guckw', 'sozlesmegucu', 'baglantigucu', 'gucu'],
    'kurulu_güç': ['kuruluguc', 'kurulugucu'],
    'aktif_enerji': ['aktifenerjituketimkwh', 'toplamt1t2t3tuketim', 'dagitimmiktari', 'aktifenerji', 'toplamtuketim', 't0tuketim', 'toplamkwh'],
    'trafo_kaybı': ['trafokaybi', 'trafokaybituketim', 't0trafokaybi', 'trafokaybit0', 'trafokaybikwh', 'aktikkayip'],
    'dagitim_bedeli': ['dagitimbedelitl', 'dagitimbedeli', 'dagitimbdltl'],
    'güç_bedeli': ['gucbedelitl', 'gucbedeli'],
    'güç_aşım': ['gucasimbedelitl', 'gucasimbedeli', 'gucasimbdltl'],
    'reaktif': ['reaktifbedelitl', 'reaktifbedeli', 'reaktiftuketim', 'ribedeli', 'rcbedeli', 'reaktifbedel', 'reaktifinduktifkapasitiftl'],
    'reaktif2': ['reaktifbedelitl2', 'rcbedeli'],
    'reaktif_tenzil': ['reaktifbedeltenziltl', 'tenzilrcbedeli', 'ilkreaktifiadetutar'],
    'reaktif_tenzil2': ['reaktifbedeltenziltl2', 'tenzilribedeli']
}


def resolve_column_indices(headers, region_mapping):
    """
    Header satırına göre sütun indekslerini dinamik olarak çözümler.
    1. Mapping dosyasında tanımlı başlık ismi eşleştirilir.
    2. Eşleşmezse bilinen eş anlamlı başlıklar (FIELD_SYNONYMS) denenir.
    3. Eşleşmezse sabit sütun harfi indeksine düşer.
    """
    fields = ['etso', 'musteri', 'tarife', 'ag_og', 'terim', 'güç_kw', 'kurulu_güç',
              'aktif_enerji', 'trafo_kaybı', 'dagitim_bedeli', 'güç_bedeli',
              'güç_aşım', 'reaktif', 'reaktif2', 'reaktif_tenzil', 'reaktif_tenzil2']
    for field in (
        list(region_mapping.get('tazminat_fields', []))
        + list(region_mapping.get('standart_disi_fields', []))
    ):
        if field not in fields:
            fields.append(field)

    if not headers:
        return {field: region_mapping.get(f'{field}_index', -1) for field in fields}

    normalized_headers = [normalize_header(h) for h in headers]
    resolved = {}

    norm_reg = normalize_region_name(region_mapping.get('region_name', ''), to_format='reference')
    force_fixed_fields = set(region_mapping.get('rules', {}).get('force_fixed_fields', []))

    for field in fields:
        header_name = region_mapping.get(field)
        fixed_idx = region_mapping.get(f'{field}_index', -1)
        found_idx = -1

        # Notlar 7 / Çamlıbel: aynı isimli iki Tarife başlığından AV zorunlu.
        if field in force_fixed_fields and 0 <= fixed_idx < len(headers):
            found_idx = fixed_idx

        # Check SPECIAL_HEADER_MAPPING first (highest priority for region-specific overrides)
        if found_idx == -1 and norm_reg in SPECIAL_HEADER_MAPPING and field in SPECIAL_HEADER_MAPPING[norm_reg]:
            sp_val = SPECIAL_HEADER_MAPPING[norm_reg][field]
            if sp_val is None:
                found_idx = -2  # Explicitly disabled
            elif isinstance(sp_val, int):
                found_idx = sp_val
            else:
                norm_sp = normalize_header(sp_val)
                if norm_sp in normalized_headers:
                    found_idx = normalized_headers.index(norm_sp)

        if found_idx == -1 and header_name and pd.notna(header_name) and str(header_name).strip() != '-':
            norm_target = normalize_header(header_name)
            if norm_target and norm_target in normalized_headers:
                found_idx = normalized_headers.index(norm_target)

        # Keep established extraction behavior: use the mapped column before synonyms.
        if found_idx == -1 and 0 <= fixed_idx < len(headers):
            found_idx = fixed_idx

        if found_idx == -1:
            for syn in FIELD_SYNONYMS.get(field, []):
                if syn in normalized_headers:
                    found_idx = normalized_headers.index(syn)
                    break

        if found_idx == -2:
            found_idx = -1

        resolved[field] = found_idx

    return resolved


def resolve_reactive_field_groups(region_mapping, region_name):
    """Reaktif ve İlk Reaktif kaynak alanlarını bölgesel kurallarla seç."""
    reactive_rule = region_mapping.get('rules', {}).get('reactive_total_rule')
    norm_rn = normalize_region_name(region_name or '', to_format='reference')
    no_tenzil_regions = [
        'Aras EDAŞ', 'Çoruh EDAŞ', 'Dicle EDAŞ', 'Fırat EDAŞ', 'Vangölü EDAŞ', 'Meram EDAŞ'
    ]

    if reactive_rule in {'single_column', 'base_columns_only'} or norm_rn in no_tenzil_regions:
        reactive_fields = ['reaktif', 'reaktif2']
    else:
        reactive_fields = ['reaktif', 'reaktif2', 'reaktif_tenzil', 'reaktif_tenzil2']

    if norm_rn in ['Trakya EDAŞ', 'Trakya Edaş', 'TREDAŞ']:
        first_reactive_fields = ['reaktif_tenzil2']
    else:
        first_reactive_fields = ['reaktif_tenzil', 'reaktif_tenzil2']

    return reactive_fields, first_reactive_fields


def sum_mapped_numeric_fields(row, indices, field_names, region_name, field_name):
    """Bir mapping alan listesindeki sayısal değerleri güvenle topla."""
    total = 0.0
    for mapped_field in field_names:
        idx = indices.get(mapped_field, -1)
        if 0 <= idx < len(row):
            total += clean_turkish_number(
                row[idx],
                region_name=region_name,
                field_name=field_name,
            )
    return total


def add_extended_financial_fields(data_row, row, indices, region_mapping, region_name):
    """Mapping 7'nin iki yeni nihai tutarını ham satır sözlüğüne ekle."""
    if not region_mapping.get('features', {}).get('extended_financial_fields'):
        return data_row

    data_row['Tazminat Bedeli'] = sum_mapped_numeric_fields(
        row,
        indices,
        region_mapping.get('tazminat_fields', []),
        region_name,
        'tazminat_bedeli',
    )
    data_row['Standart Dışı Tutar (TL)'] = sum_mapped_numeric_fields(
        row,
        indices,
        region_mapping.get('standart_disi_fields', []),
        region_name,
        'standart_disi_tutar',
    )
    return data_row


def normalize_correction_lookup(value):
    """Düzeltme lookup için NFKC, kontrollü boşluk ve casefold normalizasyonu."""
    if value is None:
        return ''
    normalized = unicodedata.normalize('NFKC', str(value)).strip()
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.casefold()


def extract_value_from_row(row, column_index, field_name=None, region_name=None):
    """Satırdan değeri çıkar"""
    if 0 <= column_index < len(row):
        value = row[column_index]
        if field_name == 'etso':
            if isinstance(value, str):
                return value.strip() if value.strip() else None
            if isinstance(value, (int, float)):
                return str(int(value)) if value == int(value) else str(value)
            return value if value else None

        norm_reg = normalize_region_name(region_name or '', to_format='reference') if region_name else ''
        if ('yeşilırmak' in str(region_name).lower() or 'yeşilırmak' in str(norm_reg).lower() or 'yesilirmak' in normalize_header(region_name)) and field_name == 'Güç kW':
            if isinstance(value, str):
                val_str = value.strip()
                if val_str and val_str.lower() not in ('boş', 'none', 'null'):
                    if val_str.endswith('-'):
                        val_str = '-' + val_str[:-1].strip()
                    return val_str

        return clean_turkish_number(value, region_name=region_name, field_name=field_name)
    return 0


def load_mapping():
    """SKF Başlıkları.xlsx'den mapping'i yükle"""
    wb = openpyxl.load_workbook(MAPPING_FILE)
    ws = wb['Başlıklar']
    rows = list(ws.iter_rows(values_only=True))

    # Header mapping
    mapping = {}
    for row in rows[1:]:  # İlk satır header
        if row[0]:  # Dağıtım Bölgesi doluysa
            region = row[0]
            mapping[region] = {
                'etso': row[1],
                'musteri': row[2],
                'tarife': row[3],
                'ag_og': row[4],
                'term': row[5],
                'guc_kw': row[6],
                'kurulu_guc': row[7],
                'aktif_enerji': row[8],
                'dagitim_bedeli': row[9],
                'guc_bedeli': row[10],
                'guc_asim': row[11],
                'reaktif': row[12]
            }

    return mapping


def dynamic_region_list(mapping_file_path=None):
    """Mapping dosyasından bölgeleri dinamik olarak yükle."""
    mapping = parse_mapping_file(mapping_file_path or MAPPING_FILE)
    return list(mapping.get('regions', {}).keys())


def load_mapping_new(mapping_file_path=None):
    """Ana 16 alanı ve şema destekliyorsa Mapping 7 uzantılarını yükle."""
    return parse_mapping_file(mapping_file_path or MAPPING_FILE)


def discover_supported_files(region_path):
    """Bir bölge klasörünü alt klasörleriyle birlikte deterministik tara."""
    MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024  # 50 MB
    supported_files = []
    for root_dir, dir_names, file_names in os.walk(region_path):
        dir_names.sort(key=str.casefold)
        for file_name in sorted(file_names, key=str.casefold):
            if file_name.startswith('~$'):
                continue
            if file_name.lower().endswith(SUPPORTED_INPUT_EXTENSIONS):
                fpath = os.path.join(root_dir, file_name)
                fsize = os.path.getsize(fpath)
                if fsize > MAX_FILE_SIZE_BYTES:
                    print(f"  [UYARI] {file_name}: {fsize / (1024*1024):.1f}MB — "
                          f"50MB sınırını aşıyor, atlanıyor")
                    continue
                supported_files.append((file_name, fpath))
    return supported_files


def read_excel_content(
    file_info,
    region_mapping,
    region_name=None,
    alt_region_mapping=None,
    source_name=None,
    warning_collector=None,
):
    """Excel'den ana 16 alanı ve koşullu Mapping 7 alanlarını çıkar.
    
    Çift formatlı bölgeler (ör: Dicle EDAŞ) için alt_region_mapping parametresi
    ile alternatif başlık eşlemesi sağlanır. Dosya açıldığında birincil mapping
    ile ETSO sütunu bulunamazsa alternatif mapping otomatik devreye girer.
    """
    data_rows = []

    if file_info['type'] == 'xlsx':
        wb = file_info['workbook']
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                continue

            headers = rows[0]
            validation_headers = headers
            validation_header_row_number = 1
            if is_known_source_header_layout(region_name):
                validation_row_index, validation_headers = find_effective_header_row(rows)
                validation_header_row_number = validation_row_index + 1

            # Sütun indekslerini dinamik resolver ile çözümler
            active_mapping = region_mapping
            indices = resolve_column_indices(headers, active_mapping)

            # Çift format algılama (Notlar 12 #2): Başlık ismine göre dinamik
            # eşleştir, sabit sütun harfine güvenme. Birincil mapping'in eşleşme
            # kalitesini alternatif ile karşılaştır ve daha iyi olanı seç.
            if alt_region_mapping:
                primary_score = sum(1 for v in indices.values() if v >= 0)
                alt_indices = resolve_column_indices(headers, alt_region_mapping)
                alt_score = sum(1 for v in alt_indices.values() if v >= 0)
                
                if alt_score > primary_score:
                    active_mapping = alt_region_mapping
                    indices = alt_indices
                    print(f"  [ÇİFT FORMAT] Alternatif format kullanılıyor ({region_name}) "
                          f"[skor: {alt_score} vs {primary_score}]")

            # Notlar 12 #3: Standart dışı dosya kontrolü – ETSO sütunu bile
            # bulunamıyorsa bu sayfa standart formatlardan biri değildir.
            # Tahmin etme, okunamadığını belirt.
            etso_idx = indices['etso']
            if etso_idx < 0:
                if warning_collector is not None and is_known_source_header_layout(region_name):
                    report_source_header_mismatches(
                        validation_headers,
                        active_mapping,
                        region_name,
                        f"{source_name or 'Excel kaynağı'} / {sheet_name}",
                        warning_collector,
                        header_row_number=validation_header_row_number,
                    )
                print(f"  [UYARI] {region_name} / {sheet_name}: ETSO sütunu bulunamadı – "
                      f"standart dışı format, atlanıyor")
                continue
            if warning_collector is not None:
                report_source_header_mismatches(
                    validation_headers,
                    active_mapping,
                    region_name,
                    f"{source_name or 'Excel kaynağı'} / {sheet_name}",
                    warning_collector,
                    header_row_number=validation_header_row_number,
                )
            musteri_idx = indices['musteri']
            tarife_idx = indices['tarife']
            ag_og_idx = indices['ag_og']
            term_idx = indices['terim']
            guc_kw_idx = indices['güç_kw']
            kurulu_guc_idx = indices['kurulu_güç']
            aktif_idx = indices['aktif_enerji']
            dagitim_idx = indices['dagitim_bedeli']
            guc_bedeli_idx = indices['güç_bedeli']
            guc_asim_idx = indices['güç_aşım']
            trafo_kaybi_idx = indices['trafo_kaybı']

            # P+R aktif enerji kuralı kontrolü
            pr_sum_rule = region_mapping.get('rules', {}).get('active_energy_rule') == 'P + R sum'

            norm_rn = normalize_region_name(region_name or '', to_format='reference')
            reaktif_fields, ilk_reaktif_fields = resolve_reactive_field_groups(
                region_mapping, region_name
            )

            # Veri satırlarını işle
            for row in rows[1:]:
                if not row or all(v is None for v in row):
                    continue

                # Reaktif ve İlk Reaktif hesaplama — DRY: core.reactive_calculator
                from core.reactive_calculator import calculate_reactive_totals
                reactive_total, ilk_reaktif_value, sakarya_x = calculate_reactive_totals(
                    row, indices, reaktif_fields, ilk_reaktif_fields, region_name
                )

                guc_val = extract_value_from_row(row, guc_kw_idx, field_name='Güç kW', region_name=region_name)
                kurulu_val = extract_value_from_row(row, kurulu_guc_idx, field_name='kurulu_güç', region_name=region_name)
                aktif_val = extract_value_from_row(row, aktif_idx, field_name='aktif_enerji', region_name=region_name)
                trafo_val = extract_value_from_row(row, trafo_kaybi_idx, field_name='trafo_kaybı', region_name=region_name)
                dagitim_val = extract_value_from_row(row, dagitim_idx, field_name='dagitim_bedeli', region_name=region_name)
                if norm_rn == 'Akdeniz EDAŞ':
                    acma_kesme_idx = find_column_index(headers, ['Enerji_Acma_Kesme_Bedeli', 'acmakesmebedeli', 'kesmebedeli'], region=region_name)
                    if acma_kesme_idx >= 0 and acma_kesme_idx < len(row):
                        ak_val = clean_turkish_number(row[acma_kesme_idx], region_name=region_name, field_name='dagitim_bedeli')
                        if ak_val:
                            dagitim_val += ak_val

                # Notlar 12 #1: Sakarya EDAŞ aktif enerji artık doğrudan mapping'den
                # 'Dağıtım Miktarı' (Sütun AR) başlığıyla geliyor. Fallback gereksiz.

                guc_bedeli_val = extract_value_from_row(row, guc_bedeli_idx, field_name='güç_bedeli', region_name=region_name)
                guc_asim_val = extract_value_from_row(row, guc_asim_idx, field_name='güç_aşım', region_name=region_name)

                fatura_turu_idx = find_column_index(headers, ['faturaturu', 'fatura_turu', 'tahakkuk_turu'], region=region_name)
                fatura_no_idx = find_column_index(headers, ['faturano', 'fatura_no', 'tedarikcifaturano'], region=region_name)
                data_row = {
                    'Dağıtım Bölgesi': region_mapping.get('region_name', ''),
                    'Etso Kodu': extract_value_from_row(row, etso_idx, field_name='etso', region_name=region_name),
                    'Müşteri': row[musteri_idx] if musteri_idx >= 0 and musteri_idx < len(row) else None,
                    'Tarife Grubu': row[tarife_idx] if tarife_idx >= 0 and tarife_idx < len(row) else None,
                    'AG OG': row[ag_og_idx] if ag_og_idx >= 0 and ag_og_idx < len(row) else None,
                    'TERİM': row[term_idx] if term_idx >= 0 and term_idx < len(row) else None,
                    'Güç kW': guc_val,
                    'KURULU GÜÇ': kurulu_val,
                    'Aktif Enerji Tüketim (kWh)': (aktif_val + trafo_val) if pr_sum_rule and trafo_kaybi_idx >= 0 else aktif_val,
                    'Dağıtım Bedeli(TL)': dagitim_val,
                    'Güç Bedeli(TL)': guc_bedeli_val,
                    'Güç Aşım Bedeli (TL)': guc_asim_val,
                    'Reaktif Bedel (TL)': reactive_total,
                    'İlk Reaktif': ilk_reaktif_value,
                    'FATURA_TURU': row[fatura_turu_idx] if fatura_turu_idx >= 0 and fatura_turu_idx < len(row) else None,
                    'FATURA_NO': row[fatura_no_idx] if fatura_no_idx >= 0 and fatura_no_idx < len(row) else None,
                }
                add_extended_financial_fields(
                    data_row, row, indices, region_mapping, region_name
                )
                data_rows.append(data_row)

    elif file_info['type'] == 'xls':
        wb = file_info['workbook']
        for sheet_idx in range(wb.nsheets):
            sheet = wb.sheet_by_index(sheet_idx)
            if sheet.nrows == 0:
                continue

            header_row_idx = 0
            row0_vals = [str(v).strip() for v in sheet.row_values(0) if v is not None and str(v).strip()]
            if len(row0_vals) <= 2 and sheet.nrows > 1:
                header_row_idx = 1

            headers = sheet.row_values(header_row_idx)

            # Sütun indekslerini dinamik resolver ile çözümler
            indices = resolve_column_indices(headers, region_mapping)
            etso_idx = indices['etso']
            # Notlar 12 #3: Standart dışı format kontrolü
            if etso_idx < 0:
                if warning_collector is not None and is_known_source_header_layout(region_name):
                    report_source_header_mismatches(
                        headers,
                        region_mapping,
                        region_name,
                        f"{source_name or 'Excel kaynağı'} / {sheet.name}",
                        warning_collector,
                        header_row_number=header_row_idx + 1,
                    )
                print(f"  [UYARI] {region_name} / {sheet.name}: ETSO sütunu bulunamadı – "
                      f"standart dışı format, atlanıyor")
                continue
            if warning_collector is not None:
                report_source_header_mismatches(
                    headers,
                    region_mapping,
                    region_name,
                    f"{source_name or 'Excel kaynağı'} / {sheet.name}",
                    warning_collector,
                    header_row_number=header_row_idx + 1,
                )
            musteri_idx = indices['musteri']
            tarife_idx = indices['tarife']
            ag_og_idx = indices['ag_og']
            term_idx = indices['terim']
            guc_kw_idx = indices['güç_kw']
            kurulu_guc_idx = indices['kurulu_güç']
            aktif_idx = indices['aktif_enerji']
            dagitim_idx = indices['dagitim_bedeli']
            guc_bedeli_idx = indices['güç_bedeli']
            guc_asim_idx = indices['güç_aşım']
            trafo_kaybi_idx = indices['trafo_kaybı']

            # P+R aktif enerji kuralı kontrolü
            pr_sum_rule = region_mapping.get('rules', {}).get('active_energy_rule') == 'P + R sum'

            reaktif_fields, ilk_reaktif_fields = resolve_reactive_field_groups(
                region_mapping, region_name
            )

            for row_idx in range(header_row_idx + 1, sheet.nrows):
                row = sheet.row_values(row_idx)
                if not row or all(v in (None, '') for v in row):
                    continue

                # Reaktif ve İlk Reaktif hesaplama — DRY: core.reactive_calculator
                from core.reactive_calculator import calculate_reactive_totals
                reactive_total, ilk_reaktif_value, sakarya_x = calculate_reactive_totals(
                    row, indices, reaktif_fields, ilk_reaktif_fields, region_name
                )

                guc_val = extract_value_from_row(row, guc_kw_idx, field_name='Güç kW', region_name=region_name)
                kurulu_val = extract_value_from_row(row, kurulu_guc_idx, field_name='kurulu_güç', region_name=region_name)
                aktif_val = extract_value_from_row(row, aktif_idx, field_name='aktif_enerji', region_name=region_name)
                trafo_val = extract_value_from_row(row, trafo_kaybi_idx, field_name='trafo_kaybı', region_name=region_name)
                dagitim_val = extract_value_from_row(row, dagitim_idx, field_name='dagitim_bedeli', region_name=region_name)
                guc_bedeli_val = extract_value_from_row(row, guc_bedeli_idx, field_name='güç_bedeli', region_name=region_name)
                guc_asim_val = extract_value_from_row(row, guc_asim_idx, field_name='güç_aşım', region_name=region_name)

                data_row = {
                    'Dağıtım Bölgesi': region_mapping.get('region_name', ''),
                    'Etso Kodu': extract_value_from_row(row, etso_idx, field_name='etso', region_name=region_name),
                    'Müşteri': row[musteri_idx] if musteri_idx >= 0 and musteri_idx < len(row) else None,
                    'Tarife Grubu': row[tarife_idx] if tarife_idx >= 0 and tarife_idx < len(row) else None,
                    'AG OG': row[ag_og_idx] if ag_og_idx >= 0 and ag_og_idx < len(row) else None,
                    'TERİM': row[term_idx] if term_idx >= 0 and term_idx < len(row) else None,
                    'Güç kW': guc_val,
                    'KURULU GÜÇ': kurulu_val,
                    'Aktif Enerji Tüketim (kWh)': (aktif_val + trafo_val) if pr_sum_rule and trafo_kaybi_idx >= 0 else aktif_val,
                    'Dağıtım Bedeli(TL)': dagitim_val,
                    'Güç Bedeli(TL)': guc_bedeli_val,
                    'Güç Aşım Bedeli (TL)': guc_asim_val,
                    'Reaktif Bedel (TL)': reactive_total,
                    'İlk Reaktif': ilk_reaktif_value
                }
                add_extended_financial_fields(
                    data_row, row, indices, region_mapping, region_name
                )
                data_rows.append(data_row)

    return data_rows


def filter_extracted_data(data, region_name=None):
    """Alt başlık/boş/pozitif-absürt satırları ele; negatif bedelleri koru.

    10 milyar TL filtresi yalnız pozitif ``bedel > threshold`` kontrolüdür.
    Akdeniz dahil negatif kayıtlar Rule #2 gereği elenmez.
    """
    filtered_data = []
    zero_bedel_filtered = 0
    negative_bedel_count = 0
    akdeniz_negative_filtered = 0
    absurd_bedel_filtered = 0

    # Threshold for absurd bedel values - 10 billion TL
    BEDEL_THRESHOLD = 10_000_000_000.0

    HEADER_ETSO_VALUES = {'duykodu', 'etso', 'etsokodu', 'tesisatno', 'tesisatnumarasi', 'sayacid', 'sayac id'}
    HEADER_CUSTOMER_VALUES = {'unvan', 'adsoyad', 'musteri', 'musteriadi', 'muhatapadi', 'aboneadi'}

    for row in data:
        bedel = row.get('Dağıtım Bedeli(TL)', 0)
        customer = row.get('Müşteri', None)
        row_region = row.get('Dağıtım Bölgesi', '')
        etso = row.get('Etso Kodu', '')

        # Skip sub-header rows where etso or customer contains literally header titles
        etso_str = str(etso or '').strip().lower()
        cust_str = str(customer or '').strip().lower()
        if etso_str in HEADER_ETSO_VALUES or cust_str in HEADER_CUSTOMER_VALUES:
            continue

        # Skip rows with absurdly large bedel values (>10 billion TL) - corrupted data
        if bedel > BEDEL_THRESHOLD:
            absurd_bedel_filtered += 1
            print(f"  [WARNING] Filtering absurd bedel value: {bedel:,.2f} TL (region: {row_region})")
            continue

        # Yeşilırmak EDAŞ specific logging for large values (>1M TL)
        if region_name and 'yeşilırmak' in region_name.lower() and abs(bedel) > 1_000_000:
            print(f"  [Yeşilırmak LARGE] Etso: {row.get('Etso Kodu')}, Customer: {row.get('Müşteri')}, Region: {row_region}, Bedel: {bedel:,.2f} TL")

        # Yeşilırmak EDAŞ specific logging for negative values
        if region_name and 'yeşilırmak' in region_name.lower() and bedel < 0:
            print(f"  [Yeşilırmak NEGATIVE] Etso: {row.get('Etso Kodu')}, Customer: {row.get('Müşteri')}, Region: {row_region}, Bedel: {bedel:,.2f} TL")

        # Rule #2: "Eksi veriler getirilmeli" - keep all negative values across all regions (including Akdeniz)
        if bedel < 0:
            negative_bedel_count += 1
            # Do NOT skip - keep negative values

        # bedel=0.0 AND etso=None AND customer=None (tamamen boş satırları) hariç tut
        if bedel == 0.0 and (etso is None or str(etso).strip() == '') and (customer is None or str(customer).strip() == ''):
            zero_bedel_filtered += 1
            continue

        filtered_data.append(row)

    return filtered_data, zero_bedel_filtered, negative_bedel_count, akdeniz_negative_filtered, absurd_bedel_filtered


def process_all_regions(base_path=None, mapping_file_path=None, status_callback=None, return_stats=False):
    """
    Tüm bölgeleri 16 ana alan ve şemaya bağlı Mapping 7 uzantılarıyla işle.

    Varsayılan parametreler eski terminal davranışını korur. Streamlit,
    dinamik klasör/mapping yolları ve istatistik dönüşü için opsiyonel
    parametreleri kullanır.
    """
    active_base_path = os.path.abspath(base_path or BASE_PATH)
    active_mapping_file = os.path.abspath(mapping_file_path or MAPPING_FILE)
    if not os.path.isdir(active_base_path):
        raise FileNotFoundError(f"Ana klasör bulunamadı: {active_base_path}")
    if not os.path.isfile(active_mapping_file):
        raise FileNotFoundError(f"Mapping dosyası bulunamadı: {active_mapping_file}")

    mapping = load_mapping_new(active_mapping_file)
    all_data = []
    total_extracted = 0
    total_zero_bedel_filtered = 0
    total_negative_bedel_count = 0
    total_akdeniz_filtered = 0
    total_absurd_bedel_filtered = 0
    total_discovered_files = 0
    total_processed_files = 0
    file_errors = []
    source_header_warnings = []
    per_region_records = {}
    correction_counts = {'Tarife Grubu': 0, 'AG OG': 0, 'TERİM': 0}
    correction_example_counts = {'Tarife Grubu': 0, 'AG OG': 0, 'TERİM': 0}
    correction_examples = []
    normalized_corrections = {}
    for output_field, lookup in mapping.get('corrections', {}).items():
        normalized_lookup = {}
        for source_value, target_value in lookup.items():
            normalized_key = normalize_correction_lookup(source_value)
            if normalized_key:
                normalized_lookup[normalized_key] = {
                    'source': source_value,
                    'target': target_value,
                }
        normalized_corrections[output_field] = normalized_lookup

    normalized_sayax_tariffs = {}
    for rule in mapping.get('sayax_tariff_rules', []):
        rule_key = tuple(
            normalize_correction_lookup(rule.get(field))
            for field in ('tarife', 'ag_og', 'terim')
        )
        if all(rule_key):
            normalized_sayax_tariffs[rule_key] = rule.get('target', '')

    print("=== VERİ ÇIKARILIYOR ===\n")

    # Dynamic region list from mapping file
    regions = list(mapping.get('regions', {}).keys())
    if status_callback:
        status_callback("Mapping dosyası okundu; dağıtım bölgeleri işleniyor.", 5)

    # Dinamik Dağıtım Adı normalizasyonu: Düzeltme sayfasının G/H sütunlarından
    # gelen casefold lookup'ı, mevcut sabit REGION_NORMALIZATION_MAPPING ile birleştir.
    # Düzeltme sayfasındaki güncel veriler önceliklidir.
    dynamic_dist_names = mapping.get('distribution_names', {})
    merged_folder_lookup = {k.casefold(): v for k, v in REGION_NORMALIZATION_MAPPING.items()}
    merged_folder_lookup.update(dynamic_dist_names)  # Düzeltme öncelikli

    # Gerçek dosya sistemindeki klasör isimlerini tara ve mapping bölgelerine eşleştir.
    # Bu sayede klasör adı "BEDAŞ" bile olsa "Boğaziçi EDAŞ" bölgesine doğru yönlenir.
    actual_folders = {}
    if os.path.isdir(active_base_path):
        for folder_name in os.listdir(active_base_path):
            folder_full = os.path.join(active_base_path, folder_name)
            if os.path.isdir(folder_full) and not folder_name.startswith('.'):
                actual_folders[folder_name.casefold()] = folder_name

    for region_index, region in enumerate(regions, start=1):
        if status_callback:
            progress_value = 5 + int(((region_index - 1) / max(1, len(regions))) * 85)
            status_callback(f"İşleniyor: {region}", progress_value)
        # Region name normalization - mapping file uses certain format
        mapped_region_name = region  # Already normalized in mapping file

        if mapped_region_name not in mapping.get('regions', {}):
            print(f"[ERROR] {region}: Mapping bulunamadı")
            continue

        region_mapping = mapping['regions'][mapped_region_name].copy()
        region_mapping['region_name'] = region
        region_mapping['features'] = mapping.get('features', {}).copy()

        # Apply region-specific rules from Notlar sheet
        rules = apply_note_rules(mapping, mapped_region_name)
        region_mapping['rules'] = rules

        # Klasör çözümleme: Önce birebir isimle dene, bulamazsan dinamik eşleştirme kullan.
        region_path = os.path.join(active_base_path, region)
        if not os.path.exists(region_path):
            # Dinamik çözümleme: Gerçek klasör isimlerinden hangisi bu mapping bölgesine
            # eşleşiyor? merged_folder_lookup'ta canonical değeri 'region' olan bir
            # klasör adı bul.
            resolved = False
            for cf_folder, real_folder_name in actual_folders.items():
                canonical = merged_folder_lookup.get(cf_folder)
                if canonical and canonical.casefold() == region.casefold():
                    region_path = os.path.join(active_base_path, real_folder_name)
                    print(f"  [DİNAMİK] Klasör '{real_folder_name}' → bölge '{region}' olarak eşleştirildi")
                    resolved = True
                    break
            if not resolved:
                print(f"[ERROR] {region}: Klasör bulunamadı")
                continue

        # Vangölü EDAŞ dahil tüm alt klasörleri derinlemesine tara.
        excel_files = discover_supported_files(region_path)
        total_discovered_files += len(excel_files)

        if not excel_files:
            print(f"[WARNING] {region}: Hiç Excel dosyası bulunamadı")
            continue

        region_all_data = []
        processed_count = 0

        # Çift format desteği: Dicle EDAŞ gibi birden fazla başlık formatına sahip
        # bölgeler için alternatif mapping'i hazırla.
        alt_region_mapping = mapping.get('alt_regions', {}).get(region)
        if alt_region_mapping:
            alt_region_mapping = alt_region_mapping.copy()
            alt_region_mapping['region_name'] = region
            alt_region_mapping['features'] = mapping.get('features', {}).copy()
            alt_region_mapping['rules'] = rules

        for fname, fpath in excel_files:
            # Region-specific file filtering (AKEDAŞ: TL_Raporu, Uludağ: 4008-TL)
            if not extract_file_filter(region, rules, fname):
                print(f"  {fname}: Dosya filtre ile atlandı (region: {region})")
                continue

            file_format = detect_file_format(fpath)
            file_was_processed = False

            if file_format == 'xlsx':
                file_info = parse_excel_file(fpath)
                if file_info and file_info['type'] == 'xlsx':
                    data = read_excel_content(
                        file_info,
                        region_mapping,
                        region_name=region,
                        alt_region_mapping=alt_region_mapping,
                        source_name=fname,
                        warning_collector=source_header_warnings,
                    )
                    region_all_data.extend(data)
                    processed_count += len(data)
                    file_was_processed = True

            elif file_format == 'xls':
                file_info = parse_excel_file(fpath)
                if file_info is None:
                    # .xls uzantılı ama aslında HTML/XML formatında, parse_html_file ile dene
                    file_info = parse_html_file(fpath)
                    if file_info and file_info['type'] == 'html':
                        data = read_html_content(
                            file_info,
                            region_mapping,
                            region_name=region,
                            source_name=fname,
                            warning_collector=source_header_warnings,
                        )
                        region_all_data.extend(data)
                        processed_count += len(data)
                        file_was_processed = True
                        print(f"  {fname}: HTML formatı (xls uzantılı) - {len(data)} satır çıkarıldı")
                    else:
                        # XML formatında olabilir, parse_xml_file ile dene
                        file_info = parse_xml_file(fpath)
                        if file_info and file_info['type'] == 'xml':
                            data = read_xml_content(
                                file_info,
                                region_mapping,
                                region_name=region,
                                source_name=fname,
                                warning_collector=source_header_warnings,
                            )
                            region_all_data.extend(data)
                            processed_count += len(data)
                            file_was_processed = True
                            print(f"  {fname}: XML formatı (xls uzantılı) - {len(data)} satır çıkarıldı")
                elif file_info and file_info['type'] == 'xls':
                    data = read_excel_content(
                        file_info,
                        region_mapping,
                        region_name=region,
                        alt_region_mapping=alt_region_mapping,
                        source_name=fname,
                        warning_collector=source_header_warnings,
                    )
                    region_all_data.extend(data)
                    processed_count += len(data)
                    file_was_processed = True
                elif file_info and file_info['type'] == 'error':
                    print(f"  {fname}: Hata - {file_info['error']}")
                    file_errors.append(f"{region} / {fname}: {file_info['error']}")

            elif file_format == 'xml':
                file_info = parse_xml_file(fpath)
                if file_info and file_info['type'] == 'xml':
                    data = read_xml_content(
                        file_info,
                        region_mapping,
                        region_name=region,
                        source_name=fname,
                        warning_collector=source_header_warnings,
                    )
                    region_all_data.extend(data)
                    processed_count += len(data)
                    file_was_processed = True
                    print(f"  {fname}: XML formatı - {len(data)} satır çıkarıldı")

            elif file_format == 'html':
                file_info = parse_html_file(fpath)
                if file_info and file_info['type'] == 'html':
                    data = read_html_content(
                        file_info,
                        region_mapping,
                        region_name=region,
                        source_name=fname,
                        warning_collector=source_header_warnings,
                    )
                    region_all_data.extend(data)
                    processed_count += len(data)
                    file_was_processed = True
                    print(f"  {fname}: HTML formatı - {len(data)} satır çıkarıldı")

            else:
                print(f"  {fname}: Bilinmeyen format")

            if file_was_processed:
                total_processed_files += 1

        total_extracted += processed_count
        print(f"* {region}: {processed_count} satır çıkarıldı ({len(excel_files)} dosya)")

        # Per-region filtering
        filtered_region_data, zero_bedel_filtered, negative_bedel_count, akdeniz_filtered, absurd_bedel_filtered = filter_extracted_data(region_all_data, region_name=region)

        all_data.extend(filtered_region_data)
        per_region_records[region] = len(filtered_region_data)

        # Per-region statistics
        if zero_bedel_filtered > 0:
            print(f"  -> {region}: {zero_bedel_filtered} records filtered (bedel=0.0 AND customer=None)")
        if negative_bedel_count > 0:
            print(f"  -> {region}: {negative_bedel_count} negative bedel kaydı (Rule #2)")
        if akdeniz_filtered > 0:
            print(f"  -> {region}: {akdeniz_filtered} negative distribution bedel values filtered")
        if absurd_bedel_filtered > 0:
            print(f"  -> {region}: {absurd_bedel_filtered} absurd bedel values filtered (>10B TL)")

        total_zero_bedel_filtered += zero_bedel_filtered
        total_negative_bedel_count += negative_bedel_count
        total_akdeniz_filtered += akdeniz_filtered
        total_absurd_bedel_filtered += absurd_bedel_filtered

    print(f"\nINFO: {total_negative_bedel_count} negatif bedel kaydı extraction'da var ve Rule #2 gereği korunuyor")

    if total_zero_bedel_filtered > 0:
        print(f"\nWARNING: {total_zero_bedel_filtered} records filtered (bedel=0.0 AND customer=None)")

    if total_akdeniz_filtered > 0:
        print(f"\nINFO: {total_akdeniz_filtered} negative distribution bedel values filtered for Akdeniz EDAŞ")

    if total_absurd_bedel_filtered > 0:
        print(f"\nINFO: {total_absurd_bedel_filtered} absurd bedel values filtered (>10B TL - corrupted data)")

    # ETSO Aggregation: Notlar 3 #1, #2, #3
    # Aynı ETSO koduna sahip kayıtları topla (bölge bazında)
    # Sayısal alanlar: SUM, Güç kW/Kurulu Güç: MAX, Metin alanları: ilk boş olmayan değer
    print(f"\n=== ETSO AGREGASYONU ===")
    print(f"Before aggregation: {len(all_data)} records")
    
    from collections import OrderedDict
    
    # Group by (Etso Kodu normalized, Dağıtım Bölgesi)
    aggregation_groups = OrderedDict()
    
    for record in all_data:
        etso_raw = record.get('Etso Kodu', '')
        etso_norm = normalize_etso_kodu(etso_raw) or str(etso_raw or '').strip()
        region = record.get('Dağıtım Bölgesi', '')
        key = (etso_norm, region)
        
        if key not in aggregation_groups:
            aggregation_groups[key] = []
        aggregation_groups[key].append(record)
    
    aggregated_data = []
    aggregation_count = 0
    
    def post_process_record(r, reg_name):
        rec = r.copy()
        
        # 1. Tarife Grubu
        tarife = str(rec.get('Tarife Grubu') or '').strip()
        if reg_name != 'Boğaziçi EDAŞ':
            for suffix in [' (Ek)', ' (ek)', ' ek', ' (EK)']:
                if tarife.endswith(suffix):
                    tarife = tarife[:-len(suffix)].strip()
        if reg_name in ['ADM EDAŞ', 'Gediz EDAŞ']:
            if tarife.upper().startswith('TICARET') or tarife.upper().startswith('TİCARET'):
                tarife = 'TICARETHAN'
            elif tarife.upper().startswith('SANAYI') or tarife.upper().startswith('SANAYİ'):
                tarife = 'SANAYI'
        elif reg_name in ['Çamlıbel EDAŞ', 'Yeşilırmak EDAŞ']:
            # Çamlıbel ve Yeşilırmak: ham değer korunur, Düzeltme tablosu dönüştürecek
            pass
        rec['Tarife Grubu'] = tarife

        # 2. AG OG
        ag_og = str(rec.get('AG OG') or '').strip()
        if reg_name in ['Çamlıbel EDAŞ', 'Yeşilırmak EDAŞ']:
            # Çamlıbel ve Yeşilırmak: ham değer korunur, Düzeltme tablosu dönüştürecek
            pass
        elif reg_name not in ['Başkent EDAŞ', 'Toroslar EDAŞ', 'Gediz EDAŞ', 'ADM EDAŞ', 'Trakya EDAŞ']:
            ag_og_upper = ag_og.upper()
            if 'AG' in ag_og_upper or 'DAG' in ag_og_upper:
                ag_og = 'AG'
            elif 'OG' in ag_og_upper or 'DOG' in ag_og_upper:
                ag_og = 'OG'
        elif reg_name == 'Toroslar EDAŞ':
            pass
        rec['AG OG'] = ag_og

        # 3. TERİM
        terim = str(rec.get('TERİM') or '').strip()
        if reg_name == 'Yeşilırmak EDAŞ':
            # Yeşilırmak: Terim, Müşteri Grubu sütunundan gelir, Düzeltme tablosu dönüştürür
            pass
        elif reg_name == 'Boğaziçi EDAŞ':
            if rec.get('Tarife Grubu') == 'Boş' or (abs(float(rec.get('Aktif Enerji Tüketim (kWh)', 0) or 0)) < 1e-4 and abs(float(rec.get('Dağıtım Bedeli(TL)', 0) or 0)) < 1e-4):
                terim = ''
            elif 'tek' in terim.lower():
                terim = 'Tek Terimli'
            elif 'cift' in terim.lower() or 'çift' in terim.lower():
                terim = 'Çift Terimli'
            else:
                terim = ''
        elif reg_name == 'Osmangazi EDAŞ':
            if 'cift terim' in terim.lower() or 'çift terim' in terim.lower():
                terim = 'Çift Terim-Tek Zamanlı    '
            elif 'tek terim' in terim.lower():
                terim = 'Tek Terim-Tek Zamanlı    '
            # Not 2: Osmangazi Güç kW (Baglanti Gucu) sıfırlanmayacak, kaynak veriden gelecek
        elif reg_name == 'Çamlıbel EDAŞ':
            # Çamlıbel: ham değer korunur, Düzeltme tablosu dönüştürecek
            pass
        else:
            if 'Tek Terim' in terim:
                terim = 'Tek Terimli'
            elif 'Çift Terim' in terim:
                terim = 'Çift Terimli'
            elif terim not in ('Tek Terimli', 'Çift Terimli'):
                # Madde 7: Bilinmeyen/eksik TERİM → Güç Bedeli'ne göre otomatik doldur
                guc_bedeli = rec.get('Güç Bedeli(TL)', 0) or 0
                try:
                    guc_bedeli_float = float(guc_bedeli)
                except (ValueError, TypeError):
                    guc_bedeli_float = 0.0
                if guc_bedeli_float > 0:
                    terim = 'Çift Terimli'
                else:
                    terim = 'Tek Terimli'
        rec['TERİM'] = terim

        # 4. Müşteri
        if reg_name == 'Çamlıbel EDAŞ':
            rec['Müşteri'] = ''

        # 5. Mapping 7 / Düzeltme: yalnız lookup'ta bulunan değerleri değiştir.
        for output_field in ['Tarife Grubu', 'AG OG', 'TERİM']:
            current_value = rec.get(output_field)
            lookup_key = normalize_correction_lookup(current_value)
            match = normalized_corrections.get(output_field, {}).get(lookup_key)
            if not match:
                continue

            target_value = match['target']
            if str(current_value) == str(target_value):
                continue

            rec[output_field] = target_value
            correction_counts[output_field] += 1
            if correction_example_counts[output_field] < 10:
                correction_examples.append({
                    'region': reg_name,
                    'field': output_field,
                    'source': current_value,
                    'target': target_value,
                })
                correction_example_counts[output_field] += 1

        # 6. Global kural: AG OG != AG/OG ve Terim == "Çift Terim" ise -> AG OG = "OG"
        final_ag_og = str(rec.get('AG OG') or '').strip()
        final_terim = str(rec.get('TERİM') or '').strip()
        if final_ag_og and final_ag_og not in ('AG', 'OG'):
            if 'Çift' in final_terim:
                rec['AG OG'] = 'OG'

        sayax_key = tuple(
            normalize_correction_lookup(rec.get(field))
            for field in ('Tarife Grubu', 'AG OG', 'TERİM')
        )
        rec["Sayax'a Atılacak Tarife"] = normalized_sayax_tariffs.get(sayax_key, '')

        return rec

    for (etso_norm, region), records in aggregation_groups.items():
        if len(records) == 1:
            record = post_process_record(records[0], region)
            aggregated_data.append(record)
            continue
        
        aggregation_count += len(records) - 1
        
        # Start with first record as base
        merged = records[0].copy()
        
        active_sum_fields = list(SUM_FIELDS)
        for optional_field in EXTENDED_SUM_FIELDS:
            if any(optional_field in source_record for source_record in records):
                active_sum_fields.append(optional_field)

        for field in active_sum_fields:
            total = 0
            for r in records:
                val = r.get(field, 0)
                if isinstance(val, str) and val.upper() == 'X':
                    continue  # Sakarya X values - skip in sum
                try:
                    total += float(val or 0)
                except (ValueError, TypeError):
                    pass
            merged[field] = total
        
        for field in MAX_FIELDS:
            max_val = 0
            max_raw = None
            for r in records:
                val = r.get(field, 0)
                if val is not None and str(val).strip() and str(val).strip().lower() not in ('none', 'boş'):
                    v_float = clean_turkish_number(val, region_name=region, field_name=field) if isinstance(val, str) else float(val or 0)
                    if v_float > max_val or max_raw is None:
                        max_val = v_float
                        max_raw = val
            merged[field] = max_raw if max_raw is not None else max_val
        
        for field in TEXT_FIELDS_AGG:
            if field == 'Tarife Grubu' and region == 'Boğaziçi EDAŞ':
                non_empty = [r.get(field) for r in records if r.get(field) and str(r.get(field)).strip().lower() not in ('none', 'boş')]
                akt = float(merged.get('Aktif Enerji Tüketim (kWh)', 0) or 0)
                dag = float(merged.get('Dağıtım Bedeli(TL)', 0) or 0)
                if non_empty:
                    # Boş olmayan bir tarife varsa her zaman onu seç
                    merged['Tarife Grubu'] = non_empty[0]
                elif (akt <= 0 and dag <= 0) or (abs(akt) < 1e-4 and abs(dag) < 1e-4):
                    merged['Tarife Grubu'] = 'Boş'
                    merged['TERİM'] = ''
                else:
                    # non_empty boş ama tuketim/bedel var - ilk kaydi kullan
                    first_val = records[0].get(field)
                    merged['Tarife Grubu'] = first_val if first_val else 'Boş'
            elif field == 'TERİM' and region == 'Boğaziçi EDAŞ':
                if merged.get('Tarife Grubu') == 'Boş':
                    merged['TERİM'] = ''
                else:
                    non_empty = [r.get('TERİM') for r in records if r.get('TERİM') and str(r.get('TERİM')).strip() and str(r.get('TERİM')).strip().lower() not in ('none', 'boş', '0')]
                    merged['TERİM'] = non_empty[0] if non_empty else 'Tek Terimli'
            elif field == 'Tarife Grubu':
                # Madde 9: 'Boş' string değerini geçersiz say; boş olmayan ilk değeri tercih et
                non_bos = [r.get(field) for r in records if r.get(field) and str(r.get(field)).strip().lower() not in ('none', 'boş')]
                if non_bos:
                    merged[field] = non_bos[0]
                else:
                    # Tüm değerler 'Boş' veya boşsa ilk kaydı kullan
                    for r in records:
                        val = r.get(field)
                        if val is not None and str(val).strip() and str(val).strip().lower() != 'none':
                            merged[field] = val
                            break
            else:
                for r in records:
                    val = r.get(field)
                    if val is not None and str(val).strip() and str(val).strip().lower() not in ('none', 'boş'):
                        merged[field] = val
                        break

        if region == 'Yeşilırmak EDAŞ' and abs(merged.get('Aktif Enerji Tüketim (kWh)', 0)) < 1e-4:
            pos_akt = sum(float(r.get('Aktif Enerji Tüketim (kWh)', 0) or 0) for r in records if float(r.get('Aktif Enerji Tüketim (kWh)', 0) or 0) > 0)
            if pos_akt > 0:
                merged['Aktif Enerji Tüketim (kWh)'] = pos_akt

        merged = post_process_record(merged, region)
        aggregated_data.append(merged)
    
    all_data = aggregated_data
    
    print(f"After aggregation: {len(all_data)} records")
    print(f"Records merged: {aggregation_count} records")
    print(f"Extraction tamamlandi: {total_extracted} satir cikarildi, {len(all_data)} satir filtering+deduplication sonrasi")
    print(f"    (Negatif bedeller korundu; yalnız tamamen boş sıfır satırlar ve pozitif absürt >10B TL kayıtlar elendi)")

    # Çıkarılan veride gerçekten bulunan eşsiz bölgeleri doğrula.
    processed_regions = {
        row.get('Dağıtım Bölgesi')
        for row in all_data
        if row.get('Dağıtım Bölgesi')
    }
    normalized_processed_regions = {
        normalize_region_name(region, to_format='reference')
        for region in processed_regions
    }
    successful_regions = [
        region for region in regions
        if normalize_region_name(region, to_format='reference') in normalized_processed_regions
    ]
    missing_regions = [region for region in regions if region not in successful_regions]

    warning_details_by_company = {}
    for warning in source_header_warnings:
        company = warning.get('Dağıtım Bölgesi') or 'Bilinmeyen dağıtım şirketi'
        company_details = warning_details_by_company.setdefault(company, set())
        company_details.add((
            warning.get('Eşlenen süt'),
            warning.get('Beklenen başlık'),
            warning.get('Sütundaki başlık'),
            warning.get('Başlığın bulunduğu sütun'),
        ))
    for company, details in warning_details_by_company.items():
        print(f"[UYARI] Format Farkı: {company} ({len(details)} farklı başlık farkı)")

    stats = {
        'expected_region_count': len(regions),
        'read_region_count': len(successful_regions),
        'successful_regions': successful_regions,
        'missing_regions': missing_regions,
        'processed_regions': sorted(processed_regions, key=str.casefold),
        'total_extracted': total_extracted,
        'total_records': len(all_data),
        'zero_filtered': total_zero_bedel_filtered,
        'negative_count': total_negative_bedel_count,
        'akdeniz_filtered': total_akdeniz_filtered,
        'absurd_filtered': total_absurd_bedel_filtered,
        'merged_duplicates': aggregation_count,
        'discovered_files': total_discovered_files,
        'processed_files': total_processed_files,
        'file_errors': file_errors,
        'per_region_records': per_region_records,
        'notes_sheet': mapping.get('notes_sheet'),
        'mapping_features': mapping.get('features', {}).copy(),
        'mapping_warnings': list(mapping.get('mapping_warnings', [])),
        'source_header_warnings': source_header_warnings,
        'correction_duplicates': list(mapping.get('correction_duplicates', [])),
        'correction_counts': correction_counts.copy(),
        'correction_total': sum(correction_counts.values()),
        'correction_examples': correction_examples,
        'standart_disi_total': sum(
            float(row.get('Standart Dışı Tutar (TL)', 0) or 0) for row in all_data
        ),
        'tazminat_total': sum(
            float(row.get('Tazminat Bedeli', 0) or 0) for row in all_data
        ),
    }

    if status_callback:
        status_callback("İşlem tamamlandı.", 100)

    if return_stats:
        return all_data, stats
    return all_data


def _format_guc_kw(value):
    """Yeşilırmak EDAŞ Güç kW string değerlerini (ör. '500,00000000000000') 2 ondalıklı float'a dönüştür."""
    if value is None or value == '':
        return ''
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    val_str = str(value).strip()
    if not val_str or val_str.lower() in ('none', 'boş'):
        return ''
    # Türkçe virgüllü string veya düz sayı
    try:
        cleaned = val_str.replace(' ', '').replace('.', '').replace(',', '.')
        return round(float(cleaned), 2)
    except (ValueError, TypeError):
        return val_str  # Dönüştürülemezse orijinali bırak


def _find_eno_file(base_path):
    """Veri klasöründe 'XX ENO AyAdı' formatında ENO dosyasını otomatik bul.
    
    Örnek: '01 ENO Ocak.xlsx', '06 ENO Haziran.xlsx'
    """
    import re
    if not base_path or not os.path.isdir(base_path):
        return None
    
    eno_pattern = re.compile(r'^\d{2}\s+ENO\s+', re.IGNORECASE)
    
    for fname in os.listdir(base_path):
        if eno_pattern.match(fname) and (fname.endswith('.xlsx') or fname.endswith('.xls')):
            fpath = os.path.join(base_path, fname)
            if os.path.isfile(fpath) and not fname.startswith('~$'):
                return fpath
    return None


def _load_eno_lookups(eno_path):
    """ENO dosyasından iki lookup tablosu oluştur:
    
    1. sayac_id_to_eic: Sayaç ID (E sütunu) → Sayaç EIC Kod (D sütunu)
    2. eic_to_abone: Sayaç EIC Kod (D sütunu) → Abone Ad-Soyad/Unvan (AD sütunu)
    
    Satır 1: üst başlık (merged), Satır 2: gerçek başlık, Satır 3+: veri
    """
    sayac_id_to_eic = {}  # str(Sayaç ID) → str(EIC Kod)
    eic_to_abone = {}     # str(EIC Kod) → str(Abone Ad-Soyad)
    
    try:
        wb = openpyxl.load_workbook(eno_path, read_only=False, data_only=True)
        ws = wb.active
        
        row_count = 0
        for row in ws.iter_rows(min_row=3, max_col=35, values_only=False):  # Satır 3'ten başla (veri)
            row_count += 1
            # D sütunu (col 4, 1-indexed) = Sayaç EIC Kod
            # E sütunu (col 5, 1-indexed) = Sayaç ID
            # AD sütunu (col 30, 1-indexed) = Abone Ad-Soyad / Unvan
            d_val = None
            e_val = None
            ad_val = None
            
            for cell in row:
                if cell.column == 4:
                    d_val = cell.value
                elif cell.column == 5:
                    e_val = cell.value
                elif cell.column == 30:
                    ad_val = cell.value
            
            if d_val is not None:
                eic_str = str(d_val).strip()
                
                # Sayaç ID → EIC Kod eşleştirmesi
                if e_val is not None:
                    # Sayaç ID'yi normalize et (float'dan int'e çevir: 20088.0 → "20088")
                    e_str = str(e_val).strip()
                    if '.' in e_str:
                        try:
                            e_str = str(int(float(e_str)))
                        except (ValueError, TypeError):
                            pass
                    sayac_id_to_eic[e_str] = eic_str
                
                # EIC Kod → Abone Ad-Soyad eşleştirmesi
                if ad_val is not None:
                    ad_str = str(ad_val).strip()
                    if ad_str:
                        eic_to_abone[eic_str] = ad_str
        
        wb.close()
        print(f"  [ENO] {row_count} satır okundu: {len(sayac_id_to_eic)} Sayaç ID→EIC, {len(eic_to_abone)} EIC→Abone")
        
    except Exception as e:
        print(f"  [ENO HATA] {eno_path}: {e}")
    
    return sayac_id_to_eic, eic_to_abone


def enrich_from_eno_file(records, base_path):
    """Çıktı kayıtlarını ENO dosyasından zenginleştir:
    
    1. ETSO Kodu '40Z' ile başlamıyorsa → ENO'daki Sayaç ID eşleşmesiyle EIC Kod'a çevir
    2. Müşteri alanı boşsa → EIC Kod ile ENO'dan Abone Ad-Soyad/Unvan getir
    
    Returns:
        dict: Zenginleştirme istatistikleri
    """
    stats = {'eno_file': None, 'etso_converted': 0, 'musteri_filled': 0, 'total_records': len(records)}
    
    eno_path = _find_eno_file(base_path)
    if not eno_path:
        print("[ENO] Veri klasöründe ENO dosyası bulunamadı (format: 'XX ENO AyAdı.xlsx')")
        return stats
    
    stats['eno_file'] = os.path.basename(eno_path)
    print(f"\n[ENO] Dosya bulundu: {stats['eno_file']}")
    
    sayac_id_to_eic, eic_to_abone = _load_eno_lookups(eno_path)
    
    if not sayac_id_to_eic and not eic_to_abone:
        print("[ENO] Lookup tabloları boş, zenginleştirme yapılamadı")
        return stats
    
    # ADIM 1: ETSO Kodu dönüşümü (Sayaç ID → EIC Kod)
    for record in records:
        etso = record.get('Etso Kodu', '')
        if etso is None:
            etso = ''
        etso_str = str(etso).strip()
        
        # Zaten 40Z formatındaysa dokunma
        if etso_str.startswith('40Z') and len(etso_str) >= 16:
            continue
        
        # Sayaç ID olarak ENO'da ara
        # Float'tan int'e normalize et (20088.0 → "20088")
        lookup_key = etso_str
        if '.' in lookup_key:
            try:
                lookup_key = str(int(float(lookup_key)))
            except (ValueError, TypeError):
                pass
        
        if lookup_key in sayac_id_to_eic:
            new_etso = sayac_id_to_eic[lookup_key]
            record['Etso Kodu'] = new_etso
            stats['etso_converted'] += 1
    
    # ADIM 2: Boş Müşteri doldurma (EIC Kod → Abone Ad-Soyad)
    for record in records:
        musteri = record.get('Müşteri', '')
        if musteri is not None and str(musteri).strip():
            continue  # Zaten dolu, dokunma
        
        etso = record.get('Etso Kodu', '')
        if etso is None:
            continue
        etso_str = str(etso).strip()
        
        if etso_str in eic_to_abone:
            record['Müşteri'] = eic_to_abone[etso_str]
            stats['musteri_filled'] += 1
    
    print(f"  [ENO] ETSO dönüştürülen: {stats['etso_converted']}, Müşteri doldurulan: {stats['musteri_filled']}")
    return stats


def save_to_excel(data, filepath):
    """Veriyi Excel'e kaydet - Formüller, Sayı Biçimleri ve Başlık Biçimlendirmesi ile"""
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Çıkarılan Veriler'
    
    extended_output_enabled = any(
        any(field in row for field in EXTENDED_OUTPUT_COLUMNS)
        for row in data
    )
    output_columns = list(STANDARD_COLUMNS)
    if extended_output_enabled:
        output_columns.extend(EXTENDED_OUTPUT_COLUMNS)

    # Header satırı ekle
    ws.append(output_columns)

    # Başlık satırı biçimlendirmesi: Kalın, CK Mavi (#305496) arka plan, beyaz yazı, ortalama
    header_font = Font(bold=True, color='FFFFFF', size=11)
    header_fill = PatternFill(fill_type='solid', fgColor='305496')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=False)
    data_align_left = Alignment(horizontal='left', vertical='center')
    for col_idx in range(1, len(output_columns) + 1):
        cell = ws.cell(1, col_idx)
        cell.number_format = 'General'
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    # Başlığı dondur (A2'den itibaren kaydır)
    ws.freeze_panes = 'A2'

    # Otomatik filtre (tüm sütunlara)
    last_col_letter = ws.cell(1, len(output_columns)).column_letter
    ws.auto_filter.ref = f'A1:{last_col_letter}1'

    COMMA_STYLE_FORMAT = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'
    
    # Veri satırları
    for idx, row in enumerate(data, start=2):
        output_row = [
            row.get('Dağıtım Bölgesi', ''),
            row.get('Etso Kodu', ''),
            row.get('Müşteri', ''),
            row.get('Tarife Grubu', ''),
            row.get('AG OG', ''),
            row.get('TERİM', ''),
            _format_guc_kw(row.get('Güç kW', '')),
            row.get('KURULU GÜÇ', ''),
            row.get('Aktif Enerji Tüketim (kWh)', ''),
            row.get('Dağıtım Bedeli(TL)', ''),
            row.get('Güç Bedeli(TL)', ''),
            row.get('Güç Aşım Bedeli (TL)', ''),
            row.get('Reaktif Bedel (TL)', ''),
            f"=J{idx}+K{idx}+L{idx}+M{idx}",           # N: KDV Matrahı (TL)
            f"=N{idx}*0.2",                              # O: KDV
            f"=N{idx}+O{idx}+Q{idx}",                    # P: Toplam (TL)
            row.get('İlk Reaktif', ''),                  # Q: İlk Reaktif Bedeli (TL)
            row.get("Sayax'a Atılacak Tarife", '')       # R: Sayax'a Atılacak Tarife
        ]
        if extended_output_enabled:
            output_row.extend([
                row.get('Standart Dışı Tutar (TL)', 0),
                row.get('Tazminat Bedeli', 0),
            ])
        ws.append(output_row)

        # Hizalama: Tum sutunlar sola hizali
        for c in range(1, len(output_columns) + 1):
            ws.cell(idx, c).alignment = data_align_left

        # Number Formatting:
        # A(1), C(3), D(4), E(5), F(6), R(18) -> General
        for c in [1, 3, 4, 5, 6, 18]:
            if c <= len(output_columns):
                ws.cell(idx, c).number_format = 'General'

        # B(2) -> Sayı ('0')
        ws.cell(idx, 2).number_format = '0'

        # G(7)..Q(17) -> Virgül Stili (Güç, Kurulu, Aktif, Bedeller, KDV Matrahı, KDV, Toplam)
        for c in range(7, 18):
            ws.cell(idx, c).number_format = COMMA_STYLE_FORMAT

        if extended_output_enabled:
            for c in range(len(STANDARD_COLUMNS) + 1, len(output_columns) + 1):
                ws.cell(idx, c).number_format = COMMA_STYLE_FORMAT

    # Kaydetmeden önce tüm sütunları max karakter uzunluğu + 2 ile genişlet.
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value is not None and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2
    
    try:
        wb.save(filepath)
        print(f"\n* Veri {filepath} konumuna kaydedildi ({len(data)} satır)")
    except PermissionError:
        print(f"\n[WARNING] {filepath} dosyası Excel'de açık olduğu için doğrudan kaydedilemedi!")
        alt_path = filepath.replace('.xlsx', '_Guncel.xlsx')
        wb.save(alt_path)
        print(f"* Veri alternatif olarak {alt_path} konumuna kaydedildi ({len(data)} satır)")


def compare_with_reference(extracted_data, reference_file):
    """Referans dosyasıyla karşılaştır - region normalization ile"""
    wb = openpyxl.load_workbook(reference_file)
    ws = wb['Dağıtımın Kestiği']
    reference_rows = list(ws.iter_rows(values_only=True))

    reference_data = []
    for row in reference_rows[1:]:  # Header hariç
        if row[0] and row[0] != 'Dağıtım Bölgesi':
            # Reference region name'i extraction formatına çevir ( Karşılaştırma için her ikisini de aynı formata çevirip key oluşturacağız )
            reference_region = normalize_region_name(row[0], to_format='reference')
            reference_data.append({
                'Dağıtım Bölgesi': reference_region,
                'Etso Kodu': row[1],
                'Müşteri': row[2],
                'Tarife Grubu': row[3],
                'AG OG': row[4],
                'TERİM': row[5],
                'Güç kW': row[6],
                'KURULU GÜÇ': row[7],
                'Aktif Enerji Tüketim (kWh)': row[8],
                'Dağıtım Bedeli(TL)': row[9],
                'Güç Bedeli(TL)': row[10],
                'Güç Aşım Bedeli (TL)': row[11],
                'Reaktif Bedel (TL)': row[12]
            })

    # Reference'daki toplam bedel
    ref_total_original = sum((r['Dağıtım Bedeli(TL)'] or 0) for r in reference_data)

    # Etso Kodu + Müşteri + Dağıtım Bedeli + Bölge ile eşleşme kontrolü
    # Müşteri değerini stringe çevir ve boşlukları temizle
    def normalize_musteri(value):
        if value is None:
            return ''
        return str(value).strip().upper()  # Case-insensitive comparison

    # Matching key doesn't include customer name - used for comparison logic
    def create_matching_key(row, region_field, etso_field, bedel_field):
        etso = normalize_etso_kodu(row[etso_field])
        bedel = (row[bedel_field] or 0)
        region = normalize_region_name(row.get(region_field), to_format='reference')
        # Key excludes customer name - we match on Etso+Bedel+Region only
        return (etso, bedel, region)

    # Dictionary yerine list-based counting kullan - duplicate handling için
    # Her key için record'ların listesini tut, duplication'da count artar
    from collections import Counter

    # Priority filtering: when same Etso+Bedel+Region has both None and non-None customers,
    # keep only non-None customer records
    def prioritize_non_none_customers(data, region_field, etso_field, bedel_field, musteri_field, original_data):
        """Filter data to prioritize non-None customers when same Etso+Bedel+Region exists with both."""
        # Group by Etso+Bedel+Region
        groups = {}
        for idx, row in enumerate(data):
            etso = normalize_etso_kodu(row[etso_field])
            bedel = (row[bedel_field] or 0)
            # Data zaten normalize edilmiş, tekrar normalize etme
            region = row.get(region_field)
            key = (etso, bedel, region)

            if key not in groups:
                groups[key] = []
            groups[key].append({'row': row, 'index': idx})

        # Debug: log Çamlıbel filtering BEFORE filtering
        camlibel_groups = [(key, items) for key, items in groups.items() if 'ÇAMLIBEL' in str(key).upper()]
        if camlibel_groups:
            print(f"\n=== DEBUG: Çamlıbel groups BEFORE filtering ===")
            for key, items in camlibel_groups:
                non_none_items = [item for item in items
                                if item['row'][musteri_field] is not None and str(item['row'][musteri_field]).strip() != '']
                print(f"Group {key}: {len(items)} items, {len(non_none_items)} non-None")
                for item in items:
                    print(f"  - {item['row'][etso_field]} | {item['row'][musteri_field]} | {item['row'][bedel_field]}")

        # For each group, filter out None customers if non-None exists
        result_indices = set()
        for key, items in groups.items():
            non_none_items = [item for item in items
                            if item['row'][musteri_field] is not None and str(item['row'][musteri_field]).strip() != '']

            # If we have non-None customers, keep only those indices; otherwise keep all
            if non_none_items:
                for item in non_none_items:
                    result_indices.add(item['index'])
            else:
                for item in items:
                    result_indices.add(item['index'])

        # Debug: log Çamlıbel filtering AFTER filtering
        if camlibel_groups:
            print(f"\n=== DEBUG: Çamlıbel groups AFTER filtering ===")
            filtered_data = [original_data[i] for i in sorted(result_indices)]
            camlibel_filtered = [d for d in filtered_data if 'ÇAMLIBEL' in str(d.get(region_field, '')).upper()]
            for item in camlibel_filtered:
                etso = normalize_etso_kodu(item[etso_field])
                bedel = (item[bedel_field] or 0)
                region = normalize_region_name(item.get(region_field), to_format='reference')
                print(f"  - {etso} | {item[musteri_field]} | {bedel} | {region}")

        return [original_data[i] for i in sorted(result_indices)]

    # Extract data zaten normalized, sadece field'ı yeniden adlandır
    extracted_records = []
    for row in extracted_data:
        extracted_records.append({
            **row,
            'normalized_region': row.get('Dağıtım Bölgesi')
        })

    # Apply priority filtering to both reference and extracted data
    reference_data_filtered = prioritize_non_none_customers(
        reference_data, 'Dağıtım Bölgesi', 'Etso Kodu', 'Dağıtım Bedeli(TL)', 'Müşteri', reference_data
    )
    extracted_data_filtered = prioritize_non_none_customers(
        extracted_records, 'normalized_region', 'Etso Kodu', 'Dağıtım Bedeli(TL)', 'Müşteri', extracted_records
    )

    # Matching keys (without customer name) kullanarak eşleşme yap
    # Reference data için matching key oluştur
    reference_matching_keys = []
    reference_key_details = []  # (matching_key, full_key, row) tuple for tracking
    for row in reference_data_filtered:
        etso_normalized = normalize_etso_kodu(row['Etso Kodu'])
        musteri_normalized = normalize_musteri(row['Müşteri'])
        bedel = round(float(row['Dağıtım Bedeli(TL)'] or 0), 2)
        # Reference data zaten normalize edilmiş, tekrar normalize etme
        normalized_region = row.get('Dağıtım Bölgesi')
        matching_key = (etso_normalized, bedel, normalized_region)
        full_key = (etso_normalized, musteri_normalized, bedel, normalized_region)
        reference_matching_keys.append(matching_key)
        reference_key_details.append((matching_key, full_key, row))

    # Extract data için matching key
    extracted_matching_keys = []
    extracted_key_details = []  # (matching_key, full_key, record) tuple for tracking
    for row in extracted_data_filtered:
        # extracted_records zaten normalized_region field'ı ile normalize edilmiş
        normalized_region = row.get('normalized_region')
        etso_normalized = normalize_etso_kodu(row['Etso Kodu'])
        musteri_normalized = normalize_musteri(row['Müşteri'])
        bedel = round(float(row['Dağıtım Bedeli(TL)'] or 0), 2)
        matching_key = (etso_normalized, bedel, normalized_region)
        full_key = (etso_normalized, musteri_normalized, bedel, normalized_region)
        extracted_matching_keys.append(matching_key)
        extracted_key_details.append((matching_key, full_key, row))

    # Counter kullanarak duplicate count'ları al (matching key bazlı)
    reference_matching_counter = Counter(reference_matching_keys)
    extracted_matching_counter = Counter(extracted_matching_keys)

    # Initialize all_matching_keys from both datasets
    all_matching_keys = set(reference_matching_keys) | set(extracted_matching_keys)
    matched_count = 0
    missing = []
    extra = []

    # Index dictionaries for efficient lookup - matching key bazlı
    # Her matching key için record listesi tut (customer name ignored in matching)
    reference_matching_index = {}
    for matching_key, full_key, row in reference_key_details:
        if matching_key not in reference_matching_index:
            reference_matching_index[matching_key] = []
        reference_matching_index[matching_key].append({'row': row, 'used': False})

    extracted_matching_index = {}
    for matching_key, full_key, record in extracted_key_details:
        if matching_key not in extracted_matching_index:
            extracted_matching_index[matching_key] = []
        extracted_matching_index[matching_key].append({'record': record, 'used': False})

    # Matching key bazlı eşleştirme
    # Her matching key için, record count'larını karşılaştır
    for matching_key in all_matching_keys:
        ref_count = reference_matching_counter.get(matching_key, 0)
        ext_count = extracted_matching_counter.get(matching_key, 0)

        # Bu matching key'e ait record count'larını al
        ref_count_key = ref_count
        ext_count_key = ext_count

        matched = min(ref_count_key, ext_count_key)
        matched_count += matched

        # Eksik kayıtları bul - matching key bazlı
        missing_count = ref_count_key - matched
        if missing_count > 0 and matching_key in reference_matching_index:
            count = 0
            for item in reference_matching_index[matching_key]:
                if not item['used'] and count < missing_count:
                    missing.append(item['row'])
                    item['used'] = True
                    count += 1

        # Fazla kayıtları bul - matching key bazlı
        extra_count = ext_count_key - matched
        if extra_count > 0 and matching_key in extracted_matching_index:
            count = 0
            for item in extracted_matching_index[matching_key]:
                if not item['used'] and count < extra_count:
                    extra.append(item['record'])
                    item['used'] = True
                    count += 1

    print(f"\n=== KARŞILAŞTIRMA SONUÇLARI ===")
    print(f"Referans dosyası: {len(reference_data)} satır")
    print(f"Referans dosyası (filtered): {len(reference_data_filtered)} satır")
    print(f"Çıkarılan veri (all): {len(extracted_data)} satır")
    print(f"Çıkarılan veri (filtered): {len(extracted_data_filtered)} satır")
    print(f"Matched records: {matched_count}")
    print(f"Eksik kayıtlar: {len(missing)}")
    print(f"Fazla kayıtlar: {len(extra)}")

    # Debug: Same Etso+Bedel values in both missing and extra
    print(f"\n=== DEBUG: Same Etso+Bedel in both lists ===")
    missing_set = {(r['Etso Kodu'], r['Dağıtım Bedeli(TL)']) for r in missing}
    extra_set = {(r['Etso Kodu'], r['Dağıtım Bedeli(TL)']) for r in extra}
    common = missing_set & extra_set
    print(f"  Common Etso+Bedel pairs: {len(common)}")
    if common:
        print(f"  All common pairs (first 10):")
        for item in list(common)[:10]:
            print(f"    Etso: {item[0]}, Bedel: {item[1]}")
        # Also show regions for these common pairs to identify normalization issues
        print(f"\n  Common pairs region analysis (first 10):")
        for etso, bedel in list(common)[:10]:
            # Find missing records with this Etso+Bedel
            missing_records = [r for r in missing if r['Etso Kodu'] == etso and r['Dağıtım Bedeli(TL)'] == bedel]
            # Find extra records with this Etso+Bedel
            extra_records = [r for r in extra if r['Etso Kodu'] == etso and r['Dağıtım Bedeli(TL)'] == bedel]
            print(f"    Etso: {etso}, Bedel: {bedel}")
            print(f"      Missing ({len(missing_records)} records):")
            for m in missing_records:
                print(f"        Region: {m['Dağıtım Bölgesi']}, Customer: {m['Müşteri']}")
            print(f"      Extra ({len(extra_records)} records):")
            for e in extra_records:
                region = e.get('normalized_region', e.get('Dağıtım Bölgesi'))
                print(f"        Region: {region}, Customer: {e['Müşteri']}")
    
    # Debug: Çamlıbel EDAŞ region specific analysis
    print(f"\n=== DEBUG: Çamlıbel EDAŞ region analysis ===")
    missing_camlıbel = [r for r in missing if 'ÇAMLIBEL' in (r['Dağıtım Bölgesi'] or '').upper()]
    extra_camlıbel = [r for r in extra if 'ÇAMLIBEL' in (r.get('normalized_region', r.get('Dağıtım Bölgesi')) or '').upper()]
    print(f"  Missing in Çamlıbel: {len(missing_camlıbel)}")
    print(f"  Extra in Çamlıbel: {len(extra_camlıbel)}")
    if missing_camlıbel or extra_camlıbel:
        print(f"  Missing Çamlıbel records:")
        for r in missing_camlıbel:
            print(f"    Etso: {r['Etso Kodu']}, Customer: {r['Müşteri']}, Bedel: {r['Dağıtım Bedeli(TL)']}, Region: {r['Dağıtım Bölgesi']}")
        print(f"  Extra Çamlıbel records:")
        for r in extra_camlıbel:
            print(f"    Etso: {r['Etso Kodu']}, Customer: {r['Müşteri']}, Bedel: {r['Dağıtım Bedeli(TL)']}, Region: {r.get('normalized_region', r.get('Dağıtım Bölgesi'))}")

    # Eksik kayıtları detaylı analiz et
    if missing:
        print(f"\n=== EKSİK KAYITLAR ANALİZİ ===")
        print(f"Toplam eksik: {len(missing)}")
        
        # Region dağılımı
        region_dist = {}
        for row in missing:
            region = row['Dağıtım Bölgesi']
            region_dist[region] = region_dist.get(region, 0) + 1
        print(f"Region distribution:")
        for region, count in sorted(region_dist.items(), key=lambda x: x[1], reverse=True)[:10]:
            print(f"  {region}: {count}")
        
        # Bedel distribution
        bedel_ranges = {'0-1000': 0, '1000-10000': 0, '10000-100000': 0, '100000+': 0}
        for row in missing:
            bedel = row['Dağıtım Bedeli(TL)'] or 0
            if bedel < 1000:
                bedel_ranges['0-1000'] += 1
            elif bedel < 10000:
                bedel_ranges['1000-10000'] += 1
            elif bedel < 100000:
                bedel_ranges['10000-100000'] += 1
            else:
                bedel_ranges['100000+'] += 1
        print(f"Bedel distribution:")
        for range_name, count in bedel_ranges.items():
            print(f"  {range_name}: {count}")

        print(f"\n[MISSING] EKSİK KAYITLAR (İlk 10):")
        for row in missing[:10]:
            print(f"  {row['Dağıtım Bölgesi']} | Etso: {row['Etso Kodu']} | Müşteri: {row['Müşteri']} | Bedel: {row['Dağıtım Bedeli(TL)']}")

    # Fazla kayıtları detaylı analiz et
    if extra:
        print(f"\n=== FAZLA KAYITLAR ANALİZİ ===")
        print(f"Toplam fazla: {len(extra)}")
        
        # Region dağılımı
        region_dist = {}
        for row in extra:
            region = row.get('normalized_region', row.get('Dağıtım Bölgesi'))
            region_dist[region] = region_dist.get(region, 0) + 1
        print(f"Region distribution:")
        for region, count in sorted(region_dist.items(), key=lambda x: x[1], reverse=True)[:10]:
            print(f"  {region}: {count}")
        
        # Bedel distribution
        bedel_ranges = {'0-1000': 0, '1000-10000': 0, '10000-100000': 0, '100000+': 0}
        for row in extra:
            bedel = row['Dağıtım Bedeli(TL)'] or 0
            if bedel < 1000:
                bedel_ranges['0-1000'] += 1
            elif bedel < 10000:
                bedel_ranges['1000-10000'] += 1
            elif bedel < 100000:
                bedel_ranges['10000-100000'] += 1
            else:
                bedel_ranges['100000+'] += 1
        print(f"Bedel distribution:")
        for range_name, count in bedel_ranges.items():
            print(f"  {range_name}: {count}")
        
        print(f"\n[EXTRA] FAZLA KAYITLAR (İlk 10):")
        for row in extra[:10]:
            region = row.get('normalized_region', row.get('Dağıtım Bölgesi'))
            print(f"  {region} | Etso: {row['Etso Kodu']} | Müşteri: {row['Müşteri']} | Bedel: {row['Dağıtım Bedeli(TL)']}")

    # Debug: extracted_data'daki bedel değerlerini kontrol et
    print(f"\n[DEBUG] extracted_data bedel kontrolü:")
    print(f"  extracted_data uzunluğu: {len(extracted_data)}")
    
    # Tüm bedel değerlerini topla ve istatistiksel analiz yap
    bedel_values = [r.get('Dağıtım Bedeli(TL)') for r in extracted_data]
    bedel_values = [v for v in bedel_values if v is not None]
    
    print(f"  Geçerli bedel sayısı: {len(bedel_values)}")
    print(f"  Min bedel: {min(bedel_values):,.2f}")
    print(f"  Max bedel: {max(bedel_values):,.2f}")
    print(f"  Ortalama: {sum(bedel_values)/len(bedel_values):,.2f}")
    
    # Büyük bedeller (>1M)
    large_bedels = [(i, v) for i, v in enumerate(extracted_data) if v.get('Dağıtım Bedeli(TL)') and abs(v.get('Dağıtım Bedeli(TL)')) > 1_000_000]
    print(f"\n  Büyük bedeller (>1M): {len(large_bedels)}")
    for i, v in large_bedels[:20]:
        print(f"    [{i}] {v.get('Dağıtım Bedeli(TL)'):,.2f} - Region: {v.get('Dağıtım Bölgesi')} - Customer: {v.get('Müşteri')}")

    # Yeşilırmak EDAŞ specific logging for large values in extracted data
    print(f"\n[DEBUG] Yeşilırmak EDAŞ bedel kontrolü:")
    yesilirmak_large = [(i, v) for i, v in enumerate(extracted_data) if v.get('Dağıtım Bedeli(TL)') and abs(v.get('Dağıtım Bedeli(TL)')) > 1_000_000 and v.get('Dağıtım Bölgesi', '').lower().find('yeşilırmak') >= 0]
    print(f"  Yeşilırmak büyük bedeller (>1M): {len(yesilirmak_large)}")
    for i, v in yesilirmak_large:
        print(f"    Etso: {v.get('Etso Kodu')}, Customer: {v.get('Müşteri')}, Region: {v.get('Dağıtım Bölgesi')}, Bedel: {v.get('Dağıtım Bedeli(TL)'):,.2f} TL")

    # Yeşilırmak negative values
    yesilirmak_negative = [(i, v) for i, v in enumerate(extracted_data) if v.get('Dağıtım Bedeli(TL)') and v.get('Dağıtım Bedeli(TL)') < 0 and v.get('Dağıtım Bölgesi', '').lower().find('yeşilırmak') >= 0]
    print(f"  Yeşilırmak negative bedeller: {len(yesilirmak_negative)}")
    for i, v in yesilirmak_negative:
        print(f"    Etso: {v.get('Etso Kodu')}, Customer: {v.get('Müşteri')}, Region: {v.get('Dağıtım Bölgesi')}, Bedel: {v.get('Dağıtım Bedeli(TL)'):,.2f} TL")

    # Toplam bedel karşılaştırması
    ref_total = sum((r['Dağıtım Bedeli(TL)'] or 0) for r in reference_data)
    ext_total = sum((r['Dağıtım Bedeli(TL)'] or 0) for r in extracted_data)

    print(f"\n=== TOPLAM BEDEL KARŞILAŞTIRMASI ===")
    print(f"Referans Toplam (orijinal): {ref_total_original:,.2f} TL")
    print(f"Referans Toplam (comparison): {ref_total:,.2f} TL")
    print(f"Çıkarılan Toplam (non-negative): {ext_total:,.2f} TL")
    print(f"Bedel Farkı: {abs(ref_total - ext_total):,.2f} TL")


def main():
    """Ana fonksiyon"""
    print("=" * 70)
    print("KAPSAMLI FATURA VERİ ÇIKARMA VE KARŞILAŞTIRMA")
    print("=" * 70)

    # Veriyi çıkar
    extracted_data = process_all_regions()

    # Region name normalization uygula (reference data ile uyumlu hale getir)
    extracted_data_normalized = []
    for row in extracted_data:
        normalized_row = row.copy()
        # Extraction region name'ini reference formatına çevir
        normalized_region = normalize_region_name(row.get('Dağıtım Bölgesi'), to_format='reference')
        normalized_row['Dağıtım Bölgesi'] = normalized_region
        extracted_data_normalized.append(normalized_row)

    # Excel'e kaydet (normalized extracted data - reference formatında region isimleri ile)
    save_to_excel(extracted_data_normalized, OUTPUT_FILE)
    save_to_excel(extracted_data_normalized, os.path.join(BASE_PATH, 'Nihai_Birlestirilmis_Faturalar.xlsx'))

    # Referansla karşılaştır (normalized extracted data kullan)
    compare_with_reference(extracted_data_normalized, REFERENCE_FILE)


if __name__ == '__main__':
    main()
