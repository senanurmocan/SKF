# -*- coding: utf-8 -*-
"""
Region and field mapping dictionaries for EDAS Bill Extraction & Analysis System.
"""

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

# ── Bölge-spesifik başlık override ve alan sinonimler ─────────────────────
# Öncelik: config/header_overrides.json > aşağıdaki hardcoded fallback
# JSON dosyası kod değişikliği gerektirmeden mapping güncellemeye izin verir.
import json as _json
import os as _os

_OVERRIDES_JSON = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'header_overrides.json')

# Hardcoded fallback — JSON yoksa veya okunamazsa bunlar kullanılır
_SPECIAL_HEADER_MAPPING_FALLBACK = {
    'AKEDAŞ': {
        'etso': 'Id',
        'musteri': None,
    },
    'Çamlıbel EDAŞ': {
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
        'güç_bedeli': 'Güç Bedeli (TL)',
        'reaktif': 'Reaktif Bedeli (TL)',
        'reaktif_tenzil': 'İlk Reaktif İade Tutar',
    },
    'Toroslar EDAŞ': {
        'ag_og': 'Gerilim Seviyesi',
    },
    'Meram EDAŞ': {
        'reaktif': 'REAKTİF TÜKETİM',
    },
    'MERAM EDAŞ': {
        'reaktif': 'REAKTİF TÜKETİM',
    },
    'MERAM': {
        'reaktif': 'REAKTİF TÜKETİM',
    },
    'Sakarya EDAŞ': {
        'reaktif': 'Reaktif Bedel (TL)',
    },
    'SAKARYA EDAŞ': {
        'reaktif': 'Reaktif Bedel (TL)',
    },
}

_FIELD_SYNONYMS_FALLBACK = {
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


def _load_overrides():
    """JSON config dosyasından override mapping'leri yükle."""
    if _os.path.exists(_OVERRIDES_JSON):
        try:
            with open(_OVERRIDES_JSON, 'r', encoding='utf-8') as f:
                data = _json.load(f)
            return (
                data.get('SPECIAL_HEADER_MAPPING', _SPECIAL_HEADER_MAPPING_FALLBACK),
                data.get('FIELD_SYNONYMS', _FIELD_SYNONYMS_FALLBACK),
            )
        except Exception:
            pass
    return _SPECIAL_HEADER_MAPPING_FALLBACK, _FIELD_SYNONYMS_FALLBACK


SPECIAL_HEADER_MAPPING, FIELD_SYNONYMS = _load_overrides()

