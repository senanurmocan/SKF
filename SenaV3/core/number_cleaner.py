# -*- coding: utf-8 -*-
"""
Number cleaning and header normalization utilities extracted from extract_and_compare.py.

Includes:
- normalize_header: Normalizes and cleans header strings for matching.
- clean_turkish_number: Smart parsing of numbers in Turkish and US formats.
- _format_guc_kw: Formats Güç kW values to 2-decimal floats.
"""

import re


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
