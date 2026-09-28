# -*- coding: utf-8 -*-
"""
Bölge ismi normalizasyonu ve ETSO kodu standartlaştırma modülü.
"""

import re

from config.mappings import (
    REGION_NORMALIZATION_MAPPING,
    REFERENCE_TO_EXTRACTION_MAPPING,
)

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
