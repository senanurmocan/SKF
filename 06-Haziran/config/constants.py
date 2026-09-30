# -*- coding: utf-8 -*-
"""
Constants and configuration for SKF Bill Extraction & Comparison System.
"""

import os
from pathlib import Path

# Sabitler — BASE_PATH: config/ içinden bir üst dizine (06-Haziran/) çıkılır
BASE_PATH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAPPING_FILE = os.path.join(BASE_PATH, 'SKF Başlıkları.xlsx')
REFERENCE_FILE = os.path.join(BASE_PATH, 'Dağıtımın Kestiği Faturalar Özet.xlsx')
OUTPUT_FILE = os.path.join(BASE_PATH, 'Çıkarılan_Veriler.xlsx')
SUPPORTED_INPUT_EXTENSIONS = ('.xls', '.xlsx', '.xml', '.html')

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
