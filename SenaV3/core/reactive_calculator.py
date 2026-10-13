# -*- coding: utf-8 -*-
"""
Reaktif ve İlk Reaktif bedel hesaplama modülü.

Bu modül, 4 farklı format okuyucusu (xlsx, xls, html, xml) tarafından
ortaklaşa kullanılan reaktif hesaplama mantığını TEK KAYNAKTA tutar.

Notlar 3 #10: İlk Reaktif alan grubu ayrı toplanır.
Sakarya EDAŞ: 'X' değeri özel olarak handle edilir.
"""

from core.number_cleaner import clean_turkish_number


def calculate_reactive_totals(cell_values, indices, reaktif_fields, ilk_reaktif_fields, region_name=None):
    """Reaktif ve İlk Reaktif toplam değerlerini hesapla.

    Bu fonksiyon 4 format okuyucunun (xlsx, xls, html, xml) ortak
    reaktif hesaplama mantığını içerir. DRY prensibine uygun olarak
    tek kaynaktan çağrılır.

    Args:
        cell_values: Satırdaki hücre değerleri (list)
        indices: Sütun indeksleri sözlüğü (dict)
        reaktif_fields: Reaktif alan isimleri listesi (list)
        ilk_reaktif_fields: İlk Reaktif alan isimleri listesi (list)
        region_name: Bölge adı (opsiyonel, hata raporlama için)

    Returns:
        tuple: (reaktif_toplam, ilk_reaktif_value, sakarya_x)
            - reaktif_toplam (float): Reaktif bedel toplamı
            - ilk_reaktif_value (float|str): İlk Reaktif toplamı veya 'X'
            - sakarya_x (bool): Sakarya 'X' değeri tespit edildi mi
    """
    # Reaktif field summing
    reaktif_toplam = 0
    sakarya_x = False
    for f_key in reaktif_fields:
        idx = indices.get(f_key, -1)
        if idx >= 0 and idx < len(cell_values):
            val = cell_values[idx]
            if val is not None:
                if str(val).strip().upper() == 'X':
                    sakarya_x = True
                else:
                    reaktif_toplam += clean_turkish_number(
                        val, region_name=region_name, field_name='reaktif'
                    )

    # İlk Reaktif hesaplama (Notlar 3 #10)
    ilk_reaktif_toplam = 0
    for f_key in ilk_reaktif_fields:
        idx = indices.get(f_key, -1)
        if idx >= 0 and idx < len(cell_values):
            val = cell_values[idx]
            if val is not None:
                if str(val).strip().upper() == 'X':
                    sakarya_x = True
                else:
                    ilk_reaktif_toplam += clean_turkish_number(
                        val, region_name=region_name, field_name='reaktif'
                    )

    ilk_reaktif_value = 'X' if sakarya_x else ilk_reaktif_toplam
    return reaktif_toplam, ilk_reaktif_value, sakarya_x
