#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Dağıtımın Kestiği Faturalar Özet.xlsx ile çıkarılan veriler arasındaki
eksik, fazla ve farklı verilerin detaylı listesini çıkartır.
"""
from pathlib import Path

import os
import sys
import openpyxl
import pandas as pd
from collections import Counter

# Add parent directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from new_mapping_parser import parse_mapping_file, apply_note_rules, excel_column_letter_to_index

# Sabitler
BASE_PATH = Path(__file__).resolve().parents[1]
MAPPING_FILE = os.path.join(BASE_PATH, 'SKF Başlıkları Güncel.xlsx')
REFERENCE_FILE = os.path.join(BASE_PATH, 'Dağıtımın Kestiği Faturalar Özet.xlsx')
EXTRACTED_FILE = os.path.join(BASE_PATH, 'Çıkarılan_Veriler.xlsx')
OUTPUT_FILE = os.path.join(BASE_PATH, 'Farklı_Veriler_Raporu.xlsx')


def normalize_etso_kodu(etso_kodu):
    """Etso Kodu'yu standartlaştırmak için fonksiyon"""
    if etso_kodu is None:
        return None
    
    if isinstance(etso_kodu, (int, float)):
        if etso_kodu == int(etso_kodu):
            return str(int(etso_kodu))
        return str(etso_kodu)
    
    etso_str = str(etso_kodu).strip()
    
    if etso_str.isdigit():
        return etso_str.lstrip('0') or '0'
    
    # Format "40Z0000006543210G" gibi başlıyorsa
    import re
    match = re.match(r'^\d{2}Z0*(\d{7})', etso_str, re.IGNORECASE)
    if match:
        result = match.group(1).lstrip('0')
        return result if result else '0'
    
    match = re.search(r'\d{7}', etso_str)
    if match:
        return match.group(0)
    
    digits = re.findall(r'\d+', etso_str)
    if digits:
        all_digits = ''.join(digits)
        if len(all_digits) >= 7:
            return all_digits[-7:].lstrip('0') or '0'
        else:
            return all_digits.lstrip('0') or '0'
    
    return etso_str


def normalize_musteri(musteri):
    """Müşteri adını standartlaştırmak için fonksiyon"""
    if musteri is None:
        return ''
    return str(musteri).strip().upper().replace('İ', 'I').replace('Ç', 'C').replace('Ş', 'S').replace('Ö', 'O').replace('Ü', 'U').replace('Ğ', 'G')


def normalize_region_name(region_name, to_format='reference'):
    """Region ismini standartlaştırmak için fonksiyon"""
    if region_name is None:
        return None

    region_name = str(region_name).strip()
    
    # Extraction formatından reference formatına çevir
    REGION_NORMALIZATION_MAPPING = {
        # Large letter formats to proper format
        "AKEDAŞ": "AKEDAŞ ( Göksu EDAŞ )",
        "ADM EDAŞ": "Aydem EDAŞ",
        "AYEDAŞ": "Ayedaş",
        "DİCLE EDAŞ": "Dicle Edaş",
        "FIRAT EDAŞ": "Fırat Edaş",
        "KAYSERI VE CIVARI": "Kayseri EDAŞ",
        "KCETAŞ": "Kayseri EDAŞ",
        "VANGÖLÜ EDAŞ": "Vangölü Edaş",
        "BOĞAZİÇİ EDAŞ": "Boğaziçi EDAŞ",
        "BAŞKENT EDAŞ": "Başkent EDAŞ",
        "SEDAŞ": "Sakarya EDAŞ",
        "TREDAŞ": "Trakya EDAŞ",
        "UEDAŞ": "Uludağ EDAŞ",
        "ÇAMLIBEL EDAŞ": "Çamlıbel EDAŞ",
        "YEŞİLIRMAK EDAŞ": "Yeşilırmak EDAŞ",
        "GDZ EDAŞ": "Gediz EDAŞ",
        "MERAM": "Meram EDAŞ",
        "OEDAŞ": "Osmangazi EDAŞ",
        "TOROSLAR EDAŞ": "Toroslar EDAŞ",
        "ÇORUH EDAŞ": "Çoruh EDAŞ",
        "ARAS EDAŞ": "Aras EDAŞ",
        # Mixed case formats to proper format (already normalized but not in mapping)
        "Ayedaş": "Ayedaş",
        "Dicle Edaş": "Dicle Edaş",
        "Fırat Edaş": "Fırat Edaş",
        "Vangölü Edaş": "Vangölü Edaş",
        "Boğaziçi EDAŞ": "Boğaziçi EDAŞ",
        "Başkent EDAŞ": "Başkent EDAŞ",
        "Sakarya EDAŞ": "Sakarya EDAŞ",
        "Trakya EDAŞ": "Trakya EDAŞ",
        "Uludağ EDAŞ": "Uludağ EDAŞ",
        "Çamlıbel EDAŞ": "Çamlıbel EDAŞ",
        "Yeşilırmak EDAŞ": "Yeşilırmak EDAŞ",
        "Gediz EDAŞ": "Gediz EDAŞ",
        "Meram EDAŞ": "Meram EDAŞ",
        "Osmangazi EDAŞ": "Osmangazi EDAŞ",
        "Kayseri EDAŞ": "Kayseri EDAŞ",
        "Toroslar EDAŞ": "Toroslar EDAŞ",
        "Akdeniz EDAŞ": "Akdeniz EDAŞ",
        "Aydem EDAŞ": "Aydem EDAŞ",
    }

    # Try to find in mapping (case insensitive)
    region_upper = region_name.upper()
    if region_upper in REGION_NORMALIZATION_MAPPING:
        return REGION_NORMALIZATION_MAPPING[region_upper]
    
    # If not in mapping, return as-is (already normalized or unknown)
    return region_name


def load_reference_data():
    """Referans dosyasını yükle - Etso Kodu bazında toplama yapar"""
    wb = openpyxl.load_workbook(REFERENCE_FILE, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    headers = rows[0]

    # Etso Kodu bazında toplama için sözlük
    aggregated_data = {}
    
    for row in rows[1:]:
        if not row or all(v is None for v in row):
            continue

        row_dict = {}
        for i, header in enumerate(headers):
            if i < len(row):
                row_dict[header] = row[i]
            else:
                row_dict[header] = None

        # Normalize region name
        if 'Dağıtım Bölgesi' in row_dict:
            row_dict['Dağıtım Bölgesi'] = normalize_region_name(row_dict['Dağıtım Bölgesi'])

        # Normalize Etso Kodu
        etso = None
        if 'Etso Kodu' in row_dict:
            etso = normalize_etso_kodu(row_dict['Etso Kodu'])
            row_dict['Etso Kodu'] = etso

        # Normalize Müşteri
        musteri_normalized = ''
        if 'Müşteri' in row_dict:
            musteri_normalized = normalize_musteri(row_dict['Müşteri'])
            row_dict['Müşteri Normalized'] = musteri_normalized

        # Etso Kodu bazında toplama
        if etso:
            if etso not in aggregated_data:
                aggregated_data[etso] = {
                    'Etso Kodu': etso,
                    'Dağıtım Bölgesi': row_dict.get('Dağıtım Bölgesi'),
                    'Müşteri Normalized': musteri_normalized,
                    'Dağıtım Bedeli(TL)': 0,
                    'original_rows': []
                }
            
            bedel = (row_dict.get('Dağıtım Bedeli(TL)') or 0)
            aggregated_data[etso]['Dağıtım Bedeli(TL)'] += bedel
            aggregated_data[etso]['original_rows'].append(row_dict)

    return list(aggregated_data.values())


def load_extracted_data():
    """Çıkarılan verileri yükle - Etso Kodu bazında toplama yapar"""
    if not os.path.exists(EXTRACTED_FILE):
        print(f"Hata: {EXTRACTED_FILE} dosyası bulunamadı!")
        return []

    wb = openpyxl.load_workbook(EXTRACTED_FILE, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    headers = rows[0]

    # Etso Kodu bazında toplama için sözlük
    aggregated_data = {}
    
    for row in rows[1:]:
        if not row or all(v is None for v in row):
            continue

        row_dict = {}
        for i, header in enumerate(headers):
            if i < len(row):
                row_dict[header] = row[i]
            else:
                row_dict[header] = None

        # Negative değerleri filtrele
        bedel = (row_dict.get('Dağıtım Bedeli(TL)') or 0)
        if bedel < 0:
            continue

        # Normalize region name
        if 'Dağıtım Bölgesi' in row_dict:
            row_dict['Dağıtım Bölgesi'] = normalize_region_name(row_dict['Dağıtım Bölgesi'])

        # Normalize Etso Kodu
        etso = None
        if 'Etso Kodu' in row_dict:
            etso = normalize_etso_kodu(row_dict['Etso Kodu'])
            row_dict['Etso Kodu'] = etso

        # Normalize Müşteri
        musteri_normalized = ''
        if 'Müşteri' in row_dict:
            musteri_normalized = normalize_musteri(row_dict['Müşteri'])
            row_dict['Müşteri Normalized'] = musteri_normalized

        # Etso Kodu bazında toplama
        if etso:
            if etso not in aggregated_data:
                aggregated_data[etso] = {
                    'Etso Kodu': etso,
                    'Dağıtım Bölgesi': row_dict.get('Dağıtım Bölgesi'),
                    'Müşteri Normalized': musteri_normalized,
                    'Dağıtım Bedeli(TL)': 0,
                    'original_rows': []
                }
            
            aggregated_data[etso]['Dağıtım Bedeli(TL)'] += bedel
            aggregated_data[etso]['original_rows'].append(row_dict)

    return list(aggregated_data.values())


def find_discrepancies():
    """Eksik, fazla ve farklı verileri bul"""
    
    print("Referans dosyası yükleniyor...")
    reference_data = load_reference_data()
    print(f"Referans dosyası: {len(reference_data)} satır")
    
    print("Çıkarılan veriler yükleniyor...")
    extracted_data = load_extracted_data()
    print(f"Çıkarılan veriler: {len(extracted_data)} satır")
    
    # Matching key oluşturulması (Etso + Bedel + Region)
    def create_matching_key(row):
        etso = row.get('Etso Kodu', '')
        bedel = (row.get('Dağıtım Bedeli(TL)') or 0)
        region = row.get('Dağıtım Bölgesi', '')
        return (etso, bedel, region)
    
    # Full key (Etso + Müşteri + Bedel + Region)
    def create_full_key(row):
        etso = row.get('Etso Kodu', '')
        musteri = row.get('Müşteri Normalized', '')
        bedel = (row.get('Dağıtım Bedeli(TL)') or 0)
        region = row.get('Dağıtım Bölgesi', '')
        return (etso, musteri, bedel, region)
    
    # Reference data için indexler
    reference_matching_index = {}
    reference_full_index = {}

    for idx, row in enumerate(reference_data):
        matching_key = create_matching_key(row)
        full_key = create_full_key(row)

        if matching_key not in reference_matching_index:
            reference_matching_index[matching_key] = []
        reference_matching_index[matching_key].append({'row': row, 'index': idx, 'used': False})

        # Boş müşteri adı durumu için özel handling
        etso, musteri, bedel, region = full_key
        if not musteri or musteri.strip() == '':
            musteri = '(BOŞ)'
            full_key = (etso, musteri, bedel, region)
        
        if full_key not in reference_full_index:
            reference_full_index[full_key] = []
        reference_full_index[full_key].append({'row': row, 'index': idx, 'used': False})

    # Extracted data için indexler
    extracted_matching_index = {}
    extracted_full_index = {}

    for idx, row in enumerate(extracted_data):
        matching_key = create_matching_key(row)
        full_key = create_full_key(row)

        if matching_key not in extracted_matching_index:
            extracted_matching_index[matching_key] = []
        extracted_matching_index[matching_key].append({'row': row, 'index': idx, 'used': False})

        # Boş müşteri adı durumu için özel handling
        etso, musteri, bedel, region = full_key
        if not musteri or musteri.strip() == '':
            musteri = '(BOŞ)'
            full_key = (etso, musteri, bedel, region)
        
        if full_key not in extracted_full_index:
            extracted_full_index[full_key] = []
        extracted_full_index[full_key].append({'row': row, 'index': idx, 'used': False})
    
    # Counters
    reference_matching_counter = Counter(create_matching_key(r) for r in reference_data)
    extracted_matching_counter = Counter(create_matching_key(r) for r in extracted_data)
    
    # All matching keys
    all_matching_keys = set(reference_matching_index.keys()) | set(extracted_matching_index.keys())
    
    # Eksik ve fazla kayıtları bul
    missing = []
    extra = []
    matched_count = 0
    
    for matching_key in all_matching_keys:
        ref_count = reference_matching_counter.get(matching_key, 0)
        ext_count = extracted_matching_counter.get(matching_key, 0)
        
        matched = min(ref_count, ext_count)
        matched_count += matched
        
        # Eksik kayıtlar
        missing_count = ref_count - matched
        if missing_count > 0 and matching_key in reference_matching_index:
            count = 0
            for item in reference_matching_index[matching_key]:
                if not item['used'] and count < missing_count:
                    missing.append(item['row'])
                    item['used'] = True
                    count += 1
        
        # Fazla kayıtlar
        extra_count = ext_count - matched
        if extra_count > 0 and matching_key in extracted_matching_index:
            count = 0
            for item in extracted_matching_index[matching_key]:
                if not item['used'] and count < extra_count:
                    extra.append(item['row'])
                    item['used'] = True
                    count += 1
    
    # Eşleşen kayıtlar için full key kontrolü (müşteri farkları için)
    # Eşleşen matching key'leri bul
    matched_matching_keys = set()
    for matching_key in all_matching_keys:
        ref_count = reference_matching_counter.get(matching_key, 0)
        ext_count = extracted_matching_counter.get(matching_key, 0)
        if min(ref_count, ext_count) > 0:
            matched_matching_keys.add(matching_key)
    
    # Eşleşen kayıtlarda müşterileri kontrol et
    different_customers = []
    for matching_key in matched_matching_keys:
        ref_items = [item for item in reference_matching_index.get(matching_key, [])]
        ext_items = [item for item in extracted_matching_index.get(matching_key, [])]
        
        # Customer name'leri karşılaştır
        ref_customers = set(item['row'].get('Müşteri Normalized', '') for item in ref_items)
        ext_customers = set(item['row'].get('Müşteri Normalized', '') for item in ext_items)
        
        if ref_customers != ext_customers:
            # Müşteri farkı var
            for item in ref_items:
                if item['row'].get('Müşteri Normalized', '') not in ext_customers:
                    different_customers.append({
                        'source': 'referans',
                        'row': item['row'],
                        'missing_customer': True
                    })
            for item in ext_items:
                if item['row'].get('Müşteri Normalized', '') not in ref_customers:
                    different_customers.append({
                        'source': 'ekstrakt',
                        'row': item['row'],
                        'missing_customer': False
                    })
    
    print(f"\n=== SONUÇLAR ===")
    print(f"Toplam eşleşen: {matched_count}")
    print(f"Toplam eksik: {len(missing)}")
    print(f"Toplam fazla: {len(extra)}")
    print(f"Müşteri farkı: {len(different_customers)}")
    
    # Region distribution
    print(f"\n=== EKSİK KAYITLAR - REGION DAĞILIMI ===")
    region_dist_missing = {}
    for row in missing:
        region = row.get('Dağıtım Bölgesi', 'UNKNOWN')
        region_dist_missing[region] = region_dist_missing.get(region, 0) + 1
    for region, count in sorted(region_dist_missing.items(), key=lambda x: x[1], reverse=True):
        print(f"  {region}: {count}")
    
    print(f"\n=== FAZLA KAYITLAR - REGION DAĞILIMI ===")
    region_dist_extra = {}
    for row in extra:
        region = row.get('Dağıtım Bölgesi', 'UNKNOWN')
        region_dist_extra[region] = region_dist_extra.get(region, 0) + 1
    for region, count in sorted(region_dist_extra.items(), key=lambda x: x[1], reverse=True):
        print(f"  {region}: {count}")
    
    # Toplam bedel
    ref_total = sum((r.get('Dağıtım Bedeli(TL)') or 0) for r in reference_data)
    ext_total = sum((r.get('Dağıtım Bedeli(TL)') or 0) for r in extracted_data)
    missing_total = sum((r.get('Dağıtım Bedeli(TL)') or 0) for r in missing)
    extra_total = sum((r.get('Dağıtım Bedeli(TL)') or 0) for r in extra)
    
    print(f"\n=== TOPLAM BEDEL ===")
    print(f"Referans: {ref_total:,.2f} TL")
    print(f"Çıkarılan: {ext_total:,.2f} TL")
    print(f"Eksik kayıtların toplamı: {missing_total:,.2f} TL")
    print(f"Fazla kayıtların toplamı: {extra_total:,.2f} TL")
    
    # Excel'e kaydet
    save_results_to_excel(missing, extra, different_customers, reference_data, extracted_data)
    
    return missing, extra, different_customers


def save_results_to_excel(missing, extra, different_customers, reference_data, extracted_data):
    """Sonuçları Excel'e kaydet"""
    
    print(f"\nSonuçlar Excel'e kaydediliyor: {OUTPUT_FILE}")
    
    wb = openpyxl.Workbook()
    
    # Sayfa 1: Eksik Kayıtlar
    ws1 = wb.active
    ws1.title = "Eksik Kayıtlar"
    
    headers = ['Dağıtım Bölgesi', 'Etso Kodu', 'Müşteri', 'Dağıtım Bedeli(TL)', 'Müşteri Normalized']
    ws1.append(headers)
    
    for row in missing:
        ws1.append([
            row.get('Dağıtım Bölgesi', ''),
            row.get('Etso Kodu', ''),
            row.get('Müşteri', ''),
            row.get('Dağıtım Bedeli(TL)', ''),
            row.get('Müşteri Normalized', '')
        ])
    
    # Sayfa 2: Fazla Kayıtlar
    ws2 = wb.create_sheet(title="Fazla Kayıtlar")
    ws2.append(headers)
    
    for row in extra:
        ws2.append([
            row.get('Dağıtım Bölgesi', ''),
            row.get('Etso Kodu', ''),
            row.get('Müşteri', ''),
            row.get('Dağıtım Bedeli(TL)', ''),
            row.get('Müşteri Normalized', '')
        ])
    
    # Sayfa 3: Müşteri Farkı
    if different_customers:
        ws3 = wb.create_sheet(title="Müşteri Farkı")
        ws3.append(['Kaynak', 'Dağıtım Bölgesi', 'Etso Kodu', 'Müşteri', 'Dağıtım Bedeli(TL)', 'Müşteri Normalized', 'Eksik Müşteri mi'])
        
        for item in different_customers:
            row = item['row']
            ws3.append([
                item['source'],
                row.get('Dağıtım Bölgesi', ''),
                row.get('Etso Kodu', ''),
                row.get('Müşteri', ''),
                row.get('Dağıtım Bedeli(TL)', ''),
                row.get('Müşteri Normalized', ''),
                item['missing_customer']
            ])
    
    # Sayfa 4: Referans Data (özet)
    ws4 = wb.create_sheet(title="Referans Data")
    ws4.append(['Dağıtım Bölgesi', 'Etso Kodu', 'Müşteri', 'Dağıtım Bedeli(TL)'])
    
    for row in reference_data[:1000]:  # İlk 1000 satır
        ws4.append([
            row.get('Dağıtım Bölgesi', ''),
            row.get('Etso Kodu', ''),
            row.get('Müşteri', ''),
            row.get('Dağıtım Bedeli(TL)', '')
        ])
    
    # Sayfa 5: Extracted Data (özet)
    ws5 = wb.create_sheet(title="Extracted Data")
    ws5.append(['Dağıtım Bölgesi', 'Etso Kodu', 'Müşteri', 'Dağıtım Bedeli(TL)'])
    
    for row in extracted_data[:1000]:  # İlk 1000 satır
        ws5.append([
            row.get('Dağıtım Bölgesi', ''),
            row.get('Etso Kodu', ''),
            row.get('Müşteri', ''),
            row.get('Dağıtım Bedeli(TL)', '')
        ])
    
    wb.save(OUTPUT_FILE)
    print(f"Sonuçlar kaydedildi!")


if __name__ == '__main__':
    missing, extra, different_customers = find_discrepancies()
    
    # Detaylı rapor
    print(f"\n=== DETAYLI RAPOR ===")
    
    print(f"\n--- EKSİK KAYITLAR (İlk 20) ---")
    for i, row in enumerate(missing[:20], 1):
        print(f"{i}. {row.get('Dağıtım Bölgesi', '')} | {row.get('Etso Kodu', '')} | {row.get('Müşteri', '')} | {row.get('DağıAPSHOT Bedeli(TL)', '')}")
    if len(missing) > 20:
        print(f"... ve {len(missing) - 20} daha")
    
    print(f"\n--- FAZLA KAYITLAR (İlk 20) ---")
    for i, row in enumerate(extra[:20], 1):
        print(f"{i}. {row.get('Dağıtım Bölgesi', '')} | {row.get('Etso Kodu', '')} | {row.get('Müşteri', '')} | {row.get('Dağıtım Bedeli(TL)', '')}")
    if len(extra) > 20:
        print(f"... ve {len(extra) - 20} daha")
    
    print(f"\n--- MÜŞTERİ FARKI (İlk 20) ---")
    for i, item in enumerate(different_customers[:20], 1):
        row = item['row']
        print(f"{i}. [{item['source']}] {row.get('Dağıtım Bölgesi', '')} | {row.get('Etso Kodu', '')} | {row.get('Müşteri', '')} | {row.get('Dağıtım Bedeli(TL)', '')}")
    if len(different_customers) > 20:
        print(f"... ve {len(different_customers) - 20} daha")
    
    print(f"\nRapor tamamlandı. Excel dosyası: {OUTPUT_FILE}")
