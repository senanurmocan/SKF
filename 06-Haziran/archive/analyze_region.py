#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Bölgesel analiz scripti - eksik ve fazla kayıtları Etso+Bedel bazlı inceler
"""
from pathlib import Path

import os
import openpyxl
from collections import Counter

BASE_PATH = Path(__file__).resolve().parents[1]
EXTRACTED_FILE = os.path.join(BASE_PATH, 'Çıkarılan_Veriler.xlsx')
REFERENCE_FILE = os.path.join(BASE_PATH, 'Dağıtımın Kestiği Faturalar Özet.xlsx')

def normalize_region_name(region_name):
    """Extraction formatından reference formatına çevir"""
    if region_name is None:
        return None
    
    region_name = str(region_name).strip()
    
    REGION_NORMALIZATION_MAPPING = {
        "AKEDAŞ": "AKEDAŞ ( Göksu EDAŞ )",
        "ADM EDAŞ": "Aydem EDAŞ",
        "AYEDAŞ": "Ayedaş",
        "Dicle EDAŞ": "Dicle Edaş",
        "Fırat EDAŞ": "Fırat Edaş",
        "Kayseri ve Civarı": "Kayseri EDAŞ",
        "Vangölü EDAŞ": "Vangölü Edaş",
        "Boğaziçi EDAŞ": "Boğaziçi EDAŞ",
        "Başkent EDAŞ": "Başkent EDAŞ",
        "Sakarya EDAŞ": "Sakarya EDAŞ",
        "Trakya EDAŞ": "Trakya EDAŞ",
        "Uludağ EDAŞ": "Uludağ EDAŞ",
        "Çamlıbel EDAŞ": "Çamlıbel EDAŞ",
        "Vangölü EDAŞ": "Vangölü Edaş",
        "Yeşilırmak EDAŞ": "Yeşilırmak EDAŞ",
        "Gediz EDAŞ": "Gediz EDAŞ",
        "Meram EDAŞ": "Meram EDAŞ",
        "Osmangazi EDAŞ": "Osmangazi EDAŞ",
        "Toroslar EDAŞ": "Toroslar EDAŞ",
        "Fırat EDAŞ": "Fırat Edaş",
        "Dicle EDAŞ": "Dicle Edaş",
        "Çoruh EDAŞ": "Çoruh EDAŞ",
        "Aras EDAŞ": "Aras EDAŞ",
    }
    
    return REGION_NORMALIZATION_MAPPING.get(region_name, region_name)

def normalize_etso_kodu(etso_kodu):
    """Etso Kodu'yu standartlaştırmak"""
    if etso_kodu is None:
        return None
    
    if isinstance(etso_kodu, (int, float)):
        if etso_kodu == int(etso_kodu):
            return str(int(etso_kodu))
        return str(etso_kodu)
    
    etso_str = str(etso_kodu).strip()
    
    import re
    if etso_str.isdigit():
        return etso_str.lstrip('0') or '0'
    
    digits = re.findall(r'\d+', etso_str)
    if digits:
        result = ''.join(digits).lstrip('0')
        return result if result else '0'
    
    return etso_str

def normalize_musteri(value):
    """Müşteri değerini normalize et"""
    if value is None:
        return ''
    return str(value).strip().upper()

def load_extracted_data():
    """Çıkarılan veriyi yükle"""
    wb = openpyxl.load_workbook(EXTRACTED_FILE)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    
    data = []
    for row in rows[1:]:
        if row[0] and row[0] != 'Dağıtım Bölgesi':
            data.append({
                'Dağıtım Bölgesi': row[0],
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
    
    return data

def load_reference_data():
    """Referans veriyi yükle"""
    wb = openpyxl.load_workbook(REFERENCE_FILE)
    ws = wb['Dağıtımın Kestiği']
    rows = list(ws.iter_rows(values_only=True))
    
    data = []
    for row in rows[1:]:
        if row[0] and row[0] != 'Dağıtım Bölgesi':
            # Reference region name'i extraction formatına çevir
            normalized_region = normalize_region_name(row[0])
            data.append({
                'Dağıtım Bölgesi': normalized_region,
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
    
    return data

def compare_data(extracted_data, reference_data):
    """Karşılaştırma yap - Etso+Müşteri+Bedel+Region key ile"""
    # Reference keys
    reference_keys = []
    for row in reference_data:
        etso = normalize_etso_kodu(row['Etso Kodu'])
        musteri = normalize_musteri(row['Müşteri'])
        bedel = row['Dağıtım Bedeli(TL)'] or 0
        region = normalize_region_name(row.get('Dağıtım Bölgesi'))
        key = (etso, musteri, bedel, region)
        reference_keys.append(key)
    
    # Extracted keys
    extracted_keys = []
    for row in extracted_data:
        etso = normalize_etso_kodu(row['Etso Kodu'])
        musteri = normalize_musteri(row['Müşteri'])
        bedel = row['Dağıtım Bedeli(TL)'] or 0
        region = normalize_region_name(row.get('Dağıtım Bölgesi'))
        key = (etso, musteri, bedel, region)
        extracted_keys.append(key)
    
    # Counters
    ref_counter = Counter(reference_keys)
    ext_counter = Counter(extracted_keys)
    
    all_keys = set(reference_keys) | set(extracted_keys)
    
    missing = []
    extra = []
    
    # Index dictionaries
    ref_index = {}
    for i, key in enumerate(reference_keys):
        if key not in ref_index:
            ref_index[key] = []
        ref_index[key].append(reference_data[i])
    
    ext_index = {}
    for i, key in enumerate(extracted_keys):
        if key not in ext_index:
            ext_index[key] = []
        ext_index[key].append(extracted_data[i])
    
    # Compare
    for key in all_keys:
        ref_count = ref_counter.get(key, 0)
        ext_count = ext_counter.get(key, 0)
        
        matched = min(ref_count, ext_count)
        
        # Missing
        missing_count = ref_count - matched
        if missing_count > 0 and key in ref_index:
            for i in range(min(missing_count, len(ref_index[key]))):
                missing.append(ref_index[key][i])
        
        # Extra
        extra_count = ext_count - matched
        if extra_count > 0 and key in ext_index:
            for i in range(min(extra_count, len(ext_index[key]))):
                extra.append(ext_index[key][i])
    
    return missing, extra

def analyze_region(missing, extra, region_name):
    """Belirli bir bölgeyi analiz et"""
    missing_region = [r for r in missing if region_name.upper() in (r['Dağıtım Bölgesi'] or '').upper()]
    extra_region = [r for r in extra if region_name.upper() in (r.get('normalized_region', r.get('Dağıtım Bölgesi')) or '').upper()]
    
    print(f"\n{'='*70}")
    print(f"ANALİZ: {region_name}")
    print(f"{'='*70}")
    print(f"Eksik kayıtlar: {len(missing_region)}")
    print(f"Fazla kayıtlar: {len(extra_region)}")
    
    if not missing_region and not extra_region:
        print("  Hiç kayıt yok")
        return
    
    # Etso+Bedel distribution for missing
    print(f"\n--- Eksik Kayıtlar (Etso+Bedel bazlı) ---")
    missing_pairs = [(r['Etso Kodu'], r['Dağıtım Bedeli(TL)']) for r in missing_region]
    missing_counter = Counter(missing_pairs)
    
    for (etso, bedel), count in missing_counter.most_common(20):
        print(f"  Etso: {etso}, Bedel: {bedel} - {count} kayıt")
        # Customer bilgisi
        customers = list(set([r['Müşteri'] for r in missing_region if r['Etso Kodu'] == etso and r['Dağıtım Bedeli(TL)'] == bedel]))
        print(f"    Müşteriler: {', '.join([str(c) for c in customers[:5]])}")
    
    # Etso+Bedel distribution for extra
    print(f"\n--- Fazla Kayıtlar (Etso+Bedel bazlı) ---")
    extra_pairs = [(r['Etso Kodu'], r['Dağıtım Bedeli(TL)']) for r in extra_region]
    extra_counter = Counter(extra_pairs)
    
    for (etso, bedel), count in extra_counter.most_common(20):
        print(f"  Etso: {etso}, Bedel: {bedel} - {count} kayıt")
        customers = list(set([r['Müşteri'] for r in extra_region if r['Etso Kodu'] == etso and r['Dağıtım Bedeli(TL)'] == bedel]))
        print(f"    Müşteriler: {', '.join([str(c) for c in customers[:5]])}")
    
    # Common pairs between missing and extra
    print(f"\n--- Ortak Etso+Bedel çiftleri (Eksik+Fazla) ---")
    missing_set = set(missing_pairs)
    extra_set = set(extra_pairs)
    common = missing_set & extra_set
    print(f"  Ortak çiftler: {len(common)}")
    
    if common:
        print(f"\n  Ortak çiftlerin detayı:")
        for etso, bedel in list(common)[:10]:
            missing_count = sum(1 for r in missing_region if r['Etso Kodu'] == etso and r['Dağıtım Bedeli(TL)'] == bedel)
            extra_count = sum(1 for r in extra_region if r['Etso Kodu'] == etso and r['Dağıtım Bedeli(TL)'] == bedel)
            print(f"    Etso: {etso}, Bedel: {bedel}")
            print(f"      Eksik: {missing_count}, Fazla: {extra_count}")
            # Customer comparison
            missing_customers = [r['Müşteri'] for r in missing_region if r['Etso Kodu'] == etso and r['Dağıtım Bedeli(TL)'] == bedel]
            extra_customers = [r['Müşteri'] for r in extra_region if r['Etso Kodu'] == etso and r['Dağıtım Bedeli(TL)'] == bedel]
            print(f"      Eksik Müşteriler: {', '.join([str(c) for c in list(set(missing_customers))[:3]])}")
            print(f"      Fazla Müşteriler: {', '.join([str(c) for c in list(set(extra_customers))[:3]])}")

def main():
    print("=== BÖLGESEL ANALİZ SCRIPTİ ===\n")
    
    # Verileri yükle
    print("Veriler yükleniyor...")
    extracted_data = load_extracted_data()
    reference_data = load_reference_data()
    print(f"Çıkarılan: {len(extracted_data)} satır")
    print(f"Referans: {len(reference_data)} satır")
    
    # Karşılaştır
    print("\nKarşılaştırılıyor...")
    missing, extra = compare_data(extracted_data, reference_data)
    print(f"Eksik kayıtlar: {len(missing)}")
    print(f"Fazla kayıtlar: {len(extra)}")
    
    #REGIONLERİ ANALİZ ET
    regions_to_analyze = [
        "Sakarya EDAŞ",
        "Boğaziçi EDAŞ", 
        "Uludağ EDAŞ",
        "AYEDAŞ",
        "Başkent EDAŞ",
        "ADM EDAŞ",
        "AKEDAŞ",
        "Trakya EDAŞ"
    ]
    
    for region in regions_to_analyze:
        analyze_region(missing, extra, region)

if __name__ == '__main__':
    import os
    main()
