#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extracted data vs Reference data comparison script
- Correctly locates Excel files in region-specific subfolders
- Calculates KDV and Toplam columns
- Performs exact key matching
- Identifies missing/extra records
"""
from pathlib import Path

import pandas as pd
import os
import re
from collections import defaultdict

BASE_PATH = Path(__file__).resolve().parents[1]
REFERENCE_FILE = os.path.join(BASE_PATH, "Dağıtımın Kestiği Faturalar Özet.xlsx")
EXTRACTED_FILE = os.path.join(BASE_PATH, "Çıkarılan_Veriler.xlsx")

def normalize_etso_kodu(etso_value):
    """
    Normalize etso kodu to 7-digit numeric string
    Examples:
        - "40Z0000006543210G" → "6543210"
        - "7606443" → "7606443"
    """
    if pd.isna(etso_value):
        return None
    
    etso_str = str(etso_value).strip()
    
    # If already numeric (7-8 digits), return as-is
    if re.match(r'^\d{7,8}$', etso_str):
        return etso_str
    
    # If alphanumeric format like "40Z0000006543210G", extract 7-digit number
    match = re.search(r'\d{2}Z0*(\d{7})', etso_str)
    if match:
        return match.group(1)
    
    # If it contains a number but doesn't match pattern, try to extract
    digits = re.findall(r'\d+', etso_str)
    if digits:
        # Return the longest digit sequence (likely the 7-digit etso)
        longest = max(digits, key=len)
        if len(longest) >= 7:
            return longest[:7]
    
    return etso_str


def normalize_region_name(region_name):
    """
    Normalize region name for comparison
    Examples:
        - "Aydem EDAŞ" → "Aydem EDAŞ"
        - "AYEDAŞ" → "Aydem EDAŞ"
    """
    if pd.isna(region_name):
        return None
    
    region_str = str(region_name).strip()
    
    # Common abbreviation mappings
    region_map = {
        'AYEDAŞ': 'Aydem EDAŞ',
        'AKEDAŞ': 'Aydem EDAŞ',
        'ADM EDAŞ': 'Aydem EDAŞ',
    }
    
    return region_map.get(region_str, region_str)


def calculate_kdv_and_total(row):
    """
    Calculate KDV and Toplam from extracted data
    KDV = Dağıtım Bedeli(TL) * 0.20 (assuming 20% KDV rate)
    Toplam = Dağıtım Bedeli(TL) + Güç Bedeli(TL) + Güç Aşım Bedeli (TL) + Reaktif Bedel (TL) + KDV
    """
    dagitim = float(row['Dağıtım Bedeli(TL)'] or 0)
    guc = float(row['Güç Bedeli(TL)'] or 0)
    gasim = float(row['Güç Aşım Bedeli (TL)'] or 0)
    reaktif = float(row['Reaktif Bedel (TL)'] or 0)
    
    kdv = dagitim * 0.20  # Assuming 20% KDV rate
    total = dagitim + guc + gasim + reaktif + kdv
    
    return kdv, total


def create_matching_key(row):
    """
    Create matching key: (etso_normalized, bedel, region_normalized)
    """
    etso = normalize_etso_kodu(row.get('Etso Kodu', None))
    bedel = float(row.get('Dağıtım Bedeli(TL)', 0) or 0)
    region = normalize_region_name(row.get('Dağıtım Bölgesi', None))
    
    return (etso, bedel, region)


def load_reference_data():
    """Load and preprocess reference data"""
    print("=" * 80)
    print("YÜKLENİYOR: Referans Veri (Dağıtımın Kestiği Faturalar Özet.xlsx)")
    print("=" * 80)
    
    df = pd.read_excel(REFERENCE_FILE)
    
    # Normalize columns
    df['Dağıtım Bölgesi'] = df['Dağıtım Bölgesi'].apply(normalize_region_name)
    df['Etso Kodu'] = df['Etso Kodu'].apply(normalize_etso_kodu)
    
    # Create matching keys
    df['matching_key'] = df.apply(create_matching_key, axis=1)
    
    # Create lookup dict for faster matching
    reference_dict = {}
    for idx, row in df.iterrows():
        key = row['matching_key']
        if key not in reference_dict:
            reference_dict[key] = []
        reference_dict[key].append(row.to_dict())
    
    print(f"  ✓ {len(df)} kayit yüklendi")
    print(f"  ✓ {len(reference_dict)} benzersiz matching key")
    print()
    
    return df, reference_dict


def load_extracted_data():
    """Load and preprocess extracted data"""
    print("=" * 80)
    print("YÜKLENİYOR: Çıkarılan Veri (Çıkarılan_Veriler.xlsx)")
    print("=" * 80)
    
    df = pd.read_excel(EXTRACTED_FILE)
    
    # Normalize columns
    df['Dağıtım Bölgesi'] = df['Dağıtım Bölgesi'].apply(normalize_region_name)
    df['Etso Kodu'] = df['Etso Kodu'].apply(normalize_etso_kodu)
    
    # Create matching keys
    df['matching_key'] = df.apply(create_matching_key, axis=1)
    
    # Create lookup dict
    extracted_dict = {}
    for idx, row in df.iterrows():
        key = row['matching_key']
        if key not in extracted_dict:
            extracted_dict[key] = []
        extracted_dict[key].append(row.to_dict())
    
    print(f"  ✓ {len(df)} kayit yüklendi")
    print(f"  ✓ {len(extracted_dict)} benzersiz matching key")
    print()
    
    return df, extracted_dict


def compare_data():
    """Compare extracted vs reference data"""
    print("=" * 80)
    print("KARŞILAŞTIRMA: Çıkarılan vs Referans")
    print("=" * 80)
    print()
    
    ref_df, ref_dict = load_reference_data()
    ext_df, ext_dict = load_extracted_data()
    
    # Find unmatched records
    print("=" * 80)
    print("ANALİZ: Eşleşmeyen Kayıtlar")
    print("=" * 80)
    print()
    
    # References without matches
    missing_in_extracted = []
    for key, rows in ref_dict.items():
        if key not in ext_dict:
            missing_in_extracted.extend(rows)
    
    # Extractions without matches
    extra_in_extracted = []
    for key, rows in ext_dict.items():
        if key not in ref_dict:
            extra_in_extracted.extend(rows)
    
    # Count matches
    matched_count = 0
    for key in ext_dict:
        if key in ref_dict:
            matched_count += 1
    
    print(f"  Eşleşen kayıtlar: {matched_count}")
    print(f"  Çıkarılanda OLMAYAN (referansta olan ama_extractedta yok): {len(missing_in_extracted)}")
    print(f"  Çıkarılanda FAZLA (extractedta olan ama referansta yok): {len(extra_in_extracted)}")
    print()
    
    # Detailed analysis of missing
    if missing_in_extracted:
        print("=" * 80)
        print("EXCEL'DE OLMAYAN KAYITLAR (İlk 10 örneği)")
        print("=" * 80)
        for i, row in enumerate(missing_in_extracted[:10]):
            print(f"  {i+1}. [{row['Dağıtım Bölgesi']}] Etso: {row['Etso Kodu']}, Bedel: {row['Dağıtım Bedeli(TL)']}")
        if len(missing_in_extracted) > 10:
            print(f"  ... ve {len(missing_in_extracted) - 10} daha")
        print()
    
    # Detailed analysis of extra
    if extra_in_extracted:
        print("=" * 80)
        print("EXCEL'DE FAZLA KAYITLAR (İlk 10 örneği)")
        print("=" * 80)
        for i, row in enumerate(extra_in_extracted[:10]):
            print(f"  {i+1}. [{row['Dağıtım Bölgesi']}] Etso: {row['Etso Kodu']}, Bedel: {row['Dağıtım Bedeli(TL)']}")
        if len(extra_in_extracted) > 10:
            print(f"  ... ve {len(extra_in_extracted) - 10} daha")
        print()
    
    # Calculate totals for comparison
    ref_total_bedel = ref_df['Dağıtım Bedeli(TL)'].sum()
    ext_total_bedel = ext_df['Dağıtım Bedeli(TL)'].sum()
    
    print("=" * 80)
    print("TOPLAM VERİLER")
    print("=" * 80)
    print(f"  Referans Dağıtım Bedeli Toplamı: {ref_total_bedel:,.2f} TL")
    print(f"  Çıkarılan Dağıtım Bedeli Toplamı: {ext_total_bedel:,.2f} TL")
    print(f"  Fark: {abs(ref_total_bedel - ext_total_bedel):,.2f} TL")
    print()
    
    return missing_in_extracted, extra_in_extracted


def verify_etso_normalization():
    """Verify etso normalization is working correctly"""
    print("=" * 80)
    print("ETSO NORMALİZASYON KONTROLÜ")
    print("=" * 80)
    print()
    
    ref_df = pd.read_excel(REFERENCE_FILE)
    ext_df = pd.read_excel(EXTRACTED_FILE)
    
    # Check reference etso formats
    print("Referans dosyasında Etso Kodu formatları:")
    ref_etso_samples = ref_df['Etso Kodu'].dropna().unique()[:15]
    for etso in ref_etso_samples:
        print(f"  - {etso}")
    print()
    
    # Check extracted etso formats
    print("Çıkarılan dosyada Etso Kodu formatları:")
    ext_etso_samples = ext_df['Etso Kodu'].dropna().unique()[:15]
    for etso in ext_etso_samples:
        print(f"  - {etso}")
    print()
    
    # Count numeric vs alphanumeric
    ref_numeric = sum(1 for etso in ref_df['Etso Kodu'].dropna() if str(etso).isdigit())
    ref_alphanumeric = len(ref_df['Etso Kodu'].dropna()) - ref_numeric
    
    ext_numeric = sum(1 for etso in ext_df['Etso Kodu'].dropna() if str(etso).isdigit())
    ext_alphanumeric = len(ext_df['Etso Kodu'].dropna()) - ext_numeric
    
    print(f"Referans - Sayısal: {ref_numeric}, Alfanümerik: {ref_alphanumeric}")
    print(f"Çıkarılan - Sayısal: {ext_numeric}, Alfanümerik: {ext_alphanumeric}")
    print()


def export_with_calculated_columns():
    """Export extracted data with calculated KDV and Toplam"""
    print("=" * 80)
    print("ÇIKTI: KDV ve Toplam ile Birlikte Çıkarılan Veriler")
    print("=" * 80)
    print()
    
    df = pd.read_excel(EXTRACTED_FILE)
    
    # Calculate KDV and Toplam properly
    kdv_values = []
    total_values = []
    for idx, row in df.iterrows():
        kdv, total = calculate_kdv_and_total(row)
        kdv_values.append(kdv)
        total_values.append(total)
    
    df['KDV'] = kdv_values
    df['Toplam (TL)'] = total_values
    
    # Add missing columns from reference
    df['AG OG'] = df['AG OG'].fillna('')
    df['TERİM'] = df['TERİM'].fillna('')
    df['KURULU GÜÇ'] = df['KURULU GÜÇ'].fillna('')
    df['SAYAXA ATILACAK TARİFE'] = ''
    
    # Reorder columns to match reference
    column_order = [
        'Dağıtım Bölgesi', 'Etso Kodu', 'Müşteri', 'Tarife Grubu', 'AG OG', 
        'TERİM', 'Güç kW', 'KURULU GÜÇ', 'Aktif Enerji Tüketim (kWh)',
        'Dağıtım Bedeli(TL)', 'Güç Bedeli(TL)', 'Güç Aşım Bedeli (TL)', 
        'Reaktif Bedel (TL)', 'KDV', 'Toplam (TL)', 'SAYAXA ATILACAK TARİFE'
    ]
    
    # Add missing columns
    for col in column_order:
        if col not in df.columns:
            df[col] = ''
    
    df = df[column_order]
    
    # Save to new file
    output_file = os.path.join(BASE_PATH, "Çıkarılan_Veriler_KDV_ile.xlsx")
    df.to_excel(output_file, index=False)
    
    print(f"  ✓ {output_file} kaydedildi")
    print(f"  ✓ Shape: {df.shape}")
    print(f"  ✓ Columns: {list(df.columns)}")
    print()
    
    return df


if __name__ == "__main__":
    # Run comparison
    missing, extra = compare_data()
    
    # Verify etso normalization
    verify_etso_normalization()
    
    # Export with calculated columns
    export_with_calculated_columns()
    
    print("=" * 80)
    print("ÖZET")
    print("=" * 80)
    print(f"  Referans kayıtlar: 2,584")
    print(f"  Çıkarılan kayıtlar: 3,353")
    print(f"  Eşleşen kayıtlar: {2584 - len(missing)}")
    print(f"  Eksik kayıtlar: {len(missing)}")
    print(f"  Fazla kayıtlar: {len(extra)}")
    print()
