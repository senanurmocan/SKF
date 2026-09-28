#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Tanımlayıcı Karşılaştırma Scripti
- Belirli bir bölge (Boğaziçi EDAŞ) için field-by-field karşılaştırma yapar
- Eksik/fazla kayıtları tanımlar ve sebeplerini gösterir
"""
from pathlib import Path

import os
import sys
import re
import openpyxl
import pandas as pd
from collections import defaultdict

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from new_mapping_parser import parse_mapping_file, apply_note_rules, excel_column_letter_to_index

BASE_PATH = Path(__file__).resolve().parents[1]
MAPPING_FILE = os.path.join(BASE_PATH, 'SKF Başlıkları Güncel.xlsx')
REFERENCE_FILE = os.path.join(BASE_PATH, 'Dağıtımın Kestiği Faturalar Özet.xlsx')

# Region name mapping (extraction -> reference)
REGION_NORMALIZATION_MAPPING = {
    "AKEDAŞ": "AKEDAŞ ( Göksu EDAŞ )",
    "ADM EDAŞ": "Aydem EDAŞ",
    "AYEDAŞ": "Ayedaş",
    "DİCLE EDAŞ": "Dicle Edaş",
    "FIRAT EDAŞ": "Fırat Edaş",
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
}


def normalize_region_name(region_name):
    """Extraction region name -> Reference region name"""
    if region_name is None:
        return None
    region_name = str(region_name).strip().upper()
    return REGION_NORMALIZATION_MAPPING.get(region_name, region_name)


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

    # Format "40Z0000006543210G" gibi başlıyorsa, sadece 7 haneli numarayı çıkar
    match = re.match(r'^\d{2}Z0*(\d{7})', etso_str, re.IGNORECASE)
    if match:
        result = match.group(1).lstrip('0')
        return result if result else '0'

    # Alternative: find any 7 consecutive digits
    match = re.search(r'\d{7}', etso_str)
    if match:
        return match.group(0)

    # Fallback: extract all digits and take last 7
    digits = re.findall(r'\d+', etso_str)
    if digits:
        all_digits = ''.join(digits)
        if len(all_digits) >= 7:
            return all_digits[-7:].lstrip('0') or '0'
        else:
            return all_digits.lstrip('0') or '0'

    return etso_str


def normalize_header(header):
    """Header'ı temizle ve normalleştir"""
    if header is None:
        return ''
    header = str(header).strip().lower()
    replacements = {
        'ı': 'i', 'İ': 'i', 'ğ': 'g', 'Ğ': 'g', 'ü': 'u', 'Ü': 'u',
        'ö': 'o', 'Ö': 'o', 'ç': 'c', 'Ç': 'c', 'ş': 's', 'Ş': 's'
    }
    for tr, en in replacements.items():
        header = header.replace(tr, en)
    header = re.sub(r'[\s_]+', '', header)
    header = re.sub(r'[^\w]', '', header)
    return header


def find_column_index(headers, target_headers):
    """Header listesinden hedef header'ı bul"""
    normalized_headers = [normalize_header(h) for h in headers]
    for target in target_headers:
        normalized_target = normalize_header(target)
        if normalized_target in normalized_headers:
            return normalized_headers.index(normalized_target)
    return -1


def load_mapping_new():
    """SKF Başlıkları Güncel.xlsx'den mapping'i yükle - 16-field mapping"""
    wb = openpyxl.load_workbook(MAPPING_FILE)
    ws = wb['Başlıklar']
    rows = list(ws.iter_rows(values_only=True))

    mapping = {}
    for row in rows[1:]:  # İlk satır header
        if row[0]:  # Dağıtım Bölgesi doluysa
            region = str(row[0]).strip()
            mapping[region] = {
                'region_name': region,
                'etso': row[1], 'musteri': row[2], 'tarife': row[3],
                'ag_og': row[4], 'term': row[5], 'guc_kw': row[6],
                'kurulu_guc': row[7], 'aktif_enerji': row[8],
                'dagitim_bedeli': row[9], 'guc_bedeli': row[10],
                'guc_asim': row[11], 'reaktif': row[12],
                'reaktif2': row[13], 'reaktif_tenzil': row[14],
                'reaktif_tenzil2': row[15],
                'etso_index': -1, 'musteri_index': -1, 'tarife_index': -1,
                'ag_og_index': -1, 'term_index': -1, 'guc_kw_index': -1,
                'kurulu_guc_index': -1, 'aktif_enerji_index': -1,
                'dagitim_bedeli_index': -1, 'guc_bedeli_index': -1,
                'guc_asim_index': -1, 'reaktif_index': -1,
                'reaktif2_index': -1, 'reaktif_tenzil_index': -1,
                'reaktif_tenzil2_index': -1
            }

    # Pre-compute column indices for each region
    for region_name, region_data in mapping.items():
        excel_file = region_data.get('excel_file', '')
        if excel_file:
            excel_path = os.path.join(BASE_PATH, 'Dağıtım', excel_file)
            if os.path.exists(excel_path):
                try:
                    wb_region = openpyxl.load_workbook(excel_path)
                    sheet = wb_region.active
                    headers = list(sheet.iter_rows(values_only=True))[0]

                    field_mapping = {
                        'etso': region_data['etso'],
                        'musteri': region_data['musteri'],
                        'tarife': region_data['tarife'],
                        'ag_og': region_data['ag_og'],
                        'term': region_data['term'],
                        'guc_kw': region_data['guc_kw'],
                        'kurulu_guc': region_data['kurulu_guc'],
                        'aktif_enerji': region_data['aktif_enerji'],
                        'dagitim_bedeli': region_data['dagitim_bedeli'],
                        'guc_bedeli': region_data['guc_bedeli'],
                        'guc_asim': region_data['guc_asim'],
                        'reaktif': region_data['reaktif'],
                        'reaktif2': region_data['reaktif2'],
                        'reaktif_tenzil': region_data['reaktif_tenzil'],
                        'reaktif_tenzil2': region_data['reaktif_tenzil2'],
                    }

                    for field_name, header_name in field_mapping.items():
                        if header_name:
                            idx = find_column_index(headers, [header_name])
                            key = f'{field_name}_index'
                            region_data[key] = idx

                except Exception as e:
                    print(f"  Hata ({region_name}): {e}")

    return mapping


def extract_region_data(region_name, region_mapping):
    """Belirli bir bölgenin Excel dosyasından veri çıkar"""
    region_data = region_mapping.get(region_name, {})
    excel_file = region_data.get('excel_file', '')
    
    if not excel_file:
        print(f"  {region_name}: Excel dosyası bulunamadı")
        return []

    excel_path = os.path.join(BASE_PATH, 'Dağıtım', excel_file)
    if not os.path.exists(excel_path):
        print(f"  {region_name}: Dosya yok - {excel_path}")
        return []

    try:
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active
        rows = list(sheet.iter_rows(values_only=True))
        
        if not rows:
            return []
        
        headers = rows[0]
        
        # Pre-computed index'leri kullan
        etso_idx = region_data.get('etso_index', -1)
        musteri_idx = region_data.get('musteri_index', -1)
        tarife_idx = region_data.get('tarife_index', -1)
        ag_og_idx = region_data.get('ag_og_index', -1)
        term_idx = region_data.get('term_index', -1)
        guc_kw_idx = region_data.get('guc_kw_index', -1)
        kurulu_guc_idx = region_data.get('kurulu_guc_index', -1)
        aktif_idx = region_data.get('aktif_enerji_index', -1)
        dagitim_idx = region_data.get('dagitim_bedeli_index', -1)
        guc_bedeli_idx = region_data.get('guc_bedeli_index', -1)
        guc_asim_idx = region_data.get('guc_asim_index', -1)
        
        reaktif_idx = region_data.get('reaktif_index', -1)
        reaktif2_idx = region_data.get('reaktif2_index', -1)
        reaktif_tenzil_idx = region_data.get('reaktif_tenzil_index', -1)
        reaktif_tenzil2_idx = region_data.get('reaktif_tenzil2_index', -1)

        data_rows = []
        for row in rows[1:]:
            if not row or all(v is None for v in row):
                continue
            
            # Reaktif summing
            reaktif_toplam = 0
            for idx in [reaktif_idx, reaktif2_idx, reaktif_tenzil_idx, reaktif_tenzil2_idx]:
                if idx >= 0 and idx < len(row):
                    val = row[idx]
                    if val is not None:
                        try:
                            if isinstance(val, str):
                                val = val.strip().replace(',', '').replace('(', '-').replace(')', '')
                            val = float(val)
                            reaktif_toplam += val
                        except (ValueError, AttributeError):
                            pass
            
            # Etso extraction
            if etso_idx >= 0 and etso_idx < len(row):
                etso_val = row[etso_idx]
                if isinstance(etso_val, str):
                    etso_normalized = normalize_etso_kodu(etso_val)
                else:
                    etso_normalized = normalize_etso_kodu(etso_val)
            else:
                etso_normalized = None
            
            # Customer name extraction
            customer_name = None
            if musteri_idx >= 0 and musteri_idx < len(row):
                customer_name = row[musteri_idx]
            
            data_row = {
                'region': region_name,
                'etso_raw': row[etso_idx] if etso_idx >= 0 else None,
                'etso_normalized': etso_normalized,
                'customer_raw': customer_name,
                'customer_normalized': customer_name,
                'tarife': row[tarife_idx] if tarife_idx >= 0 else None,
                'dagitim_bedeli': row[dagitim_idx] if dagitim_idx >= 0 else 0,
                'guc_bedeli': row[guc_bedeli_idx] if guc_bedeli_idx >= 0 else 0,
                'guc_asim': row[guc_asim_idx] if guc_asim_idx >= 0 else 0,
                'reaktif': reaktif_toplam,
                'bedel_toplam': (float(row[dagitim_idx]) if dagitim_idx >= 0 and row[dagitim_idx] else 0) +
                               (float(row[guc_bedeli_idx]) if guc_bedeli_idx >= 0 and row[guc_bedeli_idx] else 0) +
                               (float(row[guc_asim_idx]) if guc_asim_idx >= 0 and row[guc_asim_idx] else 0) +
                               reaktif_toplam
            }
            data_rows.append(data_row)

        print(f"  {region_name}: {len(data_rows)} kayıt")
        return data_rows

    except Exception as e:
        print(f"  {region_name}: Hata - {e}")
        import traceback
        traceback.print_exc()
        return []


def load_reference_data():
    """Referans dosyasını yükle"""
    try:
        df = pd.read_excel(REFERENCE_FILE)
        print(f"Referans dosyası: {len(df)} kayıt")
        
        # Customer name column'ları kontrol et
        print(f"\nReferans dosyası kolonları:")
        for col in df.columns:
            print(f"  - {col}")
        
        # Etso Kodu column'ını bul
        etso_col = None
        for col in df.columns:
            if 'etso' in col.lower() or 'kodu' in col.lower():
                etso_col = col
                break
        
        if etso_col:
            print(f"\nEtso Kodu column: {etso_col}")
            sample_etso = df[etso_col].dropna().head(5)
            print(f"Örnek Etso Kodu değerleri:")
            for val in sample_etso:
                print(f"  - {val} (type: {type(val).__name__})")
        
        # Customer name column'ını bul
        customer_cols = [col for col in df.columns if 'müşteri' in col.lower() or 'adi' in col.lower() or 'name' in col.lower()]
        if customer_cols:
            print(f"\nCustomer name column'ları:")
            for col in customer_cols:
                sample = df[col].dropna().head(3)
                print(f"  {col}: {list(sample)}")
        
        # Bedel column'larını bul
        bedel_cols = [col for col in df.columns if 'bedel' in col.lower() or 'toplam' in col.lower()]
        if bedel_cols:
            print(f"\nBedel column'ları:")
            for col in bedel_cols:
                print(f"  - {col}")
        
        return df
    
    except Exception as e:
        print(f"Referans dosyası yüklenemedi: {e}")
        import traceback
        traceback.print_exc()
        return None


def analyze_discrepancies(extracted_df, reference_df):
    """Eksik ve fazla kayıtları analiz et"""
    # Key creation
    extracted_df['key'] = extracted_df['etso_normalized'].astype(str) + '_' + \
                          extracted_df['bedel_toplam'].astype(str) + '_' + \
                          extracted_df['region'].str.upper()
    
    reference_df['key_ref'] = reference_df['Etso Kodu'].astype(str) + '_' + \
                              reference_df['Dağıtım Bedeli(TL)'].astype(str) + '_' + \
                              reference_df['Dağıtım Bölgesi'].str.upper()
    
    # Keys to sets
    extracted_keys = set(extracted_df['key'].unique())
    reference_keys = set(reference_df['key_ref'].unique())
    
    # Find missing and extra
    missing_keys = reference_keys - extracted_keys
    extra_keys = extracted_keys - reference_keys
    
    print(f"\n=== ANALİZ ===")
    print(f"Referans keys: {len(reference_keys)}")
    print(f"Extraction keys: {len(extracted_keys)}")
    print(f"Eksik kayıtlar: {len(missing_keys)}")
    print(f"Fazla kayıtlar: {len(extra_keys)}")
    
    # Sample missing records
    if len(missing_keys) > 0:
        print(f"\nEksik kayıtlar (ilk 5):")
        missing_df = reference_df[reference_df['key_ref'].isin(list(missing_keys)[:5])]
        for idx, row in missing_df.iterrows():
            print(f"  Etso: {row['Etso Kodu']}, Bedel: {row['Dağıtım Bedeli(TL)']}, Region: {row['Dağıtım Bölgesi']}")
    
    # Sample extra records
    if len(extra_keys) > 0:
        print(f"\nFazla kayıtlar (ilk 5):")
        extra_df = extracted_df[extracted_df['key'].isin(list(extra_keys)[:5])]
        for idx, row in extra_df.iterrows():
            print(f"  Etso: {row['etso_normalized']}, Bedel: {row['bedel_toplam']}, Region: {row['region']}")


def main():
    print("=== TANIMLAYICI KARŞILAŞTIRMA ===\n")
    
    # Referans dosyasını yükle
    print("Referans dosyası yükleniyor...")
    reference_df = load_reference_data()
    if reference_df is None:
        return
    
    # Mapping'i yükle
    print("\nMapping yükleniyor...")
    mapping = load_mapping_new()
    
    # Boğaziçi EDAŞ için extraction yap
    print("\nBoğaziçi EDAŞ verisi çıkarılıyor...")
    bozaiçi_data = extract_region_data('BOĞAZİÇİ EDAŞ', mapping)
    
    # Diğer tüm bölgeleri extraction yap
    print("\nTüm bölgeler extraction ediliyor...")
    all_data = []
    for region_name in ['BOĞAZİÇİ EDAŞ', 'AKDENİZ EDAŞ', 'BAŞKENT EDAŞ']:
        data = extract_region_data(region_name, mapping)
        all_data.extend(data)
    
    # DataFrame'e çevir
    extracted_df = pd.DataFrame(all_data)
    print(f"\nToplam extraction: {len(extracted_df)} kayıt")
    
    # Analyze discrepancies
    print("\n=== KARŞILAŞTIRMA ===")
    analyze_discrepancies(extracted_df, reference_df)
    
    # Debug output
    print("\n=== DEBUG OUTPUT ===")
    
    # Boğaziçi EDAŞ için header'ları göster
    print("\nBoğaziçi EDAŞ Excel dosyası header'ları:")
    excel_path = os.path.join(BASE_PATH, 'Dağıtım', 'Boğaziçi EDAŞ - Dağıtım.xlsx')
    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active
        headers = list(sheet.iter_rows(values_only=True))[0]
        for i, h in enumerate(headers):
            print(f"  [{i}] {h}")
    
    # Mapping'deki header mappings
    print("\nMapping'deki Boğaziçi EDAŞ header mappings:")
    bozaiçi_mapping = mapping.get('BOĞAZİÇİ EDAŞ', {})
    for field, header in bozaiçi_mapping.items():
        if field.endswith('_index'):
            print(f"  {field}: {header}")
    
    # Sample extracted data
    if len(extracted_df) > 0:
        print("\nSample extracted data (Boğaziçi EDAŞ):")
        bozaiçi_extracted = extracted_df[extracted_df['region'] == 'BOĞAZİÇİ EDAŞ'].head(3)
        for idx, row in bozaiçi_extracted.iterrows():
            print(f"  Etso: {row['etso_normalized']}, Customer: {row['customer_normalized']}, Bedel: {row['bedel_toplam']}")


if __name__ == '__main__':
    main()
