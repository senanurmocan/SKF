#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SKF Başlıkları Güncel mapping dosyalarını (Mapping 6/7 uyumlu) parse eder
- Excel column letter to index converter
- Region-specific mapping creation
- Notlar sheet rules implementation
"""

import pandas as pd
from collections import defaultdict
from pathlib import Path
import re


def _mapping_text(value):
    """Mapping hücresini boş/`-` değerlerinden arındırarak metne çevir."""
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    return None if not text or text == '-' else text

def excel_column_letter_to_index(letter):
    """
    Excel column letter to 0-based index
    A=0, B=1, ..., Z=25, AA=26, AB=27, ...
    """
    if not letter:
        return -1
    
    letter = str(letter).strip().upper()
    result = 0
    
    for char in letter:
        if 'A' <= char <= 'Z':
            result = result * 26 + (ord(char) - ord('A') + 1)
        else:
            return -1
    
    return result - 1  # Convert to 0-based index


def index_to_excel_column(index):
    """
    0-based index to Excel column letter
    0=A, 1=B, ..., 25=Z, 26=AA, 27=AB, ...
    """
    if index < 0:
        return ""
    
    result = ""
    while index >= 0:
        result = chr(index % 26 + ord('A')) + result
        index = index // 26 - 1
    
    return result


def _parse_distribution_name_mapping(df_corrections):
    """
    Düzeltme sayfasındaki Dağıtım Adı tablosunu (G/H sütunları) parse eder.

    G sütunu: 'Olabilecek Versiyonlar' (klasör veya veri kaynağındaki farklı yazımlar)
    H sütunu: 'Algılaması Gereken' (kanonik/standart bölge ismi)

    Döndürülen dict casefold() anahtarlıdır; böylece büyük/küçük harf
    duyarsız eşleştirme yapılır.  Örnek::

        {"bedaş": "Boğaziçi EDAŞ", "boğaziçi": "Boğaziçi EDAŞ", ...}
    """
    dist_names = {}  # casefold key -> canonical name
    if df_corrections is None or df_corrections.empty:
        return dist_names

    g_col = 6  # G sütunu (0-indexed)
    h_col = 7  # H sütunu (0-indexed)

    if h_col >= len(df_corrections.columns):
        return dist_names

    for row_idx in range(2, len(df_corrections)):
        variant = _mapping_text(df_corrections.iloc[row_idx, g_col]) if g_col < len(df_corrections.columns) else None
        canonical = _mapping_text(df_corrections.iloc[row_idx, h_col])

        if variant is None or canonical is None:
            continue

        key = variant.casefold()
        dist_names[key] = canonical

    if dist_names:
        print(f"\n  Dağıtım Adı Normalizasyon Tablosu: {len(dist_names)} kayıt okundu")

    return dist_names


def parse_mapping_file(file_path):
    """
    SKF Başlıkları Güncel mapping dosyasını parse eder.
    
    Return: {
        'regions': {
            'ADM EDAŞ': {
                'etso': 'Sayac ID',
                'musteri': 'B',  # Column letter
                'tarife': 'Trf.nominal dğr.gr.',
                'ag_og': 'Tarife tipi',
                'terim': None,
                'güç_kw': 'Sözleşme gücü',
                'kurulu_güç': 'Kurulu gücü',
                'aktif_enerji': 'Dağıtım Miktarı',
                'trafo_kaybı': None,
                'dagitim_bedeli': 'Dağıtım Bedeli',
                'güç_bedeli': 'Güç Bedeli',
                'güç_aşım': 'Güç Aşım Bedeli',
                'reaktif': 'RI Bedeli',
                'reaktif2': 'RC Bedeli',
                'reaktif_tenzil': 'Tenzil RC Bedeli',
                'reaktif_tenzil2': 'Tenzil RI Bedeli',
                'column_map': {
                    'Sayac ID': 'B',
                    'Trf.nominal dğr.gr.': 'BK',
                    ...
                }
            },
            ...
        },
        'notes': {
            'reaktif_toplam': 'Z sütunundan itibaren dolu sütunların toplamı',
            'eksi_veriler': 'Getirilmeli',
            'aktif_enerji_toplam': 'P + R (Boğaziçi, Sakarya, Uludağ, Yeşilırmak)',
            'duplicate_handling': 'Max for sözleşme/güç, sum for others',
            'akdeniz_parentheses': 'Eksi değerler parantez içi',
            'akedas_tl_raporu': 'TL_Raporu dosyaları alınmalı',
            'uludağ_5_dosya': 'Sadece 4008-TL dosyası',
            'reaktif_tenzil': 'AD + AF toplamı (Trakya için sadece AF)',
        }
    }
    """
    print("=" * 80)
    print("YENİ MAPPING DOSYASI PARSE İŞLEMİ")
    print("=" * 80)
    print()
    
    # Excel dosyasını oku
    xls = pd.ExcelFile(file_path)
    
    # Başlıklar sheetini oku
    df_headers = pd.read_excel(file_path, sheet_name="Başlıklar", header=None)
    
    # Mevcut en güncel Notlar sayfasını tercih et; eski mapping dosyalarıyla uyumluluğu koru.
    notlar_sheets = ["Notlar 9", "Notlar 8", "Notlar 7", "Notlar 6", "Notlar 5", "Notlar 4", "Notlar 3", "Notlar 2", "Notlar 1", "Notlar"]
    notes_sheet = next((sheet for sheet in notlar_sheets if sheet in xls.sheet_names), None)
    if notes_sheet is None:
        df_notes = pd.DataFrame()
    else:
        df_notes = pd.read_excel(file_path, sheet_name=notes_sheet, header=None)

    if "Düzeltme" in xls.sheet_names:
        df_corrections = pd.read_excel(file_path, sheet_name="Düzeltme", header=None)
    else:
        df_corrections = pd.DataFrame()
    
    # Mapping structure
    mapping = {
        'regions': {},
        'notes': {},
        'column_letters': {},  # Column letter to index mapping
        'notes_sheet': notes_sheet,
        'corrections': {
            'Tarife Grubu': {},
            'AG OG': {},
            'TERİM': {},
        },
        'correction_duplicates': [],
        'mapping_warnings': [],
        'features': {},
    }
    
    # Header mapping (Satır 0)
    headers = df_headers.iloc[0].tolist() if len(df_headers) > 0 else []

    header_texts = [str(value).strip() for value in headers if pd.notna(value)]
    has_tazminat_fields = any(value.startswith('Tazminat Bedeli') for value in header_texts)
    standart_disi_headers = [value for value in header_texts if value.startswith('Standart Dışı-')]
    has_standart_disi_fields = bool(standart_disi_headers)
    mapping['features'] = {
        'notlar_7': notes_sheet is not None,
        'corrections': not df_corrections.empty,
        'extended_financial_fields': has_tazminat_fields or has_standart_disi_fields,
    }

    duplicate_standard_headers = sorted({
        value for value in standart_disi_headers if standart_disi_headers.count(value) > 1
    })
    for duplicate_header in duplicate_standard_headers:
        mapping['mapping_warnings'].append(
            f"Başlıklar sayfasında yinelenen fiziksel alan: {duplicate_header}"
        )
    
    print("HEADER MAPPING YAPISI (Satır 0):")
    print("-" * 80)
    for idx, header in enumerate(headers):
        if pd.notna(header):
            print(f"  [{idx}] {header}")
    print()
    
    # Her bölge için mapping oluştur (Satır 1-21)
    print("BÖLGE MAPPING'LARI:")
    print("-" * 80)
    
    for row_idx in range(1, len(df_headers)):
        row = df_headers.iloc[row_idx].tolist()
        
        if len(row) < 2:
            continue
            
        region_name = row[0] if pd.notna(row[0]) else None
        etso_header = row[1] if pd.notna(row[1]) else None
        
        if not region_name:
            continue
        
        print(f"\n{region_name}:")

        region_mapping = {
            'etso': str(etso_header) if etso_header else None,
            'musteri': None,
            'tarife': None,
            'ag_og': None,
            'terim': None,
            'güç_kw': None,
            'kurulu_güç': None,
            'aktif_enerji': None,
            'trafo_kaybı': None,
            'dagitim_bedeli': None,
            'güç_bedeli': None,
            'güç_aşım': None,
            'reaktif': None,
            'reaktif2': None,
            'reaktif_tenzil': None,
            'reaktif_tenzil2': None,
            'tazminat_fields': [],
            'standart_disi_fields': [],
            'extended_field_metadata': [],
            'column_map': {}
        }

        # Her bölgenin kendi satırındaki gerçek kaynak başlıklarını kullan.
        # Böylece örneğin Osmangazi için jenerik "Güç kW" yerine "Baglanti Gucu" eşleşir.
        target_headers = []
        for col_idx in [1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31]:
            if col_idx < len(row):
                target_headers.append(row[col_idx])
            else:
                target_headers.append(None)

        # Satır 1-21'deki region mapping satırındaki tek sütunlar (column letters)
        # Column letters: Satır 1-21'deki 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32
        column_letters = []
        for col_idx in [2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32]:
            if col_idx < len(row):
                column_letters.append(row[col_idx])
            else:
                column_letters.append(None)

        # Map fields - header names ile column letters'i eşleştir
        # Musteri header'ı ile musteri column letter'ı eşleşmeli, tarife ile tarife, vb.
        # NOT: etso header'ı row[1]'de, column letter'ı row[2]'de
        field_names = ['etso', 'musteri', 'tarife', 'ag_og', 'terim', 'güç_kw', 'kurulu_güç',
                      'aktif_enerji', 'trafo_kaybı', 'dagitim_bedeli', 'güç_bedeli',
                      'güç_aşım', 'reaktif', 'reaktif2', 'reaktif_tenzil', 'reaktif_tenzil2']

        for i, (header, col_letter) in enumerate(zip(target_headers, column_letters)):
            if pd.notna(header) and pd.notna(col_letter):
                field = field_names[i]
                region_mapping[field] = str(header)
                region_mapping['column_map'][str(header)] = str(col_letter)

                # Convert column letter to index for later use
                col_index = excel_column_letter_to_index(str(col_letter))
                region_mapping[f'{field}_index'] = col_index
                print(f"  {field}: {header} (Column: {col_letter}, Index: {col_index})")

        # Mapping 7: üç tazminat ve fiziksel olarak dokuz standart-dışı çifti.
        # 7a/7b/8 tarihsel iç anahtarları fiziksel 7/8/9 pozisyonlarını temsil
        # eder; çıktı hesabı dokuz çiftin tamamını kayıpsız toplar.
        if mapping['features']['extended_financial_fields']:
            extended_pairs = [
                ('tazminat_1', 33, 34, 'tazminat'),
                ('tazminat_2', 35, 36, 'tazminat'),
                ('tazminat_3', 37, 38, 'tazminat'),
                ('standart_disi_1', 39, 40, 'standart_disi'),
                ('standart_disi_2', 41, 42, 'standart_disi'),
                ('standart_disi_3', 43, 44, 'standart_disi'),
                ('standart_disi_4', 45, 46, 'standart_disi'),
                ('standart_disi_5', 47, 48, 'standart_disi'),
                ('standart_disi_6', 49, 50, 'standart_disi'),
                ('standart_disi_7a', 51, 52, 'standart_disi'),
                ('standart_disi_7b', 53, 54, 'standart_disi'),
                ('standart_disi_8', 55, 56, 'standart_disi'),
            ]

            for field, header_idx, letter_idx, category in extended_pairs:
                logical_header = _mapping_text(headers[header_idx]) if header_idx < len(headers) else None
                source_header = _mapping_text(row[header_idx]) if header_idx < len(row) else None
                column_letter = _mapping_text(row[letter_idx]) if letter_idx < len(row) else None
                column_index = excel_column_letter_to_index(column_letter) if column_letter else -1

                region_mapping[field] = source_header
                region_mapping[f'{field}_index'] = column_index

                if source_header or column_letter:
                    field_list = (
                        region_mapping['tazminat_fields']
                        if category == 'tazminat'
                        else region_mapping['standart_disi_fields']
                    )
                    field_list.append(field)
                    region_mapping['extended_field_metadata'].append({
                        'field': field,
                        'category': category,
                        'logical_header': logical_header,
                        'source_header': source_header,
                        'column_letter': column_letter,
                        'column_index': column_index,
                    })
                    print(
                        f"  {field}: {source_header or '-'} "
                        f"(Column: {column_letter or '-'}, Index: {column_index})"
                    )
        
        # Store region mapping – aynı isimli birden fazla satır varsa (ör: Dicle EDAŞ
        # çift format), ikincisini alt_regions'a alternatif format olarak sakla.
        if region_name in mapping['regions']:
            if 'alt_regions' not in mapping:
                mapping['alt_regions'] = {}
            mapping['alt_regions'][region_name] = region_mapping
            print(f"  [!] '{region_name}' alternatif format olarak kaydedildi (çift format algılama)")
        else:
            mapping['regions'][region_name] = region_mapping

    # Düzeltme sayfasındaki üç bağımsız kaynak->hedef tablosunu oku.
    if not df_corrections.empty:
        correction_blocks = [
            ('Tarife Grubu', 0, 1),
            ('AG OG', 2, 3),
            ('TERİM', 4, 5),
        ]
        for output_field, source_col, target_col in correction_blocks:
            if source_col >= len(df_corrections.columns) or target_col >= len(df_corrections.columns):
                mapping['mapping_warnings'].append(
                    f"Düzeltme bloğu eksik: {output_field}"
                )
                continue

            lookup = mapping['corrections'][output_field]
            for row_idx in range(2, len(df_corrections)):
                source_value = _mapping_text(df_corrections.iloc[row_idx, source_col])
                target_value = _mapping_text(df_corrections.iloc[row_idx, target_col])
                if source_value is None and target_value is None:
                    continue
                if source_value is None or target_value is None:
                    mapping['mapping_warnings'].append(
                        f"Düzeltme satırı eksik ({output_field}, Excel satırı {row_idx + 1})"
                    )
                    continue

                if source_value in lookup:
                    duplicate = {
                        'field': output_field,
                        'source': source_value,
                        'existing_target': lookup[source_value],
                        'duplicate_target': target_value,
                        'excel_row': row_idx + 1,
                    }
                    mapping['correction_duplicates'].append(duplicate)
                    if lookup[source_value] != target_value:
                        mapping['mapping_warnings'].append(
                            f"Çelişkili Düzeltme eşlemesi: {output_field} / {source_value}"
                        )
                    continue

                lookup[source_value] = target_value

    # Düzeltme sayfasından Dağıtım Adı normalizasyon tablosunu oku (G/H sütunları).
    mapping['distribution_names'] = _parse_distribution_name_mapping(df_corrections)
    
    # Notlar sheetini parse et
    print("\n\nNOTLAR:")
    print("-" * 80)
    
    for row_idx in range(len(df_notes)):
        row = df_notes.iloc[row_idx].tolist()
        if len(row) >= 2:
            note_number = row[0]
            note_text = row[1]
            
            if pd.notna(note_number) and pd.notna(note_text):
                note_key = f"note_{int(note_number)}"
                mapping['notes'][note_key] = str(note_text)
                print(f"  {note_key}: {note_text}")
    
    return mapping


def apply_note_rules(mapping, region_name):
    """
    Notlar sayfasındaki kuralları uygula
    
    Kurallar:
    1. Reaktif bedel: Z sütunundan itibaren dolu sütunların toplamı
    2. Eksi veriler: Getirilmeli (negative values allowed)
    3. Aktif enerji toplam: P + R (Boğaziçi, Sakarya, Uludağ, Yeşilırmak)
    4. Duplicate handling: Max for sözleşme/güç, sum for others
    5. Akdeniz: Eksi değerler parantez içi
    6. AKEDAŞ: TL_Raporu dosyaları
    7. Uludağ: 4008-TL dosyası
    8. Reaktif Tenzil: AD + AF (Trakya için sadece AF)
    """
    rules = {
        'apply_negative_filter': True,  # Default: filter negative (rule #2 says: NO - include negatives)
        'reactive_total_rule': None,
        'active_energy_rule': None,
        'duplicate_handling': 'max_for_contract_power',
        'akdeniz_parentheses': False,
        'akedas_tl_raporu': False,
        'uludağ_specific': False,
        'reaktif_tenzil_rule': None,
        'force_fixed_fields': [],
    }
    
    notes = mapping['notes']
    
    # Rule #2: Eksi veriler getirilmeli
    rules['apply_negative_filter'] = False
    
    # Rule #3: Aktif enerji toplam (P + R)
    active_energy_regions = ['Boğaziçi EDAŞ', 'Sakarya EDAŞ', 'Uludağ EDAŞ', 'Yeşilırmak EDAŞ']
    if region_name in active_energy_regions:
        rules['active_energy_rule'] = 'P + R sum'
    
    # Rule #5: Akdeniz eksi değerler parantez içi
    if region_name == 'Akdeniz EDAŞ':
        rules['akdeniz_parentheses'] = True
    
    # Rule #6: AKEDAŞ TL_Raporu
    if region_name == 'AKEDAŞ':
        rules['akedas_tl_raporu'] = True
    
    # Rule #7: Uludağ 4008-TL
    if region_name == 'Uludağ EDAŞ':
        rules['uludağ_specific'] = True
    
    # Rule #8: Reaktif Tenzil
    if region_name == 'Trakya EDAŞ':
        rules['reaktif_tenzil_rule'] = 'AF only'
    else:
        rules['reaktif_tenzil_rule'] = 'AD + AF sum'

    # Notlar 7 #1 & Notlar 9 #1: Çamlıbel'de tekrarlanan Tarife başlıkları arasından AV (index 47) zorunlu.
    if region_name == 'Çamlıbel EDAŞ':
        rules['force_fixed_fields'] = ['tarife', 'ag_og', 'terim']

    # Notlar 7 #2: Bu bölgelerde Reaktif Bedel yalnız base iki alandan oluşur.
    base_reactive_regions = ['ADM EDAŞ', 'Gediz EDAŞ', 'Trakya EDAŞ', 'Uludağ EDAŞ']
    if region_name in base_reactive_regions:
        rules['reactive_total_rule'] = 'base_columns_only'
    
    return rules


def get_column_index_from_letter(column_letter, region_mapping):
    """
    Column letter (B, F, J...) → 0-based index
    """
    return excel_column_letter_to_index(column_letter)


def extract_file_filter(region_name, rules, filename):
    """
    Region-specific file filtering
    
    AKEDAŞ: TL_Raporu içeren dosyalar
    Uludağ: Sadece 4008-TL dosyası
    """
    if rules['akedas_tl_raporu']:
        return 'TL_Raporu' in filename
    
    if rules['uludağ_specific']:
        return '4008' in filename and 'TL' in filename and 'KWH' not in filename.upper()
    
    return True  # Default: include all files


if __name__ == "__main__":
    file_path = Path(__file__).resolve().parent / "SKF Başlıkları Güncel 7.xlsx"
    
    mapping = parse_mapping_file(file_path)
    
    # Test rules for specific regions
    print("\n\n" + "=" * 80)
    print("TEST: BÖLGE SPESİFİK KURALLAR")
    print("=" * 80)
    
    test_regions = ['AKEDAŞ', 'Uludağ EDAŞ', 'Trakya EDAŞ', 'Boğaziçi EDAŞ', 'Akdeniz EDAŞ']
    
    for region in test_regions:
        if region in mapping['regions']:
            rules = apply_note_rules(mapping, region)
            print(f"\n{region}:")
            for key, value in rules.items():
                print(f"  {key}: {value}")
