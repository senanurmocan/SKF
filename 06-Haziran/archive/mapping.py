"""
Mapping file parser, Excel column letter converter, and note-based business rules.
"""

import os
import pandas as pd
from config.constants import FIELD_SYNONYMS, SPECIAL_HEADER_MAPPING


def excel_column_letter_to_index(letter):
    """
    Convert Excel column letter to 0-based index.
    A=0, B=1, ..., Z=25, AA=26, AB=27, ...
    """
    if not letter or pd.isna(letter):
        return -1
    
    letter_str = str(letter).strip().upper()
    result = 0
    for char in letter_str:
        if 'A' <= char <= 'Z':
            result = result * 26 + (ord(char) - ord('A') + 1)
        else:
            return -1
    return result - 1


def index_to_excel_column(index):
    """Convert 0-based index to Excel column letter."""
    if index < 0:
        return ""
    result = ""
    while index >= 0:
        result = chr(index % 26 + ord('A')) + result
        index = index // 26 - 1
    return result


def parse_mapping_file(file_path):
    """
    Parse SKF Başlıklar Excel mapping file.
    Reads sheet 'Başlıklar' and notes sheet ('Notlar 6', 'Notlar 5', etc.).
    Extracts dynamic region-specific header names and column letters.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Mapping file not found: {file_path}")

    # Read Başlıklar sheet
    df_headers = pd.read_excel(file_path, sheet_name="Başlıklar", header=None)

    # Read Notlar sheet in order of preference
    notlar_sheets = ['Notlar 6', 'Notlar 5', 'Notlar 4', 'Notlar 3', 'Notlar 2', 'Notlar 1', 'Notlar']
    df_notes = None
    for sheet_name in notlar_sheets:
        try:
            df_notes = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            break
        except ValueError:
            continue

    if df_notes is None:
        df_notes = pd.DataFrame()

    mapping = {
        'regions': {},
        'notes': {},
        'column_letters': {}
    }

    field_names = [
        'etso', 'musteri', 'tarife', 'ag_og', 'terim', 'güç_kw', 'kurulu_güç',
        'aktif_enerji', 'trafo_kaybı', 'dagitim_bedeli', 'güç_bedeli',
        'güç_aşım', 'reaktif', 'reaktif2', 'reaktif_tenzil', 'reaktif_tenzil2'
    ]

    for row_idx in range(1, len(df_headers)):
        row = df_headers.iloc[row_idx].tolist()
        if len(row) < 2:
            continue

        region_name = row[0] if pd.notna(row[0]) else None
        etso_header = row[1] if pd.notna(row[1]) else None

        if not region_name:
            continue

        region_mapping = {
            'etso': str(etso_header) if etso_header else None,
            'musteri': None, 'tarife': None, 'ag_og': None, 'terim': None,
            'güç_kw': None, 'kurulu_güç': None, 'aktif_enerji': None,
            'trafo_kaybı': None, 'dagitim_bedeli': None, 'güç_bedeli': None,
            'güç_aşım': None, 'reaktif': None, 'reaktif2': None,
            'reaktif_tenzil': None, 'reaktif_tenzil2': None,
            'column_map': {}
        }

        # Dynamic header names from region row
        target_headers = [
            row[c] if c < len(row) and pd.notna(row[c]) else None
            for c in [1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31]
        ]

        # Column letters from region row
        column_letters = [
            row[c] if c < len(row) and pd.notna(row[c]) else None
            for c in [2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32]
        ]

        for i, (header, col_letter) in enumerate(zip(target_headers, column_letters)):
            if pd.notna(header) and pd.notna(col_letter):
                field = field_names[i]
                region_mapping[field] = str(header)
                region_mapping['column_map'][str(header)] = str(col_letter)
                col_index = excel_column_letter_to_index(str(col_letter))
                region_mapping[f'{field}_index'] = col_index

        mapping['regions'][region_name] = region_mapping

    # Parse notes
    for row_idx in range(len(df_notes)):
        row = df_notes.iloc[row_idx].tolist()
        if len(row) >= 2 and pd.notna(row[0]) and pd.notna(row[1]):
            note_key = f"note_{int(row[0])}" if isinstance(row[0], (int, float)) else str(row[0])
            mapping['notes'][note_key] = str(row[1])

    return mapping


def apply_note_rules(mapping, region_name):
    """Extract business rules derived from Note sheets for a specific region."""
    rules = {
        'apply_negative_filter': False,
        'reactive_total_rule': None,
        'active_energy_rule': None,
        'duplicate_handling': 'max_for_contract_power',
        'akdeniz_parentheses': False,
        'akedas_tl_raporu': False,
        'uludağ_specific': False,
        'reaktif_tenzil_rule': None,
    }

    active_energy_regions = ['Boğaziçi EDAŞ', 'Sakarya EDAŞ', 'Uludağ EDAŞ', 'Yeşilırmak EDAŞ']
    if region_name in active_energy_regions:
        rules['active_energy_rule'] = 'P + R sum'

    if region_name == 'Akdeniz EDAŞ':
        rules['akdeniz_parentheses'] = True

    if region_name == 'AKEDAŞ':
        rules['akedas_tl_raporu'] = True

    if region_name == 'Uludağ EDAŞ':
        rules['uludağ_specific'] = True

    if region_name == 'Trakya EDAŞ':
        rules['reaktif_tenzil_rule'] = 'AF only'
    else:
        rules['reaktif_tenzil_rule'] = 'AD + AF sum'

    return rules


def extract_file_filter(region_name, rules, filename):
    """Region-specific file filter logic."""
    if rules.get('akedas_tl_raporu'):
        return 'TL_Raporu' in filename

    if rules.get('uludağ_specific'):
        return '4008' in filename and 'TL' in filename and 'KWH' not in filename.upper()

    return True
