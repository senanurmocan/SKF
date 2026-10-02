# -*- coding: utf-8 -*-
"""
Column resolution and row value extraction module for EDAS Bill Extraction & Analysis System.

This module provides functions to dynamically resolve column indices from headers,
handle region-specific header overrides and synonyms, and extract/clean values
from Excel row data.
"""

import re
import unicodedata
import pandas as pd

from core.number_cleaner import normalize_header, clean_turkish_number
from config.mappings import (
    FIELD_SYNONYMS,
    KNOWN_ADJACENT_HEADER_SHIFTS,
    SPECIAL_HEADER_MAPPING,
)
from config.constants import EXTENDED_SUM_FIELDS

from core.normalizer import normalize_region_name


def find_column_index(headers, target_headers, region=None, field_name=None):
    """Header listesinden hedef header'ı bul"""
    # Special handling for region-specific mappings
    if region and field_name and region in SPECIAL_HEADER_MAPPING:
        special_field = SPECIAL_HEADER_MAPPING[region].get(field_name)
        if special_field:
            normalized_headers = [normalize_header(h) for h in headers]
            normalized_target = normalize_header(special_field)
            if normalized_target in normalized_headers:
                return normalized_headers.index(normalized_target)
    
    normalized_headers = [normalize_header(h) for h in headers]
    for target in target_headers:
        normalized_target = normalize_header(target)
        if normalized_target in normalized_headers:
            return normalized_headers.index(normalized_target)
    return -1


def resolve_column_indices(headers, region_mapping):
    """
    Header satırına göre sütun indekslerini dinamik olarak çözümler.
    1. Mapping dosyasında tanımlı başlık ismi eşleştirilir.
    2. Eşleşmezse bilinen eş anlamlı başlıklar (FIELD_SYNONYMS) denenir.
    3. Eşleşmezse sabit sütun harfi indeksine düşer.
    """
    fields = ['etso', 'musteri', 'tarife', 'ag_og', 'terim', 'güç_kw', 'kurulu_güç',
              'aktif_enerji', 'trafo_kaybı', 'dagitim_bedeli', 'güç_bedeli',
              'güç_aşım', 'reaktif', 'reaktif2', 'reaktif_tenzil', 'reaktif_tenzil2']
    for field in (
        list(region_mapping.get('tazminat_fields', []))
        + list(region_mapping.get('standart_disi_fields', []))
    ):
        if field not in fields:
            fields.append(field)

    if not headers:
        return {field: region_mapping.get(f'{field}_index', -1) for field in fields}

    normalized_headers = [normalize_header(h) for h in headers]
    resolved = {}

    norm_reg = normalize_region_name(region_mapping.get('region_name', ''), to_format='reference')
    force_fixed_fields = set(region_mapping.get('rules', {}).get('force_fixed_fields', []))

    for field in fields:
        header_name = region_mapping.get(field)
        fixed_idx = region_mapping.get(f'{field}_index', -1)
        found_idx = -1

        shifted_idx = KNOWN_ADJACENT_HEADER_SHIFTS.get(norm_reg, {}).get(fixed_idx)
        if (
            shifted_idx is not None
            and 0 <= shifted_idx < len(headers)
            and header_name
            and normalize_header(headers[shifted_idx]) == normalize_header(header_name)
        ):
            found_idx = shifted_idx

        # Notlar 7 / Çamlıbel: aynı isimli iki Tarife başlığından AV zorunlu.
        if found_idx == -1 and field in force_fixed_fields and 0 <= fixed_idx < len(headers):
            found_idx = fixed_idx

        # Check SPECIAL_HEADER_MAPPING first (highest priority for region-specific overrides)
        if found_idx == -1 and norm_reg in SPECIAL_HEADER_MAPPING and field in SPECIAL_HEADER_MAPPING[norm_reg]:
            sp_val = SPECIAL_HEADER_MAPPING[norm_reg][field]
            if sp_val is None:
                found_idx = -2  # Explicitly disabled
            elif isinstance(sp_val, int):
                found_idx = sp_val
            else:
                norm_sp = normalize_header(sp_val)
                if norm_sp in normalized_headers:
                    found_idx = normalized_headers.index(norm_sp)

        if found_idx == -1 and header_name and pd.notna(header_name) and str(header_name).strip() != '-':
            norm_target = normalize_header(header_name)
            if norm_target and norm_target in normalized_headers:
                found_idx = normalized_headers.index(norm_target)

        # Prioritize explicitly assigned column index from mapping if header name didn't match directly
        if found_idx == -1 and 0 <= fixed_idx < len(headers):
            found_idx = fixed_idx

        if found_idx == -1:
            for syn in FIELD_SYNONYMS.get(field, []):
                if syn in normalized_headers:
                    found_idx = normalized_headers.index(syn)
                    break

        if found_idx == -2:
            found_idx = -1

        resolved[field] = found_idx

    return resolved


def resolve_reactive_field_groups(region_mapping, region_name):
    """Reaktif ve İlk Reaktif kaynak alanlarını bölgesel kurallarla seç."""
    reactive_rule = region_mapping.get('rules', {}).get('reactive_total_rule')
    norm_rn = normalize_region_name(region_name or '', to_format='reference')
    no_tenzil_regions = [
        'Aras EDAŞ', 'Çoruh EDAŞ', 'Dicle EDAŞ', 'Fırat EDAŞ', 'Vangölü EDAŞ', 'Meram EDAŞ'
    ]

    if reactive_rule in {'single_column', 'base_columns_only'} or norm_rn in no_tenzil_regions:
        reactive_fields = ['reaktif', 'reaktif2']
    else:
        reactive_fields = ['reaktif', 'reaktif2', 'reaktif_tenzil', 'reaktif_tenzil2']

    if norm_rn in ['Trakya EDAŞ', 'Trakya Edaş', 'TREDAŞ']:
        first_reactive_fields = ['reaktif_tenzil2']
    else:
        first_reactive_fields = ['reaktif_tenzil', 'reaktif_tenzil2']

    return reactive_fields, first_reactive_fields


def sum_mapped_numeric_fields(row, indices, field_names, region_name, field_name):
    """Bir mapping alan listesindeki sayısal değerleri güvenle topla."""
    total = 0.0
    for mapped_field in field_names:
        idx = indices.get(mapped_field, -1)
        if 0 <= idx < len(row):
            total += clean_turkish_number(
                row[idx],
                region_name=region_name,
                field_name=field_name,
            )
    return total


def add_extended_financial_fields(data_row, row, indices, region_mapping, region_name):
    """Mapping 7'nin iki yeni nihai tutarını ham satır sözlüğüne ekle."""
    if not region_mapping.get('features', {}).get('extended_financial_fields'):
        return data_row

    data_row['Tazminat Bedeli'] = sum_mapped_numeric_fields(
        row,
        indices,
        region_mapping.get('tazminat_fields', []),
        region_name,
        'tazminat_bedeli',
    )
    data_row['Standart Dışı Tutar (TL)'] = sum_mapped_numeric_fields(
        row,
        indices,
        region_mapping.get('standart_disi_fields', []),
        region_name,
        'standart_disi_tutar',
    )
    return data_row


def normalize_correction_lookup(value):
    """Düzeltme lookup için NFKC, kontrollü boşluk ve casefold normalizasyonu."""
    if value is None:
        return ''
    normalized = unicodedata.normalize('NFKC', str(value)).strip()
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.casefold()


def extract_value_from_row(row, column_index, field_name=None, region_name=None):
    """Satırdan değeri çıkar"""
    if 0 <= column_index < len(row):
        value = row[column_index]
        if field_name == 'etso':
            if isinstance(value, str):
                return value.strip() if value.strip() else None
            if isinstance(value, (int, float)):
                return str(int(value)) if value == int(value) else str(value)
            return value if value else None

        norm_reg = normalize_region_name(region_name or '', to_format='reference') if region_name else ''
        if ('yeşilırmak' in str(region_name).lower() or 'yeşilırmak' in str(norm_reg).lower() or 'yesilirmak' in normalize_header(region_name)) and field_name == 'Güç kW':
            if isinstance(value, str):
                val_str = value.strip()
                if val_str and val_str.lower() not in ('boş', 'none', 'null'):
                    if val_str.endswith('-'):
                        val_str = '-' + val_str[:-1].strip()
                    return val_str

        return clean_turkish_number(value, region_name=region_name, field_name=field_name)
    return 0
