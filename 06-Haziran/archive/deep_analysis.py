#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Deep analysis for discrepancies
"""
from pathlib import Path

import openpyxl
import sys
from collections import defaultdict

BASE_PATH = Path(__file__).resolve().parents[1]
REFERENCE_FILE = BASE_PATH / 'Dağıtımın Kestiği Faturalar Özet.xlsx'
EXTRACTED_FILE = BASE_PATH / 'Çıkarılan_Veriler.xlsx'

def normalize_region_name(region_name):
    """Region ismini standartlaştırmak için fonksiyon"""
    if region_name is None:
        return None
    region_name = str(region_name).strip().upper()

    REGION_NORMALIZATION_MAPPING = {
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
        "Gediz EDAŞ": "Gediz EDAŞ",
        "Meram EDAŞ": "Meram EDAŞ",
        "Osmangazi EDAŞ": "Osmangazi EDAŞ",
        "Sakarya EDAŞ": "Sakarya EDAŞ",
        "Trakya EDAŞ": "Trakya EDAŞ",
        "Uludağ EDAŞ": "Uludağ EDAŞ",
        "Çamlıbel EDAŞ": "Çamlıbel EDAŞ",
        "Vangölü EDAŞ": "Vangölü Edaş",
        "Yeşilırmak EDAŞ": "Yeşilırmak EDAŞ",
        "Başkent EDAŞ": "Başkent EDAŞ",
        "Toroslar EDAŞ": "Toroslar EDAŞ",
        "Fırat EDAŞ": "Fırat Edaş",
        "Dicle EDAŞ": "Dicle Edaş",
        "Aras EDAŞ": "Aras EDAŞ",
        "Çoruh EDAŞ": "Çoruh EDAŞ",
    }

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

print("=== Region Mapping Issues ===")
# Check which regions don't map properly
wb = openpyxl.load_workbook(REFERENCE_FILE, data_only=True)
ws = wb.active
regions_in_ref = set()
for row in list(ws.iter_rows(values_only=True))[1:]:
    if row and row[0]:
        regions_in_ref.add(str(row[0]).strip())

print("Regions in reference that don't match mapping keys:")
unmapped = []
for r in regions_in_ref:
    normalized = normalize_region_name(r)
    if normalized != r.upper() and normalized not in ["AKEDAŞ ( Göksu EDAŞ )", "Ayedaş", "Aydem EDAŞ", "Dicle Edaş", "Fırat Edaş", "Vangölü Edaş", "Çoruh EDAŞ", "Aras EDAŞ"]:
        unmapped.append((r, normalized))

for r, n in sorted(unmapped)[:20]:
    print(f"  '{r}' -> '{n}'")

print("\n=== Check for negative values in extracted ===")
wb2 = openpyxl.load_workbook(EXTRACTED_FILE, data_only=True)
ws2 = wb2.active
negative_count = 0
negative_examples = []
for row in list(ws2.iter_rows(values_only=True))[1:]:
    if row and len(row) > 9:
        bedel = row[9]
        if bedel is not None and isinstance(bedel, (int, float)) and bedel < 0:
            negative_count += 1
            if len(negative_examples) < 5:
                negative_examples.append(row)
print(f"Negative values found: {negative_count}")
if negative_examples:
    print("Examples:")
    for ex in negative_examples:
        print(f"  Etso: {ex[1]}, Bedel: {ex[9]}, Region: {ex[0]}")

print("\n=== Check region normalization consistency ===")
# Check if 'Aydem EDAŞ' and 'ADM EDAŞ' both map to same thing
print(f"'Aydem EDAŞ' normalized: '{normalize_region_name('Aydem EDAŞ')}'")
print(f"'ADM EDAŞ' normalized: '{normalize_region_name('ADM EDAŞ')}'")
print(f"'AYDEM EDAŞ' normalized: '{normalize_region_name('AYDEM EDAŞ')}'")

print("\n=== Check if region names match after normalization ===")
wb = openpyxl.load_workbook(REFERENCE_FILE, data_only=True)
ws = wb.active
ref_regions_normalized = set()
for row in list(ws.iter_rows(values_only=True))[1:]:
    if row and row[0]:
        ref_regions_normalized.add(normalize_region_name(row[0]))

wb2 = openpyxl.load_workbook(EXTRACTED_FILE, data_only=True)
ws2 = wb2.active
ext_regions_normalized = set()
for row in list(ws2.iter_rows(values_only=True))[1:]:
    if row and row[0]:
        ext_regions_normalized.add(normalize_region_name(row[0]))

print(f"Reference normalized regions count: {len(ref_regions_normalized)}")
print(f"Extracted normalized regions count: {len(ext_regions_normalized)}")
print(f"Difference: {ref_regions_normalized - ext_regions_normalized}")

print("\n=== Sample data with normalized values ===")
wb = openpyxl.load_workbook(REFERENCE_FILE, data_only=True)
ws = wb.active
row = list(ws.iter_rows(values_only=True))[1]
print("Reference row 1:")
print(f"  Raw: Region='{row[0]}', Etso={row[1]}, Customer={row[2]}, Bedel={row[9]}")
print(f"  Normalized: Region='{normalize_region_name(row[0])}', Etso='{normalize_etso_kodu(row[1])}', Customer='{row[2]}', Bedel={row[9]}")

wb2 = openpyxl.load_workbook(EXTRACTED_FILE, data_only=True)
ws2 = wb2.active
row2 = list(ws2.iter_rows(values_only=True))[1]
print("\nExtracted row 1:")
print(f"  Raw: Region='{row2[0]}', Etso={row2[1]}, Customer={row2[2]}, Bedel={row2[9]}")
print(f"  Normalized: Region='{normalize_region_name(row2[0])}', Etso='{normalize_etso_kodu(row2[1])}', Customer='{row2[2]}', Bedel={row2[9]}")

print("\n=== Check Etso matching key ===")
# Check if Etso values match after normalization
etso1 = normalize_etso_kodu(7606443)
etso2 = normalize_etso_kodu('7606443')
print(f"Etso from reference (int 7606443): '{etso1}'")
print(f"Etso from extracted (str '7606443'): '{etso2}'")
print(f"Match: {etso1 == etso2}")
