#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SKF Başlıkları Güncel.xlsx dosyasını analiz eder
- Başlıklar sayfasını okur
- Notlar sayfasını okur
- Header ve sütun yapısını analiz eder
"""
from pathlib import Path

import pandas as pd
from collections import defaultdict

def analyze_mapping_file():
    """SKF Başlıkları Güncel.xlsx dosyasını analiz eder"""
    
    file_path = Path(__file__).resolve().parents[1] / 'SKF Başlıkları Güncel.xlsx'
    
    print("=" * 80)
    print("SKF BAŞLIKLARI GÜNCEL.XLSX ANALİZİ")
    print("=" * 80)
    print()
    
    # Excel dosyasını oku
    xls = pd.ExcelFile(file_path)
    
    print("SAYFA İSİMLERİ:")
    for sheet_name in xls.sheet_names:
        print(f"  - {sheet_name}")
    print()
    
    # Başlıklar sayfasını oku
    if "Başlıklar" in xls.sheet_names:
        print("=" * 80)
        print("BAŞLIKLAR SAYFASI")
        print("=" * 80)
        
        df = pd.read_excel(file_path, sheet_name="Başlıklar", header=None)
        
        print(f"Toplam satır sayısı: {len(df)}")
        print(f"Toplam sütun sayısı: {len(df.columns)}")
        print()
        
        # İlk satırı (header) göster
        print("İLK SATIR (Header):")
        print("-" * 80)
        for idx, col in enumerate(df.columns):
            value = df.iloc[0, idx] if idx < len(df.columns) else ""
            print(f"  Sütun {idx}: {value}")
        print()
        
        # Tüm başlıkları göster
        print("TÜM BAŞLIKLAR (İlk 5 satır):")
        print("-" * 80)
        for row_idx in range(min(5, len(df))):
            print(f"\nSatır {row_idx}:")
            for col_idx in range(len(df.columns)):
                value = df.iloc[row_idx, col_idx]
                if pd.notna(value):
                    print(f"  [{col_idx}] {value}")
        print()
        
        # Bölge isimlerini bul
        print("DAĞITIM BÖLGELERİ:")
        print("-" * 80)
        regions = []
        for row_idx in range(1, len(df)):
            first_cell = df.iloc[row_idx, 0] if len(df.columns) > 0 else None
            if pd.notna(first_cell):
                regions.append(str(first_cell))
                print(f"  {len(regions)}. {first_cell}")
        print()
        
        # Header mappings için detaylı analiz
        print("HEADER MAPPING YAPISI:")
        print("-" * 80)
        headers = df.iloc[0].tolist() if len(df) > 0 else []
        print(f"Header satırı ({len(headers)} sütun):")
        for idx, header in enumerate(headers):
            if pd.notna(header):
                print(f"  [{idx}] {header}")
        print()
    
    # Notlar sayfasını oku
    if "Notlar" in xls.sheet_names:
        print("=" * 80)
        print("NOTLAR SAYFASI")
        print("=" * 80)
        
        df_notes = pd.read_excel(file_path, sheet_name="Notlar", header=None)
        
        print(f"Toplam satır sayısı: {len(df_notes)}")
        print(f"Toplam sütun sayısı: {len(df_notes.columns)}")
        print()
        
        # Tüm notları göster
        print("NOTLAR İÇERİĞİ:")
        print("-" * 80)
        for row_idx in range(len(df_notes)):
            row_values = []
            for col_idx in range(len(df_notes.columns)):
                value = df_notes.iloc[row_idx, col_idx]
                if pd.notna(value):
                    row_values.append(str(value))
            if row_values:
                print(f"Satır {row_idx}: {' | '.join(row_values)}")
        print()
    
    # Tüm sheet'leri tek tek oku
    print("=" * 80)
    print("TÜM SAYFALARIN DETAYLI İNCELENMESİ")
    print("=" * 80)
    
    for sheet_name in xls.sheet_names:
        print(f"\n{'=' * 80}")
        print(f"SAYFA: {sheet_name}")
        print(f"{'=' * 80}")
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"Satır: {len(df)}, Sütun: {len(df.columns)}")
        print(f"Sütun isimleri: {list(df.columns)}")
        
        if len(df) > 0:
            print("\nİlk 3 satır:")
            print(df.head(3).to_string())
    
    return xls

if __name__ == "__main__":
    analyze_mapping_file()
