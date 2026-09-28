"""
Find Dağıtım Bedeli column in all region Excel files and compare with mapping file.
"""
from pathlib import Path
import os
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string

def normalize_turkish(text):
    """Normalize Turkish characters for case-insensitive comparison."""
    if not text:
        return ""
    text = str(text).lower()
    text = text.replace('ı', 'i').replace('i', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ş', 's').replace('ö', 'o').replace('ç', 'c')
    return text

def find_column_with_keyword(df, keywords, method="header=0"):
    """Find column index containing any of the keywords."""
    if method == "header=0":
        headers = df.columns.tolist()
    elif method == "header=1":
        headers = df.iloc[0].tolist()
    else:
        headers = df.iloc[0].tolist()
    
    for idx, header in enumerate(headers):
        header_str = str(header).lower() if pd.notna(header) else ""
        for keyword in keywords:
            if keyword in normalize_turkish(header_str):
                return idx, header
    return None, None

def main():
    mapping_file = Path(__file__).resolve().parents[1] / 'SKF Başlıkları Güncel.xlsx'
    base_dir = Path(__file__).resolve().parents[1]
    
    # Load mapping file
    df_mapping = pd.read_excel(mapping_file, header=None)
    
    # Build mapping dictionary: normalized region name -> column letter
    region_mapping = {}
    for row_idx in range(1, len(df_mapping)):
        bolgesi_value = df_mapping.iloc[row_idx, 0]
        column_letter = df_mapping.iloc[row_idx, 20]  # Column index 20
        
        if bolgesi_value and pd.notna(column_letter) and str(column_letter).strip() != '-':
            normalized_bolgesi = normalize_turkish(str(bolgesi_value))
            region_mapping[normalized_bolgesi] = {
                'bolgesi': str(bolgesi_value),
                'column_letter': str(column_letter).strip().upper()
            }
    
    print("Mapping file predictions:")
    for region, info in sorted(region_mapping.items()):
        print(f"  {region}: column {info['column_letter']} (index {column_index_from_string(info['column_letter']) - 1})")
    print()
    
    # Check each region
    results = []
    
    for region_folder in sorted(os.listdir(base_dir)):
        if region_folder.startswith('.') or region_folder.startswith('~') or region_folder == '__pycache__':
            continue
        
        region_path = os.path.join(base_dir, region_folder)
        if not os.path.isdir(region_path):
            continue
        
        # Find Excel files
        excel_files = [f for f in os.listdir(region_path) if f.endswith('.xlsx') and not f.startswith('~')]
        
        if not excel_files:
            print(f"NO EXCEL: {region_folder}")
            continue
        
        normalized_region = normalize_turkish(region_folder)
        expected_info = region_mapping.get(normalized_region)
        
        for excel_file in excel_files:
            excel_path = os.path.join(region_path, excel_file)
            print(f"\nChecking: {region_folder} / {excel_file}")
            
            # Try different header configurations
            for header_config in [0, 1, None]:
                try:
                    if header_config is None:
                        df = pd.read_excel(excel_path, header=None)
                        headers = df.iloc[0].tolist()
                        method_str = "header=None (row 0)"
                    elif header_config == 0:
                        df = pd.read_excel(excel_path, header=0)
                        headers = df.columns.tolist()
                        method_str = "header=0"
                    else:
                        df = pd.read_excel(excel_path, header=1)
                        headers = df.iloc[0].tolist()
                        method_str = "header=1"
                    
                    # Search for Dağıtım Bedeli
                    col_idx, col_name = find_column_with_keyword(df, ['dagitim', 'bedeli'], method=method_str)
                    
                    if col_idx is not None:
                        print(f"  Found 'Dağıtım Bedeli' with {method_str}:")
                        print(f"    Column index: {col_idx}, Column name: {col_name}")
                        if expected_info:
                            expected_idx = column_index_from_string(expected_info['column_letter']) - 1
                            print(f"    Expected from mapping: index {expected_idx} (column {expected_info['column_letter']})")
                            if col_idx != expected_idx:
                                print(f"    MISMATCH: {col_idx - expected_idx} columns offset!")
                        break
                except Exception as e:
                    print(f"  Error with {method_str}: {e}")
            
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)

if __name__ == "__main__":
    main()
