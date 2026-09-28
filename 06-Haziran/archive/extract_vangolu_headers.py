from pathlib import Path
import xml.etree.ElementTree as ET
import re

filepath = Path(__file__).resolve().parents[1] / 'Vangölü EDAŞ' / 'STDetay_202606_4008.xls'

# XML namespace'lerini tanımla
ns = {
    'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
    'html': 'http://www.w3.org/TR/REC-html40'
}

with open(filepath, 'r', encoding='windows-1254') as f:
    content = f.read()

# İlk 50.000 karakteri al (yeterli olmalı)
content = content[:50000]

# Worksheet ve Row elementlerini bul
# XML'de headers genellikle ilk Worksheet'in ilk Row'u
pattern = r'<Worksheet[^>]*ss:Name="([^"]*)"[^>]*>(.*?)</Worksheet>'
matches = re.findall(pattern, content, re.DOTALL)

print(f"Bulunan Worksheet sayısı: {len(matches)}")

if matches:
    for ws_name, ws_content in matches[:1]:  # İlk worksheet'i işle
        print(f"\n=== Worksheet: {ws_name} ===")
        
        # İlk birkaç Row'u bul
        row_pattern = r'<Row[^>]*>(.*?)</Row>'
        rows = re.findall(row_pattern, ws_content, re.DOTALL)
        
        if rows:
            # İlk row (headers)
            print("\n--- Header Row (First Row) ---")
            cells = re.findall(r'<Cell[^>]*>(.*?)</Cell>', rows[0], re.DOTALL)
            headers = []
            for cell in cells[:15]:  # İlk 15 hücreyi al
                # Data içeriğini çıkar
                data_match = re.search(r'<Data[^>]*>(.*?)</Data>', cell, re.DOTALL)
                if data_match:
                    # HTML tags temizle
                    text = re.sub(r'<[^>]+>', '', data_match.group(1)).strip()
                    headers.append(text)
                else:
                    headers.append('')
            
            print(f"Headers ({len(headers)} cells):")
            for i, h in enumerate(headers):
                print(f"  [{i}] {h}")
            
            # İlk data row
            if len(rows) > 1:
                print("\n--- First Data Row (Second Row) ---")
                cells = re.findall(r'<Cell[^>]*>(.*?)</Cell>', rows[1], re.DOTALL)
                data_values = []
                for cell in cells[:15]:
                    data_match = re.search(r'<Data[^>]*>(.*?)</Data>', cell, re.DOTALL)
                    if data_match:
                        text = re.sub(r'<[^>]+>', '', data_match.group(1)).strip()
                        data_values.append(text)
                    else:
                        data_values.append('')
                
                print(f"Data values ({len(data_values)} cells):")
                for i, v in enumerate(data_values):
                    print(f"  [{i}] {v}")
