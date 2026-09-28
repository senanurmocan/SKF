from pathlib import Path
import xml.etree.ElementTree as ET

filepath = Path(__file__).resolve().parents[1] / 'Vangölü EDAŞ' / 'STDetay_202606_4008.xls'

# XML namespace'lerini tanımla
ns = {
    'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
    'html': 'http://www.w3.org/TR/REC-html40'
}

with open(filepath, 'r', encoding='windows-1254') as f:
    content = f.read()

# XML parse et - error handling ile
try:
    root = ET.fromstring(content)
    print("XML parsed successfully")
    
    # Workbook elementini bul
    # Workbook -> Worksheet -> Table -> Row -> Cell -> Data
    
    worksheets = root.findall('.//ss:Worksheet', ns)
    print(f"Bulunan Worksheet sayısı: {len(worksheets)}")
    
    if worksheets:
        ws = worksheets[0]
        
        # Worksheet name
        ws_name = ws.get('{urn:schemas-microsoft-com:office:spreadsheet}Name', 'No Name')
        print(f"\nWorksheet name: {ws_name}")
        
        # Table elementini bul
        table = ws.find('.//ss:Table', ns)
        if table:
            rows = table.findall('.//ss:Row', ns)
            print(f"Bulunan Row sayısı: {len(rows)}")
            
            if len(rows) >= 2:
                # First row (headers)
                print("\n--- Header Row ---")
                header_cells = rows[0].findall('.//ss:Cell', ns)
                print(f"Header cell sayısı: {len(header_cells)}")
                
                headers = []
                for i, cell in enumerate(header_cells[:20]):
                    data = cell.find('.//ss:Data', ns)
                    if data is not None and data.text:
                        headers.append(data.text.strip())
                    else:
                        # Merge cell kontrolü
                        merged = cell.get('{urn:schemas-microsoft-com:office:spreadsheet}MergeAcross', '0')
                        headers.append(f"<merged:{merged}>")
                
                print(f"\nHeaders ({len(headers)}):")
                for i, h in enumerate(headers):
                    print(f"  [{i}] {h}")
                
                # Second row (first data)
                print("\n--- First Data Row ---")
                data_cells = rows[1].findall('.//ss:Cell', ns)
                print(f"Data cell sayısı: {len(data_cells)}")
                
                data_values = []
                for i, cell in enumerate(data_cells[:20]):
                    data = cell.find('.//ss:Data', ns)
                    if data is not None and data.text:
                        data_values.append(data.text.strip())
                    else:
                        merged = cell.get('{urn:schemas-microsoft-com:office:spreadsheet}MergeAcross', '0')
                        data_values.append(f"<merged:{merged}>")
                
                print(f"\nData values ({len(data_values)}):")
                for i, v in enumerate(data_values):
                    print(f"  [{i}] {v}")
                    
                if len(data_values) > 0:
                    # Total records check
                    print(f"\nTotal records (excluding header): {len(rows) - 1}")
                    
except ET.ParseError as e:
    print(f"XML Parse Error: {e}")
    # Parse error olursa, satır satır parse etmeyi dene
    print("\nTrying line-by-line parsing...")
    lines = content.split('\n')
    for i, line in enumerate(lines[0:20]):
        print(f"Line {i}: {line[:150]}")
