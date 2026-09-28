import openpyxl
import re

def normalize_region_name(region, to_format='reference'):
    """Bölge isimlerini normalleştir"""
    normalization = {
        'aydem edaş': 'AYEDAŞ', 'aydem': 'AYEDAŞ',
        'akdeniz edaş': 'Akdeniz EDAŞ', 'akdeniz': 'Akdeniz EDAŞ',
        'akedaş': 'AKEDAŞ',
        'aras edaş': 'Aras EDAŞ', 'aras': 'Aras EDAŞ',
        'boğaziçi edaş': 'Boğaziçi EDAŞ', 'boğaziçi': 'Boğaziçi EDAŞ',
        'çamlıbel edaş': 'Çamlıbel EDAŞ', 'çamlıbel': 'Çamlıbel EDAŞ',
        'çoruh edaş': 'Çoruh EDAŞ', 'çoruh': 'Çoruh EDAŞ',
        'dicle edaş': 'Dicle EDAŞ', 'dicle': 'Dicle EDAŞ',
        'firat edaş': 'Fırat EDAŞ', 'firat': 'Fırat EDAŞ',
        'gediz edaş': 'Gediz EDAŞ', 'gediz': 'Gediz EDAŞ',
        'kayseri': 'Kayseri ve Civarı', 'kayseri ve civarı': 'Kayseri ve Civarı',
        'meram edaş': 'Meram EDAŞ', 'meram': 'Meram EDAŞ',
        'osmangazi edaş': 'Osmangazi EDAŞ', 'osmangazi': 'Osmangazi EDAŞ',
        'sakarya edaş': 'Sakarya EDAŞ', 'sakarya': 'Sakarya EDAŞ',
        'toroslar edaş': 'Toroslar EDAŞ', 'toroslar': 'Toroslar EDAŞ',
        'trakya edaş': 'Trakya EDAŞ', 'trakya': 'Trakya EDAŞ',
        'uludağ edaş': 'Uludağ EDAŞ', 'uludağ': 'Uludağ EDAŞ',
        'vangölü edaş': 'Vangölü EDAŞ', 'vangölü': 'Vangölü EDAŞ',
        'yesilirmak edaş': 'Yeşilırmak EDAŞ', 'yeşilırmak': 'Yeşilırmak EDAŞ',
        'adm edaş': 'ADM EDAŞ', 'adm': 'ADM EDAŞ',
    }
    
    region = str(region).strip().lower()
    region = re.sub(r'[^\w\s]', '', region)
    
    if to_format == 'reference':
        return normalization.get(region, region.title() if ' ' in region else region.upper())
    else:
        return normalization.get(region, region.title() if ' ' in region else region.upper())

# Referans dosyayı yükle
wb_ref = openpyxl.load_workbook('Dağıtımın Kestiği Faturalar Özet.xlsx')
ws_ref = wb_ref['Dağıtımın Kestiği']
ref_rows = list(ws_ref.iter_rows(values_only=True))

reference_data = []
for row in ref_rows[1:]:
    if row[0] and row[0] != 'Dağıtım Bölgesi':
        region = normalize_region_name(row[0], to_format='extraction')
        reference_data.append({
            'Dağıtım Bölgesi': region,
            'Etso Kodu': str(row[1]) if row[1] else '',
            'Müşteri': row[2],
            'Dağıtım Bedeli(TL)': row[9]
        })

# Çıkarılan veriyi yükle
wb_ext = openpyxl.load_workbook('Çıkarılan_Veriler.xlsx')
ws_ext = wb_ext.active
ext_rows = list(ws_ext.iter_rows(values_only=True))

extracted_data = []
for row in ext_rows[1:]:
    extracted_data.append({
        'Dağıtım Bölgesi': row[0],
        'Etso Kodu': str(row[1]) if row[1] else '',
        'Müşteri': row[2],
        'Dağıtım Bedeli(TL)': row[9]
    })

# Karşılaştırma anahtarını oluştur
def make_key(row):
    musteri = '' if row['Müşteri'] is None else str(row['Müşteri']).strip()
    return (row['Etso Kodu'], musteri, row['Dağıtım Bedeli(TL)'])

ref_dict = {make_key(r): r for r in reference_data}
ext_dict = {make_key(r): r for r in extracted_data}

# Eşleşmeleri bul
matches = sum(1 for key in ref_dict if key in ext_dict)

print(f"Referans satır sayısı: {len(reference_data)}")
print(f"Çıkarılan satır sayısı: {len(extracted_data)}")
print(f"Eşleşen anahtar sayısı: {matches}")
print(f"Çıkarılan dosyada EKSİK: {len(reference_data) - matches}")
print(f"Çıkarılan dosyada FAZLA: {len(extracted_data) - matches}")

print("\n--- EKSİK KAYITLARIN ÖRNEKLERİ (Referans'da var, Çıkarılan'da yok) ---")
count = 0
for key, ref_row in list(ref_dict.items())[:50]:
    if key not in ext_dict:
        print(f"{ref_row['Dağıtım Bölgesi']} | Etso: {ref_row['Etso Kodu']} | Müşteri: {ref_row['Müşteri']} | Bedel: {ref_row['Dağıtım Bedeli(TL)']}")
        count += 1
    if count >= 10:
        break

print("\n--- FAZLA KAYITLARIN ÖRNEKLERİ (Çıkarılan'da var, Referans'ta yok) ---")
count = 0
for key, ext_row in list(ext_dict.items())[:100]:
    if key not in ref_dict:
        print(f"{ext_row['Dağıtım Bölgesi']} | Etso: {ext_row['Etso Kodu']} | Müşteri: {ext_row['Müşteri']} | Bedel: {ext_row['Dağıtım Bedeli(TL)']}")
        count += 1
    if count >= 10:
        break
