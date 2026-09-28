from pathlib import Path
import pandas as pd

ref = pd.read_excel(Path(__file__).resolve().parents[1] / 'Dağıtımın Kestiği Faturalar Özet.xlsx')
extracted = pd.read_excel(Path(__file__).resolve().parents[1] / 'Çıkarılan_Veriler.xlsx')

print('=== REGION COMPARISON ===')
ref_counts = ref.groupby('Dağıtım Bölgesi').size()
ext_counts = extracted.groupby('Dağıtım Bölgesi').size()
all_regions = sorted(set(ref_counts.index) | set(ext_counts.index))

for region in all_regions:
    r = ref_counts.get(region, 0)
    e = ext_counts.get(region, 0)
    diff = r - e
    status = '✓' if diff == 0 else '✗'
    print(f'{status} {region}: Ref={r}, Extracted={e}, Diff={diff}')

print(f'\nTotal - Ref: {len(ref)}, Extracted: {len(extracted)}')
