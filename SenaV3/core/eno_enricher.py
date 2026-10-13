# -*- coding: utf-8 -*-
"""
ENO (Enerji Piyasası Operatörü) dosyasından veri zenginleştirme modülü.

Bu modül, çıktı kayıtlarındaki Sayaç ID'lerini ENO dosyasındaki eşleştirmeleri
kullanarak EIC kodlarına dönüştürür ve eksik Müşteri/Abone bilgilerini doldurur.
"""

import os
import re
import openpyxl


def _find_eno_file(base_path):
    """Veri klasöründe 'XX ENO AyAdı' formatında ENO dosyasını otomatik bul.
    
    Örnek: '01 ENO Ocak.xlsx', '06 ENO Haziran.xlsx'
    """
    import re
    if not base_path or not os.path.isdir(base_path):
        return None
    
    eno_pattern = re.compile(r'^\d{2}\s+ENO\s+', re.IGNORECASE)
    
    for fname in os.listdir(base_path):
        if eno_pattern.match(fname) and (fname.endswith('.xlsx') or fname.endswith('.xls')):
            fpath = os.path.join(base_path, fname)
            if os.path.isfile(fpath) and not fname.startswith('~$'):
                return fpath
    return None


def _load_eno_lookups(eno_path):
    """ENO dosyasından iki lookup tablosu oluştur:
    
    1. sayac_id_to_eic: Sayaç ID (E sütunu) → Sayaç EIC Kod (D sütunu)
    2. eic_to_abone: Sayaç EIC Kod (D sütunu) → Abone Ad-Soyad/Unvan (AD sütunu)
    
    Satır 1: üst başlık (merged), Satır 2: gerçek başlık, Satır 3+: veri
    """
    sayac_id_to_eic = {}  # str(Sayaç ID) → str(EIC Kod)
    eic_to_abone = {}     # str(EIC Kod) → str(Abone Ad-Soyad)
    
    try:
        wb = openpyxl.load_workbook(eno_path, read_only=False, data_only=True)
        ws = wb.active
        
        row_count = 0
        for row in ws.iter_rows(min_row=3, max_col=35, values_only=False):  # Satır 3'ten başla (veri)
            row_count += 1
            # D sütunu (col 4, 1-indexed) = Sayaç EIC Kod
            # E sütunu (col 5, 1-indexed) = Sayaç ID
            # AD sütunu (col 30, 1-indexed) = Abone Ad-Soyad / Unvan
            d_val = None
            e_val = None
            ad_val = None
            
            for cell in row:
                if cell.column == 4:
                    d_val = cell.value
                elif cell.column == 5:
                    e_val = cell.value
                elif cell.column == 30:
                    ad_val = cell.value
            
            if d_val is not None:
                eic_str = str(d_val).strip()
                
                # Sayaç ID → EIC Kod eşleştirmesi
                if e_val is not None:
                    # Sayaç ID'yi normalize et (float'dan int'e çevir: 20088.0 → "20088")
                    e_str = str(e_val).strip()
                    if '.' in e_str:
                        try:
                            e_str = str(int(float(e_str)))
                        except (ValueError, TypeError):
                            pass
                    sayac_id_to_eic[e_str] = eic_str
                
                # EIC Kod → Abone Ad-Soyad eşleştirmesi
                if ad_val is not None:
                    ad_str = str(ad_val).strip()
                    if ad_str:
                        eic_to_abone[eic_str] = ad_str
        
        wb.close()
        print(f"  [ENO] {row_count} satır okundu: {len(sayac_id_to_eic)} Sayaç ID→EIC, {len(eic_to_abone)} EIC→Abone")
        
    except Exception as e:
        print(f"  [ENO HATA] {eno_path}: {e}")
    
    return sayac_id_to_eic, eic_to_abone


def enrich_from_eno_file(records, base_path):
    """Çıktı kayıtlarını ENO dosyasından zenginleştir:
    
    1. ETSO Kodu '40Z' ile başlamıyorsa → ENO'daki Sayaç ID eşleşmesiyle EIC Kod'a çevir
    2. Müşteri alanı boşsa → EIC Kod ile ENO'dan Abone Ad-Soyad/Unvan getir
    
    Returns:
        dict: Zenginleştirme istatistikleri
    """
    stats = {'eno_file': None, 'etso_converted': 0, 'musteri_filled': 0, 'total_records': len(records)}
    
    eno_path = _find_eno_file(base_path)
    if not eno_path:
        print("[ENO] Veri klasöründe ENO dosyası bulunamadı (format: 'XX ENO AyAdı.xlsx')")
        return stats
    
    stats['eno_file'] = os.path.basename(eno_path)
    print(f"\n[ENO] Dosya bulundu: {stats['eno_file']}")
    
    sayac_id_to_eic, eic_to_abone = _load_eno_lookups(eno_path)
    
    if not sayac_id_to_eic and not eic_to_abone:
        print("[ENO] Lookup tabloları boş, zenginleştirme yapılamadı")
        return stats
    
    # ADIM 1: ETSO Kodu dönüşümü (Sayaç ID → EIC Kod)
    for record in records:
        etso = record.get('Etso Kodu', '')
        if etso is None:
            etso = ''
        etso_str = str(etso).strip()
        
        # Zaten 40Z formatındaysa dokunma
        if etso_str.startswith('40Z') and len(etso_str) >= 16:
            continue
        
        # Sayaç ID olarak ENO'da ara
        # Float'tan int'e normalize et (20088.0 → "20088")
        lookup_key = etso_str
        if '.' in lookup_key:
            try:
                lookup_key = str(int(float(lookup_key)))
            except (ValueError, TypeError):
                pass
        
        if lookup_key in sayac_id_to_eic:
            new_etso = sayac_id_to_eic[lookup_key]
            record['Etso Kodu'] = new_etso
            stats['etso_converted'] += 1
    
    # ADIM 2: Boş Müşteri doldurma (EIC Kod → Abone Ad-Soyad)
    for record in records:
        musteri = record.get('Müşteri', '')
        if musteri is not None and str(musteri).strip():
            continue  # Zaten dolu, dokunma
        
        etso = record.get('Etso Kodu', '')
        if etso is None:
            continue
        etso_str = str(etso).strip()
        
        if etso_str in eic_to_abone:
            record['Müşteri'] = eic_to_abone[etso_str]
            stats['musteri_filled'] += 1
    
    print(f"  [ENO] ETSO dönüştürülen: {stats['etso_converted']}, Müşteri doldurulan: {stats['musteri_filled']}")
    return stats
