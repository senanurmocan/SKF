"""
ETSO Record Aggregation, Field Summation/Max Rules, and Regional Post-Processing.
"""

from collections import OrderedDict
from config.constants import SUM_FIELDS, MAX_FIELDS, TEXT_FIELDS_AGG
from core.normalizer import clean_turkish_number, normalize_etso_kodu


def post_process_record(r, reg_name):
    """Apply region-specific formatting rules after aggregation."""
    rec = r.copy()

    # 1. Tarife Grubu
    tarife = str(rec.get('Tarife Grubu') or '').strip()
    if reg_name != 'Boğaziçi EDAŞ':
        for suffix in [' (Ek)', ' (ek)', ' ek', ' (EK)']:
            if tarife.endswith(suffix):
                tarife = tarife[:-len(suffix)].strip()

    if reg_name in ['ADM EDAŞ', 'Gediz EDAŞ']:
        if tarife.upper().startswith('TICARET') or tarife.upper().startswith('TİCARET'):
            tarife = 'TICARETHAN'
        elif tarife.upper().startswith('SANAYI') or tarife.upper().startswith('SANAYİ'):
            tarife = 'SANAYI'
    elif reg_name == 'Çamlıbel EDAŞ':
        if 'Ticarethane' in tarife or 'TİCARETHANE' in tarife or 'ticarethane' in tarife or '12-' in tarife:
            tarife = 'TİCARETHANE'
        elif 'Sanayi' in tarife or 'SANAYİ' in tarife:
            tarife = 'SANAYİ'
    rec['Tarife Grubu'] = tarife

    # 2. AG OG
    ag_og = str(rec.get('AG OG') or '').strip()
    if reg_name not in ['Başkent EDAŞ', 'Toroslar EDAŞ', 'Gediz EDAŞ', 'ADM EDAŞ', 'Trakya EDAŞ']:
        ag_og_upper = ag_og.upper()
        if 'AG' in ag_og_upper or 'DAG' in ag_og_upper:
            ag_og = 'AG'
        elif 'OG' in ag_og_upper or 'DOG' in ag_og_upper:
            ag_og = 'OG'
    rec['AG OG'] = ag_og

    # 3. TERİM
    terim = str(rec.get('TERİM') or '').strip()
    if reg_name == 'Yeşilırmak EDAŞ':
        terim = ''
    elif reg_name == 'Boğaziçi EDAŞ':
        if rec.get('Tarife Grubu') == 'Boş' or (abs(float(rec.get('Aktif Enerji Tüketim (kWh)', 0) or 0)) < 1e-4 and abs(float(rec.get('Dağıtım Bedeli(TL)', 0) or 0)) < 1e-4):
            terim = ''
        elif 'tek' in terim.lower():
            terim = 'Tek Terimli'
        elif 'cift' in terim.lower() or 'çift' in terim.lower():
            terim = 'Çift Terimli'
        else:
            terim = ''
    elif reg_name == 'Osmangazi EDAŞ':
        if 'cift terim' in terim.lower() or 'çift terim' in terim.lower():
            terim = 'Çift Terim-Tek Zamanlı    '
        elif 'tek terim' in terim.lower():
            terim = 'Tek Terim-Tek Zamanlı    '
    elif reg_name == 'Çamlıbel EDAŞ':
        if 'Tek' in terim or '12-' in terim or 'TekTerim' in terim:
            terim = 'Tek Terimli'
        elif 'Çift' in terim:
            terim = 'Çift Terimli'
    else:
        if 'Tek Terim' in terim:
            terim = 'Tek Terimli'
        elif 'Çift Terim' in terim:
            terim = 'Çift Terimli'
    rec['TERİM'] = terim

    # 4. Müşteri
    if reg_name == 'Çamlıbel EDAŞ':
        rec['Müşteri'] = ''

    return rec


def aggregate_etso_records(all_data):
    """
    Group records by (ETSO Code, Region).
    SUM: Numeric monetary & consumption fields
    MAX: Contract power & Installed capacity (never summed)
    TEXT: First non-empty string value
    """
    aggregation_groups = OrderedDict()

    for record in all_data:
        etso_raw = record.get('Etso Kodu', '')
        etso_norm = normalize_etso_kodu(etso_raw) or str(etso_raw or '').strip()
        region = record.get('Dağıtım Bölgesi', '')
        key = (etso_norm, region)

        if key not in aggregation_groups:
            aggregation_groups[key] = []
        aggregation_groups[key].append(record)

    aggregated_data = []
    aggregation_count = 0

    for (etso_norm, region), records in aggregation_groups.items():
        if len(records) == 1:
            record = post_process_record(records[0], region)
            aggregated_data.append(record)
            continue

        aggregation_count += len(records) - 1
        merged = records[0].copy()

        # SUM fields
        for field in SUM_FIELDS:
            total = 0.0
            for r in records:
                val = r.get(field, 0)
                if isinstance(val, str) and val.upper() == 'X':
                    continue  # Sakarya X values skip summation
                try:
                    total += float(val or 0)
                except (ValueError, TypeError):
                    pass
            merged[field] = total

        # MAX fields
        for field in MAX_FIELDS:
            max_val = 0.0
            max_raw = None
            for r in records:
                val = r.get(field, 0)
                if val is not None and str(val).strip() and str(val).strip().lower() not in ('none', 'boş'):
                    v_float = clean_turkish_number(val, region_name=region, field_name=field) if isinstance(val, str) else float(val or 0)
                    if v_float > max_val or max_raw is None:
                        max_val = v_float
                        max_raw = val
            merged[field] = max_raw if max_raw is not None else max_val

        # TEXT fields
        for field in TEXT_FIELDS_AGG:
            if field == 'Tarife Grubu' and region == 'Boğaziçi EDAŞ':
                non_empty = [r.get(field) for r in records if r.get(field) and str(r.get(field)).strip().lower() not in ('none', 'boş')]
                akt = float(merged.get('Aktif Enerji Tüketim (kWh)', 0) or 0)
                dag = float(merged.get('Dağıtım Bedeli(TL)', 0) or 0)
                if not non_empty or (akt <= 0 and dag <= 0) or (abs(akt) < 1e-4 and abs(dag) < 1e-4):
                    merged['Tarife Grubu'] = 'Boş'
                    merged['TERİM'] = ''
                else:
                    merged['Tarife Grubu'] = non_empty[0]
            elif field == 'TERİM' and region == 'Boğaziçi EDAŞ':
                if merged.get('Tarife Grubu') == 'Boş':
                    merged['TERİM'] = ''
                else:
                    non_empty = [r.get('TERİM') for r in records if r.get('TERİM') and str(r.get('TERİM')).strip() and str(r.get('TERİM')).strip().lower() not in ('none', 'boş', '0')]
                    merged['TERİM'] = non_empty[0] if non_empty else 'Tek Terimli'
            else:
                for r in records:
                    val = r.get(field)
                    if val is not None and str(val).strip() and str(val).strip().lower() not in ('none', 'boş'):
                        merged[field] = val
                        break

        if region == 'Yeşilırmak EDAŞ' and abs(merged.get('Aktif Enerji Tüketim (kWh)', 0)) < 1e-4:
            pos_akt = sum(float(r.get('Aktif Enerji Tüketim (kWh)', 0) or 0) for r in records if float(r.get('Aktif Enerji Tüketim (kWh)', 0) or 0) > 0)
            if pos_akt > 0:
                merged['Aktif Enerji Tüketim (kWh)'] = pos_akt

        merged = post_process_record(merged, region)
        aggregated_data.append(merged)

    return aggregated_data, aggregation_count
