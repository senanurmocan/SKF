"""
Dynamic Statistical Anomaly Detector & Data Quality Auditor for EDAS Bill Datasets.
Uses per-region median/quantile baselines to detect statistical bill spikes (5x-10x above region median)
without hardcoded threshold false alarms.
"""

import numpy as np
from core.normalizer import mask_etso

ABSURD_CORRUPTION_THRESHOLD = 50_000_000_000.0  # 50 Billion TL absolute data corruption safety cap


def audit_extracted_dataset(dataset):
    """
    Performs dynamic statistical auditing per distribution region.
    Calculates per-region spending baselines (median, p90) and identifies statistical bill spikes.
    """
    warnings = []
    absurd_records = []
    dynamic_spike_records = []
    negative_cost_records = []
    missing_customer_count = 0
    missing_etso_count = 0

    region_counts = {}
    region_values = {}

    # Step 1: Collect cost distributions per region
    for row in dataset:
        region = row.get('Dağıtım Bölgesi', 'Bilinmeyen Bölge')
        dagitim = float(row.get('Dağıtım Bedeli(TL)', 0) or 0)
        guc = float(row.get('Güç Bedeli(TL)', 0) or 0)
        toplam = float(row.get('Toplam (TL)', 0) or (dagitim + guc))
        val = abs(toplam if toplam > 0 else dagitim)

        region_counts[region] = region_counts.get(region, 0) + 1

        if val > 0 and val < ABSURD_CORRUPTION_THRESHOLD:
            if region not in region_values:
                region_values[region] = []
            region_values[region].append(val)

    # Step 2: Calculate per-region statistical baselines (Median & P90)
    region_baselines = {}
    for region, vals in region_values.items():
        if len(vals) >= 3:
            median_val = float(np.median(vals))
            p90_val = float(np.percentile(vals, 90))
        else:
            median_val = float(np.median(vals)) if vals else 50000.0
            p90_val = median_val * 2.0

        # Requirement 4: Dynamic spike threshold with absolute minimum floor of 10,000,000 TL (10 Million TL)
        spike_threshold = max(median_val * 7.0, p90_val * 3.5, 10_000_000.0)
        region_baselines[region] = {
            'median': median_val,
            'p90': p90_val,
            'spike_threshold': spike_threshold
        }

    # Step 3: Evaluate rows against region dynamic statistical baselines
    for idx, row in enumerate(dataset, start=2):
        # Anomali kontrol bloğunun en başına bunu ekle:
        bedel_degeri = float(row.get('Dağıtım Bedeli(TL)', 0) or 0)
        if bedel_degeri < 10000000.0:
            continue # 10 Milyon TL'nin altındaki hiçbir faturaya anomali UYARISI VERME!

        region = row.get('Dağıtım Bölgesi', 'Bilinmeyen Bölge')
        etso = row.get('Etso Kodu', '')
        masked_etso_str = mask_etso(etso)
        musteri = row.get('Müşteri', '')
        dagitim = float(row.get('Dağıtım Bedeli(TL)', 0) or 0)
        guc = float(row.get('Güç Bedeli(TL)', 0) or 0)
        toplam = float(row.get('Toplam (TL)', 0) or (dagitim + guc))
        val = abs(toplam if toplam > 0 else dagitim)

        baseline = region_baselines.get(region, {'median': 100000.0, 'spike_threshold': 10_000_000.0})
        reg_median = baseline['median']
        spike_thresh = baseline['spike_threshold']

        # 1. Absurd Data Corruption Check (>50 Billion TL)
        if val > ABSURD_CORRUPTION_THRESHOLD:
            absurd_records.append({'row': idx, 'region': region, 'etso': masked_etso_str, 'val': val})
            warnings.append({
                'type': 'DANGER',
                'category': 'Absürt Tutar',
                'message': f"⚠️ KRİTİK: {region} bölgesinde ETSO {masked_etso_str} için {val:,.2f} TL tutarında bozuk/absürt veri tespit edildi!"
            })

        # 2. Dynamic Statistical Spike Check (Kural: 10.000.000 TL altındaki faturalara KESİNLİKLE uyarı VERME)
        elif val >= 10_000_000.0 and val >= spike_thresh:
            ratio = val / max(1.0, reg_median)
            dynamic_spike_records.append({'row': idx, 'region': region, 'etso': masked_etso_str, 'val': val, 'ratio': ratio})
            warnings.append({
                'type': 'WARNING',
                'category': 'Yüksek Tutar',
                'message': f"🔔 DİKKAT: {region} bölgesinde ETSO {masked_etso_str} için {val:,.2f} TL tutarında yüksek fatura tespit edildi (10M TL üzeri & bölge ortalamasının {ratio:.1f} katı!)."
            })

        # 3. Negative Cost Check
        if dagitim < 0 and region != 'AKEDAŞ ( Göksu EDAŞ )':
            negative_cost_records.append({'row': idx, 'region': region, 'etso': etso, 'val': dagitim})

        # 4. Missing Field Check
        if not musteri or str(musteri).strip() in ('', 'None', 'BOŞ', 'boş'):
            missing_customer_count += 1

        if not etso or str(etso).strip() in ('', 'None'):
            missing_etso_count += 1

    if negative_cost_records:
        warnings.append({
            'type': 'INFO',
            'category': 'Negatif Tutar',
            'message': f"ℹ️ BİLGİ: {len(negative_cost_records)} adet negatif bedelli kayıttan filtreleme kuralları uygulandı."
        })

    if missing_customer_count > 0:
        warnings.append({
            'type': 'INFO',
            'category': 'Eksik Müşteri',
            'message': f"ℹ️ BİLGİ: Toplam {missing_customer_count} adet kayıtta müşteri unvanı bulunmuyor (Tarife/Bedel kontrolü yapıldı)."
        })

    health_score = max(0, 100 - (len(absurd_records) * 20) - (missing_etso_count * 5))

    return {
        'total_records': len(dataset),
        'region_counts': region_counts,
        'absurd_count': len(absurd_records),
        'high_value_count': len(dynamic_spike_records),
        'negative_cost_count': len(negative_cost_records),
        'missing_customer_count': missing_customer_count,
        'missing_etso_count': missing_etso_count,
        'health_score': health_score,
        'region_baselines': region_baselines,
        'warnings': warnings
    }
