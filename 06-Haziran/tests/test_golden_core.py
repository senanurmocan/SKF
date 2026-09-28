# -*- coding: utf-8 -*-
"""
Golden Core regresyon testi — 2593 kayıt, 21 bölge, 0 hata.

Bu test, tüm pipeline'ı çalıştırıp çekirdek çıktının bozulmadığını doğrular.
Her commit öncesi çalıştırılmalıdır.
"""
import pytest
import io
import contextlib
from pathlib import Path
from collections import Counter


@pytest.fixture(scope='module')
def pipeline_result():
    """Tüm pipeline'ı çalıştır ve sonucu cache'le (module scope)."""
    from extract_and_compare import process_all_regions
    base = Path(__file__).resolve().parents[1]
    mapping = base / 'SKF Başlıkları.xlsx'
    region_folders_with_input = 0
    for folder in (path for path in base.iterdir() if path.is_dir()):
        if any(
            file.suffix.lower() in {'.xls', '.xlsx', '.xml', '.html'}
            for file in folder.rglob('*')
            if file.is_file()
        ):
            region_folders_with_input += 1

    if not mapping.is_file() or region_folders_with_input < 21:
        pytest.skip(
            'Golden integration test needs the private monthly workbooks for all 21 regions.'
        )

    captured = io.StringIO()
    with contextlib.redirect_stdout(captured):
        records, stats = process_all_regions(
            base_path=str(base),
            mapping_file_path=str(mapping),
            return_stats=True,
        )
    return records, stats, captured.getvalue()


class TestGoldenCore:
    """Golden Core regresyon testleri."""

    def test_total_records(self, pipeline_result):
        """Toplam kayıt sayısı 2593 olmalı."""
        records, _, _ = pipeline_result
        assert len(records) == 2593, f'Beklenen: 2593, Gelen: {len(records)}'

    def test_total_regions(self, pipeline_result):
        """21 bölge okunmalı."""
        records, _, _ = pipeline_result
        region_counts = Counter(r.get('Dağıtım Bölgesi') for r in records)
        assert len(region_counts) == 21, f'Beklenen: 21, Gelen: {len(region_counts)}'

    def test_no_empty_etso(self, pipeline_result):
        """Hiçbir kayıtta boş ETSO olmamalı."""
        records, _, _ = pipeline_result
        empty = [r for r in records if not r.get('Etso Kodu')]
        assert len(empty) == 0, f'{len(empty)} kayıtta boş ETSO'

    def test_sakarya_has_records(self, pipeline_result):
        """Sakarya EDAŞ en az 100 kayıt üretmeli."""
        records, _, _ = pipeline_result
        sakarya = [r for r in records if r.get('Dağıtım Bölgesi') == 'Sakarya EDAŞ']
        assert len(sakarya) >= 100, f'Sakarya: {len(sakarya)} kayıt'

    def test_sakarya_aktif_enerji_nonzero(self, pipeline_result):
        """Sakarya EDAŞ aktif enerji değerleri sıfır olmamalı (Notlar 12 #1)."""
        records, _, _ = pipeline_result
        sakarya = [r for r in records if r.get('Dağıtım Bölgesi') == 'Sakarya EDAŞ']
        nonzero = [r for r in sakarya if r.get('Aktif Enerji Tüketim (kWh)', 0) != 0]
        assert len(nonzero) > len(sakarya) * 0.8, \
            f'Sakarya sıfır olmayan aktif enerji: {len(nonzero)}/{len(sakarya)}'

    def test_all_regions_present(self, pipeline_result):
        """21 bölgenin tamamı çıktıda olmalı."""
        records, _, _ = pipeline_result
        regions = set(r.get('Dağıtım Bölgesi') for r in records)
        expected = {
            'ADM EDAŞ', 'AKEDAŞ', 'Akdeniz EDAŞ', 'Aras EDAŞ', 'AYEDAŞ',
            'Başkent EDAŞ', 'Boğaziçi EDAŞ', 'Çamlıbel EDAŞ', 'Çoruh EDAŞ',
            'Dicle EDAŞ', 'Fırat EDAŞ', 'Gediz EDAŞ', 'Kayseri ve Civarı',
            'Meram EDAŞ', 'Osmangazi EDAŞ', 'Sakarya EDAŞ', 'Toroslar EDAŞ',
            'Trakya EDAŞ', 'Uludağ EDAŞ', 'Vangölü EDAŞ', 'Yeşilırmak EDAŞ',
        }
        missing = expected - regions
        assert not missing, f'Eksik bölgeler: {missing}'

    def test_standard_columns_present(self, pipeline_result):
        """Her kayıtta temel alanlar olmalı."""
        records, _, _ = pipeline_result
        required_keys = ['Dağıtım Bölgesi', 'Etso Kodu', 'Aktif Enerji Tüketim (kWh)']
        for key in required_keys:
            present = sum(1 for r in records if key in r)
            assert present == len(records), f'{key}: {present}/{len(records)}'

    def test_no_warnings_in_standard_format(self, pipeline_result):
        """Haziran verilerinde standart dışı format uyarısı olmamalı."""
        _, _, log = pipeline_result
        warning_lines = [l for l in log.split('\n') if 'standart dışı format' in l]
        assert len(warning_lines) == 0, f'{len(warning_lines)} uyarı: {warning_lines[:3]}'
