# -*- coding: utf-8 -*-
"""
Birim testler: Bölge ismi normalizasyonu ve ETSO kodu standartlaştırma.
"""
import pytest
from core.normalizer import normalize_region_name, normalize_etso_kodu


class TestNormalizeRegionName:
    """normalize_region_name fonksiyonu testleri."""

    # --- extraction formatına çevirme ---
    def test_reference_to_extraction_basic(self):
        assert normalize_region_name('Aydem EDAŞ', to_format='extraction') == 'ADM EDAŞ'

    def test_reference_to_extraction_kayseri(self):
        assert normalize_region_name('Kayseri EDAŞ', to_format='extraction') == 'KCETAŞ'

    def test_reference_to_extraction_akedas(self):
        assert normalize_region_name('AKEDAŞ (Göksu EDAŞ)', to_format='extraction') == 'AKEDAŞ'

    # --- reference formatına çevirme ---
    def test_extraction_to_reference_sedas(self):
        assert normalize_region_name('SEDAŞ', to_format='reference') == 'Sakarya EDAŞ'

    def test_extraction_to_reference_tredas(self):
        assert normalize_region_name('TREDAŞ', to_format='reference') == 'Trakya EDAŞ'

    def test_extraction_to_reference_yedas(self):
        assert normalize_region_name('YEDAŞ', to_format='reference') == 'Yeşilırmak EDAŞ'

    # --- Bilinmeyen bölge → olduğu gibi döner ---
    def test_unknown_region(self):
        assert normalize_region_name('Bilinmeyen EDAŞ', to_format='extraction') == 'Bilinmeyen EDAŞ'

    # --- Edge cases ---
    def test_none(self):
        assert normalize_region_name(None) is None

    def test_case_insensitive(self):
        """casefold() ile Türkçe İ/ı doğru handle edilir."""
        result = normalize_region_name('sedaş', to_format='reference')
        assert result == 'Sakarya EDAŞ'


class TestNormalizeEtsoKodu:
    """normalize_etso_kodu fonksiyonu testleri."""

    def test_eic_format_40z(self):
        """40Z EIC formatından kök numaraya çevirme."""
        assert normalize_etso_kodu('40Z000000123456T') == '123456'

    def test_plain_number(self):
        assert normalize_etso_kodu('123456') == '123456'

    def test_integer(self):
        assert normalize_etso_kodu(123456) == '123456'

    def test_float_format(self):
        """Excel'den gelen float formatı."""
        assert normalize_etso_kodu('123456.0') == '123456'

    def test_long_number(self):
        assert normalize_etso_kodu('12345678') == '12345678'

    def test_none(self):
        assert normalize_etso_kodu(None) is None

    def test_empty(self):
        assert normalize_etso_kodu('') is None

    def test_dash(self):
        assert normalize_etso_kodu('-') is None

    def test_leading_zeros(self):
        """Baştaki sıfırlar kaldırılır."""
        result = normalize_etso_kodu('000123456')
        assert result == '123456'
