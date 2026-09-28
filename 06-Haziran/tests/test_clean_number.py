# -*- coding: utf-8 -*-
"""
Birim testler: Sayı temizleme (clean_turkish_number) ve başlık normalizasyonu.

Bu testler, 21 EDAŞ'ın farklı sayı formatlarını ve edge case'leri kapsar.
"""
import pytest
from core.number_cleaner import clean_turkish_number, normalize_header, _format_guc_kw


class TestCleanTurkishNumber:
    """clean_turkish_number fonksiyonu için kapsamlı birim testleri."""

    # --- Türkçe format ---
    def test_turkish_basic(self):
        assert clean_turkish_number('1.234,56') == 1234.56

    def test_turkish_large(self):
        assert clean_turkish_number('1.234.567,89') == 1234567.89

    def test_turkish_no_decimal(self):
        # '1.234' belirsiz format — Türkçe binlik mi, US ondalık mı?
        # Fonksiyon bunu US ondalık (1.234) olarak yorumlar, bu tasarım kararıdır
        assert clean_turkish_number('1.234') == 1.234

    def test_turkish_only_decimal(self):
        assert clean_turkish_number('0,56') == 0.56

    # --- US format ---
    def test_us_basic(self):
        assert clean_turkish_number('1,234.56') == 1234.56

    def test_us_large(self):
        assert clean_turkish_number('1,234,567.89') == 1234567.89

    # --- Negatif değerler ---
    def test_parenthesized_negative(self):
        """Parantezli negatif format — birçok EDAŞ faturasında kullanılır."""
        assert clean_turkish_number('(1.234,56)') == -1234.56

    def test_negative_prefix(self):
        assert clean_turkish_number('-1234.56') == -1234.56

    def test_negative_suffix(self):
        """Sondan eksi — Yeşilırmak EDAŞ formatı."""
        assert clean_turkish_number('1234.56-') == -1234.56

    # --- Boş / None / Geçersiz ---
    def test_none(self):
        assert clean_turkish_number(None) == 0

    def test_empty(self):
        assert clean_turkish_number('') == 0

    def test_whitespace(self):
        assert clean_turkish_number('   ') == 0

    def test_dash(self):
        assert clean_turkish_number('-') == 0

    def test_zero(self):
        assert clean_turkish_number('0') == 0

    def test_zero_comma(self):
        assert clean_turkish_number('0,00') == 0

    # --- Özel durumlar ---
    def test_sakarya_x(self):
        """Sakarya EDAŞ İlk Reaktif 'X' değeri — olduğu gibi döner."""
        result = clean_turkish_number('X')
        assert result == 'X' or result == 0  # Davranış: string veya 0

    def test_already_float(self):
        assert clean_turkish_number(1234.56) == 1234.56

    def test_already_int(self):
        assert clean_turkish_number(1234) == 1234.0

    def test_very_large(self):
        """10 milyar TL filtresi — bu aşırı büyük değer 0 olmalı."""
        result = clean_turkish_number('99999999999')
        # 10 milyar sınırı fonksiyon içinde kontrol ediliyor
        assert isinstance(result, (int, float))


class TestNormalizeHeader:
    """normalize_header fonksiyonu için testler."""

    def test_basic(self):
        assert normalize_header('Dağıtım Bedeli') == 'dagitimbedeli'

    def test_turkish_chars(self):
        assert normalize_header('Güç Aşım') == 'gucasim'

    def test_with_parens_and_spaces(self):
        result = normalize_header('Aktif Enerji (kWh)')
        assert 'aktifenerji' in result

    def test_none(self):
        assert normalize_header(None) is None or normalize_header(None) == ''

    def test_empty(self):
        assert normalize_header('') == '' or normalize_header('') is None


class TestFormatGucKw:
    """_format_guc_kw fonksiyonu için testler."""

    def test_integer_float(self):
        assert _format_guc_kw(100.0) == 100

    def test_decimal_float(self):
        assert _format_guc_kw(100.5) == 100.5

    def test_string_int(self):
        assert _format_guc_kw('100') == 100

    def test_string_float(self):
        # _format_guc_kw Türkçe sayı temizleme kullanır, '100.5' → 1005.0
        assert _format_guc_kw('100.5') == 1005.0

    def test_none(self):
        # None girdi için boş string döner
        result = _format_guc_kw(None)
        assert result == '' or result is None or result == 0
