# -*- coding: utf-8 -*-
"""
Birim testler: Reaktif hesaplama (calculate_reactive_totals).

DRY prensibine göre tek kaynaktan çağrılan reaktif hesaplamanın
tüm edge case'lerini doğrular.
"""
import pytest
from core.reactive_calculator import calculate_reactive_totals


class TestCalculateReactiveTotals:
    """calculate_reactive_totals fonksiyonu testleri."""

    def test_basic_reactive(self):
        """Basit reaktif toplam hesaplama."""
        cell_values = [None, None, '100,50', '200,75', '50,25', '30,00']
        indices = {'reaktif': 2, 'reaktif2': 3, 'reaktif_tenzil': 4, 'reaktif_tenzil2': 5}
        reaktif_fields = ['reaktif', 'reaktif2']
        ilk_reaktif_fields = ['reaktif_tenzil', 'reaktif_tenzil2']

        reaktif, ilk_reaktif, sakarya_x = calculate_reactive_totals(
            cell_values, indices, reaktif_fields, ilk_reaktif_fields
        )
        assert reaktif == pytest.approx(301.25, abs=0.01)
        assert ilk_reaktif == pytest.approx(80.25, abs=0.01)
        assert sakarya_x is False

    def test_sakarya_x_value(self):
        """Sakarya EDAŞ 'X' değeri — İlk Reaktif 'X' olmalı."""
        cell_values = [None, None, '100,50', None, 'X', None]
        indices = {'reaktif': 2, 'reaktif_tenzil': 4}
        reaktif_fields = ['reaktif']
        ilk_reaktif_fields = ['reaktif_tenzil']

        reaktif, ilk_reaktif, sakarya_x = calculate_reactive_totals(
            cell_values, indices, reaktif_fields, ilk_reaktif_fields
        )
        assert reaktif == pytest.approx(100.50, abs=0.01)
        assert ilk_reaktif == 'X'
        assert sakarya_x is True

    def test_empty_fields(self):
        """Hiç reaktif alan yoksa sıfır dönmeli."""
        cell_values = [None, None, None]
        indices = {}
        reaktif, ilk_reaktif, sakarya_x = calculate_reactive_totals(
            cell_values, indices, ['reaktif'], ['reaktif_tenzil']
        )
        assert reaktif == 0
        assert ilk_reaktif == 0
        assert sakarya_x is False

    def test_none_values(self):
        """None hücre değerleri atlanmalı."""
        cell_values = [None, None, None, None]
        indices = {'reaktif': 2, 'reaktif_tenzil': 3}
        reaktif, ilk_reaktif, sakarya_x = calculate_reactive_totals(
            cell_values, indices, ['reaktif'], ['reaktif_tenzil']
        )
        assert reaktif == 0
        assert ilk_reaktif == 0

    def test_negative_values(self):
        """Negatif değerler (parantezli format)."""
        cell_values = [None, None, '(100,50)', '(50,25)']
        indices = {'reaktif': 2, 'reaktif_tenzil': 3}
        reaktif, ilk_reaktif, sakarya_x = calculate_reactive_totals(
            cell_values, indices, ['reaktif'], ['reaktif_tenzil']
        )
        assert reaktif == pytest.approx(-100.50, abs=0.01)
        assert ilk_reaktif == pytest.approx(-50.25, abs=0.01)

    def test_out_of_bounds_index(self):
        """Sütun indeksi hücre listesi dışındaysa atlanmalı."""
        cell_values = ['val1', 'val2']
        indices = {'reaktif': 5}  # index > len(cell_values)
        reaktif, ilk_reaktif, sakarya_x = calculate_reactive_totals(
            cell_values, indices, ['reaktif'], []
        )
        assert reaktif == 0
