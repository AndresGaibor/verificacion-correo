"""Tests para verificacion_correo.core.gal_enricher."""

import pytest
from pathlib import Path
from openpyxl import Workbook
from verificacion_correo.core.gal_enricher import (
    get_companies_to_enrich_from_excel,
)


class TestGetCompaniesFromExcel:
    """Tests para get_companies_to_enrich_from_excel."""

    def test_returns_marked_companies(self, tmp_path):
        wb = Workbook()
        ws2 = wb.create_sheet("Compañías")
        ws2.append(['Compañía', 'Enrich'])
        ws2.append(['ORG1', 'X'])
        ws2.append(['ORG2', ''])
        ws2.append(['ORG3', 'X'])

        excel_path = tmp_path / "test.xlsx"
        wb.save(excel_path)

        result = get_companies_to_enrich_from_excel(excel_path)
        assert result == ['ORG1', 'ORG3']

    def test_empty_when_no_x(self, tmp_path):
        wb = Workbook()
        ws2 = wb.create_sheet("Compañías")
        ws2.append(['Compañía', 'Enrich'])
        ws2.append(['ORG1', ''])

        excel_path = tmp_path / "test.xlsx"
        wb.save(excel_path)

        result = get_companies_to_enrich_from_excel(excel_path)
        assert result == []
