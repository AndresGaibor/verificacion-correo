"""Tests para verificacion_correo.core.gal_exporter."""

import pytest
from pathlib import Path
from openpyxl import load_workbook
from verificacion_correo.core.gal_exporter import (
    flatten_contact_to_dict,
    extract_companies_from_contacts,
    save_to_excel,
    load_gal_cache,
)


class TestFlattenContact:
    """Tests para flatten_contact_to_dict."""

    def test_flatten_contact_basic(self):
        persona = {
            'DisplayName': 'Juan Perez',
            'EmailAddresses': [{'Value': 'juan@madrid.org'}],
            'CompanyName': 'ORGANOS JUDICIALES',
        }
        result = flatten_contact_to_dict(persona)
        assert result['nombre'] == 'Juan Perez'
        assert result['email'] == 'juan@madrid.org'
        assert result['empresa'] == 'ORGANOS JUDICIALES'

    def test_flatten_contact_no_email(self):
        persona = {
            'DisplayName': 'Test User',
            'CompanyName': 'TEST CO',
        }
        result = flatten_contact_to_dict(persona)
        assert result['email'] == ''
        assert result['empresa'] == 'TEST CO'

    def test_flatten_contact_empty_fields(self):
        persona = {}
        result = flatten_contact_to_dict(persona)
        assert result['nombre'] == ''
        assert result['email'] == ''
        assert result['empresa'] == ''


class TestExtractCompanies:
    """Tests para extract_companies_from_contacts."""

    def test_extract_companies(self):
        contacts = [
            {'empresa': 'ORG1'},
            {'empresa': 'ORG2'},
            {'empresa': 'ORG1'},
        ]
        result = extract_companies_from_contacts(contacts)
        assert result == ['ORG1', 'ORG2']

    def test_extract_companies_empty(self):
        contacts = [{'empresa': ''}, {'empresa': '  '}]
        result = extract_companies_from_contacts(contacts)
        assert result == []

    def test_extract_companies_sorted(self):
        contacts = [
            {'empresa': 'ZZZ'},
            {'empresa': 'AAA'},
            {'empresa': 'MMM'},
        ]
        result = extract_companies_from_contacts(contacts)
        assert result == ['AAA', 'MMM', 'ZZZ']


class TestSaveToExcel:
    """Tests para save_to_excel."""

    def test_save_to_excel_creates_two_sheets(self, tmp_path):
        contacts = [
            {'nombre': 'Test', 'email': 'test@test.com', 'empresa': 'TEST CO',
             'telefono': '', 'departamento': '', 'oficina': '', 'direccion': ''},
        ]
        excel_path = tmp_path / "test.xlsx"
        save_to_excel(contacts, excel_path)

        wb = load_workbook(excel_path)
        assert "Contactos" in wb.sheetnames
        assert "Compañías" in wb.sheetnames

    def test_save_to_excel_contactos_headers(self, tmp_path):
        contacts = [
            {'nombre': 'Test', 'email': 'test@test.com', 'empresa': 'TEST CO',
             'telefono': '', 'departamento': '', 'oficina': '', 'direccion': ''},
        ]
        excel_path = tmp_path / "test.xlsx"
        save_to_excel(contacts, excel_path)

        wb = load_workbook(excel_path)
        ws1 = wb["Contactos"]
        headers = [cell.value for cell in ws1[1]]
        assert headers == ['nombre', 'email', 'empresa', 'telefono', 'departamento', 'oficina', 'direccion']

    def test_save_to_excel_companias_sheet(self, tmp_path):
        contacts = [
            {'nombre': 'Test', 'email': 'test@test.com', 'empresa': 'TEST CO',
             'telefono': '', 'departamento': '', 'oficina': '', 'direccion': ''},
        ]
        excel_path = tmp_path / "test.xlsx"
        save_to_excel(contacts, excel_path)

        wb = load_workbook(excel_path)
        ws2 = wb["Compañías"]
        assert ws2.cell(1, 1).value == 'Compañía'
        assert ws2.cell(1, 2).value == 'Enrich'
        assert ws2.cell(2, 1).value == 'TEST CO'

    def test_save_to_excel_with_cache(self, tmp_path):
        contacts = [
            {'nombre': 'Test', 'email': 'test@test.com', 'empresa': 'TEST CO',
             'telefono': '', 'departamento': '', 'oficina': '', 'direccion': ''},
        ]
        excel_path = tmp_path / "test.xlsx"
        cache_path = tmp_path / "cache.json"
        save_to_excel(contacts, excel_path, cache_path)

        assert cache_path.exists()
        loaded = load_gal_cache(cache_path)
        assert len(loaded) == 1
        assert loaded[0]['nombre'] == 'Test'

    def test_save_to_excel_multiple_companies(self, tmp_path):
        contacts = [
            {'nombre': 'A', 'email': 'a@test.com', 'empresa': 'COMP A',
             'telefono': '', 'departamento': '', 'oficina': '', 'direccion': ''},
            {'nombre': 'B', 'email': 'b@test.com', 'empresa': 'COMP B',
             'telefono': '', 'departamento': '', 'oficina': '', 'direccion': ''},
            {'nombre': 'C', 'email': 'c@test.com', 'empresa': 'COMP A',
             'telefono': '', 'departamento': '', 'oficina': '', 'direccion': ''},
        ]
        excel_path = tmp_path / "test.xlsx"
        save_to_excel(contacts, excel_path)

        wb = load_workbook(excel_path)
        ws2 = wb["Compañías"]
        companies = [ws2.cell(r, 1).value for r in range(2, ws2.max_row + 1)]
        assert companies == ['COMP A', 'COMP B']
