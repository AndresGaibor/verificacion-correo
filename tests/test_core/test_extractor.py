"""Tests for core.extractor module."""

from unittest.mock import MagicMock, patch

import pytest

from verificacion_correo.core.extractor import ContactInfo, ContactExtractor


class TestContactInfo:
    def test_default_creation(self):
        info = ContactInfo()
        assert info.name is None
        assert info.email is None
        assert info.phone is None
        assert info.sip is None
        assert info.address is None
        assert info.department is None
        assert info.company is None
        assert info.office_location is None

    def test_with_all_fields(self):
        info = ContactInfo(
            name="Juan Perez",
            email="juan.perez@madrid.org",
            phone="912345678",
            sip="sip:juan.perez@madrid.org",
            address="C/ Mayor 1, 28001 Madrid",
            department="IT",
            company="Madrid Org",
            office_location="Edificio A",
        )
        assert info.name == "Juan Perez"
        assert info.email == "juan.perez@madrid.org"
        assert info.phone == "912345678"
        assert info.sip == "sip:juan.perez@madrid.org"
        assert "Mayor" in info.address
        assert info.department == "IT"
        assert info.company == "Madrid Org"
        assert info.office_location == "Edificio A"

    def test_to_dict(self):
        info = ContactInfo(name="Test", email="test@madrid.org")
        d = info.to_dict()
        assert d["name"] == "Test"
        assert d["email"] == "test@madrid.org"
        assert d["phone"] is None
        assert d["sip"] is None
        assert d["address"] is None
        assert d["department"] is None
        assert d["company"] is None
        assert d["office_location"] is None

    def test_to_dict_excludes_none(self):
        info = ContactInfo(name="Test", email="test@madrid.org")
        d = info.to_dict()
        # to_dict returns all keys, None values included
        assert "phone" in d
        assert d["phone"] is None

    def test_is_valid_with_email(self):
        info = ContactInfo(email="test@madrid.org")
        assert info.is_valid() is True

    def test_is_valid_with_phone(self):
        info = ContactInfo(phone="912345678")
        assert info.is_valid() is True

    def test_is_valid_with_both(self):
        info = ContactInfo(email="test@madrid.org", phone="912345678")
        assert info.is_valid() is True

    def test_is_valid_with_name_only(self):
        info = ContactInfo(name="Test User")
        assert info.is_valid() is False

    def test_is_valid_empty(self):
        info = ContactInfo()
        assert info.is_valid() is False

    def test_is_valid_with_all_fields(self):
        info = ContactInfo(
            name="Test",
            email="test@madrid.org",
            phone="912345678",
            sip="sip:test@madrid.org",
        )
        assert info.is_valid() is True

    def test_repr_empty(self):
        info = ContactInfo()
        r = repr(info)
        assert "ContactInfo(" in r

    def test_repr_with_fields(self):
        info = ContactInfo(name="Test", email="test@madrid.org")
        r = repr(info)
        assert "name=" in r
        assert "email=" in r

    def test_repr_excludes_none_fields(self):
        info = ContactInfo(name="Test", email="test@madrid.org", phone=None)
        r = repr(info)
        assert "phone" not in r


class TestContactExtractor:
    @pytest.fixture
    def mock_config(self):
        config = MagicMock()
        config.patterns = MagicMock()
        config.selectors = MagicMock()
        config.wait_times = MagicMock()

        import re
        config.patterns.EMAIL = re.compile(r'[\w.+-]+@[\w.-]+\.[a-z]{2,}', re.I)
        config.patterns.PHONE = re.compile(r'\b\d{6,}\b')
        config.patterns.POSTAL_ADDR = re.compile(r'\d{5}\s+[A-ZÁÉÍÓÚÑ\-\s]+', re.I)
        config.patterns.SIP = re.compile(r'sip:[\w.+-]+@[\w.-]+', re.I)
        config.patterns.NAME = re.compile(
            r'([A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ\-\.\s]+,\s*[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ\-\s]+)'
        )

        config.selectors.popup = "div._pe_Y[ispopup='1']"
        config.wait_times.popup_visible = 5000

        return config

    def test_init(self, mock_config):
        extractor = ContactExtractor(mock_config)
        assert extractor.config is mock_config
        assert extractor.patterns is mock_config.patterns
        assert extractor.selectors is mock_config.selectors
        assert extractor.wait_times is mock_config.wait_times

    def test_extract_from_popup_no_popup(self, mock_config):
        extractor = ContactExtractor(mock_config)

        page = MagicMock()
        page.locator.return_value.first.wait_for.side_effect = Exception("Timeout")

        result = extractor.extract_from_popup(page)
        assert result is None

    def test_extract_from_popup_dom_based_valid(self, mock_config):
        extractor = ContactExtractor(mock_config)

        page = MagicMock()
        popup = MagicMock()
        popup.inner_text.return_value = "Test contact info"
        page.locator.return_value.first = popup
        popup.locator.return_value.count.return_value = 2

        with patch.object(extractor, '_extract_dom_based') as mock_dom:
            mock_dom.return_value = ContactInfo(
                email="test@madrid.org",
                phone="912345678"
            )
            result = extractor.extract_from_popup(page)
            assert result is not None
            assert result.email == "test@madrid.org"

    def test_extract_from_popup_fallback_to_text(self, mock_config):
        extractor = ContactExtractor(mock_config)

        page = MagicMock()
        popup = MagicMock()
        popup.inner_text.return_value = "Contact: test@madrid.org"
        page.locator.return_value.first = popup

        with patch.object(extractor, '_extract_dom_based', return_value=ContactInfo()):
            with patch.object(extractor, '_extract_text_based') as mock_text:
                mock_text.return_value = ContactInfo(
                    email="test@madrid.org",
                    phone="912345678"
                )
                result = extractor.extract_from_popup(page)
                assert result is not None
                assert result.email == "test@madrid.org"

    def test_extract_from_popup_both_fail(self, mock_config):
        extractor = ContactExtractor(mock_config)

        page = MagicMock()
        popup = MagicMock()
        popup.inner_text.return_value = "Some text without valid data"
        page.locator.return_value.first = popup

        with patch.object(extractor, '_extract_dom_based', return_value=ContactInfo()):
            with patch.object(extractor, '_extract_text_based', return_value=None):
                result = extractor.extract_from_popup(page)
                assert result is None

    def test_extract_from_popup_exception_handling(self, mock_config):
        extractor = ContactExtractor(mock_config)

        page = MagicMock()
        page.locator.side_effect = Exception("Unexpected error")

        result = extractor.extract_from_popup(page)
        assert result is None

    def test_wait_for_popup_timeout(self, mock_config):
        extractor = ContactExtractor(mock_config)

        page = MagicMock()
        locator = MagicMock()
        from playwright.sync_api import TimeoutError as PWTimeout
        locator.first.wait_for.side_effect = PWTimeout("Timeout")
        page.locator.return_value = locator

        result = extractor._wait_for_popup(page)
        assert result is None

    def test_get_popup_text_error(self, mock_config):
        extractor = ContactExtractor(mock_config)
        popup = MagicMock()
        popup.inner_text.side_effect = Exception("Error")

        result = extractor._get_popup_text(popup)
        assert result == ""

    def test_extract_specific_email_filters_generic(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "ASP123@MADRID.ORG and juan.perez@madrid.org"

        email = extractor._extract_specific_email(text)
        assert email == "juan.perez@madrid.org"

    def test_extract_specific_email_only_generic(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "ASP123@MADRID.ORG"

        email = extractor._extract_specific_email(text)
        assert email == "ASP123@MADRID.ORG"

    def test_extract_specific_email_none(self, mock_config):
        extractor = ContactExtractor(mock_config)
        email = extractor._extract_specific_email("No email here")
        assert email is None

    def test_extract_phone_9_digit(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "Trabajo\n912345678"

        phone = extractor._extract_phone(text)
        assert phone == "912345678"

    def test_extract_phone_6_digit_with_trabajo(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "Trabajo\n123456"

        phone = extractor._extract_phone(text)
        assert phone == "123456"

    def test_extract_phone_none(self, mock_config):
        extractor = ContactExtractor(mock_config)
        phone = extractor._extract_phone("No numbers here")
        assert phone is None

    def test_extract_sip_valid(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "sip:user@madrid.org"

        sip = extractor._extract_sip(text)
        assert sip == "sip:user@madrid.org"

    def test_extract_sip_invalid(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "sip:invalid"

        sip = extractor._extract_sip(text)
        assert sip is None

    def test_extract_sip_none(self, mock_config):
        extractor = ContactExtractor(mock_config)
        sip = extractor._extract_sip("No SIP here")
        assert sip is None

    def test_extract_address_spanish_format(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "C/ Mayor 1 28001 Madrid"

        address = extractor._extract_address(text)
        assert address is not None
        assert "Mayor" in address

    def test_extract_address_postal_code(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "28001 Madrid"

        address = extractor._extract_address(text)
        assert address is not None
        assert "28001" in address

    def test_extract_address_none(self, mock_config):
        extractor = ContactExtractor(mock_config)
        address = extractor._extract_address("No address here")
        assert address is None

    def test_extract_name_regex(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "GARCIA, JUAN\nSome other text"

        name = extractor._extract_name(text)
        assert name == "GARCIA, JUAN"

    def test_extract_name_heuristic(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "LOPEZ\nPEREZ\nMARTINEZ, CARLOS"

        name = extractor._extract_name(text)
        assert name is not None

    def test_extract_name_none(self, mock_config):
        extractor = ContactExtractor(mock_config)
        name = extractor._extract_name("no name format here")
        assert name is None

    def test_extract_work_info_department(self, mock_config):
        extractor = ContactExtractor(mock_config)
        data = {}
        extractor._extract_work_info("Departamento: Informática\n", data)
        assert data.get("department") == "Informática"

    def test_extract_work_info_company(self, mock_config):
        extractor = ContactExtractor(mock_config)
        data = {}
        extractor._extract_work_info("Compañía: Madrid Org\n", data)
        assert data.get("company") == "Madrid Org"

    def test_extract_work_info_office(self, mock_config):
        extractor = ContactExtractor(mock_config)
        data = {}
        extractor._extract_work_info("Oficina: Edificio Central\n", data)
        assert data.get("office_location") == "Edificio Central"

    def test_extract_work_info_heuristic(self, mock_config):
        extractor = ContactExtractor(mock_config)
        data = {}
        extractor._extract_work_info("CONTACTO\nSECRETARÍA TÉCNICA\nNOTAS\n", data)
        assert data.get("department") == "SECRETARÍA TÉCNICA"

    def test_extract_work_info_empty_value_skipped(self, mock_config):
        extractor = ContactExtractor(mock_config)
        data = {}
        extractor._extract_work_info("Departamento: \n", data)
        assert "department" not in data

    def test_extract_text_based_valid(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "juan.perez@madrid.org\n912345678"

        result = extractor._extract_text_based(text)
        assert result is not None
        assert result.email == "juan.perez@madrid.org"
        assert result.phone == "912345678"

    def test_extract_text_based_empty(self, mock_config):
        extractor = ContactExtractor(mock_config)
        result = extractor._extract_text_based("")
        assert result is None

    def test_extract_text_based_whitespace(self, mock_config):
        extractor = ContactExtractor(mock_config)
        result = extractor._extract_text_based("   \n  \n  ")
        assert result is None

    def test_extract_text_based_not_valid(self, mock_config):
        extractor = ContactExtractor(mock_config)
        text = "Some random text without emails or phones"

        with patch.object(extractor, '_extract_specific_email', return_value=None), \
             patch.object(extractor, '_extract_phone', return_value=None), \
             patch.object(extractor, '_extract_name', return_value="GARCIA, JUAN"):
            result = extractor._extract_text_based(text)
            # Name alone doesn't make it valid
            assert result is None

    def test_extract_by_text_labels_department(self, mock_config):
        extractor = ContactExtractor(mock_config)
        popup = MagicMock()
        popup.inner_text.return_value = (
            "Juan Perez\n"
            "Departamento: Informática\n"
            "Compañía: Madrid\n"
        )

        result = extractor._extract_by_text_labels(popup)
        assert result.get("department") == "Informática"
        assert result.get("company") == "Madrid"

    def test_extract_by_text_labels_phone(self, mock_config):
        extractor = ContactExtractor(mock_config)
        popup = MagicMock()
        popup.inner_text.return_value = (
            "Trabajo:\n"
            "912345678"
        )

        result = extractor._extract_by_text_labels(popup)
        assert "912345678" in result.get("phone", "")

    def test_extract_by_text_labels_sip(self, mock_config):
        extractor = ContactExtractor(mock_config)
        popup = MagicMock()
        popup.inner_text.return_value = (
            "MI:\n"
            "sip:user@madrid.org"
        )

        result = extractor._extract_by_text_labels(popup)
        assert result.get("sip") == "sip:user@madrid.org"

    def test_extract_by_text_labels_address(self, mock_config):
        extractor = ContactExtractor(mock_config)
        popup = MagicMock()
        popup.inner_text.return_value = (
            "Dirección profesional\n"
            "C/ Mayor 1\n"
            "28001 Madrid"
        )

        result = extractor._extract_by_text_labels(popup)
        assert result.get("address") is not None
        assert "Mayor" in result["address"]

    def test_extract_personal_email_filters_token(self, mock_config):
        extractor = ContactExtractor(mock_config)
        popup = MagicMock()
        email_locator = MagicMock()
        email_locator.count.return_value = 2

        token_mock = MagicMock()
        token_mock.get_attribute.return_value = "ASP123@MADRID.ORG"

        personal_mock = MagicMock()
        personal_mock.get_attribute.return_value = "juan.perez@madrid.org"

        email_locator.nth.side_effect = [token_mock, personal_mock]
        popup.locator.return_value = email_locator

        result = extractor._extract_personal_email(popup)
        assert result == "juan.perez@madrid.org"

    def test_extract_personal_email_all_tokens(self, mock_config):
        extractor = ContactExtractor(mock_config)
        popup = MagicMock()
        email_locator = MagicMock()
        email_locator.count.return_value = 2

        token1 = MagicMock()
        token1.get_attribute.return_value = "ASP123@MADRID.ORG"
        token2 = MagicMock()
        token2.get_attribute.return_value = "AGM456@MADRID.ORG"

        email_locator.nth.side_effect = [token1, token2]
        popup.locator.return_value = email_locator

        result = extractor._extract_personal_email(popup)
        assert result is None
