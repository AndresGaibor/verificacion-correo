"""Tests for core.gal_scraper module."""

import json
from pathlib import Path
from unittest.mock import patch, MagicMock
from io import BytesIO

import pytest
from urllib.error import HTTPError, URLError

from verificacion_correo.core.gal_scraper import (
    ProgressFile,
    _flatten_persona,
    _csv_safe,
    save_to_csv,
    save_to_json,
    _build_find_people_payload,
    scrape_gal,
    PROGRESS_FILENAME,
    OUTPUT_JSON,
    OUTPUT_CSV,
    DEFAULT_BATCH_SIZE,
)

from verificacion_correo.core.api_extractor import SessionExpiredError


# ─── ProgressFile Tests ──────────────────────────────────────────────

class TestProgressFile:
    def test_load_returns_default_when_no_file(self, tmp_path):
        pf = ProgressFile(tmp_path)
        result = pf.load()
        assert result == {"offset": 0, "people": []}

    def test_exists_false_when_no_file(self, tmp_path):
        pf = ProgressFile(tmp_path)
        assert pf.exists is False

    def test_exists_true_when_file_exists(self, tmp_path):
        pf = ProgressFile(tmp_path)
        pf.save(10, [{"name": "A"}])
        assert pf.exists is True

    def test_save_creates_file(self, tmp_path):
        pf = ProgressFile(tmp_path)
        pf.save(5, [{"name": "A"}, {"name": "B"}])

        raw = json.loads((tmp_path / PROGRESS_FILENAME).read_text(encoding="utf-8"))
        assert raw["offset"] == 5
        assert raw["count"] == 2
        assert "last_update" in raw
        # save() does NOT store people
        assert "people" not in raw

    def test_save_overwrites_existing(self, tmp_path):
        pf = ProgressFile(tmp_path)
        pf.save(5, [{"name": "A"}])
        pf.save(10, [{"name": "A"}, {"name": "B"}])

        raw = json.loads((tmp_path / PROGRESS_FILENAME).read_text(encoding="utf-8"))
        assert raw["offset"] == 10
        assert raw["count"] == 2

    def test_load_returns_saved_state(self, tmp_path):
        pf = ProgressFile(tmp_path)
        pf.save(42, [{"name": "X"}])

        # load() reconstructs people from file on disk; save() writes only
        # count, but ProgressFile load() uses setdefault("people", []).
        result = pf.load()
        assert result["offset"] == 42
        assert "people" in result

    def test_load_handles_backward_compatible_missing_people(self, tmp_path):
        """File without 'people' key still loads correctly."""
        progress_file = tmp_path / PROGRESS_FILENAME
        progress_file.write_text(
            json.dumps({"offset": 20, "count": 3}), encoding="utf-8"
        )
        pf = ProgressFile(tmp_path)
        result = pf.load()
        assert result["offset"] == 20
        assert result["people"] == []
        assert result["count"] == 3

    def test_clear_removes_file(self, tmp_path):
        pf = ProgressFile(tmp_path)
        pf.save(1, [])
        assert pf.exists is True
        pf.clear()
        assert pf.exists is False

    def test_clear_no_error_when_no_file(self, tmp_path):
        pf = ProgressFile(tmp_path)
        pf.clear()  # should not raise
        assert pf.exists is False

    def test_save_empty_people_list(self, tmp_path):
        pf = ProgressFile(tmp_path)
        pf.save(0, [])
        raw = json.loads((tmp_path / PROGRESS_FILENAME).read_text(encoding="utf-8"))
        assert raw["offset"] == 0
        assert raw["count"] == 0


# ─── _flatten_persona Tests ──────────────────────────────────────────

class TestFlattenPersona:
    def test_flatten_full_persona(self):
        persona = {
            "DisplayName": "Juan Pérez",
            "EmailAddress": {"EmailAddress": "juan@madrid.org"},
            "BusinessPhoneNumbersArray": [
                {"Value": {"Number": "912345678"}}
            ],
            "ImAddress": "sip:juan@madrid.org",
            "CompanyName": "Ayuntamiento",
            "Department": "IT",
            "OfficeLocation": "Edificio Central",
            "BusinessAddressesArray": [
                {
                    "Value": {
                        "Street": "Calle Mayor 1",
                        "City": "Madrid",
                        "State": "Madrid",
                        "PostalCode": "28013",
                        "Country": "España",
                    }
                }
            ],
        }
        result = _flatten_persona(persona)
        assert result["nombre"] == "Juan Pérez"
        assert result["email"] == "juan@madrid.org"
        assert result["telefono"] == "912345678"
        assert result["sip"] == "sip:juan@madrid.org"
        assert result["empresa"] == "Ayuntamiento"
        assert result["departamento"] == "IT"
        assert result["oficina"] == "Edificio Central"
        assert "Calle Mayor 1" in result["direccion"]
        assert "Madrid" in result["direccion"]

    def test_flatten_empty_persona(self):
        result = _flatten_persona({})
        assert result["nombre"] == ""
        assert result["email"] == ""
        assert result["telefono"] == ""
        assert result["sip"] == ""
        assert result["empresa"] == ""
        assert result["departamento"] == ""
        assert result["oficina"] == ""
        assert result["direccion"] == ""

    def test_flatten_email_as_string(self):
        persona = {"DisplayName": "Test", "EmailAddress": "directo@madrid.org"}
        result = _flatten_persona(persona)
        assert result["email"] == "directo@madrid.org"

    def test_flatten_missing_optional_fields(self):
        persona = {"DisplayName": "Solo Nombre"}
        result = _flatten_persona(persona)
        assert result["nombre"] == "Solo Nombre"
        assert result["email"] == ""
        assert result["telefono"] == ""
        assert result["sip"] == ""

    def test_flatten_phone_value_as_string(self):
        persona = {
            "BusinessPhoneNumbersArray": [{"Value": "600123456"}]
        }
        result = _flatten_persona(persona)
        assert result["telefono"] == "600123456"

    def test_flatten_address_value_as_string(self):
        persona = {
            "BusinessAddressesArray": [{"Value": "Gran Vía 28, Madrid"}]
        }
        result = _flatten_persona(persona)
        assert result["direccion"] == "Gran Vía 28, Madrid"


# ─── _csv_safe Tests ─────────────────────────────────────────────────

class TestCsvSafe:
    def test_none_returns_empty(self):
        assert _csv_safe(None) == ""

    def test_string_passthrough(self):
        assert _csv_safe("hello") == "hello"

    def test_int_to_string(self):
        assert _csv_safe(42) == "42"

    def test_replaces_newlines(self):
        assert _csv_safe("line1\nline2") == "line1 line2"

    def test_replaces_carriage_return(self):
        assert _csv_safe("line1\rline2") == "line1 line2"

    def test_replaces_semicolon(self):
        assert _csv_safe("a;b") == "a,b"


# ─── save_to_csv Tests ───────────────────────────────────────────────

class TestSaveToCsv:
    def test_creates_csv_file(self, tmp_path):
        contacts = [
            {"nombre": "Ana", "email": "ana@madrid.org", "telefono": "911",
             "sip": "", "empresa": "Ayto", "departamento": "IT",
             "oficina": "", "direccion": ""},
        ]
        csv_path = tmp_path / "output.csv"
        save_to_csv(contacts, csv_path)
        assert csv_path.exists()
        content = csv_path.read_text(encoding="utf-8-sig")
        assert "Ana" in content
        assert "ana@madrid.org" in content

    def test_creates_empty_csv(self, tmp_path):
        csv_path = tmp_path / "empty.csv"
        save_to_csv([], csv_path)
        assert csv_path.exists()
        lines = csv_path.read_text(encoding="utf-8-sig").strip().splitlines()
        # header only
        assert len(lines) == 1

    def test_csv_uses_semicolon_delimiter(self, tmp_path):
        contacts = [
            {"nombre": "B", "email": "b@m.org", "telefono": "",
             "sip": "", "empresa": "", "departamento": "",
             "oficina": "", "direccion": ""},
        ]
        csv_path = tmp_path / "delim.csv"
        save_to_csv(contacts, csv_path)
        content = csv_path.read_text(encoding="utf-8-sig")
        header_line = content.splitlines()[0]
        assert ";" in header_line


# ─── save_to_json Tests ──────────────────────────────────────────────

class TestSaveToJson:
    def test_creates_json_file(self, tmp_path):
        data = [{"DisplayName": "Test", "email": "t@m.org"}]
        json_path = tmp_path / "out.json"
        save_to_json(data, json_path)
        assert json_path.exists()
        loaded = json.loads(json_path.read_text(encoding="utf-8"))
        assert len(loaded) == 1
        assert loaded[0]["DisplayName"] == "Test"

    def test_creates_empty_json_array(self, tmp_path):
        json_path = tmp_path / "empty.json"
        save_to_json([], json_path)
        assert json_path.exists()
        loaded = json.loads(json_path.read_text(encoding="utf-8"))
        assert loaded == []


# ─── _build_find_people_payload Tests ────────────────────────────────

class TestBuildFindPeoplePayload:
    def test_payload_structure(self):
        payload = _build_find_people_payload(0, 100, "test-id")
        assert payload["__type"] == "FindPeopleJsonRequest:#Exchange"
        body = payload["Body"]
        view = body["IndexedPageItemView"]
        assert view["Offset"] == 0
        assert view["MaxEntriesReturned"] == 100
        assert view["BasePoint"] == "Beginning"
        folder_id = body["ParentFolderId"]["BaseFolderId"]
        assert folder_id["Id"] == "test-id"

    def test_payload_with_offset(self):
        payload = _build_find_people_payload(200, 50, "addr-id")
        view = payload["Body"]["IndexedPageItemView"]
        assert view["Offset"] == 200
        assert view["MaxEntriesReturned"] == 50


# ─── scrape_gal Tests (mocked HTTP) ──────────────────────────────────

def _mock_session_file(tmp_path):
    """Create a valid session file for scrape_gal tests."""
    session = tmp_path / "state.json"
    session.write_text(
        json.dumps({
            "cookies": [
                {"name": "X-OWA-CANARY", "value": "canary123", "domain": ".madrid.org"},
            ]
        }),
        encoding="utf-8",
    )
    return str(session)


def _mock_response(data: dict):
    """Create a mock urllib response."""
    body = json.dumps(data).encode("utf-8")
    resp = MagicMock()
    resp.read.return_value = body
    return resp


class TestScrapeGalEmpty:
    def test_scrape_gal_empty_result(self, tmp_path):
        session = _mock_session_file(tmp_path)
        output = str(tmp_path / "output")

        with patch("verificacion_correo.core.gal_scraper._build_cookie_header", return_value="cookie=val"), \
             patch("verificacion_correo.core.gal_scraper._get_canary", return_value="canary123"), \
             patch("verificacion_correo.core.gal_scraper.urlopen") as mock_urlopen:
            mock_urlopen.return_value = _mock_response({
                "Body": {"People": []}
            })
            stats = scrape_gal(session, output_dir=output, batch_size=10, request_delay=0)

        assert stats["total"] == 0
        assert stats["expired"] is False
        assert stats["offset_end"] == 0


class TestScrapeGalWithContacts:
    def test_scrape_gal_with_contacts(self, tmp_path):
        session = _mock_session_file(tmp_path)
        output = str(tmp_path / "output")
        people = [
            {"DisplayName": "Ana", "EmailAddress": {"EmailAddress": "ana@m.org"}},
            {"DisplayName": "Luis", "EmailAddress": {"EmailAddress": "luis@m.org"}},
        ]

        with patch("verificacion_correo.core.gal_scraper._build_cookie_header", return_value="c=v"), \
             patch("verificacion_correo.core.gal_scraper._get_canary", return_value="can"), \
             patch("verificacion_correo.core.gal_scraper.urlopen") as mock_urlopen, \
             patch("verificacion_correo.core.gal_scraper.time.sleep"):
            # First call: 2 contacts; second call: empty -> stop
            mock_urlopen.side_effect = [
                _mock_response({"Body": {"People": people}}),
                _mock_response({"Body": {"People": []}}),
            ]
            stats = scrape_gal(session, output_dir=output, batch_size=10, request_delay=0)

        assert stats["total"] == 2
        assert stats["expired"] is False
        # Verify output files exist
        assert Path(stats["files"]["json"]).exists()
        assert Path(stats["files"]["csv"]).exists()


class TestScrapeGalSessionExpired:
    def test_scrape_gal_session_expired(self, tmp_path):
        session = _mock_session_file(tmp_path)
        output = str(tmp_path / "output")

        http_307 = HTTPError(
            url="https://correoweb.madrid.org/owa/service.svc?action=FindPeople",
            code=307,
            msg="Temporary Redirect",
            hdrs={},
            fp=BytesIO(b""),
        )

        with patch("verificacion_correo.core.gal_scraper._build_cookie_header", return_value="c=v"), \
             patch("verificacion_correo.core.gal_scraper._get_canary", return_value="can"), \
             patch("verificacion_correo.core.gal_scraper.urlopen", side_effect=http_307), \
             patch("verificacion_correo.core.gal_scraper.time.sleep"):
            stats = scrape_gal(session, output_dir=output, request_delay=0)

        assert stats["expired"] is True
        assert stats["total"] == 0


class TestScrapeGalNoCanary:
    def test_scrape_gal_no_canary_raises(self, tmp_path):
        session = _mock_session_file(tmp_path)
        output = str(tmp_path / "output")

        with patch("verificacion_correo.core.gal_scraper._build_cookie_header", return_value="c=v"), \
             patch("verificacion_correo.core.gal_scraper._get_canary", return_value=""):
            with pytest.raises(ValueError, match="X-OWA-CANARY"):
                scrape_gal(session, output_dir=output, request_delay=0)


class TestScrapeGalMaxContacts:
    def test_scrape_gal_respects_max_contacts(self, tmp_path):
        session = _mock_session_file(tmp_path)
        output = str(tmp_path / "output")
        people = [{"DisplayName": f"P{i}"} for i in range(10)]

        with patch("verificacion_correo.core.gal_scraper._build_cookie_header", return_value="c=v"), \
             patch("verificacion_correo.core.gal_scraper._get_canary", return_value="can"), \
             patch("verificacion_correo.core.gal_scraper.urlopen") as mock_urlopen, \
             patch("verificacion_correo.core.gal_scraper.time.sleep"):
            # First call: return only the requested batch (5), second: empty
            mock_urlopen.side_effect = [
                _mock_response({"Body": {"People": people[:5]}}),
                _mock_response({"Body": {"People": []}}),
            ]
            stats = scrape_gal(session, output_dir=output, max_contacts=5, batch_size=10, request_delay=0)

        assert stats["total"] == 5


class TestScrapeGalStopFlag:
    def test_scrape_gal_stop_flag(self, tmp_path):
        session = _mock_session_file(tmp_path)
        output = str(tmp_path / "output")
        people = [{"DisplayName": "X"}]
        stop = {"stop": True}

        with patch("verificacion_correo.core.gal_scraper._build_cookie_header", return_value="c=v"), \
             patch("verificacion_correo.core.gal_scraper._get_canary", return_value="can"), \
             patch("verificacion_correo.core.gal_scraper.urlopen") as mock_urlopen, \
             patch("verificacion_correo.core.gal_scraper.time.sleep"):
            mock_urlopen.return_value = _mock_response({"Body": {"People": people}})
            stats = scrape_gal(session, output_dir=output, stop_flag=stop, request_delay=0)

        assert stats["stopped"] is True


class TestScrapeGalForceRestart:
    def test_scrape_gal_force_restart(self, tmp_path):
        session = _mock_session_file(tmp_path)
        output = str(tmp_path / "output")

        # Create existing progress file
        out_path = Path(output)
        out_path.mkdir(parents=True, exist_ok=True)
        pf = ProgressFile(out_path)
        pf.save(50, [{"old": "data"}] * 50)

        with patch("verificacion_correo.core.gal_scraper._build_cookie_header", return_value="c=v"), \
             patch("verificacion_correo.core.gal_scraper._get_canary", return_value="can"), \
             patch("verificacion_correo.core.gal_scraper.urlopen") as mock_urlopen, \
             patch("verificacion_correo.core.gal_scraper.time.sleep"):
            mock_urlopen.return_value = _mock_response({"Body": {"People": []}})
            stats = scrape_gal(session, output_dir=output, force_restart=True, request_delay=0)

        # force_restart resets offset to 0
        assert stats["offset_end"] == 0
        assert stats["total"] == 0


# ─── scrape_gal Validation Tests ─────────────────────────────────────

class TestScrapeGalValidation:
    def test_raises_when_session_file_missing(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            scrape_gal(str(tmp_path / "nonexistent.json"),
                       output_dir=str(tmp_path / "out"),
                       request_delay=0)

    def test_raises_when_output_dir_empty(self, tmp_path):
        session = _mock_session_file(tmp_path)
        with pytest.raises(ValueError, match="output_dir"):
            scrape_gal(session, output_dir="", request_delay=0)

    def test_raises_when_max_contacts_negative(self, tmp_path):
        session = _mock_session_file(tmp_path)
        with pytest.raises(ValueError, match="max_contacts"):
            scrape_gal(session, output_dir=str(tmp_path / "out"),
                       max_contacts=-1, request_delay=0)

    def test_raises_when_batch_size_zero(self, tmp_path):
        session = _mock_session_file(tmp_path)
        with pytest.raises(ValueError, match="batch_size"):
            scrape_gal(session, output_dir=str(tmp_path / "out"),
                       batch_size=0, request_delay=0)

    def test_raises_when_request_delay_negative(self, tmp_path):
        session = _mock_session_file(tmp_path)
        with pytest.raises(ValueError, match="request_delay"):
            scrape_gal(session, output_dir=str(tmp_path / "out"),
                       request_delay=-1.0)
