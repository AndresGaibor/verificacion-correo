"""Tests for gui.service module."""

from unittest.mock import MagicMock, patch, PropertyMock

import pytest
import queue

from verificacion_correo.gui.service import GUIService


@pytest.fixture
def mock_config():
    config = MagicMock()
    config.processing.batch_size = 10
    config.get_session_file_path.return_value = "/tmp/state.json"
    return config


@pytest.fixture
def service(mock_config):
    return GUIService(mock_config)


class TestGUIServiceInit:
    def test_creates_queue(self, service):
        assert isinstance(service.progress_queue, queue.Queue)

    def test_init_state(self, service):
        assert service.is_processing is False
        assert service.should_stop is False
        assert service.current_thread is None


class TestStartProcessingRaises:
    def test_raises_when_already_processing(self, service):
        service.is_processing = True
        with pytest.raises(RuntimeError, match="Processing already active"):
            service.start_processing("/fake/path.xlsx")

    def test_start_api_processing_raises_when_already_processing(self, service):
        service.is_processing = True
        with pytest.raises(RuntimeError, match="Processing already active"):
            service.start_api_processing("/fake/path.xlsx")

    def test_start_gal_scraping_raises_when_already_processing(self, service):
        service.is_processing = True
        with pytest.raises(RuntimeError, match="Processing already active"):
            service.start_gal_scraping("/fake/output")


class TestStopProcessing:
    def test_sets_flags(self, service):
        service.stop_processing()
        assert service.should_stop is True
        assert service.is_processing is False


class TestStopGalScraping:
    def test_sets_flag(self, service):
        service._gal_stop_flag = {'stop': False}
        service.stop_gal_scraping()
        assert service._gal_stop_flag['stop'] is True

    def test_no_flag_no_error(self, service):
        if hasattr(service, '_gal_stop_flag'):
            del service._gal_stop_flag
        service.stop_gal_scraping()


class TestCheckQueue:
    def test_yields_items(self, service):
        service.progress_queue.put(('progress', {'current': 1, 'total': 5}))
        service.progress_queue.put(('complete', {'ok': True}))
        items = list(service.check_queue())
        assert len(items) == 2
        assert items[0] == ('progress', {'current': 1, 'total': 5})
        assert items[1] == ('complete', {'ok': True})

    def test_empty_queue(self, service):
        items = list(service.check_queue())
        assert items == []


class TestGetExcelSummary:
    @patch('verificacion_correo.gui.service.ExcelReader')
    def test_success(self, MockExcelReader, service):
        mock_reader = MockExcelReader.return_value
        mock_summary = MagicMock()
        mock_summary.total_emails = 20
        mock_summary.pending_count = 15
        mock_summary.processed_count = 5
        mock_summary.batches = [1, 2]
        mock_reader.read_pending_emails.return_value = mock_summary

        result = service.get_excel_summary("/fake/data.xlsx")
        assert result['total_emails'] == 20
        assert result['pending_count'] == 15
        assert result['processed_count'] == 5
        assert result['batch_count'] == 2
        MockExcelReader.assert_called_once_with("/fake/data.xlsx")
        mock_reader.read_pending_emails.assert_called_once_with(batch_size=10)

    @patch('verificacion_correo.gui.service.ExcelReader')
    def test_error(self, MockExcelReader, service):
        MockExcelReader.side_effect = FileNotFoundError("File not found")
        result = service.get_excel_summary("/missing/file.xlsx")
        assert 'error' in result
        assert "File not found" in result['error']


class TestValidateSession:
    @patch('verificacion_correo.gui.service.get_session_status')
    def test_calls_config(self, mock_get_status, service):
        mock_get_status.return_value = {'valid': True}
        result = service.validate_session()
        mock_get_status.assert_called_once_with(service.config)
        assert result == {'valid': True}
