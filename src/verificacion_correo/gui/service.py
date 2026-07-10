"""
Background service for GUI operations.

This module provides the GUIService class that handles background processing
tasks for the GUI application.
"""

import threading
import queue
from typing import Optional, Dict, Any, Callable

from verificacion_correo.core.config import Config
from verificacion_correo.core.browser import BrowserAutomation
from verificacion_correo.core.session import SessionManager, get_session_status
from verificacion_correo.core.excel import ExcelReader, ExcelWriter
from verificacion_correo.core.api_extractor import process_emails_via_api
from verificacion_correo.core.gal_scraper import scrape_gal
from verificacion_correo.utils.logging import get_logger


logger = get_logger(__name__)


class GUIService:
    """Background service for GUI operations."""

    def __init__(self, config: Config):
        """Initialize GUI service."""
        self.config = config
        self.session_manager = SessionManager(config)
        self.progress_queue = queue.Queue()
        self.current_thread: Optional[threading.Thread] = None
        self.is_processing = False
        self.should_stop = False
        self._gal_stop_flag: dict = {'stop': False}

    def validate_session(self) -> Dict[str, Any]:
        """Validate browser session."""
        return get_session_status(self.config)

    def setup_session(self, progress_callback: Callable = None) -> bool:
        """Set up browser session interactively."""
        try:
            return self.session_manager.setup_interactive_session()
        except Exception as e:
            logger.error(f"Session setup error: {e}")
            return False

    def get_excel_summary(self, excel_path: str) -> Dict[str, Any]:
        """Get Excel file summary."""
        try:
            reader = ExcelReader(excel_path)
            summary = reader.read_pending_emails(batch_size=self.config.processing.batch_size)
            return {
                'total_emails': summary.total_emails,
                'pending_count': summary.pending_count,
                'processed_count': summary.processed_count,
                'batch_count': len(summary.batches)
            }
        except Exception as e:
            logger.error(f"Excel summary error: {e}")
            return {'error': str(e)}

    def start_processing(self, excel_path: str) -> None:
        """Start email processing in background thread."""
        if self.is_processing:
            raise RuntimeError("Processing already active")

        self.should_stop = False
        self.is_processing = True

        def processing_thread():
            try:
                logger.info("Using Playwright automation engine")
                automation = BrowserAutomation(self.config)
                stats = automation.process_emails(excel_path, progress_callback=self._handle_progress)
                self.progress_queue.put(('complete', stats))
            except Exception as e:
                logger.error(f"Processing error: {e}")
                self.progress_queue.put(('error', str(e)))
            finally:
                self.is_processing = False

        self.current_thread = threading.Thread(target=processing_thread, daemon=True)
        self.current_thread.start()

    def _handle_progress(self, current: int, total: int):
        """Handle progress updates from background threads."""
        self.progress_queue.put(('progress', {'current': current, 'total': total}))

    def stop_processing(self):
        """Stop current processing."""
        self.should_stop = True
        self.is_processing = False

    def start_api_processing(self, excel_path: str) -> None:
        """Start API-based contact search in background thread."""
        if self.is_processing:
            raise RuntimeError("Processing already active")

        self.should_stop = False
        self.is_processing = True

        def api_thread():
            try:
                session_file = self.config.get_session_file_path()
                logger.info(f"Starting API search with session: {session_file}")
                result = process_emails_via_api(excel_path, session_file, progress_callback=self._handle_progress)
                self.progress_queue.put(('api_complete', result))
            except Exception as e:
                logger.error(f"API search error: {e}")
                self.progress_queue.put(('api_error', str(e)))
            finally:
                self.is_processing = False

        self.current_thread = threading.Thread(target=api_thread, daemon=True)
        self.current_thread.start()

    def start_gal_scraping(
        self,
        output_dir: str,
        max_contacts: int = 0,
        force_restart: bool = False,
    ) -> None:
        """Start GAL directory scraping in background thread."""
        if self.is_processing:
            raise RuntimeError("Processing already active")

        self.should_stop = False
        self.is_processing = True

        def gal_thread():
            try:
                session_file = self.config.get_session_file_path()
                logger.info(f"Starting GAL scraping with session: {session_file}")

                stop_flag = {'stop': False}
                self._gal_stop_flag = stop_flag

                def progress_callback(count, total):
                    self.progress_queue.put(('gal_progress', {'count': count, 'total': total}))

                result = scrape_gal(
                    session_file=session_file,
                    output_dir=output_dir,
                    max_contacts=max_contacts,
                    force_restart=force_restart,
                    progress_callback=progress_callback,
                    stop_flag=stop_flag,
                )
                self.progress_queue.put(('gal_complete', result))
            except Exception as e:
                logger.error(f"GAL scraping error: {e}")
                self.progress_queue.put(('gal_error', str(e)))
            finally:
                self.is_processing = False

        self.current_thread = threading.Thread(target=gal_thread, daemon=True)
        self.current_thread.start()

    def stop_gal_scraping(self):
        """Signal the GAL scraper to stop."""
        if hasattr(self, '_gal_stop_flag'):
            self._gal_stop_flag['stop'] = True

    def check_queue(self):
        """Check for progress updates."""
        try:
            while True:
                item = self.progress_queue.get_nowait()
                yield item
        except queue.Empty:
            pass
