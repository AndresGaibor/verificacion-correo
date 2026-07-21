"""
Background service for GUI operations.

This module provides the GUIService class that handles background processing
tasks for the GUI application.
"""

import threading
import queue
from pathlib import Path
from typing import Optional, Dict, Any, Callable

from verificacion_correo.core.config import Config
from verificacion_correo.core.browser import BrowserAutomation
from verificacion_correo.core.session import SessionManager, get_session_status
from verificacion_correo.core.excel import ExcelReader, ExcelWriter
from verificacion_correo.core.api_extractor import process_emails_via_api, validate_session_api, get_people_filters
from verificacion_correo.core.gal_scraper import scrape_gal, fetch_company_list, save_companies_cache, load_companies_cache
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

    def validate_session_api_quick(self) -> Dict[str, Any]:
        """Quick session validation via API (no Playwright, no browser)."""
        session_file = self.config.get_session_file_path()
        return validate_session_api(session_file)

    def setup_session(self, progress_callback: Callable = None) -> bool:
        """Set up browser session interactively."""
        try:
            return self.session_manager.setup_interactive_session()
        except Exception as e:
            logger.error(f"Session setup error: {e}")
            return False

    def confirm_session_ready(self):
        """Signal that the user has confirmed the session is ready to save."""
        self.session_manager.confirm_session_ready()

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

                def session_health_callback(health_info):
                    self.progress_queue.put(('session_health', health_info))

                result = process_emails_via_api(
                    excel_path, session_file,
                    progress_callback=self._handle_progress,
                    session_health_callback=session_health_callback,
                )
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
        excel_path: str,
        max_contacts: int = 0,
        force_restart: bool = False,
        company_filter: Optional[list] = None,
        address_list_id: Optional[str] = None,
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

                def session_health_callback(health_info):
                    self.progress_queue.put(('session_health', health_info))

                result = scrape_gal(
                    session_file=session_file,
                    excel_path=excel_path,
                    max_contacts=max_contacts,
                    force_restart=force_restart,
                    progress_callback=progress_callback,
                    session_health_callback=session_health_callback,
                    stop_flag=stop_flag,
                    company_filter=company_filter,
                    address_list_id=address_list_id or "fed75805-8ba2-4323-9f6d-80be7e3abc6a",
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

    def start_enrichment(self, excel_path: str) -> None:
        """Start enrichment in background thread."""
        if self.is_processing:
            raise RuntimeError("Processing already active")

        self.should_stop = False
        self.is_processing = True

        def enrich_thread():
            try:
                from verificacion_correo.core.gal_enricher import enrich_excel_by_companies, get_companies_to_enrich_from_excel
                from pathlib import Path

                excel_p = Path(excel_path)

                companies = get_companies_to_enrich_from_excel(excel_p)

                if not companies:
                    self.progress_queue.put(('enrich_complete', {
                        'error': 'No companies selected for enrichment',
                        'contacts_enriched': 0,
                        'companies_done': 0
                    }))
                    self.is_processing = False
                    return

                logger.info(f"Starting enrichment for {len(companies)} companies")

                def progress_callback(enriched_count, total):
                    self.progress_queue.put(('enrich_progress', {
                        'count': enriched_count,
                        'companies': len(companies)
                    }))

                result = enrich_excel_by_companies(
                    excel_p,
                    companies,
                    progress_callback=progress_callback,
                )
                self.progress_queue.put(('enrich_complete', result))
            except Exception as e:
                logger.error(f"Enrichment error: {e}")
                self.progress_queue.put(('enrich_error', str(e)))
            finally:
                self.is_processing = False

        self.current_thread = threading.Thread(target=enrich_thread, daemon=True)
        self.current_thread.start()

    def start_company_scan(self, address_list_id: Optional[str] = None) -> None:
        """Start company list scan from GAL in background thread."""
        if self.is_processing:
            raise RuntimeError("Processing already active")

        self.should_stop = False
        self.is_processing = True

        def scan_thread():
            try:
                session_file = self.config.get_session_file_path()
                logger.info(f"Starting company scan with session: {session_file}")

                kwargs = {}
                if address_list_id:
                    kwargs['address_list_id'] = address_list_id

                companies = fetch_company_list(session_file, **kwargs)

                # Save cache in the same output dir as GAL scraper
                output_dir = Path(self.config.get_excel_file_path()).parent / "gal"
                output_dir.mkdir(parents=True, exist_ok=True)
                save_companies_cache(companies, output_dir)

                self.progress_queue.put(('company_scan_complete', {
                    'companies': companies,
                    'count': len(companies),
                }))
            except Exception as e:
                logger.error(f"Company scan error: {e}")
                self.progress_queue.put(('company_scan_error', str(e)))
            finally:
                self.is_processing = False

        self.current_thread = threading.Thread(target=scan_thread, daemon=True)
        self.current_thread.start()

    def get_cached_companies(self) -> list:
        """Get cached company list from previous scan."""
        output_dir = Path(self.config.get_excel_file_path()).parent / "gal"
        return load_companies_cache(output_dir) or []

    def start_address_list_scan(self) -> None:
        """Start address list scan from OWA in background thread."""
        if self.is_processing:
            raise RuntimeError("Processing already active")

        self.should_stop = False
        self.is_processing = True

        def scan_thread():
            try:
                session_file = self.config.get_session_file_path()
                logger.info(f"Starting address list scan with session: {session_file}")

                address_lists = get_people_filters(session_file)

                self.progress_queue.put(('address_list_scan_complete', {
                    'address_lists': address_lists,
                    'count': len(address_lists),
                }))
            except Exception as e:
                logger.error(f"Address list scan error: {e}")
                self.progress_queue.put(('address_list_scan_error', str(e)))
            finally:
                self.is_processing = False

        self.current_thread = threading.Thread(target=scan_thread, daemon=True)
        self.current_thread.start()

    def check_queue(self):
        """Check for progress updates."""
        try:
            while True:
                item = self.progress_queue.get_nowait()
                yield item
        except queue.Empty:
            pass
