"""
Browser session management for verificacion-correo.

This module handles browser session creation, authentication, and state management
for automated interaction with the OWA interface.
"""

import os
from pathlib import Path
from typing import Optional, Dict, Any
from playwright.sync_api import sync_playwright, Browser, BrowserContext, Page, Playwright

from verificacion_correo.core.config import Config
from verificacion_correo.utils.logging import get_logger


logger = get_logger(__name__)


class SessionManager:
    """
    Manages browser sessions for OWA automation.

    Handles session creation, authentication, and persistence for reuse
    across multiple automation runs.
    """

    def __init__(self, config: Config):
        """
        Initialize session manager.

        Args:
            config: Application configuration
        """
        self.config = config
        self.session_file = Path(config.get_session_file_path())
        self._playwright: Optional[Playwright] = None
        self._browser: Optional[Browser] = None
        self._context: Optional[BrowserContext] = None

    def setup_interactive_session(self) -> bool:
        """
        Set up an interactive session for manual authentication.

        Opens a browser window for manual login and saves the session state.

        Returns:
            True if session was successfully saved, False otherwise
        """
        logger.info("Starting interactive session setup...")
        logger.info(f"Navigate to: {self.config.page_url}")

        try:
            with sync_playwright() as p:
                # Launch browser (visible for manual interaction)
                browser = p.chromium.launch(
                    headless=False,
                    slow_mo=1000  # Slow down operations to give user time to interact
                )
                context = browser.new_context(
                    viewport={'width': 1280, 'height': 720}
                )
                page = context.new_page()

                # Navigate to OWA
                page.goto(self.config.page_url)

                logger.info("Browser opened. Please log in manually.")
                logger.info("The browser will stay open for 5 minutes to allow login...")
                logger.info("You can close the browser when finished to save the session.")

                # Wait for user to complete login or close browser
                # We'll wait up to 5 minutes (300 seconds)
                try:
                    page.wait_for_timeout(300000)  # 5 minutes
                except:
                    # Browser was closed by user, which is expected
                    pass

                # Save session state
                self._ensure_session_directory()
                context.storage_state(path=str(self.session_file))

                logger.info(f"Session saved to: {self.session_file}")

                context.close()
                browser.close()

            return True

        except Exception as e:
            logger.error(f"Error setting up interactive session: {e}")
            return False

    def create_automation_context(self) -> Optional[BrowserContext]:
        """
        Create browser context with saved session for automation.

        Returns:
            BrowserContext object with saved session, or None if failed
        """
        if not self.session_file.exists():
            logger.error(f"Session file not found: {self.session_file}")
            logger.info("Run session setup first: verificacion-correo-setup")
            return None

        try:
            if not self._playwright:
                # Fix for asyncio event loop conflict when running in threads
                import asyncio
                import threading
                import sys

                # Log thread and loop state for debugging
                thread_id = threading.current_thread().ident
                thread_name = threading.current_thread().name
                logger.debug(f"Creating Playwright context in thread {thread_name} (ID: {thread_id})")

                # Strategy 1: Set environment variable to force sync API
                import os
                os.environ['PLAYWRIGHT_PYTHON_SYNC_API_IN_THREAD'] = '1'
                logger.debug("Set PLAYWRIGHT_PYTHON_SYNC_API_IN_THREAD=1")

                # Strategy 2: Remove event loop for the duration of Playwright operations
                old_loop = None
                loop_was_running = False
                try:
                    old_loop = asyncio.get_event_loop()
                    loop_was_running = old_loop.is_running() if old_loop else False
                    logger.debug(f"Found existing event loop: {old_loop}, running: {loop_was_running}")
                    asyncio.set_event_loop(None)
                    logger.debug("Event loop temporarily removed")
                except RuntimeError as e:
                    logger.debug(f"No event loop in current thread: {e}")
                    pass

                try:
                    logger.debug("Starting sync_playwright()...")
                    self._playwright = sync_playwright().start()
                    logger.debug("sync_playwright() started successfully")
                except Exception as e:
                    logger.error(f"Failed to start sync_playwright(): {e}", exc_info=True)
                    raise

            # Launch browser with session - keep event loop disabled
            try:
                logger.debug(f"Launching browser (headless={self.config.browser.headless})...")
                self._browser = self._playwright.chromium.launch(
                    headless=self.config.browser.headless
                )
                logger.debug("Browser launched successfully")
            except Exception as e:
                # Check if error is due to missing browser executable
                error_str = str(e)
                if "Executable doesn't exist" in error_str or "playwright install" in error_str.lower():
                    logger.warning("Playwright browsers not installed, attempting auto-installation...")
                    print("\n⚠️ Navegadores de Playwright no encontrados.")

                    # Try to install browsers
                    from verificacion_correo.core.first_run import ensure_playwright_browsers_installed
                    if ensure_playwright_browsers_installed():
                        # Retry browser launch after installation
                        logger.info("Retrying browser launch after installation...")
                        print("🔄 Reintentando abrir navegador...")
                        self._browser = self._playwright.chromium.launch(
                            headless=self.config.browser.headless
                        )
                        logger.debug("Browser launched successfully after auto-installation")
                    else:
                        logger.error("Failed to auto-install browsers")
                        raise Exception(
                            "No se pudieron instalar los navegadores automáticamente. "
                            "Por favor ejecuta: playwright install chromium"
                        )
                else:
                    logger.error(f"Failed to launch browser: {e}", exc_info=True)
                    raise

            # Create context with saved session state
            try:
                logger.debug(f"Creating browser context with session: {self.session_file}")
                self._context = self._browser.new_context(
                    storage_state=str(self.session_file),
                    viewport={'width': 1280, 'height': 720}
                )
                logger.debug("Browser context created successfully")
            except Exception as e:
                logger.error(f"Failed to create browser context: {e}", exc_info=True)
                raise

            # Restore event loop only after all Playwright operations complete
            if old_loop is not None:
                asyncio.set_event_loop(old_loop)
                logger.debug(f"Event loop restored: {old_loop}")

            logger.info(f"Browser context created with session: {self.session_file}")
            return self._context

        except Exception as e:
            logger.error(f"Error creating automation context: {e}", exc_info=True)

            # Restore loop even on error
            try:
                import asyncio
                if 'old_loop' in locals() and old_loop is not None:
                    asyncio.set_event_loop(old_loop)
                    logger.debug("Event loop restored after error")
            except:
                pass

            self._cleanup()
            return None

    def get_new_page(self) -> Optional[Page]:
        """
        Create a new page in the current browser context.

        Returns:
            Page object or None if no context is available
        """
        if not self._context:
            context = self.create_automation_context()
            if not context:
                return None

        try:
            page = self._context.new_page()
            logger.debug("Created new browser page")
            return page
        except Exception as e:
            logger.error(f"Error creating new page: {e}")
            return None

    def validate_session(self) -> bool:
        """
        Validate that the saved session is still valid.

        Returns:
            True if session is valid, False otherwise
        """
        if not self.session_file.exists():
            return False

        try:
            # Try to create context and navigate to OWA
            context = self.create_automation_context()
            if not context:
                return False

            page = context.new_page()

            # Use domcontentloaded instead of load (better for SPAs like OWA)
            # Increased timeout from 10s to 30s to handle slow loading
            response = page.goto(
                self.config.page_url,
                timeout=30000,
                wait_until='domcontentloaded'
            )

            # Check if we're still authenticated (not redirected to login)
            # Valid session: no redirect to login page, successful response
            is_valid = (
                response is not None and
                'login' not in page.url.lower() and
                'signin' not in page.url.lower()
            )

            # Additional validation: try to detect OWA interface element
            if is_valid:
                try:
                    # Check for a known OWA element (e.g., new message button)
                    page.wait_for_selector(
                        self.config.selectors.new_message_btn,
                        timeout=5000,
                        state='attached'
                    )
                except:
                    # Element not found, session might be invalid
                    is_valid = False

            page.close()
            self._cleanup()

            if is_valid:
                logger.info("Session validation successful")
            else:
                logger.warning("Session validation failed - session may have expired")

            return is_valid

        except Exception as e:
            logger.warning(f"Session validation error: {e}")
            self._cleanup()
            return False

    def get_session_info(self) -> Dict[str, Any]:
        """
        Get information about the current session.

        Returns:
            Dictionary with session information
        """
        if not self.session_file.exists():
            return {
                'exists': False,
                'file_path': str(self.session_file),
                'is_valid': False
            }

        try:
            # Get file stats
            stat = self.session_file.stat()

            # Try to load session data
            import json
            with open(self.session_file, 'r') as f:
                session_data = json.load(f)

            cookies_count = len(session_data.get('cookies', []))
            origins_count = len(session_data.get('origins', []))

            return {
                'exists': True,
                'file_path': str(self.session_file),
                'file_size': stat.st_size,
                'modified': stat.st_mtime,
                'cookies_count': cookies_count,
                'origins_count': origins_count,
                'is_valid': self.validate_session()
            }

        except Exception as e:
            logger.error(f"Error reading session info: {e}")
            return {
                'exists': True,
                'file_path': str(self.session_file),
                'error': str(e),
                'is_valid': False
            }

    def delete_session(self) -> bool:
        """
        Delete the saved session file.

        Returns:
            True if deletion was successful, False otherwise
        """
        try:
            if self.session_file.exists():
                self.session_file.unlink()
                logger.info(f"Session file deleted: {self.session_file}")
            else:
                logger.warning("Session file does not exist")
            return True
        except Exception as e:
            logger.error(f"Error deleting session file: {e}")
            return False

    def _ensure_session_directory(self):
        """Ensure the directory for the session file exists."""
        self.session_file.parent.mkdir(parents=True, exist_ok=True)

    def _cleanup(self):
        """Clean up browser resources."""
        if self._context:
            self._context.close()
            self._context = None

        if self._browser:
            self._browser.close()
            self._browser = None

        if self._playwright:
            self._playwright.stop()
            self._playwright = None

    def __enter__(self):
        """Context manager entry."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - cleanup resources."""
        self._cleanup()


def setup_session_interactive(config: Optional[Config] = None) -> bool:
    """
    Convenience function to set up an interactive session.

    Args:
        config: Application configuration (optional, will load default if None)

    Returns:
        True if setup was successful, False otherwise
    """
    if config is None:
        from .config import get_config
        config = get_config()

    with SessionManager(config) as session_manager:
        return session_manager.setup_interactive_session()


def validate_saved_session(config: Optional[Config] = None) -> bool:
    """
    Convenience function to validate saved session.

    Args:
        config: Application configuration (optional, will load default if None)

    Returns:
        True if session is valid, False otherwise
    """
    if config is None:
        from .config import get_config
        config = get_config()

    with SessionManager(config) as session_manager:
        return session_manager.validate_session()


def get_session_status(config: Optional[Config] = None) -> Dict[str, Any]:
    """
    Convenience function to get session status information.

    Args:
        config: Application configuration (optional, will load default if None)

    Returns:
        Dictionary with session status information
    """
    if config is None:
        from .config import get_config
        config = get_config()

    with SessionManager(config) as session_manager:
        return session_manager.get_session_info()