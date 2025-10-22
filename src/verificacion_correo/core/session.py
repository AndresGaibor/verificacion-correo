"""
Browser session management for verificacion-correo.

This module handles browser session creation, authentication, and state management
for automated interaction with the OWA interface.
"""

import os
from pathlib import Path
from typing import Optional, Dict, Any
from playwright.sync_api import sync_playwright, Browser, BrowserContext, Page, Playwright

from .config import Config
from ..utils.logging import get_logger


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
                browser = p.chromium.launch(headless=False)
                context = browser.new_context()
                page = context.new_page()

                # Navigate to OWA
                page.goto(self.config.page_url)

                logger.info("Browser opened. Please log in manually.")
                input("When finished, press ENTER in this terminal to save the session...")

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
                self._playwright = sync_playwright().start()

            # Launch browser with session
            self._browser = self._playwright.chromium.launch(
                headless=self.config.browser.headless
            )

            # Create context with saved session state
            self._context = self._browser.new_context(
                storage_state=str(self.session_file),
                viewport={'width': 1280, 'height': 720}
            )

            logger.info(f"Browser context created with session: {self.session_file}")
            return self._context

        except Exception as e:
            logger.error(f"Error creating automation context: {e}")
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
            response = page.goto(self.config.page_url, timeout=10000)

            # Check if we're still authenticated (not redirected to login)
            is_valid = (response and response.status == 200 and
                       page.url != self.config.page_url and
                       'login' not in page.url.lower())

            page.close()
            self._cleanup()

            if is_valid:
                logger.info("Session validation successful")
            else:
                logger.warning("Session validation failed - session may have expired")

            return is_valid

        except Exception as e:
            logger.warning(f"Session validation error: {e}")
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