"""
NoDriver browser management module.

Manages NoDriver (undetected Chrome) browser instances with advanced
anti-detection features. NoDriver is the successor to Undetected-Chromedriver
and is specifically designed to evade bot detection systems.
"""

import asyncio
import json
from pathlib import Path
from typing import Optional, Dict, Any
from dataclasses import dataclass

try:
    import nodriver as uc
except ImportError:
    uc = None

from .user_agents import UserAgentRotator, UserAgentConfig
from ...utils.logging import get_logger


logger = get_logger(__name__)


@dataclass
class NoDriverConfig:
    """Configuration for NoDriver browser."""
    headless: bool = False
    browser_executable_path: Optional[str] = None
    user_data_dir: Optional[str] = None
    enable_cdp: bool = True  # Chrome DevTools Protocol
    sandbox: bool = True
    lang: str = "es-ES"


class NoDriverManager:
    """
    Manages NoDriver browser instances with anti-detection features.

    NoDriver provides:
    - Automatic ChromeDriver patching
    - CDP (Chrome DevTools Protocol) leak prevention
    - Stealth mode by default
    - Session persistence support
    """

    def __init__(
        self,
        config: Optional[NoDriverConfig] = None,
        user_agent_config: Optional[UserAgentConfig] = None
    ):
        """
        Initialize NoDriver manager.

        Args:
            config: NoDriver configuration (uses defaults if None)
            user_agent_config: User-Agent rotation configuration
        """
        if uc is None:
            raise ImportError(
                "NoDriver not installed. Install with: pip install nodriver"
            )

        self.config = config or NoDriverConfig()
        self.ua_rotator = UserAgentRotator(user_agent_config or UserAgentConfig())
        self.browser = None
        self.page = None
        self._session_file = None

    async def start(
        self,
        session_file: Optional[str] = None,
        **kwargs
    ):
        """
        Start NoDriver browser instance.

        Args:
            session_file: Path to session state file (for authentication persistence)
            **kwargs: Additional arguments to pass to nodriver.start()

        Returns:
            Browser instance
        """
        logger.info("Starting NoDriver browser...")

        # Store session file path
        self._session_file = session_file

        # Get User-Agent
        user_agent = self.ua_rotator.get_user_agent()
        logger.debug(f"Using User-Agent: {user_agent[:50]}...")

        # Prepare browser arguments
        browser_args = self._get_browser_args()

        # Start browser
        try:
            self.browser = await uc.start(
                headless=self.config.headless,
                browser_executable_path=self.config.browser_executable_path,
                user_data_dir=self.config.user_data_dir,
                browser_args=browser_args,
                **kwargs
            )

            # Get main page/tab
            self.page = await self.browser.get("about:blank")

            # Apply additional anti-detection measures
            await self._apply_stealth_scripts()

            # Load session if provided
            # Try NoDriver-specific session first, then fall back to Playwright session
            nodriver_session = session_file.replace('state.json', 'nodriver_state.json') if session_file else None

            if nodriver_session and Path(nodriver_session).exists():
                logger.info(f"Using NoDriver-specific session: {nodriver_session}")
                await self._load_session(nodriver_session)
                logger.info("NoDriver session loaded successfully")
            elif session_file and Path(session_file).exists():
                logger.warning(f"NoDriver session not found, falling back to Playwright session: {session_file}")
                logger.warning("This may not work correctly. Run: python setup_nodriver_session.py")
                await self._load_session(session_file)
                logger.info("Playwright session loaded (may have compatibility issues)")

            logger.info("NoDriver browser started successfully")
            return self.browser

        except Exception as e:
            logger.error(f"Failed to start NoDriver browser: {e}")
            raise

    def _get_browser_args(self) -> list:
        """
        Get Chrome browser arguments for anti-detection.

        Returns:
            List of Chrome command-line arguments
        """
        args = [
            # User-Agent
            f'--user-agent={self.ua_rotator.get_user_agent()}',

            # Language
            f'--lang={self.config.lang}',
            f'--accept-lang={self.config.lang}',

            # Disable automation flags
            '--disable-blink-features=AutomationControlled',

            # Window size (common resolution)
            '--window-size=1920,1080',

            # Disable various detection vectors
            '--disable-dev-shm-usage',
            '--disable-software-rasterizer',
            '--disable-extensions',
            '--no-first-run',
            '--no-default-browser-check',

            # Memory and performance
            '--disable-background-timer-throttling',
            '--disable-backgrounding-occluded-windows',
            '--disable-renderer-backgrounding',
        ]

        # Sandbox
        if not self.config.sandbox:
            args.append('--no-sandbox')

        return args

    async def _apply_stealth_scripts(self):
        """
        Apply additional stealth scripts to page.

        These scripts modify browser properties that could be used
        for bot detection.
        """
        if not self.page:
            return

        stealth_js = """
        () => {
            // Override webdriver property
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined,
            });

            // Override plugins
            Object.defineProperty(navigator, 'plugins', {
                get: () => [1, 2, 3, 4, 5],
            });

            // Override languages
            Object.defineProperty(navigator, 'languages', {
                get: () => ['es-ES', 'es', 'en-US', 'en'],
            });

            // Override chrome property (should exist in real Chrome)
            if (!window.chrome) {
                window.chrome = {
                    runtime: {},
                    loadTimes: function() {},
                    csi: function() {},
                    app: {},
                };
            }

            // Override permissions
            const originalQuery = window.navigator.permissions.query;
            window.navigator.permissions.query = (parameters) => (
                parameters.name === 'notifications' ?
                    Promise.resolve({ state: Notification.permission }) :
                    originalQuery(parameters)
            );

            // Modify navigator.platform if needed
            Object.defineProperty(navigator, 'platform', {
                get: () => '%PLATFORM%'.replace('%PLATFORM%', navigator.platform),
            });

            // Remove automation-related properties
            delete navigator.__proto__.webdriver;

            // Spoof battery status (laptops have batteries)
            if ('getBattery' in navigator) {
                navigator.getBattery = () => Promise.resolve({
                    charging: Math.random() > 0.5,
                    chargingTime: Math.random() * 10000,
                    dischargingTime: Math.random() * 100000,
                    level: 0.5 + Math.random() * 0.5,
                    addEventListener: () => {},
                });
            }

            // Canvas fingerprint spoofing (basic)
            const originalToDataURL = HTMLCanvasElement.prototype.toDataURL;
            HTMLCanvasElement.prototype.toDataURL = function(type) {
                if (type === 'image/png' && this.width === 220 && this.height === 30) {
                    // Known fingerprinting canvas size
                    return originalToDataURL.apply(this, ['image/png']);
                }
                return originalToDataURL.apply(this, arguments);
            };

            // AudioContext fingerprint spoofing
            const AudioContext = window.AudioContext || window.webkitAudioContext;
            if (AudioContext) {
                const OriginalAnalyser = AudioContext.prototype.createAnalyser;
                AudioContext.prototype.createAnalyser = function() {
                    const analyser = OriginalAnalyser.apply(this, arguments);
                    const original_getFloatFrequencyData = analyser.getFloatFrequencyData;
                    analyser.getFloatFrequencyData = function(array) {
                        const ret = original_getFloatFrequencyData.apply(this, arguments);
                        // Add tiny noise
                        for (let i = 0; i < array.length; i++) {
                            array[i] += Math.random() * 0.0001;
                        }
                        return ret;
                    };
                    return analyser;
                };
            }

            // Screen resolution consistency
            Object.defineProperty(screen, 'availWidth', {
                get: () => screen.width
            });
            Object.defineProperty(screen, 'availHeight', {
                get: () => screen.height - 40  // Taskbar height
            });

            console.log('[AntiDetection] Stealth scripts applied');
        }
        """

        try:
            await self.page.evaluate(stealth_js)
            logger.debug("Stealth scripts applied successfully")
        except Exception as e:
            logger.warning(f"Failed to apply some stealth scripts: {e}")

    async def _load_session(self, session_file: str):
        """
        Load session state from file.

        Args:
            session_file: Path to session state JSON file
        """
        try:
            with open(session_file, 'r') as f:
                session_data = json.load(f)

            # Load cookies using CDP
            if 'cookies' in session_data:
                import nodriver.cdp.network as cdp_network
                for cookie in session_data['cookies']:
                    # Convert Playwright cookie format to CDP format
                    try:
                        await self.page.send(cdp_network.set_cookie(
                            name=cookie.get('name', ''),
                            value=cookie.get('value', ''),
                            domain=cookie.get('domain', ''),
                            path=cookie.get('path', '/'),
                            secure=cookie.get('secure', False),
                            http_only=cookie.get('httpOnly', False),
                            same_site=cdp_network.CookieSameSite(cookie.get('sameSite', 'None').lower()) if cookie.get('sameSite') else None,
                        ))
                    except Exception as cookie_err:
                        logger.debug(f"Failed to set cookie {cookie.get('name')}: {cookie_err}")
                        continue

            # Load localStorage (if available in session)
            if 'origins' in session_data:
                for origin in session_data['origins']:
                    if 'localStorage' in origin:
                        for item in origin['localStorage']:
                            await self.page.evaluate(
                                f"localStorage.setItem('{item['name']}', '{item['value']}')"
                            )

            logger.info(f"Session loaded from {session_file}")

        except Exception as e:
            logger.warning(f"Failed to load session: {e}")

    async def save_session(self, session_file: Optional[str] = None):
        """
        Save current session state to file.

        Args:
            session_file: Path to save session file (uses stored path if None)
        """
        save_path = session_file or self._session_file
        if not save_path:
            logger.warning("No session file path specified")
            return

        try:
            # Get cookies
            cookies = await self.page.send('Network.getAllCookies')

            # Get localStorage (basic approach)
            local_storage = await self.page.evaluate("""
                () => {
                    const ls = {};
                    for (let i = 0; i < localStorage.length; i++) {
                        const key = localStorage.key(i);
                        ls[key] = localStorage.getItem(key);
                    }
                    return ls;
                }
            """)

            # Prepare session data (Playwright-compatible format)
            session_data = {
                'cookies': cookies.get('cookies', []),
                'origins': [
                    {
                        'origin': await self.page.evaluate('location.origin'),
                        'localStorage': [
                            {'name': k, 'value': v}
                            for k, v in local_storage.items()
                        ]
                    }
                ]
            }

            # Save to file
            with open(save_path, 'w') as f:
                json.dump(session_data, f, indent=2)

            logger.info(f"Session saved to {save_path}")

        except Exception as e:
            logger.error(f"Failed to save session: {e}")

    async def get_page(self):
        """
        Get current page/tab.

        Returns:
            Current page object
        """
        if not self.page:
            if self.browser:
                # Get or create page
                tabs = await self.browser.get_all_tabs()
                if tabs:
                    self.page = tabs[0]
                else:
                    self.page = await self.browser.get("about:blank")

        return self.page

    async def close(self):
        """Close browser and cleanup resources."""
        if self.browser:
            try:
                await self.browser.stop()
                logger.info("NoDriver browser closed")
            except Exception as e:
                logger.warning(f"Error closing browser: {e}")
            finally:
                self.browser = None
                self.page = None

    async def __aenter__(self):
        """Async context manager entry."""
        await self.start()
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        """Async context manager exit."""
        await self.close()


# Convenience function
async def create_nodriver_session(
    session_file: Optional[str] = None,
    config: Optional[NoDriverConfig] = None,
    ua_config: Optional[UserAgentConfig] = None
):
    """
    Create and start NoDriver browser session.

    Convenience function for simple usage.

    Args:
        session_file: Path to session state file
        config: NoDriver configuration
        ua_config: User-Agent configuration

    Returns:
        Tuple of (browser, page)

    Example:
        >>> browser, page = await create_nodriver_session("state.json")
        >>> await page.get("https://example.com")
        >>> await browser.stop()
    """
    manager = NoDriverManager(config, ua_config)
    await manager.start(session_file)
    return manager.browser, await manager.get_page()
