"""
User-Agent rotation module for anti-detection.

Provides realistic User-Agent strings from real browsers to avoid
detection. Uses a curated list of common browser configurations
updated for 2025.
"""

import random
from typing import List, Optional
from dataclasses import dataclass


@dataclass
class UserAgentConfig:
    """Configuration for User-Agent rotation."""
    rotate: bool = True
    pool_size: int = 10
    prefer_platform: Optional[str] = None  # 'windows', 'mac', 'linux', None for random


class UserAgentRotator:
    """
    Manages User-Agent rotation with realistic browser strings.

    Provides a pool of real User-Agent strings from common browsers
    and operating systems to avoid detection through UA fingerprinting.
    """

    # Curated list of realistic User-Agent strings (updated for 2025)
    USER_AGENTS = [
        # Chrome on Windows
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 11.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",

        # Chrome on macOS
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_6_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",

        # Edge on Windows
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0",

        # Edge on macOS
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0",

        # Firefox on Windows
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:134.0) Gecko/20100101 Firefox/134.0",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:133.0) Gecko/20100101 Firefox/133.0",

        # Firefox on macOS
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:134.0) Gecko/20100101 Firefox/134.0",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 14.0; rv:134.0) Gecko/20100101 Firefox/134.0",

        # Firefox on Linux
        "Mozilla/5.0 (X11; Linux x86_64; rv:134.0) Gecko/20100101 Firefox/134.0",
        "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:134.0) Gecko/20100101 Firefox/134.0",

        # Safari on macOS
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.6 Safari/605.1.15",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.6 Safari/605.1.15",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",

        # Chrome on Linux
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
    ]

    def __init__(self, config: Optional[UserAgentConfig] = None):
        """
        Initialize User-Agent rotator.

        Args:
            config: User-Agent configuration (uses defaults if None)
        """
        self.config = config or UserAgentConfig()
        self._current_ua = None
        self._used_uas = []

        # Filter by platform if specified
        if self.config.prefer_platform:
            self._pool = self._filter_by_platform(self.config.prefer_platform)
        else:
            self._pool = self.USER_AGENTS.copy()

        # Limit pool size
        if len(self._pool) > self.config.pool_size:
            self._pool = random.sample(self._pool, self.config.pool_size)

    def _filter_by_platform(self, platform: str) -> List[str]:
        """
        Filter User-Agents by platform.

        Args:
            platform: Platform name ('windows', 'mac', 'linux')

        Returns:
            List of User-Agents for specified platform
        """
        platform_lower = platform.lower()

        if platform_lower == 'windows':
            return [ua for ua in self.USER_AGENTS if 'Windows' in ua]
        elif platform_lower in ['mac', 'macos', 'darwin']:
            return [ua for ua in self.USER_AGENTS if 'Macintosh' in ua]
        elif platform_lower == 'linux':
            return [ua for ua in self.USER_AGENTS if 'Linux' in ua or 'X11' in ua]
        else:
            return self.USER_AGENTS.copy()

    def get_user_agent(self) -> str:
        """
        Get a User-Agent string.

        Returns the same UA if rotation is disabled, otherwise
        returns a random one from the pool.

        Returns:
            User-Agent string
        """
        if not self.config.rotate:
            # Use same UA throughout session
            if not self._current_ua:
                self._current_ua = random.choice(self._pool)
            return self._current_ua

        # Rotate UA - avoid recently used ones
        available = [ua for ua in self._pool if ua not in self._used_uas[-3:]]
        if not available:
            # Reset history if we've exhausted options
            self._used_uas.clear()
            available = self._pool

        ua = random.choice(available)
        self._used_uas.append(ua)

        # Keep history manageable
        if len(self._used_uas) > 10:
            self._used_uas.pop(0)

        self._current_ua = ua
        return ua

    def get_current_user_agent(self) -> Optional[str]:
        """
        Get the currently active User-Agent without rotation.

        Returns:
            Current User-Agent string or None if not set
        """
        return self._current_ua

    def reset(self):
        """Reset rotation history and current UA."""
        self._current_ua = None
        self._used_uas.clear()

    def get_random_platform(self) -> str:
        """
        Get a random platform name.

        Useful for other fingerprinting aspects.

        Returns:
            Platform string ('Windows', 'MacIntel', or 'Linux x86_64')
        """
        if self.config.prefer_platform:
            platform = self.config.prefer_platform.lower()
            if platform == 'windows':
                return 'Win32'
            elif platform in ['mac', 'macos']:
                return 'MacIntel'
            elif platform == 'linux':
                return 'Linux x86_64'

        return random.choice(['Win32', 'MacIntel', 'Linux x86_64'])

    def get_browser_hints(self) -> dict:
        """
        Get browser client hints based on current User-Agent.

        Returns dictionary with platform, brands, etc. for modern
        browser fingerprinting.

        Returns:
            Dictionary with browser hints
        """
        ua = self.get_current_user_agent() or self.get_user_agent()

        # Extract browser and version from UA
        if 'Chrome' in ua and 'Edg' not in ua:
            browser = 'Chrome'
            # Extract version
            import re
            match = re.search(r'Chrome/(\d+)', ua)
            version = match.group(1) if match else '131'
        elif 'Edg' in ua:
            browser = 'Edge'
            match = re.search(r'Edg/(\d+)', ua)
            version = match.group(1) if match else '131'
        elif 'Firefox' in ua:
            browser = 'Firefox'
            match = re.search(r'Firefox/(\d+)', ua)
            version = match.group(1) if match else '134'
        elif 'Safari' in ua and 'Chrome' not in ua:
            browser = 'Safari'
            match = re.search(r'Version/(\d+)', ua)
            version = match.group(1) if match else '17'
        else:
            browser = 'Chrome'
            version = '131'

        # Determine platform
        if 'Windows' in ua:
            platform = 'Windows'
        elif 'Macintosh' in ua:
            platform = 'macOS'
        elif 'Linux' in ua or 'X11' in ua:
            platform = 'Linux'
        else:
            platform = 'Windows'

        return {
            'browser': browser,
            'version': version,
            'platform': platform,
            'brands': [
                {"brand": "Not A(Brand", "version": "99"},
                {"brand": browser, "version": version},
                {"brand": "Chromium", "version": version} if browser in ['Chrome', 'Edge'] else None
            ],
            'mobile': False,
            'platform': platform,
        }


# Convenience function
def get_random_user_agent(prefer_platform: Optional[str] = None) -> str:
    """
    Get a random User-Agent string.

    Convenience function for one-off UA generation.

    Args:
        prefer_platform: Preferred platform ('windows', 'mac', 'linux')

    Returns:
        User-Agent string

    Example:
        >>> ua = get_random_user_agent('windows')
        >>> print(ua)
    """
    config = UserAgentConfig(rotate=True, prefer_platform=prefer_platform)
    rotator = UserAgentRotator(config)
    return rotator.get_user_agent()
