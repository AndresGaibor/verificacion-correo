"""
Human-like delay simulation module.

Provides realistic timing patterns that mimic human behavior to avoid
detection by anti-bot systems. Uses random distributions that follow
natural human interaction patterns rather than simple uniform random delays.
"""

import random
import time
import asyncio
from typing import Tuple, Optional
from dataclasses import dataclass


@dataclass
class DelayConfig:
    """Configuration for delay ranges."""
    between_actions: Tuple[int, int] = (500, 2000)  # milliseconds
    between_emails: Tuple[int, int] = (3000, 8000)
    popup_wait: int = 2000
    after_typing: Tuple[int, int] = (200, 800)
    after_click: Tuple[int, int] = (800, 1500)
    after_close_popup: Tuple[int, int] = (1000, 2000)


class DelayManager:
    """
    Manages realistic delays with human-like timing patterns.

    Uses a combination of techniques:
    - Gaussian distribution for natural variation
    - Context-aware delays (different for different actions)
    - Anti-pattern detection (avoids predictable sequences)
    """

    def __init__(self, config: Optional[DelayConfig] = None):
        """
        Initialize delay manager.

        Args:
            config: Delay configuration (uses defaults if None)
        """
        self.config = config or DelayConfig()
        self._last_delay = 0.0
        self._delay_history = []
        self._max_history = 10

    def _gaussian_delay(self, min_ms: int, max_ms: int) -> float:
        """
        Generate delay using Gaussian distribution.

        More realistic than uniform distribution as it clusters around
        the mean while still allowing outliers.

        Args:
            min_ms: Minimum delay in milliseconds
            max_ms: Maximum delay in milliseconds

        Returns:
            Delay in seconds
        """
        # Calculate mean and std dev for Gaussian
        mean = (min_ms + max_ms) / 2
        # std_dev is roughly 1/6 of range (ensures ~99.7% within range)
        std_dev = (max_ms - min_ms) / 6

        # Generate Gaussian value and clamp to range
        delay_ms = random.gauss(mean, std_dev)
        delay_ms = max(min_ms, min(max_ms, delay_ms))

        return delay_ms / 1000.0  # Convert to seconds

    def _add_micro_variation(self, delay: float) -> float:
        """
        Add small random variation to avoid exact timing patterns.

        Args:
            delay: Base delay in seconds

        Returns:
            Delay with micro-variation
        """
        # Add ±5% variation
        variation = delay * 0.05
        return delay + random.uniform(-variation, variation)

    def _avoid_pattern(self, delay: float) -> float:
        """
        Adjust delay if it creates a detectable pattern.

        Args:
            delay: Proposed delay in seconds

        Returns:
            Adjusted delay
        """
        # If we have history, check for patterns
        if len(self._delay_history) >= 3:
            # Check if delays are too similar (pattern detection)
            recent = self._delay_history[-3:]
            if all(abs(d - delay) < 0.1 for d in recent):
                # Add extra variation to break pattern
                delay += random.uniform(-0.2, 0.3)

        return max(0.1, delay)  # Ensure minimum delay

    def _record_delay(self, delay: float):
        """Record delay in history for pattern avoidance."""
        self._delay_history.append(delay)
        if len(self._delay_history) > self._max_history:
            self._delay_history.pop(0)
        self._last_delay = delay

    def get_delay(self, min_ms: int, max_ms: int) -> float:
        """
        Get human-like delay within specified range.

        Args:
            min_ms: Minimum delay in milliseconds
            max_ms: Maximum delay in milliseconds

        Returns:
            Delay in seconds
        """
        delay = self._gaussian_delay(min_ms, max_ms)
        delay = self._add_micro_variation(delay)
        delay = self._avoid_pattern(delay)
        self._record_delay(delay)
        return delay

    def between_actions(self) -> float:
        """Get delay for between general actions."""
        min_ms, max_ms = self.config.between_actions
        return self.get_delay(min_ms, max_ms)

    def between_emails(self) -> float:
        """Get delay for between processing emails."""
        min_ms, max_ms = self.config.between_emails
        return self.get_delay(min_ms, max_ms)

    def after_typing(self) -> float:
        """Get delay after typing action."""
        min_ms, max_ms = self.config.after_typing
        return self.get_delay(min_ms, max_ms)

    def after_click(self) -> float:
        """Get delay after click action."""
        min_ms, max_ms = self.config.after_click
        return self.get_delay(min_ms, max_ms)

    def after_close_popup(self) -> float:
        """Get delay after closing popup."""
        min_ms, max_ms = self.config.after_close_popup
        return self.get_delay(min_ms, max_ms)

    def popup_wait(self) -> float:
        """Get fixed delay for popup loading."""
        # Add some variation to fixed delay
        base = self.config.popup_wait / 1000.0
        return self._add_micro_variation(base)

    def sleep(self, delay: float):
        """
        Sleep for specified duration (sync version).

        Args:
            delay: Delay in seconds
        """
        time.sleep(delay)

    async def sleep_async(self, delay: float):
        """
        Sleep for specified duration (async version).

        Args:
            delay: Delay in seconds
        """
        await asyncio.sleep(delay)


# Convenience functions for simple usage

def human_delay(min_ms: int = 500, max_ms: int = 2000) -> float:
    """
    Generate a human-like delay.

    Convenience function for one-off delays without DelayManager instance.

    Args:
        min_ms: Minimum delay in milliseconds
        max_ms: Maximum delay in milliseconds

    Returns:
        Delay in seconds

    Example:
        >>> delay = human_delay(500, 1500)
        >>> time.sleep(delay)
    """
    manager = DelayManager()
    return manager.get_delay(min_ms, max_ms)


def sleep_human(min_ms: int = 500, max_ms: int = 2000):
    """
    Sleep with human-like timing.

    Convenience function that combines delay generation and sleep.

    Args:
        min_ms: Minimum delay in milliseconds
        max_ms: Maximum delay in milliseconds

    Example:
        >>> sleep_human(1000, 3000)
    """
    delay = human_delay(min_ms, max_ms)
    time.sleep(delay)


async def sleep_human_async(min_ms: int = 500, max_ms: int = 2000):
    """
    Async sleep with human-like timing.

    Convenience function for async contexts.

    Args:
        min_ms: Minimum delay in milliseconds
        max_ms: Maximum delay in milliseconds

    Example:
        >>> await sleep_human_async(1000, 3000)
    """
    delay = human_delay(min_ms, max_ms)
    await asyncio.sleep(delay)
