"""
Human typing pattern simulation module.

Simulates realistic human typing behavior including variable speed,
occasional mistakes, and corrections to avoid bot detection.
"""

import random
import asyncio
from typing import Optional, Tuple
from dataclasses import dataclass


@dataclass
class TypingConfig:
    """Configuration for typing simulation."""
    chars_per_second: Tuple[float, float] = (2.0, 6.0)  # Min, Max chars/sec
    mistake_probability: float = 0.02  # 2% chance of mistake
    correction_delay_ms: Tuple[int, int] = (100, 300)  # Delay before correcting
    between_words_factor: float = 1.5  # Longer pause between words
    burst_typing_chance: float = 0.3  # 30% chance of fast burst
    burst_duration: int = 5  # Number of chars in burst


class TypingSimulator:
    """
    Simulates human-like typing patterns.

    Features:
    - Variable typing speed (realistic range)
    - Random mistakes with backspace corrections
    - Longer pauses between words
    - Occasional "burst typing" (rapid typing)
    - Natural rhythm variations
    """

    def __init__(self, config: Optional[TypingConfig] = None):
        """
        Initialize typing simulator.

        Args:
            config: Typing configuration (uses defaults if None)
        """
        self.config = config or TypingConfig()
        self._in_burst = False
        self._burst_remaining = 0

    def _get_char_delay(self, char: str, next_char: Optional[str] = None) -> float:
        """
        Get delay before typing character.

        Args:
            char: Current character
            next_char: Next character (for context-aware delays)

        Returns:
            Delay in seconds
        """
        # Base delay from typing speed
        min_speed, max_speed = self.config.chars_per_second
        base_delay = 1.0 / random.uniform(min_speed, max_speed)

        # Check if we're in burst typing mode
        if self._in_burst and self._burst_remaining > 0:
            self._burst_remaining -= 1
            if self._burst_remaining == 0:
                self._in_burst = False
            return base_delay * 0.4  # Much faster during burst

        # Start burst typing?
        if random.random() < self.config.burst_typing_chance:
            self._in_burst = True
            self._burst_remaining = self.config.burst_duration
            return base_delay * 0.4

        # Longer delay after space (between words)
        if char == ' ':
            return base_delay * self.config.between_words_factor

        # Longer delay after punctuation
        if char in '.,;:!?':
            return base_delay * 1.8

        # Slightly longer for uppercase (shift key)
        if char.isupper():
            return base_delay * 1.1

        # Add random variation (±20%)
        variation = base_delay * 0.2
        return base_delay + random.uniform(-variation, variation)

    def _should_make_mistake(self, char: str) -> bool:
        """
        Decide if a typing mistake should occur.

        Args:
            char: Character being typed

        Returns:
            True if mistake should occur
        """
        # No mistakes on spaces or punctuation
        if char in ' .,;:!?\n\t':
            return False

        return random.random() < self.config.mistake_probability

    def _get_mistake_char(self, correct_char: str) -> str:
        """
        Get a realistic mistake character.

        Args:
            correct_char: The correct character

        Returns:
            A nearby character on keyboard
        """
        # Keyboard layout proximity (QWERTY)
        keyboard_nearby = {
            'a': 'sq', 'b': 'vn', 'c': 'xv', 'd': 'sf', 'e': 'wr',
            'f': 'dg', 'g': 'fh', 'h': 'gj', 'i': 'uo', 'j': 'hk',
            'k': 'jl', 'l': 'k', 'm': 'n', 'n': 'bm', 'o': 'ip',
            'p': 'o', 'q': 'w', 'r': 'et', 's': 'ad', 't': 'ry',
            'u': 'yi', 'v': 'cb', 'w': 'qe', 'x': 'zc', 'y': 'tu',
            'z': 'x',
        }

        char_lower = correct_char.lower()
        if char_lower in keyboard_nearby:
            nearby = keyboard_nearby[char_lower]
            mistake = random.choice(nearby)
            # Preserve case
            return mistake.upper() if correct_char.isupper() else mistake
        else:
            # Random character if no proximity defined
            return random.choice('abcdefghijklmnopqrstuvwxyz')

    async def type_text_async(self, element, text: str):
        """
        Type text with human-like patterns (async version).

        Args:
            element: NoDriver element to type into
            text: Text to type
        """
        for i, char in enumerate(text):
            next_char = text[i + 1] if i + 1 < len(text) else None

            # Check for mistake
            if self._should_make_mistake(char):
                # Type wrong character
                mistake = self._get_mistake_char(char)
                await element.send_keys(mistake)

                # Wait before realizing mistake
                min_delay, max_delay = self.config.correction_delay_ms
                await asyncio.sleep(random.uniform(min_delay, max_delay) / 1000.0)

                # Backspace
                await element.send_keys('\b')
                await asyncio.sleep(0.05)  # Brief pause after backspace

            # Type correct character
            await element.send_keys(char)

            # Wait before next character
            if i < len(text) - 1:  # Don't wait after last char
                delay = self._get_char_delay(char, next_char)
                await asyncio.sleep(delay)

    def type_text_sync(self, element, text: str):
        """
        Type text with human-like patterns (sync version for Playwright).

        Args:
            element: Playwright element to type into
            text: Text to type
        """
        import time

        for i, char in enumerate(text):
            next_char = text[i + 1] if i + 1 < len(text) else None

            # Check for mistake
            if self._should_make_mistake(char):
                # Type wrong character
                mistake = self._get_mistake_char(char)
                element.type(mistake, delay=0)

                # Wait before realizing mistake
                min_delay, max_delay = self.config.correction_delay_ms
                time.sleep(random.uniform(min_delay, max_delay) / 1000.0)

                # Backspace
                element.press('Backspace')
                time.sleep(0.05)  # Brief pause after backspace

            # Type correct character
            element.type(char, delay=0)

            # Wait before next character
            if i < len(text) - 1:  # Don't wait after last char
                delay = self._get_char_delay(char, next_char)
                time.sleep(delay)

    async def fill_with_typing_async(self, element, text: str):
        """
        Clear and type text into element (async version).

        Args:
            element: NoDriver element to type into
            text: Text to type
        """
        # Clear existing content
        await element.clear()
        await asyncio.sleep(0.1)

        # Type new text
        await self.type_text_async(element, text)

    def fill_with_typing_sync(self, element, text: str):
        """
        Clear and type text into element (sync version for Playwright).

        Args:
            element: Playwright element to type into
            text: Text to type
        """
        import time

        # Clear existing content
        element.fill('')
        time.sleep(0.1)

        # Type new text
        self.type_text_sync(element, text)

    def get_word_typing_time(self, text: str) -> float:
        """
        Estimate time to type given text.

        Useful for setting appropriate timeouts.

        Args:
            text: Text that will be typed

        Returns:
            Estimated time in seconds
        """
        char_count = len(text)
        space_count = text.count(' ')
        punct_count = sum(1 for c in text if c in '.,;:!?')

        # Average delays
        avg_chars_per_sec = sum(self.config.chars_per_second) / 2
        avg_char_delay = 1.0 / avg_chars_per_sec

        # Calculate time
        base_time = char_count * avg_char_delay
        space_time = space_count * avg_char_delay * self.config.between_words_factor
        punct_time = punct_count * avg_char_delay * 0.8

        # Add some buffer for mistakes
        mistake_buffer = char_count * self.config.mistake_probability * 0.5

        total = base_time + space_time + punct_time + mistake_buffer
        return max(1.0, total)  # Minimum 1 second


# Convenience functions

async def type_human_async(element, text: str, config: Optional[TypingConfig] = None):
    """
    Type text with human-like pattern (async).

    Convenience function for one-off typing without TypingSimulator instance.

    Args:
        element: NoDriver element to type into
        text: Text to type
        config: Optional typing configuration

    Example:
        >>> await type_human_async(element, "hello@example.com")
    """
    simulator = TypingSimulator(config)
    await simulator.type_text_async(element, text)


def type_human_sync(element, text: str, config: Optional[TypingConfig] = None):
    """
    Type text with human-like pattern (sync).

    Convenience function for Playwright.

    Args:
        element: Playwright element to type into
        text: Text to type
        config: Optional typing configuration

    Example:
        >>> type_human_sync(element, "hello@example.com")
    """
    simulator = TypingSimulator(config)
    simulator.type_text_sync(element, text)
