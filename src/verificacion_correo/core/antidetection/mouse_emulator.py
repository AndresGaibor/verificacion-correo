"""
Human mouse movement emulation module.

Simulates realistic mouse movements using Bézier curves and natural
patterns to avoid bot detection through mouse tracking analysis.
"""

import random
import asyncio
import math
from typing import Tuple, Optional, List
from dataclasses import dataclass


@dataclass
class MouseConfig:
    """Configuration for mouse movement simulation."""
    bezier_curves: bool = True
    random_offset_px: int = 10
    move_duration_ms: Tuple[int, int] = (500, 1500)
    overshoot_chance: float = 0.15  # 15% chance of overshooting target
    pause_before_click_ms: Tuple[int, int] = (50, 150)


class MouseEmulator:
    """
    Emulates human-like mouse movements.

    Features:
    - Bézier curve trajectories (natural curved paths)
    - Random control points for variation
    - Occasional overshooting with correction
    - Random offset from exact target
    - Variable movement speed
    - Natural acceleration/deceleration
    """

    def __init__(self, config: Optional[MouseConfig] = None):
        """
        Initialize mouse emulator.

        Args:
            config: Mouse configuration (uses defaults if None)
        """
        self.config = config or MouseConfig()

    def _bezier_curve(
        self,
        start: Tuple[float, float],
        end: Tuple[float, float],
        control1: Optional[Tuple[float, float]] = None,
        control2: Optional[Tuple[float, float]] = None,
        steps: int = 50
    ) -> List[Tuple[float, float]]:
        """
        Generate points along a cubic Bézier curve.

        Args:
            start: Starting point (x, y)
            end: Ending point (x, y)
            control1: First control point (auto-generated if None)
            control2: Second control point (auto-generated if None)
            steps: Number of points along curve

        Returns:
            List of (x, y) points along the curve
        """
        # Auto-generate control points if not provided
        if control1 is None or control2 is None:
            # Generate control points that create natural curve
            dx = end[0] - start[0]
            dy = end[1] - start[1]

            # Control points offset perpendicular to direct line
            offset_x = dy * random.uniform(-0.3, 0.3)
            offset_y = -dx * random.uniform(-0.3, 0.3)

            control1 = (
                start[0] + dx * 0.33 + offset_x,
                start[1] + dy * 0.33 + offset_y
            )
            control2 = (
                start[0] + dx * 0.66 - offset_x,
                start[1] + dy * 0.66 - offset_y
            )

        # Generate points along Bézier curve
        points = []
        for i in range(steps + 1):
            t = i / steps

            # Cubic Bézier formula
            x = (
                (1 - t) ** 3 * start[0] +
                3 * (1 - t) ** 2 * t * control1[0] +
                3 * (1 - t) * t ** 2 * control2[0] +
                t ** 3 * end[0]
            )
            y = (
                (1 - t) ** 3 * start[1] +
                3 * (1 - t) ** 2 * t * control1[1] +
                3 * (1 - t) * t ** 2 * control2[1] +
                t ** 3 * end[1]
            )

            points.append((x, y))

        return points

    def _apply_random_offset(
        self,
        point: Tuple[float, float]
    ) -> Tuple[float, float]:
        """
        Apply random offset to target point.

        Humans don't click exact pixels - add realistic offset.

        Args:
            point: Original point (x, y)

        Returns:
            Point with random offset applied
        """
        offset_x = random.randint(-self.config.random_offset_px, self.config.random_offset_px)
        offset_y = random.randint(-self.config.random_offset_px, self.config.random_offset_px)

        return (point[0] + offset_x, point[1] + offset_y)

    def _get_overshoot_point(
        self,
        start: Tuple[float, float],
        end: Tuple[float, float]
    ) -> Tuple[float, float]:
        """
        Get overshoot point beyond target.

        Simulates human tendency to slightly overshoot and correct.

        Args:
            start: Starting point (x, y)
            end: Target point (x, y)

        Returns:
            Overshoot point
        """
        dx = end[0] - start[0]
        dy = end[1] - start[1]

        # Overshoot by 5-15% of distance
        overshoot_factor = random.uniform(0.05, 0.15)

        return (
            end[0] + dx * overshoot_factor,
            end[1] + dy * overshoot_factor
        )

    async def move_to_async(
        self,
        page,
        target_x: float,
        target_y: float,
        start_x: Optional[float] = None,
        start_y: Optional[float] = None
    ):
        """
        Move mouse to target position with human-like movement (async).

        Args:
            page: NoDriver page object
            target_x: Target X coordinate
            target_y: Target Y coordinate
            start_x: Starting X coordinate (current position if None)
            start_y: Starting Y coordinate (current position if None)
        """
        # Get current mouse position if start not specified
        if start_x is None or start_y is None:
            # NoDriver doesn't have direct mouse position, assume center
            viewport = await page.evaluate("() => ({width: window.innerWidth, height: window.innerHeight})")
            start_x = start_x or viewport['width'] / 2
            start_y = start_y or viewport['height'] / 2

        start = (start_x, start_y)

        # Apply random offset to target
        target_with_offset = self._apply_random_offset((target_x, target_y))

        # Should we overshoot?
        if random.random() < self.config.overshoot_chance:
            # Move to overshoot point first
            overshoot = self._get_overshoot_point(start, target_with_offset)

            if self.config.bezier_curves:
                points = self._bezier_curve(start, overshoot, steps=30)
            else:
                points = [start, overshoot]

            # Move along path to overshoot
            await self._animate_movement_async(page, points, duration_factor=0.6)

            # Brief pause at overshoot
            await asyncio.sleep(random.uniform(0.05, 0.15))

            # Correct to actual target
            correction_points = self._bezier_curve(overshoot, target_with_offset, steps=20)
            await self._animate_movement_async(page, correction_points, duration_factor=0.4)

        else:
            # Direct movement to target
            if self.config.bezier_curves:
                points = self._bezier_curve(start, target_with_offset, steps=50)
            else:
                points = [start, target_with_offset]

            await self._animate_movement_async(page, points)

    async def _animate_movement_async(
        self,
        page,
        points: List[Tuple[float, float]],
        duration_factor: float = 1.0
    ):
        """
        Animate mouse movement through points.

        Args:
            page: NoDriver page object
            points: List of (x, y) points along path
            duration_factor: Multiplier for duration (for speed variation)
        """
        if len(points) < 2:
            return

        # Calculate total duration
        min_duration, max_duration = self.config.move_duration_ms
        base_duration = random.uniform(min_duration, max_duration) * duration_factor
        delay_per_point = (base_duration / 1000.0) / len(points)

        # Move through each point
        for x, y in points:
            await page.mouse.move(int(x), int(y))
            await asyncio.sleep(delay_per_point)

    async def click_async(
        self,
        page,
        x: float,
        y: float,
        button: str = 'left',
        move_to_target: bool = True
    ):
        """
        Click at position with human-like movement (async).

        Args:
            page: NoDriver page object
            x: X coordinate
            y: Y coordinate
            button: Mouse button ('left', 'right', 'middle')
            move_to_target: Whether to move mouse to position first
        """
        if move_to_target:
            await self.move_to_async(page, x, y)

        # Brief pause before clicking
        min_pause, max_pause = self.config.pause_before_click_ms
        await asyncio.sleep(random.uniform(min_pause, max_pause) / 1000.0)

        # Click
        await page.mouse.click(int(x), int(y), button=button)

    async def click_element_async(
        self,
        page,
        element,
        move_to_target: bool = True
    ):
        """
        Click on element with human-like movement (async).

        Args:
            page: NoDriver page object
            element: Element to click
            move_to_target: Whether to move mouse to element first
        """
        # Get element bounding box
        box = await element.bounding_box()
        if not box:
            # Element not visible, just click directly
            await element.click()
            return

        # Calculate center of element
        center_x = box['x'] + box['width'] / 2
        center_y = box['y'] + box['height'] / 2

        # Click at center with random offset
        await self.click_async(page, center_x, center_y, move_to_target=move_to_target)

    def move_to_sync(
        self,
        page,
        target_x: float,
        target_y: float,
        start_x: Optional[float] = None,
        start_y: Optional[float] = None
    ):
        """
        Move mouse to target position with human-like movement (sync version for Playwright).

        Note: Playwright's sync mouse movement is limited. This provides basic
        offset and pause but not full Bézier curve animation.

        Args:
            page: Playwright page object
            target_x: Target X coordinate
            target_y: Target Y coordinate
            start_x: Starting X (unused in Playwright sync)
            start_y: Starting Y (unused in Playwright sync)
        """
        import time

        # Apply random offset to target
        target_with_offset = self._apply_random_offset((target_x, target_y))

        # Playwright sync API doesn't support smooth animation well
        # So we just move with offset and pause
        page.mouse.move(target_with_offset[0], target_with_offset[1])

        # Pause to simulate movement time
        min_duration, max_duration = self.config.move_duration_ms
        duration = random.uniform(min_duration, max_duration) / 1000.0
        time.sleep(duration)

    def click_sync(
        self,
        page,
        x: float,
        y: float,
        button: str = 'left',
        move_to_target: bool = True
    ):
        """
        Click at position with human-like movement (sync version for Playwright).

        Args:
            page: Playwright page object
            x: X coordinate
            y: Y coordinate
            button: Mouse button ('left', 'right', 'middle')
            move_to_target: Whether to move mouse to position first
        """
        import time

        if move_to_target:
            self.move_to_sync(page, x, y)

        # Brief pause before clicking
        min_pause, max_pause = self.config.pause_before_click_ms
        time.sleep(random.uniform(min_pause, max_pause) / 1000.0)

        # Click
        page.mouse.click(x, y, button=button)


# Convenience functions

async def move_and_click_async(
    page,
    x: float,
    y: float,
    config: Optional[MouseConfig] = None
):
    """
    Move to position and click with human-like movement (async).

    Convenience function for one-off movements.

    Args:
        page: NoDriver page object
        x: X coordinate
        y: Y coordinate
        config: Optional mouse configuration

    Example:
        >>> await move_and_click_async(page, 100, 200)
    """
    emulator = MouseEmulator(config)
    await emulator.click_async(page, x, y, move_to_target=True)


def move_and_click_sync(
    page,
    x: float,
    y: float,
    config: Optional[MouseConfig] = None
):
    """
    Move to position and click with human-like movement (sync).

    Convenience function for Playwright.

    Args:
        page: Playwright page object
        x: X coordinate
        y: Y coordinate
        config: Optional mouse configuration

    Example:
        >>> move_and_click_sync(page, 100, 200)
    """
    emulator = MouseEmulator(config)
    emulator.click_sync(page, x, y, move_to_target=True)
