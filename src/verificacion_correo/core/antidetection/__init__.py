"""
Anti-detection module for verificacion-correo.

This module provides advanced anti-detection techniques to evade
Microsoft OWA's anti-scraping protections, including:

- NoDriver integration (undetected Chrome automation)
- Mouse movement emulation with Bézier curves
- Human typing patterns simulation
- Random delays with realistic distributions
- User-Agent rotation

These techniques significantly improve the success rate of extracting
contact information, particularly the name field which is typically
blocked by OWA's anti-scraping measures.
"""

from .delays import human_delay, DelayManager
from .user_agents import UserAgentRotator
from .typing_simulator import TypingSimulator
from .mouse_emulator import MouseEmulator
from .nodriver_manager import NoDriverManager

__all__ = [
    'human_delay',
    'DelayManager',
    'UserAgentRotator',
    'TypingSimulator',
    'MouseEmulator',
    'NoDriverManager',
]
