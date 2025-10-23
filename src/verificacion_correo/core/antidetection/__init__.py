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

from verificacion_correo.core.antidetection.delays import human_delay, DelayManager, DelayConfig
from verificacion_correo.core.antidetection.user_agents import UserAgentRotator, UserAgentConfig
from verificacion_correo.core.antidetection.typing_simulator import TypingSimulator, TypingConfig
from verificacion_correo.core.antidetection.mouse_emulator import MouseEmulator, MouseConfig
from verificacion_correo.core.antidetection.nodriver_manager import NoDriverManager, NoDriverConfig

__all__ = [
    'human_delay',
    'DelayManager',
    'DelayConfig',
    'UserAgentRotator',
    'UserAgentConfig',
    'TypingSimulator',
    'TypingConfig',
    'MouseEmulator',
    'MouseConfig',
    'NoDriverManager',
    'NoDriverConfig',
]
