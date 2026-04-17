"""Windows adapter shim.

The first split keeps the proven Windows implementation in clipboard_manager.py
and exposes it here as the platform module. Future refactors can move the
pywin32 code into this file without changing the UI boundary.
"""

from __future__ import annotations

from clipboard_manager import WindowsPlatformServices

__all__ = ["WindowsPlatformServices"]

