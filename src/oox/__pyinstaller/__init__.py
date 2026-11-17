"""PyInstaller support: oox loads its grammar tables and element factories with
``importlib`` by namespace prefix, which PyInstaller's import analysis can't see.
PyInstaller finds the hook below through the ``pyinstaller40`` entry point, so
``pyinstaller app.py`` bundles every oox module with no extra options."""

from __future__ import annotations

import os


def get_hook_dirs() -> list[str]:
    return [os.path.dirname(__file__)]
