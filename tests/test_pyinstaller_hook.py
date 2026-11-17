"""The PyInstaller hook lists every oox module (they are imported by name at run time)."""

from __future__ import annotations

import runpy
from pathlib import Path

import oox
from oox.__pyinstaller import get_hook_dirs


def test_hook_lists_every_module():
    (hook_dir,) = get_hook_dirs()
    hook = Path(hook_dir) / "hook-oox.py"
    hidden = set(runpy.run_path(str(hook))["hiddenimports"])
    root = Path(oox.__file__).parent
    expected = {
        "oox." + ".".join(p.relative_to(root).with_suffix("").parts).removesuffix(".__init__")
        for p in root.rglob("*.py")
        if "__pyinstaller" not in p.parts and p.stem != "__init__"
    }
    assert expected <= hidden
    assert {"oox.ns.w", "oox.ns.x", "oox.ns.p", "oox.walk"} <= hidden


def test_hook_is_registered():
    text = (Path(__file__).resolve().parents[1] / "pyproject.toml").read_text(encoding="utf-8")
    assert 'hook-dirs = "oox.__pyinstaller:get_hook_dirs"' in text
