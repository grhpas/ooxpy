"""PyInstaller hook for oox: bundle every submodule, since the grammar tables
(``oox._generated.g_*``) and factories (``oox.ns.*``) are imported by name at
run time.  The ``.pyi`` stubs are for editors and are not bundled."""

import pkgutil

import oox

hiddenimports = sorted(m.name for m in pkgutil.walk_packages(oox.__path__, "oox.")
                       if not m.name.startswith("oox.__pyinstaller"))
