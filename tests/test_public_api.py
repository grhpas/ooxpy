"""The public API is pinned (docs/STABILITY.md): any change to it shows up here.

A deliberate change -- a new method, a new keyword, a rename before 1.0 --
is made by updating the snapshot in the same commit:

    python -m tools.api_surface --write

and, from 1.0 on, by following the deprecation policy for anything removed
or renamed.  The diff of tests/api/public_api.txt is the API review.
"""

from __future__ import annotations

import difflib

from tools import api_surface


def test_public_api_matches_snapshot():
    want = api_surface.SNAPSHOT.read_text(encoding="utf-8").splitlines()
    got = "\n".join(api_surface.surface()).rstrip().splitlines()
    diff = list(difflib.unified_diff(want, got, "tests/api/public_api.txt", "current", lineterm=""))
    assert not diff, ("the public API changed; if that is intended, run "
                      "`python -m tools.api_surface --write` and review the diff:\n"
                      + "\n".join(diff[:80]))


def test_snapshot_covers_the_exports():
    import oox

    text = api_surface.SNAPSHOT.read_text(encoding="utf-8")
    for name in oox.__all__:
        assert name in text
    # the classes the document methods hand out are part of the contract
    for cls in ("oox.tables.Table", "oox.slidetable.SlideTable", "oox.sections.Section",
                "oox.numbering.ListHandle"):
        assert f"class {cls}(" in text
