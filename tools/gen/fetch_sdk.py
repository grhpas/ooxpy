"""Fetch the Open XML SDK schema data at the pinned revision.

    python -m tools.gen.fetch_sdk DEST            # -> DEST/data
    python -m tools.gen.fetch_sdk DEST --revision <sha>

Only the ``data/`` folder is checked out (sparse, shallow), about 12 MB.  The
pinned revision lives in ``tools/gen/SDK_REVISION``; bump it deliberately and
regenerate (``python -m tools.gen --sdk-data DEST/data``) in the same change.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

REPO = "https://github.com/dotnet/Open-XML-SDK.git"
PINNED = (Path(__file__).parent / "SDK_REVISION").read_text().strip()


def git(*args: str, cwd: Path) -> None:
    subprocess.run(["git", *args], cwd=cwd, check=True)


def fetch(dest: Path, revision: str = PINNED) -> Path:
    dest.mkdir(parents=True, exist_ok=True)
    if not (dest / ".git").exists():
        git("init", "-q", cwd=dest)
        git("remote", "add", "origin", REPO, cwd=dest)
        git("sparse-checkout", "set", "data", cwd=dest)
    git("fetch", "-q", "--depth", "1", "--filter=blob:none", "origin", revision, cwd=dest)
    git("checkout", "-q", "FETCH_HEAD", cwd=dest)
    return dest / "data"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="python -m tools.gen.fetch_sdk", description=__doc__)
    ap.add_argument("dest", type=Path)
    ap.add_argument("--revision", default=PINNED)
    args = ap.parse_args(argv)
    data = fetch(args.dest, args.revision)
    print(f"Open XML SDK data @ {args.revision} -> {data}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
