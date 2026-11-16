"""CLI: python -m tools.gen --sdk-data path/to/Open-XML-SDK/data [--out src/oox]"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import emit, overrides, sdk_frontend


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="python -m tools.gen", description=__doc__)
    ap.add_argument("--sdk-data", required=True, type=Path,
                    help="the data/ directory of a dotnet/Open-XML-SDK checkout")
    ap.add_argument("--out", type=Path, default=Path(__file__).resolve().parents[2] / "src" / "oox")
    ap.add_argument("-v", "--verbose", action="store_true", help="print all warnings")
    ap.add_argument("--allow-other-revision", action="store_true",
                    help="generate from SDK data other than tools/gen/SDK_REVISION")
    args = ap.parse_args(argv)

    schema = sdk_frontend.load(args.sdk_data)
    pinned = (Path(__file__).parent / "SDK_REVISION").read_text().strip()
    if pinned not in schema.source and not args.allow_other_revision:
        print(f"error: SDK data is not the pinned revision {pinned}\n  ({schema.source})\n"
              "  fetch it with `python -m tools.gen.fetch_sdk DEST`, or bump tools/gen/SDK_REVISION,\n"
              "  or pass --allow-other-revision to experiment.", file=sys.stderr)
        return 2
    overrides.apply(schema)
    stats = emit.emit(schema, args.out)
    print(f"source: {schema.source}")
    for k, v in stats.items():
        print(f"  {k:>17}: {v}")
    print(f"  {'warnings':>17}: {len(schema.warnings)}")
    if args.verbose:
        for w in schema.warnings:
            print("   -", w)
    return 0


if __name__ == "__main__":
    sys.exit(main())
