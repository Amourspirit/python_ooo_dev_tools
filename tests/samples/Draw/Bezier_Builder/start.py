from __future__ import annotations
from pathlib import Path
from ooodev.utils.file_io import FileIO
from bezier_builder import BezierBuilder


def main() -> int:
    # file name: bpts0.txt or bpts1.txt or bpts2.txt or bpts3.txt
    file_name = "bpts2.txt"
    fnm = Path("tests/fixtures/data", file_name)
    p = FileIO.get_absolute_path(fnm)
    if not p.exists():
        fnm = Path("../../../../tests/fixtures/data", file_name)

    builder = BezierBuilder(fnm_point=fnm)
    builder.show()
    return 0


if __name__ == "__main__":
    SystemExit(main())
