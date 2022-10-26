from __future__ import annotations
from pathlib import Path
from ooodev.utils.file_io import FileIO
from append_slides import AppendSlides
import sys


def main() -> int:
    fnm_lst = []
    if len(sys.argv) > 1:
        fnm_lst.extend(sys.argv[1:])
    else:

        files = ("algs.odp", "points.odp")
        dir_path = Path("tests/fixtures/presentation")
        p = FileIO.get_absolute_path(dir_path)
        if not p.exists():
            dir_path = Path("../../../../tests/fixtures/presentation")
        for file in files:
            fnm_lst.append(Path(dir_path, file))

    appender = AppendSlides(*fnm_lst)
    appender.append()

    return 0


if __name__ == "__main__":
    SystemExit(main())
