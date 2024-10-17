from __future__ import annotations
from ooodev.utils.file_io import FileIO
from points_builder import PointsBuilder
import sys


def main() -> int:
    if len(sys.argv) > 1:
        data_fnm = sys.argv[1]
        FileIO.is_exist_file(data_fnm, True)
    else:
        data_fnm = "tests/fixtures/data/pointsInfo.txt"
        if not FileIO.is_exist_file(data_fnm):
            data_fnm = "../../../../tests/fixtures/data/pointsInfo.txt"
            FileIO.is_exist_file(data_fnm, True)

    pb = PointsBuilder(data_fnm)
    pb.main()

    return 0


if __name__ == "__main__":
    SystemExit(main())
