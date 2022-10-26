import sys
from pathlib import Path
from ooodev.utils.file_io import FileIO
from custom_show import CustomShow

# region maind()
def main() -> int:
    if len(sys.argv) > 1:
        idxs = [int(s) for s in sys.argv[1:]]
    else:
        # 56 is out of range. when index is out of range it is omitted
        # idxs = [5, 6, 7, 8, 56]
        idxs = [5, 6, 7, 8]
    fnm = "tests/fixtures/presentation/algs.odp"
    if not FileIO.is_exist_file(fnm):
        fnm = "../../../../tests/fixtures/presentation/algs.odp"
        FileIO.is_exist_file(fnm, True)

    fnm = Path("tests/fixtures/presentation/algs.odp")

    custom_show = CustomShow(fnm, *idxs)  # 56 is too high
    custom_show.show()
    return 0


# endregion maind()

if __name__ == "__main__":
    SystemExit(main())
