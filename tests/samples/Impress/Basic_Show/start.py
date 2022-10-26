import sys
from ooodev.utils.file_io import FileIO
from pathlib import Path
from basic_show import BasicShow

# region maind()
def main() -> int:
    if len(sys.argv) > 1:
        fnm = sys.argv[1]
        FileIO.is_exist_file(fnm, True)
    else:
        fnm = "tests/fixtures/presentation/algs.odp"
        if not FileIO.is_exist_file(fnm):
            fnm = "../../../../tests/fixtures/presentation/algs.odp"
            FileIO.is_exist_file(fnm, True)

    basic_show = BasicShow(fnm)
    basic_show.show()
    return 0


# endregion maind()

if __name__ == "__main__":
    SystemExit(main())
