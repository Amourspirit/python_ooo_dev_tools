import sys
from ooodev.utils.file_io import FileIO
from extract_text import ExtractText

# region maind()
def main() -> int:
    if len(sys.argv) == 2:
        fnm = sys.argv[1]
    else:
        fnm = "tests/fixtures/presentation/algs.odp"
        if not FileIO.is_exist_file(fnm):
            fnm = "../../../../tests/fixtures/presentation/algs.odp"
            FileIO.is_exist_file(fnm, True)

    et = ExtractText(fnm)
    et.extract()
    return 0


# endregion maind()

if __name__ == "__main__":
    SystemExit(main())
