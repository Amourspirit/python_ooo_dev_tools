import sys
from ooodev.utils.file_io import FileIO
from slides_info import SlidesInfo


# region main()
def main() -> int:
    if len(sys.argv) > 1:
        fnm = sys.argv[1]
        FileIO.is_exist_file(fnm, True)
    else:
        fnm = "tests/fixtures/presentation/algs.odp"
        if not FileIO.is_exist_file(fnm):
            fnm = "../../../../tests/fixtures/presentation/algs.odp"
            FileIO.is_exist_file(fnm, True)

    slides_info = SlidesInfo(fnm)
    slides_info.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
