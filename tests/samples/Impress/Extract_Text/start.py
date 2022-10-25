import sys
from ooodev.utils.file_io import FileIO
from pathlib import Path
from extract_text import ExtractText

# region maind()
def main() -> int:
    if len(sys.argv) == 2:
        fnm = sys.argv[1]
    else:
        fnm = Path("tests/fixtures/presentation/algs.odp")
        p = FileIO.get_absolute_path(fnm)
        if not p.exists():
            fnm = Path("../../../../tests/fixtures/presentation/algs.odp")
            p = FileIO.get_absolute_path(fnm)
        if not p.exists():
            raise FileNotFoundError("Unable to find path to algs.odp")
    et = ExtractText(fnm)
    et.extract()
    return 0


# endregion maind()

if __name__ == "__main__":
    SystemExit(main())
