from ooodev.utils.file_io import FileIO
from pathlib import Path
from basic_show import BasicShow

# region maind()
def main() -> int:
    fnm = Path("tests/fixtures/presentation/algs.odp")
    p = FileIO.get_absolute_path(fnm)
    if not p.exists():
        fnm = Path("../../../../../tests/fixtures/presentation/algs.odp")
        p = FileIO.get_absolute_path(fnm)
    if not p.exists():
        raise FileNotFoundError("Unable to find path to algs.odp")
    basic_show = BasicShow(fnm)
    # basic_show.is_automatic = True
    # basic_show.is_full_screen = False
    # basic_show.pause = 3
    # basic_show.start_with_navigator = True
    basic_show.show()
    return 0


# endregion maind()

if __name__ == "__main__":
    SystemExit(main())
