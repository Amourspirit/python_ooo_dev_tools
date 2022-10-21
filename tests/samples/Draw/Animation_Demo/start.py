from ooodev.utils.file_io import FileIO
from pathlib import Path
from animation_demo import AnimationDemo

# region maind()
def main() -> int:
    # fnm = Path("tests/fixtures/presentation/algs.odp")
    # p = FileIO.get_absolute_path(fnm)
    # if not p.exists():
    #     fnm = Path("../../../../../tests/fixtures/presentation/algs.odp")
    #     p = FileIO.get_absolute_path(fnm)
    # if not p.exists():
    #     raise FileNotFoundError("Unable to find path to algs.odp")
    show = AnimationDemo()
    show.show()
    return 0


# endregion maind()

if __name__ == "__main__":
    SystemExit(main())
