from ooodev.utils.file_io import FileIO
from pathlib import Path
from auto_show import AutoShow, FadeEffect

# region maind()
def main() -> int:
    fnm = Path("tests/fixtures/presentation/algs.odp")
    p = FileIO.get_absolute_path(fnm)
    if not p.exists():
        fnm = Path("../../../../tests/fixtures/presentation/algs.odp")
        p = FileIO.get_absolute_path(fnm)
    if not p.exists():
        raise FileNotFoundError("Unable to find path to algs.odp")
    auto_show = AutoShow(fnm)
    # auto_show.duration = 2
    # auto_show.fade_effect = FadeEffect.MOVE_FROM_LEFT
    # auto_show.end_delay = 3
    auto_show.show()
    return 0


# endregion maind()

if __name__ == "__main__":
    SystemExit(main())
