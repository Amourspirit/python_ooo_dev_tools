import sys
from ooodev.utils.file_io import FileIO
from auto_show import AutoShow  # , FadeEffect


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

    auto_show = AutoShow(fnm)
    # auto_show.duration = 2
    # auto_show.fade_effect = FadeEffect.MOVE_FROM_LEFT
    # auto_show.end_delay = 3
    auto_show.show()
    return 0


# endregion maind()

if __name__ == "__main__":
    SystemExit(main())
