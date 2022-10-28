from __future__ import annotations
from ooodev.utils.file_io import FileIO
from modify_slides import ModifySlides
import sys


def main() -> int:
    im_fnm = "tests/fixtures/image/questions.png"
    if not FileIO.is_exist_file(im_fnm):
        im_fnm = "../../../../tests/fixtures/image/questions.png"
        FileIO.is_exist_file(im_fnm, True)

    if len(sys.argv) > 1:
        fnm = sys.argv[1]
        _ = FileIO.is_exist_file(fnm, True)
    else:
        fnm = "tests/fixtures/presentation/algsSmall.ppt"
        if not FileIO.is_exist_file(fnm):
            fnm = "../../../../tests/fixtures/presentation/algsSmall.ppt"
            _ = FileIO.is_exist_file(fnm, True)

    modify = ModifySlides(fnm=fnm, im_fnm=im_fnm)
    modify.main()

    return 0


if __name__ == "__main__":
    SystemExit(main())
