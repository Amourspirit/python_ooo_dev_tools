import sys
from pathlib import Path
from ooodev.utils.file_io import FileIO
from pathlib import Path
from copy_slide import CopySlide

# region maind()
def main() -> int:
    if len(sys.argv) == 4:
        p = Path(sys.argv[1])
        from_idx = int(sys.argv[2])
        to_idx = int(sys.argv[3])

    else:
        from_idx = 2
        to_idx = 4
        fnm = "tests/fixtures/presentation/algs.odp"
        if not FileIO.is_exist_file(fnm):
            fnm = "../../../../tests/fixtures/presentation/algs.odp"
            FileIO.is_exist_file(fnm, True)
        p = Path(fnm)
    # slide indexes are zero based indexes.
    cs = CopySlide(fnm=p, from_idx=from_idx, to_idx=to_idx)
    cs.copy()
    return 0


# endregion maind()

if __name__ == "__main__":
    SystemExit(main())
