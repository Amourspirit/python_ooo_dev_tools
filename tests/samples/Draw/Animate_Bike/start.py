from __future__ import annotations
from pathlib import Path
from anim_bicycle import AnimBicycle
from ooodev.utils.file_io import FileIO

# region main()
def main() -> int:
    fnm = Path("tests/fixtures/image/Bicycle-Blue.png")
    p = FileIO.get_absolute_path(fnm)
    if not p.exists():
        fnm = Path("../../../../tests/fixtures/image/Bicycle-Blue.png")
        p = FileIO.get_absolute_path(fnm)
    if not p.exists():
        raise FileNotFoundError("Unable to find path to Bicycle-Blue.png")
    show = AnimBicycle(fnm_bike=fnm)
    show.animate()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
