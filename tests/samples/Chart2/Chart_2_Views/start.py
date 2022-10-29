from __future__ import annotations
from pathlib import Path
from chart_2_views import Chart2View
from ooodev.utils.file_io import FileIO

# region main()
def main() -> int:
    fnm = Path("tests/fixtures/calc/chartsData.ods")
    p = FileIO.get_absolute_path(fnm)
    if not p.exists():
        fnm = Path("../../../../tests/fixtures/calc/chartsData.ods")
        p = FileIO.get_absolute_path(fnm)
    if not p.exists():
        raise FileNotFoundError("Unable to find path to chartsData.ods")
    cv = Chart2View(data_fnm=fnm)
    cv.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
