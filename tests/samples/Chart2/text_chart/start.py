from __future__ import annotations
from pathlib import Path
from text_chart import TextChart
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

    tc = TextChart(data_fnm=fnm)
    tc.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
