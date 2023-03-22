from enum import Enum


class CalcStylePageKind(Enum):
    """Style Lookups for Page Styles fo Calc."""

    DEFAULT = "Default"
    REPORT = "Report"

    def __str__(self) -> str:
        return self.value
