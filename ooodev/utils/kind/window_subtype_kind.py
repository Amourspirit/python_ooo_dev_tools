from __future__ import annotations
from enum import Enum


class WindowSubtypeKind(str, Enum):
    BASE_TABLE = "BASETABLE"
    BASE_QUERY = "BASEQUERY"
    BASE_REPORT = "BASEREPORT"
    BASE_DIAGRAM = "BASEDIAGRAM"

    def __str__(self) -> str:
        return self.value
