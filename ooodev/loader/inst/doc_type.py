from __future__ import annotations
from typing import TYPE_CHECKING
from enum import Enum, IntEnum

if TYPE_CHECKING:
    from ooodev.utils.inst.lo.service import Service


class DocType(IntEnum):
    """Document Type"""

    UNKNOWN = 0
    WRITER = 1
    BASE = 2
    CALC = 3
    DRAW = 4
    IMPRESS = 5
    MATH = 6

    def __str__(self) -> str:
        return str(self.value)

    def get_doc_type_str(self) -> DocTypeStr:
        """Gets the document type as string Enum"""
        return DocTypeStr[self.name]

    def get_service(self) -> Service:
        """Gets the service type as Service Enum"""
        from ooodev.loader.inst.service import Service

        return Service[self.name]


class DocTypeStr(str, Enum):
    """Document Type as string Enum"""

    UNKNOWN = "unknown"
    WRITER = "swriter"
    BASE = "sbase"
    CALC = "scalc"
    DRAW = "sdraw"
    IMPRESS = "simpress"
    MATH = "smath"

    def __str__(self) -> str:
        return self.value

    def get_doc_type(self) -> DocType:
        """Gets the document type as DocType Enum"""
        return DocType[self.name]

    def get_service(self) -> Service:
        """Gets the service type as Service Enum"""
        from ooodev.loader.inst.service import Service

        return Service[self.name]
