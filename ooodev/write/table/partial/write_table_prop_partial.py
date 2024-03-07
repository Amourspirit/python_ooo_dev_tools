from __future__ import annotations

from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from ooodev.write.table.write_table import WriteTable


class WriteTablePropPartial:
    """A partial class for Write Document Table."""

    def __init__(self, obj: WriteTable[Any]) -> None:
        self.__write_table = obj

    @property
    def write_table(self) -> WriteTable[Any]:
        """Write Document."""
        return self.__write_table
