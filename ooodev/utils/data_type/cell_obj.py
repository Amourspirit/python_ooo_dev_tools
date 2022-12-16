from __future__ import annotations
from dataclasses import dataclass
from ..decorator import enforce
from .. import table_helper as mTb
from . import col_obj as mCol
from . import row_obj as mRow


@enforce.enforce_types
@dataclass(frozen=True)
class CellObj:
    """
    Cell Parts

    .. versionadded:: 0.8.2
    """

    col: str
    """Column such as ``A``"""
    row: int
    """Row such as ``125``"""

    def __post_init__(self):
        object.__setattr__(self, "col", self.col.upper())

    @staticmethod
    def from_str(name: str) -> CellObj:
        """
        Gets a ``CellObj`` instance from a string

        Args:
            name (str): Cell Name such ``A1``

        Returns:
            CellObj: Cell Object
        """
        return mTb.TableHelper.get_cell_obj(name)

    def __str__(self) -> str:
        return f"{self.col}{self.row}"

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, CellObj):
            return False
        return str(self) == str(other)

    @property
    def col_info(self) -> mCol.ColObj:
        """Gets Column Info"""
        try:
            return self._col_info
        except AttributeError:
            object.__setattr__(self, "_col_info", mCol.ColObj(self.col))
        return self._col_info

    @property
    def row_info(self) -> mRow.RowObj:
        """Gets Row Info"""
        try:
            return self._row_info
        except AttributeError:
            object.__setattr__(self, "_row_info", mRow.RowObj(self.row))
        return self._row_info
