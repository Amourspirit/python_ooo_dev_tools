from __future__ import annotations
import string
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
        # split will cover if a range is passed in, return first cell
        cells = name.split(":")
        col_start = cells[0].rstrip(string.digits).upper()
        row_start = mTb.TableHelper.row_name_to_int(cells[0])
        return CellObj(col=col_start, row=row_start)

    @staticmethod
    def from_idx(col_idx: int, row_idx: int) -> CellObj:
        """
        Gets a ``CellObj`` from zero-based col and row indexes

        Args:
            col_idx (int): Column index
            row_idx (int): Row Index

        Returns:
            CellObj: Cell object
        """
        col = mTb.TableHelper.make_column_name(col=col_idx, zero_index=True)
        return CellObj(col=col, row=row_idx + 1)

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
