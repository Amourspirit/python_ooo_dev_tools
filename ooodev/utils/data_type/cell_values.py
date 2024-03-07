from __future__ import annotations
import contextlib
from dataclasses import dataclass
from typing import cast, TYPE_CHECKING
import uno
from ooo.dyn.table.cell_address import CellAddress

from ooodev.loader import lo as mLo
from ooodev.utils import table_helper as mTb
from ooodev.utils.decorator import enforce
from ooodev.loader.inst.doc_type import DocType

if TYPE_CHECKING:
    from ooodev.calc.calc_doc import CalcDoc


@enforce.enforce_types
@dataclass(frozen=True)
class CellValues:
    """
    Cell Parts

    .. versionadded:: 0.8.2
    """

    col: int
    """Column such as ``1``. Must be non-negative integer value."""
    row: int
    """Row such as ``125``. Must be non-negative integer value."""
    sheet_idx: int = -1
    """
    Sheet index that this cell value belongs to.
    ``-1`` means no sheet is defined for this instance.
    ``-2`` means no sheet is defined for this instance and not attempt to get the sheet should be made.
    """

    def __post_init__(self):
        # it is ok for sheet_idx to be negative.
        if self.col < 0:
            raise ValueError(f"Value of {self.col} is out of range. Value must be greater then 0")
        if self.row < 0:
            raise ValueError(f"Value of {self.row} is out of range. Value must be greater then or equal to 0")

    def __str__(self) -> str:
        return mTb.TableHelper.make_cell_name(row=self.row, col=self.col, zero_index=True)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, CellValues):
            return self.sheet_idx == other.sheet_idx and self.col == other.col and self.row == other.row
        return str(self) == other.upper() if isinstance(other, str) else False

    def get_cell_address(self) -> CellAddress:
        """
        Gets a cell address

        Raises:
            ValueError: If ``sheet_idx`` is a negative value.

        Returns:
            CellAddress: Cell Address
        """
        if self.sheet_idx < 0:
            raise ValueError(f"Cannot convert to CellAddress because sheet_idx is negative: {self.sheet_idx}")
        return CellAddress(Sheet=self.sheet_idx, Column=self.col, Row=self.row)

    @staticmethod
    def from_cell(cell_val: str | CellAddress | CellValues | mCellObj.CellObj) -> CellValues:
        """
        Gets cell values

        Args:
            cell_val (str | CellAddress | mCellObj.CellObj | CellValues): Cell value. If ``cell_val`` is CellValues instance then that instance is returned.

        Returns:
            CellValues: Cell values representing ``cell_val``
        """
        if isinstance(cell_val, CellValues):
            return cell_val

        if isinstance(cell_val, str):
            # split will cover if a range is passed in, return first cell
            parts = mTb.TableHelper.get_cell_parts(cell_val)
            idx = -2
            row = parts.row - 1
            col = mTb.TableHelper.col_name_to_int(parts.col, True)
            if parts.sheet:
                with contextlib.suppress(Exception):
                    # pylint: disable=no-member
                    if mLo.Lo.is_loaded and mLo.Lo.current_doc.DOC_TYPE == DocType.CALC:
                        doc = cast("CalcDoc", mLo.Lo.current_doc)
                        sheet = doc.get_sheet(sheet_name=parts.sheet)
                        idx = sheet.get_sheet_index()
        elif isinstance(cell_val, mCellObj.CellObj):
            idx = cell_val.sheet_idx
            row = cell_val.row - 1
            col = cell_val.col_obj.index
        else:
            idx = cell_val.Sheet
            row = cell_val.Row
            col = cell_val.Column
        return CellValues(col=col, row=row, sheet_idx=idx)


from ooodev.utils.data_type import cell_obj as mCellObj
