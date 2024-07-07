from __future__ import annotations
import contextlib
from typing import Any, cast, overload, TYPE_CHECKING
from dataclasses import dataclass
import uno
from ooo.dyn.table.cell_range_address import CellRangeAddress

from ooodev.loader import lo as mLo
from ooodev.utils import table_helper as mTb
from ooodev.utils.decorator import enforce
from ooodev.loader.inst.doc_type import DocType


if TYPE_CHECKING:
    from ooo.lo.table.cell_address import CellAddress
    from ooodev.calc.calc_doc import CalcDoc


@enforce.enforce_types
@dataclass(frozen=True)
class RangeValues:
    """
    Range Parts. Intended to be zero-based indexes.

    .. versionchanged:: 0.32.0
        Added support for ``__contains__`` and method. If sheet_idx is set to -2 then no attempt is made to get the sheet index from spreadsheet.

    .. versionadded:: 0.8.2
    """

    col_start: int
    """Column start such as ``0``. Must be non-negative integer"""
    col_end: int
    """Column end such as ``3``. Must be non-negative integer"""
    row_start: int
    """Row start such as ``0``. Must be non-negative integer"""
    row_end: int
    """Row end such as ``125``. Must be non-negative integer"""
    sheet_idx: int = -1
    """Sheet index that this range value belongs to"""

    # region dunder

    def __post_init__(self):
        cr_vals = (self.col_start, self.row_start, self.col_end, self.row_end)
        for val in cr_vals:
            if val < 0:
                raise ValueError(
                    f"All indexes must be greater than 0. Column Range ({self.col_start}:{self.col_end}), Row Range - ({self.row_start}:{self.row_end})"
                )

        col_start = self.col_start
        col_end = self.col_end
        if col_start > col_end:
            col_start, col_end = col_end, col_start
            object.__setattr__(self, "col_start", col_start)
            object.__setattr__(self, "col_end", col_end)

        row_start = self.row_start
        row_end = self.row_end
        if row_start > row_end:
            row_start, row_end = row_end, row_start
            object.__setattr__(self, "row_start", row_start)
            object.__setattr__(self, "row_end", row_end)

        if self.sheet_idx == -1:
            with contextlib.suppress(Exception):
                # pylint: disable=no-member
                if mLo.Lo.is_loaded and mLo.Lo.current_doc.DOC_TYPE == DocType.CALC:
                    doc = cast("CalcDoc", mLo.Lo.current_doc)
                    sheet = doc.get_active_sheet()
                    idx = sheet.get_sheet_index()
                    object.__setattr__(self, "sheet_idx", idx)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, RangeValues):
            return (
                self.col_start == other.col_start
                and self.col_end == other.col_end
                and self.row_start == other.row_start
                and self.row_end == other.row_end
                and self.sheet_idx == other.sheet_idx
            )
        if isinstance(other, mRngObj.RangeObj):
            return str(self) == str(other)
        return str(self) == other.upper() if isinstance(other, str) else False

    def __str__(self) -> str:
        start = mTb.TableHelper.make_cell_name(row=self.row_start, col=self.col_start, zero_index=True)
        end = mTb.TableHelper.make_cell_name(row=self.row_end, col=self.col_end, zero_index=True)
        return f"{start}:{end}"

    def __copy__(self) -> RangeValues:
        return RangeValues(
            col_start=self.col_start,
            col_end=self.col_end,
            row_start=self.row_start,
            row_end=self.row_end,
            sheet_idx=self.sheet_idx,
        )

    # endregion dunder

    # region Math
    # region add_rows()
    @overload
    def add_rows(self, num: int) -> RangeValues: ...

    @overload
    def add_rows(self, num: int, to_end: bool) -> RangeValues: ...

    def add_rows(self, num: int, to_end: bool = True) -> RangeValues:
        """
        Gets a new instance with rows added

        Args:
            num (int): Number of row to add.
            to_end (bool, optional): If  ``True`` then adds to ``row_end``; Otherwise, subtracts from ``row_start``. Defaults to ``True``.

        Raises:
            ValueError: If adding rows result is negative

        Returns:
            RangeValues: Instance representing added value
        """
        row_start = self.row_start
        row_end = self.row_end
        if to_end:
            row_end += num
            # could be adding a negative num
            if row_end < 0:
                raise ValueError(f"RangeValues.add_rows(): adding {num} rows to the end results in negative row end.")
        else:
            row_start -= num
            if row_start < 0:
                raise ValueError(
                    f"RangeValues.add_rows(): adding {num} rows to the start results in negative row start."
                )

        if row_start > row_end:
            row_start, row_end = row_end, row_start

        return RangeValues(
            col_start=self.col_start,
            col_end=self.col_end,
            row_start=row_start,
            row_end=row_end,
            sheet_idx=self.sheet_idx,
        )

    # endregion add_rows()

    # region subtract_rows()
    @overload
    def subtract_rows(self, num: int) -> RangeValues: ...

    @overload
    def subtract_rows(self, num: int, from_end: bool) -> RangeValues: ...

    def subtract_rows(self, num: int, from_end: bool = True) -> RangeValues:
        """
        Gets a new instance with rows subtracted

        Args:
            num (int): Number of row to add.
            from_end (bool, optional): If  ``True`` then subtracts from ``row_end``; Otherwise, adds to ``row_start``. Defaults to ``True``.

        Raises:
            ValueError: If subtracting rows result is negative

        Returns:
            RangeValues: Instance representing subtracted value
        """
        row_start = self.row_start
        row_end = self.row_end
        if from_end:
            row_end -= num
            if row_end < 0:
                raise ValueError(
                    f"RangeValues.subtract_rows(): subtracting {num} rows from the end results in negative row end."
                )
        else:
            row_start += num
            # could be subtracting a negative number
            if row_start < 0:
                raise ValueError(
                    f"RangeValues.subtract_rows(): subtracting {num} rows from the start results in negative row start."
                )

        if row_start > row_end:
            row_start, row_end = row_end, row_start

        return RangeValues(
            col_start=self.col_start,
            col_end=self.col_end,
            row_start=row_start,
            row_end=row_end,
            sheet_idx=self.sheet_idx,
        )

    # endregion subtract_rows()
    # region add_cols()
    @overload
    def add_cols(self, num: int) -> RangeValues: ...

    @overload
    def add_cols(self, num: int, to_end: bool) -> RangeValues: ...

    def add_cols(self, num: int, to_end: bool = True) -> RangeValues:
        """
        Gets a new instance with cols added

        Args:
            num (int): Number of row to add.
            to_end (bool, optional): If  ``True`` then adds to ``col_end``; Otherwise, subtracts from ``col_start``. Defaults to ``True``.

        Raises:
            ValueError: If adding cols result is negative

        Returns:
            RangeValues: Instance representing added value
        """
        col_start = self.col_start
        col_end = self.col_end
        if to_end:
            col_end += num
            # could be adding a negative num
            if col_end < 0:
                raise ValueError(f"RangeValues.add_cols(): adding {num} cols to the end results in negative col end.")
        else:
            col_start -= num
            if col_start < 0:
                raise ValueError(
                    f"RangeValues.add_cols(): adding {num} cols to the start results in negative col start."
                )

        if col_start > col_end:
            col_start, col_end = col_end, col_start

        return RangeValues(
            col_start=col_start,
            col_end=col_end,
            row_start=self.row_start,
            row_end=self.row_end,
            sheet_idx=self.sheet_idx,
        )

    # endregion add_cols()
    # region subtract_cols()
    @overload
    def subtract_cols(self, num: int) -> RangeValues: ...

    @overload
    def subtract_cols(self, num: int, from_end: bool) -> RangeValues: ...

    def subtract_cols(self, num: int, from_end: bool = True) -> RangeValues:
        """
        Gets a new instance with cols subtracted

        Args:
            num (int): Number of cols to subtract.
            from_end (bool, optional): If  ``True`` then subtracts from ``col_end``; Otherwise, adds to ``col_start``. Defaults to ``True``.

        Raises:
            ValueError: If subtracting cols result is negative

        Returns:
            RangeValues: Instance representing subtracted value
        """
        col_start = self.col_start
        col_end = self.col_end
        if from_end:
            col_end -= num
            if col_end < 0:
                raise ValueError(
                    f"RangeValues.subtract_rows(): subtracting {num} rows from the end results in negative row end."
                )
        else:
            col_start += num
            # could be subtracting a negative number
            if col_start < 0:
                raise ValueError(
                    f"RangeValues.subtract_rows(): subtracting {num} rows from the start results in negative row start."
                )

        if col_start > col_end:
            col_start, col_end = col_end, col_start

        return RangeValues(
            col_start=col_start,
            col_end=col_end,
            row_start=self.row_start,
            row_end=self.row_end,
            sheet_idx=self.sheet_idx,
        )

    # endregion subtract_cols()
    # endregion Math

    # region from_range()
    @overload
    @staticmethod
    def from_range(range_val: CellRangeAddress) -> RangeValues: ...

    @overload
    @staticmethod
    def from_range(range_val: mRngObj.RangeObj) -> RangeValues: ...

    @overload
    @staticmethod
    def from_range(range_val: str) -> RangeValues: ...

    @staticmethod
    def from_range(range_val: str | mRngObj.RangeObj | CellRangeAddress) -> RangeValues:
        """
        Gets a ``RangeValues`` instance from a range

        Args:
            range (str | mRngObj.RangeObj | CellRangeAddress): Range as object or string (``A2:G23``).

        Returns:
            RangeValues: Object representing range values.
        """
        if isinstance(range_val, mRngObj.RangeObj):
            col_start = mTb.TableHelper.col_name_to_int(range_val.col_start, True)
            col_end = mTb.TableHelper.col_name_to_int(range_val.col_end, True)
            row_start = range_val.row_start - 1
            row_end = range_val.row_end - 1
            sheet_idx = range_val.sheet_idx
        elif isinstance(range_val, str):
            parts = mTb.TableHelper.get_range_parts(range_val)
            col_start = mTb.TableHelper.col_name_to_int(parts.col_start, True)
            col_end = mTb.TableHelper.col_name_to_int(parts.col_end, True)
            row_start = parts.row_start - 1
            row_end = parts.row_end - 1
            sheet_idx = -2
            if parts.sheet:
                with contextlib.suppress(Exception):
                    # pylint: disable=no-member
                    if mLo.Lo.is_loaded and mLo.Lo.current_doc.DOC_TYPE == DocType.CALC:
                        doc = cast("CalcDoc", mLo.Lo.current_doc)
                        sheet = doc.get_sheet(sheet_name=parts.sheet)
                        sheet_idx = sheet.get_sheet_index()
        else:
            # CellRange
            col_start = range_val.StartColumn
            col_end = range_val.EndColumn
            row_start = range_val.StartRow
            row_end = range_val.EndRow
            sheet_idx = range_val.Sheet

        return RangeValues(
            col_start=col_start, row_start=row_start, col_end=col_end, row_end=row_end, sheet_idx=sheet_idx
        )

    # endregion from_range()

    def copy(self) -> RangeValues:
        """
        Gets a copy of the instance

        Returns:
            RangeValues: Copy of the instance

        .. versionadded:: 0.47.5
        """
        return self.__copy__()

    def get_range_obj(self) -> mRngObj.RangeObj:
        """
        Gets a ``RangeObj``

        Returns:
            RangeObj: Range object.
        """
        return mRngObj.RangeObj.from_range(self)

    def get_cell_range_address(self) -> CellRangeAddress:
        """
        Gets a Cell Range Address

        Returns:
            CellRangeAddress: Cell range Address
        """
        return CellRangeAddress(
            Sheet=self.sheet_idx,
            StartColumn=self.col_start,
            StartRow=self.row_start,
            EndColumn=self.col_end,
            EndRow=self.row_end,
        )

    def is_single_col(self) -> bool:
        """
        Gets if instance is a single column or multi-column

        Returns:
            bool: ``True`` if single column; Otherwise, ``False``

        Note:
            If instance is a single cell address then ``True`` is returned.
        """
        return self.col_start == self.col_end

    def is_single_row(self) -> bool:
        """
        Gets if instance is a single row or multi-row

        Returns:
            bool: ``True`` if single row; Otherwise, ``False``

        Note:
            If instance is a single cell address then ``True`` is returned.
        """
        return self.row_start == self.row_end

    def is_single_cell(self) -> bool:
        """
        Gets if a instance is a single cell or a range

        Returns:
            bool: ``True`` if single cell; Otherwise, ``False``
        """
        return self.is_single_col() and self.is_single_row()

    # region contains()
    def __contains__(self, value: Any) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            value (CellObj): Cell object
            value (CellAddress): Cell address
            value (CellValues): Cell Values
            value (str): Cell name

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.

        Note:
            If cell input contains sheet info the it is use in comparison.
            Otherwise sheet is ignored.

        .. versionadded:: 0.32.0
        """
        return self.contains(value)

    @overload
    def contains(self, cell_obj: mCellObj.CellObj) -> bool: ...

    @overload
    def contains(self, cell_addr: CellAddress) -> bool: ...

    @overload
    def contains(self, cell_vals: mCellVals.CellValues) -> bool: ...

    @overload
    def contains(self, cell_name: str) -> bool: ...

    def contains(self, *args, **kwargs) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_obj (CellObj): Cell object
            cell_addr (CellAddress): Cell address
            cell_vals (CellValues): Cell Values
            cell_name (str): Cell name

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.

        Note:
            If cell input contains sheet info the it is use in comparison.
            Otherwise sheet is ignored.
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell_obj", "cell_vals", "cell_name", "cell_addr")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("contains() got an unexpected keyword argument")
            for key in valid_keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("contains() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        arg = kargs[1]
        if isinstance(arg, mCellObj.CellObj):
            cvs = arg.get_cell_values()
            col = cvs.col
            row = cvs.row
            idx = cvs.sheet_idx
        elif isinstance(arg, mCellVals.CellValues):
            col = arg.col
            row = arg.row
            idx = arg.sheet_idx
        elif isinstance(arg, str):
            cvs = mCellVals.CellValues.from_cell(arg)
            col = cvs.col
            row = cvs.row
            idx = cvs.sheet_idx
        else:
            # CellAddress
            ca = cast("CellAddress", arg)
            col = ca.Column
            row = ca.Row
            idx = ca.Sheet

        contains = True
        if idx >= 0:
            contains = contains and self.sheet_idx == idx
        contains = contains and self.col_start <= col
        contains = contains and self.row_start <= row
        contains = contains and self.col_end >= col
        contains = contains and self.row_end >= row
        return contains

    # endregion contains()


from ooodev.utils.data_type import cell_obj as mCellObj
from ooodev.utils.data_type import cell_values as mCellVals
from ooodev.utils.data_type import range_obj as mRngObj
