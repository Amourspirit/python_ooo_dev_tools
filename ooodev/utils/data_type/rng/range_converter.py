from __future__ import annotations
from typing import Any, cast, overload, Sequence, TYPE_CHECKING, Tuple
import contextlib
import uno
from com.sun.star.container import XNamed
from com.sun.star.sheet import XCellAddressable
from com.sun.star.sheet import XCellRangeAddressable

from ooodev.events.args.event_args import EventArgs
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.data_type.cell_obj import CellObj
from ooodev.utils.data_type.cell_values import CellValues
from ooodev.utils.data_type.range_obj import RangeObj
from ooodev.utils.data_type.range_values import RangeValues
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.table_helper import TableHelper
from ooodev.loader.inst.doc_type import DocType

if TYPE_CHECKING:
    from com.sun.star.table import CellRangeAddress
    from com.sun.star.table import XCellRange
    from com.sun.star.sheet import XSpreadsheet
    from com.sun.star.table import CellAddress
    from com.sun.star.table import XCell
    from ooodev.calc.calc_doc import CalcDoc


class RangeConverter(LoInstPropsPartial):
    EVENT_RANGE_CREATING = "range_converter_range_creating"
    EVENT_CELL_CREATING = "range_converter_cell_creating"

    def __init__(self, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)

    def get_cell_values(self, cell: Any) -> CellValues:
        """
        Gets the cell values from a cell like object.

        Args:
            cell (XCell): Cell.

        Returns:
            Tuple[int, int]: Column and Row.

        See Also:
            - :py:meth:`~ooodev.utils.data_type.rng.range_converter.get_cell_obj`
        """
        try:
            cell_obj = self.get_cell_obj(cell)
            return cell_obj.get_cell_values()
        except Exception as e:
            raise TypeError("cell must be an object that can be converted to a CellObj.") from e

    def get_range_values(self, rng: Any) -> RangeValues:
        """
        Gets the range values from a range like object.

        Args:
            rng (Any): Any value that can be converted to RangeObj.

        Returns:
            Tuple[int, int]: Column and Row.

        See Also:
            - :py:meth:`~ooodev.utils.data_type.rng.range_converter.get_range_obj`
        """
        try:
            range_obj = self.get_range_obj(rng)
            return range_obj.get_range_values()
        except Exception as e:
            raise TypeError("rng must be an object that can be converted to RangeObj.") from e

    def get_safe_quoted_name(self, name: str) -> str:
        """
        Returns the name quoted if it is not alphanumeric.

        Args:
            name (str): Name

        Returns:
            str: Quoted name
        """
        return name if name.isalnum() else f"'{name}'"

    # region Cell Converters to String

    def get_cell_str_from_col_row(self, col: int, row: int) -> str:
        """
        Gets the cell as string from column and row.

        Args:
            col (int): Column. Zero Based column index.
            row (int): Row. Zero Based row index.

        Returns:
            str: Cell as string in the format of ``A1``.
        """
        return f"{self.column_number_str(col)}{row + 1}"

    def get_cell_str_from_addr(self, addr: CellAddress) -> str:
        """
        Gets the cell as string from a cell address.

        Args:
            addr (CellAddress): Cell Address.

        Returns:
            str: Cell as string in the format of ``A1``.
        """
        return self.get_cell_str_col_row(col=addr.Column, row=addr.Row)

    def get_cell_str_from_cell(self, cell: XCell) -> str:
        """
        Gets the cell as string from a cell.

        Args:
            cell (XCell): Cell.

        Returns:
            str: Cell as string in the format of ``A1``.
        """
        addr = self.get_cell_address_from_cell(cell)
        return self.get_cell_str_from_addr(addr)

    def get_cell_str_from_cell_obj(self, cell_obj: CellObj) -> str:
        """
        Gets the cell as string from a cell object.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            str: Cell as string in the format of ``A1``.
        """
        return str(cell_obj)

    # region get_cell() method
    @overload
    def get_cell(self, col: int, row: int) -> str:
        """
        Gets the cell as string from column and row.

        Args:
            col (int): Column. Zero Based column index.
            row (int): Row. Zero Based row index.

        Returns:
            str: Cell as string in the format of ``A1``.
        """
        ...

    @overload
    def get_cell(self, addr: CellAddress) -> str:
        """
        Gets the cell as string from a cell address.

        Args:
            addr (CellAddress): Cell Address.

        Returns:
            str: Cell as string in the format of ``A1``.
        """
        ...

    @overload
    def get_cell(self, cell: XCell) -> str:
        """
        Gets the cell as string from a cell.

        Args:
            cell (XCell): Cell.

        Returns:
            str: Cell as string in the format of ``A1``.
        """
        ...

    @overload
    def get_cell(self, cell_obj: CellObj) -> str:
        """
        Gets the cell as string from a cell object.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            str: Cell as string in the format of ``A1``.
        """
        return str(cell_obj)

    def get_cell(self, *args, **kwargs) -> str:
        """
        Gets the cell as string from a cell.

        Args:
            col (int): Column. Zero Based column index.
            row (int): Row. Zero Based row index.
            addr (CellAddress): Cell Address.
            cell (XCell): Cell.
            cell_obj (CellObj): Cell Object.


        Returns:
            str: Cell as string in the format of ``A1``.
        """
        # get_cell_str_from_col_row(self, col: int, row: int)
        # get_cell_str_from_addr(self, addr: CellAddress)
        # get_cell_str_from_cell(self, cell: XCell)
        # def get_cell_str_from_cell_obj(self, cell_obj: CellObj) -> str:
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = {"col", "addr", "cell", "cell_obj"}
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_cell() got an unexpected keyword argument")
            for key in valid_keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            if "row" in kwargs:
                ka[2] = kwargs["row"]
            return ka

        if count not in (1, 2):
            raise TypeError("get_cell() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            return self.get_cell_str_from_col_row(kargs[1], kargs[2])
        arg1 = kargs[1]
        if hasattr(arg1, "typeName") and getattr(arg1, "typeName") == "com.sun.star.table.CellAddress":
            return self.get_cell_str_from_addr(arg1)

        if mInfo.Info.is_instance(arg1, CellObj):
            return self.get_cell_str_from_cell_obj(arg1)

        return self.get_cell_str_from_cell(arg1)

    # endregion get_cell() method

    # endregion Cell Converters to String

    # region Cell Convertor to CellObj
    def _create_cell_obj(self, col: str, row: int, sheet_idx: int = -2) -> CellObj:
        """
        Creates a cell object

        Args:
            col (str): Column such as ``A``
            row (int): Row such as ``1``
            sheet_idx (int, optional): Sheet index that this cell value belongs to.
            kwargs: Additional arguments to pass to the event data.

        Returns:
            CellObj: Cell object.
        """
        return CellObj(
            col=col,
            row=row,
            sheet_idx=sheet_idx,
        )

    def get_cell_obj_from_col_row(self, col: int, row: int, sheet_idx: int = -2) -> CellObj:
        """
        Gets the cell as CellObj from column and row.

        Args:
            col (int): Column. Zero Based column index.
            row (int): Row. Zero Based row index.
            sheet_idx (int, optional): Sheet index that this cell value belongs to.
                A value of ``-1`` means the sheet index is not set and an attempt
                will be made to discover the sheet index from current document if it is a Calc document.
                A value of ``-2`` means no attempt is made to discover the sheet index.
                Default is ``-2``.

        Returns:
            CellObj: Cell Object.
        """
        col_str = TableHelper.make_column_name(col=col, zero_index=True)
        return self._create_cell_obj(col=col_str, row=row + 1, sheet_idx=sheet_idx)

    def get_cell_obj_from_addr(self, addr: CellAddress) -> CellObj:
        """
        Gets the cell as CellObj from a cell address.

        Args:
            addr (CellAddress): Cell Address.

        Returns:
            CellObj: Cell Object.
        """
        col_str = TableHelper.make_column_name(col=addr.Column, zero_index=True)
        return self._create_cell_obj(col=col_str, row=addr.Row + 1, sheet_idx=addr.Sheet)

    def get_cell_obj_from_cell(self, cell: XCell) -> CellObj:
        """
        Gets the cell as CellObj from a cell.

        Args:
            cell (XCell): Cell.

        Returns:
            CellObj: Cell Object.
        """
        addr = self.get_cell_address_from_cell(cell)
        return self.get_cell_obj_from_addr(addr)

    def get_cell_obj_from_cell_obj(self, val: CellValues) -> CellObj:
        """
        Gets the cell as CellObj from CellValues.

        Args:
            val (CellValues): Cell values.

        Returns:
            CellObj: Cell Object.

        Hint:
            - ``CellValues`` can be imported from ``ooodev.utils.data_type.cell_values``
        """
        col_str = TableHelper.make_column_name(col=val.col, zero_index=True)
        return self._create_cell_obj(col=col_str, row=val.row + 1, sheet_idx=val.sheet_idx)

    def get_cell_obj_tuple(self, values: Tuple[int, int] | Tuple[int, int, int]) -> CellObj:
        """
        Gets the cell as CellObj from CellValues.

        Args:
            val (Tuple[int, int], Tuple[int, int, int]): Cell values.
                Tuple of (col, row) or (col, row, sheet_idx). All values are Zero Based.

        Returns:
            CellObj: Cell Object.
        """
        col = TableHelper.make_column_name(values[0], True)
        row = values[1] + 1
        sheet_idx = values[2] if len(values) == 3 else -2
        return self._create_cell_obj(col=col, row=row, sheet_idx=sheet_idx)

    def get_cell_obj_from_str(self, name: str) -> CellObj:
        """
        Gets the cell as CellObj from a cell name.

        Args:
            name (str): Cell name such as as ``A23`` or ``Sheet1.A23``

        Returns:
            CellObj: Cell Object.
        """
        parts = TableHelper.get_cell_parts(name)
        idx = self.get_sheet_index(parts.sheet) if parts.sheet else -2
        return self._create_cell_obj(col=parts.col, row=parts.row, sheet_idx=idx)

    # region get_cell_obj() method
    @overload
    def get_cell_obj(self, values: Tuple[int, int] | Tuple[int, int, int]) -> CellObj:
        """
        Gets the cell as CellObj from CellValues.

        Args:
            val (Tuple[int, int], Tuple[int, int, int]): Cell values.
                Tuple of (col, row) or (col, row, sheet_idx). All values are Zero Based.

        Returns:
            CellObj: Cell Object.
        """
        ...

    @overload
    def get_cell_obj(self, col: int, row: int, sheet_idx: int = ...) -> CellObj:
        """
        Gets the cell as CellObj from column and row.

        Args:
            col (int): Column. Zero Based column index.
            row (int): Row. Zero Based row index.
            sheet_idx (int, optional): Sheet index that this cell value belongs to.
                A value of ``-1`` means the sheet index is not set and an attempt
                will be made to discover the sheet index from current document if it is a Calc document.
                A value of ``-2`` means no attempt is made to discover the sheet index.
                Default is ``-2``.

        Returns:
            CellObj: Cell Object.
        """
        ...

    @overload
    def get_cell_obj(self, addr: CellAddress) -> CellObj:
        """
        Gets the cell as CellObj from a cell address.

        Args:
            addr (CellAddress): Cell Address.

        Returns:
            CellObj: Cell Object.
        """
        ...

    @overload
    def get_cell_obj(self, cell: XCell) -> CellObj:
        """
        Gets the cell as CellObj from a cell.

        Args:
            cell (XCell): Cell.

        Returns:
            CellObj: Cell Object.
        """
        ...

    @overload
    def get_cell_obj(self, val: CellValues) -> CellObj:
        """
        Gets the cell as CellObj from CellValues.

        Args:
            val (CellValues): Cell values.

        Returns:
            CellObj: Cell Object.

        Hint:
            - ``CellValues`` can be imported from ``ooodev.utils.data_type.cell_values``
        """
        ...

    @overload
    def get_cell_obj(self, name: str) -> CellObj:
        """
        Gets the cell as CellObj from a cell name.

        Args:
            name (str): Cell name such as as ``A23`` or ``Sheet1.A23``

        Returns:
            CellObj: Cell Object.
        """
        ...

    @overload
    def get_cell_obj(self, cell_obj: CellObj) -> CellObj:
        """
        Gets CellObj. Returns the same object.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            CellObj: Cell Object.
        """
        ...

    def get_cell_obj(self, *args, **kwargs) -> CellObj:
        """
        Gets the cell as string from a cell.

        Args:
            col (int): Column. Zero Based column index.
            row (int): Row. Zero Based row index.
            sheet_idx (int, optional): Sheet index that this cell value belongs to. Default is ``-2``.
            addr (CellAddress): Cell Address.
            cell (XCell): Cell.
            val (CellValues): Cell values.
            name (str): Cell name such as as ``A23`` or ``Sheet1.A23``
            cell_obj (CellObj): Cell Object.
            values (Tuple[int, int], Tuple[int, int, int]): Cell values.


        Returns:
            CellObj: Cell Object.
        """
        # def get_cell_obj_from_col_row(self, col: int, row: int, sheet_idx: int = -1)
        # def get_cell_obj_from_addr(self, addr: CellAddress)
        # def get_cell_obj_from_cell(self, cell: XCell)
        # def get_cell_obj_from_cell_obj(self, val: CellValues)
        # def get_cell_obj_from_str(self, name: str)

        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = {"col", "row", "addr", "cell", "val", "name", "values", "cell_obj", "sheet_idx"}
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_cell_obj() got an unexpected keyword argument")
            keys = {"col", "addr", "cell", "val", "name", "values", "cell_obj"}
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            if "row" in kwargs:
                ka[2] = kwargs["row"]
            if count == 2:
                return ka
            if "sheet_idx" in kwargs:
                ka[3] = kwargs["sheet_idx"]
            return ka

        if count not in (1, 2, 3):
            raise TypeError("get_cell_obj() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            return self.get_cell_obj_from_col_row(kargs[1], kargs[2], -2)
        if count == 3:
            return self.get_cell_obj_from_col_row(kargs[1], kargs[2], kargs[3])
        arg1 = kargs[1]
        if mInfo.Info.is_instance(arg1, CellObj):
            return arg1
        if hasattr(arg1, "typeName") and getattr(arg1, "typeName") == "com.sun.star.table.CellAddress":
            return self.get_cell_obj_from_addr(arg1)
        if mInfo.Info.is_instance(arg1, tuple):
            return self.get_cell_obj_tuple(arg1)
        if mInfo.Info.is_instance(arg1, CellValues):
            return self.get_cell_obj_from_cell_obj(arg1)
        if mInfo.Info.is_instance(arg1, str):
            return self.get_cell_obj_from_str(arg1)
        return self.get_cell_obj_from_cell(arg1)

    # endregion get_cell_obj() method
    # endregion Cell Convertor to CellObj

    # region Cell Converter to CellAddress
    def get_cell_address_from_cell(self, cell: XCell) -> CellAddress:
        """
        Gets a cell address from a cell.

        Args:
            cell (XCell): Cell.

        Returns:
            CellAddress: Cell address.
        """
        addr = mLo.Lo.qi(XCellAddressable, cell, True)
        return addr.getCellAddress()

    def get_cell_address_sheet(self, sheet: XSpreadsheet, cell_name: str) -> CellAddress:
        """
        Gets a cell address from a sheet using the cell name.

        Args:
            sheet (XSpreadsheet): Sheet.
            cell_name (str): Cell name.

        Returns:
            CellAddress: Cell address.
        """
        cell_range = sheet.getCellRangeByName(cell_name)
        start_cell = self.get_cell_from_cell_range(cell_range=cell_range, col=0, row=0)
        return self.get_cell_address_from_cell(start_cell)

    # endregion Cell Converter to CellAddress

    # region Cell Converters to XCell
    def get_cell_from_cell_range(self, cell_range: XCellRange, col: int, row: int) -> XCell:
        """
        Gets a cell from a cell range.

        Args:
            cell_range (XCellRange): Cell Range.
            col (int): Column. Zero Based column index.
            row (int): Row. Zero Based row index.

        Returns:
            XCell: Cell.
        """
        return cell_range.getCellByPosition(col, row)

    def get_cell_from_sheet_addr(self, sheet: XSpreadsheet, addr: CellAddress) -> XCell:
        """
        Gets a cell from a sheet using the cell address.

        Args:
            sheet (XSpreadsheet): Sheet.
            addr (CellAddress): Cell Address.

        Returns:
            XCell: Cell.
        """
        # not using Sheet value in addr
        return self.get_cell_by_position(sheet=sheet, col=addr.Column, row=addr.Row)

    def get_cell_by_position(self, sheet: XSpreadsheet, col: int, row: int) -> XCell:
        """
        Gets a cell from a sheet using the column and row.

        Args:
            sheet (XSpreadsheet): Sheet.
            col (int): Column. Zero Based column index.
            row (int): Row. Zero Based row index.

        Returns:
            XCell: Cell.
        """
        return sheet.getCellByPosition(col, row)

    def get_cell_sheet_cell(self, sheet: XSpreadsheet, rng_name: str) -> XCell:
        """
        Gets a cell from a sheet using the cell name.

        Args:
            sheet (XSpreadsheet): Sheet.
            rng_name (str): Range name such as 'A1:D7'.

        Returns:
            XCell: Cell.

        Note:
            In spreadsheets valid names may be ``A1:C5`` or ``$B$2`` or even defined names for cell ranges such as ``MySpecialCell``.
        """
        cell_range = sheet.getCellRangeByName(rng_name)
        return self.get_cell_from_cell_range(cell_range=cell_range, col=0, row=0)

    # endregion Cell Converters to XCell

    # region Range Converters to String
    def get_rng_str_from_cell_rng(self, cell_range: XCellRange, sheet: XSpreadsheet) -> str:
        """
        Return as str using the name taken from the sheet works, Sheet1.A1:B2

        Args:
            cell_range (XCellRange): Cell Range.
            sheet (XSpreadsheet): Sheet.

        Returns:
            str: Range as string.
        """
        return self.get_range_str_from_cr_addr_sheet(
            self.get_cell_address_from_cell_range(cell_range=cell_range), sheet
        )

    def get_range_str_from_cr_addr_sheet(self, cr_addr: CellRangeAddress, sheet: XSpreadsheet) -> str:
        """
        Return as str using the name taken from the sheet works, Sheet1.A1:B2

        Args:
            cr_addr (CellRangeAddress): Cell Range Address.
            sheet (XSpreadsheet): Sheet.

        Returns:
            str: Range as string.
        """
        return f"{self.get_sheet_name(sheet=sheet)}.{self.get_range_str_from_cr_addr(cr_addr)}"

    def get_range_str_from_cr_addr(self, cr_addr: CellRangeAddress) -> str:
        """
        Gets the range as string from a Cell Range Address.

        Args:
            cr_addr (CellRangeAddress): Cell Range Address.

        Returns:
            str: Range as string. In the format of ``A1:B2``.
        """
        result = f"{self.get_cell_str_col_row(cr_addr.StartColumn, cr_addr.StartRow)}:"
        result += f"{self.get_cell_str_col_row(cr_addr.EndColumn, cr_addr.EndRow)}"
        return result

    def get_cell_str_col_row(self, col: int, row: int) -> str:
        """LO Safe Method"""
        if col < 0 or row < 0:
            mLo.Lo.print("Cell position is negative; using A1")
            return "A1"
        return f"{self.column_number_str(col)}{row + 1}"

    def get_cell_address_from_cell_range(self, cell_range: XCellRange) -> CellRangeAddress:
        """
        Gets the cell address from a cell range.

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            CellRangeAddress: Cell address.
        """
        addr = mLo.Lo.qi(XCellRangeAddressable, cell_range, True)
        return addr.getRangeAddress()  # type: ignore

    def column_number_str(self, col: int) -> str:
        """
        Creates a column Name from zero base column number.

        Columns are numbered starting at 0 where 0 corresponds to ``A``
        They run as ``A-Z``, ``AA-AZ``, ``BA-BZ``, ..., ``IV``

        Args:
            col (int): Zero based column index

        Returns:
            str: Column Name
        """
        num = col + 1  # shift to one based.
        return TableHelper.make_column_name(num)

    def get_sheet_name(self, sheet: XSpreadsheet, safe_quote: bool = True) -> str:
        """
        Gets the name of a sheet.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            safe_quote (bool, optional): If True, returns quoted (in single quotes) sheet name if the sheet name is not alphanumeric.
                Default is ``True``.

        Returns:
            str: Name of sheet
        """
        xnamed = mLo.Lo.qi(XNamed, sheet, True)
        name = xnamed.getName()
        return self.get_safe_quoted_name(name) if safe_quote else name

    # endregion Range Converters to String

    # region Range Converters to RangeObj
    def _create_range_obj(
        self, col_start: str, col_end: str, row_start: int, row_end: int, sheet_idx: int = -2
    ) -> RangeObj:
        """
        Creates a cell range object

        Args:
            col_start (str): Column start such as ``A``
            col_end (str): Column end such as ``C``
            row_start (int): Row start such as ``1``
            row_end (int): Row end such as ``125``
            sheet_idx (int, optional): Sheet index that this range value belongs to.
                A value of ``-1`` means the sheet index is not set and an attempt
                will be made to discover the sheet index from current document if it is a Calc document.
                A value of ``-2`` means no attempt is made to discover the sheet index.
                Default is ``-2``.
            kwargs: Additional arguments to pass to the event data.

        Returns:
            RangeObj: Range object.
        """
        return RangeObj(
            col_start=col_start,
            col_end=col_end,
            row_start=row_start,
            row_end=row_end,
            sheet_idx=sheet_idx,
        )

    def rng_from_cell_obj(self, cell_obj: CellObj) -> RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            RangeObj: Range object.
        """

        return self._create_range_obj(
            col_start=cell_obj.col,
            col_end=cell_obj.col,
            row_start=cell_obj.row,
            row_end=cell_obj.row,
            sheet_idx=cell_obj.sheet_idx,
        )

    def rng_from_position(
        self, col_start: int, row_start: int, col_end: int, row_end: int, sheet_idx: int = -2
    ) -> RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            col_start (int): Zero-based start column index.
            row_start (int): Zero-based start row index.
            col_end (int): Zero-based end column index.
            row_end (int): Zero-based end row index.
            sheet_idx (int, optional): Zero-based sheet index that this range value belongs to. Default is -1.

        Returns:
            RangeObj: Range object.
        """
        col_start_str = TableHelper.make_column_name(col=col_start, zero_index=True)
        col_end_str = TableHelper.make_column_name(col=col_end, zero_index=True)
        return self._create_range_obj(
            col_start=col_start_str,
            col_end=col_end_str,
            row_start=row_start + 1,
            row_end=row_end + 1,
            sheet_idx=sheet_idx,
        )

    def rng_from_cell_rng_addr(self, addr: CellRangeAddress) -> RangeObj:
        """
        Gets a range Object representing a range from a cell range address.

        Args:
            addr (CellRangeAddress): Cell Range Address.

        Returns:
            RangeObj: Range object.
        """
        col_start = TableHelper.make_column_name(addr.StartColumn, True)
        col_end = TableHelper.make_column_name(addr.EndColumn, True)
        row_start = addr.StartRow + 1
        row_end = addr.EndRow + 1
        sheet_idx = addr.Sheet
        return self._create_range_obj(
            col_start=col_start,
            col_end=col_end,
            row_start=row_start,
            row_end=row_end,
            sheet_idx=sheet_idx,
        )

    def rng_from_cell_rng_value(self, rng: RangeValues) -> RangeObj:
        """
        Gets a range Object representing a range from a cell range address.

        Args:
            rng (RangeValues): Cell Range Values.

        Returns:
            RangeObj: Range object.

        Hint:
            - ``RangeValues`` can be imported from ``ooodev.utils.data_type.range_values``
        """
        col_start = TableHelper.make_column_name(rng.col_start, True)
        col_end = TableHelper.make_column_name(rng.col_end, True)
        row_start = rng.row_start + 1
        row_end = rng.row_end + 1
        sheet_idx = rng.sheet_idx
        return self._create_range_obj(
            col_start=col_start,
            col_end=col_end,
            row_start=row_start,
            row_end=row_end,
            sheet_idx=sheet_idx,
        )

    def rng_from_cell_rng(self, cell_range: XCellRange) -> RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            RangeObj: Range object.
        """
        addr = self.get_cell_address_from_cell_range(cell_range)
        return self.rng_from_cell_rng_addr(addr)

    def rng_from_str(self, rng_name: str) -> RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            rng_name (str): Range as string such as ``Sheet1.A1:C125`` or ``A1:C125``.
                Cell name is also valid such as ``A1``.

        Returns:
            RangeObj: Range object.
        """
        if self.is_cell_range_name(rng_name):
            parts = TableHelper.get_range_parts(rng_name)
            col_start = parts.col_start
            col_end = parts.col_end
            row_start = parts.row_start
            row_end = parts.row_end
        else:
            parts = TableHelper.get_cell_parts(rng_name)
            col_start = parts.col
            col_end = parts.col
            row_start = parts.row
            row_end = parts.row
        idx = self.get_sheet_index(parts.sheet) if parts.sheet else -2
        return self._create_range_obj(
            col_start=col_start,
            col_end=col_end,
            row_start=row_start,
            row_end=row_end,
            sheet_idx=idx,
        )

    # region get_range_obj() method

    @overload
    def get_range_obj(self, cell_obj: CellObj) -> RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def get_range_obj(self, range_obj: RangeObj) -> RangeObj:
        """
        Gets a range object. Returns the same object.

        Args:
            range_obj (RangeObj): Range Object

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def get_range_obj(
        self, col_start: int, row_start: int, col_end: int, row_end: int, sheet_idx: int = ...
    ) -> RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            col_start (int): Zero-based start column index.
            row_start (int): Zero-based start row index.
            col_end (int): Zero-based end column index.
            row_end (int): Zero-based end row index.
            sheet_idx (int, optional): Zero-based sheet index that this range value belongs to. Default is -1.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def get_range_obj(self, addr: CellRangeAddress) -> RangeObj:
        """
        Gets a range Object representing a range from a cell range address.

        Args:
            addr (CellRangeAddress): Cell Range Address.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def get_range_obj(self, rng: RangeValues) -> RangeObj:
        """
        Gets a range Object representing a range from a cell range address.

        Args:
            rng (RangeValues): Cell Range Values.

        Returns:
            RangeObj: Range object.

        Hint:
            - ``RangeValues`` can be imported from ``ooodev.utils.data_type.range_values``
        """
        ...

    @overload
    def get_range_obj(self, cell_range: XCellRange) -> RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def get_range_obj(self, rng_name: str) -> RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            rng_name (str): Range as string such as ``Sheet1.A1:C125`` or ``A1:C125``

        Returns:
            RangeObj: Range object.
        """
        ...

    def get_range_obj(self, *args, **kwargs) -> RangeObj:
        """
        Gets the cell as string from a cell.

        Args:
            range_obj (RangeObj): Range Object.
            col_start (int): Zero-based start column index.
            row_start (int): Zero-based start row index.
            col_end (int): Zero-based end column index.
            row_end (int): Zero-based end row index.
            sheet_idx (int, optional): Zero-based sheet index that this range value belongs to. Default is ``-2``.
            addr (CellRangeAddress): Cell Range Address.
            rng (RangeValues): Cell Range Values.
            cell_range (XCellRange): Cell Range.
            rng_name (str): Range as string such as ``Sheet1.A1:C125`` or ``A1:C125``



        Returns:
            RangeObj: Range Object.
        """
        # rng_from_cell_obj(self, cell_obj: CellObj)
        # rng_from_position(self, col_start: int, row_start: int, col_end: int, row_end: int, sheet_idx: int = -1)
        # rng_from_cell_rng_addr(self, addr: CellRangeAddress)
        # rng_from_cell_rng_value(self, rng: RangeValues)
        # rng_from_cell_rng(self, cell_range: XCellRange)
        # rng_from_str(self, rng: str)

        ordered_keys = (1, 2, 3, 4, 5)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = {
                "range_obj",
                "cell_obj",
                "addr",
                "rng",
                "cell_range",
                "rng_name",
                "row_start",
                "col_start",
                "col_end",
                "row_end",
                "sheet_idx",
            }
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_range_obj() got an unexpected keyword argument")
            keys = valid_keys = {"range_obj", "cell_obj", "col_start", "addr", "rng", "cell_range", "rng_name"}
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            if "row_start" in kwargs:
                ka[2] = kwargs["row_start"]
            if "col_end" in kwargs:
                ka[3] = kwargs["col_end"]
            if "row_end" in kwargs:
                ka[4] = kwargs["row_end"]
            if count == 4:
                return ka
            if "sheet_idx" in kwargs:
                ka[5] = kwargs["sheet_idx"]
            return ka

        if count not in (1, 4, 5):
            raise TypeError("get_range_obj() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 4:
            return self.rng_from_position(kargs[1], kargs[2], kargs[3], kargs[4], -2)
        if count == 5:
            return self.rng_from_position(kargs[1], kargs[2], kargs[3], kargs[4], kargs[5])
        arg1 = kargs[1]
        if mInfo.Info.is_instance(arg1, RangeObj):
            return arg1
        if hasattr(arg1, "typeName") and getattr(arg1, "typeName") == "com.sun.star.table.CellRangeAddress":
            return self.rng_from_cell_rng_addr(arg1)
        if mInfo.Info.is_instance(arg1, CellObj):
            return self.rng_from_cell_obj(arg1)
        if mInfo.Info.is_instance(arg1, RangeValues):
            return self.rng_from_cell_rng_value(arg1)
        if mInfo.Info.is_instance(arg1, str):
            return self.rng_from_str(arg1)
        return self.rng_from_cell_rng(arg1)

    # endregion get_range_obj() method

    def get_range_from_2d(self, data: Sequence[Sequence[Any]]) -> RangeObj:
        """
        Creates a range object from a 2D array of data.

        Args:
            data (Sequence[Sequence[Any]]): 2D array of data.

        Returns:
            RangeObj: Range object.
        """
        row_count = len(data)
        col_count = len(data[0])
        return self.rng_from_position(0, 0, col_count - 1, row_count - 1)

    # endregion Range Converters to RangeObj

    def get_offset_range_obj(self, range_obj: RangeObj) -> RangeObj:
        """
        Gets a new range object with an offset from the original range object.

        The returned range will have Start column of ``A``, Start row of ``1``.

        Args:
            range_obj (RangeObj): Range Object.
            col_offset (int): Column offset.
            row_offset (int): Row offset.

        Returns:
            RangeObj: Range Object.

        Note:
            A ``RangeConverter.EVENT_RANGE_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col_start: Column start such as ``A``
            - col_end: Column end such as ``C``
            - row_start: Row start such as ``1``
            - row_end: Row end such as ``125``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to. Default is -1.
        """
        row_count = range_obj.row_count  # current_rng.row_end - current_rng.row_start
        col_count = range_obj.col_count
        sheet_idx = range_obj.sheet_idx
        return self._create_range_obj(
            col_start="A",
            col_end=TableHelper.make_column_name(col_count),
            row_start=1,
            row_end=row_count,
            sheet_idx=sheet_idx,
        )

    def get_sheet_index(self, key: str | int = -1) -> int:
        """
        Gets the sheet index from the current Calc document.

        Args:
            key (str | int, optional): Sheet name or Sheet index.
                A value of ``-1`` means get active sheet index. Defaults to ``-1``

        Returns:
            int: Sheet index or ``-1`` if not found. If current doc is not a Calc doc then ``-2`` is returned.
        """
        idx = -1
        with contextlib.suppress(Exception):
            # pylint: disable=no-member
            if mLo.Lo.is_loaded and mLo.Lo.current_doc.DOC_TYPE == DocType.CALC:
                doc = cast("CalcDoc", mLo.Lo.current_doc)
                if isinstance(key, str):
                    sheet = doc.get_sheet(sheet_name=key)
                else:
                    sheet = doc.get_active_sheet() if key < 0 else doc.sheets[key]
                idx = sheet.get_sheet_index()
            else:
                idx = -2
        return idx

    def is_cell_range_name(self, s: str) -> bool:
        """
        Gets if is a cell name or a cell range.

        Args:
            s (str): cell name such as 'A1' or range name such as 'B3:E7'

        Returns:
            bool: True if range name; Otherwise, False
        """
        return ":" in s

    def is_single_cell_range(self, cr_addr: CellRangeAddress) -> bool:
        """
        Gets if a cell address is a single cell or a range.

        Args:
            cr_addr (CellRangeAddress): cell range address

        Returns:
            bool: ``True`` if single cell; Otherwise, ``False``
        """
        return cr_addr.StartColumn == cr_addr.EndColumn and cr_addr.StartRow == cr_addr.EndRow

    def is_single_column_range(self, cr_addr: CellRangeAddress) -> bool:
        """
        Gets if a cell address is a single column or multi-column.

        Args:
            cr_addr (CellRangeAddress): cell range address

        Returns:
            bool: ``True`` if single column; Otherwise, ``False``

        Note:
            If ``cr_addr`` is a single cell address then ``True`` is returned.
        """
        return cr_addr.StartColumn == cr_addr.EndColumn

    def is_single_row_range(self, cr_addr: CellRangeAddress) -> bool:
        """
        Gets if a cell address is a single row or multi-row.

        Args:
            cr_addr (CellRangeAddress): cell range address

        Returns:
            bool: ``True`` if single row; Otherwise, ``False``

        Note:
            If ``cr_addr`` is a single cell address then ``True`` is returned.
        """
        return cr_addr.StartRow == cr_addr.EndRow
