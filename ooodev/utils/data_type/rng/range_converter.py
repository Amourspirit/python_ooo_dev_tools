from __future__ import annotations
from typing import overload, TYPE_CHECKING, Tuple
import uno
from com.sun.star.container import XNamed
from com.sun.star.sheet import XCellAddressable
from com.sun.star.sheet import XCellRangeAddressable

from ooodev.events.args.event_args import EventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.data_type.cell_obj import CellObj
from ooodev.utils.data_type.cell_values import CellValues
from ooodev.utils.data_type.range_obj import RangeObj
from ooodev.utils.data_type.range_values import RangeValues
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.table_helper import TableHelper

if TYPE_CHECKING:
    from com.sun.star.table import CellRangeAddress
    from com.sun.star.table import XCellRange
    from com.sun.star.sheet import XSpreadsheet
    from com.sun.star.table import CellAddress
    from com.sun.star.table import XCell


class RangeConverter(LoInstPropsPartial, EventsPartial):
    EVENT_RANGE_CREATING = "range_converter_range_creating"
    EVENT_CELL_CREATING = "range_converter_cell_creating"

    def __init__(self, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        EventsPartial.__init__(self)

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
    def _create_cell_obj(self, col: str, row: int, sheet_idx: int = -1, **kwargs) -> CellObj:
        """
        Creates a cell object

        Args:
            col (str): Column such as ``A``
            row (int): Row such as ``1``
            sheet_idx (int, optional): Sheet index that this cell value belongs to.
            kwargs: Additional arguments to pass to the event data.

        Returns:
            CellObj: Cell object.

        Note:
            By Default when a cell object is created it will check for a sheet index when the ``sheet_idx`` is less then 0.
            In this case we do not want this. This converter may be used in other places such as in a Write Table.
            For this reason an event is triggered that allows for a sheet index to be set.
            For the purposes of the method, the sheet index is set to -1. By Default.
            If the sheet index is less than 0, the sheet index will be set to 0 at creation time to avoid a check for a spreadsheet index.
            After the cell object is created, the sheet index will be set to -1.
        """
        args = EventArgs(source=self)
        event_data = {
            "col": col,
            "row": row,
            "sheet_idx": sheet_idx,
        }
        if kwargs:
            event_data.update(kwargs)
        args.event_data = event_data
        self.trigger_event(RangeConverter.EVENT_CELL_CREATING, args)
        col = args.event_data.get("col", col)
        row = args.event_data.get("row", row)
        sheet_idx = args.event_data.get("sheet_idx", sheet_idx)

        idx = max(sheet_idx, 0)

        cell_obj = CellObj(
            col=col,
            row=row,
            sheet_idx=idx,
        )
        if sheet_idx < 0:
            object.__setattr__(cell_obj, "sheet_idx", -1)
        return cell_obj

    def get_cell_obj_from_col_row(self, col: int, row: int, sheet_idx: int = -1) -> CellObj:
        """
        Gets the cell as CellObj from column and row.

        Args:
            col (int): Column. Zero Based column index.
            row (int): Row. Zero Based row index.
            sheet_idx (int, optional): Sheet index that this cell value belongs to. Default is ``-1``.

        Returns:
            CellObj: Cell Object.

        Note:
            A ``RangeConverter.EVENT_CELL_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col: Column start such as ``A``
            - row: Row start such as ``1``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to. Default is the value of ``sheet_idx``.
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

        Note:
            A ``RangeConverter.EVENT_CELL_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col: Column start such as ``A``
            - row: Row start such as ``1``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to. Default is the value of ``addr.Sheet``.
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

        Note:
            A ``RangeConverter.EVENT_CELL_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col: Column start such as ``A``
            - row: Row start such as ``1``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to.
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

        Note:
            A ``RangeConverter.EVENT_CELL_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col: Column start such as ``A``
            - row: Row start such as ``1``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to. Default is the value of ``val.sheet_idx``.

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

        Note:
            A ``RangeConverter.EVENT_CELL_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col: Column start such as ``A``
            - row: Row start such as ``1``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to.
        """
        col = TableHelper.make_cell_name(values[0], True)
        row = values[1] + 1
        sheet_idx = values[2] if len(values) == 3 else -1
        return self._create_cell_obj(col=col, row=row, sheet_idx=sheet_idx)

    def get_cell_obj_from_str(self, name: str) -> CellObj:
        """
        Gets the cell as CellObj from a cell name.

        Args:
            name (str): Cell name such as as ``A23`` or ``Sheet1.A23``

        Returns:
            CellObj: Cell Object.

        Note:
            If a range name such as ``A23:G45`` or ``Sheet1.A23:G45`` then only the first cell is used.

            A ``RangeConverter.EVENT_CELL_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col: Column start such as ``A``
            - row: Row start such as ``23``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to.
            - sheet_name: Sheet name if applicable. May be empty string.

            If a sheet name is present it is passed to the event data in the ``sheet_name`` key.
            If the sheet name is not present ``sheet_name`` key will be an empty string.

            If there is a sheet name is will not be converted into a sheet index.
            This must be done manually by setting the ``sheet_idx`` key in the event data.
        """
        parts = TableHelper.get_cell_parts(name)
        return self._create_cell_obj(col=parts.col, row=parts.row, sheet_idx=-1, sheet_name=parts.sheet)

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
    def get_cell_obj(self, col: int, row: int, sheet_idx: int = -1) -> CellObj:
        """
        Gets the cell as CellObj from column and row.

        Args:
            col (int): Column. Zero Based column index.
            row (int): Row. Zero Based row index.
            sheet_idx (int, optional): Sheet index that this cell value belongs to. Default is ``-1``.

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
            name (cell_obj): Cell Object

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
            sheet_idx (int, optional): Sheet index that this cell value belongs to. Default is ``-1``.
            addr (CellAddress): Cell Address.
            cell (XCell): Cell.
            val (CellValues): Cell values.
            name (str): Cell name such as as ``A23`` or ``Sheet1.A23``


        Returns:
            CellObj: Cell Object.

        Note:
            If a range name such as ``A23:G45`` or ``Sheet1.A23:G45`` then only the first cell is used.

            A ``RangeConverter.EVENT_CELL_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col: Column start such as ``A``
            - row: Row start such as ``23``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to.
            - sheet_name: Sheet name if applicable. May be empty string. Key may not be present.
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
            valid_keys = {"col", "addr", "cell", "val", "name", "values", "cell_obj"}
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_cell_obj() got an unexpected keyword argument")
            for key in valid_keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            if "row" in kwargs:
                ka[2] = kwargs["row"]
            if "sheet_idx" in kwargs:
                ka[3] = kwargs["sheet_idx"]
            return ka

        if count not in (1, 3):
            raise TypeError("get_cell_obj() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 3:
            return self.get_cell_obj_from_col_row(kargs[1], kargs[2], kargs[3])
        arg1 = kargs[1]
        if hasattr(arg1, "typeName") and getattr(arg1, "typeName") == "com.sun.star.table.CellAddress":
            return self.get_cell_obj_from_addr(arg1)
        if mInfo.Info.is_instance(arg1, CellObj):
            return arg1
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
        self, col_start: str, col_end: str, row_start: int, row_end: int, sheet_idx: int = -1, **kwargs
    ) -> RangeObj:
        """
        Creates a cell range object

        Args:
            col_start (str): Column start such as ``A``
            col_end (str): Column end such as ``C``
            row_start (int): Row start such as ``1``
            row_end (int): Row end such as ``125``
            sheet_idx (int, optional): Sheet index that this range value belongs to.
            kwargs: Additional arguments to pass to the event data.

        Returns:
            RangeObj: Range object.

        Note:
            By Default when a range object is created it will check for a sheet index when the ``sheet_idx`` is less then 0.
            In this case we do not want this. This converter may be used in other places such as in a Write Table.
            For this reason an event is triggered that allows for a sheet index to be set.
            For the purposes of the method, the sheet index is set to -1. By Default.
            If the sheet index is less than 0, the sheet index will be set to 0 at creation time to avoid a check for a spreadsheet index.
            After the range object is created, the sheet index will be set to -1.
        """
        args = EventArgs(source=self)
        event_data = {
            "col_start": col_start,
            "col_end": col_end,
            "row_start": row_start,
            "row_end": row_end,
            "sheet_idx": sheet_idx,
        }
        if kwargs:
            event_data.update(kwargs)
        args.event_data = event_data
        self.trigger_event(RangeConverter.EVENT_RANGE_CREATING, args)
        col_start = args.event_data.get("col_start", col_start)
        col_end = args.event_data.get("col_end", col_end)
        row_start = args.event_data.get("row_start", row_start)
        row_end = args.event_data.get("row_end", row_end)
        sheet_idx = args.event_data.get("sheet_idx", sheet_idx)

        idx = max(sheet_idx, 0)

        rng_obj = RangeObj(
            col_start=col_start,
            col_end=col_end,
            row_start=row_start,
            row_end=row_end,
            sheet_idx=idx,
        )
        if sheet_idx < 0:
            object.__setattr__(rng_obj, "sheet_idx", -1)
        return rng_obj

    def rng_from_cell_obj(self, cell_obj: CellObj) -> RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            RangeObj: Range object.

        Note:
            A ``RangeConverter.EVENT_RANGE_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col_start: Column start such as ``A``
            - col_end: Column end such as ``C``
            - row_start: Row start such as ``1``
            - row_end: Row end such as ``125``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to. Default is the value of ``cell_obj.sheet_idx``.
        """

        return self._create_range_obj(
            col_start=cell_obj.col,
            col_end=cell_obj.col,
            row_start=cell_obj.row,
            row_end=cell_obj.row,
            sheet_idx=cell_obj.sheet_idx,
        )

    def rng_from_position(
        self, col_start: int, row_start: int, col_end: int, row_end: int, sheet_idx: int = -1
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

        Note:
            A ``RangeConverter.EVENT_RANGE_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col_start: Column start such as ``A``
            - col_end: Column end such as ``C``
            - row_start: Row start such as ``1``
            - row_end: Row end such as ``125``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to. Default is -1.

            By default ``sheet_idx`` will be ``-1`` meaning no sheet index is set.
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

        Note:
            A ``RangeConverter.EVENT_RANGE_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col_start: Column start such as ``A``
            - col_end: Column end such as ``C``
            - row_start: Row start such as ``1``
            - row_end: Row end such as ``125``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to.
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
            addr (RangeValues): Cell Range Values.

        Returns:
            RangeObj: Range object.

        Note:
            A ``RangeConverter.EVENT_RANGE_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col_start: Column start such as ``A``
            - col_end: Column end such as ``C``
            - row_start: Row start such as ``1``
            - row_end: Row end such as ``125``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to.
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

        Note:
            A ``RangeConverter.EVENT_RANGE_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col_start: Column start such as ``A``
            - col_end: Column end such as ``C``
            - row_start: Row start such as ``1``
            - row_end: Row end such as ``125``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to.
        """
        addr = self.get_cell_address_from_cell_range(cell_range)
        return self.rng_from_cell_rng_addr(addr)

    def rng_from_str(self, rng: str) -> RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            rng (str): Range as string such as ``Sheet1.A1:C125`` or ``A1:C125``

        Returns:
            RangeObj: Range object.

        Note:
            A ``RangeConverter.EVENT_RANGE_CREATING`` event is triggered that allows for a sheet index to be set and any other range object args to be set.
            The ``EventArgs.event_data`` is a dictionary and contains the following keys:

            - col_start: Column start such as ``A``
            - col_end: Column end such as ``C``
            - row_start: Row start such as ``1``
            - row_end: Row end such as ``125``
            - sheet_idx: Sheet index, if applicable, that this range value belongs to. Default is -1.
            - sheet_name: Sheet name if applicable. May be empty string.

            If a sheet name is present it is passed to the event data in the ``sheet_name`` key.
            If the sheet name is not present ``sheet_name`` key will be an empty string.

            If there is a sheet name is will not be converted into a sheet index.
            This must be done manually by setting the ``sheet_idx`` key in the event data.
        """
        parts = TableHelper.get_range_parts(rng)
        col_start = parts.col_start
        col_end = parts.col_end
        row_start = parts.row_start
        row_end = parts.row_end
        sheet_name = parts.sheet
        return self._create_range_obj(
            col_start=col_start,
            col_end=col_end,
            row_start=row_start,
            row_end=row_end,
            sheet_name=sheet_name,
        )

    # region get_range_obj() method

    def get_range_obj(self, *args, **kwargs) -> RangeObj:
        """
        Gets the cell as string from a cell.

        Args:
            col (int): Column. Zero Based column index.

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
            - sheet_name: Sheet name if applicable. May be empty string. Key may not be present.
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
            valid_keys = {"cell_obj", "col_start", "addr", "rng", "cell_range"}
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_range_obj() got an unexpected keyword argument")
            for key in valid_keys:
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
            if "sheet_idx" in kwargs:
                ka[5] = kwargs["sheet_idx"]
            return ka

        if count not in (1, 5):
            raise TypeError("get_range_obj() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 5:
            return self.rng_from_position(kargs[1], kargs[2], kargs[3], kargs[4], kargs[5])
        arg1 = kargs[1]
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
    # endregion Range Converters to RangeObj
