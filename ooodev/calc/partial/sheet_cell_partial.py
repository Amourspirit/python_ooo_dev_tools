from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno

from com.sun.star.table import XCell
from com.sun.star.table import XCellRange

from ooodev.office import calc as mCalc
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.data_type import cell_obj as mCellObj
from ooodev.loader.inst.lo_inst import LoInst

from ooodev.calc import calc_cell as mCalcCell
from ooodev.calc import calc_sheet as mCalcSheet

if TYPE_CHECKING:
    from com.sun.star.table import CellAddress


class SheetCellPartial:
    """Partial Class for sheet Cell"""

    def __init__(self, owner: mCalcSheet.CalcSheet, lo_inst: LoInst | None = None) -> None:
        self.__owner = owner
        self.__lo_inst = mLo.Lo.current_lo if lo_inst is None else lo_inst

    def __getitem__(self, _val: Any) -> mCalcCell.CalcCell:
        # print(f"Getting item at index {index}")
        if isinstance(_val, tuple):
            return self.get_cell(int(_val[0]), int(_val[1]))
        return self.get_cell(_val)

    # region get_cell()
    @overload
    def get_cell(self, cell: XCell) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            cell (XCell): Cell

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, addr: CellAddress) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            addr (CellAddress): Cell Address

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, cell_name: str) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            cell_name (str): Cell Name such as 'A1'

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, cell_obj: mCellObj.CellObj) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            cell_obj: (CellObj): Cell object

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, col: int, row: int) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            col (int): Cell column
            row (int): cell row

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, cell_range: XCellRange) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            cell_range (XCellRange): Cell Range

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, cell_range: XCellRange, col: int, row: int) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            cell_range (XCellRange): Cell Range
            col (int): Cell column
            row (int): cell row

        Returns:
            CalcCell: cell
        """
        ...

    def get_cell(self, *args, **kwargs) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            addr (CellAddress): Cell Address
            cell_name (str): Cell Name such as 'A1'
            cell_obj: (CellObj): Cell object
            cell_range (XCellRange): Cell Range
            col (int): Cell column
            row (int): cell row
            cell (XCell): Cell

        Returns:
            CalcCell: cell
        """
        # pylint: disable=protected-access
        # valid overloads
        # def get_cell(self, cell: XCell)
        # def get_cell(self, addr: CellAddress)
        # def get_cell(self, cell_name: str)
        # def get_cell(self, cell_obj: CellObj)
        # def get_cell(self, cell_range: XCellRange)
        # def get_cell(self, col: int, row: int)
        # def get_cell(self, cell_range: XCellRange, col: int, row: int)

        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell", "addr", "cell_name", "cell_obj", "cell_range", "col", "row")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_cell() got an unexpected keyword argument")
            keys = ("cell", "addr", "cell_name", "cell_obj", "cell_range", "col")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("row", "col")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("row", None)
            return ka

        if count not in (1, 2, 3):
            raise TypeError("get_cell() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        arg1 = kargs[1]
        if count == 1:
            # def get_cell(self, cell: XCell)
            # def get_cell(self, cell_range: XCellRange)

            # def get_cell(self, addr: CellAddress)
            # def get_cell(self, cell_name: str)
            # def get_cell(self, cell_obj: CellObj)
            if self.__lo_inst.is_uno_interfaces(arg1, XCell):
                x_cell = arg1
            elif self.__lo_inst.is_uno_interfaces(arg1, XCellRange):
                with LoContext(self.__lo_inst):
                    x_cell = mCalc.Calc.get_cell(cell_range=arg1)
            else:
                with LoContext(self.__lo_inst):
                    x_cell = mCalc.Calc.get_cell(self.__owner.component, arg1)
        elif count == 2:
            # def get_cell(self, col: int, row: int)
            with LoContext(self.__lo_inst):
                x_cell = mCalc.Calc._get_cell_sheet_col_row(sheet=self.__owner.component, col=kargs[1], row=kargs[2])
        else:
            # def get_cell(self, cell_range: XCellRange, col: int, row: int)
            with LoContext(self.__lo_inst):
                x_cell = mCalc.Calc._get_cell_cell_rng(cell_range=kargs[1], col=kargs[2], row=kargs[3])
        with LoContext(self.__lo_inst):
            cell_obj = mCalc.Calc.get_cell_obj(cell=x_cell)
        return mCalcCell.CalcCell(owner=self.__owner, cell=cell_obj)

    # endregion get_cell()
