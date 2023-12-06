from __future__ import annotations
from typing import Any, Sequence, TYPE_CHECKING
import uno

if TYPE_CHECKING:
    from com.sun.star.awt import Point
    from com.sun.star.sheet import XSheetAnnotation
    from com.sun.star.sheet import XGoalSeek
    from .calc_sheet import CalcSheet
    from . import calc_cell_cursor as mCalcCellCursor
else:
    XSheetAnnotation = object

from ooodev.exceptions import ex as mEx
from ooodev.proto.style_obj import StyleT
from ooodev.units import UnitT
from ooodev.utils.data_type import cell_obj as mCellObj
from ooodev.utils.type_var import Row, Table
from ooodev.office import calc as mCalc
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils import lo as mLo
from ooodev.adapter.sheet.sheet_cell_comp import SheetCellComp


class CalcCell(SheetCellComp, QiPartial, PropPartial):
    def __init__(self, owner: CalcSheet, cell: str | mCellObj.CellObj) -> None:
        self.__owner = owner
        self.__cell_obj = mCellObj.CellObj.from_cell(cell)
        # don't use owner.get_cell() here because it will be recursive.
        sheet_cell = mCalc.Calc.get_cell(sheet=owner.component, cell_obj=self.__cell_obj)
        SheetCellComp.__init__(self, sheet_cell)  # type: ignore
        QiPartial.__init__(self, component=sheet_cell, lo_inst=mLo.Lo.current_lo)
        PropPartial.__init__(self, component=sheet_cell, lo_inst=mLo.Lo.current_lo)

    def create_cursor(self) -> mCalcCellCursor.CalcCellCursor:
        """
        Creates a cell cursor to travel in the given range context.

        Returns:
            CalcCellCursor: Cell cursor
        """
        return self.calc_sheet.create_cursor_by_range(cell_obj=self.__cell_obj)

    # region Cell Properties

    def is_first_row(self) -> bool:
        """Determines if this cell is in the first row of the sheet."""
        return self.__cell_obj.row == 1

    def is_first_column(self) -> bool:
        """Determines if this cell is in the first column of the sheet."""
        return self.__cell_obj.col == "A"

    # endregion Cell Properties

    def goal_seek(
        self,
        gs: XGoalSeek,
        formula_cell_name: str | mCellObj.CellObj,
        result: int | float,
    ) -> float:
        """
        Calculates a value which gives a specified result in a formula.

        Args:
            gs (XGoalSeek): Goal seeking value for cell
            formula_cell_name (str | CellObj): formula cell name
            result (int, float): float or int, result of the goal seek

        Raises:
            GoalDivergenceError: If goal divergence is greater than 0.1

        Returns:
            float: result of the goal seek
        """
        return mCalc.Calc.goal_seek(
            gs=gs,
            sheet=self.calc_sheet.component,
            cell_name=self.cell_obj,
            formula_cell_name=formula_cell_name,
            result=result,
        )

    def get_cell_position(self) -> Point:
        """
        Gets a cell name as a Point.

        - ``Point.X`` is column zero-based index.
        - ``Point.Y`` is row zero-based index.

        Returns:
            Point: cell name as Point with X as col and Y as row
        """
        return mCalc.Calc.get_cell_position(cell_name=self.__cell_obj)

    # region Other Cells

    def get_cell_left(self) -> CalcCell:
        """
        Gets the cell to the left of this cell.

        Raises:
            CellError: If cell is in the first column of the sheet.

        Returns:
            CalcCell: Cell to the left of this cell.
        """
        if self.is_first_column():
            raise mEx.CellError(f"Cell {self.__cell_obj} is in the first column of the sheet.")
        cell_obj = self.__cell_obj.left
        return CalcCell(owner=self.calc_sheet, cell=cell_obj)

    def get_cell_right(self) -> CalcCell:
        """
        Gets the cell to the right of this cell.

        Raises:
            CellError: If cell is in the last column of the sheet.

        Returns:
            CalcCell: Cell to the right of this cell.
        """
        cell_obj = self.__cell_obj.right
        return CalcCell(owner=self.calc_sheet, cell=cell_obj)

    def get_cell_up(self) -> CalcCell:
        """
        Gets the cell above this cell.

        Returns:
            CalcCell: Cell above this cell.
        """
        if self.is_first_row():
            raise mEx.CellError(f"Cell {self.__cell_obj} is in the first row of the sheet.")
        cell_obj = self.__cell_obj.up
        return CalcCell(owner=self.calc_sheet, cell=cell_obj)

    def get_cell_down(self) -> CalcCell:
        """
        Gets the cell below this cell.

        Returns:
            CalcCell: Cell below this cell.
        """
        cell_obj = self.__cell_obj.down
        return CalcCell(owner=self.calc_sheet, cell=cell_obj)

    # endregion Other Cells

    def add_annotation(self, msg: str, is_visible=True) -> XSheetAnnotation:
        """
        Adds an annotation to a cell and makes the annotation visible.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Name of cell to add annotation such as 'A1'
            msg (str): Annotation Text
            set_visible (bool): Determines if the annotation is set visible

        Raises:
            MissingInterfaceError: If interface is missing

        Returns:
            XSheetAnnotation: Cell annotation that was added
        """
        return mCalc.Calc.add_annotation(
            sheet=self.calc_sheet.component, cell_name=str(self.__cell_obj), msg=msg, is_visible=is_visible
        )

    def get_annotation(self) -> XSheetAnnotation:
        """
        Gets the annotation of a cell.

        Returns:
            XSheetAnnotation: Cell annotation
        """
        return mCalc.Calc.get_annotation(sheet=self.calc_sheet.component, cell_name=self.__cell_obj)

    def get_annotation_str(self) -> str:
        """
        Gets text of an annotation for a cell.

        Returns:
            str: Cell annotation text
        """
        return mCalc.Calc.get_annotation_str(sheet=self.calc_sheet.component, cell_name=self.__cell_obj)

    def get_string(self) -> str:
        """
        Gets the value of a cell as a string.

        Returns:
            str: Cell value as string.
        """
        return mCalc.Calc.get_string(cell=self.component)

    def get_cell_str(self) -> str:
        """
        Gets the cell as a string in format of ``A1``

        Returns:
            str: Cell as str
        """
        return str(self.__cell_obj)

    def get_type_enum(self) -> mCalc.Calc.CellTypeEnum:
        """
        Gets enum representing the Type

        Returns:
            CellTypeEnum: Enum of cell type
        """
        return mCalc.Calc.get_type_enum(cell=self.component)

    def get_type_string(self) -> str:
        """
        Gets String representing the Type

        Returns:
            str: String of cell type
        """
        return str(self.get_type_enum())

    def get_num(self) -> float:
        """
        Get cell value a float

        Returns:
            float: Cell value as float. If cell value cannot be converted then 0.0 is returned.
        """
        return mCalc.Calc.get_num(cell=self.component)

    def get_val(self) -> Any | None:
        """
        Gets cell value

        Returns:
            Any | None: Cell value cell has a value; Otherwise, None
        """
        return mCalc.Calc.get_val(cell=self.component)

    # region set_array_cell()
    def set_array_cell(self, values: Table, styles: Sequence[StyleT] | None = None) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            range_name (str): Range to insert data such as 'A1:E12'.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.
        """
        if styles is None:
            self.calc_sheet.set_array_cell(cell_obj=self.__cell_obj, values=values)
        else:
            self.calc_sheet.set_array_cell(cell_obj=self.__cell_obj, values=values, styles=styles)

    # endregion set_array_cell()
    def set_date(self, day: int, month: int, year: int) -> None:
        """
        Writes a date with standard date format into a spreadsheet

        Args:
            day (int): Date day part.
            month (int): Date month part.
            year (int): Date year part.
        """
        mCalc.Calc.set_date(
            sheet=self.calc_sheet.component, cell_name=self.__cell_obj, day=day, month=month, year=year
        )

    # region set_row()
    def set_row(self, values: Row) -> None:
        """
        Inserts a row of data into spreadsheet

        Raises:
            MissingInterfaceError: if unable to obtain interface

        Args:
            values (Row): Row Data.
        """
        self.calc_sheet.set_row(
            values=values,
            cell_obj=self.__cell_obj,
        )

    # endregion set_row()

    def set_row_height(
        self,
        height: int | UnitT,
    ) -> None:
        """
        Sets column width. height is in ``mm``, e.g. 6

        Args:
            height (int, UnitT): Width in ``mm`` units or :ref:`proto_unit_obj`.
            idx (int): Index of Row

        Raises:
            CancelEventError: If SHEET_ROW_HEIGHT_SETTING event is canceled.

        Returns:
            None:

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_HEIGHT_SETTING` :eventref:`src-docs-sheet-event-row-height-setting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_HEIGHT_SET` :eventref:`src-docs-sheet-event-row-height-set`

        Note:
            Event args ``index`` is set to ``idx`` value, ``event_data`` is set to ``height`` value (``mm100`` units).
        """
        index = self.__cell_obj.row - 1
        self.calc_sheet.set_row_height(height=height, idx=index)

    def set_style(self, styles: Sequence[StyleT]) -> None:
        """
        Sets style for cell

        Args:
            styles (Sequence[StyleT]): One or more styles to apply to cell.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        mCalc.Calc.set_style_cell(sheet=self.calc_sheet.component, cell_obj=self.__cell_obj, styles=styles)

    def set_val(self, value: Any, styles: Sequence[StyleT] | None = None) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell.
            cell_obj (CellObj): Cell Object.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:
        """
        if styles is None:
            self.calc_sheet.set_val(value=value, cell_obj=self.__cell_obj)
        else:
            self.calc_sheet.set_val(value=value, cell_obj=self.__cell_obj, styles=styles)

    def split_window(self) -> None:
        """
        Splits window

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            cell_name (str): Cell to preform split on. e.g. 'C4'

        Returns:
            None:

        See Also:
            :ref:`ch23_splitting_panes`
        """
        mCalc.Calc.split_window(doc=self.calc_sheet.calc_doc.component, cell_name=str(self.__cell_obj))

    # region Properties
    @property
    def calc_sheet(self) -> CalcSheet:
        """Sheet that owns this cell."""
        return self.__owner

    @property
    def cell_obj(self) -> mCellObj.CellObj:
        """Cell object."""
        return self.__cell_obj

    @property
    def position(self) -> Point:
        """
        Contains the position of the top left cell of this range in the sheet (in 1/100 mm).

        This property contains the absolute position in the whole sheet,
        not the position in the visible area.
        """
        return self.component.Position

    # endregion Properties
