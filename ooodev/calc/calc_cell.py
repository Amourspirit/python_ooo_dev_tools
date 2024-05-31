from __future__ import annotations
from typing import Any, overload, Sequence, TYPE_CHECKING
import uno
from com.sun.star.uno import RuntimeException

from ooodev.mock import mock_g
from ooodev.adapter.sheet.sheet_cell_comp import SheetCellComp
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.partial.area.fill_color_partial import FillColorPartial
from ooodev.format.inner.partial.calc.alignment.properties_partial import PropertiesPartial as AlignPropertiesPartial
from ooodev.format.inner.partial.calc.alignment.text_align_partial import TextAlignPartial
from ooodev.format.inner.partial.calc.alignment.text_orientation_partial import TextOrientationPartial
from ooodev.format.inner.partial.calc.borders.calc_borders_partial import CalcBordersPartial
from ooodev.format.inner.partial.calc.cell_protection.cell_protection_partial import CellProtectionPartial
from ooodev.format.inner.partial.calc.font.font_effects_partial import FontEffectsPartial
from ooodev.format.inner.partial.font.font_only_partial import FontOnlyPartial
from ooodev.format.inner.partial.font.font_partial import FontPartial
from ooodev.format.inner.partial.numbers.numbers_numbers_partial import NumbersNumbersPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.office import calc as mCalc
from ooodev.units.unit_mm import UnitMM
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.data_type import cell_obj as mCellObj
from ooodev.utils.data_type.generic_unit_point import GenericUnitPoint
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.gen_util import NULL_OBJ
from ooodev.utils.type_var import Row, Table
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.format.inner.partial.style.style_property_partial import StylePropertyPartial
from ooodev.calc.partial.calc_cell_prop_partial import CalcCellPropPartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial
from ooodev.adapter.table.cell_properties2_partial_props import CellProperties2PartialProps

if TYPE_CHECKING:
    from com.sun.star.awt import Point
    from com.sun.star.sheet import SolverConstraint  # struct
    from com.sun.star.sheet import XGoalSeek
    from com.sun.star.sheet import XSheetAnnotation
    from com.sun.star.text import XTextRange
    from ooo.dyn.sheet.solver_constraint_operator import SolverConstraintOperator
    from ooodev.proto.style_obj import StyleT
    from ooodev.events.args.cancel_event_args import CancelEventArgs
    from ooodev.format.calc.style import StyleCellKind
    from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
    from ooodev.calc.calc_sheet import CalcSheet
    from ooodev.calc import calc_cell_cursor as mCalcCellCursor
    from ooodev.calc.calc_cell_text_cursor import CalcCellTextCursor
    from ooodev.units.unit_obj import UnitT
    from ooodev.calc.controls.cell_control import CellControl
    from ooodev.calc.cell.custom_prop import CustomProp
else:
    XSheetAnnotation = Any
    UnitT = Any


class CalcCell(
    LoInstPropsPartial,
    SheetCellComp,
    EventsPartial,
    CellProperties2PartialProps,
    QiPartial,
    PropPartial,
    StylePartial,
    ServicePartial,
    CalcCellPropPartial,
    CalcSheetPropPartial,
    CalcDocPropPartial,
    FontOnlyPartial,
    FontEffectsPartial,
    FontPartial,
    TextAlignPartial,
    TextOrientationPartial,
    AlignPropertiesPartial,
    FillColorPartial,
    CalcBordersPartial,
    CellProtectionPartial,
    NumbersNumbersPartial,
    StylePropertyPartial,
):
    def __init__(self, owner: CalcSheet, cell: str | mCellObj.CellObj, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        self._cell_obj = mCellObj.CellObj.from_cell(cell)
        # don't use owner.get_cell() here because it will be recursive.
        sheet_cell = mCalc.Calc.get_cell(sheet=owner.component, cell_obj=self._cell_obj)
        SheetCellComp.__init__(self, sheet_cell)  # type: ignore
        EventsPartial.__init__(self)
        CellProperties2PartialProps.__init__(self, component=sheet_cell)  # type: ignore
        QiPartial.__init__(self, component=sheet_cell, lo_inst=self.lo_inst)
        PropPartial.__init__(self, component=sheet_cell, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=sheet_cell)
        ServicePartial.__init__(self, component=sheet_cell, lo_inst=self.lo_inst)
        CalcCellPropPartial.__init__(self, obj=self)
        CalcSheetPropPartial.__init__(self, obj=owner)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)
        FontOnlyPartial.__init__(self, factory_name="ooodev.calc.cell", component=sheet_cell, lo_inst=self.lo_inst)
        FontEffectsPartial.__init__(self, factory_name="ooodev.calc.cell", component=sheet_cell, lo_inst=self.lo_inst)
        FontPartial.__init__(
            self, factory_name="ooodev.general_style.text", component=sheet_cell, lo_inst=self.lo_inst
        )
        TextAlignPartial.__init__(self, factory_name="ooodev.calc.cell", component=sheet_cell, lo_inst=self.lo_inst)
        TextOrientationPartial.__init__(
            self, factory_name="ooodev.calc.cell", component=sheet_cell, lo_inst=self.lo_inst
        )
        AlignPropertiesPartial.__init__(
            self, factory_name="ooodev.calc.cell", component=sheet_cell, lo_inst=self.lo_inst
        )
        FillColorPartial.__init__(self, factory_name="ooodev.calc.cell", component=sheet_cell, lo_inst=lo_inst)
        CalcBordersPartial.__init__(self, factory_name="ooodev.calc.cell", component=sheet_cell, lo_inst=lo_inst)
        CellProtectionPartial.__init__(self, component=sheet_cell)
        NumbersNumbersPartial.__init__(
            self, factory_name="ooodev.number.numbers", component=sheet_cell, lo_inst=lo_inst
        )
        StylePropertyPartial.__init__(self, component=sheet_cell, property_name="CellStyle")
        self._control = None
        self._custom_properties = None
        self._init_events()

    def _init_events(self) -> None:
        self._fn_on_before_style_number_number = self._on_before_style_number_number
        self._fn_on_style_by_name_default_prop_setting = self._on_style_by_name_default_prop_setting
        self.subscribe_event(event_name="before_style_number_number", callback=self._fn_on_before_style_number_number)
        self.subscribe_event(
            event_name="style_by_name_default_prop_setting", callback=self._fn_on_style_by_name_default_prop_setting
        )

    def _on_before_style_number_number(self, src: Any, event: CancelEventArgs) -> None:
        event.event_data["component"] = self.calc_doc.component

    def _on_style_by_name_default_prop_setting(self, src: Any, event: KeyValCancelArgs) -> None:
        # this event is triggered by StylePropertyPartial.style_by_name()
        # when property is setting default value this is triggered.
        # In this case we want the style to be set to the default property value.
        event.default = True

    # region SimpleTextPartial Overrides

    def create_text_cursor(self) -> CalcCellTextCursor:
        """
        Creates a text cursor to travel in the given range context.

        Cursor can be used to insert text, paragraphs, hyperlinks, and other text content.

        Returns:
            CalcCellTextCursor: Text cursor

        .. versionadded:: 0.28.4
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.calc.calc_cell_text_cursor import CalcCellTextCursor

        cursor = self.component.createTextCursor()
        return CalcCellTextCursor(owner=self, cursor=cursor, lo_inst=self.lo_inst)

    def create_text_cursor_by_range(self, text_position: XTextRange) -> CalcCellTextCursor:
        """
        The initial position is set to ``text_position``.

        Cursor can be used to insert text, paragraphs, hyperlinks, and other text content.

        Args:
            text_position (XTextRange): The initial position of the new text cursor.

        Returns:
            CalcCellTextCursor: The new text cursor.

        .. versionadded:: 0.28.4
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.calc.calc_cell_text_cursor import CalcCellTextCursor

        cursor = self.component.createTextCursorByRange(text_position)
        return CalcCellTextCursor(owner=self, cursor=cursor, lo_inst=self.lo_inst)

    # endregion SimpleTextPartial Overrides

    def create_cursor(self) -> mCalcCellCursor.CalcCellCursor:
        """
        Creates a cell cursor to travel in the given range context.

        Returns:
            CalcCellCursor: Cell cursor
        """
        return self.calc_sheet.create_cursor_by_range(cell_obj=self._cell_obj)

    # region StylePropertyPartial overrides

    def style_by_name(self, name: str | StyleCellKind = "") -> None:
        """
        Assign a style by name to the component.

        Args:
            name (str, StyleCellKind, optional): The name of the style to apply. ``StyleCellKind`` contains various style names.
                If not provided, the default style is applied.

        Raises:
            CancelEventError: If the event ``before_style_by_name`` is cancelled and not handled.

        Returns:
            None:

        Hint:
            - ``StyleCellKind`` can be imported from ``ooodev.format.calc.style``
        """
        super().style_by_name(name=str(name))

    # endregion StylePropertyPartial overrides

    # region dunder methods
    def __eq__(self, other: Any) -> bool:
        """Compares two instances of CalcCell."""
        if isinstance(other, CalcCell):
            return self.cell_obj == other.cell_obj
        return False

    def __ne__(self, other: Any) -> bool:
        """Compares two instances of CalcCell."""
        return not self.__eq__(other)

    # endregion dunder methods

    # region Cell Properties

    def is_first_row(self) -> bool:
        """Determines if this cell is in the first row of the sheet."""
        return self._cell_obj.row == 1

    def is_first_column(self) -> bool:
        """Determines if this cell is in the first column of the sheet."""
        return self._cell_obj.col == "A"

    # endregion Cell Properties
    def goto(self) -> None:
        """
        Go to this cell in the spreadsheet.

        Returns:
            None:

        Attention:
            :py:meth:`~.utils.lo.Lo.dispatch_cmd` method is called along with any of its events.

            Dispatch command is ``GoToCell``.

        .. versionadded:: 0.20.2
        """
        _ = self.calc_sheet.goto_cell(cell_obj=self._cell_obj)

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
        return mCalc.Calc.get_cell_position(cell_name=self._cell_obj)

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
            raise mEx.CellError(f"Cell {self._cell_obj} is in the first column of the sheet.")
        cell_obj = self._cell_obj.left
        return CalcCell(owner=self.calc_sheet, cell=cell_obj, lo_inst=self.lo_inst)

    def get_cell_right(self) -> CalcCell:
        """
        Gets the cell to the right of this cell.

        Raises:
            CellError: If cell is in the last column of the sheet.

        Returns:
            CalcCell: Cell to the right of this cell.
        """
        cell_obj = self._cell_obj.right
        return CalcCell(owner=self.calc_sheet, cell=cell_obj, lo_inst=self.lo_inst)

    def get_cell_up(self) -> CalcCell:
        """
        Gets the cell above this cell.

        Returns:
            CalcCell: Cell above this cell.
        """
        if self.is_first_row():
            raise mEx.CellError(f"Cell {self._cell_obj} is in the first row of the sheet.")
        cell_obj = self._cell_obj.up
        return CalcCell(owner=self.calc_sheet, cell=cell_obj, lo_inst=self.lo_inst)

    def get_cell_down(self) -> CalcCell:
        """
        Gets the cell below this cell.

        Returns:
            CalcCell: Cell below this cell.
        """
        cell_obj = self._cell_obj.down
        return CalcCell(owner=self.calc_sheet, cell=cell_obj, lo_inst=self.lo_inst)

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
        with LoContext(self.lo_inst):
            result = mCalc.Calc.add_annotation(
                sheet=self.calc_sheet.component, cell_name=str(self._cell_obj), msg=msg, is_visible=is_visible
            )
        return result

    def get_annotation(self) -> XSheetAnnotation:
        """
        Gets the annotation of a cell.

        Returns:
            XSheetAnnotation: Cell annotation
        """
        return mCalc.Calc.get_annotation(sheet=self.calc_sheet.component, cell_name=self._cell_obj)

    def get_annotation_str(self) -> str:
        """
        Gets text of an annotation for a cell.

        Returns:
            str: Cell annotation text
        """
        return mCalc.Calc.get_annotation_str(sheet=self.calc_sheet.component, cell_name=self._cell_obj)

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
        return str(self._cell_obj)

    def get_type_enum(self) -> mCalc.Calc.CellTypeEnum:
        """
        Gets enum representing the Type

        Returns:
            CellTypeEnum: Enum of cell type
        """
        # does not need context manager
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
            self.calc_sheet.set_array_cell(cell_obj=self._cell_obj, values=values)
        else:
            self.calc_sheet.set_array_cell(cell_obj=self._cell_obj, values=values, styles=styles)

    # endregion set_array_cell()
    def set_date(self, day: int, month: int, year: int) -> None:
        """
        Writes a date with standard date format into a spreadsheet

        Args:
            day (int): Date day part.
            month (int): Date month part.
            year (int): Date year part.
        """
        with LoContext(self.lo_inst):
            mCalc.Calc.set_date(
                sheet=self.calc_sheet.component, cell_name=self._cell_obj, day=day, month=month, year=year
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
            cell_obj=self._cell_obj,
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
        index = self._cell_obj.row - 1
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
        mCalc.Calc.set_style_cell(sheet=self.calc_sheet.component, cell_obj=self._cell_obj, styles=styles)

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
            self.calc_sheet.set_val(value=value, cell_obj=self._cell_obj)
        else:
            self.calc_sheet.set_val(value=value, cell_obj=self._cell_obj, styles=styles)

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
        with LoContext(self.lo_inst):
            mCalc.Calc.split_window(doc=self.calc_sheet.calc_doc.component, cell_name=str(self._cell_obj))

    # region make_constraint()
    @overload
    def make_constraint(self, num: int | float, op: str) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        Args:
            num (Number): Constraint number such as float or int.
            op (str): Operation such as ``<=``.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.
        """
        ...

    @overload
    def make_constraint(self, num: int | float, op: SolverConstraintOperator) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        Args:
            num (Number): Constraint number such as float or int.
            op (SolverConstraintOperator): Operation such as ``SolverConstraintOperator.EQUAL``.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.

        Hint:
            - ``SolverConstraintOperator`` can be imported from ``ooo.dyn.sheet.solver_constraint_operator``
        """
        ...

    def make_constraint(self, num: int | float, op: SolverConstraintOperator | str) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        Args:
            num (Number): Constraint number such as float or int.
            op (str | SolverConstraintOperator): Operation such as ``<=``.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.

        Hint:
            - ``SolverConstraintOperator`` can be imported from ``ooo.dyn.sheet.solver_constraint_operator``
        """
        if isinstance(op, str):
            return mCalc.Calc.make_constraint(num=num, op=op, sheet=self.calc_sheet.component, cell_obj=self._cell_obj)
        else:
            return mCalc.Calc.make_constraint(
                num=num, op=op, sheet=self.calc_sheet.component, cell_name=str(self._cell_obj)
            )

    # endregion make_constraint()

    # region Refresh
    def refresh(self) -> None:
        """
        Refreshes this instance cell.

        This method should be call if the sheet has rows or columns inserted or deleted since this instance was created that affect the cell address.

        Raises:
            CellDeletedError: If the cell has been deleted.

        Returns:
            None:

        See Also:
            - :py:meth:`~.calc_cell.CalcCell.is_cell_deleted`

        .. versionadded:: 0.45.0
        """
        if self.is_cell_deleted():
            raise mEx.CellDeletedError(f"Cell {self._cell_obj} has been deleted.")
        self._custom_properties = None
        cell = self.component.getCellAddress()
        cell_obj = mCellObj.CellObj.from_cell(cell)
        if self._cell_obj != cell_obj:
            self._cell_obj = cell_obj

    def is_cell_deleted(self) -> bool:
        """
        Determines if this cell has been deleted.

        Returns:
            bool: ``True`` if cell is deleted; Otherwise, ``False``.

        .. versionadded:: 0.45.2
        """
        try:
            assert self.component.AbsoluteName
        except RuntimeException:
            return True
        return False

    # endregion Refresh

    # region Custom Properties Methods
    def _get_custom_properties(self) -> CustomProp:
        # pylint: disable=import-outside-toplevel
        # pylint: disable=redefined-outer-name
        if self._custom_properties is not None:
            return self._custom_properties
        from ooodev.calc.cell.custom_prop import CustomProp

        self._custom_properties = CustomProp(cell=self)
        return self._custom_properties

    # endregion Custom Properties

    # region Properties Methods

    # region Properties

    # region    Custom Properties
    def get_custom_property(self, name: str, default: Any = NULL_OBJ) -> Any:
        """
        Gets a custom property.

        Args:
            name (str): The name of the property.
            default (Any, optional): The default value to return if the property does not exist.

        Raises:
            AttributeError: If the property is not found.

        Returns:
            Any: The value of the property.
        """
        cp = self._get_custom_properties()
        return cp.get_custom_property(name=name, default=default)

    def get_custom_properties(self) -> DotDict:
        """
        Gets custom properties.

        Returns:
            DotDict: custom properties.

        Hint:
            DotDict is a class that allows you to access dictionary keys as attributes or keys.
            DotDict can be imported from ``ooodev.utils.helper.dot_dict.DotDict``.
        """
        cp = self._get_custom_properties()
        return cp.get_custom_properties()

    def has_custom_property(self, name: str) -> bool:
        """
        Gets if a custom property exists.

        Args:
            name (str): The name of the property to check.

        Returns:
            bool: ``True`` if the property exists, otherwise ``False``.
        """
        cp = self._get_custom_properties()
        return cp.has_custom_property(name=name)

    def has_custom_properties(self) -> bool:
        """
        Gets if a custom properties exists.

        Returns:
            bool: ``True`` if the properties exists, otherwise ``False``.
        """
        cp = self._get_custom_properties()
        return cp.has_custom_properties()

    def set_custom_property(self, name: str, value: Any):
        """
        Sets a custom property.

        Args:
            name (str): The name of the property.
            value (Any): The value of the property.

        Raises:
            AttributeError: If the property is a forbidden key.
        """
        cp = self._get_custom_properties()
        cp.set_custom_property(name=name, value=value)

    def set_custom_properties(self, properties: DotDict) -> None:
        """
        Sets custom properties.

        Args:
            properties (DotDict): custom properties to set.

        Hint:
            DotDict is a class that allows you to access dictionary keys as attributes or keys.
            DotDict can be imported from ``ooodev.utils.helper.dot_dict.DotDict``.
        """
        cp = self._get_custom_properties()
        cp.set_custom_properties(properties=properties)

    def remove_custom_property(self, name: str) -> None:
        """
        Removes a custom property.

        Args:
            name (str): The name of the property to remove.

        Raises:
            AttributeError: If the property is a forbidden key.

        Returns:
            None:
        """
        cp = self._get_custom_properties()
        cp.remove_custom_property(name=name)

    def remove_custom_properties(self) -> None:
        """
        Removes any custom properties.

        Returns:
            None:
        """
        cp = self._get_custom_properties()
        cp.remove_custom_properties()

    # endregion Custom Properties

    @property
    def cell_obj(self) -> mCellObj.CellObj:
        """Cell object."""
        return self._cell_obj

    @property
    def position(self) -> GenericUnitPoint[UnitMM, float]:
        """
        Gets the Position of the cell in ``UnitMM`` Values.

        Contains the position of the top left cell of this range.

        This property contains the absolute position in the whole sheet,
        not the position in the visible area.

        .. versionchanged:: 0.20.1
            Now return :ref:`generic_unit_point` instead of ``Point``.
        """
        ps = self.component.Position
        return GenericUnitPoint(UnitMM.from_mm100(ps.X), UnitMM.from_mm100(ps.Y))

    @property
    def value(self) -> str | float | None:
        """
        Gets/Sets the value of cell.

        If the cell has a value, then the value is returned; Otherwise, None is returned.

        .. versionadded:: 0.20.1
        """
        return self.get_val()

    @value.setter
    def value(self, value: Any) -> None:
        """Sets value of cell."""
        self.set_val(value=value)

    @property
    def control(self) -> CellControl:
        """
        Gets access to class for managing cell control.

        Returns:
            CellControl: Cell control instance.
        """
        # pylint: disable=import-outside-toplevel
        # pylint: disable=redefined-outer-name
        if self._control is None:
            from ooodev.calc.controls.cell_control import CellControl

            self._control = CellControl(self, self.lo_inst)
            self._control.add_event_observers(self.event_observer)
        return self._control

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.calc.calc_cell_text_cursor import CalcCellTextCursor
    from ooodev.calc.controls.cell_control import CellControl
    from ooodev.calc.cell.custom_prop import CustomProp
