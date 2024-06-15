from __future__ import annotations
from typing import Any, cast, List, overload, Sequence, TYPE_CHECKING
import uno

from com.sun.star.sheet import XCellSeries
from com.sun.star.table import XCellRange
from ooo.dyn.sheet.cell_flags import CellFlagsEnum as CellFlagsEnum

from ooodev.mock import mock_g
from ooodev.adapter.sheet.sheet_cell_range_comp import SheetCellRangeComp
from ooodev.events.calc_named_event import CalcNamedEvent
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
from ooodev.utils import file_io as mFile
from ooodev.utils.color import CommonColor
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.data_type.generic_unit_size import GenericUnitSize
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.utils import info as mInfo
from ooodev.format.inner.partial.style.style_property_partial import StylePropertyPartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial
from ooodev.calc import calc_cell as mCalcCell
from ooodev.adapter.table.cell_properties2_partial_props import CellProperties2PartialProps


if TYPE_CHECKING:
    from com.sun.star.sheet import SheetCellRange
    from com.sun.star.table import CellAddress
    from ooo.dyn.table.cell_range_address import CellRangeAddress
    from ooodev.events.args.cancel_event_args import CancelEventArgs
    from ooodev.proto.style_obj import StyleT
    from ooodev.utils.color import Color
    from ooodev.utils.data_type.cell_obj import CellObj
    from ooodev.utils.data_type.cell_values import CellValues
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.size import Size
    from ooodev.utils.kind.chart2_types import ChartTemplateBase, ChartTypes as ChartTypes
    from ooodev.utils.type_var import Table, TupleArray, FloatTable, Row, PathOrStr
    from ooodev.format.calc.style import StyleCellKind
    from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
    from ooodev.calc import calc_cell_cursor as mCalcCellCursor
    from ooodev.calc.calc_sheet import CalcSheet
    from ooodev.calc.chart2.table_chart import TableChart
    from ooodev.calc.controls.cell_range_control import CellRangeControl
else:
    CellRangeAddress = Any
    ImgExportT = Any


class CalcCellRange(
    LoInstPropsPartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
    SheetCellRangeComp,
    EventsPartial,
    CellProperties2PartialProps,
    QiPartial,
    PropPartial,
    StylePartial,
    ServicePartial,
    TheDictionaryPartial,
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
    """Represents a calc cell range."""

    def __init__(self, owner: CalcSheet, rng: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (CalcSheet): Sheet that owns this cell range.
            rng (Any): Range object.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        CalcSheetPropPartial.__init__(self, obj=owner)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)
        if self.lo_inst.is_uno_interfaces(rng, XCellRange):
            with LoContext(self.lo_inst):
                self._range_obj = mCalc.Calc.get_range_obj(cell_range=rng)
            cell_range = rng
        else:
            with LoContext(self.lo_inst):
                self._range_obj = mCalc.Calc.get_range_obj(rng)
                cell_range = mCalc.Calc.get_cell_range(sheet=self.calc_sheet.component, range_obj=self._range_obj)
        SheetCellRangeComp.__init__(self, cell_range)  # type: ignore
        EventsPartial.__init__(self)
        CellProperties2PartialProps.__init__(self, component=cell_range)  # type: ignore
        QiPartial.__init__(self, component=cell_range, lo_inst=self.lo_inst)  # type: ignore
        PropPartial.__init__(self, component=cell_range, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=cell_range)
        TheDictionaryPartial.__init__(self)
        ServicePartial.__init__(self, component=cell_range, lo_inst=self.lo_inst)
        FontOnlyPartial.__init__(self, factory_name="ooodev.calc.cell_rng", component=cell_range, lo_inst=self.lo_inst)
        FontEffectsPartial.__init__(
            self, factory_name="ooodev.calc.cell_rng", component=cell_range, lo_inst=self.lo_inst
        )
        FontPartial.__init__(
            self, factory_name="ooodev.general_style.text", component=cell_range, lo_inst=self.lo_inst
        )
        TextAlignPartial.__init__(
            self, factory_name="ooodev.calc.cell_rng", component=cell_range, lo_inst=self.lo_inst
        )
        TextOrientationPartial.__init__(
            self, factory_name="ooodev.calc.cell_rng", component=cell_range, lo_inst=self.lo_inst
        )
        AlignPropertiesPartial.__init__(
            self, factory_name="ooodev.calc.cell_rng", component=cell_range, lo_inst=self.lo_inst
        )
        FillColorPartial.__init__(self, factory_name="ooodev.calc.cell_rng", component=cell_range, lo_inst=lo_inst)
        CalcBordersPartial.__init__(self, factory_name="ooodev.calc.cell_rng", component=cell_range, lo_inst=lo_inst)
        CellProtectionPartial.__init__(self, component=cell_range)
        NumbersNumbersPartial.__init__(
            self, factory_name="ooodev.number.numbers", component=cell_range, lo_inst=lo_inst
        )
        StylePropertyPartial.__init__(self, component=cell_range, property_name="CellStyle")
        self._control = None
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

    # region dunder methods
    def __eq__(self, other: Any) -> bool:
        """Compares two instances of CalcCellRange."""
        if isinstance(other, CalcCellRange):
            return self.range_obj == other.range_obj
        return False

    def __ne__(self, other: Any) -> bool:
        """Compares two instances of CalcCellRange."""
        return not self.__eq__(other)

    # endregion dunder methods

    # region StylePropertyPartial overrides

    def style_by_name(self, name: str | StyleCellKind = "") -> None:
        """
        Assign a style by name to the component.

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

    # region Chart2
    def insert_chart(
        self,
        *,
        cell_name: str = "",
        width: int = 16,
        height: int = 9,
        diagram_name: ChartTemplateBase | str = "Column",
        color_bg: Color | None = CommonColor.PALE_BLUE,
        color_wall: Color | None = CommonColor.LIGHT_BLUE,
        **kwargs,
    ) -> TableChart:
        """
        Insert a new chart.


        Args:
            cell_name (str, optional): Cell name such as ``A1``.
            width (int, optional): Width. Default ``16``.
            height (int, optional): Height. Default ``9``.
            diagram_name (ChartTemplateBase | str): Diagram Name. Defaults to ``Column``.
            color_bg (:py:data:`~.utils.color.Color`, optional): Color Background. Defaults to ``CommonColor.PALE_BLUE``.
                If set to ``None`` then no color is applied.
            color_wall (:py:data:`~.utils.color.Color`, optional): Color Wall. Defaults to ``CommonColor.LIGHT_BLUE``.
                If set to ``None`` then no color is applied.

        Keyword Arguments:
            chart_name (str, optional): Chart name
            is_row (bool, optional): Determines if the data is row data or column data.
            first_cell_as_label (bool, optional): Set is first row is to be used as a label.
            set_data_point_labels (bool, optional): Determines if the data point labels are set.

        Raises:
            ChartError: If error occurs

        Returns:
            TableChart: Chart Document that was created and inserted into the sheet.

        Note:
            **Keyword Arguments** are to mostly be ignored.
            If finer control over chart creation is needed then **Keyword Arguments** can be used.

        Note:
            See **Open Office Wiki** - `The Structure of Charts <https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Structure_of_Charts>`__ for more information.

        See Also:
            - :py:class:`~.color.CommonColor`
            - :ref:`ooodev.utils.kind.chart2_types`
        """
        from ooodev.office.chart2 import Chart2

        # from ..utils.kind.chart2_types import ChartTemplateBase, ChartTypeNameBase, ChartTypes as ChartTypes
        _ = Chart2.insert_chart(
            sheet=self.calc_sheet.component,
            cells_range=self._range_obj.get_cell_range_address(),
            cell_name=cell_name,
            width=width,
            height=height,
            diagram_name=diagram_name,
            color_bg=color_bg,
            color_wall=color_wall,
            **kwargs,
        )
        return self.calc_sheet.charts[-1]

    # endregion Chart2

    def change_style(self, style_name: str) -> bool:
        """
        Changes style of a range of cells.

        Args:
            style_name (str): Name of style to apply.
            range_obj (RangeObj): Range Object.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        return self.calc_sheet.change_style(style_name=style_name, range_obj=self._range_obj)

    def clear_cells(self, cell_flags: CellFlagsEnum | None = None) -> bool:
        """
        Clears the specified contents of the cell range

        If ``cell_flags`` is not specified then
        cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared

        Raises:
            MissingInterfaceError: If XSheetOperation interface cannot be obtained.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_CLEARING` :eventref:`src-docs-cell-event-clearing`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_CLEARED` :eventref:`src-docs-cell-event-cleared`

        Args:
            cell_flags (CellFlagsEnum, optional): Cell flags to clear. Default ``None``.

        Returns:
            bool: True if cells are cleared; Otherwise, False

        Note:
            Events arg for this method have a ``cell`` type of ``XCellRange``.

            Events arg ``event_data`` is a dictionary containing ``cell_flags``.

        Hint:
            - ``CellFlagsEnum`` can be imported from ``ooo.dyn.sheet.cell_flags``
        """
        if cell_flags is None:
            result = mCalc.Calc.clear_cells(sheet=self.calc_sheet.component, range_val=self._range_obj)
        else:
            result = mCalc.Calc.clear_cells(
                sheet=self.calc_sheet.component, range_val=self._range_obj, cell_flags=cell_flags
            )
        return result

    def delete_cells(self, is_shift_left: bool) -> bool:
        """
        Deletes cell in a spreadsheet

        Args:
            is_shift_left (bool): If True then cell are shifted left; Otherwise, cells are shifted up.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_DELETING` :eventref:`src-docs-cell-event-deleting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_DELETED` :eventref:`src-docs-cell-event-deleted`

        Returns:
            bool: True if cells are deleted; Otherwise, False

        Note:
            Events args for this method have a ``cell`` type of ``XCellRange``

        Note:
            Event args ``event_data`` is a dictionary containing ``is_shift_left``.
        """
        return mCalc.Calc.delete_cells(
            sheet=self.calc_sheet.component, range_obj=self._range_obj, is_shift_left=is_shift_left
        )

    def get_cell_series(self) -> XCellSeries:
        """
        Gets the cell series for the current range.

        Returns:
            XCellSeries: Cell series
        """
        return self.qi(XCellSeries, True)

    def get_cell_range(self) -> XCellRange:
        """
        Gets the ``XCellRange`` for the current range.

        Returns:
            XCellRange: Cell series
        """
        return self.qi(XCellRange, True)

    def get_col(self) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        return self.calc_sheet.get_col(range_obj=self._range_obj)

    def get_row(self) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            range_obj (RangeObj): Range Object

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        return self.calc_sheet.get_row(range_obj=self._range_obj)

    def get_cell_range_address(self) -> CellRangeAddress:
        """
        Gets the cell range address for the current range.

        Returns:
            CellRangeAddress: Cell range address
        """
        return self._range_obj.get_cell_range_address()

    def get_range_size(self) -> Size:
        """
        Gets the size of the range.

        Returns:
            ~ooodev.utils.data_type.size.Size: Size, Width is number of Columns and Height is number of Rows
        """
        return mCalc.Calc.get_range_size(range_obj=self._range_obj)

    def get_range_str(self) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``

        If ``sheet`` is included the format ``Sheet1.A1:B2`` is returned; Otherwise,
        ``A1:B2`` format is returned.

        Returns:
            str: range as string
        """
        return str(self._range_obj)

    def is_single_cell_range(self) -> bool:
        """
        Gets if a cell address is a single cell or a range

        Returns:
            bool: ``True`` if single cell; Otherwise, ``False``
        """
        return mCalc.Calc.is_single_cell_range(self.get_cell_range_address())

    def is_single_column_range(self) -> bool:
        """
        Gets if a cell address is a single column or multi-column

        Returns:
            bool: ``True`` if single column; Otherwise, ``False``
        """
        return mCalc.Calc.is_single_column_range(self.get_cell_range_address())

    def is_single_row_range(self) -> bool:
        """
        Gets if a cell address is a single row or multi-row

        Returns:
            bool: ``True`` if single row; Otherwise, ``False``
        """
        return mCalc.Calc.is_single_row_range(self.get_cell_range_address())

    def set_array(self, values: Table, styles: Sequence[StyleT] | None = None) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        if styles:
            self.calc_sheet.set_array(values=values, range_obj=self._range_obj, styles=styles)
        else:
            self.calc_sheet.set_array(values=values, range_obj=self._range_obj)

    def set_array_range(self, values: Table, styles: Sequence[StyleT] | None = None) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        if styles:
            self.calc_sheet.set_array_range(range_obj=self._range_obj, values=values, styles=styles)
        else:
            self.calc_sheet.set_array_range(range_obj=self._range_obj, values=values)

    def set_cell_range_array(self, values: Table, styles: Sequence[StyleT] | None = None) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        if styles:
            self.calc_sheet.set_cell_range_array(cell_range=self.component, values=values, styles=styles)
        else:
            self.calc_sheet.set_cell_range_array(cell_range=self.component, values=values)

    def is_merged_cells(self) -> bool:
        """
        Gets is a range of cells is merged.

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        return self.calc_sheet.is_merged_cells(range_obj=self._range_obj)

    def merge_cells(self, center: bool = False) -> None:
        """
        Merges a range of cells

        Args:
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:

        See Also:
            - :py:meth:`.Calc.unmerge_cells`
            - :py:meth:`.Calc.is_merged_cells`
        """
        self.calc_sheet.merge_cells(range_obj=self._range_obj, center=center)

    def unmerge_cells(self) -> None:
        """
        Removes merging from a range of cells

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            None:

        See Also:
            - :py:meth:`.Calc.merge_cells`
            - :py:meth:`.Calc.is_merged_cells`
        """
        self.calc_sheet.unmerge_cells(range_obj=self._range_obj)

    def set_val(self, value: Any) -> None:
        """
        Set the value of the very first cell in the range.

        Useful for merged cells.

        Args:
            value (Any): Value to set.
        """
        cell_obj = self.range_obj.start
        cell = mCalcCell.CalcCell(owner=self.calc_sheet, cell=cell_obj, lo_inst=self.lo_inst)
        cell.set_val(value=value)

    def select(self) -> None:
        """
        Selects the range of cells represented by this instance.

        Returns:
            None:

        .. versionadded:: 0.20.3
        """
        _ = mCalc.Calc.set_selected_range(
            doc=self.calc_sheet.calc_doc.component, sheet=self.calc_sheet.component, range_val=self._range_obj
        )

    def get_address(self) -> CellRangeAddress:
        """
        Gets Range Address.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        return self.calc_sheet.get_address(range_obj=self._range_obj)

    def get_array(self) -> TupleArray:
        """
        Gets a 2-Dimensional array of values from a range of cells.

        Returns:
            TupleArray: 2-Dimensional array of values.
        """
        return self.calc_sheet.get_array(range_obj=self._range_obj)

    def get_float_array(self) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        return self.calc_sheet.get_float_array(range_obj=self._range_obj)

    def get_val(self) -> Any:
        """
        Get the value of the very first cell in the range.

        Useful for merged cells.

        Returns:
            Any: Value of cell.
        """
        cell_obj = self.range_obj.start
        cell = mCalcCell.CalcCell(owner=self.calc_sheet, cell=cell_obj)
        return cell.get_val()

    def create_cursor(self) -> mCalcCellCursor.CalcCellCursor:
        """
        Creates a cell cursor to travel in the given range context.

        Returns:
            CalcCellCursor: Cell cursor
        """
        return self.calc_sheet.create_cursor_by_range(range_obj=self._range_obj)

    # region contains()

    @overload
    def contains(self, cell_obj: CellObj) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_obj (CellObj): Cell object

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.
        """
        ...

    @overload
    def contains(self, cell_addr: CellAddress) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_addr (CellAddress): Cell address

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.
        """
        ...

    @overload
    def contains(self, cell_vals: CellValues) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_vals (CellValues): Cell Values

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.
        """
        ...

    @overload
    def contains(self, cell_name: str) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_name (str): Cell name

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.
        """
        ...

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
        return self.range_obj.contains(*args, **kwargs)

    # endregion contains()

    # region export
    def export_as_image(self, fnm: PathOrStr, resolution: int = 96) -> None:
        """
        Exports a range of cells as an image.

        If the filename extension is ``.png`` then the image is exported as a PNG;
        Otherwise, image is exported as a JPEG.

        Args:
            fnm (PathOrStr): Filename to export to.
            resolution (int, optional): Resolution in dpi. Defaults to 96.

        Returns:
            None:

        Example:

            .. code-block:: python

                from pathlib import Path
                from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
                from ooodev.events.args.event_args_export import EventArgsExport
                from ooodev.calc.filter.export_png import ExportPngT
                from ooodev.calc import CalcNamedEvent

                def on_exporting(source: Any, args: CancelEventArgsExport[ExportJpgT]) -> None:
                    args.event_data["compression"] = 9

                def on_exported(source: Any, args: EventArgsGeneric[ImgExportT]) -> None:
                    print(f"Image has been exported to {args.event_data['file']}")

                rng = sheet.get_range(range_name="A1:M4")
                rng.subscribe_event(CalcNamedEvent.EXPORTING_RANGE_JPG, on_exporting)
                rng.subscribe_event(CalcNamedEvent.EXPORTED_RANGE_JPG, on_exported)
                pth = Path("./export_range.jpg")
                rng.export_as_image(pth)

        See Also:

            - Live LibreOffice Example `Export Calc Sheet Range as Image <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_export_calc_image>`__
            - :py:meth:`~ooodev.calc.CalcCellRange.export_jpg`
            - :py:meth:`~ooodev.calc.CalcCellRange.export_png`

        .. versionadded:: 0.20.3
        """
        ext = mFile.FileIO.get_ext(fnm=fnm)
        if ext and ext.lower() == "png":
            self.export_png(fnm=fnm, resolution=resolution)
        else:
            self.export_jpg(fnm=fnm, resolution=resolution)

    def export_jpg(self, fnm: PathOrStr, resolution: int = 96) -> None:
        """
        Exports page as jpg image.

        Args:
            fnm (PathOrStr, optional): Image file name.
            resolution (int, optional): Resolution in dpi. Defaults to ``96``.

        Raises:
            ValueError: If ``fnm`` is empty.
            CancelEventError: If ``EXPORTING_RANGE_JPG`` event is canceled.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~ooodev.events.calc_named_event.CalcNamedEvent.EXPORTING_RANGE_JPG` :eventref:`src-docs-event-cancel-export`
                - :py:attr:`~ooodev.events.calc_named_event.CalcNamedEvent.EXPORTED_RANGE_JPG` :eventref:`src-docs-event-export`

        Returns:
            None:

        Note:
            On exporting event is :ref:`cancel_event_args_export`.
            On exported event is :ref:`event_args_export`.
            Args ``event_data`` is a :py:class:`~ooodev.calc.filter.export_jpg.ExportJpgT` dictionary.

        Example:

            .. code-block:: python

                from pathlib import Path
                from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
                from ooodev.events.args.event_args_export import EventArgsExport
                from ooodev.calc.filter.export_jpg import ExportJpgT
                from ooodev.calc import CalcNamedEvent

                def on_exporting(source: Any, args: CancelEventArgsExport[ExportJpgT]) -> None:
                    args.event_data["quality"] = 80

                def on_exported(source: Any, args: EventArgsExport[ExportJpgT]) -> None:
                    print(f"Image has been exported to {args.fnm}")

                rng = sheet.get_range(range_name="A1:M4")
                rng.subscribe_event(CalcNamedEvent.EXPORTING_RANGE_JPG, on_exporting)
                rng.subscribe_event(CalcNamedEvent.EXPORTED_RANGE_JPG, on_exported)
                pth = Path("./export_range.jpg")
                rng.export_jpg(pth)

        See Also:
            Live LibreOffice Example `Export Calc Sheet Range as Image <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_export_calc_image>`__

        .. versionadded:: 0.21.0
        """
        from ooodev.calc.export.range_jpg import RangeJpg

        def on_exporting(source: Any, args: Any) -> None:
            self.trigger_event(CalcNamedEvent.EXPORTING_RANGE_JPG, args)

        def on_exported(source: Any, args: Any) -> None:
            self.trigger_event(CalcNamedEvent.EXPORTED_RANGE_JPG, args)

        exporter = RangeJpg(cell_range=self, lo_inst=self.lo_inst)
        exporter.subscribe_event_exporting(on_exporting)
        exporter.subscribe_event_exported(on_exported)
        exporter.export(fnm=fnm, resolution=resolution)

    def export_png(self, fnm: PathOrStr, resolution: int = 96) -> None:
        """
        Exports page as jpg image.

        Args:
            fnm (PathOrStr, optional): Image file name.
            resolution (int, optional): Resolution in dpi. Defaults to ``96``.

        Raises:
            ValueError: If ``fnm`` is empty.
            CancelEventError: If ``EXPORTING_RANGE_PNG`` event is canceled.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~ooodev.events.calc_named_event.CalcNamedEvent.EXPORTING_RANGE_PNG` :eventref:`src-docs-event-cancel-export`
                - :py:attr:`~ooodev.events.calc_named_event.CalcNamedEvent.EXPORTED_RANGE_PNG` :eventref:`src-docs-event-export`

        Returns:
            None:

        Note:
            On exporting event is :ref:`cancel_event_args_export`.
            On exported event is :ref:`event_args_export`.
            Args ``event_data`` is a :py:class:`~ooodev.calc.filter.export_png.ExportPngT` dictionary.


        Example:

            .. code-block:: python

                from pathlib import Path
                from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
                from ooodev.events.args.event_args_export import EventArgsExport
                from ooodev.calc.filter.export_png import ExportPngT
                from ooodev.calc import CalcNamedEvent

                def on_exporting(source: Any, args: CancelEventArgsExport[ExportPngT]) -> None:
                    args.event_data["compression"] = 9

                def on_exported(source: Any, args: EventArgsExport[ExportPngT]) -> None:
                    print(f"Image has been exported to {args.fnm}")

                rng = sheet.get_range(range_name="A1:M4")
                rng.subscribe_event(CalcNamedEvent.EXPORTING_RANGE_JPG, on_exporting)
                rng.subscribe_event(CalcNamedEvent.EXPORTED_RANGE_JPG, on_exported)
                pth = Path("./export_range.png")
                rng.export_png(pth)

        See Also:
            Live LibreOffice Example `Export Calc Sheet Range as Image <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_export_calc_image>`__

        .. versionadded:: 0.21.0
        """
        from ooodev.calc.export.range_png import RangePng

        def on_exporting(source: Any, args: Any) -> None:
            self.trigger_event(CalcNamedEvent.EXPORTING_RANGE_PNG, args)

        def on_exported(source: Any, args: Any) -> None:
            self.trigger_event(CalcNamedEvent.EXPORTED_RANGE_PNG, args)

        exporter = RangePng(cell_range=self, lo_inst=self.lo_inst)
        exporter.subscribe_event_exporting(on_exporting)
        exporter.subscribe_event_exported(on_exported)
        exporter.export(fnm=fnm, resolution=resolution)

    # endregion export
    @overload
    def highlight(self, headline: str) -> mCalcCell.CalcCell:
        """
        Draw a light blue colored border around the range and write a headline in the
        top-left cell of the range.

        Args:
            headline (str): Headline.

        Returns:
            CalcCell: Cell object. First cell of range that headline ia applied on.
        """
        ...

    @overload
    def highlight(self, headline: str, color: Color) -> mCalcCell.CalcCell:
        """
        Draw a colored border around the range and write a headline in the
        top-left cell of the range.

        Args:
            headline (str): Headline.
            color (~ooodev.utils.color.Color): RGB color.

        Returns:
            CalcCell: Cell object. First cell of range that headline ia applied on.
        """
        ...

    def highlight(self, headline: str, color: Color = CommonColor.LIGHT_BLUE) -> mCalcCell.CalcCell:
        """
        Draw a colored border around the range and write a headline in the
        top-left cell of the range.

        Args:
            headline (str): Headline.
            color (~ooodev.utils.color.Color): RGB color.

        Returns:
            CalcCell: Cell object. First cell of range that headline ia applied on.
        """
        cell = mCalc.Calc.highlight_range(
            sheet=self.calc_sheet.component, headline=headline, range_obj=self._range_obj, color=color
        )
        cell_obj = mCalc.Calc.get_cell_obj(cell=cell)
        return mCalcCell.CalcCell(owner=self.calc_sheet, cell=cell_obj, lo_inst=self.lo_inst)

    def remove_border(self) -> None:
        """
        Removes border from range of cells.

        Returns:
            None:

        .. versionadded:: 0.25.2
        """
        _ = mCalc.Calc.remove_border(sheet=self.calc_sheet.component, range_obj=self._range_obj)

    def refresh(self) -> None:
        """
        Refreshes the range of cells.

        This method should be call if the sheet has rows or columns inserted or deleted since this instance as been created.

        Returns:
            None:

        .. versionadded:: 0.45.0
        """
        rng = self.component.getRangeAddress()
        rng_obj = RangeObj.from_range(rng)
        if rng_obj != self._range_obj:
            self._range_obj = rng_obj

    # region Static Methods

    @classmethod
    def from_obj(cls, obj: Any, lo_inst: LoInst | None = None) -> CalcCellRange | None:
        """
        Creates a CalcCellRange from an object.

        Args:
            obj (Any): Object to create CalcCellRange from. Can be a CalcCellRange, CalcCellRangePropPartial, or any object that can be converted to a CalcCellRange such as a cell.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to ``None``.

        Returns:
            CalcSheet: CalcSheet if found; Otherwise, ``None``
        
        .. versionadded:: 0.46.0
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.calc.calc_sheet import CalcSheet

        if mInfo.Info.is_instance(obj, CalcCellRange):
            return obj

        calc_sheet = CalcSheet.from_obj(obj=obj, lo_inst=lo_inst)
        if calc_sheet is None:
            return None

        if hasattr(obj, "component"):
            obj = obj.component

        if not hasattr(obj, "getImplementationName"):
            return None

        cell_rng = None
        imp_name = obj.getImplementationName()
        if imp_name == "ScCellRangeObj":
            cell_rng = obj

        if cell_rng is None:
            return None

        addr = cast("CellRangeAddress", cell_rng.getRangeAddress())

        return cls(owner=calc_sheet, rng=addr, lo_inst=lo_inst)

    # endregion Static Methods

    # region Properties

    @property
    def range_obj(self) -> RangeObj:
        """Range object."""
        return self._range_obj

    @property
    def size(self) -> GenericUnitSize[UnitMM, float]:
        """Gets the size of the cell range in ``UnitMM`` Values."""
        if not self.support_service("com.sun.star.sheet.SheetCellRange"):
            raise mEx.ServiceNotSupported("com.sun.star.sheet.SheetCellRange")
        comp = cast("SheetCellRange", self.component)
        sz = comp.Size
        return GenericUnitSize(UnitMM.from_mm100(sz.Width), UnitMM.from_mm100(sz.Height))

    @property
    def control(self) -> CellRangeControl:
        """
        Gets access to class for managing cell control.

        Returns:
            CellRangeControl: Cell Range Control instance.
        """
        # pylint: disable=import-outside-toplevel
        # pylint: disable=redefined-outer-name
        if self._control is None:
            from ooodev.calc.controls.cell_range_control import CellRangeControl

            self._control = CellRangeControl(self, self.lo_inst)
            self._control.add_event_observers(self.event_observer)
        return self._control

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.office.chart2 import Chart2
    from ooodev.calc.export.range_jpg import RangeJpg
    from ooodev.calc.export.range_png import RangePng
