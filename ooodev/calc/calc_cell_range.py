from __future__ import annotations
from typing import Any, List, overload, Sequence, TYPE_CHECKING
import uno


from com.sun.star.sheet import XCellSeries
from com.sun.star.table import XCellRange
from com.sun.star.frame import XStorable
from ooo.dyn.sheet.cell_flags import CellFlagsEnum as CellFlagsEnum
from ooo.dyn.beans.property_state import PropertyState  # enum


if TYPE_CHECKING:
    from com.sun.star.table import CellAddress
    from ooo.dyn.table.cell_range_address import CellRangeAddress
    from ooodev.proto.style_obj import StyleT
    from ooodev.utils.data_type.cell_obj import CellObj
    from ooodev.utils.data_type.cell_values import CellValues
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.size import Size
    from ooodev.utils.type_var import Table, TupleArray, FloatTable, Row, PathOrStr
    from . import calc_cell_cursor as mCalcCellCursor
    from .calc_sheet import CalcSheet
    from ooodev.events.event_data.img_export_t import ImgExportT
else:
    CellRangeAddress = Any
    ImgExportT = Any

from ooodev.adapter.sheet.sheet_cell_range_comp import SheetCellRangeComp
from ooodev.calc import CalcNamedEvent
from ooodev.events.args.cancel_event_args_generic import CancelEventArgsGeneric
from ooodev.events.args.event_args_generic import EventArgsGeneric
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import calc as mCalc
from ooodev.utils import file_io as mFile
from ooodev.utils import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from . import calc_cell as mCalcCell


class CalcCellRange(SheetCellRangeComp, QiPartial, PropPartial, StylePartial, EventsPartial):
    """Represents a calc cell range."""

    def __init__(self, owner: CalcSheet, rng: Any) -> None:
        """
        Constructor

        Args:
            owner (CalcSheet): Sheet that owns this cell range.
            rng (Any): Range object.
        """
        self.__owner = owner
        if mLo.Lo.is_uno_interfaces(rng, XCellRange):
            self.__range_obj = mCalc.Calc.get_range_obj(cell_range=rng)
            cell_range = rng
        else:
            self.__range_obj = mCalc.Calc.get_range_obj(rng)
            cell_range = mCalc.Calc.get_cell_range(sheet=self.calc_sheet.component, range_obj=self.__range_obj)
        SheetCellRangeComp.__init__(self, cell_range)  # type: ignore
        QiPartial.__init__(self, component=cell_range, lo_inst=mLo.Lo.current_lo)  # type: ignore
        PropPartial.__init__(self, component=cell_range, lo_inst=mLo.Lo.current_lo)  # type: ignore
        StylePartial.__init__(self, component=cell_range)
        EventsPartial.__init__(self)
        # self.__doc = doc

    def change_style(self, style_name: str) -> bool:
        """
        Changes style of a range of cells.

        Args:
            style_name (str): Name of style to apply.
            range_obj (RangeObj): Range Object.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        return self.calc_sheet.change_style(style_name=style_name, range_obj=self.__range_obj)

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
        """
        if cell_flags is None:
            return mCalc.Calc.clear_cells(sheet=self.calc_sheet.component, range_val=self.__range_obj)
        return mCalc.Calc.clear_cells(
            sheet=self.calc_sheet.component, range_val=self.__range_obj, cell_flags=cell_flags
        )

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
            sheet=self.calc_sheet.component, range_obj=self.__range_obj, is_shift_left=is_shift_left
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
        return self.calc_sheet.get_col(range_obj=self.__range_obj)

    def get_row(self) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            range_obj (RangeObj): Range Object

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        return self.calc_sheet.get_row(range_obj=self.__range_obj)

    def get_cell_range_address(self) -> CellRangeAddress:
        """
        Gets the cell range address for the current range.

        Returns:
            CellRangeAddress: Cell range address
        """
        return self.__range_obj.get_cell_range_address()

    def get_range_size(self) -> Size:
        """
        Gets the size of the range.

        Returns:
            ~ooodev.utils.data_type.size.Size: Size, Width is number of Columns and Height is number of Rows
        """
        return mCalc.Calc.get_range_size(range_obj=self.__range_obj)

    def get_range_str(self) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``

        If ``sheet`` is included the format ``Sheet1.A1:B2`` is returned; Otherwise,
        ``A1:B2`` format is returned.

        Returns:
            str: range as string
        """
        return str(self.__range_obj)

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
            self.calc_sheet.set_array(values=values, range_obj=self.__range_obj, styles=styles)
        else:
            self.calc_sheet.set_array(values=values, range_obj=self.__range_obj)

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
            self.calc_sheet.set_array_range(range_obj=self.__range_obj, values=values, styles=styles)
        else:
            self.calc_sheet.set_array_range(range_obj=self.__range_obj, values=values)

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
        mCalc.Calc.set_style_range(sheet=self.calc_sheet.component, range_obj=self.__range_obj, styles=styles)

    def is_merged_cells(self) -> bool:
        """
        Gets is a range of cells is merged.

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        return self.calc_sheet.is_merged_cells(range_obj=self.__range_obj)

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
        self.calc_sheet.merge_cells(range_obj=self.__range_obj, center=center)

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
        self.calc_sheet.unmerge_cells(range_obj=self.__range_obj)

    def set_val(self, value: Any) -> None:
        """
        Set the value of the very first cell in the range.

        Useful for merged cells.

        Args:
            value (Any): Value to set.
        """
        cell_obj = self.range_obj.start
        cell = mCalcCell.CalcCell(owner=self.calc_sheet, cell=cell_obj)
        cell.set_val(value=value)

    def select(self) -> None:
        """
        Selects the range of cells represented by this instance.

        Returns:
            None:

        .. versionadded:: 0.20.3
        """
        _ = mCalc.Calc.set_selected_range(
            doc=self.calc_sheet.calc_doc.component, sheet=self.calc_sheet.component, range_val=self.__range_obj
        )

    def get_address(self) -> CellRangeAddress:
        """
        Gets Range Address.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        return self.calc_sheet.get_address(range_obj=self.__range_obj)

    def get_array(self) -> TupleArray:
        """
        Gets a 2-Dimensional array of values from a range of cells.

        Returns:
            TupleArray: 2-Dimensional array of values.
        """
        return self.calc_sheet.get_array(range_obj=self.__range_obj)

    def get_float_array(self) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        return self.calc_sheet.get_float_array(range_obj=self.__range_obj)

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
        return self.calc_sheet.create_cursor_by_range(range_obj=self.__range_obj)

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
    def export_as_image(self, fnm: PathOrStr) -> None:
        """
        Exports a range of cells as an image.

        If the filename extension is ``.png`` then the image is exported as a PNG;
        Otherwise, image is exported as a JPEG.

        Args:
            fnm (PathOrStr): Filename to export to.

        Returns:
            None:

        Raises:
            CancelEventError: If ``RANGE_EXPORTING_IMAGE`` is canceled.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.RANGE_EXPORTING_IMAGE` :eventref:`src-docs-event-cancel-generic`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.RANGE_EXPORTED_IMAGE` :eventref:`src-docs-event-generic`

        Note:
            Event args ``event_data`` an instance of ``~ooodev.events.event_data.img_export_t.ImgExportT``.

        Example:

            .. code-block:: python

                from ooodev.events.args.cancel_event_args_generic import CancelEventArgsGeneric
                from ooodev.events.args.event_args_generic import EventArgsGeneric
                from ooodev.events.event_data.img_export_t import ImgExportT
                from ooodev.calc import CalcNamedEvent

                def on_exporting(source: Any, args: CancelEventArgsGeneric[ImgExportT]) -> None:
                    args.event_data["compression"] = 9

                def on_exported(source: Any, args: EventArgsGeneric[ImgExportT]) -> None:
                    print(f"Image has been exported to {args.event_data['file']}")

                rng = sheet.get_range(range_name="A1:M4")
                rng.subscribe_event(CalcNamedEvent.RANGE_EXPORTING_IMAGE, on_exporting)
                rng.subscribe_event(CalcNamedEvent.RANGE_EXPORTED_IMAGE, on_exported)
                rng.export_as_image(pth)

        See Also:
            `Export Calc Sheet Range as Image Example <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_export_calc_image>`__

        .. versionadded:: 0.20.3
        """
        ext = mFile.FileIO.get_ext(fnm=fnm)
        if ext and ext.lower() == "png":
            self._export_as_png(fnm=fnm)
        else:
            self._export_as_jpg(fnm=fnm)

    def _export_as_jpg(self, fnm: PathOrStr) -> None:
        """
        Exports a range of cells as an jpeg.

        Returns:
            None:
        """
        # https://ask.libreoffice.org/t/export-as-png-with-macro/74337/11
        # https://ask.libreoffice.org/t/how-to-export-cell-range-to-images/57828/2

        event_data: ImgExportT = {
            "image_type": "jpg",
            "compression": 9,
            "interlaced": 1,
            "translucent": 0,
            "pixel_width": 1192,
            "pixel_height": 1673,
            "logical_width": 27522,
            "logical_height": 38628,
            "file": fnm,
        }
        cargs = CancelEventArgsGeneric(source=self, event_data=event_data)
        self.trigger_event(CalcNamedEvent.RANGE_EXPORTING_IMAGE, cargs)
        if cargs.cancel and cargs.handled is False:
            raise mEx.CancelEventError(cargs)

        dv = PropertyState.DIRECT_VALUE
        # name, handle, value, state
        filter_data = (
            ("Compression", 0, cargs.event_data["compression"], dv),
            ("Interlaced", 0, cargs.event_data["interlaced"], dv),
            ("Translucent", 0, cargs.event_data["translucent"], dv),
            ("PixelWidth", 0, cargs.event_data["pixel_width"], dv),
            ("PixelHeight", 0, cargs.event_data["pixel_height"], dv),
            ("LogicalWidth", 0, cargs.event_data["logical_width"], dv),
            ("LogicalHeight", 0, cargs.event_data["logical_height"], dv),
        )
        self._export_as_img(fnm=event_data["file"], filter_name="calc_jpg_Export", filter_data=filter_data)
        self.trigger_event(CalcNamedEvent.RANGE_EXPORTED_IMAGE, EventArgsGeneric.from_args(cargs))

    def _export_as_png(self, fnm: PathOrStr) -> None:
        """
        Exports a range of cells as an jpeg.

        Returns:
            None:
        """
        event_data: ImgExportT = {
            "image_type": "png",
            "compression": 6,
            "interlaced": 1,
            "translucent": 1,
            "pixel_width": 256,
            "pixel_height": 194,
            "logical_width": 6772,
            "logical_height": 5132,
            "file": fnm,
        }
        cargs = CancelEventArgsGeneric(source=self, event_data=event_data)
        self.trigger_event(CalcNamedEvent.RANGE_EXPORTING_IMAGE, cargs)
        if cargs.cancel and cargs.handled is False:
            raise mEx.CancelEventError(cargs)

        dv = PropertyState.DIRECT_VALUE
        # name, handle, value, state
        filter_data = (
            ("Compression", 0, cargs.event_data["compression"], dv),
            ("Interlaced", 0, cargs.event_data["interlaced"], dv),
            ("Translucent", 0, cargs.event_data["translucent"], dv),
            ("PixelWidth", 0, cargs.event_data["pixel_width"], dv),
            ("PixelHeight", 0, cargs.event_data["pixel_height"], dv),
            ("LogicalWidth", 0, cargs.event_data["logical_width"], dv),
            ("LogicalHeight", 0, cargs.event_data["logical_height"], dv),
        )
        self._export_as_img(fnm=event_data["file"], filter_name="calc_png_Export", filter_data=filter_data)
        self.trigger_event(CalcNamedEvent.RANGE_EXPORTED_IMAGE, EventArgsGeneric.from_args(cargs))

    def _export_as_img(self, fnm: PathOrStr, filter_name: str, filter_data: tuple) -> None:
        url = mFile.FileIO.fnm_to_url(fnm=fnm)
        # capture the current selection.
        current_sel = self.calc_sheet.get_selected()
        self.select()

        args = mProps.Props.make_props(FilterName=filter_name, FilterData=filter_data, SelectionOnly=True)
        storable = self.calc_sheet.calc_doc.qi(XStorable, True)
        storable.storeToURL(url, args)  # save PNG
        # restore previous selection.
        current_sel.select()

    # endregion export

    # region Properties
    @property
    def calc_sheet(self) -> CalcSheet:
        """Sheet that owns this cell."""
        return self.__owner

    @property
    def range_obj(self) -> RangeObj:
        """Range object."""
        return self.__range_obj

    # endregion Properties
