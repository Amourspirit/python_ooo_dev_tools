# region imports
from __future__ import annotations
from typing import Any, cast, Iterable, Sequence, TYPE_CHECKING, Tuple
import contextlib
import uno  # pylint: disable=unused-import

# pylint: disable=useless-import-alias
from ooo.dyn.style.horizontal_alignment import HorizontalAlignment as HorizontalAlignment
from ooo.dyn.view.selection_type import SelectionType
from com.sun.star.awt.grid import XMutableGridDataModel

from ooodev.adapter.awt.grid.grid_selection_events import GridSelectionEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.loader import lo as mLo
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.utils.table_helper import TableHelper
from ooodev.utils.type_var import Table
from ooodev.adapter.awt.grid.uno_control_grid_model_partial import UnoControlGridModelPartial
from ooodev.dialog.dl_control.ctl_base import DialogControlBase


if TYPE_CHECKING:
    from com.sun.star.awt.grid import UnoControlGrid  # service
    from com.sun.star.awt.grid import UnoControlGridModel  # service
# endregion imports


class CtlGrid(DialogControlBase, UnoControlGridModelPartial, GridSelectionEvents):
    """Class for Grid Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlGrid) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlGrid): Grid Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlGridModelPartial.__init__(self)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        GridSelectionEvents.__init__(self, trigger_args=generic_args, cb=self._on_grid_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_grid_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addSelectionListener(self.events_listener_grid_selection)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlGrid:
        return cast("UnoControlGrid", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlGrid``"""
        return "com.sun.star.awt.UnoControlGrid"

    def get_model(self) -> UnoControlGridModel:
        """Gets the Model for the control"""
        return cast("UnoControlGridModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.GRID_CONTROL``"""
        return DialogControlKind.GRID_CONTROL

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.GRID_CONTROL``"""
        return DialogControlNamedKind.GRID_CONTROL

    # endregion Overrides

    # region Data
    def set_table_data(
        self,
        data: Table,
        *,
        widths: Sequence[int] | None = None,
        align: Iterable[HorizontalAlignment] | str | None = None,
        row_header_width: int = 10,
        has_colum_headers: bool | None = None,
        has_row_headers: bool | None = None,
    ) -> None:
        """
        Set the data in a table control. Preexisting data is cleared.

        Args:
            data (Table): 2D Sequence of data that is the data to set.
            widths (Sequence[int] | None, optional): Specifies Column Widths. If number of widths is less then the number of columns,
                the last width is used for the remaining columns.
                If omitted then each column is auto-sized to fill out the table width.
            align (Iterable[HorizontalAlignment] | str | None, optional): Specifies column alignments. See Note Below.
            row_header_width (int, optional): Specifies the width of the row header. Defaults to ``10``.
            has_colum_headers (bool | None, optional): Specifies if the data has a column header. If omitted the table's ShowColumnHeader property is used. Defaults to ``None``.
            has_row_headers (bool | None, optional): Specifies if the data has a row header. If omitted the table's ShowRowHeader property is used. Defaults to ``None``.

        Raises:
            ValueError: if not a valid UnoControlGrid or if no data model.

        Returns:
            None:

        Note:
            ``align`` can be a string of ``"L"``, ``"R"``, or ``"C"`` for left, right, or center alignment or
            a list of ``HorizontalAlignment`` values. If ``align`` values is lest then the number of columns,
            then the remaining columns will be aligned left.

            If ``has_colum_headers`` is ``True`` then the first row of data is used for the column headers.
            If ``table.Model.ShowColumnHeader`` is ``False``, then the column header row is not used.

            If ``has_colum_headers`` is ``False`` and ``table.Model.ShowColumnHeader`` is ``True``
            then the column headers are set to the default column names such as (A, B, C, D).

            If ``has_row_headers`` is ``True`` then the first row of data is used for the row headers.
            If ``table.Model.ShowRowHeader`` is ``False``, then the row header is not used.


            If ``has_row_headers`` is ``False`` and ``table.Model.ShowRowHeader`` is ``True``
            then the row headers are set to the default row names such as (1, 2, 3, 4).

        Example:
            .. code-block:: python

                # other code
                tab_sz = self._ctl_tab.getPosSize()
                ctl_table1 = Dialogs.insert_table_control(
                    dialog_ctrl=self._tab_table,
                    x=tab_sz.X + self._padding,
                    y=tab_sz.Y + self._padding,
                    width=tab_sz.Width - (self._padding * 2),
                    height=300,
                    grid_lines=True,
                    col_header=True,
                    row_header=True,
                )

                tbl = ... # get data as 2d sequence
                Dialogs.set_table_data(
                    table=ctl_table1,
                    data=tbl,
                    align="RLC", # first column right, second left, third center. All others left
                    widths=(75, 60, 100, 40), # does not need to add up to total width, a factor will be used to auto size where needed.
                    has_row_headers=True,
                    has_colum_headers=True,
                )
        """
        # set_table_data() will handle to many or to few widths
        # widths are applied by using a scale factor to the table width
        table = self.view

        tbl_size = table.getSize()

        model = cast("UnoControlGridModel", table.getModel())
        data_model = model.GridDataModel
        if not data_model:
            raise ValueError("No data model")
        data_model = mLo.Lo.qi(XMutableGridDataModel, data_model, True)

        # Erase any pre-existing data and columns
        data_model.removeAllRows()
        if data_model.ColumnCount > 0:
            # reverse indexes to start removing from the end
            for i in range(data_model.ColumnCount - 1, -1, -1):
                model.ColumnModel.removeColumn(i)

        # Get the headers from data
        use_col_headers = False
        use_row_headers = False
        if has_colum_headers is None:
            if model.ShowColumnHeader:
                use_col_headers = True
        elif has_colum_headers:
            use_col_headers = True

        if has_row_headers is None:
            if model.ShowRowHeader:
                use_row_headers = True
        elif has_row_headers:
            use_row_headers = True

        col_headers = data[0][1:] if has_row_headers else data[0]

        # Create the columns
        for i, header in enumerate(col_headers):
            column = model.ColumnModel.createColumn()
            if use_col_headers:
                column.Title = str(header)
            elif model.ShowColumnHeader:
                column.Title = TableHelper.make_column_name(i, zero_index=True)
            model.ColumnModel.addColumn(column)

        # Manage row headers width
        if has_row_headers and model.ShowRowHeader:
            header_width_row = row_header_width
            model.RowHeaderWidth = header_width_row
        else:
            header_width_row = 0

        # Size the columns. Column sizing cannot be done before all the columns are added
        len_col_headers = len(col_headers)
        len_widths = 0
        if widths:
            len_widths = len(widths)
            # Size the columns proportionally with their relative widths
            rel_width = 0.0
            # Compute the sum of the relative widths
            for i, width in enumerate(widths):
                if i + 1 >= len_col_headers:
                    break
                rel_width += width
            # if widths have less values then columns, add the rest with the last value of widths.
            if len_widths < len_col_headers:
                last_width = widths[-1]
                for i in range(len_widths, len_col_headers):
                    rel_width += last_width

            # Set absolute column widths
            # initial testing showed that columns are sized using this factor method even
            # if the factoring is not done here.
            if rel_width > 0:
                width_factor = (tbl_size.Width - header_width_row) / rel_width
            else:
                width_factor = 1.0

            for i, width in enumerate(widths):
                if i + 1 > len_col_headers:
                    break
                model.ColumnModel.getColumn(i).ColumnWidth = int(width * width_factor)
            # if widths have less values then columns, calculate the rest with the last value of widths.
            if len_widths < len_col_headers:
                last_width = widths[-1]
                for i in range(len_widths, len_col_headers):
                    model.ColumnModel.getColumn(i).ColumnWidth = int(last_width * width_factor)
        else:
            # Size header and columns evenly
            width = (tbl_size.Width - header_width_row) // len_col_headers
            for i in range(len_col_headers):
                model.ColumnModel.getColumn(i).ColumnWidth = width

        # Initialize the column alignment

        def get_align(s: str):
            s = s.lower()
            if s == "l":
                return HorizontalAlignment.LEFT
            elif s == "r":
                return HorizontalAlignment.RIGHT
            elif s == "c":
                return HorizontalAlignment.CENTER
            return HorizontalAlignment.LEFT

        if align:
            if isinstance(align, str):
                align = [get_align(s) for s in align.replace(" ", "")]
            elif not isinstance(align, list):
                align = list(align)
        else:
            align = [HorizontalAlignment.LEFT for _ in range(len_col_headers)]

        while len(align) > len_col_headers:
            _ = align.pop()

        while len(align) < len_col_headers:
            align.append(HorizontalAlignment.LEFT)

        # Feed the table with data
        # skip column headers row
        if use_col_headers is False:
            rng_start = 0
        else:
            rng_start = 1
        for i in range(rng_start, len(data)):
            row = data[i][1:] if use_row_headers else data[i]
            if not isinstance(row, tuple):
                row = tuple(row)
            row_header_text = ""
            if use_row_headers and model.ShowRowHeader:
                row_header_text = str(data[i][0])
            elif model.ShowRowHeader:
                if rng_start == 0:
                    row_header_text = str(i + 1)
                else:
                    row_header_text = str(i)
            data_model.addRow(row_header_text, row)

        for i, alignment in enumerate(align):
            model.ColumnModel.getColumn(i).HorizontalAlign = alignment  # type: ignore

    # endregion Data

    # region Properties
    @property
    def model(self) -> UnoControlGridModel:
        # pylint: disable=no-member
        return cast("UnoControlGridModel", super().model)

    @property
    def view(self) -> UnoControlGrid:
        # pylint: disable=no-member
        return cast("UnoControlGrid", super().view)

    @property
    def horizontal_scrollbar(self) -> bool:
        """
        Gets or sets if a horizontal scrollbar should be added to the dialog.
        Same as ``h_scroll`` property.
        """
        return self.h_scroll

    @horizontal_scrollbar.setter
    def horizontal_scrollbar(self, value: bool) -> None:
        """Sets the horizontal scrollbar"""
        self.h_scroll = value

    @property
    def vertical_scrollbar(self) -> bool:
        """
        Gets or sets if a vertical scrollbar should be added to the dialog.
        Same as ``v_scroll`` property.
        """
        return self.v_scroll

    @vertical_scrollbar.setter
    def vertical_scrollbar(self, value: bool) -> None:
        self.v_scroll = value

    @property
    def list_count(self) -> int:
        """Gets the number of items in the combo box"""
        with contextlib.suppress(Exception):
            return self.model.GridDataModel.RowCount
        return 0

    @property
    def list_index(self) -> int:
        """
        Gets which row index is selected in the gird.

        Returns:
            Index of the first selected row or ``-1`` if no rows are selected.
        """
        with contextlib.suppress(Exception):
            model = self.model
            if model.SelectionModel == SelectionType.SINGLE:
                return self.view.getCurrentRow()
            sel = cast(Tuple[int, ...], self.view.getSelectedRows())
            if sel:
                return sel[0]
        return -1

    # endregion Properties
