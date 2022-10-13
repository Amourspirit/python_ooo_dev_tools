# region Imports
from __future__ import annotations
from typing import List, Tuple, overload
from enum import Enum

from com.sun.star.chart import XChartDocument
from com.sun.star.container import XNameAccess
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.table import XTableChart
from com.sun.star.table import XTableCharts
from com.sun.star.table import XTableChartsSupplier
from com.sun.star.document import XEmbeddedObjectSupplier
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.chart import XDiagram
from com.sun.star.drawing import XDrawPageSupplier
from com.sun.star.drawing import XShape

from ooo.dyn.table.cell_range_address import CellRangeAddress
from ooo.dyn.awt.rectangle import Rectangle

from . import lo as mLo
from . import props as mProps
from ..exceptions import ex as mEx

# endregion Imports
class Chart:
    # private static final String CHART_CLASSID = "12dcae26-281f-416f-a234-c3086127382e";

    # hart label type
    # public static final int NO_LABEL = 0
    # public static final int SHOW_NUMBER = 1
    # public static final int SHOW_PERCENT = 2
    # public static final int SHOW_CATEGORY = 4
    # public static final int CHECK_LEGEND = 16

    class DiagramEnum(str, Enum):
        AREA_DIAGRAM = "AreaDiagram"
        BAR_DIAGRAM = "BarDiagram"
        BUBBLE_DIAGRAM = "BubbleDiagram"
        DONUT_DIAGRAM = "DonutDiagram"
        FILLED_NET_DIAGRAM = "FilledNetDiagram"
        LINE_DIAGRAM = "LineDiagram"
        NET_DIAGRAM = "NetDiagram"
        PIE_DIAGRAM = "PieDiagram"
        STOCK_DIAGRAM = "StockDiagram"
        XY_DIAGRAM = "XYDiagram"

        def __str__(self) -> str:
            return self._value_

    @overload
    @classmethod
    def insert_chart(
        cls,
        sheet: XSpreadsheet,
        chart_name: str,
        cells_range: CellRangeAddress,
        width: int,
        height: int,
        diagram_name: str | None,
    ) -> XChartDocument:
        ...

    @overload
    @classmethod
    def insert_chart(
        cls,
        sheet: XSpreadsheet,
        chart_name: str,
        cells_range: CellRangeAddress,
        width: int,
        height: int,
        diagram_name: str | None,
        x: int,
        y: int,
    ) -> XChartDocument:
        ...

    @classmethod
    def insert_chart(
        cls,
        sheet: XSpreadsheet,
        chart_name: str,
        cells_range: CellRangeAddress,
        width: int,
        height: int,
        diagram_name: str | None,
        x: int = 1,
        y: int = 1,
    ) -> XChartDocument:
        """
        Insert a chart using the given name as name of the OLE object and
        the range as corresponding
        range of data to be used for rendering.  The chart is placed in the sheet
        for charts at position (1,1) extending as large as given in chartSize.

        The diagram name must be the name of a diagram service (i.e one
        in "com.sun.star.chart.") that can be
        instantiated via the factory of the chart document

        Args:
            sheet (XSpreadsheet): Spreadsheet
            chart_name (str): Name of chart
            cells_range (CellRangeAddress): Cells range
            width (int): Width
            height (int): Height
            diagram_name (str | None): Diagram Name
            x (int, optional): Chart X position. Defaults to 1.
            y (int, optional): Chart y Position. Defaults to 1.

        Raises:
            ChartNoAccessError: If unable to get access to chart table.
            ChartExistingError: If chart already exist.

        Returns:
            XChartDocument: _description_
        """
        tbl_charts = cls.get_table_charts(sheet)
        tc_access = mLo.Lo.qi(XNameAccess, tbl_charts)
        if tc_access is None:
            raise mEx.ChartNoAccessError("Unable to get name access to chart table")

        if tc_access.hasByName(chart_name):
            raise mEx.ChartExistingError(f"A chart table called {chart_name} already exist")

        rect = Rectangle(x * 1000, y * 1000, width * 1000, height * 1000)

        addrs = [cells_range]
        # first boolean: has column headers?; second boolean: has row headers?
        tbl_charts.addNewByName(chart_name, rect, addrs, True, True)
        # 2nd last arg: whether the topmost row of the source data will be used
        # to set labels for the category axis or the legend.
        # last arg: whether the leftmost column of the source data will be
        # used to set labels for the category axis or the legend.

        tbl_chart = cls.get_table_chart(tc_access, chart_name)
        chart_doc = cls.get_chart_doc(tbl_chart)

        if diagram_name is not None:
            cls.set_chart_type(chart_doc, diagram_name)
        return chart_doc

    @classmethod
    def get_table_chart_access(cls, sheet: XSpreadsheet) -> XNameAccess:
        """
        Get XnameAccess interface of sheet charts

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Raises:
            MissingInterfaceError if ``XNameAccess`` interface is not available

        Returns:
            XNameAccess: Named Access interface
        """
        tbl_charts = cls.get_table_charts(sheet)
        return mLo.Lo.qi(XNameAccess, tbl_charts, True)

    @staticmethod
    def get_table_charts(sheet: XSpreadsheet) -> XTableCharts:
        """
        Gets Table Charts of a sheet

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Raises:
            MissingInterfaceError if access to ``XTableChartsSupplier`` interface is not available.

        Returns:
            XTableCharts: Table Charts interface
        """
        # get the supplier for the charts
        charts_supplier = mLo.Lo.qi(XTableChartsSupplier, sheet, True)
        return charts_supplier.getCharts()

    @classmethod
    def get_chart_names_list(cls, sheet: XSpreadsheet) -> Tuple[str]:
        """
        Gets Tuple of chart names as strings

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            Tuple[str]: Tuple of strings or empyty tuple if no charts are found
        """
        try:
            tbl_charts = cls.get_table_charts(sheet)
        except mEx.MissingInterfaceError:
            return ()
        return tbl_charts.getElementNames()

    @staticmethod
    def get_table_chart(tc_access: XNameAccess, chart_name: str) -> XTableChart:
        """
        Gets a table chart.

        Args:
            tc_access (XNameAccess): Name Access
            chart_name (str): Chart Name

        Raises:
            ChartError: If unable to get Table Chart

        Returns:
            XTableChart: Table Chart Interface.
        """
        try:
            tbl_chart = mLo.Lo.qi(XTableChart, tc_access.getByName(chart_name))
            return tbl_chart
        except Exception as e:
            raise mEx.ChartError(f"Could Not Access: {chart_name}") from e

    # region get_chart_doc()

    @staticmethod
    def _get_chart_doc_tbl_chart(tbl_chart: XTableChart) -> XChartDocument:
        """
        Gets chart document

        Args:
            tbl_chart (XTableChart): Table Chart

        Returns:
            XChartDocument: Chart Document Interface
        """
        eos = mLo.Lo.qi(XEmbeddedObjectSupplier, tbl_chart, True)
        intf = eos.getEmbeddedObject()
        return mLo.Lo.qi(XChartDocument, intf, True)

    @classmethod
    def _get_chart_doc_sheet_chart_name(cls, sheet: XSpreadsheet, chart_name: str) -> XChartDocument:
        try:
            tc_access = cls.get_table_chart_access(sheet)
        except mEx.MissingInterfaceError as e:
            raise mEx.ChartError("Unable to get name access to table chart") from e

        if not tc_access.hasByName(chart_name):
            raise mEx.ChartNotExistingError(f'No table chart called "{chart_name}" found')

        try:
            tbl_chart = cls.get_table_chart(tc_access=tc_access, chart_name=chart_name)
        except mEx.ChartError as e:
            raise mEx.ChartError("Unable to get name access to chart table") from e
        return cls._get_chart_doc_tbl_chart(tbl_chart)

    @overload
    @classmethod
    def get_chart_doc(cls, tbl_chart: XTableChart) -> XChartDocument:
        ...

    @overload
    @classmethod
    def get_chart_doc(cls, sheet: XSpreadsheet, chart_name: str) -> XChartDocument:
        ...

    @classmethod
    def get_chart_doc(cls, *args, **kwargs) -> XChartDocument:
        """
        Gets chart document

        Args:
            tbl_chart (XTableChart): Table Chart
            sheet: SpreadSheet
            chart_name (str): Chart Name

        Returns:
            XChartDocument: Chart Document Interface
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("tbl_chart", "sheet", "chart_name")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_chart_doc() got an unexpected keyword argument")
            keys = ("tbl_chart", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            ka[2] = ka.get("chart_name", None)
            return ka

        if not count in (1, 2):
            raise TypeError("get_chart_doc() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            return cls._get_chart_doc_tbl_chart(kargs[1])

        return cls._get_chart_doc_sheet_chart_name(kargs[1], kargs[2])

    # endregion get_chart_doc()

    @staticmethod
    def set_chart_type(chart_doc: XChartDocument, diagram_name: Chart.DiagramEnum | str) -> None:
        """
        Sets chart diagram name

        Args:
            chart_doc (XChartDocument): Chart Document
            diagram_name (DiagramEnum | str): Diagram Name

        Raises:
            ChartError: If unable to set diagram name
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, chart_doc)
            chart_type = f"com.sun.star.chart.{diagram_name}"
            diagram = mLo.Lo.qi(XDiagram, msf.createInstance(chart_type), True)
            chart_doc.setDiagram(diagram)
        except Exception as e:
            raise mEx.ChartError(f"Could not set the chart type to: {chart_type}") from e

    @staticmethod
    def get_chart_type(chart_doc: XChartDocument) -> str:
        """
        Gets chart document chart type

        Args:
            chart_doc (XChartDocument): Chart Document

        Returns:
            str: Chart Type
        """
        return chart_doc.getDiagram().getDiagramType()

    @classmethod
    def remove_chart(cls, sheet: XSpreadsheet, chart_name: str) -> bool:
        """
        Removes a chart from spreadsheet current charts

        Args:
            sheet (XSpreadsheet): Spreadsheet
            chart_name (str): Chart Name to remove

        Returns:
            bool: ``True`` if chart is found and removed; Otherwise, ``False``
        """
        try:
            tbl_charts = cls.get_table_charts(sheet)
            tc_access = mLo.Lo.qi(XNameAccess, tbl_charts, True)
        except mEx.MissingInterfaceError:
            mLo.Lo.print("Unable to get name access to chart table")
            return False

        if tc_access.hasByName(chart_name):
            tbl_charts.removeByName(chart_name)
            mLo.Lo.print(f'Chart table "{chart_name}" removed.')
            return True

        mLo.Lo.print(f'Chart table "{chart_name}" not found.')
        return False

    @staticmethod
    def set_visible(sheet: XSpreadsheet, is_visible: bool) -> None:
        """
        Set visible state of charts.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            is_visible (bool): Visible State. ``True`` shows charts, ``False`` hides charts.
        """
        # get draw page supplier for chart sheet
        page_supplier = mLo.Lo.qi(XDrawPageSupplier, sheet, True)
        draw_page = page_supplier.getDrawPage()
        num_shapes = draw_page.getCount()
        chart_class_id = str(mLo.Lo.CLSID.CHART)
        for i in range(num_shapes):
            try:
                shape = mLo.Lo.qi(XShape, draw_page.getByIndex(i))
                class_id = str(mProps.Props.get_property(obj=shape, name="CLSID"))

                if class_id.casefold() == chart_class_id:
                    mProps.Props.set_property(obj=shape, name="Visible", value=is_visible)
            except Exception:
                pass

    @staticmethod
    def get_chart_shape(sheet: XSpreadsheet) -> XShape | None:
        """
        Gets the shape of the first found chart

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            XShape | None: Shape if found; Otherwise, ``None``.
        """
        # get draw page supplier for chart sheet
        page_supplier = mLo.Lo.qi(XDrawPageSupplier, sheet, True)
        draw_page = page_supplier.getDrawPage()
        num_shapes = draw_page.getCount()
        chart_class_id = str(mLo.Lo.CLSID.CHART)
        shape = None
        for i in range(num_shapes):
            try:
                shape = mLo.Lo.qi(XShape, draw_page.getByIndex(i))
                class_id = str(mProps.Props.get_property(obj=shape, name="CLSID"))

                if class_id.casefold() == chart_class_id:
                    break
            except Exception:
                pass
        if shape is not None:
            mLo.Lo.print("Found a chart")
        return shape

    # region Chart inside Draw

    # work in progress. Need to create Draw class

    # endregion Chart inside Draw
