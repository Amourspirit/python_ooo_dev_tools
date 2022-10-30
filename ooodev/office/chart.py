# region Imports
from __future__ import annotations
from typing import Tuple, overload
from enum import Enum

import uno
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
from com.sun.star.drawing import XDrawPage
from com.sun.star.beans import XPropertySet
from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextContent
from com.sun.star.chart import XAxisXSupplier
from com.sun.star.chart import XAxisYSupplier
from com.sun.star.chart import XTwoAxisYSupplier
from com.sun.star.chart import XChartDataArray

from ..utils.kind.chart_diagram_kind import ChartDiagramKind as ChartDiagramKind
from ..utils.kind.drawing_shape_kind import DrawingShapeKind as DrawingShapeKind

from ooo.dyn.table.cell_range_address import CellRangeAddress as CellRangeAddress
from ooo.dyn.awt.rectangle import Rectangle
from ooo.dyn.text.vert_orientation import VertOrientation
from ooo.dyn.text.hori_orientation import HoriOrientation
from ooo.dyn.chart.chart_data_caption import ChartDataCaptionEnum as ChartDataCaptionEnum
from ooo.dyn.chart.data_label_placement import DataLabelPlacementEnum as DataLabelPlacementEnum
from ooo.dyn.chart.chart_solid_type import ChartSolidTypeEnum as ChartSolidTypeEnum
from ooo.dyn.chart.chart_symbol_type import ChartSymbolType as ChartSymbolType
from ooo.dyn.chart.chart_regression_curve_type import ChartRegressionCurveType as ChartRegressionCurveType

from ..utils import lo as mLo
from ..utils import props as mProps
from . import draw as mDraw
from ..utils import info as mInfo
from ..exceptions import ex as mEx
from ..utils.data_type.intensity import Intensity

# endregion Imports
class Chart:
    """Chart Class"""

    # region insert_chart()
    @overload
    @classmethod
    def insert_chart(
        cls, *, slide: XDrawPage, x: int, y: int, width: int, height: int, diagram_name: ChartDiagramKind | str
    ) -> XChartDocument:
        ...

    @overload
    @classmethod
    def insert_chart(
        cls, *, doc: XTextDocument, x: int, y: int, width: int, height: int, diagram_name: ChartDiagramKind | str
    ) -> XChartDocument:
        ...

    @overload
    @classmethod
    def insert_chart(
        cls,
        *,
        sheet: XSpreadsheet,
        chart_name: str,
        cells_range: CellRangeAddress,
        width: int,
        height: int,
        diagram_name: ChartDiagramKind | str,
    ) -> XChartDocument:
        ...

    @overload
    @classmethod
    def insert_chart(
        cls,
        *,
        sheet: XSpreadsheet,
        chart_name: str,
        cells_range: CellRangeAddress,
        x: int,
        y: int,
        width: int,
        height: int,
        diagram_name: ChartDiagramKind | str,
    ) -> XChartDocument:
        ...

    @classmethod
    def _insert_chart_slide(
        cls, slide: XDrawPage, x: int, y: int, width: int, height: int, diagram_name: ChartDiagramKind | str
    ) -> XChartDocument:
        try:
            ashape = mDraw.Draw.add_shape(
                slide=slide,
                shape_type=DrawingShapeKind.OLE2_SHAPE,
                x=x,
                y=y,
                width=width,
                height=height,
            )
            chart_doc = cls._get_chart_doc_shape(ashape)
            cls.set_chart_type(chart_doc=chart_doc, diagram_name=diagram_name)
            return chart_doc
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error inserting chart slide") from e

    @classmethod
    def _insert_chart_sheet(
        cls,
        sheet: XSpreadsheet,
        chart_name: str,
        cells_range: CellRangeAddress,
        x: int,
        y: int,
        width: int,
        height: int,
        diagram_name: ChartDiagramKind | str,
    ) -> XChartDocument:
        try:
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
        except mEx.ChartNoAccessError:
            raise
        except mEx.ChartExistingError:
            raise
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError(f'Error inserting "{chart_name}" into sheet') from e

    @classmethod
    def _insert_chart_doc(
        cls, doc: XTextDocument, x: int, y: int, width: int, height: int, diagram_name: ChartDiagramKind | str
    ) -> XChartDocument:
        try:
            tc = mLo.Lo.create_instance_msf(XTextContent, "com.sun.star.text.TextEmbeddedObject", raise_err=True)
            ps = mLo.Lo.qi(XPropertySet, tc, True)
            ps.setPropertyValue("CLSID", str(mLo.Lo.CLSID.CHART))

            xtext = doc.getText()
            cursor = xtext.createTextCursor()

            # insert embedded object in text -> object will be created
            xtext.insertTextContent(cursor, tc, True)

            # set size and position
            shape = mLo.Lo.qi(XShape, tc, True)
            shape.setSize(mDraw.Draw.Size(width * 1_000, height * 1_000))

            ps.setPropertyValue("VertOrient", VertOrientation.NONE)
            ps.setPropertyValue("HoriOrient", HoriOrientation.NONE)
            ps.setPropertyValue("VertOrientPosition", y * 1_000)
            ps.setPropertyValue("HoriOrientPosition", x * 1_000)

            # retrieve the chart document as model of the OLE shape
            chart_doc = mLo.Lo.qi(XChartDocument, ps.getPropertyValue("Model"))
            if chart_doc is not None:
                cls.set_chart_type(chart_doc=chart_doc, diagram_name=diagram_name)
        except Exception as e:
            raise mEx.ChartError("Unable to insert chart") from e

    @classmethod
    def insert_chart(cls, *args, **kwargs) -> XChartDocument:
        """
        Insert a chart using the given name as name of the OLE object and
        the range as corresponding range of data to be used for rendering.
        The chart is placed in the sheet for charts at position (1,1)
        extending as large as given in chart size.

        The diagram name must be the name of a diagram service (i.e one
        in ``com.sun.star.chart``) that can be
        instantiated via the factory of the chart document

        Args:
            slide (XDrawPage): Draw page
            doc (XTextDocument): Document
            sheet (XSpreadsheet): Spreadsheet
            chart_name (str): Name of chart
            cells_range (CellRangeAddress): Cells range
            width (int): Width
            height (int): Height
            diagram_name (DiagramKind | str): Diagram Name
            x (int, optional): Chart X position. Defaults to 1.
            y (int, optional): Chart y Position. Defaults to 1.

        Raises:
            ChartNoAccessError: If unable to get access to chart table.
            ChartExistingError: If chart already exist.
            ChartError: If unable to insert chart.

        Returns:
            XChartDocument: Chart Document
        """
        if len(args) > 0:
            raise TypeError(f"Insert_chart() takes 0 positional arguments but {len(args)} was given")

        count = len(kwargs)

        if not count in (6, 8):
            raise TypeError("insert_chart() got an invalid number of arguments")

        if count == 6:
            if "slide" in kwargs:
                # insert_chart( slide: XDrawPage, x: int, y: int, width: int, height: int, diagram_name: ChartDiagramKind | str) -> XChartDocument:
                return cls._insert_chart_slide(**kwargs)
            if "doc" in kwargs:
                # insert_chart(doc: XTextDocument, x: int, y: int, width: int, height: int, diagram_name: ChartDiagramKind | str) -> XChartDocument:
                return cls._insert_chart_doc(**kwargs)
            # else
            # insert_chart(sheet: XSpreadsheet, chart_name: str, cells_range: CellRangeAddress, width: int, height: int, diagram_name: ChartDiagramKind | str,) -> XChartDocument:
            kargs = kwargs.copy()
            kargs["x"] = 1
            kargs["y"] = 1
            return cls._insert_chart_sheet(**kargs)

        return cls._insert_chart_sheet(**kwargs)

    # endregion insert_chart()

    @classmethod
    def get_table_chart_access(cls, sheet: XSpreadsheet) -> XNameAccess:
        """
        Get XnameAccess interface of sheet charts

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Raises:
            ChartError: if error occurs.

        Returns:
            XNameAccess: Named Access interface
        """
        try:
            tbl_charts = cls.get_table_charts(sheet)
            return mLo.Lo.qi(XNameAccess, tbl_charts, True)
        except Exception as e:
            raise mEx.ChartError("Error getting table chart access") from e

    @staticmethod
    def get_table_charts(sheet: XSpreadsheet) -> XTableCharts:
        """
        Gets Table Charts of a sheet

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Raises:
            ChartError: if error occurs.

        Returns:
            XTableCharts: Table Charts interface
        """
        try:
            # get the supplier for the charts
            charts_supplier = mLo.Lo.qi(XTableChartsSupplier, sheet, True)
            return charts_supplier.getCharts()
        except Exception as e:
            raise mEx.ChartError("Error getting table chart access") from e

    @classmethod
    def get_chart_names_list(cls, sheet: XSpreadsheet) -> Tuple[str, ...]:
        """
        Gets Tuple of chart names as strings

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            Tuple[str]: Tuple of strings or empty tuple if no charts are found
        """
        try:
            tbl_charts = cls.get_table_charts(sheet)
        except mEx.ChartError:
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
            tbl_chart = mLo.Lo.qi(XTableChart, tc_access.getByName(chart_name), True)
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
        try:
            eos = mLo.Lo.qi(XEmbeddedObjectSupplier, tbl_chart, True)
            intf = eos.getEmbeddedObject()
            return mLo.Lo.qi(XChartDocument, intf, True)
        except Exception as e:
            raise mEx.ChartError("Unable to get chart.") from e

    @classmethod
    def _get_chart_doc_sheet_chart_name(cls, sheet: XSpreadsheet, chart_name: str) -> XChartDocument:
        tc_access = cls.get_table_chart_access(sheet)

        if not tc_access.hasByName(chart_name):
            raise mEx.ChartNotExistingError(f'No table chart called "{chart_name}" found')

        tbl_chart = cls.get_table_chart(tc_access=tc_access, chart_name=chart_name)
        return cls._get_chart_doc_tbl_chart(tbl_chart)

    @staticmethod
    def _get_chart_doc_shape(shape: XShape) -> XChartDocument:
        try:
            # change the LOL shape into a chart
            try:
                shape_props = mLo.Lo.qi(XPropertySet, shape, True)
            except mEx.MissingInterfaceError as ex:
                raise mEx.ChartError("Unable to access shape properties") from ex
            # set the class id for charts
            # shape_props.setPropertyValue("CLSID", str(mLo.Lo.CLSID.CHART))

            # retrieve the chart document as model of the OLE shape
            chart_doc = mLo.Lo.qi(XChartDocument, shape_props.getPropertyValue("Model"), True)
            return chart_doc
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Couldn't change the OLE shape into a chart.") from e

    @overload
    @classmethod
    def get_chart_doc(cls, shape: XShape) -> XChartDocument:
        ...

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
            shape (XShape): Shape object
            tbl_chart (XTableChart): Table Chart
            sheet: SpreadSheet
            chart_name (str): Chart Name

        Raises:
            ChartNotExistingError: If chart is not existing.
            ChartError: If chart error occurs.

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
            valid_keys = ("tbl_chart", "shape", "sheet", "chart_name")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_chart_doc() got an unexpected keyword argument")
            keys = ("tbl_chart", "sheet", "shape")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            ka[2] = kwargs.get("chart_name", None)
            return ka

        if not count in (1, 2):
            raise TypeError("get_chart_doc() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            # def get_chart_doc(cls, shape: XShape)
            # def get_chart_doc(cls, tbl_chart: XTableChart)
            shape = mLo.Lo.qi(XShape, kargs[1])
            if shape is None:
                return cls._get_chart_doc_tbl_chart(kargs[1])
            else:
                return cls._get_chart_doc_shape(shape)

        return cls._get_chart_doc_sheet_chart_name(kargs[1], kargs[2])

    # endregion get_chart_doc()

    @staticmethod
    def set_chart_type(chart_doc: XChartDocument, diagram_name: ChartDiagramKind | str) -> None:
        """
        Sets chart diagram name

        Args:
            chart_doc (XChartDocument): Chart Document
            diagram_name (DiagramEnum | str): Diagram Name

        Raises:
            ChartError: If error occurs.
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, chart_doc, True)
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
        try:
            result = chart_doc.getDiagram().getDiagramType()
            if result is None:
                raise mEx.ChartError("Error, result of get chart type is None")
            return result
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error getting chart type")

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
        mLo.Lo.print(f'Removing chart "{chart_name}"')
        try:
            tbl_charts = cls.get_table_charts(sheet)
            tc_access = mLo.Lo.qi(XNameAccess, tbl_charts, True)
        except mEx.ChartError as ce:
            mLo.Lo.print(f'Error get table charts for "{chart_name}"')
            mLo.Lo.print(f"  {ce}")
            return False
        except mEx.MissingInterfaceError as e:
            mLo.Lo.print("Unable to get name access to chart table")
            mLo.Lo.print(f"  {ce}")
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

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        Note:
            Any chart on sheet errors when attempting to set visibility then then it is skipped
            and no error is raised.
        """

        # get draw page supplier for chart sheet
        try:
            page_supplier = mLo.Lo.qi(XDrawPageSupplier, sheet, True)
            draw_page = page_supplier.getDrawPage()
            num_shapes = draw_page.getCount()
            chart_class_id = str(mLo.Lo.CLSID.CHART)
        except Exception as e:
            raise mEx.ChartError("Error setting visibility of sheet") from e

        for i in range(num_shapes):
            try:
                shape = mLo.Lo.qi(XShape, draw_page.getByIndex(i), True)
                class_id = str(mProps.Props.get(shape, "CLSID"))

                if class_id.casefold() == chart_class_id:
                    mProps.Props.set(shape, Visible=is_visible)
            except Exception as e:
                mLo.Lo.print("Error setting visibility of chart")
                mLo.Lo.print(f"  {e}")

    @staticmethod
    def get_chart_shape(sheet: XSpreadsheet) -> XShape:
        """
        Gets the shape of the first found chart

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Raises:
            ShapeMissingError: If unable to find a chart on the sheet.
            ShapeError: If any other error occurs.

        Returns:
            XShape: Chart Shape.
        """
        try:
            # get draw page supplier for chart sheet
            page_supplier = mLo.Lo.qi(XDrawPageSupplier, sheet, True)
            draw_page = page_supplier.getDrawPage()
            num_shapes = draw_page.getCount()
            chart_class_id = str(mLo.Lo.CLSID.CHART)
            shape = None
            for i in range(num_shapes):
                try:
                    shape = mLo.Lo.qi(XShape, draw_page.getByIndex(i), True)
                    class_id = str(mProps.Props.get(shape, "CLSID"))

                    if class_id.casefold() == chart_class_id:
                        break
                except Exception:
                    pass
            if shape is None:
                raise mEx.ShapeMissingError("Unable to find any chart on sheet")
            else:
                mLo.Lo.print("Found a chart")
            return shape
        except mEx.ShapeMissingError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Error getting chart shape") from e

    # region adjust properties
    @staticmethod
    def get_title(chart_doc: XChartDocument) -> str:
        """
        Gets title for a chart document.

        Args:
            chart_doc (XChartDocument): Chart document

        Raises:
            ChartError: If error occurs.

        Returns:
            str: Chart Title
        """
        try:
            title = mProps.Props.get(chart_doc.getTitle(), "String")
            if title is None:
                raise mEx.ChartError("Error getting title for chart document. Property returned None")
            return str(title)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error geting chart document title") from e

    @staticmethod
    def set_title(chart_doc: XChartDocument, title: str) -> None:
        """
        Set title for a chart document

        Args:
            chart_doc (XChartDocument): Chart document
            title (str): Title text

        Raises:
            ChartError: If error occurs

        Returns:
            None:
        """
        try:
            shape = chart_doc.getTitle()
            mProps.Props.set(shape, String=title)
        except Exception as e:
            raise mEx.ChartError("Error setting title for chart document.") from e

    @staticmethod
    def get_sub_title(chart_doc: XChartDocument) -> str:
        """
        Gets subtitle for a chart document.

        Args:
            chart_doc (XChartDocument): Chart document

        Raises:
            ChartError: If error occurs.

        Returns:
            str: Chart Subtitle
        """
        try:
            sub_title = mProps.Props.get(chart_doc.getSubTitle(), "String")
            if sub_title is None:
                raise mEx.ChartError("Error getting subtitle for chart document. Property returned None")
            return str(sub_title)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error geting chart document subtitle") from e

    @staticmethod
    def set_subtitle(chart_doc: XChartDocument, subtitle: str) -> None:
        """
        Set subtitle for a chart document

        Args:
            chart_doc (XChartDocument): Chart document
            subtitle (str): Subtitle text

        Raises:
            ChartError: If error occurs

        Returns:
            None:
        """
        try:
            # in java this next line was in error and should have been getSubTitle()
            # chartDoc.getTitle();
            shape = chart_doc.getSubTitle()
            mProps.Props.set(shape, String=subtitle)
            # mProps.Props.set_property(shape, "HasSubTitle", True)
        except Exception as e:
            raise mEx.ChartError("Error setting title for chart document.") from e

    @staticmethod
    def set_x_axis_title(chart_doc: XChartDocument, title: str) -> None:
        """
        Sets the X Axis Title

        Args:
            chart_doc (XChartDocument): Chart Document
            title (str): Title text

        Raises:
            DiagramNotExistingError: If ``chart_doc`` diagram does not exist
            ServiceNotSupported: If ``com.sun.star.chart.ChartAxisXSupplier`` service is not supported by diagram.
            ChartError: If any other error occurs.

        Returns:
            None:
        """
        try:
            diagram = chart_doc.getDiagram()
            if diagram is None:
                raise mEx.DiagramNotExistingError("Diagram does not exist")

            if not mInfo.Info.support_service(diagram, "com.sun.star.chart.ChartAxisXSupplier"):
                raise mEx.ServiceNotSupported("com.sun.star.chart.ChartAxisXSupplier")

            mProps.Props.set(diagram, HasXAxisDescription=True)
            axis = mLo.Lo.qi(XAxisXSupplier, diagram, True)
            title_shape = axis.getXAxisTitle()
            mProps.Props.set_property(title_shape, "String", title)
            # mProps.Props.set_property(title_shape, "HasXAxisTitle", True)
        except mEx.DiagramNotExistingError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting title for X Axis of chart document.") from e

    @staticmethod
    def set_y_axis_title(chart_doc: XChartDocument, title: str) -> None:
        """
        Sets the Y Axis Title

        Args:
            chart_doc (XChartDocument): Chart Document
            title (str): Title text

        Raises:
            DiagramNotExistingError: If ``chart_doc`` diagram does not exist
            ServiceNotSupported: If ``com.sun.star.chart.ChartAxisYSupplier`` service is not supported by diagram.
            ChartError: If any other error occurs.

        Returns:
            None:
        """
        try:
            diagram = chart_doc.getDiagram()
            if diagram is None:
                raise mEx.DiagramNotExistingError("Diagram does not exist")

            if not mInfo.Info.support_service(diagram, "com.sun.star.chart.ChartAxisYSupplier"):
                raise mEx.ServiceNotSupported("com.sun.star.chart.ChartAxisYSupplier")

            mProps.Props.set(diagram, HasYAxisDescription=True)
            axis = mLo.Lo.qi(XAxisYSupplier, diagram, True)
            title_shape = axis.getYAxisTitle()
            mProps.Props.set(title_shape, String=title)
            # mProps.Props.set(title_shape, HasYAxisTitle=True)
        except mEx.DiagramNotExistingError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting title for Y Axis of chart document.") from e

    @staticmethod
    def set_second_y_axis(chart_doc: XChartDocument, title: str) -> None:
        """
        Sets the second Y Axis Title

        Args:
            chart_doc (XChartDocument): Chart Document
            title (str): Title text

        Raises:
            DiagramNotExistingError: If ``chart_doc`` diagram does not exist
            ServiceNotSupported: If ``com.sun.star.chart.ChartTwoAxisYSupplier`` service is not supported by diagram.
            ChartError: If any other error occurs.

        Returns:
            None:
        """
        try:
            diagram = chart_doc.getDiagram()
            if diagram is None:
                raise mEx.DiagramNotExistingError("Diagram does not exist")

            if not mInfo.Info.support_service(diagram, "com.sun.star.chart.ChartTwoAxisYSupplier"):
                raise mEx.ServiceNotSupported("com.sun.star.chart.ChartTwoAxisYSupplier")

            mProps.Props.set(diagram, HasSecondaryYAxis=True)
            axis = mLo.Lo.qi(XTwoAxisYSupplier, diagram, True)

            # y2_ps = axis.getSecondaryYAxis()
            # mProps.Props.show_props("Second y-axis", y2_ps)

            y2_title_ps = mLo.Lo.qi(XPropertySet, axis.getYAxisTitle(), True)
            mProps.Props.set(y2_title_ps, String=title)
            # mProps.Props.show_props("Second y-axis title", y2_title_ps)

        except mEx.DiagramNotExistingError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting second Y Axis of chart document.") from e

    @staticmethod
    def view_legend(chart_doc: XChartDocument, is_visible: bool) -> None:
        """
        Set the visibility of chart document legend

        Args:
            chart_doc (XChartDocument): Chart Document
            is_visible (bool): Visibility

        Raises:
            ServiceNotSupported: If ``com.sun.star.chart.ChartDocument`` service is not supported.
            ChartError: If any other error occurs.

        Returns:
            None:
        """
        # this was showLegend in java
        if not mInfo.Info.support_service(chart_doc, "com.sun.star.chart.ChartDocument"):
            raise mEx.ServiceNotSupported("com.sun.star.chart.ChartDocument")

        try:

            mProps.Props.set_property(chart_doc, "HasLegend", is_visible)
        except Exception as e:
            raise mEx.ChartError(f'Error setting chart document legend visibility to "{is_visible}"') from e

    @staticmethod
    def set_use_horizontals(chart_doc: XChartDocument, is_vertical: bool = True) -> None:
        """
        Determines if the bars of a Bar Diagram chart are drawn vertically or horizontally.
        Default is vertical.

        If ``is_vertical`` is ``False`` you get a column chart rather than a bar chart.

        Args:
            chart_doc (XChartDocument): Bar Diagram chart document
            is_vertical (bool, optional): Determines if chart is drawn vertically or horizontally. Default ``True``.

        Raises:
            DiagramNotExistingError: If ``chart_doc`` diagram does not exist
            ServiceNotSupported: If ``com.sun.star.chart.BarDiagram`` service is not supported by diagram.
            ChartError: If any other error occurs.

        Returns:
            None:
        """
        # was useHorizontals in java
        try:
            diagram = chart_doc.getDiagram()
            if diagram is None:
                raise mEx.DiagramNotExistingError("Diagram does not exist")

            if not mInfo.Info.support_service(diagram, "com.sun.star.chart.BarDiagram"):
                raise mEx.ServiceNotSupported("com.sun.star.chart.BarDiagram")

            mProps.Props.set_property(diagram, "Vertical", is_vertical)

        except mEx.DiagramNotExistingError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting use horizontals of chart document.") from e

    @staticmethod
    def set_use_lines(chart_doc: XChartDocument, has_lines: bool) -> None:
        """
        Sets if the chart type has lines connecting the data points or contains just symbols.

        Args:
            chart_doc (XChartDocument): ``LineDiagram`` or ``XYDiagram`` chart document.
            has_lines (bool): Determines if the chart type has lines connecting the data points or contains just symbols.

        Raises:
            DiagramNotExistingError: If ``chart_doc`` diagram does not exist
            ServiceNotSupported: If chart document is not a ``LineDiagram`` or ``XYDiagram``
            ChartError: If any other error occurs.

        Returns:
            None:
        """
        # usesLines in java
        try:
            # com.sun.star.chart.XYDiagram inherits from LineDiagram
            diagram = chart_doc.getDiagram()
            if diagram is None:
                raise mEx.DiagramNotExistingError("Diagram does not exist")

            if not mInfo.Info.support_service(
                diagram, "com.sun.star.chart.LineDiagram", "com.sun.star.chart.XYDiagram"
            ):
                raise mEx.ServiceNotSupported("com.sun.star.chart.LineDiagram", "com.sun.star.chart.XYDiagram")

            mProps.Props.set(diagram, Lines=has_lines)

        except mEx.ServiceNotSupported:
            raise
        except mEx.DiagramNotExistingError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting use lines chart document.") from e

    @staticmethod
    def set_data_caption(chart_doc: XChartDocument, label_types: ChartDataCaptionEnum) -> None:
        """
        Sets how the caption of data points is displayed.

        Args:
            chart_doc (XChartDocument): Chart Document
            label_types (ChartDataCaptionEnum): Flags, specifies how the caption of data points is displayed.

        Raises:
            DiagramNotExistingError: If ``chart_doc`` diagram does not exist
            ServiceNotSupported: If ``com.sun.star.chart.Diagram`` service is not supported by diagram.
            ChartError: If any other error occurs.

        Returns:
            None:
        """
        # ChartDataCaptionEnum is IntFlags enum.
        try:
            diagram = chart_doc.getDiagram()
            if diagram is None:
                raise mEx.DiagramNotExistingError("Diagram does not exist")

            if not mInfo.Info.support_service(diagram, "com.sun.star.chart.Diagram"):
                raise mEx.ServiceNotSupported("com.sun.star.chart.Diagram")

            mProps.Props.set(diagram, DataCaption=int(label_types))

        except mEx.ServiceNotSupported:
            raise
        except mEx.DiagramNotExistingError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting data caption of chart document.") from e

    @staticmethod
    def set_data_placement(chart_doc: XChartDocument, placement: DataLabelPlacementEnum) -> None:
        """
        Sets the relative position for the data label.

        Args:
            chart_doc (XChartDocument): Chart Document.
            placement (DataLabelPlacementEnum): Specifies a relative position for the data label.

        Raises:
            DiagramNotExistingError: If ``chart_doc`` diagram does not exist
            ServiceNotSupported: If ``com.sun.star.chart.ChartDataPointProperties`` service is not supported by diagram.
            ChartError: If any other error occurs.

        Returns:
            None:
        """
        try:
            diagram = chart_doc.getDiagram()
            if diagram is None:
                raise mEx.DiagramNotExistingError("Diagram does not exist")

            if not mInfo.Info.support_service(diagram, "com.sun.star.chart.ChartDataPointProperties"):
                raise mEx.ServiceNotSupported("com.sun.star.chart.ChartDataPointProperties")

            data_arr = mLo.Lo.qi(XChartDataArray, chart_doc.getData(), True)
            adata = data_arr.getData()
            num_points = len(adata)
            for i in range(num_points):
                # first parameter is the index of the point, the second one is the series
                point_props = diagram.getDataPointProperties(i, 0)

                # point_props.setPropertyValue("CharHeight", 14.0)
                # point_props.setPropertyValue("CharWeight", FontWeight.BOLD)
                # point_props.setPropertyValue("CharColor", 0x993366)
                point_props.setPropertyValue("LabelPlacement", int(placement))

            if mInfo.Info.support_service(diagram, "com.sun.star.chart.PieDiagram"):
                # for pie chart
                ps = diagram.getDataPointProperties(0, 0)
            else:
                ps = diagram.getDataRowProperties(0)
            mProps.Props.set(ps, LabelPlacement=int(placement))
            mProps.Props.show_props("Data Row", ps)

        except mEx.ServiceNotSupported:
            raise
        except mEx.DiagramNotExistingError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting data placement of chart document.") from e

    @staticmethod
    def show_data_array(chart_doc: XChartDocument) -> None:
        """
        Prints data array info to console.

        Args:
            chart_doc (XChartDocument): Chart Document.
        """
        chart_data = chart_doc.getData()
        data_arr = mLo.Lo.qi(XChartDataArray, chart_data, True)
        data = data_arr.getData()
        if not data:
            print("No data found")
        else:
            print(f"No. of Data columns: {len(data[0])}")
            print(f"No. of Data rows: {len(data)}")
            for row in data:
                for col in row:
                    print(f"  {col}")
        print()
        row_descs = data_arr.getRowDescriptions()
        if not row_descs:
            print("No row description found")
        else:
            print(f"No. of rows: {len(row_descs)}")
            for row in row_descs:
                print(f'  "{row}"')
        print()

        col_descs = data_arr.getColumnDescriptions()
        if not col_descs:
            print("No column description found")
        else:
            print(f"No. of columns: {len(row_descs)}")
            for col in col_descs:
                print(f'  "{col}"')
        print()

    @staticmethod
    def set_area_transparency(chart_doc: XChartDocument, val: Intensity) -> None:
        """
        Set the transparency of chart areas surrounding the diagram/

        Args:
            chart_doc (XChartDocument): Chart Document.
            val (Intensity): Transparency intensity, Higher values is more transparent.

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            mProps.Props.set(chart_doc.getArea(), FillTransparence=val.Value)
        except Exception as e:
            raise mEx.ChartError("Error setting chart transparency") from e

    # region set_3d()
    @overload
    @staticmethod
    def set_3d(chart_doc: XChartDocument, is_3d: bool) -> None:
        ...

    @overload
    @staticmethod
    def set_3d(chart_doc: XChartDocument, is_3d: bool, solid_type: ChartSolidTypeEnum) -> None:
        ...

    @staticmethod
    def set_3d(chart_doc: XChartDocument, is_3d: bool, solid_type: ChartSolidTypeEnum | None = None) -> None:
        """
        Sets chart 3d option.

        Args:
            chart_doc (XChartDocument): Chart Document.
            is_3d (bool): Specifies if chart is 3d.
            solid_type (ChartSolidTypeEnum): Chart Solid Type. Defaults to ``ChartSolidTypeEnum.CYLINDER``.

        Raises:
            DiagramNotExistingError: If ``chart_doc`` diagram does not exist
            ChartError: If any other error occurs.

        Returns:
            None:
        """
        try:
            diagram = chart_doc.getDiagram()
            if diagram is None:
                raise mEx.DiagramNotExistingError("Diagram does not exist")

            is_vert = bool(mProps.Props.get(diagram, "Vertical"))
            if solid_type is None:
                solid_type = ChartSolidTypeEnum.CYLINDER
            mProps.Props.set(diagram, Dim3D=is_3d, SolidType=int(solid_type), Vertical=is_vert)

        except mEx.DiagramNotExistingError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting chart to 3d.") from e

    # endregion set_3d()

    @staticmethod
    def set_symbol(chart_doc: XChartDocument, has_symbol: bool) -> None:
        """
        Sets chart symbol.

        Args:
            chart_doc (XChartDocument): Chart Document.
            has_symbol (bool): Specifies if chart is 3d.

        Raises:
            DiagramNotExistingError: If ``chart_doc`` diagram does not exist
            ServiceNotSupported: If chart document is not a ``LineDiagram`` or ``XYDiagram``
            ChartError: If any other error occurs.

        Returns:
            None:
        """
        # In this interface, only the two values ChartSymbolType::NONE and
        # ChartSymbolType::AUTO are supported. Later versions may support the selection of the symbols shape.

        # XYDiagram inherits LineDiagram
        try:
            diagram = chart_doc.getDiagram()
            if diagram is None:
                raise mEx.DiagramNotExistingError("Diagram does not exist")

            if not mInfo.Info.support_service(
                diagram, "com.sun.star.chart.LineDiagram", "com.sun.star.chart.XYDiagram"
            ):
                raise mEx.ServiceNotSupported("com.sun.star.chart.LineDiagram", "com.sun.star.chart.XYDiagram")

            if has_symbol:
                mProps.Props.set(diagram, SymbolType=ChartSymbolType.AUTO)
            else:
                mProps.Props.set(diagram, SymbolType=ChartSymbolType.NONE)

        except mEx.ServiceNotSupported:
            raise
        except mEx.DiagramNotExistingError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting chart symbol.") from e

    @staticmethod
    def set_trend(chart_doc: XChartDocument, curve_type: ChartRegressionCurveType) -> None:
        """
        Sets chart trend

        Args:
            chart_doc (XChartDocument): Chart Document.
            curve_type (ChartRegressionCurveType): determines a type of regression for the data row values.

        Raises:
            DiagramNotExistingError: If ``chart_doc`` diagram does not exist
            ServiceNotSupported: If ``com.sun.star.chart.ChartStatistics`` service is not supported by diagram.
            ChartError: If any other error occurs.

        Returns:
            None:
        """
        # determines a type of regression for the data row values.
        try:
            diagram = chart_doc.getDiagram()
            if diagram is None:
                raise mEx.DiagramNotExistingError("Diagram does not exist")

            if not mInfo.Info.support_service(diagram, "com.sun.star.chart.ChartStatistics"):
                raise mEx.ServiceNotSupported("com.sun.star.chart.ChartStatistics")

            # data_arr = mLo.Lo.qi(XChartDataArray, chart_doc.getData(), True)
            # data = data_arr.getData()
            # num_points = len(data)

            if (
                curve_type == ChartRegressionCurveType.NONE
                or curve_type == ChartRegressionCurveType.LINEAR
                or curve_type == ChartRegressionCurveType.LOGARITHM
                or curve_type == ChartRegressionCurveType.EXPONENTIAL
                or curve_type == ChartRegressionCurveType.POWER
            ):
                mProps.Props.set(diagram, RegressionCurves=curve_type)
            else:
                mLo.Lo.print(f"Did not recognize curve type: {curve_type}")

        except mEx.ServiceNotSupported:
            raise
        except mEx.DiagramNotExistingError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting chart trend.") from e

    @staticmethod
    def set_data_row_descriptions(chart_doc: XChartDocument, *descs: str) -> None:
        """
        Set data row descriptions for chart.

        Args:
            chart_doc (XChartDocument): Chart Document.
            *descs (str): Variable length argument list of descriptions. Must match the same number of descriptions in the chart.

        Raises:
            ValueError: If wrong number of descriptions are provided.
            ChartError: If any other error occurs

        Returns:
            None:

        Note:
            This method works, but affects the graphic data point position.
        """
        try:
            chart_data = chart_doc.getData()
            data_arr = mLo.Lo.qi(XChartDataArray, chart_data, True)

            row_descs = data_arr.getRowDescriptions()
            if not row_descs:
                mLo.Lo.print("No row description found")
                return

            if len(row_descs) != len(descs):
                raise ValueError(f"Row length mismatch; No. of rows == {len(row_descs)}")
            data_arr.setRowDescriptions(descs)
        except ValueError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting chart row descriptions.") from e

    @staticmethod
    def show_data_row_props(chart_doc: XChartDocument) -> None:
        """
        Prints data row info to console.

        Args:
            chart_doc (XChartDocument): Chart Document.
        """
        try:
            diagram = chart_doc.getDiagram()
            if diagram is None:
                raise mEx.DiagramNotExistingError("Diagram does not exist")

            data_arr = mLo.Lo.qi(XChartDataArray, chart_doc.getData(), True)
            data = data_arr.getData()
            num_points = len(data)
            print(f"No. of points: {num_points}")

            for i in range(num_points):
                point_ps = diagram.getDataPointProperties(i, 0)
                mProps.Props.show_props(f"{i}. Data Point", point_ps)

            if mInfo.Info.support_service(diagram, "com.sun.star.chart.PieDiagram"):
                # for pie chart
                ps = diagram.getDataPointProperties(0, 0)
            else:
                ps = diagram.getDataRowProperties(0)

            mProps.Props.show_props("Data Row", ps)
        except Exception as e:
            print("Could not get DataPointProperties:")
            print(f"  {e}")

    # endregion adjust properties

    # endregion background colors


__all__ = ("Chart",)
