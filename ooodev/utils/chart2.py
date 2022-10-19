# region Imports
from __future__ import annotations
from enum import Enum
from random import random
from typing import List, Tuple, cast, overload

from com.sun.star.beans import XPropertySet
from com.sun.star.chart2 import XAxis
from com.sun.star.chart2 import XChartDocument
from com.sun.star.chart2 import XChartType
from com.sun.star.chart2 import XChartTypeContainer
from com.sun.star.chart2 import XChartTypeTemplate
from com.sun.star.chart2 import XCoordinateSystem
from com.sun.star.chart2 import XCoordinateSystemContainer
from com.sun.star.chart2 import XDataSeries
from com.sun.star.chart2 import XDataSeriesContainer
from com.sun.star.chart2 import XDiagram
from com.sun.star.chart2 import XFormattedString
from com.sun.star.chart2 import XLegend
from com.sun.star.chart2 import XRegressionCurve
from com.sun.star.chart2 import XRegressionCurveContainer
from com.sun.star.chart2 import XScaling
from com.sun.star.chart2 import XTitle
from com.sun.star.chart2 import XTitled
from com.sun.star.chart2.data import XDataProvider
from com.sun.star.chart2.data import XDataSink
from com.sun.star.chart2.data import XDataSource
from com.sun.star.chart2.data import XLabeledDataSequence
from com.sun.star.container import XNameAccess
from com.sun.star.document import XEmbeddedObjectSupplier
from com.sun.star.drawing import XDrawPage
from com.sun.star.drawing import XDrawPageSupplier
from com.sun.star.drawing import XShape
from com.sun.star.embed import XComponentSupplier
from com.sun.star.embed import XEmbeddedObject
from com.sun.star.graphic import XGraphic
from com.sun.star.lang import XComponent
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.table import XTableChart
from com.sun.star.table import XTableChartsSupplier
from com.sun.star.util import XNumberFormatsSupplier

from . import color as mColor
from . import file_io as mFileIo
from . import gui as mGui
from . import images_lo as mImgLo
from . import info as mInfo
from . import lo as mLo
from . import props as mProps
from ..exceptions import ex as mEx
from ..office import calc as mCalc
from .data_type.angle import Angle as AngleType
from .kind.axis_kind import AxisKind as AxisKindEnum
from .kind.chart2_types import ChartTypeBase, ChartTypes
from .kind.curve_kind import CurveKind as CurveKindEnum
from .kind.data_point_label_type_kind import DataPointLabelTypeKind as DataPointLabelTypeKindEnum
from .kind.data_point_lable_placement_kind import DataPointLabelPlacementKind as DataPointLabelPlacementKindEnum
from .kind.line_style_name_kind import LineStyleNameKind

from ooo.dyn.awt.rectangle import Rectangle
from ooo.dyn.chart.chart_data_row_source import ChartDataRowSource
from ooo.dyn.chart.error_bar_style import ErrorBarStyle
from ooo.dyn.chart2.axis_orientation import AxisOrientation
from ooo.dyn.chart2.axis_type import AxisTypeEnum
from ooo.dyn.chart2.data_point_geometry3_d import DataPointGeometry3DEnum
from ooo.dyn.chart2.data_point_label import DataPointLabel
from ooo.dyn.drawing.fill_style import FillStyle as FillStyleEnum
from ooo.dyn.drawing.line_style import LineStyle as LineStyleEnum
from ooo.dyn.lang.locale import Locale
from ooo.dyn.table.cell_range_address import CellRangeAddress

# endregion Imports


class Chart2:
    _CHART_NAME = "chart$$_"

    # region Type Alias
    Angle = AngleType
    AxisKind = AxisKindEnum
    AxisType = AxisTypeEnum
    CurveKind = CurveKindEnum
    DataPointGeometry3D = DataPointGeometry3DEnum
    DataPointLabelPlacementKind = DataPointLabelPlacementKindEnum
    DataPointLabelTypeKind = DataPointLabelTypeKindEnum
    DiagramNames = ChartTypes
    FillStyle = FillStyleEnum
    LineStyle = LineStyleEnum
    LineStyleName = LineStyleNameKind
    # endregion Type Alias

    # region insert a chart
    @classmethod
    def insert_chart(
        cls,
        sheet: XSpreadsheet,
        cells_range: CellRangeAddress,
        cell_name: str,
        width: int,
        height: int,
        diagram_name: ChartTypeBase | str,
        color_bg: mColor.Color = mColor.CommonColor.PALE_BLUE,
        color_wall: mColor.Color = mColor.CommonColor.LIGHT_BLUE,
    ) -> XChartDocument:
        chart_name = Chart2._CHART_NAME + str(int(random() * 10_000))
        cls.add_table_chart(
            sheet=sheet,
            chart_name=chart_name,
            cells_range=cells_range,
            cell_name=cell_name,
            width=width,
            height=height,
        )
        chart_doc = cls.get_chart_doc(sheet, chart_name)

        # assign chart template to the chart's diagram
        diagram = chart_doc.getFirstDiagram()
        ct_template = cls.set_templeate(chart_doc=chart_doc, diagram=diagram, diagram_name=diagram_name)

        has_cats = cls.has_categories(diagram_name)

        dp = chart_doc.getDataProvider()

        ps = mProps.Props.make_props(
            CellRangeRepresentation=mCalc.Calc.get_range_str(cells_range, sheet),
            DataRowSource=ChartDataRowSource.COLUMNS,
            FirstCellAsLabel=True,
            HasCategories=has_cats,
        )
        ds = dp.createDataSource(ps)

        # add data source to chart template
        args = mProps.Props.make_props(HasCategories=has_cats)
        ct_template.changeDiagramData(diagram, ds, args)

        # apply style settings to chart doc
        # background and wall colors
        cls.set_background_colors(chart_doc, color_bg, color_wall)

        if has_cats:
            cls.set_data_point_labels(chart_doc, Chart2.DataPointLabelTypeKind.NUMBER)

        return chart_doc

    @staticmethod
    def add_table_chart(
        sheet: XSpreadsheet, chart_name: str, cells_range: CellRangeAddress, cell_name: str, width: int, height: int
    ) -> None:
        charts_supp = mLo.Lo.qi(XTableChartsSupplier, sheet, True)
        tbl_charts = charts_supp.getCharts()

        pos = mCalc.Calc.get_cell_pos(sheet, cell_name)
        rect = Rectangle(X=pos.X, Y=pos.Y, Width=width * 1000, Height=height * 1000)
        addrs = (cells_range,)

        # 2nd last arg: whether the topmost row of the source data will
        # be used to set labels for the category axis or the legend;
        # last arg: whether the leftmost column of the source data will
        # be used to set labels for the category axis or the legend.
        tbl_charts.addNewByName(chart_name, rect, addrs, True, True)

    @staticmethod
    def set_template(
        chart_doc: XChartDocument, diagram: XDiagram, diagram_name: ChartTypeBase | str
    ) -> XChartTypeTemplate:
        ct_man = chart_doc.getChartTypeManager()
        msf = mLo.Lo.qi(XMultiServiceFactory, ct_man, True)
        template_nm = f"com.sun.star.chart2.template.{diagram_name}"
        ct_template = mLo.Lo.qi(XChartTypeTemplate, msf.createInstance(template_nm))
        if ct_template is None:
            mLo.Lo.print(f'Could not create chart template "{diagram_name}"; using a column chart instead')
            ct_template = mLo.Lo.qi(
                XChartTypeTemplate, msf.createInstance("com.sun.star.chart2.template.Column", True)
            )

        ct_template.changeDiagram(diagram)
        return ct_template

    @staticmethod
    def has_categories(diagram_name: ChartTypeBase | str) -> bool:
        """
        Gets if diagram name has categories

        Args:
            diagram_name (ChartTypeBase | str): Diagram Name

        Returns:
            bool: ``True`` if has categories; Otherwise, ``False``.
        """
        # All the chart templates, except for scatter and bubble use
        # categories on the x-axis
        dn = str(diagram_name).lower()
        non_cats = ("scatter", "bubble")
        for non_cat in non_cats:
            if non_cat in dn:
                return False
        return True

    # endregion insert a chart

    # region get a chart
    @classmethod
    def get_chart_doc(cls, sheet: XSpreadsheet, chart_name: str) -> XChartDocument:
        tbl_chart = cls.get_table_chart(sheet, chart_name)
        eos = mLo.Lo.qi(XEmbeddedObjectSupplier, tbl_chart, True)
        return mLo.Lo.qi(XChartDocument, eos.getEmbeddedObject(), True)

    @staticmethod
    def get_table_chart(sheet: XSpreadsheet, chart_name: str) -> XTableChart:
        charts_supp = mLo.Lo.qi(XTableChartsSupplier, sheet, True)
        tbl_charts = charts_supp.getCharts()
        tc_access = mLo.Lo.qi(XNameAccess, tbl_charts, True)
        tbl_chart = mLo.Lo.qi(XTableChart, tc_access.getByName(chart_name))
        return tbl_chart

    @staticmethod
    def get_chart_templates(chart_doc: XChartDocument) -> List[str]:
        ct_man = chart_doc.getChartTypeManager()
        return mInfo.Info.get_available_services(ct_man)

    # endregion get a chart

    # region titles
    @classmethod
    def set_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
        # return XTilte so it may have futher styles applied
        titled = mLo.Lo.qi(XTitled, chart_doc, True)
        xtitle = cls.create_title(title)
        titled.setTitleObject(xtitle)
        fname = mInfo.Info.get_font_general_name()
        cls.set_x_title_font(xtitle, fname, 14)
        return xtitle

    @staticmethod
    def create_title(title: str) -> XTitle:
        xtitle = mLo.Lo.create_instance_mcf(XTitle, "com.sun.star.chart2.Title", raise_err=True)
        xtitle_str = mLo.Lo.create_instance_mcf(
            XFormattedString, "com.sun.star.chart2.FormattedString", raise_err=True
        )
        xtitle_str.setString(title)
        title_arr = (xtitle_str,)
        xtitle.setText(title_arr)
        return xtitle

    @staticmethod
    def set_x_title_font(xtitle: XTitle, font_name: str, pt_size: int) -> None:
        fo_strs = xtitle.getText()
        if fo_strs:
            mProps.Props.set_property(fo_strs[0], "CharFontName", font_name)
            mProps.Props.set_property(fo_strs[0], "CharHeight", pt_size)

    @staticmethod
    def get_title(chart_doc: XChartDocument) -> XTitle:
        xtilted = mLo.Lo.qi(XTitled, chart_doc, True)
        return xtilted.getTitleObject()

    @classmethod
    def set_subtitle(cls, chart_doc: XChartDocument, subtitle: str) -> XTitle:
        diagram = chart_doc.getFirstDiagram()
        titled = mLo.Lo.qi(XTitled, diagram, True)
        title = cls.create_title(subtitle)
        titled.setTitleObject(title)
        fname = mInfo.Info.get_font_general_name()
        cls.set_x_title_font(title, fname, 12)
        return title

    @staticmethod
    def get_subtitle(chart_doc: XChartDocument) -> XTitle:
        diagram = chart_doc.getFirstDiagram()
        titled = mLo.Lo.qi(XTitled, diagram, True)
        return titled.getTitleObject()

    # endregion titles

    # region Axis
    @classmethod
    def get_axis(cls, chart_doc: XChartDocument, axis_val: Chart2.AxisKind, idx: int) -> XAxis:
        coord_sys = cls.get_coord_system(chart_doc)
        return coord_sys.getAxisByDimension(int(axis_val), idx)

    @classmethod
    def get_x_axis(cls, chart_doc: XChartDocument) -> XAxis:
        return cls.get_axis(chart_doc=chart_doc, axis_val=Chart2.AxisKind.X, idx=0)

    @classmethod
    def get_y_axis(cls, chart_doc: XChartDocument) -> XAxis:
        return cls.get_axis(chart_doc=chart_doc, axis_val=Chart2.AxisKind.Y, idx=0)

    @classmethod
    def get_x_axis2(cls, chart_doc: XChartDocument) -> XAxis:
        return cls.get_axis(chart_doc=chart_doc, axis_val=Chart2.AxisKind.X, idx=1)

    @classmethod
    def get_y_axis2(cls, chart_doc: XChartDocument) -> XAxis:
        return cls.get_axis(chart_doc=chart_doc, axis_val=Chart2.AxisKind.Y, idx=1)

    @classmethod
    def set_axis_title(cls, chart_doc: XChartDocument, title: str, axis_val: Chart2.AxisKind, idx: int) -> XTitle:
        axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
        titled_axis = mLo.Lo.qi(XTitled, axis, True)
        xtitle = cls.create_title(title)
        titled_axis.setTitleObject(xtitle)
        fname = mInfo.Info.get_font_general_name()
        cls.set_x_title_font(xtitle, fname, 12)
        return xtitle

    @classmethod
    def set_x_axis_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=Chart2.AxisKind.X, idx=0)

    @classmethod
    def set_y_axis_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=Chart2.AxisKind.Y, idx=0)

    @classmethod
    def set_x_axis2_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=Chart2.AxisKind.X, idx=1)

    @classmethod
    def set_y_axis2_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=Chart2.AxisKind.Y, idx=1)

    @classmethod
    def get_axis_title(cls, chart_doc: XChartDocument, axis_val: Chart2.AxisKind, idx: int) -> XTitle:
        axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
        titled_axis = mLo.Lo.qi(XTitled, axis, True)
        return titled_axis.getTitleObject()

    @classmethod
    def get_x_axis_title(cls, chart_doc: XChartDocument) -> XTitle:
        return cls.get_axis_title(chart_doc=chart_doc, axis_val=Chart2.AxisKind.X, idx=0)

    @classmethod
    def get_y_axis_title(cls, chart_doc: XChartDocument) -> XTitle:
        return cls.get_axis_title(chart_doc=chart_doc, axis_val=Chart2.AxisKind.Y, idx=0)

    @classmethod
    def get_x_axis2_title(cls, chart_doc: XChartDocument) -> XTitle:
        return cls.get_axis_title(chart_doc=chart_doc, axis_val=Chart2.AxisKind.X, idx=1)

    @classmethod
    def get_y_axis2_title(cls, chart_doc: XChartDocument) -> XTitle:
        return cls.get_axis_title(chart_doc=chart_doc, axis_val=Chart2.AxisKind.Y, idx=1)

    @classmethod
    def rotate_axis_title(
        cls, chart_doc: XChartDocument, axis_val: Chart2.AxisKind, idx: int, angle: Chart2.Angle
    ) -> None:
        xtitle = cls.get_axis_title(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
        mProps.Props.set_property(xtitle, "TextRotation", angle.Value)

    @classmethod
    def rotate_x_axis_title(cls, chart_doc: XChartDocument, angle: Chart2.Angle) -> None:
        cls.rotate_axis_title(chart_doc=chart_doc, axis_val=Chart2.AxisKind.X, idx=0, angle=angle)

    @classmethod
    def rotate_y_axis_title(cls, chart_doc: XChartDocument, angle: Chart2.Angle) -> None:
        cls.rotate_axis_title(chart_doc=chart_doc, axis_val=Chart2.AxisKind.Y, idx=0, angle=angle)

    @classmethod
    def rotate_x_axis2_title(cls, chart_doc: XChartDocument, angle: Chart2.Angle) -> None:
        cls.rotate_axis_title(chart_doc=chart_doc, axis_val=Chart2.AxisKind.X, idx=1, angle=angle)

    @classmethod
    def rotate_y_axis2_title(cls, chart_doc: XChartDocument, angle: Chart2.Angle) -> None:
        cls.rotate_axis_title(chart_doc=chart_doc, axis_val=Chart2.AxisKind.Y, idx=1, angle=angle)

    @classmethod
    def show_axis_label(cls, chart_doc: XChartDocument, axis_val: Chart2.AxisKind, idx: int, is_visible: bool) -> None:
        axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
        mProps.Props.set_property(axis, "Show", is_visible)

    @classmethod
    def show_x_axis_label(cls, chart_doc: XChartDocument, is_visible: bool) -> None:
        cls.show_axis_label(chart_doc=chart_doc, axis_val=Chart2.AxisKind.X, idx=0, is_visible=is_visible)

    @classmethod
    def show_y_axis_label(cls, chart_doc: XChartDocument, is_visible: bool) -> None:
        cls.show_axis_label(chart_doc=chart_doc, axis_val=Chart2.AxisKind.Y, idx=0, is_visible=is_visible)

    @classmethod
    def show_x_axis2_label(cls, chart_doc: XChartDocument, is_visible: bool) -> None:
        cls.show_axis_label(chart_doc=chart_doc, axis_val=Chart2.AxisKind.X, idx=1, is_visible=is_visible)

    @classmethod
    def show_y_axis2_label(cls, chart_doc: XChartDocument, is_visible: bool) -> None:
        cls.show_axis_label(chart_doc=chart_doc, axis_val=Chart2.AxisKind.Y, idx=1, is_visible=is_visible)

    @classmethod
    def scale_axis(
        cls, chart_doc: XChartDocument, axis_val: Chart2.AxisKind, idx: int, scale_type: Chart2.CurveKind
    ) -> XAxis:
        axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
        sd = axis.getScaleData()
        s = None
        if scale_type == Chart2.CurveKind.LINEAR:
            s = "LinearScaling"
        elif scale_type == Chart2.CurveKind.LOGARITHMIC:
            s = "LogarithmicScaling"
        elif scale_type == Chart2.CurveKind.EXPONENTIAL:
            s = "ExponentialScaling"
        elif scale_type == Chart2.CurveKind.POWER:
            s = "PowerScaling"
        if s is None:
            mLo.Lo.print(f'Did not reconize scaling type: "{scale_type}"')
        else:
            sd.Scaling = mLo.Lo.create_instance_mcf(XScaling, f"com.sun.star.chart2.{s}", True)
        axis.setScaleData(sd)
        return axis

    @classmethod
    def scale_x_axis(cls, chart_doc: XChartDocument, scale_type: Chart2.CurveKind) -> XAxis:
        return cls.scale_axis(chart_doc=chart_doc, axis_val=Chart2.AxisKind.X, idx=0, scale_type=scale_type)

    @classmethod
    def scale_y_axis(cls, chart_doc: XChartDocument, scale_type: Chart2.CurveKind) -> XAxis:
        return cls.scale_axis(chart_doc=chart_doc, axis_val=Chart2.AxisKind.Y, idx=0, scale_type=scale_type)

    @classmethod
    def print_scale_data(cls, axis_name: str, axis: XAxis) -> None:
        sd = axis.getScaleData()
        print(f"Scaled Data for: {axis_name}")
        print(f"  Minimum: {sd.Minimum}")
        print(f"  Maximum: {sd.Maximum}")
        print(f"  Origin: {sd.Origin}")

        if sd.Orientation == AxisOrientation.MATHEMATICAL:
            print("  Orientation: mathematical")
        else:
            print("  Orientation: reverse")
        print(f"  Scaling: {mInfo.Info.get_implementation_name(sd.Scaling)}")
        print(f"  AxisType: {cls.get_axis_type_string(Chart2.AxisType(sd.AxisType))}")
        print(f"  AutoDateAxis: {sd.AutoDateAxis}")
        print(f"  ShiftedCategoryPosition: {sd.ShiftedCategoryPosition}")
        print(f"  IncrementData: {sd.IncrementData}")
        print(f"  TimeIncrement: {sd.TimeIncrement}")

    @staticmethod
    def get_axis_type_string(axis_type: Chart2.AxisType) -> str:
        if axis_type == Chart2.AxisType.REALNUMBER:
            return "real numbers"
        elif axis_type == Chart2.AxisType.PERCENT:
            return "percentages"
        elif axis_type == Chart2.AxisType.CATEGORY:
            return "categories"
        elif axis_type == Chart2.AxisType.SERIES:
            return "series names"
        elif axis_type == Chart2.AxisType.DATE:
            return "dates"
        else:
            raise mEx.UnKnownError("axis_type is of unknow type")

    # endregion Axis

    # region grid lines
    # region set_grid_lines()
    @overload
    @classmethod
    def set_grid_lines(cls, chart_doc: XChartDocument, axis_val: Chart2.AxisKind) -> XPropertySet:
        ...

    @overload
    @classmethod
    def set_grid_lines(cls, chart_doc: XChartDocument, axis_val: Chart2.AxisKind, idx: int) -> XPropertySet:
        ...

    @classmethod
    def set_grid_lines(cls, chart_doc: XChartDocument, axis_val: Chart2.AxisKind, idx: int = 0) -> XPropertySet:
        axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
        props = axis.getGridProperties()
        mProps.Props.set_property(props, "LineStyle", Chart2.LineStyle.DASH)
        mProps.Props.set_property(props, "LineDashName", str(Chart2.LineStyleName.FINE_DOTTED))
        return props

    # endregion set_grid_lines()

    # endregion grid lines

    # region legend
    @staticmethod
    def view_legend(chart_doc: XChartDocument, is_visible: bool) -> None:
        diagram = chart_doc.getFirstDiagram()
        legend = diagram.getLegend()
        if is_visible and legend is None:
            leg = mLo.Lo.create_instance_mcf(XLegend, "com.sun.star.chart2.Legend", raise_err=True)
            mProps.Props.set_property(leg, "LineStyle", Chart2.LineStyle.NONE)
            mProps.Props.set_property(leg, "FillStyle", Chart2.FillStyle.SOLID)
            mProps.Props.set_property(leg, "FillTransparence", 100)
            diagram.setLegend(leg)

        mProps.Props.set_property(leg, "Show", is_visible)

    # endregion legend

    # region background colors
    @staticmethod
    def set_background_colors(chart_doc: XChartDocument, bg_color: mColor.Color, wall_color: mColor.Color) -> None:
        if int(bg_color) > 0:
            bg_ps = chart_doc.getPageBackground()
            # mProps.Props.show_props("Background", bg_ps)
            mProps.Props.set_property(bg_ps, "FillBackground", True)
            mProps.Props.set_property(bg_ps, "FillStyle", Chart2.FillStyle.SOLID)
            mProps.Props.set_property(bg_ps, "FillColor", int(bg_color))

        if int(wall_color) > 0:
            diagram = chart_doc.getFirstDiagram()
            wall_ps = diagram.getWall()
            # mProps.Props.show_props("Wall", wall_ps)
            mProps.Props.set_property(wall_ps, "FillBackground", True)
            mProps.Props.set_property(wall_ps, "FillStyle", Chart2.FillStyle.SOLID)
            mProps.Props.set_property(wall_ps, "FillColor", int(wall_color))

    # endregion background colors

    # region access data source and series
    @staticmethod
    def create_data_series() -> XDataSeries:
        try:
            ds = mLo.Lo.create_instance_mcf(XDataSeries, "com.sun.star.chart2.DataSeries", raise_err=True)
            return ds
        except Exception as e:
            raise mEx.ChartError("Error creating data series") from e

    # region get_data_series()
    @overload
    @classmethod
    def get_data_series(cls, chart_doc: XChartDocument) -> Tuple[XDataSeries, ...]:
        ...

    @overload
    @classmethod
    def get_data_series(cls, chart_doc: XChartDocument, chart_type: str) -> Tuple[XDataSeries, ...]:
        ...

    @classmethod
    def get_data_series(cls, chart_doc: XChartDocument, chart_type: str = "") -> Tuple[XDataSeries, ...]:
        if chart_type:
            xchart_type = cls.find_chart_type(chart_doc, chart_type)
        else:
            xchart_type = cls.get_chart_type(chart_doc)
        ds_con = mLo.Lo.qi(XDataSeriesContainer, xchart_type, True)
        return ds_con.getDataSeries()

    # endregion get_data_series()

    # region get_data_source()
    @overload
    @classmethod
    def get_data_source(cls, chart_doc: XChartDocument) -> XDataSource:
        ...

    @overload
    @classmethod
    def get_data_source(cls, chart_doc: XChartDocument, chart_type: str) -> XDataSource:
        ...

    @classmethod
    def get_data_source(cls, chart_doc: XChartDocument, chart_type: str = "") -> XDataSource:
        dsa = cls.get_data_series(chart_doc=chart_doc, chart_type=chart_type)
        ds = mLo.Lo.qi(XDataSource, dsa[0], True)
        return ds

    # endregion get_data_source()

    # endregion access data source and series

    # region chart types
    @staticmethod
    def get_coord_system(chart_doc: XChartDocument) -> XCoordinateSystem:
        try:
            diagram = chart_doc.getFirstDiagram()
            coord_sys_con = mLo.Lo.qi(XCoordinateSystemContainer, diagram, True)
            coord_sys = coord_sys_con.getCoordinateSystems()
            if coord_sys:
                if len(coord_sys) > 1:
                    mLo.Lo.print(f"No. of coord systems: {len(coord_sys)}; using first.")
            return coord_sys[0]  # will raise error if coord_sys is empyt or none
        except Exception as e:
            raise mEx.ChartError("Error unable to get coord_system") from e

    @classmethod
    def get_chart_types(cls, chart_doc: XChartDocument) -> Tuple[XChartType, ...]:
        coord_sys = cls.get_coord_system(chart_doc)
        ct_con = mLo.Lo.qi(XChartTypeContainer, coord_sys, True)
        return ct_con.getChartTypes()

    @classmethod
    def get_chart_type(cls, chart_doc: XChartDocument) -> XChartType:
        chart_types = cls.get_chart_types(chart_doc)
        return chart_types[0]

    @classmethod
    def find_chart_type(cls, chart_doc: XChartDocument, chart_type: str) -> XChartType:
        try:
            srch_name = f"com.sun.star.chart2.{chart_type.lower()}"
            chart_types = cls.get_chart_types(chart_doc)
            for ct in chart_types:
                ct_name = ct.getChartType().lower()
                if ct_name == srch_name:
                    return ct
        except Exception as e:
            raise mEx.ChartError(f'Error Finding chart for "{chart_type}"') from e
        raise mEx.NotFoundError(f'Chart for type "{chart_type}" was not found')

    @classmethod
    def print_chart_types(cls, chart_doc: XChartDocument) -> None:
        chart_types = cls.get_chart_types(chart_doc)
        if len(chart_types) > 1:
            print(f"No. of chart tyeps: {len(chart_types)}")
            for ct in chart_types:
                print(f"  {ct.getChartType()}")
        else:
            print(f"Chart Type: {chart_types[0].getChartType()}")
        print()

    @classmethod
    def add_chart_type(cls, chart_doc: XChartDocument, chart_type: str) -> XChartType:
        ct = mLo.Lo.create_instance_mcf(XChartType, f"com.sun.star.chart2.{chart_type}", raise_err=True)
        coord_sys = cls.get_coord_system(chart_doc)
        ct_con = mLo.Lo.qi(XChartTypeContainer, coord_sys, True)
        ct_con.addChartType(ct)
        return ct

    # endregion chart types

    # region using a data source
    @staticmethod
    def show_data_source_args(chart_doc: XChartDocument, data_source: XDataSource) -> None:
        dp = chart_doc.getDataProvider()
        ps = dp.detectArguments()
        mProps.Props.show_props("Data Source arguments", ps)

    @staticmethod
    def print_labled_seqs(data_source: XDataSource) -> None:
        data_seqs = data_source.getDataSequences()
        print(f"No. of sequeneces in data source: {len(data_seqs)}")
        for seq in data_seqs:
            label_seq = seq.getLabel().getData()
            print(f"{label_seq[0]} :")
            vals_seq = seq.getValues().getData()
            for val in vals_seq:
                print(f"  {val}")
            print()
            sr_rep = seq.getValues().getSourceRangeRepresentation()
            print(f"  Source range: {sr_rep}")
        print()

    @staticmethod
    def get_chart_data(data_source: XDataSource, idx: int) -> Tuple[float]:
        data_seqs = data_source.getDataSequences()
        if idx < 0 or idx >= len(data_seqs):
            raise IndexError(f"Index value of {idx} is out of of range")

        vals_seq = data_seqs[idx].getValues().getData()
        vals: List[float] = []
        for val in vals_seq:
            vals.append(float(val))
        return tuple(vals)

    # endregion using a data source

    # region using data series point props
    @classmethod
    def get_data_points_props(cls, chart_doc: XChartDocument, idx: int) -> List[XPropertySet]:
        data_series_arr = cls.get_data_series(chart_doc=chart_doc)
        if idx < 0 or idx >= len(data_series_arr):
            raise IndexError(f"Index value of {idx} is out of of range")

        props_lst: List[XPropertySet] = []
        i = 0
        while True:
            try:
                props = data_series_arr[idx].getDataPointByIndex(i)
                if props is not None:
                    props_lst.append(props)
                i += 1
            except Exception:
                props = None

            if props is None:
                break
        if len(props_lst) == 0:
            mLo.Lo.print(f"No Series at index {idx}")
        return props_lst

    @classmethod
    def get_data_point_props(cls, chart_doc: XChartDocument, series_idx: int, idx: int) -> XPropertySet:
        props = cls.get_data_points_props(chart_doc=chart_doc, idx=series_idx)
        if not props:
            raise mEx.NotFoundError("No Datapoints found to get XPropertySet from")

        if idx < 0 or idx >= len(props):
            raise IndexError(f"Index value of {idx} is out of of range")

        return props[idx]

    @classmethod
    def set_data_point_labels(cls, chart_doc: XChartDocument, label_type: Chart2.DataPointLabelTypeKind) -> None:
        data_series_arr = cls.get_data_series(chart_doc=chart_doc)
        for data_series in data_series_arr:
            dp_label = cast(DataPointLabel, mProps.Props.get_property(data_series, "Label"))
            dp_label.ShowNumber = False
            dp_label.ShowCategoryName = False
            dp_label.ShowLegendSymbol = False
            if label_type == Chart2.DataPointLabelTypeKind.NUMBER:
                dp_label.ShowNumber = True
            elif label_type == Chart2.DataPointLabelTypeKind.PERCENT:
                dp_label.ShowNumber = True
                dp_label.ShowNumberInPercent = True
            elif label_type == Chart2.DataPointLabelTypeKind.CATEGORY:
                dp_label.ShowCategoryName = True
            elif label_type == Chart2.DataPointLabelTypeKind.SYMBOL:
                dp_label.ShowLegendSymbol = True
            elif label_type == Chart2.DataPointLabelTypeKind.NONE:
                pass
            else:
                raise mEx.UnKnownError("label_type is of unknow type")

            mProps.Props.set_property(data_series, "Label", dp_label)

    @classmethod
    def set_chart_shape_3d(cls, chart_doc: XChartDocument, shape: Chart2.DataPointGeometry3D) -> None:
        data_series_arr = cls.get_data_series(chart_doc=chart_doc)
        for data_series in data_series_arr:
            mProps.Props.set_property(data_series, "Geometry3D", int(shape))

    @classmethod
    def dash_lines(cls, chart_doc: XChartDocument) -> None:
        data_series_arr = cls.get_data_series(chart_doc=chart_doc)
        for data_series in data_series_arr:
            mProps.Props.set_property(data_series, "LineStyle", Chart2.LineStyle.DASH)
            mProps.Props.set_property(data_series, "LineDashName", str(Chart2.LineStyleName.FINE_DASHED))

    @staticmethod
    def color_stock_bars(ct: XChartType, w_day_color: mColor.Color, b_day_color: mColor.Color) -> None:
        if ct.getChartType() == "com.sun.star.chart2.CandleStickChartType":
            ps = mLo.Lo.qi(XPropertySet, mProps.Props.get_property(ct, "WhiteDay"), True)
            mProps.Props.set_property(ps, "FillColor", int(w_day_color))

            ps = mLo.Lo.qi(XPropertySet, mProps.Props.get_property(ct, "BlackDay"), True)
            mProps.Props.set_property(ps, "FillColor", int(b_day_color))
        else:
            raise mEx.NotSupportedError(f'Only candel stick charts supported. "{ct.getChartType()}" not supported.')

    # endregion using data series point props

    # region regression
    @staticmethod
    def create_curve(curve_kind: Chart2.CurveKind) -> XRegressionCurve:
        rc = mLo.Lo.create_instance_mcf(XRegressionCurve, curve_kind.to_namespace(), raise_err=True)
        return rc

    @classmethod
    def draw_regression_curve(cls, chart_doc: XChartDocument, curve_kind: Chart2.CurveKind) -> None:
        data_series_arr = cls.get_data_series(chart_doc=chart_doc)
        rc_con = mLo.Lo.qi(XRegressionCurveContainer, data_series_arr[0], True)
        curve = cls.create_curve(curve_kind)
        rc_con.addRegressionCurve(curve)

        ps = curve.getEquationProperties()
        mProps.Props.set_property(ps, "ShowCorrelationCoefficient", True)
        mProps.Props.set_property(ps, "ShowEquation", True)

        key = cls.get_number_format_key(chart_doc=chart_doc, nf_str="0.00")  # 2 dp
        if key != -1:
            mProps.Props.set_property(ps, "NumberFormat", key)

    @staticmethod
    def get_number_format_key(chart_doc: XChartDocument, nf_str: str) -> int:
        xfs = mLo.Lo.qi(XNumberFormatsSupplier, chart_doc, True)
        n_formats = xfs.getNumberFormats()
        key = int(n_formats.queryKey(nf_str, Locale("en", "us", ""), False))
        if key == -1:
            mLo.Lo.print(f'Could not access key for number format: "{nf_str}"')
        return key

    @staticmethod
    def get_curve_type(curve: XRegressionCurve) -> CurveKind:
        services = set(mInfo.Info.get_services(curve))
        if Chart2.CurveKind.LINEAR.to_namespace() in services:
            return Chart2.CurveKind.LINEAR
        elif Chart2.CurveKind.LOGARITHMIC.to_namespace() in services:
            return Chart2.CurveKind.LOGARITHMIC
        elif Chart2.CurveKind.EXPONENTIAL.to_namespace() in services:
            return Chart2.CurveKind.EXPONENTIAL
        elif Chart2.CurveKind.POWER.to_namespace() in services:
            return Chart2.CurveKind.POWER
        elif Chart2.CurveKind.POLYNOMIAL.to_namespace() in services:
            return Chart2.CurveKind.POLYNOMIAL
        elif Chart2.CurveKind.MOVING_AVERAGE.to_namespace() in services:
            return Chart2.CurveKind.MOVING_AVERAGE
        else:
            raise mEx.UnKnownError("Could not identify trend type of curve")

    @classmethod
    def eval_curve(cls, chart_doc: XChartDocument, curve: XRegressionCurve) -> None:
        curve_calc = curve.getCalculator()
        degree = 1
        ct = cls.get_curve_type(curve)
        if ct != Chart2.CurveKind.LINEAR:
            degree = 2  # assumes POLYNOMIAL trend has degree == 2

        # degree, forceIntercept, interceptValue, period (for moving average)
        curve_calc.setRegressionProperties(degree, False, 0.0, 2)

        data_source = cls.get_data_source(chart_doc)
        # cls.print_labled_seqs(data_source)

        xvals = cls.get_chart_data(data_source=data_source, idx=0)
        yvals = cls.get_chart_data(data_source=data_source, idx=0)
        curve_calc.recalculateRegression(xvals, yvals)

        print(f"  Curve equations: {curve_calc.getRepresentation()}")
        cc = curve_calc.getCorrelationCoefficient()
        print(f"  R^2 value: {(cc*cc):.3f}")

    @classmethod
    def calc_regressions(cls, chart_doc: XChartDocument) -> None:
        def curve_info(curve_kind: Chart2.CurveKind) -> None:
            curve = cls.create_curve(curve_kind=curve_kind)
            print(f"{curve_kind.label} regression curve:")
            cls.eval_curve(chart_doc=chart_doc, curve=curve)
            print()

        curve_info(Chart2.CurveKind.LINEAR)
        curve_info(Chart2.CurveKind.LOGARITHMIC)
        curve_info(Chart2.CurveKind.EXPONENTIAL)
        curve_info(Chart2.CurveKind.POWER)
        curve_info(Chart2.CurveKind.POLYNOMIAL)
        curve_info(Chart2.CurveKind.MOVING_AVERAGE)

    # endregion regression

    # region add data to a chart
    @staticmethod
    def create_ld_seq(dp: XDataProvider, role: str, data_label: str, data_range: str) -> XLabeledDataSequence:
        # reate labeled data sequence using label and data;
        # the data is for the specified role

        # create data sequence for the label
        lbl_seq = dp.createDataSequenceByRangeRepresentation(data_label)

        # reate data sequence for the data and role
        data_seq = dp.createDataSequenceByRangeRepresentation(data_range)

        ds_ps = mLo.Lo.qi(XPropertySet, data_seq, True)

        # specify data role (type)
        mProps.Props.set_property(ds_ps, "Role", role)
        # mProps.Props.show_props("Data Sequence", ds_ps)

        # create new labeled data sequence using sequences
        ld_seq = mLo.Lo.create_instance_mcf(
            XLabeledDataSequence, "com.sun.star.chart2.data.LabeledDataSequence", raise_err=True
        )
        ld_seq.setLabel(lbl_seq)
        ld_seq.setValues(data_seq)
        return ld_seq

    @classmethod
    def set_y_error_bars(cls, chart_doc: XChartDocument, data_label: str, data_range: str) -> None:
        error_bars_ps = mLo.Lo.create_instance_mcf(XPropertySet, "com.sun.star.chart2.ErrorBar", raise_err=True)
        mProps.Props.set_property(error_bars_ps, "ShowPositiveError", True)
        mProps.Props.set_property(error_bars_ps, "ShowNegativeError", True)
        mProps.Props.set_property(error_bars_ps, "ErrorBarStyle", ErrorBarStyle.FROM_DATA)

        # convert into data sink
        data_sink = mLo.Lo.qi(XDataSink, error_bars_ps, True)

        # use data provider to create labelled data sequences
        # for the +/- error ranges
        dp = chart_doc.getDataProvider()

        pos_err_seq = cls.create_ld_seq(
            dp=dp, role="error-bars-y-positive", data_label=data_label, data_range=data_range
        )
        neg_err_seq = cls.create_ld_seq(
            dp=dp, role="error-bars-y-negative", data_label=data_label, data_range=data_range
        )

        ld_seq = (pos_err_seq, neg_err_seq)

        # store the error bar data sequences in the data sink
        data_sink.setData(ld_seq)
        # mProps.Props.show_obj_props("Error Bar", error_bars_ps)
        # "ErrorBarRangePositive" and "ErrorBarRangeNegative"
        # will now have ranges they are read-only

        # store error bar in data series
        data_series_arr = cls.get_data_series(chart_doc=chart_doc)
        # print(f'No. of data serice: {len(data_series_arr)}')
        data_series = data_series_arr[0]
        # mProps.Props.show_obj_props("Data Series 0", data_series)
        mProps.Props.set_property(data_series, "ErrorBarY", error_bars_ps)

    @classmethod
    def add_stock_line(cls, chart_doc: XChartDocument, data_label: str, data_range: str) -> None:
        # add (empty) line chart to the doc
        ct = cls.add_chart_type(chart_doc=chart_doc, chart_type="LineChartType")
        data_series_cnt = mLo.Lo.qi(XDataSeriesContainer, ct, True)

        # create (empty) data series in the line chart
        ds = mLo.Lo.create_instance_mcf(XDataSeries, "com.sun.star.chart2.DataSeries", raise_err=True)

        mProps.Props.set_property(ds, "Color", int(mColor.CommonColor.RED))
        data_series_cnt.addDataSeries(ds)

        # add data to series by treating it as a data sink
        data_sink = mLo.Lo.qi(XDataSink, ds, True)

        # add data as y values
        dp = chart_doc.getDataProvider()
        dl_seq = cls.create_ld_seq(dp=dp, role="values-y", data_label=data_label, data_range=data_range)
        ld_seq_arr = (dl_seq,)
        data_sink.setData(ld_seq_arr)

    @classmethod
    def add_cat_labels(cls, chart_doc: XChartDocument, data_label: str, data_range: str) -> None:
        dp = chart_doc.getDataProvider()
        dl_seq = cls.create_ld_seq(dp=dp, role="categories", data_label=data_label, data_range=data_range)

        axis = cls.get_axis(chart_doc=chart_doc, axis_val=Chart2.AxisKind.X, idx=0)
        sd = axis.getScaleData()
        sd.Categories = dl_seq
        axis.setScaleData(sd)

        # abel the data points with these category values
        cls.set_data_point_labels(chart_doc=chart_doc, label_type=Chart2.DataPointLabelTypeKind.CATEGORY)

    # endregion add data to a chart

    # region chart shape and image
    @staticmethod
    def get_chart_shape(sheet: XSpreadsheet) -> XShape:
        shape = None
        try:
            page_supp = mLo.Lo.qi(XDrawPageSupplier, sheet, True)
            draw_page = page_supp.getDrawPage()
            num_shapes = draw_page.getCount()
            chart_classid = mLo.Lo.CLSID.CHART.value
            for i in range(num_shapes):
                try:
                    shape = mLo.Lo.qi(XShape, draw_page.getByIndex(i), True)
                    classid = str(mProps.Props.get_property(shape, "CLSID")).lower()
                    if classid == chart_classid:
                        break
                except Exception:
                    shape = None
                    # continue on, just because got an error does not mean shape will not be found
        except Exception as e:
            raise mEx.ShapeError("Error getting shape from sheet") from e
        if shape is None:
            raise mEx.ShapeMissingError("Unalbe to find Chart Shape")
        return shape

    @classmethod
    def copy_chart(cls, ssdoc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        chart_shape = cls.get_chart_shape(sheet=sheet)
        doc = mLo.Lo.qi(XComponent, ssdoc, True)
        supp = mGui.GUI.get_selection_supplier(doc)
        supp.select(chart_shape)
        mLo.Lo.dispatch_cmd("Copy")

    @classmethod
    def get_chart_draw_page(cls, sheet: XSpreadsheet) -> XDrawPage:
        chart_shape = cls.get_chart_shape(sheet)
        embedded_chart = mLo.Lo.qi(XEmbeddedObject, mProps.Props.get_property(chart_shape, "EmbeddedObject"), True)
        comp_supp = mLo.Lo.qi(XComponentSupplier, embedded_chart, True)
        xclosable = comp_supp.getComponent()
        supp_page = mLo.Lo.qi(XDrawPageSupplier, xclosable, True)
        return supp_page.getDrawPage()

    @classmethod
    def get_chart_image(cls, sheet: XSpreadsheet) -> XGraphic:
        chart_shape = cls.get_chart_shape(sheet)

        graphic = mLo.Lo.qi(XGraphic, mProps.Props.get_property(chart_shape, "Graphic"), True)

        tmp_fnm = mFileIo.FileIO.create_temp_file("png")
        mImgLo.ImagesLo.save_graphic(pic=graphic, fnm=tmp_fnm, im_format="png")
        im = mImgLo.ImagesLo.load_graphic_file(tmp_fnm)
        mFileIo.FileIO.delete_file(tmp_fnm)
        return im

    # endregion chart shape and image
