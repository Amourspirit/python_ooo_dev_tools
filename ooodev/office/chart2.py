# region Imports
from __future__ import annotations
from random import random
from typing import List, Tuple, cast, overload, TYPE_CHECKING

import uno

# XChartTypeTemplate import error in LO >+ 7.4
from com.sun.star.chart2 import XChartTypeTemplate
from com.sun.star.beans import XPropertySet
from com.sun.star.chart2 import XAxis
from com.sun.star.chart2 import XChartDocument
from com.sun.star.chart2 import XChartType
from com.sun.star.chart2 import XChartTypeContainer
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

from . import calc as mCalc
from ..exceptions import ex as mEx
from ..utils import color as mColor
from ..utils import file_io as mFileIo
from ..utils import gui as mGui
from ..utils import images_lo as mImgLo
from ..utils import info as mInfo
from ..utils import lo as mLo
from ..utils import props as mProps
from ..utils.data_type.angle import Angle as Angle
from ..utils.kind.axis_kind import AxisKind as AxisKind
from ..utils.kind.chart2_data_role_kind import DataRoleKind as DataRoleKind
from ..utils.kind.chart2_types import ChartTemplateBase, ChartTypeNameBase, ChartTypes as ChartTypes
from ..utils.kind.curve_kind import CurveKind as CurveKind
from ..utils.kind.data_point_label_type_kind import DataPointLabelTypeKind as DataPointLabelTypeKind
from ..utils.kind.line_style_name_kind import LineStyleNameKind as LineStyleNameKind

from ooo.dyn.awt.rectangle import Rectangle
from ooo.dyn.chart.chart_data_row_source import ChartDataRowSource
from ooo.dyn.chart.error_bar_style import ErrorBarStyle
from ooo.dyn.chart2.axis_orientation import AxisOrientation
from ooo.dyn.chart2.axis_type import AxisTypeEnum as AxisTypeKind
from ooo.dyn.chart2.data_point_geometry3_d import DataPointGeometry3DEnum as DataPointGeometry3DEnum
from ooo.dyn.chart2.data_point_label import DataPointLabel
from ooo.dyn.drawing.fill_style import FillStyle as FillStyle
from ooo.dyn.drawing.line_style import LineStyle as LineStyle
from ooo.dyn.lang.locale import Locale
from ooo.dyn.table.cell_range_address import CellRangeAddress

if TYPE_CHECKING:
    from ooo.lo.chart2.data_point_properties import DataPointProperties

# endregion Imports


class Chart2:
    """Chart 2 Class"""

    _CHART_NAME = "chart$$_"

    # region insert a chart
    @classmethod
    def insert_chart(
        cls,
        sheet: XSpreadsheet,
        cells_range: CellRangeAddress,
        cell_name: str,
        width: int,
        height: int,
        diagram_name: ChartTemplateBase | str,
        color_bg: mColor.Color = mColor.CommonColor.PALE_BLUE,
        color_wall: mColor.Color = mColor.CommonColor.LIGHT_BLUE,
    ) -> XChartDocument:
        """
        Insert a new chart

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cells_range (CellRangeAddress): Cell range address
            cell_name (str): Cell name such as ``A1``
            width (int): Width
            height (int): Height
            diagram_name (ChartTemplateBase | str): Diagram Name
            color_bg (Color, optional): Color Background. Defaults to ``CommonColor.PALE_BLUE``.
            color_wall (Color, optional): Color Wall. Defaults to ``CommonColor.LIGHT_BLUE``.

        Raises:
            ChartError: If error occurs

        Returns:
            XChartDocument: Chart Document that was created and inserted

        See Also:
            :py:class:`~.color.CommonColor`

        Hint:
            .. include:: ../../resources/utils/chart2_lookup_chart_tmpl.rst
        """
        try:
            # type check that diagram_name is ChartTemplateBase | str
            mInfo.Info.is_type_enum_multi(
                alt_type="str", enum_type=ChartTemplateBase, enum_val=diagram_name, arg_name="diagram_name"
            )
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
            ct_template = cls.set_template(chart_doc=chart_doc, diagram=diagram, diagram_name=diagram_name)

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
                cls.set_data_point_labels(chart_doc, DataPointLabelTypeKind.NUMBER)

            return chart_doc
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error inserting chart") from e

    @staticmethod
    def add_table_chart(
        sheet: XSpreadsheet, chart_name: str, cells_range: CellRangeAddress, cell_name: str, width: int, height: int
    ) -> None:
        """
        Adds a new table chart at a given cell name and size, using cells range.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            chart_name (str): Chart name.
            cells_range (CellRangeAddress): Cell Range
            cell_name (str): Cell Name such as ``A1``
            width (int): Chart Width
            height (int): Chart Height

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            charts_supp = mLo.Lo.qi(XTableChartsSupplier, sheet, True)
            tbl_charts = charts_supp.getCharts()

            pos = mCalc.Calc.get_cell_pos(sheet, cell_name)
            rect = Rectangle(X=pos.X, Y=pos.Y, Width=width * 1_000, Height=height * 1_000)
            addrs = (cells_range,)

            # 2nd last arg: whether the topmost row of the source data will
            # be used to set labels for the category axis or the legend;
            # last arg: whether the leftmost column of the source data will
            # be used to set labels for the category axis or the legend.
            tbl_charts.addNewByName(chart_name, rect, addrs, True, True)
        except Exception as e:
            raise mEx.ChartError("Error adding table chart") from e

    @staticmethod
    def set_template(
        chart_doc: XChartDocument, diagram: XDiagram, diagram_name: ChartTemplateBase | str
    ) -> XChartTypeTemplate:
        """
        Sets template of chart

        Args:
            chart_doc (XChartDocument): Chart Document
            diagram (XDiagram): diagram
            diagram_name (ChartTemplateBase | str): Diagram template name

        Raises:
            ChartError: If error occurs.

        Returns:
            XInterface: Chart Template

        Note:
            If unable to create template from ``diagram_name`` for any reason then
            ``ChartTypes.Column.TEMPLATE_STACKED.COLUMN`` (``"Column"``) is used as a fallback.

        Hint:
            .. include:: ../../resources/utils/chart2_lookup_chart_tmpl.rst
        """
        # XChartTypeTemplate does not seem to be supported by LO 7.4 ( gets import error )
        # Available interfaces com.sun.star.chart2.template.Column: (also XInterface)
        # com.sun.star.beans.XFastPropertySet
        # com.sun.star.beans.XMultiPropertySet
        # com.sun.star.beans.XMultiPropertyStates
        # com.sun.star.beans.XPropertySet
        # com.sun.star.beans.XPropertyState
        # com.sun.star.lang.XServiceName
        # com.sun.star.lang.XTypeProvider
        # com.sun.star.style.XStyleSupplier
        # com.sun.star.uno.XWeak

        # in LO 7.3 com.sun.star.chart2.XChartTypeTemplate is included

        # ensure diagram_name is ChartTemplateBase | str
        mInfo.Info.is_type_enum_multi(
            alt_type=str, enum_type=ChartTemplateBase, enum_val=diagram_name, arg_name="diagram_name"
        )

        try:
            ct_man = chart_doc.getChartTypeManager()
            msf = mLo.Lo.qi(XMultiServiceFactory, ct_man, True)
            template_nm = f"com.sun.star.chart2.template.{diagram_name}"
            ct_template = mLo.Lo.qi(XChartTypeTemplate, msf.createInstance(template_nm))
            if ct_template is None:
                mLo.Lo.print(f'Could not create chart template "{diagram_name}"; using a column chart instead')
                ct_template = mLo.Lo.qi(
                    XChartTypeTemplate, msf.createInstance("com.sun.star.chart2.template.Column"), True
                )

            ct_template.changeDiagram(diagram)
            return ct_template
        except Exception as e:
            raise mEx.ChartError("Error setting chart template") from e

    @staticmethod
    def has_categories(diagram_name: ChartTemplateBase | str) -> bool:
        """
        Gets if diagram name has categories

        Args:
            diagram_name (ChartTemplateBase | str): Diagram Name

        Returns:
            bool: ``True`` if has categories; Otherwise, ``False``.

        Hint:
            .. include:: ../../resources/utils/chart2_lookup_chart_tmpl.rst
        """
        # All the chart templates, except for scatter and bubble use
        # categories on the x-axis

        # Ensure diagram_name ChartTemplateBase | str
        mInfo.Info.is_type_enum_multi(
            alt_type="str", enum_type=ChartTemplateBase, enum_val=diagram_name, arg_name="diagram_name"
        )

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
        """
        Gets the chart document from the sheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            chart_name (str): Chart Name

        Raises:
            ChartError: If error occurs

        Returns:
            XChartDocument: Spreadsheet document.
        """
        try:
            tbl_chart = cls.get_table_chart(sheet, chart_name)
            eos = mLo.Lo.qi(XEmbeddedObjectSupplier, tbl_chart, True)
            return mLo.Lo.qi(XChartDocument, eos.getEmbeddedObject(), True)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError(f'Error getting chart document for chart "{chart_name}"') from e

    @staticmethod
    def get_table_chart(sheet: XSpreadsheet, chart_name: str) -> XTableChart:
        """
        Gets the named table chart from the sheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            chart_name (str): Chart Name

        Raises:
            ChartError: If error occurs.

        Returns:
            XTableChart: Table Chart
        """
        try:
            charts_supp = mLo.Lo.qi(XTableChartsSupplier, sheet, True)
            tbl_charts = charts_supp.getCharts()
            tc_access = mLo.Lo.qi(XNameAccess, tbl_charts, True)
            tbl_chart = mLo.Lo.qi(XTableChart, tc_access.getByName(chart_name))
            return tbl_chart
        except Exception as e:
            raise mEx.ChartError(f'Error getting table chart for chart "{chart_name}"') from e

    @staticmethod
    def get_chart_templates(chart_doc: XChartDocument) -> List[str]:
        """
        Gets a list of chart templates (services).

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs

        Returns:
            List[str]: List of chart templates
        """
        try:
            ct_man = chart_doc.getChartTypeManager()
            return mInfo.Info.get_available_services(ct_man)
        except Exception as e:
            raise mEx.ChartError("Error getting chart templates") from e

    # endregion get a chart

    # region titles
    @classmethod
    def set_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
        """
        Sets the title of chart

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title as string

        Raises:
            ChartError: If error occurs.

        Returns:
            XTitle: Title Object

        Note:
            The title is set to a font size of ``14`` and the font applied is
            the font returned by :py:meth:`.Info.get_font_general_name`

        Hint:
            The returned Title object can be passed to :py:meth:`~.Chart2.set_x_title_font` to
            change default font.
        """
        try:
            # return XTilte so it may have futher styles applied
            titled = mLo.Lo.qi(XTitled, chart_doc, True)
            xtitle = cls.create_title(title)
            titled.setTitleObject(xtitle)
            fname = mInfo.Info.get_font_general_name()
            cls.set_x_title_font(xtitle, fname, 14)
            return xtitle
        except Exception as e:
            raise mEx.ChartError("Error setting title for chart") from e

    @staticmethod
    def create_title(title: str) -> XTitle:
        """
        Creates a title object

        Args:
            title (str): Title text.

        Raises:
            ChartError: If error occurs.

        Returns:
            XTitle: Title object.
        """
        try:
            xtitle = mLo.Lo.create_instance_mcf(XTitle, "com.sun.star.chart2.Title", raise_err=True)
            xtitle_str = mLo.Lo.create_instance_mcf(
                XFormattedString, "com.sun.star.chart2.FormattedString", raise_err=True
            )
            xtitle_str.setString(title)
            title_arr = (xtitle_str,)
            xtitle.setText(title_arr)
            return xtitle
        except Exception as e:
            raise mEx.ChartError(f'Error creating title for: "{title}"') from e

    @staticmethod
    def set_x_title_font(xtitle: XTitle, font_name: str, pt_size: int) -> None:
        """
        Sets X title font.

        Args:
            xtitle (XTitle): Title instance.
            font_name (str): Font Name
            pt_size (int): Font point size

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`.Info.get_font_general_name`
        """
        try:
            fo_strs = xtitle.getText()
            if fo_strs:
                mProps.Props.set_property(fo_strs[0], "CharFontName", font_name)
                mProps.Props.set_property(fo_strs[0], "CharHeight", pt_size)
        except Exception as e:
            raise mEx.ChartError("Error setting x title font") from e

    @staticmethod
    def get_title(chart_doc: XChartDocument) -> XTitle:
        """
        Gets Title from chart

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            XTitle: Title object.
        """
        try:
            xtilted = mLo.Lo.qi(XTitled, chart_doc, True)
            return xtilted.getTitleObject()
        except Exception as e:
            raise mEx.ChartError("Error getting title from chart") from e

    @classmethod
    def set_subtitle(cls, chart_doc: XChartDocument, subtitle: str) -> XTitle:
        """
        Gets subtitle

        Args:
            chart_doc (XChartDocument): Chart Document.
            subtitle (str): Subtitle text.

        Raises:
            ChartError: If error occurs

        Returns:
            XTitle: Title object

        Note:
            The subtitle is set to a font size of ``12`` and the font applied is
            the font returned by :py:meth:`.Info.get_font_general_name`

        Hint:
            The returned Title object can be passed to :py:meth:`~.Chart2.set_x_title_font` to
            change default font.
        """
        try:
            diagram = chart_doc.getFirstDiagram()
            titled = mLo.Lo.qi(XTitled, diagram, True)
            title = cls.create_title(subtitle)
            titled.setTitleObject(title)
            fname = mInfo.Info.get_font_general_name()
            cls.set_x_title_font(title, fname, 12)
            return title
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError(f'Error setting subtitle "{subtitle}" for chart') from e

    @staticmethod
    def get_subtitle(chart_doc: XChartDocument) -> XTitle:
        """
        Gets subtitle from chart

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            XTitle: Title object.
        """
        try:
            diagram = chart_doc.getFirstDiagram()
            titled = mLo.Lo.qi(XTitled, diagram, True)
            return titled.getTitleObject()
        except Exception as e:
            raise mEx.ChartError("Error getting subtitle from chart") from e

    # endregion titles

    # region Axis
    @classmethod
    def get_axis(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int) -> XAxis:
        """
        Gets axis

        Args:
            chart_doc (XChartDocument): Chart Document
            axis_val (AxisKind): Axis Kind
            idx (int): Axis index

        Raises:
            ChartError: If error occurs.

        Returns:
            XAxis: Axis object.
        """
        try:
            coord_sys = cls.get_coord_system(chart_doc)
            result = coord_sys.getAxisByDimension(int(axis_val), idx)
            if result is None:
                raise mEx.UnKnownError("None Value: getAxisByDimension() returned None")
            return result
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error getting Axis for chart") from e

    @classmethod
    def get_x_axis(cls, chart_doc: XChartDocument) -> XAxis:
        """
        Get Chart X axis.

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            XAxis: Chart X Axis object.

        See Also:
            :py:meth:`~.Chart2.get_axis`
        """
        return cls.get_axis(chart_doc=chart_doc, axis_val=AxisKind.X, idx=0)

    @classmethod
    def get_y_axis(cls, chart_doc: XChartDocument) -> XAxis:
        """
        Get Chart Y axis.

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            XAxis: Chart Y Axis object.

        See Also:
            :py:meth:`~.Chart2.get_axis`
        """
        return cls.get_axis(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=0)

    @classmethod
    def get_x_axis2(cls, chart_doc: XChartDocument) -> XAxis:
        """
        Get Chart X axis2.

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            XAxis: Chart X Axis2 object.

        See Also:
            :py:meth:`~.Chart2.get_axis`
        """
        return cls.get_axis(chart_doc=chart_doc, axis_val=AxisKind.X, idx=1)

    @classmethod
    def get_y_axis2(cls, chart_doc: XChartDocument) -> XAxis:
        """
        Get Chart Y axis2.

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            XAxis: Chart Y Axis2 object.

        See Also:
            :py:meth:`~.Chart2.get_axis`
        """
        return cls.get_axis(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=1)

    @classmethod
    def set_axis_title(cls, chart_doc: XChartDocument, title: str, axis_val: AxisKind, idx: int) -> XTitle:
        """
        Sets axis title.

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title text.
            axis_val (AxisKind): Axis kind.
            idx (int): Index

        Raises:
            ChartError: If error occurs.

        Returns:
            XTitle: Axis Title.

        Note:
            The title is set to a font size of ``12`` and the font applied is
            the font returned by :py:meth:`.Info.get_font_general_name`

        Hint:
            The returned Title object can be passed to :py:meth:`~.Chart2.set_x_title_font` to
            change default font.

        See Also:
            :py:meth:`~.Chart2.get_axis`
        """
        try:
            axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
            titled_axis = mLo.Lo.qi(XTitled, axis, True)
            xtitle = cls.create_title(title)
            titled_axis.setTitleObject(xtitle)
            fname = mInfo.Info.get_font_general_name()
            cls.set_x_title_font(xtitle, fname, 12)
            return xtitle
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError(f'Error setting axis tile: "{title}" for chart') from e

    @classmethod
    def set_x_axis_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
        """
        Sets X axis Title

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title Text

        Returns:
            XTitle: Title object

        See Also:
            :py:meth:`~.Chart2.set_axis_title`
        """
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.X, idx=0)

    @classmethod
    def set_y_axis_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
        """
        Sets Y axis Title

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title Text

        Returns:
            XTitle: Title object

        See Also:
            :py:meth:`~.Chart2.set_axis_title`
        """
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.Y, idx=0)

    @classmethod
    def set_x_axis2_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
        """
        Sets X axis2 Title

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title Text

        Returns:
            XTitle: Title object

        See Also:
            :py:meth:`~.Chart2.set_axis_title`
        """
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.X, idx=1)

    @classmethod
    def set_y_axis2_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
        """
        Sets Y axis2 Title

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title Text

        Returns:
            XTitle: Title object

        See Also:
            :py:meth:`~.Chart2.set_axis_title`
        """
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.Y, idx=1)

    @classmethod
    def get_axis_title(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int) -> XTitle:
        """
        Gets axis Title

        Args:
            chart_doc (XChartDocument): Chart Document
            axis_val (AxisKind): Axis Kind.
            idx (int): Index

        Raises:
            ChartError: If error occurs

        Returns:
            XTitle: Title object.

        See Also:
            :py:meth:`~.Chart2.get_axis`
        """
        try:
            axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
            titled_axis = mLo.Lo.qi(XTitled, axis, True)
            result = titled_axis.getTitleObject()
            if result is None:
                raise mEx.UnKnownError("None Value: getTitleObject() return a value of None")
            return result
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error getting axis title") from e

    @classmethod
    def get_x_axis_title(cls, chart_doc: XChartDocument) -> XTitle:
        """
        Gets X axis title.

        Args:
            chart_doc (XChartDocument): Chart Document

        Raises:
            ChartError: If error occurs

        Returns:
            XTitle: Title object.

        See Also:
            :py:meth:`~.Chart2.get_axis_title`
        """
        return cls.get_axis_title(chart_doc=chart_doc, axis_val=AxisKind.X, idx=0)

    @classmethod
    def get_y_axis_title(cls, chart_doc: XChartDocument) -> XTitle:
        """
        Gets Y axis title.

        Args:
            chart_doc (XChartDocument): Chart Document

        Raises:
            ChartError: If error occurs

        Returns:
            XTitle: Title object.

        See Also:
            :py:meth:`~.Chart2.get_axis_title`
        """
        return cls.get_axis_title(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=0)

    @classmethod
    def get_x_axis2_title(cls, chart_doc: XChartDocument) -> XTitle:
        """
        Gets X axis2 title.

        Args:
            chart_doc (XChartDocument): Chart Document

        Raises:
            ChartError: If error occurs

        Returns:
            XTitle: Title object.

        See Also:
            :py:meth:`~.Chart2.get_axis_title`
        """
        return cls.get_axis_title(chart_doc=chart_doc, axis_val=AxisKind.X, idx=1)

    @classmethod
    def get_y_axis2_title(cls, chart_doc: XChartDocument) -> XTitle:
        """
        Gets Y axis2 title.

        Args:
            chart_doc (XChartDocument): Chart Document

        Raises:
            ChartError: If error occurs

        Returns:
            XTitle: Title object.

        See Also:
            :py:meth:`~.Chart2.get_axis_title`
        """
        return cls.get_axis_title(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=1)

    @classmethod
    def rotate_axis_title(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int, angle: Angle) -> None:
        """
        Rotates axis title.

        Args:
            chart_doc (XChartDocument): Chart Document.
            axis_val (AxisKind): Axis kind.
            idx (int): Index
            angle (Angle): Angle

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            xtitle = cls.get_axis_title(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
            mProps.Props.set_property(xtitle, "TextRotation", angle.Value)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error while trying to rotate axis title") from e

    @classmethod
    def rotate_x_axis_title(cls, chart_doc: XChartDocument, angle: Angle) -> None:
        """
        Rotates X axis title.

        Args:
            chart_doc (XChartDocument): Chart Document.
            angle (Angle): Angle

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`~.Chart2.rotate_axis_title`
        """
        cls.rotate_axis_title(chart_doc=chart_doc, axis_val=AxisKind.X, idx=0, angle=angle)

    @classmethod
    def rotate_y_axis_title(cls, chart_doc: XChartDocument, angle: Angle) -> None:
        """
        Rotates Y axis title.

        Args:
            chart_doc (XChartDocument): Chart Document.
            angle (Angle): Angle

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`~.Chart2.rotate_axis_title`
        """
        cls.rotate_axis_title(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=0, angle=angle)

    @classmethod
    def rotate_x_axis2_title(cls, chart_doc: XChartDocument, angle: Angle) -> None:
        """
        Rotates X axis2 title.

        Args:
            chart_doc (XChartDocument): Chart Document.
            angle (Angle): Angle

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`~.Chart2.rotate_axis_title`
        """
        cls.rotate_axis_title(chart_doc=chart_doc, axis_val=AxisKind.X, idx=1, angle=angle)

    @classmethod
    def rotate_y_axis2_title(cls, chart_doc: XChartDocument, angle: Angle) -> None:
        """
        Rotates Y axis2 title.

        Args:
            chart_doc (XChartDocument): Chart Document.
            angle (Angle): Angle

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`~.Chart2.rotate_axis_title`
        """
        cls.rotate_axis_title(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=1, angle=angle)

    @classmethod
    def show_axis_label(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int, is_visible: bool) -> None:
        """
        Sets the visibility for chart axis label.

        Args:
            chart_doc (XChartDocument): Chart Document.
            axis_val (AxisKind): Axis kind.
            idx (int): Index
            is_visible (bool): Visible state

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
            mProps.Props.set_property(axis, "Show", is_visible)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error while setting axis label visibility") from e

    @classmethod
    def show_x_axis_label(cls, chart_doc: XChartDocument, is_visible: bool) -> None:
        """
        Sets the visibility for chart X axis label.

        Args:
            chart_doc (XChartDocument): Chart Document.
            is_visible (bool): Visible state

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`~.Chart2.show_axis_label`
        """
        cls.show_axis_label(chart_doc=chart_doc, axis_val=AxisKind.X, idx=0, is_visible=is_visible)

    @classmethod
    def show_y_axis_label(cls, chart_doc: XChartDocument, is_visible: bool) -> None:
        """
        Sets the visibility for chart Y axis label.

        Args:
            chart_doc (XChartDocument): Chart Document.
            is_visible (bool): Visible state

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`~.Chart2.show_axis_label`
        """
        cls.show_axis_label(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=0, is_visible=is_visible)

    @classmethod
    def show_x_axis2_label(cls, chart_doc: XChartDocument, is_visible: bool) -> None:
        """
        Sets the visibility for chart X axis2 label.

        Args:
            chart_doc (XChartDocument): Chart Document.
            is_visible (bool): Visible state

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`~.Chart2.show_axis_label`
        """
        cls.show_axis_label(chart_doc=chart_doc, axis_val=AxisKind.X, idx=1, is_visible=is_visible)

    @classmethod
    def show_y_axis2_label(cls, chart_doc: XChartDocument, is_visible: bool) -> None:
        """
        Sets the visibility for chart Y axis2 label.

        Args:
            chart_doc (XChartDocument): Chart Document.
            is_visible (bool): Visible state

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`~.Chart2.show_axis_label`
        """
        cls.show_axis_label(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=1, is_visible=is_visible)

    @classmethod
    def scale_axis(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int, scale_type: CurveKind) -> XAxis:
        """
        Scales the chart axis.

        Args:
            chart_doc (XChartDocument): Chart Document.
            axis_val (AxisKind): Axis kind.
            idx (int): Index
            scale_type (CurveKind): Scale kind

        Raises:
            ChartError: If error occurs.

        Returns:
            XAxis: Axis object.

        Note:
            Supported types of ``scale_type`` are ``LINEAR``, ``LOGARITHMIC``, ``EXPONENTIAL`` and ``POWER``.
            If ``scale_type``  is not supported then the ``Scaling`` is not set.
        """
        try:
            axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
            sd = axis.getScaleData()
            s = None
            if scale_type == CurveKind.LINEAR:
                s = "LinearScaling"
            elif scale_type == CurveKind.LOGARITHMIC:
                s = "LogarithmicScaling"
            elif scale_type == CurveKind.EXPONENTIAL:
                s = "ExponentialScaling"
            elif scale_type == CurveKind.POWER:
                s = "PowerScaling"
            if s is None:
                mLo.Lo.print(f'Did not reconize scaling type: "{scale_type}"')
            else:
                sd.Scaling = mLo.Lo.create_instance_mcf(XScaling, f"com.sun.star.chart2.{s}", raise_err=True)
            axis.setScaleData(sd)
            return axis
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting axis scale") from e

    @classmethod
    def scale_x_axis(cls, chart_doc: XChartDocument, scale_type: CurveKind) -> XAxis:
        """
        Scales the chart X axis.

        Args:
            chart_doc (XChartDocument): Chart Document.
            scale_type (CurveKind): Scale kind

        Raises:
            ChartError: If error occurs.

        Returns:
            XAxis: Axis object.

        Note:
            Supported types of ``scale_type`` are ``LINEAR``, ``LOGARITHMIC``, ``EXPONENTIAL`` and ``POWER``.
            If ``scale_type``  is not supported then the ``Scaling`` is not set.

        See Also:
            :py:meth:`~.Chart2.scale_axis`
        """
        return cls.scale_axis(chart_doc=chart_doc, axis_val=AxisKind.X, idx=0, scale_type=scale_type)

    @classmethod
    def scale_y_axis(cls, chart_doc: XChartDocument, scale_type: CurveKind) -> XAxis:
        """
        Scales the chart Y axis.

        Args:
            chart_doc (XChartDocument): Chart Document.
            scale_type (CurveKind): Scale kind

        Raises:
            ChartError: If error occurs.

        Returns:
            XAxis: Axis object.

        Note:
            Supported types of ``scale_type`` are ``LINEAR``, ``LOGARITHMIC``, ``EXPONENTIAL`` and ``POWER``.
            If ``scale_type``  is not supported then the ``Scaling`` is not set.

        See Also:
            :py:meth:`~.Chart2.scale_axis`
        """
        return cls.scale_axis(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=0, scale_type=scale_type)

    @classmethod
    def print_scale_data(cls, axis_name: str, axis: XAxis) -> None:
        """
        Prints axis info to console

        Args:
            axis_name (str): Axis Name
            axis (XAxis): Axis
        """
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
        print(f"  AxisType: {cls.get_axis_type_string(AxisTypeKind(sd.AxisType))}")
        print(f"  AutoDateAxis: {sd.AutoDateAxis}")
        print(f"  ShiftedCategoryPosition: {sd.ShiftedCategoryPosition}")
        print(f"  IncrementData: {sd.IncrementData}")
        print(f"  TimeIncrement: {sd.TimeIncrement}")

    @staticmethod
    def get_axis_type_string(axis_type: AxisTypeKind) -> str:
        """
        Gets axis type as string.

        Args:
            axis_type (AxisTypeKind): Axis Type

        Raises:
            UnKnownError: If unable to determine ``axis_type``

        Returns:
            str: Axis type as string.
        """
        if axis_type == AxisTypeKind.REALNUMBER:
            return "real numbers"
        elif axis_type == AxisTypeKind.PERCENT:
            return "percentages"
        elif axis_type == AxisTypeKind.CATEGORY:
            return "categories"
        elif axis_type == AxisTypeKind.SERIES:
            return "series names"
        elif axis_type == AxisTypeKind.DATE:
            return "dates"
        else:
            raise mEx.UnKnownError("axis_type is of unknow type")

    # endregion Axis

    # region grid lines
    # region set_grid_lines()
    @overload
    @classmethod
    def set_grid_lines(cls, chart_doc: XChartDocument, axis_val: AxisKind) -> XPropertySet:
        ...

    @overload
    @classmethod
    def set_grid_lines(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int) -> XPropertySet:
        ...

    @classmethod
    def set_grid_lines(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int = 0) -> XPropertySet:
        """
        Set the grid lines for a chart.

        Args:
            chart_doc (XChartDocument): Chart Document.
            axis_val (AxisKind): Axis kind.
            idx (int, optional): Index. Defaults to 0.

        Raises:
            ChartError: If error occurs.

        Returns:
            XPropertySet: Property Set of Grid Properties.
        """
        try:
            axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
            props = axis.getGridProperties()
            mProps.Props.set_property(props, "LineStyle", LineStyle.DASH)
            mProps.Props.set_property(props, "LineDashName", str(LineStyleNameKind.FINE_DOTTED))
            return props
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting gird lines for chart") from e

    # endregion set_grid_lines()

    # endregion grid lines

    # region legend
    @staticmethod
    def view_legend(chart_doc: XChartDocument, is_visible: bool) -> None:
        """
        Sets charts legend visibility.

        Args:
            chart_doc (XChartDocument): Chart Document.
            is_visible (bool): Visible State

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            diagram = chart_doc.getFirstDiagram()
            legend = diagram.getLegend()
            if is_visible and legend is None:
                leg = mLo.Lo.create_instance_mcf(XLegend, "com.sun.star.chart2.Legend", raise_err=True)
                mProps.Props.set_property(leg, "LineStyle", LineStyle.NONE)
                mProps.Props.set_property(leg, "FillStyle", FillStyle.SOLID)
                mProps.Props.set_property(leg, "FillTransparence", 100)
                diagram.setLegend(leg)

            mProps.Props.set_property(leg, "Show", is_visible)
        except Exception as e:
            raise mEx.ChartError("Error while setting legend visibility") from e

    # endregion legend

    # region background colors
    @staticmethod
    def set_background_colors(chart_doc: XChartDocument, bg_color: mColor.Color, wall_color: mColor.Color) -> None:
        """
        Set the background colors for a chart

        Args:
            chart_doc (XChartDocument): Chart Document.
            bg_color (Color): Color Value for background
            wall_color (Color): Color Value for wall

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:class:`~.color.CommonColor`
        """
        try:
            if int(bg_color) > 0:
                bg_ps = chart_doc.getPageBackground()
                # bg_dpp = cast("DataPointProperties", bg_ps)
                # bg_dpp.FillBackground = True
                # bg_dpp.FillStyle = FillStyle.SOLID
                # bg_dpp.FillColor = int(bg_color)
                #
                # there is a bug with chart_doc.getPageBackground()
                # it is suppose to return XProperySet, which it does but
                # XProperySet methods such as  setPropertyValue and getPropertySetInfo are missing.
                # see also: Props._set_by_attribute()
                #
                # mProps.Props.show_props("Background", bg_ps)
                mProps.Props.set(bg_ps, FillBackground=True, FillStyle=FillStyle.SOLID, FillColor=int(bg_color))

            if int(wall_color) > 0:
                diagram = chart_doc.getFirstDiagram()
                wall_ps = diagram.getWall()
                # wall_dpp = cast("DataPointProperties", wall_ps)
                # wall_dpp.FillBackground = True
                # wall_dpp.FillStyle = FillStyle.SOLID
                # wall_dpp.FillColor = int(wall_color)
                #
                # there is a bug with diagram.getWall()
                # it is suppose to return XProperySet, which it does but
                # XProperySet methods such as  setPropertyValue and getPropertySetInfo are missing.
                # see also: Props._set_by_attribute()
                #
                # mProps.Props.show_props("Wall", wall_ps)
                mProps.Props.set(wall_ps, FillBackground=True, FillStyle=FillStyle.SOLID, FillColor=int(wall_color))
        except Exception as e:
            raise mEx.ChartError("Error setting background colors") from e

    # endregion background colors

    # region access data source and series
    @staticmethod
    def create_data_series() -> XDataSeries:
        """
        Gets an instance of Chart2 DataSeries.

        Raises:
            ChartError: If error occurs.

        Returns:
            XDataSeries: Data Series instance.
        """
        try:
            ds = mLo.Lo.create_instance_mcf(XDataSeries, "com.sun.star.chart2.DataSeries", raise_err=True)
            return ds
        except Exception as e:
            raise mEx.ChartError("Error, unable to create XDataSeries interface") from e

    # region get_data_series()
    @overload
    @classmethod
    def get_data_series(cls, chart_doc: XChartDocument) -> Tuple[XDataSeries, ...]:
        ...

    @overload
    @classmethod
    def get_data_series(cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase) -> Tuple[XDataSeries, ...]:
        ...

    @overload
    @classmethod
    def get_data_series(cls, chart_doc: XChartDocument, chart_type: str) -> Tuple[XDataSeries, ...]:
        ...

    @classmethod
    def get_data_series(
        cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase | str = ""
    ) -> Tuple[XDataSeries, ...]:
        """
        Gets data series for a chart of a given chart type.

        Args:
            chart_doc (XChartDocument): Chart Document
            chart_type (ChartTypeNameBase | str, optional): Chart Type.

        Raises:
            NotFoundError: If chart is not found
            ChartError: If any other error occurs.

        Returns:
            Tuple[XDataSeries, ...]: Data Series

        Hint:
            .. include:: ../../resources/utils/chart2_lookup_chart_name.rst

        See Also:
            :py:meth:`.Chart2.find_chart_type`
        """
        try:
            if chart_type:
                xchart_type = cls.find_chart_type(chart_doc, chart_type)
            else:
                xchart_type = cls.get_chart_type(chart_doc)
            ds_con = mLo.Lo.qi(XDataSeriesContainer, xchart_type, True)
            return ds_con.getDataSeries()
        except Exception as e:
            raise mEx.ChartError("Error getting chart data series") from e

    # endregion get_data_series()

    # region get_data_source()
    @overload
    @classmethod
    def get_data_source(cls, chart_doc: XChartDocument) -> XDataSource:
        ...

    @overload
    @classmethod
    def get_data_source(cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase) -> XDataSource:
        ...

    @overload
    @classmethod
    def get_data_source(cls, chart_doc: XChartDocument, chart_type: str) -> XDataSource:
        ...

    @classmethod
    def get_data_source(cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase | str = "") -> XDataSource:
        """
        Get data source of a chart for a given chart type.

        This method assumes that the programmer wants the first data source in the data series.
        This is adequate for most charts which only use one data source.

        Args:
            chart_doc (XChartDocument): Chart Document
            chart_type (ChartTypeNameBase | str): Chart type.

        Raises:
            NotFoundError: If chart is not found
            ChartError: If any other error occurs.

        Returns:
            XDataSource: Chart data source

        Hint:
            .. include:: ../../resources/utils/chart2_lookup_chart_name.rst

        See Also:
            :py:meth:`~.Chart2.get_data_series`
        """
        try:
            dsa = cls.get_data_series(chart_doc=chart_doc, chart_type=chart_type)
            ds = mLo.Lo.qi(XDataSource, dsa[0], True)
            return ds
        except mEx.NotFoundError:
            raise
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error getting data source for chart") from e

    # endregion get_data_source()

    # endregion access data source and series

    # region chart types
    @staticmethod
    def get_coord_system(chart_doc: XChartDocument) -> XCoordinateSystem:
        """
        Gets coordinate system

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            XCoordinateSystem: Coordinate system object.
        """
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
        """
        Gets chart types for a chart.

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            Tuple[XChartType, ...]: Tuple of chart types.

        See Also:
            :py:meth:`~.Chart2.get_chart_type`
        """
        try:
            coord_sys = cls.get_coord_system(chart_doc)
            ct_con = mLo.Lo.qi(XChartTypeContainer, coord_sys, True)
            result = ct_con.getChartTypes()
            if result is None:
                raise mEx.UnKnownError("None Value: getChartTypes() returned a value of None")
            return result
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error getting chart types") from e

    @classmethod
    def get_chart_type(cls, chart_doc: XChartDocument) -> XChartType:
        """
        Gets a chart type for a chart

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            XChartType: Chart type object.

        See Also:
            :py:meth:`~.Chart2.get_chart_types`
        """
        try:
            chart_types = cls.get_chart_types(chart_doc)
            return chart_types[0]
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error getting chart type") from e

    @classmethod
    def find_chart_type(cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase | str) -> XChartType:
        """
        Finds a chart for a given chart type.

        Args:
            chart_doc (XChartDocument): Chart Document
            chart_type (ChartTypeNameBase | str): Chart type.

        Raises:
            NotFoundError: If chart is not found
            ChartError: If any other error occurs.

        Returns:
            XChartType: Found chart type.

        Hint:
            .. include:: ../../resources/utils/chart2_lookup_chart_name.rst
        """
        # Ensure chart_type is ChartTypeNameBase | str
        mInfo.Info.is_type_enum_multi(
            alt_type="str", enum_type=ChartTypeNameBase, enum_val=chart_type, arg_name="chart_type"
        )
        try:
            srch_name = f"com.sun.star.chart2.{str(chart_type).lower()}"
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
        """
        Prints chart types to the console

        Args:
            chart_doc (XChartDocument): Chart Document.
        """
        chart_types = cls.get_chart_types(chart_doc)
        if len(chart_types) > 1:
            print(f"No. of chart tyeps: {len(chart_types)}")
            for ct in chart_types:
                print(f"  {ct.getChartType()}")
        else:
            print(f"Chart Type: {chart_types[0].getChartType()}")
        print()

    # region add_chart_type()
    @overload
    @classmethod
    def add_chart_type(cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase) -> XChartType:
        ...

    @overload
    @classmethod
    def add_chart_type(cls, chart_doc: XChartDocument, chart_type: str) -> XChartType:
        ...

    @classmethod
    def add_chart_type(cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase | str) -> XChartType:
        """
        Adds a chart type.

        Args:
            chart_doc (XChartDocument): Chart Document
            chart_type (ChartTypeNameBase | str): Chart type.

        Raises:
            ChartError: If error occurs

        Returns:
            XChartType: Chart type object

        Hint:
            .. include:: ../../resources/utils/chart2_lookup_chart_name.rst
        """
        mInfo.Info.is_type_enum_multi(
            alt_type="str", enum_type=ChartTypeNameBase, enum_val=chart_type, arg_name="chart_type"
        )
        try:
            ct = mLo.Lo.create_instance_mcf(XChartType, f"com.sun.star.chart2.{chart_type}", raise_err=True)
            coord_sys = cls.get_coord_system(chart_doc)
            ct_con = mLo.Lo.qi(XChartTypeContainer, coord_sys, True)
            ct_con.addChartType(ct)
            return ct
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error adding chart type") from e

    # endregion add_chart_type()

    # endregion chart types

    # region using a data source
    @staticmethod
    def show_data_source_args(chart_doc: XChartDocument, data_source: XDataSource) -> None:
        """
        Prints data information to console

        Args:
            chart_doc (XChartDocument): Chart Document.
            data_source (XDataSource): Data Source.

        Returns:
            None:
        """
        dp = chart_doc.getDataProvider()
        ps = dp.detectArguments(data_source)
        mProps.Props.show_props("Data Source arguments", ps)

    @staticmethod
    def print_labeled_seqs(data_source: XDataSource) -> None:
        """
        Prints labeled sequence information to console.

        A diagnostic function for printing all the labeled data sequences stored in an XDataSource:

        Args:
            data_source (XDataSource): Data Source.

        Returns:
            None:
        """
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
        """
        Gets chart data

        Args:
            data_source (XDataSource): Data Source.
            idx (int): Index

        Raises:
            IndexError: If index is out of range.
            ChartError: If any other error occurs.

        Returns:
            Tuple[float]: Chart data.
        """
        try:
            data_seqs = data_source.getDataSequences()
            if idx < 0 or idx >= len(data_seqs):
                raise IndexError(f"Index value of {idx} is out of of range")

            vals_seq = data_seqs[idx].getValues().getData()
            vals: List[float] = []
            for val in vals_seq:
                vals.append(float(val))
            return tuple(vals)
        except IndexError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error getting chart data") from e

    # endregion using a data source

    # region using data series point props
    @classmethod
    def get_data_points_props(cls, chart_doc: XChartDocument, idx: int) -> List[XPropertySet]:
        """
        Gets all the properties for the data in the specified series

        Args:
            chart_doc (XChartDocument): Chart Document.
            idx (int): Index

        Raises:
            IndexError: If idx is out of range.

        Returns:
            List[XPropertySet]: Property set list.
        """
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
        if len(props_lst) > 0:
            mLo.Lo.print(f"No Series at index {idx}")
        return props_lst

    @classmethod
    def get_data_point_props(cls, chart_doc: XChartDocument, series_idx: int, idx: int) -> XPropertySet:
        """
        Get the proprieties for a specific index within the data points.

        Args:
            chart_doc (XChartDocument): Chart Document.
            series_idx (int): Series Index
            idx (int): Index to extract from the datapoints data

        Raises:
            NotFoundError: If ``series_idx`` did not find any data points.
            IndexError: If an index is out of range.

        Returns:
            XPropertySet: A single property set from the data points series.

        See Also:
            :py:meth:`~.Chart2.get_data_points_props`
        """
        props = cls.get_data_points_props(chart_doc=chart_doc, idx=series_idx)
        if not props:
            raise mEx.NotFoundError("No Datapoints found to get XPropertySet from")

        if idx < 0 or idx >= len(props):
            raise IndexError(f"Index value of {idx} is out of of range")

        return props[idx]

    @classmethod
    def set_data_point_labels(cls, chart_doc: XChartDocument, label_type: DataPointLabelTypeKind) -> None:
        """
        Sets the data point label of a chart

        Args:
            chart_doc (XChartDocument): Chart Document
            label_type (DataPointLabelTypeKind): Label Type

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            data_series_arr = cls.get_data_series(chart_doc=chart_doc)
            for data_series in data_series_arr:
                dp_label = cast(DataPointLabel, mProps.Props.get_property(data_series, "Label"))
                dp_label.ShowNumber = False
                dp_label.ShowCategoryName = False
                dp_label.ShowLegendSymbol = False
                if label_type == DataPointLabelTypeKind.NUMBER:
                    dp_label.ShowNumber = True
                elif label_type == DataPointLabelTypeKind.PERCENT:
                    dp_label.ShowNumber = True
                    dp_label.ShowNumberInPercent = True
                elif label_type == DataPointLabelTypeKind.CATEGORY:
                    dp_label.ShowCategoryName = True
                elif label_type == DataPointLabelTypeKind.SYMBOL:
                    dp_label.ShowLegendSymbol = True
                elif label_type == DataPointLabelTypeKind.NONE:
                    pass
                else:
                    raise mEx.UnKnownError("label_type is of unknow type")

                mProps.Props.set_property(data_series, "Label", dp_label)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting data point labels") from e

    @classmethod
    def set_chart_shape_3d(cls, chart_doc: XChartDocument, shape: DataPointGeometry3DEnum) -> None:
        """
        Sets chart 3d shape

        Args:
            chart_doc (XChartDocument): Chart Document.
            shape (DataPointGeometry3DEnum): Shape kind

        Raises:
            ChartError: If an error occurs.

        Returns:
            None:
        """
        try:
            data_series_arr = cls.get_data_series(chart_doc=chart_doc)
            for data_series in data_series_arr:
                mProps.Props.set_property(data_series, "Geometry3D", int(shape))
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting chart shape 3d") from e

    @classmethod
    def dash_lines(cls, chart_doc: XChartDocument) -> None:
        """
        Sets chart data series to dashed lines.

        Args:
            chart_doc (XChartDocument): Chart Document

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            data_series_arr = cls.get_data_series(chart_doc=chart_doc)
            for data_series in data_series_arr:
                mProps.Props.set_property(data_series, "LineStyle", LineStyle.DASH)
                mProps.Props.set_property(data_series, "LineDashName", str(LineStyleNameKind.FINE_DASHED))
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting chart sash lines") from e

    @staticmethod
    def color_stock_bars(ct: XChartType, w_day_color: mColor.Color, b_day_color: mColor.Color) -> None:
        """
        Set color of stock bars for a ``CandleStickChartType`` chart.

        Args:
            ct (XChartType): Chart Type.
            w_day_color (Color): Chart white day color
            b_day_color (Color): Chart black day color

        Raises:
            NotSupportedError: If Chart is not of type ``CandleStickChartType``
            ChartError: If any other error occurs.

        Returns:
            None:

        See Also:
            :py:class:`~.color.CommonColor`
        """
        try:
            if ct.getChartType() == "com.sun.star.chart2.CandleStickChartType":
                # there is a bug with white_day_ps and black_day_ps
                # they are suppose br XProperySet, which they are but
                # XProperySet methods such as  setPropertyValue and getPropertySetInfo are missing.
                # see also: Props._set_by_attribute()
                white_day_ps = mLo.Lo.qi(XPropertySet, mProps.Props.get(ct, "WhiteDay"), True)
                white_day_dpp = cast("DataPointProperties", white_day_ps)
                white_day_dpp.FillColor = int(w_day_color)
                # mProps.Props.set_property(white_day_ps, "FillColor", int(w_day_color))

                black_day_ps = mLo.Lo.qi(XPropertySet, mProps.Props.get(ct, "BlackDay"), True)
                black_day_dpp = cast("DataPointProperties", black_day_ps)
                black_day_dpp.FillColor = int(b_day_color)
                # mProps.Props.set_property(black_day_ps, "FillColor", int(b_day_color))
            else:
                raise mEx.NotSupportedError(
                    f'Only candel stick charts supported. "{ct.getChartType()}" not supported.'
                )
        except mEx.NotSupportedError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error coloring stock bars") from e

    # endregion using data series point props

    # region regression
    @staticmethod
    def create_curve(curve_kind: CurveKind) -> XRegressionCurve:
        """
        Creates a regression curve.

        Matches the regression constants defined in ``curve_kind`` to regression services offered by the API:

        Args:
            curve_kind (CurveKind): Curve kind.

        Raises:
            ChartError: If error occurs.

        Returns:
            XRegressionCurve: Regression Curve object.
        """
        try:
            rc = mLo.Lo.create_instance_mcf(XRegressionCurve, curve_kind.to_namespace(), raise_err=True)
            return rc
        except Exception as e:
            raise mEx.ChartError("Error creating curve") from e

    @classmethod
    def draw_regression_curve(cls, chart_doc: XChartDocument, curve_kind: CurveKind) -> None:
        """
        Draws a regression curve.

        Args:
            chart_doc (XChartDocument): Chart Document
            curve_kind (CurveKind): Curve kind.

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
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
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error drawing regression curve") from e

    @staticmethod
    def get_number_format_key(chart_doc: XChartDocument, nf_str: str) -> int:
        """
        Converts a number format string into a number format key, which can be assigned to
        ``NumberFormat`` property.

        Args:
            chart_doc (XChartDocument): Chart Document
            nf_str (str): Number format string.

        Raises:
            ChartError: If error occurs.

        Returns:
            int: Number format key.

        Note:
            The string-to-key conversion is straight forward if you know what number format string to use,
            but there's little documentation on them. Probably the best approach is to use the Format
            Cells menu item in a spreadsheet document, and examine the dialog
        """
        try:
            xfs = mLo.Lo.qi(XNumberFormatsSupplier, chart_doc, True)
            n_formats = xfs.getNumberFormats()
            key = int(n_formats.queryKey(nf_str, Locale("en", "us", ""), False))
            if key == -1:
                mLo.Lo.print(f'Could not access key for number format: "{nf_str}"')
            return key
        except Exception as e:
            raise mEx.ChartError("Error getting number format key") from e

    @staticmethod
    def get_curve_type(curve: XRegressionCurve) -> CurveKind:
        """
        Gets curve kind from regression object.

        Args:
            curve (XRegressionCurve): Regression curve object.

        Raises:
            NotFoundError: If unable to detect curve kind.

        Returns:
            CurveKind: Curve Kind
        """
        services = set(mInfo.Info.get_services(curve))
        if CurveKind.LINEAR.to_namespace() in services:
            return CurveKind.LINEAR
        elif CurveKind.LOGARITHMIC.to_namespace() in services:
            return CurveKind.LOGARITHMIC
        elif CurveKind.EXPONENTIAL.to_namespace() in services:
            return CurveKind.EXPONENTIAL
        elif CurveKind.POWER.to_namespace() in services:
            return CurveKind.POWER
        elif CurveKind.POLYNOMIAL.to_namespace() in services:
            return CurveKind.POLYNOMIAL
        elif CurveKind.MOVING_AVERAGE.to_namespace() in services:
            return CurveKind.MOVING_AVERAGE
        else:
            raise mEx.NotFoundError("Could not identify trend type of curve")

    @classmethod
    def eval_curve(cls, chart_doc: XChartDocument, curve: XRegressionCurve) -> None:
        """
        Uses ``XRegressionCurve.getCalculator()`` to access the ``XRegressionCurveCalculator`` interface.
        It sets up the data and parameters for a particular curve, and prints the results of curve fitting to the console.

        Args:
            chart_doc (XChartDocument): Chart Document
            curve (XRegressionCurve): Regression Curve object.
        """
        curve_calc = curve.getCalculator()
        degree = 1
        ct = cls.get_curve_type(curve)
        if ct != CurveKind.LINEAR:
            degree = 2  # assumes POLYNOMIAL trend has degree == 2

        # degree, forceIntercept, interceptValue, period (for moving average)
        # the last are for setRegressionProperties is movingType
        #   See: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1chart2_1_1MovingAverageType.html
        # movingType Only if regression type is "Moving Average" 1, 3 or 3
        #   see: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XRegressionCurveCalculator.html#ae65112de1214e140d9ce7b28ffb09292
        # Because thi sis not a Moving Average setting to 0
        curve_calc.setRegressionProperties(degree, False, 0.0, 2, 0)

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
        """
        Several different regression functions are calculated using the chart's data.
        Their equations and ``R2`` values are printed to the console

        Args:
            chart_doc (XChartDocument): Chart Document
        """

        def curve_info(curve_kind: CurveKind) -> None:
            curve = cls.create_curve(curve_kind=curve_kind)
            print(f"{curve_kind.label} regression curve:")
            cls.eval_curve(chart_doc=chart_doc, curve=curve)
            print()

        curve_info(CurveKind.LINEAR)
        curve_info(CurveKind.LOGARITHMIC)
        curve_info(CurveKind.EXPONENTIAL)
        curve_info(CurveKind.POWER)
        curve_info(CurveKind.POLYNOMIAL)
        curve_info(CurveKind.MOVING_AVERAGE)

    # endregion regression

    # region add data to a chart
    @staticmethod
    def create_ld_seq(
        dp: XDataProvider, role: DataRoleKind | str, data_label: str, data_range: str
    ) -> XLabeledDataSequence:
        """
        Creates a ``XLabeledDataSequence`` instance from two ``XDataSequence`` objects,
        one acting as a label the other as data.

        The ``XDataSequence`` object representing the data must have its "Role" property
        set to indicate the type of the data.

        Args:
            dp (XDataProvider): Data Provider
            role (DataRoleKind | str): Role that indicate the type of the data.
            data_label (str): Data label
            data_range (str): Data range

        Raises:
            ChartError: If error occurs

        Returns:
            XLabeledDataSequence: Labeled data sequence object.
        """
        # reate labeled data sequence using label and data;
        # the data is for the specified role

        try:
            # create data sequence for the label
            lbl_seq = dp.createDataSequenceByRangeRepresentation(data_label)

            # reate data sequence for the data and role
            data_seq = dp.createDataSequenceByRangeRepresentation(data_range)

            ds_ps = mLo.Lo.qi(XPropertySet, data_seq, True)

            # specify data role (type)
            mProps.Props.set_property(ds_ps, "Role", str(role))
            # mProps.Props.show_props("Data Sequence", ds_ps)

            # create new labeled data sequence using sequences
            ld_seq = mLo.Lo.create_instance_mcf(
                XLabeledDataSequence, "com.sun.star.chart2.data.LabeledDataSequence", raise_err=True
            )
            ld_seq.setLabel(lbl_seq)
            ld_seq.setValues(data_seq)
            return ld_seq
        except Exception as e:
            raise mEx.ChartError("Error creating LD sequence") from e

    @classmethod
    def set_y_error_bars(cls, chart_doc: XChartDocument, data_label: str, data_range: str) -> None:
        """
        Sets Y error Bars

        Args:
            chart_doc (XChartDocument): Chart Document
            data_label (str): Data Label
            data_range (str): Data Range

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            error_bars_ps = mLo.Lo.create_instance_mcf(XPropertySet, "com.sun.star.chart2.ErrorBar", raise_err=True)
            mProps.Props.set(
                error_bars_ps, ShowPositiveError=True, ShowNegativeError=True, ErrorBarStyle=ErrorBarStyle.FROM_DATA
            )

            # convert into data sink
            data_sink = mLo.Lo.qi(XDataSink, error_bars_ps, True)

            # use data provider to create labelled data sequences
            # for the +/- error ranges
            dp = chart_doc.getDataProvider()

            pos_err_seq = cls.create_ld_seq(
                dp=dp, role=DataRoleKind.ERROR_BARS_Y_POSITIVE, data_label=data_label, data_range=data_range
            )
            neg_err_seq = cls.create_ld_seq(
                dp=dp, role=DataRoleKind.ERROR_BARS_Y_NEGATIVE, data_label=data_label, data_range=data_range
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
            mProps.Props.set(data_series, ErrorBarY=error_bars_ps)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error Setting y error bars") from e

    @classmethod
    def add_stock_line(cls, chart_doc: XChartDocument, data_label: str, data_range: str) -> None:
        """
        Add stock line to chart.

        Args:
            chart_doc (XChartDocument): Chart Document
            data_label (str): Data label
            data_range (str): Data range

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            # add (empty) line chart to the doc
            ct = cls.add_chart_type(chart_doc=chart_doc, chart_type=ChartTypes.Line.NAMED.LINE_CHART)
            data_series_cnt = mLo.Lo.qi(XDataSeriesContainer, ct, True)

            # create (empty) data series in the line chart
            ds = mLo.Lo.create_instance_mcf(XDataSeries, "com.sun.star.chart2.DataSeries", raise_err=True)

            mProps.Props.set_property(ds, "Color", int(mColor.CommonColor.RED))
            data_series_cnt.addDataSeries(ds)

            # add data to series by treating it as a data sink
            data_sink = mLo.Lo.qi(XDataSink, ds, True)

            # add data as y values
            dp = chart_doc.getDataProvider()
            dl_seq = cls.create_ld_seq(dp=dp, role=DataRoleKind.VALUES_Y, data_label=data_label, data_range=data_range)
            ld_seq_arr = (dl_seq,)
            data_sink.setData(ld_seq_arr)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error adding stock line") from e

    @classmethod
    def add_cat_labels(cls, chart_doc: XChartDocument, data_label: str, data_range: str) -> None:
        """
        Add Category Labels

        Args:
            chart_doc (XChartDocument): Chart Document
            data_label (str): Data label
            data_range (str): Data range

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            dp = chart_doc.getDataProvider()
            dl_seq = cls.create_ld_seq(
                dp=dp, role=DataRoleKind.CATEGORIES, data_label=data_label, data_range=data_range
            )

            axis = cls.get_axis(chart_doc=chart_doc, axis_val=AxisKind.X, idx=0)
            sd = axis.getScaleData()
            sd.Categories = dl_seq
            axis.setScaleData(sd)

            # abel the data points with these category values
            cls.set_data_point_labels(chart_doc=chart_doc, label_type=DataPointLabelTypeKind.CATEGORY)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error adding category lables") from e

    # endregion add data to a chart

    # region chart shape and image
    @staticmethod
    def get_chart_shape(sheet: XSpreadsheet) -> XShape:
        """
        Gets chart shape

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Raises:
            ShapeMissingError: If shape is ``None``.
            ShapeError: If any other error occurs.

        Returns:
            XShape: Shape object.
        """
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
        """
        Copies a chart using a dispatch command.

        Args:
            ssdoc (XSpreadsheetDocument): Spreadsheet Document
            sheet (XSpreadsheet): Spreadsheet

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            chart_shape = cls.get_chart_shape(sheet=sheet)
            doc = mLo.Lo.qi(XComponent, ssdoc, True)
            supp = mGui.GUI.get_selection_supplier(doc)
            supp.select(chart_shape)
            mLo.Lo.dispatch_cmd("Copy")
        except Exception as e:
            raise mEx.ChartError("Error in attempt to copy chart") from e

    @classmethod
    def get_chart_draw_page(cls, sheet: XSpreadsheet) -> XDrawPage:
        """
        Gets chart draw page.

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Raises:
            ChartError: If error occurs.

        Returns:
            XDrawPage: Draw Page object
        """
        try:
            chart_shape = cls.get_chart_shape(sheet)
            embedded_chart = mLo.Lo.qi(XEmbeddedObject, mProps.Props.get_property(chart_shape, "EmbeddedObject"), True)
            comp_supp = mLo.Lo.qi(XComponentSupplier, embedded_chart, True)
            xclosable = comp_supp.getComponent()
            supp_page = mLo.Lo.qi(XDrawPageSupplier, xclosable, True)
            result = supp_page.getDrawPage()
            if result is None:
                raise mEx.UnKnownError("None Value: getDrawPage() returned a value of None")
            return result
        except Exception as e:
            raise mEx.ChartError("Error getting chart draw page") from e

    @classmethod
    def get_chart_image(cls, sheet: XSpreadsheet) -> XGraphic:
        """
        Get chart image as ``XGraphic``.

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Raises:
            ChartError: If error occurs.

        Returns:
            XGraphic: Graphic object
        """
        try:
            chart_shape = cls.get_chart_shape(sheet)

            graphic = mLo.Lo.qi(XGraphic, mProps.Props.get_property(chart_shape, "Graphic"), True)

            tmp_fnm = mFileIo.FileIO.create_temp_file("png")
            mImgLo.ImagesLo.save_graphic(pic=graphic, fnm=tmp_fnm, im_format="png")
            im = mImgLo.ImagesLo.load_graphic_file(tmp_fnm)
            mFileIo.FileIO.delete_file(tmp_fnm)
            return im
        except Exception as e:
            raise mEx.ChartError("Error getting chart image") from e

    # endregion chart shape and image


__all__ = ("Chart2",)
