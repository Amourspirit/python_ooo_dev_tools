# region Imports
from __future__ import annotations
import contextlib
from random import random
from typing import Any, List, Sequence, Tuple, cast, overload, TYPE_CHECKING

import uno

# XChartTypeTemplate import error in LO 7.4.0 to 7.4.3, Corrected in Lo 7.5
from com.sun.star.beans import XPropertySet
from com.sun.star.chart import XChartDocument as XChartDocumentOld
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

from ooo.dyn.awt.rectangle import Rectangle
from ooo.dyn.chart.chart_data_row_source import ChartDataRowSource
from ooo.dyn.chart.error_bar_style import ErrorBarStyle
from ooo.dyn.chart2.axis_orientation import AxisOrientation
from ooo.dyn.chart2.axis_type import AxisTypeEnum as AxisTypeKind
from ooo.dyn.chart2.data_point_geometry3_d import DataPointGeometry3DEnum as DataPointGeometry3DEnum
from ooo.dyn.drawing.fill_style import FillStyle as FillStyle
from ooo.dyn.drawing.line_style import LineStyle as LineStyle
from ooo.dyn.lang.locale import Locale
from ooo.dyn.table.cell_range_address import CellRangeAddress

from ooodev.events.event_singleton import _Events
from ooodev.office import calc as mCalc
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.chart2_named_event import Chart2NamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.units.angle import Angle as Angle
from ooodev.utils import color as mColor
from ooodev.utils import file_io as mFileIo
from ooodev.gui import gui as mGui
from ooodev.utils import images_lo as mImgLo
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.kind.axis_kind import AxisKind as AxisKind
from ooodev.utils.kind.chart2_data_role_kind import DataRoleKind as DataRoleKind
from ooodev.utils.kind.chart2_types import ChartTemplateBase, ChartTypeNameBase, ChartTypes as ChartTypes
from ooodev.utils.kind.curve_kind import CurveKind as CurveKind
from ooodev.utils.kind.data_point_label_type_kind import DataPointLabelTypeKind as DataPointLabelTypeKind
from ooodev.utils.kind.line_style_name_kind import LineStyleNameKind as LineStyleNameKind


if TYPE_CHECKING:
    from com.sun.star.drawing import OLE2Shape
    from com.sun.star.chart2 import DataPointLabel
    from ooodev.proto.style_obj import StyleT
else:
    StyleT = Any
# endregion Imports

# https://wiki.documentfoundation.org/Documentation/DevGuide/Charts


class Chart2:
    """Chart 2 Class"""

    _CHART_NAME = "chart$$_"

    # region insert/remove a chart
    @classmethod
    def insert_chart(
        cls,
        *,
        sheet: XSpreadsheet | None = None,
        cells_range: CellRangeAddress | None = None,
        cell_name: str = "",
        width: int = 16,
        height: int = 9,
        diagram_name: ChartTemplateBase | str = "Column",
        color_bg: mColor.Color | None = mColor.CommonColor.PALE_BLUE,
        color_wall: mColor.Color | None = mColor.CommonColor.LIGHT_BLUE,
        **kwargs,
    ) -> XChartDocument:
        """
        Insert a new chart.

        |lo_unsafe|

        Args:
            sheet (XSpreadsheet, optional): Spreadsheet
            cells_range (CellRangeAddress, optional): Cell range address. Defaults to current selected cells.
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
            XChartDocument: Chart Document that was created and inserted

        Note:
            **Keyword Arguments** are to mostly be ignored.
            If finer control over chart creation is needed then **Keyword Arguments** can be used.

        Note:
            See **Open Office Wiki** - `The Structure of Charts <https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Structure_of_Charts>`__ for more information.

        See Also:
            - :py:class:`~.color.CommonColor`
            - :ref:`ooodev.utils.kind.chart2_types`

        Hint:
            .. include:: ../../resources/utils/chart2_lookup_chart_tmpl.rst

        .. versionchanged:: 0.8.1
            All parameters made optional. Added ``chart_name`` parameter.
        """
        try:
            # type check that diagram_name is ChartTemplateBase | str
            mInfo.Info.is_type_enum_multi(
                alt_type="str", enum_type=ChartTemplateBase, enum_val=diagram_name, arg_name="diagram_name"
            )
            doc = None
            if sheet is None:
                doc = mCalc.Calc.get_ss_doc(mLo.Lo.this_component)
                sheet = mCalc.Calc.get_active_sheet(doc)
            if sheet is None:
                raise mEx.NoneError("unable to obtain sheet, Calc.get_active_sheet() is None")

            if cells_range is None:
                if doc is None:
                    doc = mCalc.Calc.get_ss_doc(mLo.Lo.this_component)
                cells_range = mCalc.Calc.get_selected_addr(doc)
            if cells_range is None:
                raise mEx.NoneError("unable to obtain cells_range, Calc.get_selected_addr() is None")

            if not cell_name:
                cell_name = mCalc.Calc.get_cell_str(col=cells_range.EndColumn + 1, row=cells_range.StartRow)

            # sourcery skip: low-code-quality, use-or-for-fallback
            chart_name = kwargs.get("chart_name", None)
            if not chart_name:
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
            if "is_row" in kwargs:
                is_row = bool(kwargs["is_row"])
            else:
                is_row = mCalc.Calc.is_single_row_range(cells_range)

            is_single = mCalc.Calc.is_single_column_range(cells_range) or mCalc.Calc.is_single_row_range(cells_range)

            if is_row:
                data_row_source = ChartDataRowSource.ROWS
            else:
                data_row_source = ChartDataRowSource.COLUMNS

            # assign chart template to the chart's diagram
            diagram = chart_doc.getFirstDiagram()
            ct_template = cls.set_template(chart_doc=chart_doc, diagram=diagram, diagram_name=diagram_name)

            arg_first_cell = False
            arg_has_cats = False

            if "set_data_point_labels" in kwargs:
                has_cats = bool(kwargs["set_data_point_labels"])
                arg_has_cats = True
            else:
                has_cats = cls.has_categories(diagram_name)

            dp = chart_doc.getDataProvider()

            if "first_cell_as_label" in kwargs:
                first_cell_as_lbl = bool(kwargs["first_cell_as_label"])
                arg_first_cell = True
            else:
                first_cell_as_lbl = True

            if is_single:
                if not arg_has_cats:
                    has_cats = False
                # get the cell first value
                if not arg_first_cell:
                    first_val = mCalc.Calc.get_val(sheet=sheet, col=cells_range.StartColumn, row=cells_range.StartRow)
                    if isinstance(first_val, float):
                        first_cell_as_lbl = False

            ps = mProps.Props.make_props(
                CellRangeRepresentation=mCalc.Calc.get_range_str(cells_range, sheet),
                DataRowSource=data_row_source,
                FirstCellAsLabel=first_cell_as_lbl,
                HasCategories=has_cats,
            )
            ds = dp.createDataSource(ps)  # type: ignore

            # add data source to chart template
            args = mProps.Props.make_props(HasCategories=has_cats)
            ct_template.changeDiagramData(diagram, ds, args)  # type: ignore

            # apply style settings to chart doc
            # background and wall colors
            if color_bg is None:
                color_bg = mColor.Color(-1)
            if color_wall is None:
                color_wall = mColor.Color(-1)
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

        |lo_safe|

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
            tbl_charts.addNewByName(chart_name, rect, addrs, True, True)  # type: ignore
        except Exception as e:
            raise mEx.ChartError("Error adding table chart") from e

    @staticmethod
    def set_template(
        chart_doc: XChartDocument, diagram: XDiagram, diagram_name: ChartTemplateBase | str
    ) -> XChartTypeTemplate:
        """
        Sets template of chart.

        |lo_safe|

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
        # this should be fixe in LO 7.4.4 +
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
            alt_type="str", enum_type=ChartTemplateBase, enum_val=diagram_name, arg_name="diagram_name"
        )

        try:
            ct_man = chart_doc.getChartTypeManager()
            msf = mLo.Lo.qi(XMultiServiceFactory, ct_man, True)
            if isinstance(diagram_name, ChartTemplateBase):
                template_nm = diagram_name.to_namespace()
            else:
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
        Gets if diagram name has categories.

        |lo_safe|

        Args:
            diagram_name (ChartTemplateBase | str): Diagram Name/

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
        return all(non_cat not in dn for non_cat in non_cats)

    @staticmethod
    def remove_chart(sheet: XSpreadsheet, chart_name: str) -> bool:
        """
        Removes a chart from Spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            chart_name (str): Chart Name

        Returns:
            bool: ``True`` if chart was removed; Otherwise, ``False``

        .. versionadded:: 0.8.1
        """
        charts_supp = mLo.Lo.qi(XTableChartsSupplier, sheet, True)
        tbl_charts = charts_supp.getCharts()
        if tbl_charts.hasByName(chart_name):
            tbl_charts.removeByName(chart_name)
            return True
        return False

    # endregion insert/remove a chart

    # region get a chart
    @classmethod
    def get_chart_doc(cls, sheet: XSpreadsheet, chart_name: str) -> XChartDocument:
        """
        Gets the chart document from the sheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            chart_name (str): Chart Name.

        Raises:
            ChartError: If error occurs.

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
        Gets the named table chart from the sheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            chart_name (str): Chart Name.

        Raises:
            ChartError: If error occurs.

        Returns:
            XTableChart: Table Chart.
        """
        try:
            charts_supp = mLo.Lo.qi(XTableChartsSupplier, sheet, True)
            tbl_charts = charts_supp.getCharts()
            tc_access = mLo.Lo.qi(XNameAccess, tbl_charts, True)
            return mLo.Lo.qi(XTableChart, tc_access.getByName(chart_name), True)
        except Exception as e:
            raise mEx.ChartError(f'Error getting table chart for chart "{chart_name}"') from e

    @staticmethod
    def get_chart_templates(chart_doc: XChartDocument) -> List[str]:
        """
        Gets a list of chart templates (services).

        |lo_safe|

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
    def set_title(cls, chart_doc: XChartDocument, title: str, styles: Sequence[StyleT] | None = None) -> XTitle:
        """
        Sets the title of chart.

        |lo_unsafe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title as string.
            styles (Sequence[StyleT], optional): Styles to apply to title.

        Raises:
            ChartError: If error occurs.

        Returns:
            XTitle: Title Object

        Note:
            The title has a default font size of ``14`` and the font name applied is
            the font returned by :py:meth:`.Info.get_font_general_name`.

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.
        """
        try:
            # return x_title so it may have further styles applied
            titled = mLo.Lo.qi(XTitled, chart_doc, True)
            if styles is None:
                x_title = cls._create_title(title, 14)
            else:
                x_title = cls._create_title(title, 14, styles)
            titled.setTitleObject(x_title)

            return x_title
        except Exception as e:
            raise mEx.ChartError("Error setting title for chart") from e

    @classmethod
    def _create_title(cls, title: str, font_size: int, styles: Sequence[StyleT] | None = None) -> XTitle:
        """LO UN-Safe method."""
        try:
            x_title = mLo.Lo.create_instance_mcf(XTitle, "com.sun.star.chart2.Title", raise_err=True)
            x_title_str = mLo.Lo.create_instance_mcf(
                XFormattedString, "com.sun.star.chart2.FormattedString", raise_err=True
            )
            x_title_str.setString(title)

            # set default font. Styles can override the default.
            fname = mInfo.Info.get_font_general_name()
            cls.set_x_title_font(x_title, fname, font_size)

            title_arr = (x_title_str,)
            x_title.setText(title_arr)
            # Shape style will not be applied. Need to use style_title() after title is created.
            if styles:
                title_styles = [style for style in styles if not style.support_service("com.sun.star.drawing.Shape")]
                cls._style_title(xtitle=x_title, styles=title_styles)

            return x_title
        except Exception as e:
            raise mEx.ChartError(f'Error creating title for: "{title}"') from e

    @classmethod
    def create_title(cls, title: str, styles: Sequence[StyleT] | None = None) -> XTitle:
        """
        Creates a title object.

        |lo_unsafe|

        Args:
            title (str): Title text.
            styles (Sequence[StyleT], optional): Styles to apply to title.

        Raises:
            ChartError: If error occurs.

        Returns:
            XTitle: Title object.

        Note:
            The title has a default font size of ``14`` and the font name applied is
            the font returned by :py:meth:`.Info.get_font_general_name`.

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.
        """
        return cls._create_title(title, 14, styles)

    @staticmethod
    def set_x_title_font(xtitle: XTitle, font_name: str, pt_size: int) -> None:
        """
        Sets X title font.

        |lo_safe|

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
            if fo_strs := xtitle.getText():
                mProps.Props.set_property(fo_strs[0], "CharFontName", font_name)
                mProps.Props.set_property(fo_strs[0], "CharHeight", pt_size)
        except Exception as e:
            raise mEx.ChartError("Error setting x title font") from e

    @staticmethod
    def get_title(chart_doc: XChartDocument) -> XTitle:
        """
        Gets Title from chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            XTitle: Title object.
        """
        try:
            x_tilted = mLo.Lo.qi(XTitled, chart_doc, True)
            return x_tilted.getTitleObject()
        except Exception as e:
            raise mEx.ChartError("Error getting title from chart") from e

    @classmethod
    def set_subtitle(cls, chart_doc: XChartDocument, subtitle: str, styles: Sequence[StyleT] | None = None) -> XTitle:
        """
        Gets subtitle.

        |lo_unsafe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            subtitle (str): Subtitle text.
            styles (Sequence[StyleT], optional): Styles to apply to subtitle.

        Raises:
            ChartError: If error occurs.

        Returns:
            XTitle: Title object.

        Note:
            The subtitle is set to a font size of ``12`` and the font applied is
            the font returned by :py:meth:`.Info.get_font_general_name`.

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.
        """
        try:
            diagram = chart_doc.getFirstDiagram()
            titled = mLo.Lo.qi(XTitled, diagram, True)
            title = cls._create_title(title=subtitle, font_size=12, styles=styles)
            titled.setTitleObject(title)
            return title
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError(f'Error setting subtitle "{subtitle}" for chart') from e

    @staticmethod
    def get_subtitle(chart_doc: XChartDocument) -> XTitle:
        """
        Gets subtitle from chart.

        |lo_safe|

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
        Gets axis.

        |lo_safe|

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

        |lo_safe|

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

        |lo_safe|

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

        |lo_safe|

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

        |lo_safe|

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
    def set_axis_title(
        cls,
        chart_doc: XChartDocument,
        title: str,
        axis_val: AxisKind,
        idx: int,
        styles: Sequence[StyleT] | None = None,
    ) -> XTitle:
        """
        Sets axis title.

        |lo_unsafe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title text.
            axis_val (AxisKind): Axis kind.
            idx (int): Index
            styles (Sequence[StyleT], optional): Styles to apply to title.

        Raises:
            ChartError: If error occurs.

        Returns:
            XTitle: Axis Title.

        Note:
            The title is set to a font size of ``12`` and the font applied is
            the font returned by :py:meth:`.Info.get_font_general_name`

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.

        See Also:
            :py:meth:`~.Chart2.get_axis`
        """
        try:
            axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
            titled_axis = mLo.Lo.qi(XTitled, axis, True)
            xtitle = cls._create_title(title=title, font_size=12, styles=styles)
            titled_axis.setTitleObject(xtitle)
            fname = mInfo.Info.get_font_general_name()
            cls.set_x_title_font(xtitle, fname, 12)
            return xtitle
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError(f'Error setting axis tile: "{title}" for chart') from e

    @classmethod
    def set_x_axis_title(cls, chart_doc: XChartDocument, title: str, styles: Sequence[StyleT] | None = None) -> XTitle:
        """
        Sets X axis Title.

        |lo_unsafe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title Text.
            styles (Sequence[StyleT], optional): Styles to apply to title.

        Returns:
            XTitle: Title object.

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.

        See Also:
            :py:meth:`~.Chart2.set_axis_title`
        """
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.X, idx=0, styles=styles)

    @classmethod
    def set_y_axis_title(cls, chart_doc: XChartDocument, title: str, styles: Sequence[StyleT] | None = None) -> XTitle:
        """
        Sets Y axis Title.

        |lo_unsafe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title Text.
            styles (Sequence[StyleT], optional): Styles to apply to title.

        Returns:
            XTitle: Title object

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.

        See Also:
            :py:meth:`~.Chart2.set_axis_title`
        """
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.Y, idx=0, styles=styles)

    @classmethod
    def set_x_axis2_title(
        cls, chart_doc: XChartDocument, title: str, styles: Sequence[StyleT] | None = None
    ) -> XTitle:
        """
        Sets X axis2 Title.

        |lo_unsafe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title Text.
            styles (Sequence[StyleT], optional): Styles to apply to title.

        Returns:
            XTitle: Title object.

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.

        See Also:
            :py:meth:`~.Chart2.set_axis_title`
        """
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.X, idx=1, styles=styles)

    @classmethod
    def set_y_axis2_title(
        cls, chart_doc: XChartDocument, title: str, styles: Sequence[StyleT] | None = None
    ) -> XTitle:
        """
        Sets Y axis2 Title.

        |lo_unsafe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            title (str): Title Text.
            styles (Sequence[StyleT], optional): Styles to apply to title.

        Returns:
            XTitle: Title object.

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.

        See Also:
            :py:meth:`~.Chart2.set_axis_title`
        """
        return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.Y, idx=1, styles=styles)

    @classmethod
    def get_axis_title(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int) -> XTitle:
        """
        Gets axis Title.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            axis_val (AxisKind): Axis Kind.
            idx (int): Index.

        Raises:
            ChartError: If error occurs.

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

        |lo_safe|

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

        |lo_safe|

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

        |lo_safe|

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

        |lo_safe|

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
    def rotate_axis_title(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int, angle: Angle | int) -> None:
        """
        Rotates axis title.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            axis_val (AxisKind): Axis kind.
            idx (int): Index
            angle (Angle, int): Angle

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            rotation = Angle(int(angle))
            xtitle = cls.get_axis_title(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
            mProps.Props.set(xtitle, TextRotation=rotation.value)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error while trying to rotate axis title") from e

    @classmethod
    def rotate_x_axis_title(cls, chart_doc: XChartDocument, angle: Angle | int) -> None:
        """
        Rotates X axis title.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            angle (Angle, int): Angle

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`~.Chart2.rotate_axis_title`
        """
        cls.rotate_axis_title(chart_doc=chart_doc, axis_val=AxisKind.X, idx=0, angle=angle)

    @classmethod
    def rotate_y_axis_title(cls, chart_doc: XChartDocument, angle: Angle | int) -> None:
        """
        Rotates Y axis title.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            angle (Angle, int): Angle

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`~.Chart2.rotate_axis_title`
        """
        cls.rotate_axis_title(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=0, angle=angle)

    @classmethod
    def rotate_x_axis2_title(cls, chart_doc: XChartDocument, angle: Angle | int) -> None:
        """
        Rotates X axis2 title.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            angle (Angle, int): Angle

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`~.Chart2.rotate_axis_title`
        """
        cls.rotate_axis_title(chart_doc=chart_doc, axis_val=AxisKind.X, idx=1, angle=angle)

    @classmethod
    def rotate_y_axis2_title(cls, chart_doc: XChartDocument, angle: Angle | int) -> None:
        """
        Rotates Y axis2 title.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            angle (Angle, int): Angle

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

        |lo_safe|

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

        |lo_safe|

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

        |lo_safe|

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

        |lo_safe|

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

        |lo_safe|

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

        |lo_unsafe|

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
                mLo.Lo.print(f'Did not recognize scaling type: "{scale_type}"')
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

        |lo_unsafe|

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

        |lo_unsafe|

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
        Prints axis info to console.

        |lo_safe|

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

        |lo_safe|

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
            raise mEx.UnKnownError("axis_type is of unknown type")

    # endregion Axis

    # region grid lines
    # region set_grid_lines()
    @overload
    @classmethod
    def set_grid_lines(cls, chart_doc: XChartDocument, axis_val: AxisKind) -> XPropertySet: ...

    @overload
    @classmethod
    def set_grid_lines(
        cls, chart_doc: XChartDocument, axis_val: AxisKind, *, styles: Sequence[StyleT]
    ) -> XPropertySet: ...

    @overload
    @classmethod
    def set_grid_lines(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int) -> XPropertySet: ...

    @overload
    @classmethod
    def set_grid_lines(
        cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int, styles: Sequence[StyleT]
    ) -> XPropertySet: ...

    @classmethod
    def set_grid_lines(
        cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int = 0, styles: Sequence[StyleT] | None = None
    ) -> XPropertySet:
        """
        Set the grid lines for a chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            axis_val (AxisKind): Axis kind.
            idx (int, optional): Index. Defaults to ``0``.
            styles (Sequence[StyleT], optional): Styles to apply.

        Raises:
            ChartError: If error occurs.

        Returns:
            XPropertySet: Property Set of Grid Properties.

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.grid </src/format/ooodev.format.chart2.direct.grid>`.

        See Also:
            - :ref:`help_chart2_format_direct_grid_line_properties`
        """
        try:
            axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
            props = axis.getGridProperties()
            mProps.Props.set_property(props, "LineStyle", LineStyle.DASH)
            mProps.Props.set_property(props, "LineDashName", str(LineStyleNameKind.FINE_DOTTED))
            if styles:
                for style in styles:
                    style.apply(props)
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

        |lo_unsafe|

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
                mProps.Props.set(leg, LineStyle=LineStyle.NONE, FillStyle=FillStyle.SOLID, FillTransparence=100)
                diagram.setLegend(leg)
                mProps.Props.set(leg, Show=is_visible)
        except Exception as e:
            raise mEx.ChartError("Error while setting legend visibility") from e

    @staticmethod
    def get_legend(chart_doc: XChartDocument) -> XLegend | None:
        """
        Gets chart legend.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            XLegend: Legend object or ``None`` if no legend exists.

        .. versionadded:: 0.9.4
        """
        try:
            diagram = chart_doc.getFirstDiagram()
            return diagram.getLegend()
        except Exception as e:
            raise mEx.ChartError("Error while getting legend") from e

    # endregion legend

    # region Styles
    @classmethod
    def style_grid(cls, chart_doc: XChartDocument, axis_val: AxisKind, styles: Sequence[StyleT], idx: int = 0) -> None:
        """
        Style Grid.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            axis_val (AxisKind): Axis kind.
            styles (Sequence[StyleT]): Styles to apply.
            idx (int, optional): Index. Defaults to ``0``.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.grid </src/format/ooodev.format.chart2.direct.grid>`.

        See Also:
            - :ref:`help_chart2_format_direct_grid_line_properties`
        """
        with contextlib.suppress(Exception):
            axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
            props = axis.getGridProperties()

            if props is None:
                return

            for style in styles:
                style.apply(props)

    @staticmethod
    def style_background(chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles background of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart background.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in the following:

                - :doc:`ooodev.format.chart2.direct.general.area </src/format/ooodev.format.chart2.direct.general.area>`
                - :doc:`ooodev.format.chart2.direct.general.borders </src/format/ooodev.format.chart2.direct.general.borders>`
                - :doc:`ooodev.format.chart2.direct.general.transparency </src/format/ooodev.format.chart2.direct.general.transparency>`

        See Also:
            - :ref:`help_chart2_format_direct_general_borders`
            - :ref:`help_chart2_format_direct_general_area`
            - :ref:`help_chart2_format_direct_general_transparency`

        .. versionadded:: 0.9.0
        """
        bg_ps = chart_doc.getPageBackground()
        for style in styles:
            style.apply(bg_ps)

    @staticmethod
    def style_wall(chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles Wall of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart wall.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.wall </src/format/ooodev.format.chart2.direct.wall>` subpackages.

        See Also:
            - :ref:`help_chart2_format_direct_wall_floor_area`

        .. versionadded:: 0.9.0
        """
        if wall := chart_doc.getFirstDiagram().getWall():
            for style in styles:
                style.apply(wall)

    @staticmethod
    def style_floor(chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles Floor of 3D chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart floor.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.wall </src/format/ooodev.format.chart2.direct.wall>` subpackages.

        See Also:
            - :ref:`help_chart2_format_direct_wall_floor_area`

        .. versionadded:: 0.9.4
        """
        if floor := chart_doc.getFirstDiagram().getFloor():
            for style in styles:
                style.apply(floor)

    @classmethod
    def style_data_point(cls, chart_doc: XChartDocument, series_idx: int, idx: int, styles: Sequence[StyleT]) -> None:
        """
        Styles a data point of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            series_idx (int): Series Index.
            idx (int): Index to extract from the data points data.
                If ``idx=-1`` then the last data point is styled.
            styles (Sequence[StyleT]): One or more styles to apply chart data point.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in the following packages.

                - :doc:`ooodev.format.chart2.direct.series.data_series </src/format/ooodev.format.chart2.direct.series.data_series>`
                - :doc:`ooodev.format.chart2.direct.series.data_labels </src/format/ooodev.format.chart2.direct.series.data_labels>`

        See Also:
            - :ref:`help_chart2_format_direct_series_series`
            - :ref:`help_chart2_format_direct_series_labels`

        .. versionadded:: 0.9.0
        """

        pp = cls.get_data_point_props(chart_doc=chart_doc, series_idx=series_idx, idx=idx)

        for style in styles:
            style.apply(pp)

    @classmethod
    def style_data_series(cls, chart_doc: XChartDocument, styles: Sequence[StyleT], idx: int = -1) -> None:
        """
        Styles one or more data series of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart data series.
            idx (int, optional): Zero based series index. If value is ``-1`` then styles all data series are styled,
                Otherwise only data series specified by index is styled. Defaults to ``-1``.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in the following packages.

                - :doc:`ooodev.format.chart2.direct.series.data_series </src/format/ooodev.format.chart2.direct.series.data_series>`
                - :doc:`ooodev.format.chart2.direct.series.data_labels </src/format/ooodev.format.chart2.direct.series.data_labels>`

        See Also:
            - :ref:`help_chart2_format_direct_series_series`
            - :ref:`help_chart2_format_direct_series_labels`

        .. versionadded:: 0.9.4
        """
        if idx < 0:
            idx = -1

        series = cls.get_data_series(chart_doc=chart_doc)
        if idx == -1:
            for itm in series:
                for style in styles:
                    style.apply(itm)
        else:
            if idx < 0 or idx >= len(series):
                raise IndexError(f"Index value of {idx} is out of of range")
            itm = series[idx]
            for style in styles:
                style.apply(itm)

    @classmethod
    def style_legend(cls, chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles legend of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart legend.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in the following packages.

                - :doc:`ooodev.format.chart2.direct.legend </src/format/ooodev.format.chart2.direct.legend>`
                - :doc:`ooodev.format.chart2.direct.general.position_size </src/format/ooodev.format.chart2.direct.general.position_size>`

        See Also:
            - :ref:`help_chart2_format_direct_legend`

        .. versionadded:: 0.9.0
        """
        applied_styles = 0
        legend_shape = None
        for style in styles:
            if style.support_service("com.sun.star.drawing.Shape"):
                chart_old = mLo.Lo.qi(XChartDocumentOld, chart_doc, True)
                legend_shape = chart_old.getLegend()
                break
        if legend_shape:
            for style in styles:
                if style.support_service("com.sun.star.drawing.Shape"):
                    style.apply(legend_shape)
                    applied_styles += 1
        if len(styles) == applied_styles:
            return
        if legend := cls.get_legend(chart_doc=chart_doc):
            for style in styles:
                if not style.support_service("com.sun.star.drawing.Shape"):
                    style.apply(legend)

    @classmethod
    def _style_title(cls, xtitle: XTitle, styles: Sequence[StyleT]) -> None:
        """LO Safe Method"""
        # sourcery skip: last-if-guard, move-assign, use-named-expression
        title_styles = [style for style in styles if not style.support_service("com.sun.star.drawing.Shape")]
        applied_styles = 0
        if title_styles:
            for style in title_styles:
                if style.support_service("com.sun.star.chart2.Title"):
                    style.apply(xtitle)
                    applied_styles += 1
            if len(title_styles) == applied_styles:
                return
            fo_strs = xtitle.getText()
            if fo_strs:
                fo_first = fo_strs[0]
                for style in title_styles:
                    if not style.support_service("com.sun.star.chart2.Title"):
                        style.apply(fo_first)

    @classmethod
    def style_title(cls, chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles title of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart title.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.

        See Also:
            - :ref:`help_chart2_format_direct_title`

        .. versionadded:: 0.9.4
        """
        title_styles = [style for style in styles if not style.support_service("com.sun.star.drawing.Shape")]
        if shape_styles := [style for style in styles if style.support_service("com.sun.star.drawing.Shape")]:
            chart_old = mLo.Lo.qi(XChartDocumentOld, chart_doc, True)
            if title_shape := chart_old.getTitle():
                for style in shape_styles:
                    style.apply(title_shape)

        if title_styles:
            xtitle = cls.get_title(chart_doc=chart_doc)
            if xtitle is None:
                return
            cls._style_title(xtitle=xtitle, styles=title_styles)

    @classmethod
    def style_subtitle(cls, chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles subtitle of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart subtitle.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.

        See Also:
            - :ref:`help_chart2_format_direct_title`

        .. versionadded:: 0.9.4
        """
        title_styles = [style for style in styles if not style.support_service("com.sun.star.drawing.Shape")]
        if shape_styles := [style for style in styles if style.support_service("com.sun.star.drawing.Shape")]:
            chart_old = mLo.Lo.qi(XChartDocumentOld, chart_doc, True)
            if subtitle_shape := chart_old.getSubTitle():
                for style in shape_styles:
                    style.apply(subtitle_shape)

        if title_styles:
            if xtitle := cls.get_subtitle(chart_doc=chart_doc):
                cls._style_title(xtitle=xtitle, styles=title_styles)

    @classmethod
    def style_x_axis(cls, chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles X axis of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart X axis.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.axis </src/format/ooodev.format.chart2.direct.axis>` subpackages.

        See Also:
            - :ref:`help_chart2_format_direct_axis`

        .. versionadded:: 0.9.4
        """
        try:
            axis = cls.get_x_axis(chart_doc=chart_doc)
        except mEx.ChartError:
            return

        for style in styles:
            style.apply(axis)

    @classmethod
    def style_x_axis2(cls, chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles X axis2 of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart X axis2.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.axis </src/format/ooodev.format.chart2.direct.axis>` subpackages.

        See Also:
            - :ref:`help_chart2_format_direct_axis`

        .. versionadded:: 0.9.4
        """
        try:
            axis = cls.get_x_axis2(chart_doc=chart_doc)
        except mEx.ChartError:
            return

        for style in styles:
            style.apply(axis)

    @classmethod
    def style_y_axis(cls, chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles Y axis of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart Y axis.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.axis </src/format/ooodev.format.chart2.direct.axis>` subpackages.

        See Also:
            - :ref:`help_chart2_format_direct_axis`

        .. versionadded:: 0.9.4
        """
        try:
            axis = cls.get_y_axis(chart_doc=chart_doc)
        except mEx.ChartError:
            return

        for style in styles:
            style.apply(axis)

    @classmethod
    def style_y_axis2(cls, chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles Y axis2 of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart Y axis2.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.axis </src/format/ooodev.format.chart2.direct.axis>` subpackages.

        See Also:
            - :ref:`help_chart2_format_direct_axis`

        .. versionadded:: 0.9.4
        """
        try:
            axis = cls.get_y_axis2(chart_doc=chart_doc)
        except mEx.ChartError:
            return

        for style in styles:
            style.apply(axis)

    @classmethod
    def style_x_axis_title(cls, chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles X axis title of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart X axis title.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.

        .. versionadded:: 0.9.4
        """
        title_styles = [style for style in styles if not style.support_service("com.sun.star.drawing.Shape")]
        if shape_styles := [style for style in styles if style.support_service("com.sun.star.drawing.Shape")]:
            diagram = chart_doc.getDiagram()  # type: ignore
            if title_shape := diagram.getXAxisTitle():
                for style in shape_styles:
                    style.apply(title_shape)
        if title_styles:
            if xtitle := cls.get_x_axis_title(chart_doc=chart_doc):
                cls._style_title(xtitle=xtitle, styles=title_styles)

    @classmethod
    def style_y_axis_title(cls, chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles X axis title of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart Y axis title.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.

        .. versionadded:: 0.9.4
        """

        title_styles = [style for style in styles if not style.support_service("com.sun.star.drawing.Shape")]
        if shape_styles := [style for style in styles if style.support_service("com.sun.star.drawing.Shape")]:
            diagram = chart_doc.getDiagram()  # type: ignore
            if title_shape := diagram.getYAxisTitle():
                for style in shape_styles:
                    style.apply(title_shape)
        if title_styles:
            if xtitle := cls.get_y_axis_title(chart_doc=chart_doc):
                cls._style_title(xtitle=xtitle, styles=title_styles)

    @classmethod
    def style_x_axis2_title(cls, chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles X axis2 title of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart X axis2 title.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.

        .. versionadded:: 0.9.4
        """
        xtitle = cls.get_x_axis2_title(chart_doc=chart_doc)
        if xtitle is None:
            return
        cls._style_title(xtitle=xtitle, styles=styles)

    @classmethod
    def style_y_axis2_title(cls, chart_doc: XChartDocument, styles: Sequence[StyleT]) -> None:
        """
        Styles X axis2 title of chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            styles (Sequence[StyleT]): One or more styles to apply chart Y axis2 title.

        Returns:
            None:

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>` subpackages.

        .. versionadded:: 0.9.4
        """
        xtitle = cls.get_y_axis2_title(chart_doc=chart_doc)
        if xtitle is None:
            return
        cls._style_title(xtitle=xtitle, styles=styles)

    # endregion Styles

    # region background colors

    @staticmethod
    def set_background_colors(chart_doc: XChartDocument, bg_color: mColor.Color, wall_color: mColor.Color) -> None:
        """
        Set the background colors for a chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            bg_color (~ooodev.utils.color.Color): Color Value for background. If value < ``0`` then no color is set.
            wall_color (~ooodev.utils.color.Color): Color Value for wall. If value < ``0`` then no color is set.

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :py:class:`~.color.CommonColor`
        """
        try:
            if int(bg_color) >= 0:
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

            if int(wall_color) >= 0:
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

        |lo_unsafe|

        Raises:
            ChartError: If error occurs.

        Returns:
            XDataSeries: Data Series instance.
        """
        try:
            return mLo.Lo.create_instance_mcf(XDataSeries, "com.sun.star.chart2.DataSeries", raise_err=True)
        except Exception as e:
            raise mEx.ChartError("Error, unable to create XDataSeries interface") from e

    # region get_data_series()
    @overload
    @classmethod
    def get_data_series(cls, chart_doc: XChartDocument) -> Tuple[XDataSeries, ...]: ...

    @overload
    @classmethod
    def get_data_series(cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase) -> Tuple[XDataSeries, ...]: ...

    @overload
    @classmethod
    def get_data_series(cls, chart_doc: XChartDocument, chart_type: str) -> Tuple[XDataSeries, ...]: ...

    @classmethod
    def get_data_series(
        cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase | str = ""
    ) -> Tuple[XDataSeries, ...]:
        """
        Gets data series for a chart of a given chart type.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document
            chart_type (ChartTypeNameBase, str, optional): Chart Type.

        Raises:
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
                x_chart_type = cls.find_chart_type(chart_doc, chart_type)
            else:
                x_chart_type = cls.get_chart_type(chart_doc)
            ds_con = mLo.Lo.qi(XDataSeriesContainer, x_chart_type, True)
            return ds_con.getDataSeries()
        except Exception as e:
            raise mEx.ChartError("Error getting chart data series") from e

    # endregion get_data_series()

    # region get_data_source()
    @overload
    @classmethod
    def get_data_source(cls, chart_doc: XChartDocument) -> XDataSource: ...

    @overload
    @classmethod
    def get_data_source(cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase) -> XDataSource: ...

    @overload
    @classmethod
    def get_data_source(cls, chart_doc: XChartDocument, chart_type: str) -> XDataSource: ...

    @classmethod
    def get_data_source(cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase | str = "") -> XDataSource:
        """
        Get data source of a chart for a given chart type.

        This method assumes that the programmer wants the first data source in the data series.
        This is adequate for most charts which only use one data source.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            chart_type (ChartTypeNameBase | str): Chart type.

        Raises:
            NotFoundError: If chart is not found.
            ChartError: If any other error occurs.

        Returns:
            XDataSource: Chart data source.

        Hint:
            .. include:: ../../resources/utils/chart2_lookup_chart_name.rst

        See Also:
            :py:meth:`~.Chart2.get_data_series`
        """
        try:
            dsa = cls.get_data_series(chart_doc=chart_doc, chart_type=chart_type)
            return mLo.Lo.qi(XDataSource, dsa[0], True)
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
        Gets coordinate system.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.

        Raises:
            ChartError: If error occurs.

        Returns:
            XCoordinateSystem: Coordinate system object.
        """
        # sourcery skip: merge-nested-ifs
        try:
            diagram = chart_doc.getFirstDiagram()
            coord_sys_con = mLo.Lo.qi(XCoordinateSystemContainer, diagram, True)
            coord_sys = coord_sys_con.getCoordinateSystems()
            if coord_sys:
                if len(coord_sys) > 1:
                    mLo.Lo.print(f"No. of coord systems: {len(coord_sys)}; using first.")
            return coord_sys[0]  # will raise error if coord_sys is empty or none
        except Exception as e:
            raise mEx.ChartError("Error unable to get coord_system") from e

    @classmethod
    def get_chart_types(cls, chart_doc: XChartDocument) -> Tuple[XChartType, ...]:
        """
        Gets chart types for a chart.

        |lo_safe|

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
        Gets a chart type for a chart.

        |lo_safe|

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

        |lo_safe|

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
            if isinstance(chart_type, ChartTypeNameBase):
                srch_name = chart_type.to_namespace().lower()
            else:
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
        Prints chart types to the console.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
        """
        chart_types = cls.get_chart_types(chart_doc)
        if len(chart_types) > 1:
            print(f"No. of chart types: {len(chart_types)}")
            for ct in chart_types:
                print(f"  {ct.getChartType()}")
        else:
            print(f"Chart Type: {chart_types[0].getChartType()}")
        print()

    # region add_chart_type()
    @overload
    @classmethod
    def add_chart_type(cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase) -> XChartType: ...

    @overload
    @classmethod
    def add_chart_type(cls, chart_doc: XChartDocument, chart_type: str) -> XChartType: ...

    @classmethod
    def add_chart_type(cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase | str) -> XChartType:
        """
        Adds a chart type.

        |lo_unsafe|

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
        Prints data information to console.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            data_source (XDataSource): Data Source.

        Returns:
            None:
        """
        dp = chart_doc.getDataProvider()
        ps = dp.detectArguments(data_source)
        mProps.Props.show_props("Data Source arguments", ps)  # type: ignore

    @staticmethod
    def print_labeled_seqs(data_source: XDataSource) -> None:
        """
        Prints labeled sequence information to console.

        |lo_safe|

        A diagnostic function for printing all the labeled data sequences stored in an XDataSource:

        Args:
            data_source (XDataSource): Data Source.

        Returns:
            None:
        """
        data_seqs = data_source.getDataSequences()
        print(f"No. of sequences in data source: {len(data_seqs)}")
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
    def get_chart_data(data_source: XDataSource, idx: int) -> Tuple[float, ...]:
        """
        Gets chart data.

        |lo_safe|

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

            vals_seq = cast(Tuple[float, ...], data_seqs[idx].getValues().getData())
            vals: List[float] = [float(val) for val in vals_seq]
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
        Gets all the properties for the data in the specified series.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            idx (int): Index.

        Raises:
            IndexError: If idx is out of range.

        Returns:
            List[XPropertySet]: Property set list.
        """
        # BUG: there is a bug in the API, getDataPointByIndex() is suppose to return XPropertySet,
        #   which is does, however, it does not properly implement the XPropertySet interface.
        #   setPropertyValue() and getPropertyValue() are not implemented.
        #   The Props.set() method can handle this because is has a fallback to set attributes using the setattr() method of python.
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
        if not props_lst:
            mLo.Lo.print(f"No Data Series at index {idx}")
        return props_lst

    @classmethod
    def get_data_point_props(cls, chart_doc: XChartDocument, series_idx: int, idx: int) -> XPropertySet:
        """
        Get the proprieties for a specific index within the data points.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            series_idx (int): Series Index
            idx (int): Index to extract from the data points data.
                If ``idx=-1`` then the last data point is returned.

        Raises:
            NotFoundError: If ``series_idx`` did not find any data points.
            IndexError: If an index is out of range.

        Returns:
            XPropertySet: A single property set from the data points series.

        See Also:
            :py:meth:`~.Chart2.get_data_points_props`

        .. versionchanged:: 0.9.0
            ``idx`` value of ``-1`` returns last data point.
        """
        props = cls.get_data_points_props(chart_doc=chart_doc, idx=series_idx)
        if not props:
            raise mEx.NotFoundError("No Data points found to get XPropertySet from")

        if idx == -1:
            return props.pop()

        if idx < 0 or idx >= len(props):
            raise IndexError(
                f"Index value of {idx} is out of of range; use 0 to {len(props) - 1} or -1 for last item."
            )

        return props[idx]

    @classmethod
    def set_data_point_labels(cls, chart_doc: XChartDocument, label_type: DataPointLabelTypeKind) -> None:
        """
        Sets the data point label of a chart.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            label_type (DataPointLabelTypeKind): Label Type.

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            data_series_arr = cls.get_data_series(chart_doc=chart_doc)
            for data_series in data_series_arr:
                dp_label = cast("DataPointLabel", mProps.Props.get_property(data_series, "Label"))
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
                elif label_type != DataPointLabelTypeKind.NONE:
                    raise mEx.UnKnownError("label_type is of unknown type")

                mProps.Props.set_property(data_series, "Label", dp_label)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error setting data point labels") from e

    @classmethod
    def set_chart_shape_3d(cls, chart_doc: XChartDocument, shape: DataPointGeometry3DEnum) -> None:
        """
        Sets chart 3d shape.

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            shape (DataPointGeometry3DEnum): Shape kind.

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

        |lo_safe|

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

        |lo_safe|

        Args:
            ct (XChartType): Chart Type.
            w_day_color (~ooodev.utils.color.Color): Chart white day color
            b_day_color (~ooodev.utils.color.Color): Chart black day color

        Raises:
            NotSupportedError: If Chart is not of type ``CandleStickChartType``
            ChartError: If any other error occurs.

        Returns:
            None:

        See Also:
            :py:class:`~.color.CommonColor`
        """
        # sourcery skip: remove-unnecessary-else, swap-if-else-branches
        try:
            if ct.getChartType() == "com.sun.star.chart2.CandleStickChartType":
                # there is a bug with white_day_ps and black_day_ps
                # they are suppose br XProperySet, which they are but
                # XProperySet methods such as  setPropertyValue and getPropertySetInfo are missing.
                # see also: Props._set_by_attribute()
                white_day_ps = mLo.Lo.qi(XPropertySet, mProps.Props.get(ct, "WhiteDay"), True)
                mProps.Props.set(white_day_ps, FillColor=int(w_day_color))

                black_day_ps = mLo.Lo.qi(XPropertySet, mProps.Props.get(ct, "BlackDay"), True)
                mProps.Props.set(black_day_ps, FillColor=int(b_day_color))
            else:
                raise mEx.NotSupportedError(
                    f'Only candle stick charts supported. "{ct.getChartType()}" not supported.'
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

        |lo_unsafe|

        Args:
            curve_kind (CurveKind): Curve kind.

        Raises:
            ChartError: If error occurs.

        Returns:
            XRegressionCurve: Regression Curve object.
        """
        try:
            return mLo.Lo.create_instance_mcf(XRegressionCurve, curve_kind.to_namespace(), raise_err=True)
        except Exception as e:
            raise mEx.ChartError("Error creating curve") from e

    @classmethod
    def draw_regression_curve(
        cls, chart_doc: XChartDocument, curve_kind: CurveKind, styles: Sequence[StyleT] | None = None
    ) -> XPropertySet:
        """
        Draws a regression curve.

        |lo_unsafe|

        Args:
            chart_doc (XChartDocument): Chart Document
            curve_kind (CurveKind): Curve kind.
            styles (Sequence[StyleT], optional): Styles to apply to the curve. Defaults to ``None``.

        Raises:
            ChartError: If error occurs.

        Returns:
            XPropertySet: Regression curve property set.

        Hint:
            Styles that can be applied are found in the following subpackages:

                - :doc:`ooodev.format.chart2.direct.title </src/format/ooodev.format.chart2.direct.title>`
                - :doc:`ooodev.format.chart2.direct.general.numbers </src/format/ooodev.format.chart2.direct.general.numbers>`

        .. versionchanged:: 0.9.4
            Added ``styles`` argument, and now returns the regression curve property set.
        """
        try:
            data_series_arr = cls.get_data_series(chart_doc=chart_doc)
            rc_con = mLo.Lo.qi(XRegressionCurveContainer, data_series_arr[0], True)
            curve = cls.create_curve(curve_kind)
            rc_con.addRegressionCurve(curve)

            ps = curve.getEquationProperties()
            mProps.Props.set(ps, ShowCorrelationCoefficient=True, ShowEquation=True)

            key = cls.get_number_format_key(chart_doc=chart_doc, nf_str="0.00")  # 2 dp
            if key != -1:
                mProps.Props.set(ps, NumberFormat=key)
            if styles:
                supported = (
                    "com.sun.star.chart2.RegressionEquation",
                    "com.sun.star.drawing.FillProperties",
                    "com.sun.star.drawing.LineProperties",
                    "com.sun.star.style.CharacterProperties",
                )

                for style in styles:
                    if style.support_service(*supported):
                        style.apply(ps)
            return ps
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error drawing regression curve") from e

    @staticmethod
    def get_number_format_key(chart_doc: XChartDocument, nf_str: str) -> int:
        """
        Converts a number format string into a number format key, which can be assigned to
        ``NumberFormat`` property.

        |lo_safe|

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
            # locale = Locale("en", "us", "")
            # locale = mInfo.Info.language_locale
            # note the empty locale for default locale
            key = int(n_formats.queryKey(nf_str, Locale(), False))
            if key == -1:
                mLo.Lo.print(f'Could not access key for number format: "{nf_str}"')
            return key
        except Exception as e:
            raise mEx.ChartError("Error getting number format key") from e

    @staticmethod
    def get_curve_type(curve: XRegressionCurve) -> CurveKind:
        """
        Gets curve kind from regression object.

        |lo_safe|

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

        |lo_safe|

        Args:
            chart_doc (XChartDocument): Chart Document
            curve (XRegressionCurve): Regression Curve object.

        Returns:
            None:
        """
        curve_calc = curve.getCalculator()
        ct = cls.get_curve_type(curve)
        degree = 2 if ct != CurveKind.LINEAR else 1
        # degree, forceIntercept, interceptValue, period (for moving average)
        # the last are for setRegressionProperties is movingType
        #   See: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1chart2_1_1MovingAverageType.html
        # movingType Only if regression type is "Moving Average" 1, 3 or 3
        #   see: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XRegressionCurveCalculator.html#ae65112de1214e140d9ce7b28ffb09292
        # Because thi sis not a Moving Average setting to 0
        curve_calc.setRegressionProperties(degree, False, 0.0, 2, 0)

        data_source = cls.get_data_source(chart_doc)

        x_vals = cls.get_chart_data(data_source=data_source, idx=0)
        y_vals = cls.get_chart_data(data_source=data_source, idx=0)
        curve_calc.recalculateRegression(x_vals, y_vals)

        print(f"  Curve equations: {curve_calc.getRepresentation()}")
        cc = curve_calc.getCorrelationCoefficient()
        print(f"  R^2 value: {(cc*cc):.3f}")

    @classmethod
    def calc_regressions(cls, chart_doc: XChartDocument) -> None:
        """
        Several different regression functions are calculated using the chart's data.
        Their equations and ``R2`` values are printed to the console

        |lo_unsafe|

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

        |lo_unsafe|

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
        # create labeled data sequence using label and data;
        # the data is for the specified role

        try:
            # create data sequence for the label
            lbl_seq = dp.createDataSequenceByRangeRepresentation(data_label)

            # create data sequence for the data and role
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
        Sets Y error Bars.

        |lo_unsafe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            data_label (str): Data Label.
            data_range (str): Data Range.

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
            # print(f'No. of data series: {len(data_series_arr)}')
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

        |lo_unsafe|

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

            mProps.Props.set(ds, Color=int(mColor.CommonColor.RED))
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
        Add Category Labels.

        |lo_unsafe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            data_label (str): Data label.
            data_range (str): Data range.

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

            # label the data points with these category values
            cls.set_data_point_labels(chart_doc=chart_doc, label_type=DataPointLabelTypeKind.CATEGORY)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error adding category labels") from e

    # endregion add data to a chart

    # region chart shape and image
    @staticmethod
    def get_chart_shape(sheet: XSpreadsheet, chart_name: str = "") -> XShape:
        """
        Gets chart shape.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            chart_name (str, optional): Chart name. Defaults to "". If ``""`` then first chart is returned.

        Raises:
            ShapeMissingError: If shape is ``None``.
            ShapeError: If any other error occurs.

        Returns:
            XShape: Shape object.

        .. versionchanged:: 0.9.4
            Added chart_name parameter.
        """
        shape = None
        if chart_name:
            chart_name = chart_name.casefold()
        try:
            page_supp = mLo.Lo.qi(XDrawPageSupplier, sheet, True)
            draw_page = page_supp.getDrawPage()
            num_shapes = draw_page.getCount()
            chart_class_id = mLo.Lo.CLSID.CHART.value
            for i in range(num_shapes):
                try:
                    shape = cast("OLE2Shape", mLo.Lo.qi(XShape, draw_page.getByIndex(i), True))
                    class_id = str(mProps.Props.get(shape, "CLSID")).lower()
                    if chart_name:
                        if class_id == chart_class_id and chart_name == shape.PersistName.casefold():
                            break
                    elif class_id == chart_class_id:
                        break
                except Exception:
                    shape = None
                    # continue on, just because got an error does not mean shape will not be found
        except Exception as e:
            raise mEx.ShapeError("Error getting shape from sheet") from e
        if shape is None:
            raise mEx.ShapeMissingError("Unable to find Chart Shape")
        return shape

    @classmethod
    def copy_chart(cls, ssdoc: XSpreadsheetDocument, sheet: XSpreadsheet, chart_name: str = "") -> None:
        """
        Copies a chart to the clipboard using a dispatch command.

        |lo_unsafe|

        Args:
            ssdoc (XSpreadsheetDocument): Spreadsheet Document.
            sheet (XSpreadsheet): Spreadsheet.
            chart_name (str, optional): Chart name. Defaults to "". If ``""`` then first chart is copied.

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        .. versionchanged:: 0.9.4
            Added chart_name parameter.
        """
        try:
            chart_shape = cls.get_chart_shape(sheet=sheet, chart_name=chart_name)
            doc = mLo.Lo.qi(XComponent, ssdoc, True)
            supp = mGui.GUI.get_selection_supplier(doc)
            supp.select(chart_shape)
            mLo.Lo.dispatch_cmd("Copy")
        except Exception as e:
            raise mEx.ChartError("Error in attempt to copy chart") from e

    @classmethod
    def get_chart_draw_page(cls, sheet: XSpreadsheet, chart_name: str = "") -> XDrawPage:
        """
        Gets chart draw page.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            chart_name (str, optional): Chart name. Defaults to "". If ``""`` then first chart Draw Page is returned.

        Raises:
            ChartError: If error occurs.

        Returns:
            XDrawPage: Draw Page object

        .. versionchanged:: 0.9.4
            Added chart_name parameter.
        """
        try:
            chart_shape = cls.get_chart_shape(sheet=sheet, chart_name=chart_name)
            embedded_chart = mLo.Lo.qi(XEmbeddedObject, mProps.Props.get_property(chart_shape, "EmbeddedObject"), True)
            comp_supp = mLo.Lo.qi(XComponentSupplier, embedded_chart, True)
            x_closable = comp_supp.getComponent()
            supp_page = mLo.Lo.qi(XDrawPageSupplier, x_closable, True)
            result = supp_page.getDrawPage()
            if result is None:
                raise mEx.UnKnownError("None Value: getDrawPage() returned a value of None")
            return result
        except Exception as e:
            raise mEx.ChartError("Error getting chart draw page") from e

    @classmethod
    def get_chart_image(cls, sheet: XSpreadsheet, chart_name: str = "") -> XGraphic:
        """
        Get chart image as ``XGraphic``.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            chart_name (str, optional): Chart name. Defaults to "". If ``""`` then first chart image is returned.

        Raises:
            ChartError: If error occurs.

        Returns:
            XGraphic: Graphic object

        .. versionchanged:: 0.9.4
            Added chart_name parameter.
        """
        try:
            chart_shape = cls.get_chart_shape(sheet=sheet, chart_name=chart_name)

            graphic = mLo.Lo.qi(XGraphic, mProps.Props.get(chart_shape, "Graphic"), True)

            tmp_fnm = mFileIo.FileIO.create_temp_file("png")
            mImgLo.ImagesLo.save_graphic(pic=graphic, fnm=tmp_fnm, im_format="png")
            im = mImgLo.ImagesLo.load_graphic_file(tmp_fnm)
            mFileIo.FileIO.delete_file(tmp_fnm)
            return im
        except Exception as e:
            raise mEx.ChartError("Error getting chart image") from e

    # endregion chart shape and image

    # region Lock Controllers

    @classmethod
    def lock_controllers(cls, chart_doc: XChartDocument) -> bool:
        """
        Suspends some notifications to the controllers which are used for display updates.

        The calls to :py:meth:`~.chart2.Chart2.lock_controllers` and :py:meth:`~.chart2.Chart2.unlock_controllers`
        may be nested and even overlapping, but they must be in pairs.
        While there is at least one lock remaining, some notifications for
        display updates are not broadcast.

        |lo_safe|

        Returns:
            bool: False if ``CONTROLLERS_LOCKING`` event is canceled; Otherwise, True

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.chart2_named_event.Chart2NamedEvent.CONTROLLERS_LOCKING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.chart2_named_event.Chart2NamedEvent.CONTROLLERS_LOCKED` :eventref:`src-docs-event`

        See Also:

            - :py:class:`~.chart2.Chart2ControllerLock`
            - :py:meth:`~.chart2.Chart2.unlock_controllers`

        .. versionadded:: 0.9.4
        """
        # much faster updates as screen is basically suspended
        cargs = CancelEventArgs(Chart2.lock_controllers.__qualname__)
        _Events().trigger(Chart2NamedEvent.CONTROLLERS_LOCKING, cargs)
        if cargs.cancel:
            return False
        chart_doc.lockControllers()
        _Events().trigger(Chart2NamedEvent.CONTROLLERS_LOCKED, EventArgs(cls))
        return True

    @staticmethod
    def unlock_controllers(chart_doc: XChartDocument) -> bool:
        """
        Resumes the notifications which were suspended by :py:meth:`~.chart2.Chart2.lock_controllers`.

        The calls to :py:meth:`~.chart2.Chart2.lock_controllers` and :py:meth:`~.chart2.Chart2.unlock_controllers`
        may be nested and even overlapping, but they must be in pairs.
        While there is at least one lock remaining, some notifications for
        display updates are not broadcast.

        |lo_safe|

        Returns:
            bool: False if ``CONTROLLERS_UNLOCKING`` event is canceled; Otherwise, True

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.chart2_named_event.Chart2NamedEvent.CONTROLLERS_UNLOCKING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.chart2_named_event.Chart2NamedEvent.CONTROLLERS_UNLOCKED` :eventref:`src-docs-event`

        See Also:

            - :py:class:`~.chart2.Chart2ControllerLock`
            - :py:meth:`~.chart2.Chart2.lock_controllers`

        .. versionadded:: 0.9.4
        """
        cargs = CancelEventArgs(Chart2.unlock_controllers.__qualname__)
        _Events().trigger(Chart2NamedEvent.CONTROLLERS_UNLOCKING, cargs)
        if cargs.cancel:
            return False
        if chart_doc.hasControllersLocked():
            chart_doc.unlockControllers()
        _Events().trigger(Chart2NamedEvent.CONTROLLERS_UNLOCKED, EventArgs.from_args(cargs))
        return True

    # endregion Lock Controllers


class Chart2ControllerLock:
    """
    Context manager for Locking Chart2 Controller

    In the following example ControllerLock is called using ``with``.

    All code inside the ``with Chart2ControllerLock.ControllerLock(chart_doc)`` block is updated
    with controller locked. This means the ui will not update chart until the block is done.
    A soon as the block is processed the controller is unlocked and the ui is updated.

    Example:

        .. code::

            from ooodev.utils.color import CommonColor
            from ooodev.office.chart2 import Chart2
            from ooodev.office.chart2 import Chart2ControllerLock
            from ooodev.format.chart2.direct.title.area import Color as ChartTitleBgColor
            from ooodev.format.chart2.direct.title.font import Font as TitleFont
            from ooodev.format.chart2.direct.title.borders import LineProperties as TitleBorderLineProperties, BorderLineKind
            # ... other imports

            with Chart2ControllerLock.ControllerLock(chart_doc):
                title_area_bg_color = ChartTitleBgColor(CommonColor.LIGHT_YELLOW)
                title_font = TitleFont(b=True, size=14, color=CommonColor.DARK_GREEN)
                title_border = TitleBorderLineProperties(style=BorderLineKind.DASH_DOT, width=1.0, color=CommonColor.DARK_RED)
                Chart2.style_title(chart_doc=chart_doc, styles=[title_area_bg_color, title_font, title_border])
                # ... other code

    .. versionadded:: 0.9.4
    """

    def __init__(self, chart_doc: XChartDocument):
        self._chart_doc = chart_doc
        Chart2.lock_controllers(chart_doc=self._chart_doc)

    def __enter__(self) -> XChartDocument:
        return self._chart_doc

    def __exit__(self, exc_type, exc_val, exc_tb):
        Chart2.unlock_controllers(self._chart_doc)


__all__ = ("Chart2", "Chart2ControllerLock")
