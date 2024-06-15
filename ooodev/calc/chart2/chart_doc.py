from __future__ import annotations
from typing import Any, overload, List, Sequence, TYPE_CHECKING, Tuple
from ooodev.mock import mock_g
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.chart2.chart_document_comp import ChartDocumentComp
from ooodev.adapter.document.storage_change_event_events import StorageChangeEventEvents
from ooodev.adapter.util.close_events import CloseEvents
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.partial.area.fill_color_partial import FillColorPartial
from ooodev.format.inner.partial.area.transparency.gradient_partial import GradientPartial
from ooodev.format.inner.partial.area.transparency.transparency_partial import TransparencyPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_gradient_partial import ChartFillGradientPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_hatch_partial import ChartFillHatchPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_img_partial import ChartFillImgPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_pattern_partial import ChartFillPatternPartial
from ooodev.format.inner.partial.chart2.borders.border_line_properties_partial import BorderLinePropertiesPartial
from ooodev.loader import lo as mLo
from ooodev.office import chart2 as mChart2
from ooodev.utils.color import Color
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.calc.chart2.kind.chart_axis_kind import ChartAxisKind
from ooodev.calc.chart2.kind.chart_title_kind import ChartTitleKind
from ooodev.calc.chart2.kind.chart_diagram_kind import ChartDiagramKind
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from com.sun.star.chart2 import ChartDocument  # service
    from com.sun.star.chart2 import XRegressionCurve
    from ooodev.events.args.cancel_event_args import CancelEventArgs
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.proto.style_obj import StyleT
    from ooodev.utils.comp.prop import Prop
    from ooodev.utils.kind.chart2_types import ChartTypeNameBase
    from ooodev.utils.kind.curve_kind import CurveKind
    from ooodev.utils.kind.data_point_label_type_kind import DataPointLabelTypeKind
    from ooodev.calc.chart2.chart_axis import ChartAxis
    from ooodev.calc.chart2.chart_data_series import ChartDataSeries
    from ooodev.calc.chart2.chart_diagram import ChartDiagram
    from ooodev.calc.chart2.chart_error_bar import ChartErrorBar
    from ooodev.calc.chart2.chart_title import ChartTitle
    from ooodev.calc.chart2.chart_type import ChartType
    from ooodev.calc.chart2.table_chart import TableChart
    from ooodev.calc.chart2.coordinate.coordinate_general import CoordinateGeneral
    from ooodev.calc.chart2.regression_curve.regression_curve import RegressionCurve
    from ooodev.calc.chart2.data.data_provider import DataProvider
else:
    CoordinateGeneral = Any
    StyleT = Any
    RegressionCurve = Any
    Prop = Any


class ChartDoc(
    LoInstPropsPartial,
    ChartDocumentComp,
    ModifyEvents,
    PropPartial,
    QiPartial,
    ServicePartial,
    TheDictionaryPartial,
    EventsPartial,
    ChartDocPropPartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
    CloseEvents,
    StorageChangeEventEvents,
    FillColorPartial,
    ChartFillGradientPartial,
    ChartFillImgPartial,
    ChartFillPatternPartial,
    ChartFillHatchPartial,
    BorderLinePropertiesPartial,
    TransparencyPartial,
    GradientPartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
):
    """
    Class for managing Chart2 ChartDocument Component.
    """

    def __init__(self, owner: TableChart, component: ChartDocument, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Chart2 ChartDocument Component.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        ChartDocumentComp.__init__(self, component=component)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        TheDictionaryPartial.__init__(self)
        EventsPartial.__init__(self)
        ChartDocPropPartial.__init__(self, chart_doc=self)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore # pylint: disable=no-member
        ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        CloseEvents.__init__(self, trigger_args=generic_args, cb=self._on_close_events_add_remove)
        StorageChangeEventEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_storage_change_events_add_remove
        )
        PropertyChangeImplement.__init__(self, component=component, trigger_args=generic_args)  # type: ignore
        VetoableChangeImplement.__init__(self, component=component, trigger_args=generic_args)  # type: ignore
        FillColorPartial.__init__(self, factory_name="ooodev.chart2.general", component=component, lo_inst=lo_inst)
        pg_bg = self.component.getPageBackground()
        ChartFillGradientPartial.__init__(self, factory_name="ooodev.chart2.general", component=pg_bg, lo_inst=lo_inst)
        ChartFillImgPartial.__init__(self, factory_name="ooodev.chart2.general", component=pg_bg, lo_inst=lo_inst)
        ChartFillPatternPartial.__init__(self, factory_name="ooodev.chart2.general", component=pg_bg, lo_inst=lo_inst)
        ChartFillHatchPartial.__init__(self, factory_name="ooodev.chart2.general", component=pg_bg, lo_inst=lo_inst)
        BorderLinePropertiesPartial.__init__(self, factory_name="ooodev.chart2.line", component=pg_bg, lo_inst=lo_inst)
        TransparencyPartial.__init__(self, factory_name="ooodev.chart2.general", component=pg_bg, lo_inst=lo_inst)
        GradientPartial.__init__(self, factory_name="ooodev.chart2.general", component=pg_bg, lo_inst=lo_inst)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)
        CalcSheetPropPartial.__init__(self, obj=owner.calc_sheet)
        self._owner = owner
        self._axis_x = None
        self._axis2_x = None
        self._axis_y = None
        self._axis2_y = None
        self._first_diagram = None
        self._init_events()

    # region Lazy Listeners

    def _on_modify_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addModifyListener(self.events_listener_modify)  # type: ignore
        event.remove_callback = True

    def _on_close_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addCloseListener(self.events_listener_close)  # type: ignore
        event.remove_callback = True

    def _on_storage_change_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addStorageChangeListener(self.events_listener_storage_change_event)  # type: ignore
        event.remove_callback = True

    # endregion Lazy Listeners

    # region context manage
    def __enter__(self) -> ChartDoc:
        self.lock_controllers()
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.unlock_controllers()

    # endregion context manage

    # region Events
    def _init_events(self) -> None:
        """
        Initialize Events
        """
        self._fn_on_area_fill_color_changing = self._on_area_fill_color_changing
        self.subscribe_event("before_style_area_color", self._fn_on_area_fill_color_changing)
        self.subscribe_event("before_style_area_color_get", self._fn_on_area_fill_color_changing)

    def _on_area_fill_color_changing(self, source: Any, event: CancelEventArgs) -> None:
        """
        On Area Fill Color Changing
        """
        event.event_data["this_component"] = self.component.getPageBackground()

    # endregion Events

    # region GradientPartial Overrides
    def _GradientPartial_transparency_get_chart_doc(self) -> XChartDocument | None:
        return self.component

    # endregion GradientPartial Overrides

    # region ChartDocumentPartial Overrides
    def get_data_provider(self) -> DataProvider:
        """
        Returns the currently set data provider.

        This may be an internal one, if createInternalDataProvider() has been called before, or an external one if XDataReceiver.attachDataProvider() has been called.
        """
        # pylint: disable=import-outside-toplevel
        from .data.data_provider import DataProvider

        dp = self.component.getDataProvider()
        return DataProvider(owner=self, component=dp, lo_inst=self.lo_inst)

    # endregion ChartDocumentPartial Overrides

    # region Methods
    def set_title(self, title: str) -> ChartTitle[ChartDoc]:
        """Adds a Chart Title."""
        # pylint: disable=import-outside-toplevel
        from com.sun.star.chart2 import XTitled
        from com.sun.star.chart2 import XTitle
        from com.sun.star.chart2 import XFormattedString
        from .chart_title import ChartTitle

        x_title = self.lo_inst.create_instance_mcf(XTitle, "com.sun.star.chart2.Title", raise_err=True)
        x_title_str = self.lo_inst.create_instance_mcf(
            XFormattedString, "com.sun.star.chart2.FormattedString", raise_err=True
        )
        x_title_str.setString(title)

        title_arr = (x_title_str,)
        x_title.setText(title_arr)

        titled = self.qi(XTitled, True)
        titled.setTitleObject(x_title)
        return ChartTitle(
            owner=self,
            chart_doc=self,
            component=titled.getTitleObject(),
            title_kind=ChartTitleKind.TITLE,
            lo_inst=self.lo_inst,
        )

    def get_title(self) -> ChartTitle[ChartDoc] | None:
        """Gets the Chart Title Component."""
        # pylint: disable=import-outside-toplevel
        from com.sun.star.chart2 import XTitled
        from .chart_title import ChartTitle

        titled = self.qi(XTitled, True)
        comp = titled.getTitleObject()
        if comp is None:
            return None
        return ChartTitle(
            owner=self, chart_doc=self, component=comp, title_kind=ChartTitleKind.TITLE, lo_inst=self.lo_inst
        )

    def set_bg_color(self, color: Color) -> None:
        """Sets the background color."""
        mChart2.Chart2.set_background_colors(self.component, bg_color=color, wall_color=Color(-1))

    def set_wall_color(self, color: Color) -> None:
        """Sets the wall color."""
        mChart2.Chart2.set_background_colors(self.component, bg_color=Color(-1), wall_color=color)

    # region get_data_series()

    @overload
    def get_data_series(self) -> Tuple[ChartDataSeries[ChartDoc], ...]:
        """
        Gets data series for a chart of a given chart type.

        Returns:
            Tuple[XDataSeries, ...]: Data Series
        """
        ...

    @overload
    def get_data_series(self, chart_type: ChartTypeNameBase) -> Tuple[ChartDataSeries[ChartDoc], ...]:
        """
        Gets data series for a chart of a given chart type.

        Args:
            chart_type (ChartTypeNameBase): Chart Type.

        Returns:
            Tuple[XDataSeries, ...]: Data Series

        See Also:
            :ref:`ooodev.utils.kind.chart2_types`
        """
        ...

    @overload
    def get_data_series(self, chart_type: str) -> Tuple[ChartDataSeries[ChartDoc], ...]:
        """
        Gets data series for a chart of a given chart type.

        Args:
            chart_type (str): Chart Type.

        Returns:
            Tuple[XDataSeries, ...]: Data Series
        """
        ...

    def get_data_series(self, chart_type: ChartTypeNameBase | str = "") -> Tuple[ChartDataSeries[ChartDoc], ...]:
        """
        Gets data series for a chart of a given chart type.

        Args:
            chart_type (ChartTypeNameBase, str, optional): Chart Type.

        Raises:
            ChartError: If any other error occurs.

        Returns:
            Tuple[XDataSeries, ...]: Data Series

        See Also:
            - :py:meth:`ooodev.office.chart2.get_data_series`
            - :ref:`ooodev.utils.kind.chart2_types`
        """
        # pylint: disable=import-outside-toplevel
        from .chart_data_series import ChartDataSeries

        data_series = mChart2.Chart2.get_data_series(chart_doc=self.component, chart_type=chart_type)
        series = tuple(
            ChartDataSeries(owner=self, chart_doc=self, component=comp, lo_inst=self.lo_inst) for comp in data_series
        )
        return series  # type: ignore

    # endregion get_data_series()

    def get_templates(self) -> List[str]:
        """
        Gets a list of chart templates (services).

        Raises:
            ChartError: If error occurs

        Returns:
            List[str]: List of chart templates
        """
        return mChart2.Chart2.get_chart_templates(self.component)

    def set_y_error_bar(self, data_label: str, data_range: str) -> ChartErrorBar:
        """
        Set Error Bar

        Raises:
            ChartError: If any error occurs.

        Returns:
            ChartErrorBar: Chart Error Bar
        """
        # pylint: disable=import-outside-toplevel
        from .chart_error_bar import ChartErrorBar
        from ooo.dyn.chart.error_bar_style import ErrorBarStyle
        from ooodev.utils.kind.chart2_data_role_kind import DataRoleKind

        try:
            eb = ChartErrorBar(chart_doc=self, lo_inst=self.lo_inst)
            eb.set_property(ShowPositiveError=True, ShowNegativeError=True, ErrorBarStyle=ErrorBarStyle.FROM_DATA)
            dp = self.get_data_provider()
            with LoContext(self.lo_inst):
                pos_err_seq = mChart2.Chart2.create_ld_seq(
                    dp=dp.component,
                    role=DataRoleKind.ERROR_BARS_Y_POSITIVE,
                    data_label=data_label,
                    data_range=data_range,
                )
                neg_err_seq = mChart2.Chart2.create_ld_seq(
                    dp=dp.component,
                    role=DataRoleKind.ERROR_BARS_Y_NEGATIVE,
                    data_label=data_label,
                    data_range=data_range,
                )
            ld_seq = (pos_err_seq, neg_err_seq)
            eb.set_data(ld_seq)
            data_series = self.get_data_series()[0]
            data_series.set_property(ErrorBarY=eb.component)
            return eb
        except Exception as e:
            raise mEx.ChartError("Error setting error bar", e)

    def add_stock_line(self, data_label: str, data_range: str) -> None:
        """
        Add Stock Line

        Args:
            data_label (str): Data Label
            data_range (str): Data Range

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        with LoContext(self.lo_inst):
            mChart2.Chart2.add_stock_line(self.component, data_label, data_range)

    def add_chart_type(self, chart_type: ChartTypeNameBase | str) -> ChartType[CoordinateGeneral]:
        """
        Add Chart Type

        Args:
            chart_type (ChartTypeNameBase): Chart Type

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        See Also:
            :ref:`ooodev.utils.kind.chart2_types`
        """
        # pylint: disable=import-outside-toplevel
        from .chart_type import ChartType
        from com.sun.star.chart2 import XChartType

        try:
            ct = self.lo_inst.create_instance_mcf(XChartType, f"com.sun.star.chart2.{chart_type}", raise_err=True)
            coord_sys = self.first_diagram.get_coordinate_system()
            if coord_sys is None:
                raise mEx.ChartError("Coordinate System not found")
            coord_sys.add_chart_type(ct)
            return ChartType(owner=coord_sys, chart_doc=self, component=ct, lo_inst=self.lo_inst)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error adding chart type", e)

    def add_cat_labels(self, data_label: str, data_range: str) -> None:
        """
        Add Category Labels.

        Args:
            data_label (str): Data label.
            data_range (str): Data range.

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        dp = self.get_data_provider()
        dp.add_cat_labels(data_label, data_range)

    def create_curve(self, curve_kind: CurveKind) -> RegressionCurve:
        """
        Creates a regression curve.

        Matches the regression constants defined in ``curve_kind`` to regression services offered by the API:

        Args:
            curve_kind (CurveKind): Curve kind.

        Raises:
            ChartError: If error occurs.

        Returns:
            XRegressionCurve: Regression Curve object.

        Hint:
            - ``CurveKind`` can be imported from ``ooodev.utils.kind.curve_kind``.
        """
        # pylint: disable=import-outside-toplevel
        from .regression_curve.regression_curve import RegressionCurve
        from com.sun.star.chart2 import XRegressionCurve

        try:
            curve = self.lo_inst.create_instance_mcf(XRegressionCurve, curve_kind.to_namespace(), raise_err=True)
            return RegressionCurve(owner=self, component=curve, lo_inst=self.lo_inst)
        except Exception as e:
            raise mEx.ChartError("Error creating curve") from e

    def get_number_format_key(self, nf_str: str) -> int:
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
        # pylint: disable=import-outside-toplevel
        from ooo.dyn.lang.locale import Locale

        try:
            n_formats = self.get_number_formats()
            # locale = Locale("en", "us", "")
            # locale = mInfo.Info.language_locale
            # note the empty locale for default locale
            key = int(n_formats.queryKey(nf_str, Locale(), False))
            if key == -1:
                self.lo_inst.print(f'Could not access key for number format: "{nf_str}"')
            return key
        except Exception as e:
            raise mEx.ChartError("Error getting number format key") from e

    def dash_lines(self) -> None:
        """
        Sets chart data series to dashed lines.

        Args:
            chart_doc (XChartDocument): Chart Document

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        mChart2.Chart2.dash_lines(self.component)

    def draw_regression_curve(
        self, curve_kind: CurveKind, styles: Sequence[StyleT] | None = None
    ) -> Prop[RegressionCurve]:
        """
        Draws a regression curve.

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

            - ``CurveKind`` can be imported from ``ooodev.utils.kind.curve_kind``.

        .. versionchanged:: 0.9.4
            Added ``styles`` argument, and now returns the regression curve property set.
        """
        # pylint: disable=import-outside-toplevel
        from com.sun.star.chart2 import XRegressionCurveContainer

        try:
            data_series_arr = self.get_data_series()
            rc_con = self.lo_inst.qi(XRegressionCurveContainer, data_series_arr[0].component, True)
            curve = self.create_curve(curve_kind)
            rc_con.addRegressionCurve(curve.component)

            ps = curve.get_equation_properties()
            ps.set_property(ShowCorrelationCoefficient=True, ShowEquation=True)

            key = self.get_number_format_key(nf_str="0.00")  # 2 dp
            if key != -1:
                ps.set_property(NumberFormat=key)
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

    def eval_curve(self, curve: XRegressionCurve) -> None:
        """
        Uses ``XRegressionCurve.getCalculator()`` to access the ``XRegressionCurveCalculator`` interface.
        It sets up the data and parameters for a particular curve, and prints the results of curve fitting to the console.

        Args:
            curve (XRegressionCurve): Regression Curve object.

        Returns:
            None:
        """
        mChart2.Chart2.eval_curve(self.component, curve)

    def calc_regressions(self) -> None:
        """
        Calculate regressions.

        Several different regression functions are calculated using the chart's data.
        Their equations and ``R2`` values are printed to the console

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        with LoContext(self.lo_inst):
            mChart2.Chart2.calc_regressions(self.component)

    def find_chart_type(self, chart_type: ChartTypeNameBase | str) -> ChartType[ChartDoc]:
        """
        Finds a chart for a given chart type.

        Args:
            chart_type (ChartTypeNameBase | str): Chart type.

        Raises:
            NotFoundError: If chart is not found
            ChartError: If any other error occurs.

        Returns:
            ChartType[ChartDoc]: Found chart type.

        See Also:
            :ref:`ooodev.utils.kind.chart2_types`
        """
        # pylint: disable=import-outside-toplevel
        from .chart_type import ChartType

        found_type = mChart2.Chart2.find_chart_type(chart_doc=self.component, chart_type=chart_type)
        return ChartType(owner=self, chart_doc=self, component=found_type, lo_inst=self.lo_inst)

    def set_data_point_labels(self, label_type: DataPointLabelTypeKind) -> None:
        """
        Set data point labels for a given chart type.

        Args:
            label_type (DataPointLabelTypeKind): Data point label type.

        Raises:
            ChartError: If any error occurs.

        Returns:
            None:


        Hint:
            - ``DataPointLabelTypeKind`` can be imported from ``ooodev.utils.kind.data_point_label_type_kind``.
        """
        ds_arr = self.get_data_series()
        for ds in ds_arr:
            ds.set_data_point_labels(label_type=label_type)

    @property
    def owner(self) -> TableChart:
        """Gets the owner."""
        return self._owner

    @property
    def first_diagram(self) -> ChartDiagram:
        """Gets the first diagram."""
        if self._first_diagram is None:
            # pylint: disable=import-outside-toplevel
            from .chart_diagram import ChartDiagram

            diagram = self.get_first_diagram()
            self._first_diagram = ChartDiagram(
                owner=self, component=diagram, diagram_kind=ChartDiagramKind.FIRST, lo_inst=self.lo_inst
            )
        return self._first_diagram

    @property
    def axis_x(self) -> ChartAxis:
        """Gets the X Axis Component."""
        if self._axis_x is None:
            # pylint: disable=import-outside-toplevel
            from .chart_axis import ChartAxis

            axis = mChart2.Chart2.get_x_axis(self.component)
            self._axis_x = ChartAxis(owner=self, axis_kind=ChartAxisKind.X, component=axis, lo_inst=self.lo_inst)
        return self._axis_x

    @property
    def axis2_x(self) -> ChartAxis | None:
        """Gets the X Axis Component."""
        if self._axis2_x is None:
            # pylint: disable=import-outside-toplevel
            from .chart_axis import ChartAxis

            try:
                axis = mChart2.Chart2.get_x_axis2(self.component)
            except mEx.ChartError:
                return None
            self._axis2_x = ChartAxis(owner=self, axis_kind=ChartAxisKind.X2, component=axis, lo_inst=self.lo_inst)
        return self._axis2_x

    @property
    def axis_y(self) -> ChartAxis:
        """Gets the Y Axis Component."""
        if self._axis_y is None:
            # pylint: disable=import-outside-toplevel
            from .chart_axis import ChartAxis

            axis = mChart2.Chart2.get_y_axis(self.component)
            self._axis_y = ChartAxis(owner=self, component=axis, axis_kind=ChartAxisKind.Y, lo_inst=self.lo_inst)
        return self._axis_y

    @property
    def axis2_y(self) -> ChartAxis | None:
        """Gets the Y Axis Component."""
        if self._axis2_y is None:
            # pylint: disable=import-outside-toplevel
            from .chart_axis import ChartAxis

            try:
                axis = mChart2.Chart2.get_y_axis2(self.component)
            except mEx.ChartError:
                return None
            self._axis2_y = ChartAxis(owner=self, component=axis, axis_kind=ChartAxisKind.Y2, lo_inst=self.lo_inst)
        return self._axis2_y

    # endregion Methods


if mock_g.FULL_IMPORT:
    from com.sun.star.chart2 import XChartType
    from com.sun.star.chart2 import XFormattedString
    from com.sun.star.chart2 import XRegressionCurve
    from com.sun.star.chart2 import XRegressionCurveContainer
    from com.sun.star.chart2 import XTitle
    from com.sun.star.chart2 import XTitled
    from ooo.dyn.chart.error_bar_style import ErrorBarStyle
    from ooo.dyn.lang.locale import Locale
    from ooodev.utils.kind.chart2_data_role_kind import DataRoleKind
    from .chart_axis import ChartAxis
    from .chart_data_series import ChartDataSeries
    from .chart_diagram import ChartDiagram
    from .chart_error_bar import ChartErrorBar
    from .chart_title import ChartTitle
    from .chart_type import ChartType
    from .data.data_provider import DataProvider
    from .regression_curve.regression_curve import RegressionCurve
