.. _ch31:

*******************************
Chapter 31. XY (Scatter) Charts
*******************************

.. topic:: Overview

    A Scatter Chart (with Regressions); Calculating Regressions; Drawing a Regression Curve; Changing Axis Scales; Adding Error Bars

    Examples: |chart_2_views|_

This chapter continues using the |chart_2_views|_ example from previous chapters, but looks at how various kinds of scatter charts can be generated from spreadsheet data.

A scatter chart is a good way to display (``x``, ``y``) coordinate data since the x-axis values are treated as numbers not categories.
In addition, regression functions can be calculated and displayed, the axis scales can be changed, and error bars added.

The relevant lines of |chart_2_views_py|_ are:

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 38, 40, 42

        # Chart2View.main() ofchart_2_views.py
        def main(self) -> None:
            _ = Lo.load_office(connector=Lo.ConnectPipe(), opt=Lo.Options(verbose=True))

            try:
                doc = Calc.open_doc(fnm=self._data_fnm)
                GUI.set_visible(is_visible=True, odoc=doc)
                sheet = Calc.get_sheet(doc=doc)

                chart_doc = None
                if self._chart_kind == ChartKind.AREA:
                    chart_doc = self._area_chart(doc=doc, sheet=sheet
                elif self._chart_kind == ChartKind.BAR:
                    chart_doc = self._bar_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.BUBBLE_LABELED:
                    chart_doc = self._labeled_bubble_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.COLUMN:
                    chart_doc = self._col_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.COLUMN_LINE:
                    chart_doc = self._col_line_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.COLUMN_MULTI:
                    chart_doc = self._mult_col_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.DONUT:
                    chart_doc = self._donut_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.HAPPY_STOCK:
                    chart_doc = self._happy_stock_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.LINE:
                    chart_doc = self._line_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.LINES:
                    chart_doc = self._lines_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.NET:
                    chart_doc = self._net_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.PIE:
                    chart_doc = self._pie_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.PIE_3D:
                    chart_doc = self._pie_3d_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.SCATTER:
                    chart_doc = self._scatter_chart(doc=doc, sheet=sheet) # sections 1-3
                elif self._chart_kind == ChartKind.SCATTER_LINE_ERROR:
                    chart_doc = self._scatter_line_error_chart(doc=doc, sheet=sheet) # section 5
                elif self._chart_kind == ChartKind.SCATTER_LINE_LOG:
                    chart_doc = self._scatter_line_log_chart(doc=doc, sheet=sheet) # section 4
                elif self._chart_kind == ChartKind.STOCK_PRICES:
                    chart_doc = self._stock_prices_chart(doc=doc, sheet=sheet)

                # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch31_scatter_with_regressions:

31.1 A Scatter Chart (with Regressions)
=======================================

``_scatter_chart()`` in |chart_2_views_py|_ utilizes the "Ice Cream Sales vs. Temperature" table in |ods_doc| (see :numref:`ch31fig_ice_cream_vs_temp_tpl`) to generate the scatter chart in :numref:`ch31fig_tbl_fig1`.

..
    figure 1

.. cssclass:: screen_shot invert

    .. _ch31fig_ice_cream_vs_temp_tpl:
    .. figure:: https://user-images.githubusercontent.com/4193389/206932359-cf84e7ac-1178-464c-a0ea-b6fc9c71a901.png
        :alt: The Ice Cream Sales vs. Temperature Table
        :figclass: align-center

        :The "Ice Cream Sales vs. Temperature" Table.

..
    figure 2

.. cssclass:: screen_shot

    .. _ch31fig_tbl_fig1:
    .. figure:: https://user-images.githubusercontent.com/4193389/206932404-6b5fa353-faa4-42ca-b04a-5ca359655b7b.png
        :alt: Scatter Chart for the Table in previous figure.
        :figclass: align-center

        :Scatter Chart for the Table in :numref:`ch31fig_ice_cream_vs_temp_tpl`.

Note that the x-axis in :numref:`ch31fig_ice_cream_vs_temp_tpl` is numerical, showing values ranging between ``10.0`` and ``26.0``.
This range is calculated automatically by the template.

.. tabs::

    .. code-tab:: python

        # 
        def _scatter_chart(
            self, doc: XSpreadsheetDocument, sheet: XSpreadsheet
        ) -> XChartDocument:
            # uses the "Ice Cream Sales vs Temperature" table
            range_addr = Calc.get_address(sheet=sheet, range_name="A110:B122")
            chart_doc = Chart2.insert_chart(
                sheet=sheet,
                cells_range=range_addr,
                cell_name="C109",
                width=16,
                height=11,
                diagram_name=ChartTypes.XY.TEMPLATE_LINE.SCATTER_SYMBOL,
            )
            Calc.goto_cell(cell_name="A104", doc=doc)

            Chart2.set_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A109")
            )
            Chart2.set_x_axis_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A110")
            )
            Chart2.set_y_axis_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B110")
            )
            Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))

            # Chart2.calc_regressions(chart_doc)
            # Chart2.draw_regression_curve(chart_doc=chart_doc, curve_kind=CurveKind.LINEAR)
            return XChartDocument

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

If the :py:meth:`.Chart2.calc_regressions` line is uncommented then several different regression functions are calculated using the chart's data.
Their equations and |R2| values are printed as shown below:

::

    Linear regression curve:
      Curve equation: f(x) = 30.09x - 159.5
      R^2 value: 0.917
    
    Logarithmic regression curve:
      Curve equation: f(x) = 544.1 ln(x) - 1178
      R^2 value: 0.921
    
    Exponential regression curve:
      Curve equation: f(x) = 81.62 exp( 0.0826 x )
      R^2 value: 0.865
    
    Power regression curve:
      Curve equation: f(x) = 4.545 x^1.525
      R^2 value: 0.906
    
    Polynomial regression curve:
      Curve equation: f(x) =  - 0.5384x^2 + 50.24x - 340.1
      R^2 value: 0.921
    
    Moving average regression curve:
      Curve equation: Moving average trend line with period = %PERIOD
      R^2 value: NaN

A logarithmic or quadratic polynomial are the best matches, but linear is a close third.
The "moving average" |R2| result is ``NaN`` (Not-a-Number) since no average of period 2 matches the data.

If the :py:meth:`.Chart2.draw_regression_curve` call is uncommented, the chart drawing will include a linear regression line and its equation and |R2| value (see :numref:`ch31fig_scatter_chart_linear_regression_fig1_tbl`).

..
    figure 3

.. cssclass:: screen_shot

    .. _ch31fig_scatter_chart_linear_regression_fig1_tbl:
    .. figure:: https://user-images.githubusercontent.com/4193389/207115403-9d19443d-d0c1-44b5-9c07-e735868185ce.png
        :alt: Scatter Chart with Linear Regression Line for the Table in Figure 1 of this chapter
        :figclass: align-center

        :Scatter Chart with Linear Regression Line for the Table in :numref:`ch31fig_ice_cream_vs_temp_tpl`.

The regression function is ``f(x) = 30.09x - 159.47``, and the ``C`` value is ``0.92`` (to ``2`` dp).
If the constant (``curve_kind``) is changed to :py:attr:`.CurveKind.LOGARITHMIC` in the call to :py:meth:`.Chart2.draw_regression_curve` then the generated function is ``f(x) = 544.1 ln(x) – 1178`` with an |R2| value of ``0.92``.
Other regression curves are represented by constants :py:attr:`.CurveKind.EXPONENTIAL`, :py:attr:`.CurveKind.POWER`, :py:attr:`.CurveKind.POLYNOMIAL`, and :py:attr:`.CurveKind.MOVING_AVERAGE`.

.. _ch31_calculating_regressions:

31.2 Calculating Regressions
============================

:py:meth:`.Chart2.calc_regressions` is:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def calc_regressions(cls, chart_doc: XChartDocument) -> None:

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

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Chart2.create_curve` matches the regression constants defined in :py:class:`~.kind.curve_kind.CurveKind` to regression services offered by the API:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def create_curve(curve_kind: CurveKind) -> XRegressionCurve:
            try:
                rc = Lo.create_instance_mcf(XRegressionCurve, curve_kind.to_namespace(), raise_err=True)
                return rc
            except Exception as e:
                raise ChartError("Error creating curve") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

There are seven regression curve services in the chart2 module, all of which support the XRegressionCurve_ interface, as shown in :numref:`ch31fig_regression_curve_srv`.

..
    figure 4

.. cssclass:: diagram invert

    .. _ch31fig_regression_curve_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/207119312-8cc964f2-9869-40fb-a678-73fa4accdcb0.png
        :alt: The Regression Curve Services
        :figclass: align-center

        :The RegressionCurve_ Services

The RegressionCurve_ service shown in :numref:`ch31fig_regression_curve_srv` is not a superclass for the other services.
Also note that the regression curve service for power functions is called ``PotentialRegressionCurve``.

:py:meth:`.Chart2.eval_curve` uses ``XRegressionCurve.getCalculator()`` to access the XRegressionCurveCalculator_ interface.
It sets up the data and parameters for a particular curve, and prints the results of curve fitting:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def eval_curve(cls, chart_doc: XChartDocument, curve: XRegressionCurve) -> None:
            curve_calc = curve.getCalculator()
            degree = 1
            ct = cls.get_curve_type(curve)
            if ct != CurveKind.LINEAR:
                degree = 2  # assumes POLYNOMIAL trend has degree == 2

            curve_calc.setRegressionProperties(degree, False, 0.0, 2, 0)

            data_source = cls.get_data_source(chart_doc)
            # cls.print_labled_seqs(data_source)

            xvals = cls.get_chart_data(data_source=data_source, idx=0)
            yvals = cls.get_chart_data(data_source=data_source, idx=0)
            curve_calc.recalculateRegression(xvals, yvals)

            print(f"  Curve equations: {curve_calc.getRepresentation()}")
            cc = curve_calc.getCorrelationCoefficient()
            print(f"  R^2 value: {(cc*cc):.3f}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The calculation is configured by calling ``XRegressionCurveCalculator.setRegressionProperties()``, and carried out by ``XRegressionCurveCalculator.recalculateRegression()``.

The degree argument of ``setRegressionProperties()`` specifies the polynomial curve's degree, which is hard coded to be quadratic (:abbreviation:`i.e.` a degree of ``2``).
The period argument is used when a moving average curve is being fitted.

``recalculateRegression()`` requires two arrays of ``x`` and ``y`` axis values for the scatter points.
These are obtained from the chart's data source by calling :py:meth:`.Chart2.get_data_source` which returns the XDataSource_ interface for the DataSeries_ service.

:numref:`ch31fig_data_series_detail` shows the XDataSource_, XRegressionCurveContainer_, and XDataSink_ interfaces of the DataSeries_ service.

..
    figure 5

.. cssclass:: diagram invert

    .. _ch31fig_data_series_detail:
    .. figure:: https://user-images.githubusercontent.com/4193389/207121574-944ec294-e17f-4ab3-8d1d-1c6a69f96443.png
        :alt: More Detailed DataSeries Service.
        :figclass: align-center

        :More Detailed DataSeries_ Service.

In previous chapters, only used the XDataSeries_ interface, which offers access to the data points in the chart.
The XDataSource_ interface, which is read-only, gives access to the underlying data that was used to create the points.
The data is stored as an array of XLabeledDataSequence_ objects; each object contains a label and a sequence of data.

:py:meth:`.Chart2.get_data_source` is defined as:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def get_data_source(
            cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase | str = ""
        ) -> XDataSource:
            try:
                dsa = cls.get_data_series(chart_doc=chart_doc, chart_type=chart_type)
                ds = Lo.qi(XDataSource, dsa[0], True)
                return ds
            except NotFoundError:
                raise
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error getting data source for chart") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This method assumes that the programmer wants the first data source in the data series.
This is adequate for most charts which only use one data source.

:py:meth:`.Chart2.print_labeled_seqs` is a diagnostic function for printing all the labeled data sequences stored in an XDataSource_:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def print_labeled_seqs(data_source: XDataSource) -> None:
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

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

When these function is applied to the data source for the scatter chart, the following is printed:

::

    No. of sequences in data source: 2
    Temperature °C :  14.2  16.4  11.9  15.2  18.5  22.1  19.4
                      25.1  23.4  18.1  22.6  17.2
    Source range: $examples.$A$111:$A$122
    
    Ice Cream Sales :  215.0  325.0  185.0  332.0  406.0  522.0
                       412.0  614.0  544.0  421.0  445.0  408.0
    Source range: $examples.$B$111:$B$122

This output shows that the data source consists of two XLabeledDataSequence_ objects, representing the ``x`` and ``y`` values in the data source (see :numref:`ch31fig_ice_cream_vs_temp_tpl`).
These objects' data are extracted as arrays by calls to :py:meth:`.Chart2.get_chart_data`:

.. tabs::

    .. code-tab:: python

        # in Chart2 class part of eval_curve()
        # ...
        data_source = cls.get_data_source(chart_doc)
        cls.print_labled_seqs(data_source)

        xvals = cls.get_chart_data(data_source=data_source, idx=0)
        yvals = cls.get_chart_data(data_source=data_source, idx=0)
        curve_calc.recalculateRegression(xvals, yvals)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

When ``recalculateRegression()`` has finished, various results about the fitted curve can be extracted from the XRegressionCurveCalculator_ variable, ``curve_calc``.
:py:meth:`~.Chart2.eval_curve` prints the function string (using ``getRepresentation()``) and the |R2| value (using ``getCorrelationCoefficient()``).

.. _ch31_drawing_regression_curve:

31.3 Drawing a Regression Curve
===============================

One of the surprising things about drawing a regression curve is that there's no need to explicitly calculate the curve's function with XRegressionCurveCalculator_.
Instead :py:meth:`.Chart2.draw_regression_curve` only has to initialize the curve via the data series' XRegressionCurveContainer_ interface (see :numref:`ch31fig_data_series_detail`).

:py:meth:`~.Chart2.draw_regression_curve` is:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def draw_regression_curve(
            cls, chart_doc: XChartDocument, curve_kind: CurveKind
        ) -> None:
            try:
                data_series_arr = cls.get_data_series(chart_doc=chart_doc)
                rc_con = Lo.qi(XRegressionCurveContainer, data_series_arr[0], True)
                curve = cls.create_curve(curve_kind)
                rc_con.addRegressionCurve(curve)

                ps = curve.getEquationProperties()
                Props.set_property(ps, "ShowCorrelationCoefficient", True)
                Props.set_property(ps, "ShowEquation", True)

                key = cls.get_number_format_key(chart_doc=chart_doc, nf_str="0.00")  # 2 dp
                if key != -1:
                    Props.set_property(ps, "NumberFormat", key)
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error drawing regression curve") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The XDataSeries_ interface for the first data series in the chart is converted to XRegressionCurveContainer_, and an XRegressionCurve_ instance added to it.
This triggers the calculation of the curve's function.
The rest of :py:meth:`~.Chart2.draw_regression_curve` deals with how the function information is displayed on the chart.

``XRegressionCurve.getEquationProperties()`` returns a property set which is an instance of the RegressionCurveEquation_ service class, shown in :numref:`ch31fig_regression_curve_equation_cls`.

..
    figure 6

.. cssclass:: diagram invert

    .. _ch31fig_regression_curve_equation_cls:
    .. figure:: https://user-images.githubusercontent.com/4193389/207125967-7c32a3ce-e46c-4df2-b0fd-99c9cbf0288b.png
        :alt: The Regression Curve Equation Property Class.
        :figclass: align-center

        :The RegressionCurveEquation_ Property Class.

RegressionCurveEquation_ inherits properties related to character, fill, and line, since it controls how the curve, function string, and |R2| value are drawn on the chart.
These last two are made visible by setting the ``ShowEquation`` and ``ShowCorrelationCoefficient`` properties to ``True``, which are defined in RegressionCurveEquation_.

Another useful property is ``NumberFormat`` which can be used to reduce the number of decimal places used when printing the function and |R2| value.

:py:meth:`.Chart2.get_number_format_key` converts a number format string into a number format key, which is assigned to the ``NumberFormat`` property:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def get_number_format_key(chart_doc: XChartDocument, nf_str: str) -> int:
            try:
                xfs = Lo.qi(XNumberFormatsSupplier, chart_doc, True)
                n_formats = xfs.getNumberFormats()
                key = int(n_formats.queryKey(nf_str, Locale("en", "us", ""), False))
                if key == -1:
                    Lo.print(f'Could not access key for number format: "{nf_str}"')
                return key
            except Exception as e:
                raise ChartError("Error getting number format key") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The string-to-key conversion is straight forward if you know what number format string to use, but there's little documentation on them.
Probably the best approach is to use the Format ``->`` Cells menu item in a spreadsheet document, and examine the dialog in :numref:`ch31fig_format_cells_dialog`.

..
    figure 7

.. cssclass:: screen_shot invert

    .. _ch31fig_format_cells_dialog:
    .. figure:: https://user-images.githubusercontent.com/4193389/207127667-29c4b51c-1c0a-4376-9345-2564752014dc.png
        :alt: The Format Cells Dialog
        :figclass: align-center

        :The Format Cells Dialog.

When you select a given category and format, the number format string is shown in the "Format Code" field at the bottom of the dialog.
:numref:`ch31fig_format_cells_dialog` shows that the format string for two decimal place numbers is ``0.00``.
This string should be passed to :py:meth:`~.Chart2.get_number_format_key` in :py:meth:`~.Chart2.draw_regression_curve`:

.. tabs::

    .. code-tab:: python

        key = cls.get_number_format_key(chart_doc=chart_doc, nf_str="0.00")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch31_changing_axis_scales:

31.4 Changing Axis Scales
=========================

Another way to understand scatter data is by changing the chart's axis scaling.
Alternatives to linear are logarithmic, exponential, or power, although is seems that the latter two cause the chart to be drawn incorrectly.

``_scatter_line_log_chart()`` in |chart_2_views_py|_ utilizes the "Power Function Test" table in |ods_doc| (see :numref:`ch31fig_pwr_fn_tst_tbl`).

..
    figure 8

.. cssclass:: screen_shot invert

    .. _ch31fig_pwr_fn_tst_tbl:
    .. figure:: https://user-images.githubusercontent.com/4193389/207130267-dde50520-364d-4e8e-b067-f9d2da2f99a2.png
        :alt: The Power Function Test Table.
        :figclass: align-center

        :The "Power Function Test" Table.

The formula ``=4.1*POWER(E<number>, 3.2)`` is used (:abbreviation:`i.e.` 4.1x\ :sup:`3.2`) to generate the ``Actual`` column from the ``Input`` column's cells.
Then I manually rounded the results and copied them into the "Output" column.

The data range passed to the :py:meth:`.Chart.insert_chart` uses the ``Input`` and ``Output`` columns of the table in :numref:`ch31fig_pwr_fn_tst_tbl`.
The generated scatter chart in :numref:`ch31fig_chart_for_fig8` uses log scaling for the axes, and fits a power function to the data points.

..
    figure 9

.. cssclass:: screen_shot

    .. _ch31fig_chart_for_fig8:
    .. figure:: https://user-images.githubusercontent.com/4193389/207132261-982d1775-4e8b-48b4-9da0-547ab0e3c636.png
        :alt: Scatter Chart for the Table in previous figure.
        :figclass: align-center

        :Scatter Chart for the Table in :numref:`ch31fig_pwr_fn_tst_tbl`.

The power function fits the data so well that the black regression line lies over the blue data curve.
The regression function is ``f(x) = 3.89 x^2.32`` (:abbreviation:`i.e.` 3.89x\ :sup:`2.32` ) with |R2| = ``1.00``, which is close to the power formula used to generate the ``Actual`` column data.

``_scatter_line_log_chart()`` is:

.. tabs::

    .. code-tab:: python

        # Chart2View._scatter_line_log_chart() in chart_2_views.py
        def _scatter_line_log_chart(
            self, doc: XSpreadsheetDocument, sheet: XSpreadsheet
        ) -> XChartDocument:
            # uses the "Power Function Test" table
            range_addr = Calc.get_address(sheet=sheet, range_name="E110:F120")
            chart_doc = Chart2.insert_chart(
                sheet=sheet,
                cells_range=range_addr,
                cell_name="A121",
                width=20,
                height=11,
                diagram_name=ChartTypes.XY.TEMPLATE_LINE.SCATTER_LINE_SYMBOL,
            )
            Calc.goto_cell(cell_name="A121", doc=doc)

            Chart2.set_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E109")
            )
            Chart2.set_x_axis_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E110")
            )
            Chart2.set_y_axis_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="F110")
            )
            Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))

            # change x- and y- axes to log scaling
            x_axis = Chart2.scale_x_axis(chart_doc=chart_doc, scale_type=CurveKind.LOGARITHMIC)
            _ = Chart2.scale_y_axis(chart_doc=chart_doc, scale_type=CurveKind.LOGARITHMIC)
            Chart2.draw_regression_curve(chart_doc=chart_doc, curve_kind=CurveKind.POWER)
            return chart_doc

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Chart2.scale_x_axis` and :py:meth:`~.Chart2.scale_y_axis` call the more general :py:meth:`~.Chart2.scale_axis` method:

.. tabs::

    .. code-tab:: python

        # in Chart2 Class
        @classmethod
        def scale_x_axis(cls, chart_doc: XChartDocument, scale_type: CurveKind) -> XAxis:
            return cls.scale_axis(chart_doc=chart_doc, axis_val=AxisKind.X, idx=0, scale_type=scale_type)

        @classmethod
        def scale_y_axis(cls, chart_doc: XChartDocument, scale_type: CurveKind) -> XAxis:
            return cls.scale_axis(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=0, scale_type=scale_type)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Chart2.scale_axis` utilizes ``XAxis.getScaleData()`` and ``XAxis.setScaleData()`` to access and modify the axis scales:

.. tabs::

    .. code-tab:: python

        # in Chart2 Class
        @classmethod
        def scale_axis(
            cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int, scale_type: CurveKind
        ) -> XAxis:
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
                    Lo.print(f'Did not reconize scaling type: "{scale_type}"')
                else:
                    sd.Scaling = Lo.create_instance_mcf(XScaling, f"com.sun.star.chart2.{s}", raise_err=True)
                axis.setScaleData(sd)
                return axis
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error setting axis scale") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The different scaling services all support the XScaling_ interface, as illustrated by :numref:`ch31fig_scaling_srv`.

..
    figure 10

.. cssclass:: screen_shot invert

    .. _ch31fig_scaling_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/207135718-cd68b66d-7489-4ad9-8682-f78098cdd50c.png
        :alt: The Scaling Services.
        :figclass: align-center

        :The Scaling Services.

.. _ch31_adding_error_bars:

31.5 Adding Error Bars
======================

``_scatter_line_error_chart()`` in |chart_2_views_py|_ employs the "Impact Data : 1018 Cold Rolled" table in |ods_doc| (see :numref:`ch31fig_impact_tbl`).

..
    figure 11

.. cssclass:: screen_shot invert

    .. _ch31fig_impact_tbl:
    .. figure:: https://user-images.githubusercontent.com/4193389/207136984-246b7a19-78da-40dc-adb3-942dbb1d24c6.png
        :alt: The Impact Data : 1018 Cold Rolled Table.
        :figclass: align-center

        :The "Impact Data : 1018 Cold Rolled" Table.

The data range passed to the :py:meth:`.Chart.insert_chart` uses the ``Temperature`` and ``Mean`` columns of the table; the ``Stderr`` column is added separately to generate error bars along the ``y-axis``.
The resulting scatter chart is shown in :numref:`ch31fig_scatter_chart_fig11`.

..
    figure 12

.. cssclass:: screen_shot

    .. _ch31fig_scatter_chart_fig11:
    .. figure:: https://user-images.githubusercontent.com/4193389/207137386-71fdd02b-2b95-43d3-b5d8-670703481447.png
        :alt: Scatter Chart with Error Bars for the Table in previous figure.
        :figclass: align-center

        :Scatter Chart with Error Bars for the Table in :numref:`ch31fig_impact_tbl`.

``_scatter_line_error_chart()`` is:

.. tabs::

    .. code-tab:: python

        # Chart2View._scatter_line_error_chart() in chart_2_views.py
        def _scatter_line_error_chart(
            self, doc: XSpreadsheetDocument, sheet: XSpreadsheet
        ) -> XChartDocument:
            range_addr = Calc.get_address(sheet=sheet, range_name="A142:B146")
            chart_doc = Chart2.insert_chart(
                sheet=sheet,
                cells_range=range_addr,
                cell_name="F115",
                width=14,
                height=16,
                diagram_name=ChartTypes.XY.TEMPLATE_LINE.SCATTER_LINE_SYMBOL,
            )
            Calc.goto_cell(cell_name="A123", doc=doc)

            Chart2.set_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A141")
            )
            Chart2.set_x_axis_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A142")
                )
            Chart2.set_y_axis_title(
                chart_doc=chart_doc, title="Impact Energy (Joules)"
            )
            Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))

            Lo.print("Adding y-axis error bars")
            sheet_name = Calc.get_sheet_name(sheet)
            error_label = f"{sheet_name}.C142"
            error_range = f"{sheet_name}.C143:C146"
            Chart2.set_y_error_bars(
                chart_doc=chart_doc, data_label=error_label, data_range=error_range
            )
            return chart_doc

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The new feature in ``_scatter_line_error_chart()`` is the call to :py:meth:`.Chart2.set_y_error_bars`, which is explained over the next four subsections.

.. _ch31_creating_new_chart_data:

31.5.1 Creating New Chart Data
------------------------------

The secret to adding extra data to a chart is ``XDataSink.setData()``.
XDataSink_ is yet another interface for the DataSeries service (see :numref:`ch31fig_data_series_detail`).

There are several stages required, which are depicted in :numref:`ch31fig_xdata_sink_add_to_chart`.

..
    figure 13

.. cssclass:: diagram invert

    .. _ch31fig_xdata_sink_add_to_chart:
    .. figure:: https://user-images.githubusercontent.com/4193389/207138850-c5578c57-ee64-4911-b4d1-601f99e4226f.png
        :alt: Using XDataSink to Add Data to a Chart.
        :figclass: align-center

        :Using XDataSink_ to Add Data to a Chart.

The DataProvider_ service produces two XDataSequence_ objects which are combined to become a XLabeledDataSequence_ object.
An array of these objects is passed to ``XDataSink.setData()``.

The DataProvider_ service is accessed with one line of code:

.. tabs::

    .. code-tab:: python

        dp = chart_doc.getDataProvider() # XDataProvider

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Chart2.create_ld_seq` creates a XLabeledDataSequence_ instance from two XDataSequence_ objects, one acting as a label the other as data.
The XDataSequence_ object representing the data must have its ``Role`` property set to indicate the type of the data.

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def create_ld_seq(
            dp: XDataProvider, role: DataRoleKind | str, data_label: str, data_range: str
        ) -> XLabeledDataSequence:
            try:
                # create data sequence for the label
                lbl_seq = dp.createDataSequenceByRangeRepresentation(data_label)

                # reate data sequence for the data and role
                data_seq = dp.createDataSequenceByRangeRepresentation(data_range)

                ds_ps = Lo.qi(XPropertySet, data_seq, True)

                # specify data role (type)
                Props.set_property(ds_ps, "Role", str(role))
                # Props.show_props("Data Sequence", ds_ps)

                # create new labeled data sequence using sequences
                ld_seq = Lo.create_instance_mcf(
                    XLabeledDataSequence,
                    "com.sun.star.chart2.data.LabeledDataSequence",
                    raise_err=True
                )
                ld_seq.setLabel(lbl_seq)
                ld_seq.setValues(data_seq)
                return ld_seq
            except Exception as e:
                raise ChartError("Error creating LD sequence") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Four arguments are passed to :py:meth:`~.Chart2.create_ld_seq`: a reference to the XDataProvider_ interface, a role string, a label, and a data range. For example:

.. tabs::

    .. code-tab:: python

        sheet_name = Calc.get_sheet_name(sheet)
        data_label = f"{sheet_name}.C142"
        data_range = f"{sheet_name}.C143:C146"
        lds = Chart2.create_ld_seq(
            dp=dp, role=DataRoleKind.ERROR_BARS_Y_POSITIVE, data_label=data_label, data_range=data_range
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``role`` constants are defined in :py:class:`~.kind.chart2_data_role_kind.DataRoleKind`.

``XDataSink.setData()`` can accept multiple XLabeledDataSequence_ objects in an array, making it possible to add several kinds of data to the chart at once.
This is just as well since it is easier to add two XLabeledDataSequence_ objects, one for the error bars above the data points (:abbreviation:`i.e.` up the ``y-axis``),
and another for the error bars below the points (:abbreviation:`i.e.` down the ``y-axis``).
The code for doing this:

.. tabs::

    .. code-tab:: python

        # in Chart2.set_y_error_bars(); see section 5.4 below
        # convert into data sink
        data_sink = Lo.qi(XDataSink, error_bars_ps, True)

        dp = chart_doc.getDataProvider()
        pos_err_seq = cls.create_ld_seq(
            dp=dp,
            role=DataRoleKind.ERROR_BARS_Y_POSITIVE,
            data_label=data_label,
            data_range=data_range
        )
        neg_err_seq = cls.create_ld_seq(
            dp=dp,
            role=DataRoleKind.ERROR_BARS_Y_NEGATIVE,
            data_label=data_label,
            data_range=data_range
        )

        ld_seq = (pos_err_seq, neg_err_seq)
        # store the error bar data sequences in the data sink
        data_sink.setData(ld_seq)

        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This code fragment leaves two topics unexplained: how the data sink is initially created, and how the data sink is linked to the chart.

.. _ch31_creating_data_sink:

31.5.2 Creating the Data Sink
-----------------------------

The data sink for error bars relies on the ErrorBar_ service, which is shown in :numref:`ch31fig_error_bar_srv`.

..
    figure 14

.. cssclass:: diagram invert

    .. _ch31fig_error_bar_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/207145870-185b0893-dfa2-4254-8720-fdf4160b1525.png
        :alt: The ErrorBar Service
        :figclass: align-center

        :The ErrorBar_ Service

The ErrorBar_ service stores error bar properties and implements the XDataSink_ interface.
The following code fragment creates an instance of the ErrorBar_ service, sets some of its properties, and converts it to an XDataSink_:

.. tabs::

    .. code-tab:: python

        # in Chart2.set_y_error_bars(); see section 5.4 below
        error_bars_ps = Lo.create_instance_mcf(
            XPropertySet, "com.sun.star.chart2.ErrorBar", raise_err=True
        )
        Props.set(
            error_bars_ps,
            ShowPositiveError=True,
            ShowNegativeError=True,
            ErrorBarStyle=ErrorBarStyle.FROM_DATA
        )

        # convert into data sink
        data_sink = Lo.qi(XDataSink, error_bars_ps, True)

        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch31_linking_data_sink_chart:

31.5.3 Linking the Data Sink to the Chart
-----------------------------------------

Once the data sink has been filled with XLabeledDataSequence_ objects, it can be linked to the data series in the chart.
For error bars this is done via the properties ``ErrorBarX`` and ``ErrorBarY``.
For example, the following code assigns a data sink to the data series' ``ErrorBarY`` property:

.. tabs::

    .. code-tab:: python

        # in Chart2.set_y_error_bars(); see section 5.4 below
        # ...
        # store error bar in data series
        data_series_arr = cls.get_data_series(chart_doc=chart_doc)
        data_series = data_series_arr[0]
        Props.set(data_series, ErrorBarY=error_bars_ps)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Note that the value assigned to ``ErrorBarY`` is not an XDataSink_ interface (:abbreviation:`i.e.` not ``data_sink`` from the earlier code fragment) but its property set (:abbreviation:`i.e.` ``props``).

.. _ch31_bring_together:

31.5.4 Bringing it All Together
-------------------------------

:py:meth:`.Chart2.set_y_error_bars` combines the previous code fragments into a single method: the data sink is created (as a property set),
XLabeledDataSequence_ data is added to it, and then the sink is linked to the chart's data series:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def set_y_error_bars(
            cls, chart_doc: XChartDocument, data_label: str, data_range: str
        ) -> None:
            try:
                error_bars_ps = Lo.create_instance_mcf(
                    XPropertySet, "com.sun.star.chart2.ErrorBar", raise_err=True
                )
                Props.set(
                    error_bars_ps,
                    ShowPositiveError=True,
                    ShowNegativeError=True,
                    ErrorBarStyle=ErrorBarStyle.FROM_DATA
                )

                # convert into data sink
                data_sink = Lo.qi(XDataSink, error_bars_ps, True)

                # use data provider to create labelled data sequences
                # for the +/- error ranges
                dp = chart_doc.getDataProvider()

                pos_err_seq = cls.create_ld_seq(
                    dp=dp,
                    role=DataRoleKind.ERROR_BARS_Y_POSITIVE,
                    data_label=data_label,
                    data_range=data_range
                )
                neg_err_seq = cls.create_ld_seq(
                    dp=dp,
                    role=DataRoleKind.ERROR_BARS_Y_NEGATIVE,
                    data_label=data_label,
                    data_range=data_range
                )

                ld_seq = (pos_err_seq, neg_err_seq)

                # store the error bar data sequences in the data sink
                data_sink.setData(ld_seq)
                # Props.show_obj_props("Error Bar", error_bars_ps)
                # "ErrorBarRangePositive" and "ErrorBarRangeNegative"
                # will now have ranges they are read-only

                # store error bar in data series
                data_series_arr = cls.get_data_series(chart_doc=chart_doc)
                # print(f'No. of data serice: {len(data_series_arr)}')
                data_series = data_series_arr[0]
                # Props.show_obj_props("Data Series 0", data_series)
                Props.set(data_series, ErrorBarY=error_bars_ps)
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error Setting y error bars") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This is not our last visit to DataSink_ and XDataSink_. Their features show up again in the next chapter.

.. |R2| replace:: R\ :sup:`2`

.. |ods_doc| replace:: ``chartsData.ods``

.. |chart_2_views| replace:: Chart2 Views
.. _chart_2_views: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/chart2/Chart_2_Views

.. |chart_2_views_py| replace:: chart_2_views.py
.. _chart_2_views_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/chart2/Chart_2_Views/chart_2_views.py

.. _DataProvider: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1data_1_1DataProvider.html
.. _DataSeries: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1DataSeries.html
.. _DataSink: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1data_1_1DataSink.html
.. _ErrorBar: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1ErrorBar.html
.. _RegressionCurve: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1RegressionCurve.html
.. _RegressionCurveEquation: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1RegressionCurveEquation.html
.. _XDataProvider: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1data_1_1XDataProvider.html
.. _XDataSequence: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1data_1_1XDataSequence.html
.. _XDataSeries: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XDataSeries.html
.. _XDataSink: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1data_1_1XDataSink.html
.. _XDataSource: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1data_1_1XDataSource.html
.. _XLabeledDataSequence: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1data_1_1XLabeledDataSequence.html
.. _XRegressionCurve: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XRegressionCurve.html
.. _XRegressionCurveCalculator: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XRegressionCurveCalculator.html
.. _XRegressionCurveContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XRegressionCurveContainer.html
.. _XScaling: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XScaling.html
