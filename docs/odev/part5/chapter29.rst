.. _ch29:

*************************
Chapter 29. Column Charts
*************************

.. topic:: Overview

    Creating a Chart Title; Creating Axis Titles; Rotating Axis Titles; What Chart Templates are Available?; Multiple Columns; 3D Pizazz; The Column and Line Chart

    Examples: |chart_2_views|_

All the chart examples in the next four chapters come from the same program, |chart_2_views|_, which loads a spreadsheet from ``chartsData.ods``.
Depending on the function, a different table in the sheet is used to create a chart from a template.
The main() function of |chart_2_views_py|_ is:

.. tabs::

    .. code-tab:: python

        # Chart2View.main() ofchart_2_views.py
        def main(self) -> None:
            _ = Lo.load_office(connector=Lo.ConnectPipe(), opt=Lo.Options(verbose=True))

            try:
                doc = Calc.open_doc(fnm=self._data_fnm)
                GUI.set_visible(is_visible=True, odoc=doc)
                sheet = Calc.get_sheet(doc=doc)

                chart_doc = None
                if self._chart_kind == ChartKind.AREA:
                    chart_doc = self._area_chart(doc=doc, sheet=sheet)
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
                    chart_doc = self._scatter_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.SCATTER_LINE_ERROR:
                    chart_doc = self._scatter_line_error_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.SCATTER_LINE_LOG:
                    chart_doc = self._scatter_line_log_chart(doc=doc, sheet=sheet)
                elif self._chart_kind == ChartKind.STOCK_PRICES:
                    chart_doc = self._stock_prices_chart(doc=doc, sheet=sheet)

                if chart_doc:
                    Chart2.print_chart_types(chart_doc)

                    template_names = Chart2.get_chart_templates(chart_doc)
                    Lo.print_names(template_names, 1)

                Lo.delay(2000)
                msg_result = MsgBox.msgbox(
                    "Do you wish to close document?",
                    "All done",
                    boxtype=MessageBoxType.QUERYBOX,
                    buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
                )
                if msg_result == MessageBoxResultsEnum.YES:
                    Lo.close_doc(doc=doc, deliver_ownership=True)
                    Lo.close_office()
                else:
                    print("Keeping document open")
            except Exception:
                Lo.close_office()
                raise

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Lets assume that ``self._chart_kind == ChartKind.COLUMN`` for now.

``_col_chart()`` utilizes the "Sneakers Sold this Month" table in ``chartsData.ods`` (see :numref:`ch29fig_sneakers_sold_month_tbl`) to generate the column chart in :numref:`ch29fig_chart_for_sneaker_sold_month_tbl`.

..
    figure 1

.. cssclass:: screen_shot invert

    .. _ch29fig_sneakers_sold_month_tbl:
    .. figure:: https://user-images.githubusercontent.com/4193389/206542426-9721a34d-851e-42e7-b6cd-83f0582f8f71.png
        :alt: Sneakers Sold this Month Table
        :figclass: align-center

        :The "Sneakers Sold this Month" Table.

..
    figure 2

.. cssclass:: screen_shot

    .. _ch29fig_chart_for_sneaker_sold_month_tbl:
    .. figure:: https://user-images.githubusercontent.com/4193389/206542602-82abadea-7317-4edd-b100-db7870ca1bc0.png
        :alt: The Column Chart for previous Table
        :figclass: align-center

        :The Column Chart for the Table in :numref:`ch29fig_sneakers_sold_month_tbl`.

``_col_chart()`` is:

.. tabs::

    .. code-tab:: python

        # Chart2View._col_chart() of chart_2_views.py
        def _col_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> XChartDocument:
            # draw a column chart;
            # uses "Sneakers Sold this Month" table
            range_addr = Calc.get_address(sheet=sheet, range_name="A2:B8")
            chart_doc = Chart2.insert_chart(
                sheet=sheet,
                cells_range=range_addr,
                cell_name="C3",
                width=15,
                height=11,
                diagram_name=ChartTypes.Column.TEMPLATE_STACKED.COLUMN,
            )
            Calc.goto_cell(cell_name="A1", doc=doc)

            Chart2.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A1"))
            Chart2.set_x_axis_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A2")
            )
            Chart2.set_y_axis_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B2")
            )
            Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
            return chart_doc

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The column chart created by :py:meth:`.Chart2.insert_chart` utilizes the cell range ``A2:B8``, which spans the two columns of the table, but not the title in cell ``A1``.
The ``C3`` argument specifies where the top-left corner of the chart will be positioned in the sheet, and ``15x11`` are the dimensions of the image in millimeters.

:py:meth:`.Calc.goto_cell` causes the application window's view of the spreadsheet to move so that cell ``A1`` is visible, which lets the user see the sneakers table and the chart together.

If the three set methods and ``rotateYAxisTitle()`` are left out of ``_col_chart()``, then the generated chart will have no titles as in :numref:`ch29fig_col_chart_for_tbl_sneaker_sold`.

..
    figure 3

.. cssclass:: screen_shot

    .. _ch29fig_col_chart_for_tbl_sneaker_sold:
    .. figure:: https://user-images.githubusercontent.com/4193389/206544345-5717d5c2-268f-49a6-a775-baaf1c375a92.png
        :alt: The Column Chart for the Table in The Sneakers Sold this Month Table, with no Titles.
        :figclass: align-center

        :The Column Chart for the Table in :numref:`ch29fig_chart_for_sneaker_sold_month_tbl`, with no Titles.

.. _ch29_creading_chart_title:

29.1 Creating a Chart Title
===========================

:py:meth:`.Chart2.set_title` is passed a string which becomes the chart's title. For example:

.. tabs::

    .. code-tab:: python

        # part of _col_chart() in Chart2View class
        Chart2.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A1"))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

utilizes the string from cell ``A1`` of the spreadsheet (see :numref:`ch29fig_sneakers_sold_month_tbl`).

Setting a title requires three interfaces: XTitled_, XTitle_, and XFormattedString_.
XTitled_ is utilized by several chart services, as shown in :numref:`ch29fig_srv_using_xtitled`.

..
    figure 4

.. cssclass:: diagram invert

    .. _ch29fig_srv_using_xtitled:
    .. figure:: https://user-images.githubusercontent.com/4193389/206546297-c4ad8a86-8840-434e-849a-1fc7a34c3976.png
        :alt: Services Using the XTitled Interface
        :figclass: align-center

        :Services Using the XTitled_ Interface.

The XChartDocument_ interface is converted into XTitled_ by :py:meth:`.Chart2.set_title`, so an XTitle_ object can be assigned to the chart:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def get_title(chart_doc: XChartDocument) -> XTitle:
            try:
                xtilted = Lo.qi(XTitled, chart_doc, True)
                return xtilted.getTitleObject()
            except Exception as e:
                raise ChartError("Error getting title from chart") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The XTitle_ object is an instance of the Title_ service which inherits a wide assortment of properties related to the text's paragraph, fill, and line styling, as shown in :numref:`ch29fig_title_srv`.

..
    figure 5

.. cssclass:: diagram invert

    .. _ch29fig_title_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/206548076-1598bf2a-55ed-450a-b2f8-febf270e8ff3.png
        :alt: The Title Service.
        :figclass: align-center

        :The Title_ Service.

Text is added to the XTitle_ object by :py:meth:`.Chart2.create_title`, as an XFormattedString_ array:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def create_title(title: str) -> XTitle:
            try:
                xtitle = Lo.create_instance_mcf(XTitle, "com.sun.star.chart2.Title", raise_err=True)
                xtitle_str = Lo.create_instance_mcf(
                    XFormattedString, "com.sun.star.chart2.FormattedString", raise_err=True
                )
                xtitle_str.setString(title)
                title_arr = (xtitle_str,)
                xtitle.setText(title_arr)
                return xtitle
            except Exception as e:
                raise ChartError(f'Error creating title for: "{title}"') from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The use of an XFormattedString_ tuple (``title_arr = (xtitle_str,)``) may seem to be overkill when the title is a single string,
but it also allows character properties to be associated with the string through XFormattedString2_, as shown in :numref:`ch29fig_fmt_str_srv`.

..
    figure 6

.. cssclass:: diagram invert

    .. _ch29fig_fmt_str_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/206551469-cba0a06d-a534-4c20-843d-2977b05501d1.png
        :alt: The FormattedString Service
        :figclass: align-center

        :The FormattedString_ Service.

Character properties allow the font and point size of the title to be changed to :spelling:word:`Arial` ``14pt`` by :py:meth:`.Chart2.set_x_title_font`:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def set_x_title_font(xtitle: XTitle, font_name: str, pt_size: int) -> None:
            try:
                fo_strs = xtitle.getText()
                if fo_strs:
                    Props.set_property(fo_strs[0], "CharFontName", font_name)
                    Props.set_property(fo_strs[0], "CharHeight", pt_size)
            except Exception as e:
                raise ChartError("Error setting x title font") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The ``CharFontName`` and ``CharHeight`` properties come from the CharacterProperties_ class.

.. _ch29_creating_axis_titles:

29.2 Creating Axis Titles
=========================

Setting the axes titles needs a reference to the XAxis_ interface.
Incidentally, this interface name is a little misleading since ``X`` is the naming convention for interfaces, not a reference to the ``x-axis``.

:numref:`ch28fig_chart_doc_hirarchy` shows that the XAxis_ interface is available via the XCoordinateSystem_ interface,
which can be obtained by calling :py:meth:`.Chart2.get_coord_system`.
``XCoordinateSystem.getAxisByDimension()`` can then be employed to get an axis reference.
This is implemented by :py:meth:`.Chart2.get_axis`:

.. tabs::

    .. code-tab:: python

        # in chart2 class
        @classmethod
        def get_axis(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int) -> XAxis:
            try:
                coord_sys = cls.get_coord_system(chart_doc)
                result = coord_sys.getAxisByDimension(int(axis_val), idx)
                if result is None:
                    raise UnKnownError("None Value: getAxisByDimension() returned None")
                return result
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error getting Axis for chart") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    :py:class:`~.kind.axis_kind.AxisKind`

``XCoordinateSystem.getAxisByDimension()`` takes two integer arguments: the first represents the axis (``x``, ``y``, or ``z``), while the second is a primary or secondary index (``0`` or ``1``) for the chosen axis.
:py:class:`~.chart2.Chart2` includes wrapper functions for :py:meth:`.Chart2.get_axis` for the most common cases:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def get_x_axis(cls, chart_doc: XChartDocument) -> XAxis:
            return cls.get_axis(chart_doc=chart_doc, axis_val=AxisKind.X, idx=0)

        @classmethod
        def get_y_axis(cls, chart_doc: XChartDocument) -> XAxis:
            return cls.get_axis(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=0)

        @classmethod
        def get_x_axis2(cls, chart_doc: XChartDocument) -> XAxis:
            return cls.get_axis(chart_doc=chart_doc, axis_val=AxisKind.X, idx=1)

        @classmethod
        def get_y_axis2(cls, chart_doc: XChartDocument) -> XAxis:
            return cls.get_axis(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=1)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Chart2.set_axis_title` calls :py:meth:`.Chart2.get_axis` to get a reference to the correct axis, and then reuses many of the methods described earlier for setting the chart title:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def set_axis_title(
            cls, chart_doc: XChartDocument, title: str, axis_val: AxisKind, idx: int
        ) -> XTitle:
            try:
                axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
                titled_axis = Lo.qi(XTitled, axis, True)
                xtitle = cls.create_title(title)
                titled_axis.setTitleObject(xtitle)
                fname = Info.get_font_general_name()
                cls.set_x_title_font(xtitle, fname, 12)
                return xtitle
            except ChartError:
                raise
            except Exception as e:
                raise ChartError(f'Error setting axis tile: "{title}" for chart') from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

As with :py:meth:`.Chart2.get_axis`, :py:class:`~.chart2.Chart2` includes wrapper methods for :py:meth:`.Chart2.set_axis_title` to simplify common axis cases:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def set_x_axis_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
            return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.X, idx=0)

        @classmethod
        def set_y_axis_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
            return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.Y, idx=0)

        @classmethod
        def set_x_axis2_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
            return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.X, idx=1)

        @classmethod
        def set_y_axis2_title(cls, chart_doc: XChartDocument, title: str) -> XTitle:
            return cls.set_axis_title(chart_doc=chart_doc, title=title, axis_val=AxisKind.Y, idx=1)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch29_rotating_axis_titles:

29.3 Rotating Axis Titles
=========================

The default orientation for titles is horizontal, which is fine for the chart and ``x-axis`` titles, but can cause the ``y-axis`` title to occupy too much horizontal space.
The solution is to call :py:meth:`.Chart2.rotate_y_axis_title` with an angle (usually 90 degrees) to turn the text counter-clockwise so it's vertically orientated (see :numref:`ch29fig_chart_for_sneaker_sold_month_tbl`).

The implementation accesses the XTitle_ interface for the axis title, and then modifies its ``TextRotation`` property from the Title_ service (see :numref:`ch29fig_title_srv`).

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def rotate_y_axis_title(cls, chart_doc: XChartDocument, angle: Angle) -> None:
            cls.rotate_axis_title(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=0, angle=angle)

        @classmethod
        def rotate_axis_title(
            cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int, angle: Angle
        ) -> None:
            try:
                xtitle = cls.get_axis_title(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
                Props.set(xtitle, TextRotation=angle.Value)
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error while trying to rotate axis title") from e

        @classmethod
        def get_axis_title(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int) -> XTitle:
            try:
                axis = cls.get_axis(chart_doc=chart_doc, axis_val=axis_val, idx=idx)
                titled_axis = Lo.qi(XTitled, axis, True)
                result = titled_axis.getTitleObject()
                if result is None:
                    raise UnKnownError("None Value: getTitleObject() return a value of None")
                return result
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error getting axis title") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch29_what_chart_templates:

29.4 What Chart Templates are Available?
========================================

``_col_chart()`` in |chart_2_views_py|_ returns its XChartDocument_ reference.
This isn't necessary for rendering the chart, but allows the reference to be passed to :py:meth:`.Chart2.get_chart_templates`:

.. tabs::

    .. code-tab:: python

        # in main() of chart_2_views.py
        # ...
        chart_doc = self._col_chart(doc=doc, sheet=sheet)
        # ...
        template_names = Chart2.get_chart_templates(chart_doc)
        Lo.print_names(template_names, 1)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The only way to list the chart templates supported by the ``chart2`` module (:abbreviation:`i.e.` those shown in :numref:`ch28tblchart_types_and_template_names`) is by querying an existing chart document.
That's the purpose of :py:meth:`.Chart2.get_chart_templates`:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def get_chart_templates(chart_doc: XChartDocument) -> List[str]:
            try:
                ct_man = chart_doc.getChartTypeManager()
                return Info.get_available_services(ct_man)
            except Exception as e:
                raise ChartError("Error getting chart templates") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Normally XChartTypeManager_ is used to create a template instance, but :py:meth:`.Info.get_available_services` accesses its ``XMultiServiceFactory.getAvailableServiceNames()``
method to list the names of all its supported services, which are templates:

.. tabs::

    .. code-tab:: python

        # in Info class
        @staticmethod
        def get_available_services(obj: object) -> List[str]:
            services: List[str] = []
            try:
                sf = Lo.qi(XMultiServiceFactory, obj, True)
                service_names = sf.getAvailableServiceNames()
                services.extend(service_names)
                services.sort()
            except Exception as e:
                Lo.print(e)
                raise Exception() from e
            return services

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The output lists has 64 names, same as :numref:`ch28tblchart_types_and_template_names`, starting and ending like so:

::

    com.sun.star.chart2.template.Area
    com.sun.star.chart2.template.Bar
    com.sun.star.chart2.template.Bubble
    com.sun.star.chart2.template.Column
    :
    com.sun.star.chart2.template.ThreeDLineDeep
    com.sun.star.chart2.template.ThreeDPie
    com.sun.star.chart2.template.ThreeDPieAllExploded
    com.sun.star.chart2.template.ThreeDScatter

.. _ch29_multiple_col:

29.5 Multiple Columns
=====================

The ``_mult_col_chart()`` method in |chart_2_views_py|_ uses a table containing three columns of data (see :numref:`ch29fig_tbl_most_colleges_by_state`)
to generate two column graphs in the same chart, as in :numref:`ch29fig_multi_col_chart_frm_07`.

..
    figure 7

.. cssclass:: screen_shot invert

    .. _ch29fig_tbl_most_colleges_by_state:
    .. figure:: https://user-images.githubusercontent.com/4193389/206601488-c64ac4e5-0cac-47bb-94bc-0533fdee782c.png
        :alt: The States with the Most Colleges Table
        :figclass: align-center

        :The "States with the Most Colleges" Table.

..
    figure 8

.. cssclass:: screen_shot

    .. _ch29fig_multi_col_chart_frm_07:
    .. figure:: https://user-images.githubusercontent.com/4193389/206601866-cc0dbe49-6343-406b-8925-57d53df2b969.png
        :alt: A Multiple Column Chart Generated from the Table in previous figure
        :figclass: align-center

        :A Multiple Column Chart Generated from the Table in :numref:`ch29fig_tbl_most_colleges_by_state`.

``_mult_col_chart()`` is:

.. tabs::

    .. code-tab:: python

        # 
        def _mult_col_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> XChartDocument:
            range_addr = Calc.get_address(sheet=sheet, range_name="E15:G21")
            d_name = ChartTypes.Column.TEMPLATE_STACKED.COLUMN
            # d_name = ChartTypes.Column.TEMPLATE_PERCENT.COLUMN_DEEP_3D
            # d_name = ChartTypes.Column.TEMPLATE_PERCENT.COLUMN_FLAT_3D
            chart_doc = Chart2.insert_chart(
                sheet=sheet,
                cells_range=range_addr,
                cell_name="A22",
                width=20,
                height=11,
                diagram_name=d_name,
            )
            ChartTypes.Column.TEMPLATE_STACKED.COLUMN
            Calc.goto_cell(cell_name="A13", doc=doc)

            Chart2.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E13"))
            Chart2.set_x_axis_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E15")
            )
            Chart2.set_y_axis_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="F14")
            )
            Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
            Chart2.view_legend(chart_doc=chart_doc, is_visible=True)

            # for the 3D versions
            # Chart2.show_axis_label(chart_doc=chart_doc, axis_val=AxisKind.Z, idx=0, is_visible=False)
            # Chart2.set_chart_shape_3d(chart_doc=chart_doc, shape=DataPointGeometry3DEnum.CYLINDER)
            return chart_doc

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The same ``Column`` chart template is used as in ``_col_chart()``, and the additional column of data is treated as an extra column graph.
The chart title and axis titles are added in the same way as before, and a legend is included by calling :py:meth:`.Chart2.view_legend`:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def view_legend(chart_doc: XChartDocument, is_visible: bool) -> None:
            try:
                diagram = chart_doc.getFirstDiagram()
                legend = diagram.getLegend()
                if is_visible and legend is None:
                    leg = Lo.create_instance_mcf(XLegend, "com.sun.star.chart2.Legend", raise_err=True)
                    Props.set(
                        leg,
                        LineStyle=LineStyle.NONE,
                        FillStyle=FillStyle.SOLID,
                        FillTransparence=100
                    )
                    diagram.setLegend(leg)

                Props.set(leg, Show=is_visible)
            except Exception as e:
                raise ChartError("Error while setting legend visibility") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The legend is accessible via the chart Diagram_ service.
:py:meth:`~.Chart2.view_legend` creates an instance, and sets a few properties to make it look nicer.

:numref:`ch29fig_legned_srv` shows the Legend service, which defines several properties, and inherits many others from FillProperties_ and LineProperties_.
The ``LineStyle``, ``FillStyle``, and ``FillTransparence`` properties utilized in :py:meth:`~.Chart2.view_legend` come from the inherited property classes, but ``Show`` is from the Legend_ service.

..
    figure 9

.. cssclass:: diagram invert

    .. _ch29fig_legned_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/206604671-eb2735fd-c6e4-4a3c-b7a8-39350dee90ec.png
        :alt: The Legend Service.
        :figclass: align-center

        :The Legend_ Service.

The XLegend_ interface contains no methods, and is used only to access the properties in its defining service.

.. _ch29_3d_pizazz:

29.6 3D Pizazz
==============

You may not be a fan of 3D charts which are often harder to understand than their 2D equivalents, even if they do look more "hi-tech".
But if you really want a 3D version of a chart, it's mostly just a matter of changing the template name in the call to :py:meth:`.Chart2.insert_chart`.

If ``d_name`` were were set to enum value of ``ChartTypes.Column.TEMPLATE_PERCENT.COLUMN_DEEP_3D`` or string value of ``ThreeDColumnDeep``
or enum value of ``ChartTypes.Column.TEMPLATE_PERCENT.COLUMN_FLAT_3D`` or string value of ``ThreeDColumnFlat`` in ``_mult_col_chart()``, then the charts in :numref:`ch29fig_deep_flat_col_charts` appear.

..
    figure 10

.. cssclass:: screen_shot

    .. _ch29fig_deep_flat_col_charts:
    .. figure:: https://user-images.githubusercontent.com/4193389/206615092-b69c0154-ae99-4a9b-aa9c-26d2078aea29.png
        :alt: Deep and Flat 3D Column Charts
        :figclass: align-center
        :width: 440px

        :Deep and Flat 3D Column Charts

``deep`` orders the two 3D graphs along the ``z-axis``, and labels the axis.

The ``x-axis`` labels are rotated automatically in the top-most chart of :numref:`ch29fig_deep_flat_col_charts` because the width of the chart wasn't sufficient to draw them horizontally,
and that's caused the graphs to be squashed into less vertical space.

``_mult_col_chart()`` contains two commented out lines which illustrate how a 3D graph can be changed:

.. tabs::

    .. code-tab:: python

        # in _mult_col_chart()...
        # hide labels
        Chart2.show_axis_label(chart_doc=chart_doc, axis_val=AxisKind.Z, idx=0, is_visible=False)
        Chart2.set_chart_shape_3d(chart_doc=chart_doc, shape=DataPointGeometry3DEnum.CYLINDER)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    DataPointGeometry3D_

:py:meth:`.Chart2.show_axis_label` is passed the boolean ``False`` to switch off the display of the ``z-axis`` labels.
:py:meth:`.Chart2.set_chart_shape_3d` changes the shape of the columns; in this case to cylinders, as in :numref:`ch29fig_modified_deep_3d_col_chart`.

..
    figure 11

.. cssclass:: screen_shot

    .. _ch29fig_modified_deep_3d_col_chart:
    .. figure:: https://user-images.githubusercontent.com/4193389/206613695-f91ca702-ce14-4c6c-9c6d-ed0ad3776022.png
        :alt: Modified Deep 3D Column Chart
        :figclass: align-center
        :width: 550px

        :Modified Deep 3D Column Chart.

:py:meth:`.Chart2.show_axis_label` uses :py:meth:`.Chart2.get_axis` to access the XAxis_ interface, and then modifies its ``Show`` property:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def get_axis(cls, chart_doc: XChartDocument, axis_val: AxisKind, idx: int) -> XAxis:
            try:
                coord_sys = cls.get_coord_system(chart_doc)
                result = coord_sys.getAxisByDimension(int(axis_val), idx)
                if result is None:
                    raise UnKnownError("None Value: getAxisByDimension() returned None")
                return result
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error getting Axis for chart") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The Axis_ service contains a large assortment of properties, and inherits character and line properties depicted in :numref:`ch29fig_axis_srv`.

..
    figure 12

.. cssclass:: diagram invert

    .. _ch29fig_axis_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/206617330-0cfe4198-a0d4-4c42-b558-f7f9dfaab2f2.png
        :alt: The Axis Service.
        :figclass: align-center

        :The Axis_ Service.

:py:meth:`.Chart2.set_chart_shape_3d` affects the data ``points`` (which in a 3D column chart are boxes by default).
This requires access to the XDataSeries_ array of data points by calling :py:meth:`.Chart2.get_data_series`,
and then the ``Geometry3D`` property in the DataSeries_ service is modified.
:numref:`ch28fig_coordinate_system_service` shows the service and its interfaces, and most of its properties are inherited from the DataPointProperties_ class, including ``Geometry3D``.
The code for :py:meth:`.Chart2.set_chart_shape_3d`:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def set_chart_shape_3d(cls, chart_doc: XChartDocument, shape: DataPointGeometry3DEnum) -> None:
            try:
                data_series_arr = cls.get_data_series(chart_doc=chart_doc)
                for data_series in data_series_arr:
                    Props.set_property(data_series, "Geometry3D", int(shape))
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error setting chart shape 3d") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch29_the_col_line_chart:

29.7 The Column and Line Chart
==============================

Another way to display the multiple columns of data in the "States with the Most Colleges" table (:numref:`ch29fig_tbl_most_colleges_by_state`) is to draw a column and line chart.
The column is generated from the first data column, and the line graph uses the second column.
The result is shown in :numref:`ch29fig_col_line_of_data_07_tbl`.

..
    figure 13

.. cssclass:: screen_shot

    .. _ch29fig_col_line_of_data_07_tbl:
    .. figure:: https://user-images.githubusercontent.com/4193389/206618602-67c11866-a308-4077-9811-ec6aa1dd1576.png
        :alt: A Column and Line Chart Generated from the Table in Figure 7
        :figclass: align-center
        :width: 550px

        :A Column and Line Chart Generated from the Table in :numref:`ch29fig_tbl_most_colleges_by_state`.

``_col_line_chart()`` in |chart_2_views_py|_ generates :numref:`ch29fig_col_line_of_data_07_tbl`:

.. tabs::

    .. code-tab:: python

        # Chart2View._col_line_chart() in chart_2_views.py
        def _col_line_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> XChartDocument:
            range_addr = Calc.get_address(sheet=sheet, range_name="E15:G21")
            chart_doc = Chart2.insert_chart(
                sheet=sheet,
                cells_range=range_addr,
                cell_name="B3",
                width=20,
                height=11,
                diagram_name=ChartTypes.ColumnAndLine.TEMPLATE_STACKED.COLUMN_WITH_LINE,
            )
            Calc.goto_cell(cell_name="A13", doc=doc)

            Chart2.set_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E13")
            )
            Chart2.set_x_axis_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E15")
            )
            Chart2.set_y_axis_title(
                chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="F14")
            )
            Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
            Chart2.view_legend(chart_doc=chart_doc, is_visible=True)
            return chart_doc

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

It's nearly identical to ``_mul_col_chart()`` except for ``ChartTypes.ColumnAndLine.TEMPLATE_STACKED.COLUMN_WITH_LINE`` passed to :py:meth:`.Chart2.insert_chart`.

A chart's coordinate system may utilize multiple chart types.
Up to now a chart template (:abbreviation:`i.e.` ``Column``) has been converted to a single chart type (:abbreviation:`i.e.` ``ColumnChartType``) by the chart API (specifically by the chart type manager),
but the ``ColumnWithLine`` template is different. The manager implements that template using two chart types, ``ColumnChartType`` and ``LineChartType``.
This is reported by :py:meth:`.Chart2.insert_chart` calling :py:meth:`.Chart2.print_chart_types`:

::

    No. of chart types: 2
      com.sun.star.chart2.ColumnChartType
      com.sun.star.chart2.LineChartType

:py:meth:`.Chart2.print_chart_types` uses :py:meth:`.Chart2.get_chart_types`, which was defined earlier:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def print_chart_types(cls, chart_doc: XChartDocument) -> None:
            chart_types = cls.get_chart_types(chart_doc)
            if len(chart_types) > 1:
                print(f"No. of chart types: {len(chart_types)}")
                for ct in chart_types:
                    print(f"  {ct.getChartType()}")
            else:
                print(f"Chart Type: {chart_types[0].getChartType()}")
            print()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Why is this separation of a single template into two chart types important?
The short answer is that it complicates the search for a chart template's data.
For example :py:meth:`.Chart2.get_chart_type` returns the first chart type in the XChartType_ array since most templates only use a single chart type:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def get_chart_type(cls, chart_doc: XChartDocument) -> XChartType:
            try:
                chart_types = cls.get_chart_types(chart_doc)
                return chart_types[0]
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error getting chart type") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This method is insufficient for examining a chart created with the ``ColumnWithLine`` template since the XChartType_ array holds two chart types.
A programmer will have to use :py:meth:`.Chart2.find_chart_type`, which searches the array for the specified chart type:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def find_chart_type(
            cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase | str
        ) -> XChartType:
            # Ensure chart_type is ChartTypeNameBase | str
            Info.is_type_enum_multi(
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
                raise ChartError(f'Error Finding chart for "{chart_type}"') from e
            raise NotFoundError(f'Chart for type "{chart_type}" was not found')

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

For example, the following call returns a reference to the line chart type:

.. tabs::

    .. code-tab:: python

        line_ct = Chart2.find_chart_type(chart_doc=chart_doc, chart_type="LineChartType") # XChartType

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The simple :py:meth:`~.Chart2.get_chart_type` is used in :py:meth:`.Chart2.get_data_series`:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def get_data_series(
            cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase | str = ""
        ) -> Tuple[XDataSeries, ...]:
            try:
                if chart_type:
                    xchart_type = cls.find_chart_type(chart_doc, chart_type)
                else:
                    xchart_type = cls.get_chart_type(chart_doc)
                ds_con = Lo.qi(XDataSeriesContainer, xchart_type, True)
                return ds_con.getDataSeries()
            except Exception as e:
                raise ChartError("Error getting chart data series") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

When ``chart_type`` is omitted it means that :py:meth:`.Chart2.get_data_series` can only access the data associated with the column (the first chart type) in a ``ColumnWithLine`` chart document.

When ``chart_type`` is included it requires a ``chart_type`` argument to get the correct chart type. For example, the call:

.. tabs::

    .. code-tab:: python

        ds = Chart2.get_data_series(chart_doc=chart_doc, chart_type=ChartTypes.Line.NAMED.LINE_CHART)
        #   chart_type could also be "LineChartType"

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

returns the data series associated with the line chart type.


.. |chart_2_views| replace:: Chart2 Views
.. _chart_2_views: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/chart2/Chart_2_Views

.. |chart_2_views_py| replace:: chart_2_views.py
.. _chart_2_views_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/chart2/Chart_2_Views/chart_2_views.py

.. _CharacterProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
.. _Diagram: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1Diagram.html
.. _FillProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1FillProperties.html
.. _FormattedString: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1FormattedString.html
.. _Legend: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1Legend.html
.. _LineProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineProperties.html
.. _Title: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1Title.html
.. _XAxis: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XAxis.html
.. _XChartDocument: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XChartDocument.html
.. _XChartTypeManager: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XChartTypeManager.html
.. _XCoordinateSystem: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XCoordinateSystem.html
.. _XFormattedString: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XFormattedString.html
.. _XFormattedString2: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XFormattedString2.html
.. _XLegend: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XLegend.html
.. _XTitle: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XTitle.html
.. _XTitled: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XTitled.html
.. _DataPointGeometry3D: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1chart2_1_1DataPointGeometry3D.html
.. _Axis: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1Axis.html
.. _XDataSeries: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XDataSeries.html
.. _DataSeries: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1DataSeries.html
.. _DataPointProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1DataPointProperties.html
.. _XChartType: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XChartType.html
