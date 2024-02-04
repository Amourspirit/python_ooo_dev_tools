.. _help_chart2_format_direct_series_series_options:

Chart2 Direct Series Data Series Options
========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Depending on the chart type the Data Series Options dialog will have different options.
For example a Pie chart will have a options dialog similar to :numref:`bb11ad9f-1220-41e7-921a-a69e96a47880`
whereas  a column chart will have a dialog similar to :numref:`1510f063-94f3-4b7c-a8ba-6be4c222f32e`.

The :py:mod:`ooodev.format.chart2.direct.series.data_series.options` module contains classes for the various options.

Calls to the :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>` method are used to set the data series options of a Chart.

Figures
^^^^^^^

.. cssclass:: screen_shot

    .. _bb11ad9f-1220-41e7-921a-a69e96a47880:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bb11ad9f-1220-41e7-921a-a69e96a47880
        :alt: Data Series Options Dialog Simple
        :figclass: align-center
        :width: 450px

        Data Series Options Dialog Simple

.. cssclass:: screen_shot

    .. _1510f063-94f3-4b7c-a8ba-6be4c222f32e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1510f063-94f3-4b7c-a8ba-6be4c222f32e
        :alt: Data Series Options Dialog Complex
        :figclass: align-center
        :width: 450px

        Data Series Options Dialog Complex

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 28, 29

        import uno
        from ooodev.format.chart2.direct.series.data_series.options import Orientation
        from ooodev.utils.data_type.angle import Angle
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient, PresetGradientKind
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("pie_chart.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="pie_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.BLUE_LIGHT3, width=0.7)
                chart_grad = ChartGradient.from_preset(chart_doc, PresetGradientKind.TEAL_BLUE)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                orient = Orientation(chart_doc=chart_doc, clockwise=True, angle=Angle(45))
                Chart2.style_data_series(chart_doc=chart_doc, styles=[orient])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Options for Orientation
-----------------------

Charts such as Pie and Donut have a Orientation option as shown in :numref:`bb11ad9f-1220-41e7-921a-a69e96a47880`.

With the :py:class:`~ooodev.format.chart2.direct.series.data_series.options.Orientation` class we can set the angle and direction of the chart.

Before formatting the chart is seen in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.options import Orientation
        # ... other code

        orient = Orientation(chart_doc=chart_doc, clockwise=True, angle=Angle(45))
        Chart2.style_data_series(chart_doc=chart_doc, styles=[orient])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`6066b9d9-a91a-4a58-855d-754a7fe24de6` and :numref:`44d8288f-2902-4951-84a7-2417e79181dd`.

.. cssclass:: screen_shot

    .. _6066b9d9-a91a-4a58-855d-754a7fe24de6:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/6066b9d9-a91a-4a58-855d-754a7fe24de6
        :alt: Chart with orientation set to clockwise and angle set to 45 degrees
        :figclass: align-center
        :width: 450px

        Chart with orientation set to clockwise and angle set to ``45`` degrees

.. cssclass:: screen_shot

    .. _44d8288f-2902-4951-84a7-2417e79181dd:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/44d8288f-2902-4951-84a7-2417e79181dd
        :alt: Chart Data Series options Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series options Dialog

Align Data Series
-----------------

The :py:class:`~ooodev.format.chart2.direct.series.data_series.options.AlignSeries` class can be used to align the data series.

In this example we set the plot options of a column chart as seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

The ``primary_y_axis`` parameter is used to set the alignment of the data series.
If ``True`` this the primary y-axis is used, if ``False`` the secondary y-axis is used.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.options import AlignSeries
        # ... other code

        align_options = AlignSeries(chart_doc, primary_y_axis=False)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[align_options])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`4b1bd75c-e191-46a2-8e5e-381619f2ca7a` and :numref:`d051087e-7c53-4f3d-aecc-827bd725ef4f`.

.. cssclass:: screen_shot

    .. _4b1bd75c-e191-46a2-8e5e-381619f2ca7a:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4b1bd75c-e191-46a2-8e5e-381619f2ca7a
        :alt: Chart with data series alignment set to secondary y-axis
        :figclass: align-center
        :width: 450px

        Chart with data series alignment set to secondary y-axis

.. cssclass:: screen_shot

    .. _d051087e-7c53-4f3d-aecc-827bd725ef4f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d051087e-7c53-4f3d-aecc-827bd725ef4f
        :alt: Chart Data Series options Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series options Dialog

Settings Options
----------------

The :py:class:`~ooodev.format.chart2.direct.series.data_series.options.Settings` class can be used to set the settings of the data series.

In this example we set the plot options of a column chart as seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.options import Settings
        # ... other code

        setting_options = Settings(
            chart_doc=chart_doc, spacing=150, overlap=22, side_by_side=True
        )
        Chart2.style_data_series(chart_doc=chart_doc, styles=[setting_options])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`6b406b23-68c3-4d75-a36c-a7a7f2df7d02` and :numref:`5d5bd1bf-5232-4847-9996-a24596c5bfd8`.

.. cssclass:: screen_shot

    .. _6b406b23-68c3-4d75-a36c-a7a7f2df7d02:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/6b406b23-68c3-4d75-a36c-a7a7f2df7d02
        :alt: Chart with data series without legend
        :figclass: align-center
        :width: 450px

        Chart with data series without legend

.. cssclass:: screen_shot

    .. _5d5bd1bf-5232-4847-9996-a24596c5bfd8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/5d5bd1bf-5232-4847-9996-a24596c5bfd8
        :alt: Chart Data Series options Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series options Dialog

Options for Plot
----------------

Some charts such as Pie and Donut have simple Plot options as shown in :numref:`bb11ad9f-1220-41e7-921a-a69e96a47880`.

Other charts have more complex Plot options as shown in :numref:`1510f063-94f3-4b7c-a8ba-6be4c222f32e`.

Before formatting the chart is seen in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

Simple Plot Options
^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.chart2.direct.series.data_series.options.PlotSimple` class can be used to set the hidden cell values.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.options import PlotSimple
        # ... other code

        plot_options = PlotSimple(chart_doc=chart_doc, hidden_cell_values=False)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[plot_options])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`4d67d921-c574-4fe9-9505-543608a600b7`.

.. cssclass:: screen_shot

    .. _4d67d921-c574-4fe9-9505-543608a600b7:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4d67d921-c574-4fe9-9505-543608a600b7
        :alt: Chart Data Series options Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series options Dialog

Complex Plot Options
^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.chart2.direct.series.data_series.options.Plot` class can be used to set the complex options.

In this example we set the plot options of a column chart as seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.options import Plot, MissingValueKind
        # ... other code

        plot_options = Plot(
            chart_doc=chart_doc, missing_values=MissingValueKind.USE_ZERO, hidden_cell_values=False
        )
        Chart2.style_data_series(chart_doc=chart_doc, styles=[plot_options])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`4b69ef08-775e-4574-a552-db1cb001b4c8`.

.. cssclass:: screen_shot

    .. _4b69ef08-775e-4574-a552-db1cb001b4c8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4b69ef08-775e-4574-a552-db1cb001b4c8
        :alt: Chart Data Series options Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series options Dialog

Legend Options
--------------

The :py:class:`~ooodev.format.chart2.direct.series.data_series.options.LegendEntry` class can be used to set the legend visibility of the data series.

In this example we set the plot options of a column chart as seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.options import LegendEntry
        # ... other code

        legend_options = LegendEntry(chart_doc, hide_legend=True)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[legend_options])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`42e38398-7258-4bd2-9de7-232fc8e8df7a` and :numref:`bf56acb0-5486-4ff8-898b-d4a1d5e14661`.

.. cssclass:: screen_shot

    .. _42e38398-7258-4bd2-9de7-232fc8e8df7a:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/42e38398-7258-4bd2-9de7-232fc8e8df7a
        :alt: Chart with data series without legend
        :figclass: align-center
        :width: 450px

        Chart with data series without legend

.. cssclass:: screen_shot

    .. _bf56acb0-5486-4ff8-898b-d4a1d5e14661:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bf56acb0-5486-4ff8-898b-d4a1d5e14661
        :alt: Chart Data Series options Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series options Dialog

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.series.data_series.options.Orientation`
        - :py:class:`ooodev.format.chart2.direct.series.data_series.options.AlignSeries`
        - :py:class:`ooodev.format.chart2.direct.series.data_series.options.Settings`
        - :py:class:`ooodev.format.chart2.direct.series.data_series.options.PlotSimple`
        - :py:class:`ooodev.format.chart2.direct.series.data_series.options.Plot`
        - :py:class:`ooodev.format.chart2.direct.series.data_series.options.LegendEntry`