.. _help_chart2_format_direct_series_series_borders:

Chart2 Direct Series Data Series Borders
========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.series.data_series.borders.LineProperties` class gives the same options as the Chart Data Series Borders dialog
as seen in :numref:`c4cc6299-704d-40be-8a8b-68daff8c5eef`.

Calls to the :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
and :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>` methods are used to set the data series borders of a Chart.

.. cssclass:: screen_shot

    .. _c4cc6299-704d-40be-8a8b-68daff8c5eef:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/c4cc6299-704d-40be-8a8b-68daff8c5eef
        :alt: Chart Data Series Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Borders Default Dialog

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 27, 28

        import uno
        from ooodev.format.chart2.direct.series.data_series.borders import LineProperties as SeriesLineProperties
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient, PresetGradientKind
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("col_chart.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="col_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.BLUE_LIGHT3, width=0.7)
                chart_grad = ChartGradient.from_preset(chart_doc, PresetGradientKind.TEAL_BLUE)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                data_series_border = SeriesLineProperties(color=StandardColor.MAGENTA_DARK1, width=0.75)
                Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_border])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting Line Properties
-----------------------

The :py:class:`~ooodev.format.chart2.direct.general.borders.LineProperties` class is used to set the data series border line properties.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        data_series_border = SeriesLineProperties(color=StandardColor.MAGENTA_DARK1, width=0.75)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_border])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`f462c874-3624-4eaa-898f-ea79e4b98bc4` and :numref:`cc6bba18-1fcd-4188-a0c5-14e8dbed654d`.


.. cssclass:: screen_shot

    .. _f462c874-3624-4eaa-898f-ea79e4b98bc4:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f462c874-3624-4eaa-898f-ea79e4b98bc4
        :alt: Chart with data series border set
        :figclass: align-center
        :width: 450px

        Chart with data series border set

.. cssclass:: screen_shot

    .. _cc6bba18-1fcd-4188-a0c5-14e8dbed654d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/cc6bba18-1fcd-4188-a0c5-14e8dbed654d
        :alt: Chart Data Series Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Borders Default Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=1, styles=[data_series_border])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`8a2b209b-b856-43fb-9df1-9f74bad97d96`.


.. cssclass:: screen_shot

    .. _8a2b209b-b856-43fb-9df1-9f74bad97d96:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/8a2b209b-b856-43fb-9df1-9f74bad97d96
        :alt: Chart with data point border set
        :figclass: align-center
        :width: 450px

        Chart with data point border set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_wall_floor_area`
        - :ref:`help_chart2_format_direct_series_labels_borders`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
        - :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.series.data_series.borders.LineProperties`