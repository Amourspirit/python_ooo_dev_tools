.. _help_chart2_format_direct_static_series_series_transparency:

Chart2 Direct Series Data Series Transparency (Static)
======================================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Overview
--------

Classes in the :py:mod:`ooodev.format.chart2.direct.series.data_series.transparency` module contains classes that are used to set the data series transparency of a chart.

Calls to the :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>` method are used to set the data series transparency of a Chart.

Calls to the :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
and :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>` methods are used to set the data series transparency of a Chart.

.. seealso::

    - :ref:`help_chart2_format_direct_series_series_transparency`

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 27, 28

        import uno
        from ooodev.format.chart2.direct.series.data_series.transparency import Transparency as SeriesTransparency
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient, PresetGradientKind
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

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

                data_series_transparency = SeriesTransparency(value=50)
                Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_transparency])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Transparency
------------

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Setting Transparency
^^^^^^^^^^^^^^^^^^^^

The :py:class:`ooodev.format.chart2.direct.series.data_series.transparency.Transparency` class can be used to set the data series transparency of a chart.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.transparency import Transparency as SeriesTransparency
        # ... other code

        data_series_transparency = SeriesTransparency(value=50)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_transparency])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`1c71f71a-ea08-4d47-abbb-55738998a182` and :numref:`ea9c0a9a-d069-49dd-99a7-314894eea02e`.

.. cssclass:: screen_shot

    .. _1c71f71a-ea08-4d47-abbb-55738998a182:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1c71f71a-ea08-4d47-abbb-55738998a182
        :alt: Chart with data series transparency set
        :figclass: align-center
        :width: 450px

        Chart with data series transparency set

.. cssclass:: screen_shot

    .. _ea9c0a9a-d069-49dd-99a7-314894eea02e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ea9c0a9a-d069-49dd-99a7-314894eea02e
        :alt: Chart Data Series Area Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Area Transparency Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_data_point(
            chart_doc=chart_doc, series_idx=0, idx=-1, styles=[data_series_transparency]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`7cbe60a0-cbc8-4c50-8d79-f69fe0c055ae`.

.. cssclass:: screen_shot

    .. _7cbe60a0-cbc8-4c50-8d79-f69fe0c055ae:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7cbe60a0-cbc8-4c50-8d79-f69fe0c055ae
        :alt: Chart with data point transparency set
        :figclass: align-center
        :width: 450px

        Chart with data point transparency set


Gradient Transparency
---------------------

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Setting Gradient
^^^^^^^^^^^^^^^^

The :py:class:`ooodev.format.chart2.direct.series.data_series.transparency.Gradient` class can be used to set the data series gradient transparency of a chart.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.transparency import Gradient as SeriesGradient
        from ooodev.format.chart2.direct.series.data_series.transparency import IntensityRange
        from ooodev.utils.data_type.angle import Angle
        # ... other code

        data_series_grad_transparency = SeriesGradient(
            chart_doc=chart_doc, angle=Angle(30), grad_intensity=IntensityRange(0, 100)
        )
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_grad_transparency])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`f2eea034-d414-4e70-9fe8-701968ad1304` and :numref:`392d7295-8cbb-4eed-8955-8ba481ea0db8`.

.. cssclass:: screen_shot

    .. _f2eea034-d414-4e70-9fe8-701968ad1304:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f2eea034-d414-4e70-9fe8-701968ad1304
        :alt: Chart data series with gradient transparency set
        :figclass: align-center
        :width: 450px

        Chart data series with gradient transparency set

.. cssclass:: screen_shot

    .. _392d7295-8cbb-4eed-8955-8ba481ea0db8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/392d7295-8cbb-4eed-8955-8ba481ea0db8
        :alt: Chart Data Series Area Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Area Transparency Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_data_point(
            chart_doc=chart_doc, series_idx=0, idx=-1, styles=[data_series_grad_transparency]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`bd61630d-0f6d-45ed-bcb0-f194c233b81e`.

.. cssclass:: screen_shot

    .. _bd61630d-0f6d-45ed-bcb0-f194c233b81e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bd61630d-0f6d-45ed-bcb0-f194c233b81e
        :alt: Chart data point with gradient transparency set
        :figclass: align-center
        :width: 450px

        Chart data point with gradient transparency set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_chart2_format_direct_series_series_transparency`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_wall_floor_area`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
        - :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.series.data_series.transparency.Transparency`
        - :py:class:`ooodev.format.chart2.direct.series.data_series.transparency.Gradient`
