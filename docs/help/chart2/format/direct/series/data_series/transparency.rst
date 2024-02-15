.. _help_chart2_format_direct_series_series_transparency:

Chart2 Direct Series Data Series Transparency
=============================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Overview
--------

The Data Series and Data Point of a Chart can be styled using the various ``style_*`` methods of
the :py:class:`~ooodev.calc.chart2.chart_data_series.ChartDataSeries` and :py:class:`~ooodev.calc.chart2.chart_data_point.ChartDataPoint` classes.


Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 29

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor
        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "col_chart.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_border_line(
                    color=StandardColor.BLUE_LIGHT3,
                    width=0.7,
                )
                _ = chart_doc.style_area_gradient_from_preset(
                    preset=PresetGradientKind.TEAL_BLUE,
                )

                ds = chart_doc.get_data_series()[0]
                ds.style_area_transparency_transparency(50)

                Lo.delay(1_000)
                doc.close()
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

The ``style_area_transparency_transparency()`` method can be called to set the data series transparency of a chart.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        ds = chart_doc.get_data_series()[0]
        ds.style_area_transparency_transparency(50)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`1c71f71a-ea08-4d47-abbb-55738998a182_1` and :numref:`ea9c0a9a-d069-49dd-99a7-314894eea02e_1`.

.. cssclass:: screen_shot

    .. _1c71f71a-ea08-4d47-abbb-55738998a182_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1c71f71a-ea08-4d47-abbb-55738998a182
        :alt: Chart with data series transparency set
        :figclass: align-center
        :width: 450px

        Chart with data series transparency set

.. cssclass:: screen_shot

    .. _ea9c0a9a-d069-49dd-99a7-314894eea02e_1:

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
        ds = chart_doc.get_data_series()[0]
        dp = ds[-1]
        dp.style_area_transparency_transparency(50)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`7cbe60a0-cbc8-4c50-8d79-f69fe0c055ae_1`.

.. cssclass:: screen_shot

    .. _7cbe60a0-cbc8-4c50-8d79-f69fe0c055ae_1:

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

The ``style_area_transparency_gradient()`` method can be called to set the data series gradient transparency of a chart.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.utils.data_type.intensity_range import IntensityRange
        from ooodev.utils.data_type.angle import Angle
        # ... other code

        ds = chart_doc.get_data_series()[0]
        ds.style_area_transparency_gradient(
            angle=30,
            grad_intensity=IntensityRange(0, 100),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`f2eea034-d414-4e70-9fe8-701968ad1304_1` and :numref:`392d7295-8cbb-4eed-8955-8ba481ea0db8_1`.

.. cssclass:: screen_shot

    .. _f2eea034-d414-4e70-9fe8-701968ad1304_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f2eea034-d414-4e70-9fe8-701968ad1304
        :alt: Chart data series with gradient transparency set
        :figclass: align-center
        :width: 450px

        Chart data series with gradient transparency set

.. cssclass:: screen_shot

    .. _392d7295-8cbb-4eed-8955-8ba481ea0db8_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/392d7295-8cbb-4eed-8955-8ba481ea0db8
        :alt: Chart Data Series Area Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Area Transparency Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.utils.data_type.intensity_range import IntensityRange

        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[-1]
        dp.style_area_transparency_gradient(
            angle=30,
            grad_intensity=IntensityRange(0, 100),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`bd61630d-0f6d-45ed-bcb0-f194c233b81e_1`.

.. cssclass:: screen_shot

    .. _bd61630d-0f6d-45ed-bcb0-f194c233b81e_1:

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
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_wall_floor_area`
        - :py:class:`~ooodev.loader.Lo`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`