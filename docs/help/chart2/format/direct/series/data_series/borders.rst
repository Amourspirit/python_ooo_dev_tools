.. _help_chart2_format_direct_series_series_borders:

Chart2 Direct Series Data Series Borders
========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The Data Series and Data Point of a Chart can be styled using the various ``style_*`` methods of
the :py:class:`~ooodev.calc.chart2.chart_data_series.ChartDataSeries` and :py:class:`~ooodev.calc.chart2.chart_data_point.ChartDataPoint` classes.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 29,30,31,32

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
                ds.style_border_line(
                    color=StandardColor.MAGENTA_DARK1,
                    width=0.75,
                )

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting Line Properties
-----------------------

The ``style_border_line()`` method is called to set the data series border line properties.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        ds = chart_doc.get_data_series()[0]
        ds.style_border_line(
            color=StandardColor.MAGENTA_DARK1,
            width=0.75,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`f462c874-3624-4eaa-898f-ea79e4b98bc4_1` and :numref:`cc6bba18-1fcd-4188-a0c5-14e8dbed654d_1`.


.. cssclass:: screen_shot

    .. _f462c874-3624-4eaa-898f-ea79e4b98bc4_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f462c874-3624-4eaa-898f-ea79e4b98bc4
        :alt: Chart with data series border set
        :figclass: align-center
        :width: 450px

        Chart with data series border set

.. cssclass:: screen_shot

    .. _cc6bba18-1fcd-4188-a0c5-14e8dbed654d_1:

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
        ds = chart_doc.get_data_series()[0]
        dp = ds[1]
        dp.style_border_line(
            color=StandardColor.MAGENTA_DARK1,
            width=0.75,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`8a2b209b-b856-43fb-9df1-9f74bad97d96_1`.


.. cssclass:: screen_shot

    .. _8a2b209b-b856-43fb-9df1-9f74bad97d96_1:

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
        - :py:class:`~ooodev.loader.Lo`