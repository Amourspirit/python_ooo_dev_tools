.. _help_chart2_format_direct_legend_position:

Chart2 Direct Legend Position
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 1

Overview
--------

The Legend parts of a Chart can be styled using the various ``style_*`` methods of the :py:class:`~ooodev.calc.chart2.chart_legend.ChartLegend` class.

Here we will see how the ``style_position()`` method can be used to set the position of the legend.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 40,41,42,43,44

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooo.dyn.chart2.legend_position import LegendPosition 
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor
        from ooodev.utils.data_type.color_range import ColorRange
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.format.inner.direct.chart2.title.alignment.direction import DirectionModeKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "piechart.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_border_line(
                    color=StandardColor.BRICK,
                    width=1,
                )
                _ = chart_doc.style_area_gradient(
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(
                        StandardColor.GREEN_DARK4,
                        StandardColor.TEAL_LIGHT2,
                    ),
                )
                legend = chart_doc.first_diagram.get_legend()
                if legend is None:
                    raise ValueError("Legend is None")

                _ = legend.style_position(
                    pos=LegendPosition.PAGE_END,
                    no_overlap=True,
                    mode=DirectionModeKind.LR_TB,
                )

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the Legend Position
---------------------------

.. tabs::

    .. code-tab:: python

        from ooo.dyn.chart2.legend_position import LegendPosition
        from ooodev.format.inner.direct.chart2.title.alignment.direction import DirectionModeKind
        # ... other code

        _ = legend.style_position(
            pos=LegendPosition.PAGE_END,
            no_overlap=True,
            mode=DirectionModeKind.LR_TB,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`14a18c7b-b6ae-4b8a-baac-b1929fca5b2d_1` and :numref:`0c899f9c-4b39-4607-8553-c3bc4b8ec29f_1`.


.. cssclass:: screen_shot

    .. _14a18c7b-b6ae-4b8a-baac-b1929fca5b2d_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/14a18c7b-b6ae-4b8a-baac-b1929fca5b2d
        :alt: Chart with Legend Set to Bottom
        :figclass: align-center
        :width: 450px

        Chart with Legend Set to Bottom

.. cssclass:: screen_shot

    .. _0c899f9c-4b39-4607-8553-c3bc4b8ec29f_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/0c899f9c-4b39-4607-8553-c3bc4b8ec29f
        :alt: Chart Legend Position Dialog
        :figclass: align-center
        :width: 450px

        Chart Legend Position Dialog

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_general_area`
        - :ref:`help_chart2_format_direct_legend_transparency`
        - :py:class:`~ooodev.loader.Lo`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.legend.position_size.Position`