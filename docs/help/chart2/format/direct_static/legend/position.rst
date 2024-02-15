.. _help_chart2_format_direct_static_legend_position:

Chart2 Direct Legend Position
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 1

Overview
--------

The :py:class:`ooodev.format.chart2.direct.legend.position_size.Position` class is used to specify the position of the legend.

Calls to the :py:meth:`Chart2.style_legend() <ooodev.office.chart2.Chart2.style_legend>` and method is used to style legend.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 39,40,41

        import uno
        from ooodev.format.chart2.direct.legend.position_size import (
            Position as ChartLegendPosition,
            LegendPosition,
            DirectionModeKind,
        )
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
        from ooodev.format.chart2.direct.general.area import GradientStyle, ColorRange
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.kind.zoom_kind import ZoomKind
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("pie_chart.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, ZoomKind.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="pie_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.BRICK, width=1)
                chart_grad = ChartGradient(
                    chart_doc=chart_doc,
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(StandardColor.GREEN_DARK4, StandardColor.TEAL_LIGHT2),
                )
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                legend_pos = ChartLegendPosition(
                    pos=LegendPosition.PAGE_END, no_overlap=True, mode=DirectionModeKind.LR_TB
                )
                Chart2.style_legend(chart_doc=chart_doc, styles=[legend_pos])

                Lo.delay(1_000)
                Lo.close_doc(doc)
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

        from ooodev.format.chart2.direct.legend.position_size import (
            Position as ChartLegendPosition,
            LegendPosition,
            DirectionModeKind,
        )
        # ... other code

        legend_pos = ChartLegendPosition(
            pos=LegendPosition.PAGE_END, no_overlap=True, mode=DirectionModeKind.LR_TB
        )
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_pos])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`14a18c7b-b6ae-4b8a-baac-b1929fca5b2d` and :numref:`0c899f9c-4b39-4607-8553-c3bc4b8ec29f`.


.. cssclass:: screen_shot

    .. _14a18c7b-b6ae-4b8a-baac-b1929fca5b2d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/14a18c7b-b6ae-4b8a-baac-b1929fca5b2d
        :alt: Chart with Legend Set to Bottom
        :figclass: align-center
        :width: 450px

        Chart with Legend Set to Bottom

.. cssclass:: screen_shot

    .. _0c899f9c-4b39-4607-8553-c3bc4b8ec29f:

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
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.legend.position_size.Position`