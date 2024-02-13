.. _help_chart2_format_direct_static_legend_borders:

Chart2 Direct Legend Border
===========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.legend.borders.LineProperties` class gives the same options as the Chart Legend Borders dialog
as seen in :numref:`41bf0361-0952-4c53-adbe-14dae5a2e2f3`.

Calls to the :py:meth:`Chart2.style_legend() <ooodev.office.chart2.Chart2.style_legend>` and method is used to style legend.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 35,36,37,38

        import uno
        from ooodev.format.chart2.direct.legend.borders import LineProperties as LegendLineProperties
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

                legend_line_style = LegendLineProperties(
                    color=StandardColor.BRICK, width=0.8, transparency=20
                )
                Chart2.style_legend(chart_doc=chart_doc, styles=[legend_line_style])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Applying Line Properties
------------------------

The :py:class:`~ooodev.format.chart2.direct.legend.borders.LineProperties` class is used to set the title and subtitle border line properties.

Before formatting the chart is seen in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.legend.borders import LineProperties as LegendLineProperties
        # ... other code

        legend_line_style = LegendLineProperties(color=StandardColor.BRICK, width=0.8, transparency=20)
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_line_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`7286300e-82e5-494f-b7c7-dce2e5cac0f8` and :numref:`41bf0361-0952-4c53-adbe-14dae5a2e2f3`.


.. cssclass:: screen_shot

    .. _7286300e-82e5-494f-b7c7-dce2e5cac0f8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7286300e-82e5-494f-b7c7-dce2e5cac0f8
        :alt: Chart with title border set
        :figclass: align-center
        :width: 450px

        Chart with title border set

.. cssclass:: screen_shot

    .. _41bf0361-0952-4c53-adbe-14dae5a2e2f3:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/41bf0361-0952-4c53-adbe-14dae5a2e2f3
        :alt: Chart Data Series Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Borders Default Dialog

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.legend.borders.LineProperties`
