.. _help_chart2_format_direct_legend_font:

Chart2 Direct Legend Font
=========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.legend.font.Font` combines some of the option of :ref:`help_chart2_format_direct_legend_font_only`
and :ref:`help_chart2_format_direct_legend_font_effects` into a single class. Also a few more options are added.

Calls to the :py:meth:`Chart2.style_legend() <ooodev.office.chart2.Chart2.style_legend>` and method is used to style legend.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 35,36

        import uno
        from ooodev.format.chart2.direct.legend.font import Font as LegendFont
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
        from ooodev.format.chart2.direct.general.area import GradientStyle, ColorRange
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.kind.zoom_kind import ZoomKind
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc(Path.cwd() / "tmp" / "pie_chart.ods")
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

                legend_font_style = LegendFont(b=True, color=StandardColor.PURPLE, size=12)
                Chart2.style_legend(chart_doc=chart_doc, styles=[legend_font_style])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Apply the font effects to the Legend
------------------------------------

Before formatting the chart is visible in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.legend.font import Font as LegendFont
        # ... other code

        legend_font_style = LegendFont(b=True, color=StandardColor.PURPLE, size=12)
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_font_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`b120a95d-fa1c-4ef1-89f1-5308464b2962`.

.. cssclass:: screen_shot

    .. _b120a95d-fa1c-4ef1-89f1-5308464b2962:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/b120a95d-fa1c-4ef1-89f1-5308464b2962
        :alt: Chart with Legend font applied
        :figclass: align-center
        :width: 520px

        Chart with Legend font applied


Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_legend_font_only`
        - :ref:`help_chart2_format_direct_legend_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.legend.font.Font`
