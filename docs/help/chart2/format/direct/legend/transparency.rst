.. _help_chart2_format_direct_legend_transparency:

Chart2 Direct Legend Transparency
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Classes in the :py:mod:`ooodev.format.chart2.direct.legend.transparency` module can be used to set the legend transparency.

Calls to the :py:meth:`Chart2.style_legend() <ooodev.office.chart2.Chart2.style_legend>` and method is used to style legend.

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 41,42,43,44

        import uno
        from ooodev.format.chart2.direct.legend.area import Color as LegendAreaColor
        from ooodev.format.chart2.direct.legend.transparency import (
            Transparency as LegendTransparency,
            Gradient as LegendGradient,
            IntensityRange,
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

                legend_color_style = LegendAreaColor(color=StandardColor.GREEN_LIGHT2)
                legend_bg_transparency_style = LegendTransparency(50)
                Chart2.style_legend(
                    chart_doc=chart_doc, styles=[legend_bg_transparency_style, legend_color_style]
                )

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

Before formatting the chart is seen in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

Setting Transparency
^^^^^^^^^^^^^^^^^^^^

The :py:class:`ooodev.format.chart2.direct.legend.transparency.Transparency` class can be used to set the transparency of a chart legend.

The Transparency needs a background color in order to view the transparency. See: :ref:`help_chart2_format_direct_legend_area`.

.. tabs::

    .. code-tab:: python

        ooodev.format.chart2.direct.legend.transparency import Transparency as LegendTransparency

        # ... other code
        legend_bg_transparency_style = LegendTransparency(50)
        Chart2.style_legend(
            chart_doc=chart_doc, styles=[legend_bg_transparency_style, legend_color_style]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`cd864a77-de1d-45d6-b74c-56914b2ffb99` and :numref:`de4b284c-8e3f-4b55-9d61-7e23344e01f5`.

.. cssclass:: screen_shot

    .. _cd864a77-de1d-45d6-b74c-56914b2ffb99:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/cd864a77-de1d-45d6-b74c-56914b2ffb99
        :alt: Chart with transparency applied to legend
        :figclass: align-center
        :width: 450px

        Chart with transparency applied to legend

.. cssclass:: screen_shot

    .. _de4b284c-8e3f-4b55-9d61-7e23344e01f5:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/de4b284c-8e3f-4b55-9d61-7e23344e01f5
        :alt: Chart Legend Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Legend Transparency Dialog

Gradient Transparency
---------------------

Before formatting the chart is seen in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

Setting Gradient
^^^^^^^^^^^^^^^^

The :py:class:`ooodev.format.chart2.direct.legend.transparency.Gradient` class can be used to set the gradient transparency of a legend.

Like the Transparency the Gradient Transparency needs a background color in order to view the transparency. See: :ref:`help_chart2_format_direct_legend_area`.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 2,3,4,5,9,10,11

        from ooodev.format.chart2.direct.legend.area import Color as LegendAreaColor
        from ooodev.format.chart2.direct.legend.transparency import (
            Gradient as LegendGradient,
            IntensityRange,
        )
        # ... other code

        legend_color_style = LegendAreaColor(color=StandardColor.GREEN_LIGHT2)
        legend_transparency_gradient = LegendGradient(
            chart_doc, angle=90, grad_intensity=IntensityRange(0, 100)
        )
        Chart2.style_legend(
            chart_doc=chart_doc, styles=[legend_transparency_gradient, legend_color_style]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The results can bee seen in :numref:`a84c06d4-33b7-4edf-b171-4d9f65cc38ad` and :numref:`37e8d8b9-3aa5-48ac-97ba-880d80489d85`.

.. cssclass:: screen_shot

    .. _a84c06d4-33b7-4edf-b171-4d9f65cc38ad:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a84c06d4-33b7-4edf-b171-4d9f65cc38ad
        :alt: Chart with legend gradient transparency
        :figclass: align-center
        :width: 450px

        Chart with legend gradient transparency

.. cssclass:: screen_shot

    .. _37e8d8b9-3aa5-48ac-97ba-880d80489d85:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/37e8d8b9-3aa5-48ac-97ba-880d80489d85
        :alt: Chart Legend Gradient Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Legend Gradient Transparency Dialog

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_legend_area`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.legend.transparency.Transparency`
        - :py:class:`ooodev.format.chart2.direct.legend.transparency.Gradient`