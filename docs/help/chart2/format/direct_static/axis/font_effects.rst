.. _help_chart2_format_direct_static_axis_font_effects:

Chart2 Direct Axis Font Effects (Static)
========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.axis.font.FontEffects` class gives you the similar options for data labels
as :numref:`56218305-bd42-4488-a65d-92883e781496` Font Effects Dialog, but without the dialog.

Methods for formatting the number of an axis are:

    - :py:meth:`Chart2.style_x_axis() <ooodev.office.chart2.Chart2.style_x_axis>`
    - :py:meth:`Chart2.style_x_axis2() <ooodev.office.chart2.Chart2.style_x_axis2>`
    - :py:meth:`Chart2.style_y_axis() <ooodev.office.chart2.Chart2.style_y_axis>`
    - :py:meth:`Chart2.style_y_axis2() <ooodev.office.chart2.Chart2.style_y_axis2>`

.. seealso::

    - :ref:`help_chart2_format_direct_axis_font_effects`

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 35, 36, 37, 38, 39, 40

        import uno
        from ooodev.format.chart2.direct.axis.font import FontEffects as AxisFontEffects
        from ooodev.format.chart2.direct.axis.font import FontUnderlineEnum, FontLine
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
        from ooodev.format.chart2.direct.general.area import GradientStyle, ColorRange, Offset
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("bon_voyage.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="Object 1")

                chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK2, width=0.9)
                chart_grad = ChartGradient(
                    chart_doc=chart_doc,
                    step_count=0,
                    offset=Offset(41, 50),
                    style=GradientStyle.RADIAL,
                    grad_color=ColorRange(StandardColor.TEAL, StandardColor.YELLOW_DARK1),
                )
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                axis_font_effects = AxisFontEffects(
                    color=StandardColor.RED,
                    underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE),
                    shadowed=True,
                )
                Chart2.style_y_axis(chart_doc=chart_doc, styles=[axis_font_effects])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Apply the font effects Axis
---------------------------

Before formatting the chart is seen in :numref:`3adb4ebc-83d9-44c6-9bba-6c92e11f3b0a`.

Style Y-Axis
""""""""""""

.. tabs::

    .. code-tab:: python


        from ooodev.format.chart2.direct.axis.font import FontEffects as AxisFontEffects
        from ooodev.format.chart2.direct.axis.font import FontUnderlineEnum, FontLine
        # ... other code

        axis_font_effects = AxisFontEffects(
            color=StandardColor.RED,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE),
            shadowed=True,
        )
        Chart2.style_y_axis(chart_doc=chart_doc, styles=[axis_font_effects])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`6debc66b-7157-4450-ab72-83ac2524c0af` and :numref:`56218305-bd42-4488-a65d-92883e781496`.

.. cssclass:: screen_shot

    .. _6debc66b-7157-4450-ab72-83ac2524c0af:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/6debc66b-7157-4450-ab72-83ac2524c0af
        :alt: Chart with Y-Axis font effects applied
        :figclass: align-center
        :width: 520px

        Chart with Y-Axis font effects applied

    .. _56218305-bd42-4488-a65d-92883e781496:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/56218305-bd42-4488-a65d-92883e781496
        :alt: Chart Data Labels Dialog Font Effects
        :figclass: align-center
        :width: 450px

        Chart Data Labels Dialog Font Effects

Style Secondary Y-Axis
""""""""""""""""""""""

.. tabs::

    .. code-tab:: python


        # ... other code
        Chart2.style_y_axis2(chart_doc=chart_doc, styles=[axis_font_effects])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`d85fff5e-49c4-48ea-b7cc-7c4c14b72b52`.

.. cssclass:: screen_shot

    .. _d85fff5e-49c4-48ea-b7cc-7c4c14b72b52:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d85fff5e-49c4-48ea-b7cc-7c4c14b72b52
        :alt: Chart with Y-Axis font effects applied
        :figclass: align-center
        :width: 520px

        Chart with Y-Axis font effects applied

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_chart2_format_direct_axis_font_effects`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_axis`
        - :ref:`help_chart2_format_direct_axis_font_only`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_x_axis() <ooodev.office.chart2.Chart2.style_x_axis>`
        - :py:meth:`Chart2.style_x_axis2() <ooodev.office.chart2.Chart2.style_x_axis2>`
        - :py:meth:`Chart2.style_y_axis() <ooodev.office.chart2.Chart2.style_y_axis>`
        - :py:meth:`Chart2.style_y_axis2() <ooodev.office.chart2.Chart2.style_y_axis2>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.axis.font.FontEffects`