.. _help_chart2_format_direct_static_axis_font_only:

Chart2 Direct Axis Font Only (Static)
=====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.axis.font.FontOnly` class gives you similar options
as :numref:`618f07b1-23a8-46b9-9aea-0ef4243bb0e2` Font Dialog, but without the dialog.

Methods for formatting the number of an axis are:

    - :py:meth:`Chart2.style_x_axis() <ooodev.office.chart2.Chart2.style_x_axis>`
    - :py:meth:`Chart2.style_x_axis2() <ooodev.office.chart2.Chart2.style_x_axis2>`
    - :py:meth:`Chart2.style_y_axis() <ooodev.office.chart2.Chart2.style_y_axis>`
    - :py:meth:`Chart2.style_y_axis2() <ooodev.office.chart2.Chart2.style_y_axis2>`

.. seealso::

    - :ref:`help_chart2_format_direct_axis_font_only`

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 34,35

        import uno
        from ooodev.format.chart2.direct.axis.font import FontOnly as AxisFontOnly
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
                doc = Calc.open_doc(Path.cwd() / "tmp" / "bon_voyage.ods")
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

                axis_font = AxisFontOnly(name="Lucida Calligraphy", size=14, font_style="italic")
                Chart2.style_y_axis(chart_doc=chart_doc, styles=[axis_font])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Apply Font Only
---------------

Before formatting the chart is seen in :numref:`3adb4ebc-83d9-44c6-9bba-6c92e11f3b0a`.

Apply to Y-Axis
"""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.axis.font import FontOnly as AxisFontOnly
        # ... other code

        axis_font = AxisFontOnly(name="Lucida Calligraphy", size=14, font_style="italic")
        Chart2.style_y_axis(chart_doc=chart_doc, styles=[axis_font])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`4f6b0e7e-c772-4f57-9cf7-8971dc88c2a3` and :numref:`618f07b1-23a8-46b9-9aea-0ef4243bb0e2`.

.. cssclass:: screen_shot

    .. _4f6b0e7e-c772-4f57-9cf7-8971dc88c2a3:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4f6b0e7e-c772-4f57-9cf7-8971dc88c2a3
        :alt: Chart with Y-Axis Font set
        :figclass: align-center
        :width: 450px

        Chart with Y-Axis Font set


.. cssclass:: screen_shot

    .. _618f07b1-23a8-46b9-9aea-0ef4243bb0e2:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/618f07b1-23a8-46b9-9aea-0ef4243bb0e2
        :alt: Chart Y-Axis Dialog Font
        :figclass: align-center
        :width: 450px

        Chart Y-Axis Dialog Font

Apply to X-Axis
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_X_axis(chart_doc=chart_doc, styles=[axis_font])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`1090675b-1837-412e-b430-c0519a460c18`.

.. cssclass:: screen_shot

    .. _1090675b-1837-412e-b430-c0519a460c18:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1090675b-1837-412e-b430-c0519a460c18
        :alt: Chart with Y-Axis Font set
        :figclass: align-center
        :width: 450px

        Chart with Y-Axis Font set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_chart2_format_direct_axis_font_only`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_axis`
        - :ref:`help_chart2_format_direct_axis_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_x_axis() <ooodev.office.chart2.Chart2.style_x_axis>`
        - :py:meth:`Chart2.style_x_axis2() <ooodev.office.chart2.Chart2.style_x_axis2>`
        - :py:meth:`Chart2.style_y_axis() <ooodev.office.chart2.Chart2.style_y_axis>`
        - :py:meth:`Chart2.style_y_axis2() <ooodev.office.chart2.Chart2.style_y_axis2>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.axis.font.FontOnly`