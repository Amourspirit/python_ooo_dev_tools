.. _help_chart2_format_direct_static_axis_line:

Chart2 Direct Axis Line (Static)
================================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.axis.line.LineProperties` class gives the same options
as seen in :numref:`ee039628-896d-4ab2-a4c8-843d643a5579`.

Methods for formatting the number of an axis are:

    - :py:meth:`Chart2.style_x_axis() <ooodev.office.chart2.Chart2.style_x_axis>`
    - :py:meth:`Chart2.style_x_axis2() <ooodev.office.chart2.Chart2.style_x_axis2>`
    - :py:meth:`Chart2.style_y_axis() <ooodev.office.chart2.Chart2.style_y_axis>`
    - :py:meth:`Chart2.style_y_axis2() <ooodev.office.chart2.Chart2.style_y_axis2>`

.. seealso::

    - :ref:`help_chart2_format_direct_axis_line`

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 34,35

        import uno
        from ooodev.format.chart2.direct.axis.line import LineProperties as AxisLineProperties
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
        from ooodev.format.chart2.direct.general.area import GradientStyle, ColorRange, Offset
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.gui import GUI
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

                axis_line_props = AxisLineProperties(color=StandardColor.TEAL, width=0.75)
                Chart2.style_x_axis(chart_doc=chart_doc, styles=[axis_line_props])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply Line to Axis
------------------

The :py:class:`~ooodev.format.chart2.direct.axis.line.LineProperties` class is used to set the Axis line properties.

Before formatting the chart is seen in :numref:`3adb4ebc-83d9-44c6-9bba-6c92e11f3b0a`.

Apply to Y-Axis
"""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.axis.line import LineProperties as AxisLineProperties
        # ... other code

        axis_line_props = AxisLineProperties(color=StandardColor.TEAL, width=0.75)
        Chart2.style_y_axis(chart_doc=chart_doc, styles=[axis_line_props])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`1c1711ce-1169-4106-8925-c7790dbad0e0` and :numref:`ee039628-896d-4ab2-a4c8-843d643a5579`


.. cssclass:: screen_shot

    .. _1c1711ce-1169-4106-8925-c7790dbad0e0:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1c1711ce-1169-4106-8925-c7790dbad0e0
        :alt: Chart with Y-Axis line set
        :figclass: align-center
        :width: 450px

        Chart with Y-Axis line set

.. cssclass:: screen_shot

    .. _ee039628-896d-4ab2-a4c8-843d643a5579:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ee039628-896d-4ab2-a4c8-843d643a5579
        :alt: Chart Y-Axis Line Dialog
        :figclass: align-center
        :width: 450px

        Chart Y-Axis Line Dialog

Apply to X-Axis
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        Chart2.style_x_axis(chart_doc=chart_doc, styles=[axis_line_props])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`ae063a29-1ebc-442e-9f9b-7d9dba8f64ad`


.. cssclass:: screen_shot

    .. _ae063a29-1ebc-442e-9f9b-7d9dba8f64ad:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ae063a29-1ebc-442e-9f9b-7d9dba8f64ad
        :alt: Chart with Y-Axis line set
        :figclass: align-center
        :width: 450px

        Chart with Y-Axis line set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_chart2_format_direct_axis_line`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_axis`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_x_axis() <ooodev.office.chart2.Chart2.style_x_axis>`
        - :py:meth:`Chart2.style_x_axis2() <ooodev.office.chart2.Chart2.style_x_axis2>`
        - :py:meth:`Chart2.style_y_axis() <ooodev.office.chart2.Chart2.style_y_axis>`
        - :py:meth:`Chart2.style_y_axis2() <ooodev.office.chart2.Chart2.style_y_axis2>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.axis.line.LineProperties`