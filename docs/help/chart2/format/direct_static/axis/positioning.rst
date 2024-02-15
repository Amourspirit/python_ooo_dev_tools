.. _help_chart2_format_direct_static_axis_positioning:

Chart2 Direct Axis Positioning
==============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:mod:`ooodev.format.chart2.direct.axis.positioning` namespace gives you the similar options for axis positioning
as :numref:`9800e630-952c-4947-ba95-f4d5c456284f` Axis Positioning Dialog, but without the dialog.


.. cssclass:: screen_shot

    .. _9800e630-952c-4947-ba95-f4d5c456284f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9800e630-952c-4947-ba95-f4d5c456284f
        :alt: Chart Axis Positioning Dialog
        :figclass: align-center
        :width: 520px

        Chart Axis Positioning Dialog


Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.chart2.direct.axis.positioning import AxisLine, ChartAxisPosition
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
        from ooodev.format.chart2.direct.general.area import GradientStyle, ColorRange, Offset
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.kind.zoom_kind import ZoomKind
        from ooodev.loader.lo import Lo


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("bon_voyage.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, ZoomKind.ZOOM_100_PERCENT)

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

                axis_line_style = AxisLine(cross=ChartAxisPosition.VALUE, value=2000)
                Chart2.style_x_axis(chart_doc=chart_doc, styles=[axis_line_style])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Positioning
-----------

Apply Axis Line
^^^^^^^^^^^^^^^

Before formatting the chart is seen in :numref:`3adb4ebc-83d9-44c6-9bba-6c92e11f3b0a`.

In this example the axis line is positioned at the value ``2000`` and applied to the x-axis.
The axis position is set using the :py:class:`ooodev.format.chart2.direct.axis.positioning.AxisLine` class.

.. tabs::

    .. code-tab:: python


        from ooodev.format.chart2.direct.axis.positioning import AxisLine, ChartAxisPosition
        # ... other code

        axis_line_style = AxisLine(cross=ChartAxisPosition.VALUE, value=2000)
        Chart2.style_x_axis(chart_doc=chart_doc, styles=[axis_line_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The result of running the above can be seen in :numref:`8a888665-d494-402a-9301-4a045a5233b9` and  :numref:`28b188cb-e601-4a0f-99c4-45255e78f92a`.

.. cssclass:: screen_shot

    .. _8a888665-d494-402a-9301-4a045a5233b9:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/8a888665-d494-402a-9301-4a045a5233b9
        :alt: Chart X-Axis Positioning with Axis Line set to value of 2000
        :figclass: align-center
        :width: 520px

        Chart X-Axis Positioning with Axis Line set to value of 2000

.. cssclass:: screen_shot

    .. _28b188cb-e601-4a0f-99c4-45255e78f92a:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/28b188cb-e601-4a0f-99c4-45255e78f92a
        :alt: Chart X-Axis Positioning Dialog with Axis Line set
        :figclass: align-center
        :width: 520px

        Chart X-Axis Positioning Dialog with Axis Line set

Apply Axis Position
^^^^^^^^^^^^^^^^^^^

Before formatting the chart is seen in :numref:`3adb4ebc-83d9-44c6-9bba-6c92e11f3b0a`.

For x-axis Position Dialog the Axis position can be set using the :py:class:`ooodev.format.chart2.direct.axis.positioning.PositionAxis` class.

.. tabs::

    .. code-tab:: python


        from ooodev.format.chart2.direct.axis.positioning import PositionAxis
        # ... other code

        position_axis_style = PositionAxis(False)
        Chart2.style_x_axis(chart_doc=chart_doc, styles=[position_axis_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The result of running the above can be seen in :numref:`baaab89e-eb06-4436-848d-5bbb19b3b906`.

.. cssclass:: screen_shot

    .. _baaab89e-eb06-4436-848d-5bbb19b3b906:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/baaab89e-eb06-4436-848d-5bbb19b3b906
        :alt: Chart X-Axis Positioning Dialog with Axis Line set
        :figclass: align-center
        :width: 520px

        Chart X-Axis Positioning Dialog with Axis Line set

Apply Positioning Labels
^^^^^^^^^^^^^^^^^^^^^^^^

Before formatting the chart is seen in :numref:`3adb4ebc-83d9-44c6-9bba-6c92e11f3b0a`.

The Label position can be set using the :py:class:`ooodev.format.chart2.direct.axis.positioning.LabelPosition` class.

.. tabs::

    .. code-tab:: python


        from ooodev.format.chart2.direct.axis.positioning import LabelPosition, ChartAxisLabelPosition
        # ... other code

        label_position_style = LabelPosition(ChartAxisLabelPosition.NEAR_AXIS_OTHER_SIDE)
        Chart2.style_y_axis(chart_doc=chart_doc, styles=[label_position_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The result of running the above can be seen in :numref:`486ad4fd-c710-4d42-a512-ea0084ea232b` and :numref:`500f2097-72bd-48e1-b21d-dec6a14f722f`.

.. cssclass:: screen_shot

    .. _486ad4fd-c710-4d42-a512-ea0084ea232b:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/486ad4fd-c710-4d42-a512-ea0084ea232b
        :alt: Chart with Y-Axis Label set other side
        :figclass: align-center
        :width: 520px

        Chart with Y-Axis Label set other side

.. cssclass:: screen_shot

    .. _500f2097-72bd-48e1-b21d-dec6a14f722f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/500f2097-72bd-48e1-b21d-dec6a14f722f
        :alt: Chart Y-Axis Positioning Dialog with Labels set
        :figclass: align-center
        :width: 520px

        Chart Y-Axis Positioning Dialog with Labels set

Apply Interval Marks
^^^^^^^^^^^^^^^^^^^^

Interval marks can be set using the :py:class:`ooodev.format.chart2.direct.axis.positioning.IntervalMarks` class.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 1,7,8,9

        from ooodev.format.chart2.direct.axis.positioning import IntervalMarks
        from ooodev.format.chart2.direct.axis.positioning import MarkKind, ChartAxisMarkPosition
        from ooodev.format.chart2.direct.axis.positioning import LabelPosition, ChartAxisLabelPosition
        # ... other code

        label_position_style = LabelPosition(ChartAxisLabelPosition.NEAR_AXIS_OTHER_SIDE)
        interval_marks_style = IntervalMarks(
            major=MarkKind.OUTSIDE, minor=MarkKind.NONE, pos=ChartAxisMarkPosition.AT_LABELS_AND_AXIS
        )
        Chart2.style_y_axis(chart_doc=chart_doc, styles=[label_position_style, interval_marks_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The result of running the above can be seen in :numref:`5df9a764-17be-4714-8541-2d672a076845`.

.. cssclass:: screen_shot

    .. _5df9a764-17be-4714-8541-2d672a076845:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/5df9a764-17be-4714-8541-2d672a076845
        :alt: Chart Y-Axis Positioning Dialog with Labels set
        :figclass: align-center
        :width: 520px

        Chart Y-Axis Positioning Dialog with Labels set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_axis`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_x_axis() <ooodev.office.chart2.Chart2.style_x_axis>`
        - :py:meth:`Chart2.style_x_axis2() <ooodev.office.chart2.Chart2.style_x_axis2>`
        - :py:meth:`Chart2.style_y_axis() <ooodev.office.chart2.Chart2.style_y_axis>`
        - :py:meth:`Chart2.style_y_axis2() <ooodev.office.chart2.Chart2.style_y_axis2>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.axis.positioning.AxisLine`
        - :py:class:`ooodev.format.chart2.direct.axis.positioning.PositionAxis`
        - :py:class:`ooodev.format.chart2.direct.axis.positioning.LabelPosition`
        - :py:class:`ooodev.format.chart2.direct.axis.positioning.IntervalMarks`