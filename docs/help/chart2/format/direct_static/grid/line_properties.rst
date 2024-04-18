.. _help_chart2_format_direct_static_grid_line_properties:

Chart2 Direct Grid Line Properties (Static)
===========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.grid.LineProperties` class can be used to set the line properties of a chart grid.

.. seealso::

    - :ref:`help_chart2_format_direct_grid_line_properties`

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 28,29,30,31

        import uno
        from ooodev.format.chart2.direct.general.area import Color as ChartColor
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.grid import BorderLineKind
        from ooodev.format.chart2.direct.grid import LineProperties as GridLineProperties
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2, AxisKind
        from ooodev.utils.color import StandardColor
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("col_chart.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="col_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.BLUE_DARK2, width=0.7)
                chart_color = ChartColor(color=StandardColor.DEFAULT_BLUE)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_color, chart_bdr_line])

                grid_style = GridLineProperties(
                    style=BorderLineKind.CONTINUOUS, color=StandardColor.RED, width=0.5
                )
                Chart2.style_grid(chart_doc=chart_doc, axis_val=AxisKind.Y, styles=[grid_style])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting Grid Line Properties
----------------------------

Before setting chart formatting is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

The formatting is applied to the y-axis (``axis_val=AxisKind.Y``) of the grid with a call to :py:meth:`Chart2.set_grid_lines() <ooodev.office.chart2.Chart2.set_grid_lines>`.
The :py:class:`~ooodev.format.inner.preset.preset_border_line.BorderLineKind` enum is used to select the line style.
The :py:class:`~ooodev.format.inner.preset.preset_color.StandardColor` enum is used to select the line color.
The line width is set to ``0.5`` millimeters.

.. tabs::

    .. code-tab:: python

        grid_style = GridLineProperties(
            style=BorderLineKind.CONTINUOUS, color=StandardColor.RED, width=0.5
        )
        Chart2.style_grid(chart_doc=chart_doc, axis_val=AxisKind.Y, styles=[grid_style])


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`236972816-9fe26f3f-2cc1-499b-9411-0fe8e8334140` and :numref:`236972933-34a2d2f1-4f10-499e-a598-ced11bef0d5a`


.. cssclass:: screen_shot

    .. _236972816-9fe26f3f-2cc1-499b-9411-0fe8e8334140:

    .. figure:: https://user-images.githubusercontent.com/4193389/236972816-9fe26f3f-2cc1-499b-9411-0fe8e8334140.png
        :alt: Chart with border set to green
        :figclass: align-center
        :width: 450px

        Chart with border set to green

.. cssclass:: screen_shot

    .. _236972933-34a2d2f1-4f10-499e-a598-ced11bef0d5a:

    .. figure:: https://user-images.githubusercontent.com/4193389/236972933-34a2d2f1-4f10-499e-a598-ced11bef0d5a.png
        :alt: Chart Area Borders Default Dialog Modified
        :figclass: align-center
        :width: 450px

        Chart Area Borders Default Dialog Modified

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_chart2_format_direct_grid_line_properties`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_general_area`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.set_grid_lines() <ooodev.office.chart2.Chart2.set_grid_lines>`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`