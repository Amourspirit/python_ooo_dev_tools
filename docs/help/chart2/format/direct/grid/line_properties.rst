.. _help_chart2_format_direct_grid_line_properties:

Chart2 Direct Grid Line Properties
==================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_gird_line()`` method can be used to set the line properties of a chart grid.

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 25,26,27,28,29

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooodev.format.inner.preset.preset_border_line import BorderLineKind
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.utils.color import StandardColor
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "col_chart.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_border_line(
                    color=StandardColor.BLUE_DARK2,
                    width=0.7,
                )
                _ = chart_doc.style_area_color(color=StandardColor.DEFAULT_BLUE)
                chart_doc.axis_y.style_gird_line(
                    style=BorderLineKind.CONTINUOUS,
                    color=StandardColor.RED,
                    width=0.5,
                )

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting Grid Style
------------------

Before setting chart formatting is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

The formatting is applied to the y-axis of the grid with a call to ``style_gird_line()``.
The :py:class:`~ooodev.format.inner.preset.preset_border_line.BorderLineKind` enum is used to select the line style.
The :py:class:`~ooodev.format.inner.preset.preset_color.StandardColor` enum is used to select the line color.
The line width is set to ``0.5`` millimeters.

.. tabs::

    .. code-tab:: python
    
        from ooodev.format.inner.preset.preset_border_line import BorderLineKind

        # ... other code
        chart_doc.axis_y.style_gird_line(
            style=BorderLineKind.CONTINUOUS,
            color=StandardColor.RED,
            width=0.5,
        )


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`236972816-9fe26f3f-2cc1-499b-9411-0fe8e8334140_1` and :numref:`236972933-34a2d2f1-4f10-499e-a598-ced11bef0d5a_1`


.. cssclass:: screen_shot

    .. _236972816-9fe26f3f-2cc1-499b-9411-0fe8e8334140_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236972816-9fe26f3f-2cc1-499b-9411-0fe8e8334140.png
        :alt: Chart with border set to green
        :figclass: align-center
        :width: 450px

        Chart with border set to green

.. cssclass:: screen_shot

    .. _236972933-34a2d2f1-4f10-499e-a598-ced11bef0d5a_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236972933-34a2d2f1-4f10-499e-a598-ced11bef0d5a.png
        :alt: Chart Area Borders Default Dialog Modified
        :figclass: align-center
        :width: 450px

        Chart Area Borders Default Dialog Modified

Getting Grid Style
------------------

You can get the font style using the ``style_gird_line_get()`` method.


.. tabs::

    .. code-tab:: python

        # ... other code
        f_style = chart_doc.axis_y.style_gird_line_get()
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_general_area`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`