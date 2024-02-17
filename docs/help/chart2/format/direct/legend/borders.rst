.. _help_chart2_format_direct_legend_borders:

Chart2 Direct Legend Border
===========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The Legend parts of a Chart can be styled using the various ``style_*`` methods of the :py:class:`~ooodev.calc.chart2.chart_legend.ChartLegend` class.

Here we will see how to set options that are seen in :numref:`41bf0361-0952-4c53-adbe-14dae5a2e2f3`.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 37,38,39,40,41

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor
        from ooodev.utils.data_type.color_range import ColorRange

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "piechart.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_border_line(
                    color=StandardColor.BRICK,
                    width=1,
                )
                _ = chart_doc.style_area_gradient(
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(
                        StandardColor.GREEN_DARK4,
                        StandardColor.TEAL_LIGHT2,
                    ),
                )
                legend = chart_doc.first_diagram.get_legend()
                if legend is None:
                    raise ValueError("Legend is None")
                _ = legend.style_border_line(
                    color=StandardColor.BRICK,
                    width=0.8,
                    transparency=20,
                )

                f_style = legend.style_border_line_get()
                assert f_style is not None

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Applying Border Style
---------------------

The ``style_border_line()`` method is used to set the title and subtitle border line properties.

Before formatting the chart is seen in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.legend.borders import LineProperties as LegendLineProperties
        # ... other code

        _ = legend.style_border_line(
            color=StandardColor.BRICK,
            width=0.8,
            transparency=20,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`7286300e-82e5-494f-b7c7-dce2e5cac0f8_1` and :numref:`41bf0361-0952-4c53-adbe-14dae5a2e2f3_1`.


.. cssclass:: screen_shot

    .. _7286300e-82e5-494f-b7c7-dce2e5cac0f8_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7286300e-82e5-494f-b7c7-dce2e5cac0f8
        :alt: Chart with title border set
        :figclass: align-center
        :width: 450px

        Chart with title border set

.. cssclass:: screen_shot

    .. _41bf0361-0952-4c53-adbe-14dae5a2e2f3_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/41bf0361-0952-4c53-adbe-14dae5a2e2f3
        :alt: Chart Data Series Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Borders Default Dialog

Getting Border Style
--------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = legend.style_border_line_get()
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
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.legend.borders.LineProperties`
