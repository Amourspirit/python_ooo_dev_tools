.. _help_chart2_format_direct_legend_font:

Chart2 Direct Legend Font
=========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The Legend parts of a Chart can be styled using the various ``style_*`` methods of the :py:class:`~ooodev.calc.chart2.chart_legend.ChartLegend` class.

The ``style_font_general()`` method combines some of the option of ``style_font()`` and ``style_font_effects()`` into a single class. Also a few more options are added.


Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 37

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
                _ = legend.style_font_general(b=True, color=StandardColor.PURPLE, size=12)

                # f_style = legend.style_font_get()
                # assert f_style is not None

                Lo.delay(1_000)
                doc.close()
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

        # ... other code
        _ = legend.style_font_general(b=True, color=StandardColor.PURPLE, size=12)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`b120a95d-fa1c-4ef1-89f1-5308464b2962_1`.

.. cssclass:: screen_shot

    .. _b120a95d-fa1c-4ef1-89f1-5308464b2962_1:

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
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.legend.font.Font`
