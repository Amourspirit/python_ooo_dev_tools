.. _help_chart2_format_direct_legend_font_effects:

Chart2 Direct Legend Font Effects
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The Legend parts of a Chart can be styled using the various ``style_*`` methods of the :py:class`~ooodev.calc.chart2.chart_legend.ChartLegend` class.

Here we will see how to set options that are seen in :numref:`dc7dfc17-6b99-4d07-86a4-b9539f3d61f9`.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 39,40,41,42,43,44,45

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.font_underline import FontUnderlineEnum
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
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
                _ = legend.style_font_effect(
                    color=StandardColor.PURPLE,
                    underline=FontLine(
                        line=FontUnderlineEnum.BOLDWAVE,
                        color=StandardColor.GREEN_DARK2,
                    ),
                )

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

        from ooo.dyn.awt.font_underline import FontUnderlineEnum
        from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
        # ... other code

        _ = legend.style_font_effect(
            color=StandardColor.PURPLE,
            underline=FontLine(
                line=FontUnderlineEnum.BOLDWAVE,
                color=StandardColor.GREEN_DARK2,
            ),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`c53a62d3-75dd-456f-ae95-8a62f1160feb_1` and :numref:`dc7dfc17-6b99-4d07-86a4-b9539f3d61f9_1`.

.. cssclass:: screen_shot

    .. _c53a62d3-75dd-456f-ae95-8a62f1160feb_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/c53a62d3-75dd-456f-ae95-8a62f1160feb
        :alt: Chart with Legend font effects applied
        :figclass: align-center
        :width: 520px

        Chart with Legend font effects applied

    .. _dc7dfc17-6b99-4d07-86a4-b9539f3d61f9_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/dc7dfc17-6b99-4d07-86a4-b9539f3d61f9
        :alt: Chart Legend Font Effects Dialog
        :figclass: align-center
        :width: 450px

        Chart Legend Font Effects Dialog

Get the font effects for the Legend
-----------------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = legend.style_font_effect_get()
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
        - :ref:`help_chart2_format_direct_legend_font`
        - :ref:`help_chart2_format_direct_legend_font_only`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.legend.font.FontEffects`