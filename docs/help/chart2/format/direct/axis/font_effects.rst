.. _help_chart2_format_direct_axis_font_effects:

Chart2 Direct Axis Font Effects
===============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_font_effect()`` method gives similar options for data labels
as :numref:`56218305-bd42-4488-a65d-92883e781496` Font Effects Dialog, but without the dialog.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 39, 40, 41, 42, 43, 44, 45

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooo.dyn.awt.font_underline import FontUnderlineEnum
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.utils.color import StandardColor
        from ooodev.loader.lo import Lo
        from ooodev.utils.data_type.color_range import ColorRange
        from ooodev.utils.data_type.offset import Offset
        from ooodev.format.inner.direct.write.char.font.font_effects import FontLine


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "bon_voyage.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_border_line(
                    color=StandardColor.GREEN_DARK2, width=0.9
                )
                # _ = chart_doc.style_area_color(
                    color=StandardColor.GREEN_LIGHT2
                )
                _ = chart_doc.style_area_gradient(
                    step_count=0,
                    offset=Offset(41, 50),
                    style=GradientStyle.RADIAL,
                    grad_color=ColorRange(
                        StandardColor.TEAL, StandardColor.YELLOW_DARK1
                    ),
                )
                _ = chart_doc.axis_y.style_font_effect(
                    color=StandardColor.RED,
                    underline=FontLine(
                        line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE
                    ),
                    shadowed=True,
                )

                Lo.delay(1_000)
                doc.close()
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


        from ooo.dyn.awt.font_underline import FontUnderlineEnum
        from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
        # ... other code

        _ = chart_doc.axis_y.style_font_effect(
            color=StandardColor.RED,
            underline=FontLine(
                line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE
            ),
            shadowed=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`6debc66b-7157-4450-ab72-83ac2524c0af_1` and :numref:`56218305-bd42-4488-a65d-92883e781496_1`.

.. cssclass:: screen_shot

    .. _6debc66b-7157-4450-ab72-83ac2524c0af_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/6debc66b-7157-4450-ab72-83ac2524c0af
        :alt: Chart with Y-Axis font effects applied
        :figclass: align-center
        :width: 520px

        Chart with Y-Axis font effects applied

    .. _56218305-bd42-4488-a65d-92883e781496_1:

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
        y2_axis = chart_doc.axis2_y
        if y2_axis is not None:
            _ = y2_axis.style_font_effect(
                color=StandardColor.RED,
                underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE),
                shadowed=True,
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`d85fff5e-49c4-48ea-b7cc-7c4c14b72b52_1`.

.. cssclass:: screen_shot

    .. _d85fff5e-49c4-48ea-b7cc-7c4c14b72b52_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d85fff5e-49c4-48ea-b7cc-7c4c14b72b52
        :alt: Chart with Y-Axis font effects applied
        :figclass: align-center
        :width: 520px

        Chart with Y-Axis font effects applied

Getting the font effects style
""""""""""""""""""""""""""""""

For all Axis properties you can get the font effects style using the ``style_font_effect_get()`` method.


.. tabs::

    .. code-tab:: python

        # ... other code
        y2_axis = chart_doc.axis2_y
        if y2_axis is not None:
            f_style = chart_doc.axis_y.style_font_effect_get()
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
        - :ref:`help_chart2_format_direct_axis`
        - :ref:`help_chart2_format_direct_axis_font_only`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.axis.font.FontEffects`