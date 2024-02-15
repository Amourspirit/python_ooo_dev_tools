.. _help_chart2_format_direct_title_font_effects:

Chart2 Direct Title/Subtitle Font Effects
=========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_font_effect()`` method is used to apply font effects to the Title and Subtitle of a Chart.
as :numref:`efd83001-e9a2-41c4-9c00-7771ec355a1e` Font Effects Dialog, but without the dialog.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 41,42,43,44,45

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooo.dyn.awt.font_underline import FontUnderlineEnum
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor
        from ooodev.utils.data_type.color_range import ColorRange
        from ooodev.format.inner.direct.write.char.font.font_effects import FontLine

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
                    color=StandardColor.PURPLE_DARK1,
                    width=0.7,
                )
                _ = chart_doc.style_area_gradient(
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(
                        StandardColor.BLUE_DARK1,
                        StandardColor.PURPLE_LIGHT2,
                    ),
                )

                title = chart_doc.get_title()
                if title is None:
                    raise ValueError("Title not found")

                title.style_font_effect(
                    color=StandardColor.RED,
                    underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE),
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


Apply the font effects to the data labels
-----------------------------------------

The ``style_font_effect()`` method is used to apply the font effects to the Title and Subtitle.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Apply to Title
""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooo.dyn.awt.font_underline import FontUnderlineEnum
        from ooodev.format.inner.direct.write.char.font.font_effects import FontLine

        # ... other code
        title = chart_doc.get_title()
        if title is None:
            raise ValueError("Title not found")

        title.style_font_effect(
            color=StandardColor.RED,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE),
            shadowed=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`ac3be6e7-4924-45b5-a60f-dfc63c585afc_1` and :numref:`efd83001-e9a2-41c4-9c00-7771ec355a1e_1`.

.. cssclass:: screen_shot

    .. _ac3be6e7-4924-45b5-a60f-dfc63c585afc_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ac3be6e7-4924-45b5-a60f-dfc63c585afc
        :alt: Chart with title font effects applied
        :figclass: align-center
        :width: 520px

        Chart with title font effects applied

    .. _efd83001-e9a2-41c4-9c00-7771ec355a1e_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/efd83001-e9a2-41c4-9c00-7771ec355a1e
        :alt: Chart Title Dialog Font Effects
        :figclass: align-center
        :width: 450px

        Chart Title Dialog Font Effects

Apply to Subtitle
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooo.dyn.awt.font_underline import FontUnderlineEnum
        from ooodev.format.inner.direct.write.char.font.font_effects import FontLine

        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")

        sub_title.style_font_effect(
            color=StandardColor.RED,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE),
            shadowed=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`623c1da6-eafc-4695-a89e-ea0ae3ff994f_1`.

.. cssclass:: screen_shot

    .. _623c1da6-eafc-4695-a89e-ea0ae3ff994f_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/623c1da6-eafc-4695-a89e-ea0ae3ff994f
        :alt: Chart with subtitle font effects applied
        :figclass: align-center
        :width: 520px

        Chart with subtitle font effects applied

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_title_font_only`
        - :ref:`help_chart2_format_direct_title_font`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`