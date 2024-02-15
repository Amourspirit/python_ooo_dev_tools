.. _help_chart2_format_direct_static_title_font_effects:

Chart2 Direct Title/Subtitle Font Effects
=========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.title.font.FontEffects` class gives you the similar options
as :numref:`efd83001-e9a2-41c4-9c00-7771ec355a1e` Font Effects Dialog, but without the dialog.

Calls to the :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>` and
:py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>` methods are used to set the Title and Subtitle font effects of a Chart.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 35,36,37,38,39,40

        import uno
        from ooodev.format.chart2.direct.title.font import FontEffects as TitleFontEffects
        from ooodev.format.chart2.direct.title.font import FontUnderlineEnum, FontLine
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
        from ooodev.format.chart2.direct.general.area import GradientStyle, ColorRange
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("pie_flat_chart.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="pie_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.PURPLE_DARK1, width=0.7)
                chart_grad = ChartGradient(
                    chart_doc=chart_doc,
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(StandardColor.BLUE_DARK1, StandardColor.PURPLE_LIGHT2),
                )
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                title_font_effect = TitleFontEffects(
                    color=StandardColor.RED,
                    underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE),
                    shadowed=True,
                )
                Chart2.style_title(chart_doc=chart_doc, styles=[title_font_effect])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Apply the font effects to the data labels
-----------------------------------------

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Apply to Title
""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.font import FontEffects as TitleFontEffects
        from ooodev.format.chart2.direct.title.font import FontUnderlineEnum, FontLine
        # ... other code

        title_font_effect = TitleFontEffects(
            color=StandardColor.RED,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE),
            shadowed=True,
        )
        Chart2.style_title(chart_doc=chart_doc, styles=[title_font_effect])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`ac3be6e7-4924-45b5-a60f-dfc63c585afc` and :numref:`efd83001-e9a2-41c4-9c00-7771ec355a1e`.

.. cssclass:: screen_shot

    .. _ac3be6e7-4924-45b5-a60f-dfc63c585afc:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ac3be6e7-4924-45b5-a60f-dfc63c585afc
        :alt: Chart with title font effects applied
        :figclass: align-center
        :width: 520px

        Chart with title font effects applied

    .. _efd83001-e9a2-41c4-9c00-7771ec355a1e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/efd83001-e9a2-41c4-9c00-7771ec355a1e
        :alt: Chart Title Dialog Font Effects
        :figclass: align-center
        :width: 450px

        Chart Title Dialog Font Effects

Apply to Subtitle
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[title_font_effect])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`623c1da6-eafc-4695-a89e-ea0ae3ff994f`.

.. cssclass:: screen_shot

    .. _623c1da6-eafc-4695-a89e-ea0ae3ff994f:

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
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>`
        - :py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.title.font.FontEffects`