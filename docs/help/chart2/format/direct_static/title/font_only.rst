.. _help_chart2_format_direct_static_title_font_only:

Chart2 Direct Title/Subtitle Font Only
======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.title.font.FontOnly` class gives you the similar options
as :numref:`0c6b3393-b925-42c1-9bc5-8b604459de9a` Font Dialog, but without the dialog.

Calls to the :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>` and
:py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>` methods are used to set the Title and Subtitle font of a Chart.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 34,35

        import uno
        from ooodev.format.chart2.direct.title.font import FontOnly as TitleFontOnly
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

                title_font = TitleFontOnly(name="Lucida Calligraphy", size=14, font_style="italic")
                Chart2.style_title(chart_doc=chart_doc, styles=[title_font])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Apply the Font
--------------

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Apply to Title
""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.font import FontOnly as TitleFontOnly

        # ... other code
        title_font = TitleFontOnly(name="Lucida Calligraphy", size=14, font_style="italic")
        Chart2.style_title(chart_doc=chart_doc, styles=[title_font])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`0bd83e10-35ea-4ba3-bff9-04548d2ad0e0` and :numref:`0c6b3393-b925-42c1-9bc5-8b604459de9a`.

.. cssclass:: screen_shot

    .. _0bd83e10-35ea-4ba3-bff9-04548d2ad0e0:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/0bd83e10-35ea-4ba3-bff9-04548d2ad0e0
        :alt: Chart with Title Font set
        :figclass: align-center
        :width: 450px

        Chart with Title Font set


.. cssclass:: screen_shot

    .. _0c6b3393-b925-42c1-9bc5-8b604459de9a:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/0c6b3393-b925-42c1-9bc5-8b604459de9a
        :alt: Chart Data Labels Dialog Font
        :figclass: align-center
        :width: 450px

        Chart Data Labels Dialog Font

Apply to Subtitle
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[title_font])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`6427af0a-2fad-4f6a-b390-813c9503eced`.

.. cssclass:: screen_shot

    .. _6427af0a-2fad-4f6a-b390-813c9503eced:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/6427af0a-2fad-4f6a-b390-813c9503eced
        :alt: Chart with Subtitle Font set
        :figclass: align-center
        :width: 450px

        Chart with Subtitle Font set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_title_font_effects`
        - :ref:`help_chart2_format_direct_title_font`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>`
        - :py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.title.font.FontOnly`