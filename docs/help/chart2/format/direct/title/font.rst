.. _help_chart2_format_direct_title_font:

Chart2 Direct Title/Subtitle Font
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.title.font.Font` class contains some properties from both :ref:`Font Only <help_chart2_format_direct_title_font_only>` and :ref:`Font Effects <help_chart2_format_direct_title_font_effects>`.
This class is more general purpose.

Calls to the :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>` and
:py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>` methods are used to set the Title and Subtitle font of a Chart.

Because :py:class:`~ooodev.format.chart2.direct.title.font.Font` class is more general purpose, not all properties are guaranteed to work with titles.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 34,35

        import uno
        from ooodev.format.chart2.direct.title.font import Font as TitleFont
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
        from ooodev.format.chart2.direct.general.area import GradientStyle, ColorRange
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc(Path.cwd() / "tmp" / "pie_flat_chart.ods")
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

                title_font_effect = TitleFont(b=True, i=True, u=True, color=StandardColor.PURPLE_DARK2, shadowed=True)
                Chart2.style_title(chart_doc=chart_doc, styles=[title_font_effect])

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

        from ooodev.format.chart2.direct.title.font import Font as TitleFont

        # ... other code
        title_font_effect = TitleFont(
            b=True, i=True, u=True, color=StandardColor.PURPLE_DARK2, shadowed=True
        )
        Chart2.style_title(chart_doc=chart_doc, styles=[title_font_effect])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`eaa1eab4-687c-466a-a7fd-2c126f7b1e2f`.

.. cssclass:: screen_shot

    .. _eaa1eab4-687c-466a-a7fd-2c126f7b1e2f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/eaa1eab4-687c-466a-a7fd-2c126f7b1e2f
        :alt: Chart with Title Font set
        :figclass: align-center
        :width: 450px

        Chart with Title Font set


Apply to Subtitle
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[title_font_effect])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None



Running the above code will produce the following output shown in :numref:`bb19afad-c492-4f6f-a7bf-89d2323b1c77`.

.. cssclass:: screen_shot

    .. _bb19afad-c492-4f6f-a7bf-89d2323b1c77:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bb19afad-c492-4f6f-a7bf-89d2323b1c77
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
        - :ref:`help_chart2_format_direct_title_font_only`
        - :ref:`help_chart2_format_direct_title_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>`
        - :py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.title.font.Font`