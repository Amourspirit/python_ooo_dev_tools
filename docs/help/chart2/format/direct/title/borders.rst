.. _help_chart2_format_direct_title_borders:

Chart2 Direct Title/Subtitle Borders
====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.title.borders.LineProperties` class gives the same options as the Chart Data Series Borders dialog
as seen in :numref:`a31ee22f-14cc-43ef-844f-7a078ec1abd9`.

Calls to the :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>` and
:py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>` methods are used to set the Title and Subtitle borders of a Chart.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 34, 35

        import uno
        from ooodev.format.chart2.direct.title.borders import LineProperties as TitleLineProperties
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

                title_border = TitleLineProperties(color=StandardColor.MAGENTA_DARK1, width=0.75)
                Chart2.style_title(chart_doc=chart_doc, styles=[title_border])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Applying Line Properties
------------------------

The :py:class:`~ooodev.format.chart2.direct.title.borders.LineProperties` class is used to set the title and subtitle border line properties.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Apply to Title
""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.borders import LineProperties as TitleLineProperties
        # ... other code

        title_border = TitleLineProperties(color=StandardColor.MAGENTA_DARK1, width=0.75)
        Chart2.style_title(chart_doc=chart_doc, styles=[title_border])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9b8faf7e-9cfa-407d-880c-1efce5b012fe` and :numref:`a31ee22f-14cc-43ef-844f-7a078ec1abd9`.


.. cssclass:: screen_shot

    .. _9b8faf7e-9cfa-407d-880c-1efce5b012fe:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9b8faf7e-9cfa-407d-880c-1efce5b012fe
        :alt: Chart with title border set
        :figclass: align-center
        :width: 450px

        Chart with title border set

.. cssclass:: screen_shot

    .. _a31ee22f-14cc-43ef-844f-7a078ec1abd9:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a31ee22f-14cc-43ef-844f-7a078ec1abd9
        :alt: Chart Data Series Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Borders Default Dialog

Apply to Subtitle
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[title_border])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`27378b9f-41c0-4975-8b14-161133e81ca0`.


.. cssclass:: screen_shot

    .. _27378b9f-41c0-4975-8b14-161133e81ca0:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/27378b9f-41c0-4975-8b14-161133e81ca0
        :alt: Chart with subtitle border set
        :figclass: align-center
        :width: 450px

        Chart with subtitle border set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>`
        - :py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.title.borders.LineProperties`
