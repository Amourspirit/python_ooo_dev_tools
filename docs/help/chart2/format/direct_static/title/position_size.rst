.. _help_chart2_format_direct_static_title_position_size:

Chart2 Direct Title/Subtitle Position (Static)
==============================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.title.position_size.Position` class is used to set the position of the chart title.

Calls to the :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>` and
:py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>` methods are used to set the Title and Subtitle formatting of a Chart.

.. seealso::

    - :ref:`help_chart2_format_direct_title_position_size`

Setup
-----

General setup used to run the examples in this page.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 34,35

        import uno
        from ooodev.format.chart2.direct.title.position_size import Position as TitlePosition
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

                title_pos = TitlePosition(7.1, 66.3)
                Chart2.style_title(chart_doc=chart_doc, styles=[title_pos])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Applying Position
-----------------

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

By default the :py:class:`~ooodev.format.chart2.direct.title.position_size.Position` class uses millimeters as the unit of measure.

Apply to Title
^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.position_size import Position as TitlePosition
        # ... other code

        title_pos = TitlePosition(7.1, 66.3)
        Chart2.style_title(chart_doc=chart_doc, styles=[title_pos])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`3c13137c-0b86-47b5-9b34-ee52902aff0f` and :numref:`e92ab05a-6093-43ce-a83b-14862827ec35`.

.. cssclass:: screen_shot

    .. _3c13137c-0b86-47b5-9b34-ee52902aff0f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3c13137c-0b86-47b5-9b34-ee52902aff0f
        :alt: Chart with title position set
        :figclass: align-center
        :width: 450px

        Chart with title position set

.. cssclass:: screen_shot

    .. _bfd22d03-f4d8-4d1e-9759-b773051c79df:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bfd22d03-f4d8-4d1e-9759-b773051c79df
        :alt: Chart Title Position and Size Dialog
        :figclass: align-center

        Chart Title Position and Size Dialog

Apply to Subtitle
^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[title_pos])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`3ee5c63f-f82e-4958-9d6c-cde4eaaf3f4f`.

.. cssclass:: screen_shot

    .. _3ee5c63f-f82e-4958-9d6c-cde4eaaf3f4f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3ee5c63f-f82e-4958-9d6c-cde4eaaf3f4f
        :alt: Chart with subtitle position set
        :figclass: align-center
        :width: 450px

        Chart with subtitle position set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_chart2_format_direct_title_position_size`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_title`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>`
        - :py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.title.position_size.Position`
