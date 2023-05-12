.. _help_chart2_format_direct_series_labels_font_only:

Chart2 Direct Series Data Labels Font Only
==========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.series.data_labels.font.FontOnly` class gives you the similar options for data labels
as :numref:`f4bbd523-c10f-483c-a9c8-3d370dd19433` Font Dialog, but without the dialog.

Calls to the :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>` method are used to set the data labels font of a Chart.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 27,28

        import uno
        from ooodev.format.chart2.direct.series.data_labels.font import FontOnly as LblFontOnly
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient, PresetGradientKind
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("col_chart.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="col_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.BLUE_LIGHT3, width=0.7)
                chart_grad = ChartGradient.from_preset(chart_doc, PresetGradientKind.TEAL_BLUE)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                data_lbl_font = LblFontOnly(name="Lucida Calligraphy", size=14, font_style="italic")
                Chart2.style_data_series(chart_doc=chart_doc, styles=[data_lbl_font])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Apply the font to Data Labels
-----------------------------

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

.. tabs::

    .. code-tab:: python

        data_lbl_font = LblFontOnly(name="Lucida Calligraphy", size=14, font_style="italic")
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_lbl_font])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`f4bbd523-c10f-483c-a9c8-3d370dd19433` and :numref:`2641c2d6-6efb-4c59-a747-13f7e0c3ed5c`.

.. cssclass:: screen_shot

    .. _f4bbd523-c10f-483c-a9c8-3d370dd19433:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f4bbd523-c10f-483c-a9c8-3d370dd19433
        :alt: Chart with Data Labels Font set
        :figclass: align-center
        :width: 450px

        Chart with Data Labels Font set


.. cssclass:: screen_shot

    .. _2641c2d6-6efb-4c59-a747-13f7e0c3ed5c:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2641c2d6-6efb-4c59-a747-13f7e0c3ed5c
        :alt: Chart Data Labels Dialog Font
        :figclass: align-center
        :width: 450px

        Chart Data Labels Dialog Font

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_series_labels_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.series.data_labels.font.FontOnly`