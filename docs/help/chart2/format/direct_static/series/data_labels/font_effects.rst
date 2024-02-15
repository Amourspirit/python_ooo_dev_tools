.. _help_chart2_format_direct_static_series_labels_font_effects:

Chart2 Direct Series Data Labels Font Effects (Static)
======================================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.series.data_labels.font.FontEffects` class gives you the similar options for data labels
as :numref:`c6c13eb7-cf41-4aab-a88e-1bf6ac0b77b5` Font Effects Dialog, but without the dialog.

Calls to the :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
and :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>` methods are used to set the data labels font effects of a Chart.

.. seealso::

    - :ref:`help_chart2_format_direct_series_labels_font_effects`

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 29,30,31,32,33,34

        import uno
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient, PresetGradientKind
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.series.data_labels.font import FontEffects as LblFontEffects
        from ooodev.format.chart2.direct.series.data_labels.font import FontLine, FontUnderlineEnum
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import CommonColor
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc(Path.cwd() / "tmp" / "col_chart.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="col_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.BLUE_LIGHT3, width=0.7)
                chart_grad = ChartGradient.from_preset(chart_doc, PresetGradientKind.TEAL_BLUE)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                data_lbl_font = LblFontEffects(
                    color=CommonColor.RED,
                    underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE),
                    shadowed=True,
                )
                Chart2.style_data_series(chart_doc=chart_doc, styles=[data_lbl_font])

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

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        data_lbl_font = LblFontEffects(
            color=CommonColor.RED,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE),
            shadowed=True,
        )
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_lbl_font])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`a609d760-cf92-44b3-aafd-31a5b8a79759` and :numref:`c6c13eb7-cf41-4aab-a88e-1bf6ac0b77b5`.

.. cssclass:: screen_shot

    .. _a609d760-cf92-44b3-aafd-31a5b8a79759:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a609d760-cf92-44b3-aafd-31a5b8a79759
        :alt: Chart with data series labels with font effects applied
        :figclass: align-center
        :width: 520px

        Chart with data series labels with font effects applied

    .. _c6c13eb7-cf41-4aab-a88e-1bf6ac0b77b5:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/c6c13eb7-cf41-4aab-a88e-1bf6ac0b77b5
        :alt: Chart Data Labels Dialog Font Effects
        :figclass: align-center
        :width: 450px

        Chart Data Labels Dialog Font Effects

Style Data Point
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_data_point(
            chart_doc=chart_doc, series_idx=0, idx=4, styles=[data_lbl_font]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`d828b120-fd13-4d10-8c12-cd3f4970d0e0`.

.. cssclass:: screen_shot

    .. _d828b120-fd13-4d10-8c12-cd3f4970d0e0:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d828b120-fd13-4d10-8c12-cd3f4970d0e0
        :alt: Chart with data point label with font effects applied
        :figclass: align-center
        :width: 520px

        Chart with data point label with font effects applied

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_chart2_format_direct_series_labels_font_effects`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_series_labels_font_only`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
        - :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.series.data_labels.font.FontOnly`