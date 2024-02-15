.. _help_chart2_format_direct_static_series_labels_borders:

Chart2 Direct Series Data Labels Borders
========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.series.data_labels.borders.LineProperties` class gives the same options as the Chart Data Labels Border dialog
as seen in :numref:`e970bb90-2e58-442f-89c7-bae07efa6237`.

Calls to the :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
and :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>` methods are used to set the data labels borders of a Chart.

.. cssclass:: screen_shot

    .. _e970bb90-2e58-442f-89c7-bae07efa6237:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/e970bb90-2e58-442f-89c7-bae07efa6237
        :alt: Chart Data Labels Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Labels Borders Default Dialog

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 27, 28

        import uno
        from ooodev.format.chart2.direct.series.data_labels.borders import LineProperties as LblLineProperties
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient, PresetGradientKind
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

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

                data_lbl_border = LblLineProperties(color=StandardColor.MAGENTA_DARK1, width=0.75)
                Chart2.style_data_series(chart_doc=chart_doc, styles=[data_lbl_border])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting Line Properties
-----------------------

The :py:class:`~ooodev.format.chart2.direct.series.data_labels.borders.LineProperties` class is used to set the data labels border line properties.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        data_lbl_border = LblLineProperties(color=StandardColor.MAGENTA_DARK1, width=0.75)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_lbl_border])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9a4c1076-d28b-4d6d-9924-cad9ddf69e6e` and :numref:`9dc146b5-8b46-4e6f-8cf1-f3a014827533`


.. cssclass:: screen_shot

    .. _9a4c1076-d28b-4d6d-9924-cad9ddf69e6e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9a4c1076-d28b-4d6d-9924-cad9ddf69e6e
        :alt: Chart with series data labels border set
        :figclass: align-center
        :width: 450px

        Chart with series data labels border set

.. cssclass:: screen_shot

    .. _9dc146b5-8b46-4e6f-8cf1-f3a014827533:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9dc146b5-8b46-4e6f-8cf1-f3a014827533
        :alt: Chart Data Labels Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Labels Borders Default Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=2, styles=[data_lbl_border])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`bcd85dc8-5f30-4810-890a-a8ef0ee8c377`.

.. cssclass:: screen_shot

    .. _bcd85dc8-5f30-4810-890a-a8ef0ee8c377:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bcd85dc8-5f30-4810-890a-a8ef0ee8c377
        :alt: Chart with point data labels border set
        :figclass: align-center
        :width: 450px

        Chart with point data labels border set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_series_series_borders`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
        - :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.series.data_labels.borders.LineProperties`