.. _help_chart2_format_direct_series_labels_font_effects:

Chart2 Direct Series Data Labels Font Effects
=============================================

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

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 31,32,33,34,35

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.font_underline import FontUnderlineEnum
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind
        from ooodev.utils.color import StandardColor
        from ooodev.format.inner.direct.write.char.font.font_effects import FontLine

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "col_chart.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_border_line(
                    color=StandardColor.BLUE_LIGHT3,
                    width=0.7,
                )
                _ = chart_doc.style_area_gradient_from_preset(
                    preset=PresetGradientKind.TEAL_BLUE,
                )

                ds = chart_doc.get_data_series()[0]
                ds.style_font_effect(
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

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooo.dyn.awt.font_underline import FontUnderlineEnum
        from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
        # ... other code
        ds = chart_doc.get_data_series()[0]
        ds.style_font_effect(
            color=StandardColor.RED,
            underline=FontLine(
                line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE
            ),
            shadowed=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`a609d760-cf92-44b3-aafd-31a5b8a79759_1` and :numref:`c6c13eb7-cf41-4aab-a88e-1bf6ac0b77b5_1`.

.. cssclass:: screen_shot

    .. _a609d760-cf92-44b3-aafd-31a5b8a79759_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a609d760-cf92-44b3-aafd-31a5b8a79759
        :alt: Chart with data series labels with font effects applied
        :figclass: align-center
        :width: 520px

        Chart with data series labels with font effects applied

    .. _c6c13eb7-cf41-4aab-a88e-1bf6ac0b77b5_1:

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
        ds = chart_doc.get_data_series()[0]
        dp = ds[4]
        dp.style_font_effect(
            color=StandardColor.RED,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.BLUE),
            shadowed=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`d828b120-fd13-4d10-8c12-cd3f4970d0e0_1`.

.. cssclass:: screen_shot

    .. _d828b120-fd13-4d10-8c12-cd3f4970d0e0_1:

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
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_series_labels_font_only`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`