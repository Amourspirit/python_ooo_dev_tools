.. _help_chart2_format_direct_series_labels_font_only:

Chart2 Direct Series Data Labels Font Only
==========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.series.data_labels.font.FontOnly` class gives you similar options for data labels
as :numref:`f4bbd523-c10f-483c-a9c8-3d370dd19433` Font Dialog, but without the dialog.

Calls to the :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
and :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>` methods are used to set the data labels font of a Chart.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 29

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind
        from ooodev.utils.color import StandardColor

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
                ds.style_font(name="Lucida Calligraphy", size=14, font_style="italic")

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Apply the font to Data Labels
-----------------------------

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        ds = chart_doc.get_data_series()[0]
        ds.style_font(name="Lucida Calligraphy", size=14, font_style="italic")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`f4bbd523-c10f-483c-a9c8-3d370dd19433_1` and :numref:`2641c2d6-6efb-4c59-a747-13f7e0c3ed5c_1`.

.. cssclass:: screen_shot

    .. _f4bbd523-c10f-483c-a9c8-3d370dd19433_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f4bbd523-c10f-483c-a9c8-3d370dd19433
        :alt: Chart with Data Series Labels Font set
        :figclass: align-center
        :width: 450px

        Chart with Data Series Labels Font set


.. cssclass:: screen_shot

    .. _2641c2d6-6efb-4c59-a747-13f7e0c3ed5c_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2641c2d6-6efb-4c59-a747-13f7e0c3ed5c
        :alt: Chart Data Labels Dialog Font
        :figclass: align-center
        :width: 450px

        Chart Data Labels Dialog Font

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[0]
        dp.style_font(name="Lucida Calligraphy", size=14, font_style="italic")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`93bf56fc-d122-4fa0-8630-a3a2ae87ef80_1`.

.. cssclass:: screen_shot

    .. _93bf56fc-d122-4fa0-8630-a3a2ae87ef80_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/93bf56fc-d122-4fa0-8630-a3a2ae87ef80
        :alt: Chart with Data Point Label Font set
        :figclass: align-center
        :width: 450px

        Chart with Data Point Label Font set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_series_labels_font_effects`
        - :py:class:`~ooodev.loader.Lo`