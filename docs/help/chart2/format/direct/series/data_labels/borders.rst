.. _help_chart2_format_direct_series_labels_borders:

Chart2 Direct Series Data Labels Borders
========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_label_border_line()`` method gives the same options as the Chart Data Labels Border dialog
as seen in :numref:`e970bb90-2e58-442f-89c7-bae07efa6237`.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 28,29,30,31

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor
        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

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
                ds.style_label_border_line(
                    color=StandardColor.MAGENTA_DARK1,
                    width=0.75,
                )

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting Line Properties
-----------------------

The ``style_label_border_line()`` method is used to set the data labels border line properties.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
"""""""""""""""""

Setting Style
~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        # ... other code
        ds.style_label_border_line(
            color=StandardColor.MAGENTA_DARK1,
            width=0.75,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9a4c1076-d28b-4d6d-9924-cad9ddf69e6e_1` and :numref:`9dc146b5-8b46-4e6f-8cf1-f3a014827533_1`


.. cssclass:: screen_shot

    .. _9a4c1076-d28b-4d6d-9924-cad9ddf69e6e_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9a4c1076-d28b-4d6d-9924-cad9ddf69e6e
        :alt: Chart with series data labels border set
        :figclass: align-center
        :width: 450px

        Chart with series data labels border set

.. cssclass:: screen_shot

    .. _9dc146b5-8b46-4e6f-8cf1-f3a014827533_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9dc146b5-8b46-4e6f-8cf1-f3a014827533
        :alt: Chart Data Labels Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Labels Borders Default Dialog

Getting Style
~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        # ... other code
        f_style = ds.style_label_border_line_get()
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style Data Point
""""""""""""""""

Setting Style
~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[2]
        dp.style_label_border_line(
            color=StandardColor.MAGENTA_DARK1,
            width=0.75,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`bcd85dc8-5f30-4810-890a-a8ef0ee8c377_1`.

.. cssclass:: screen_shot

    .. _bcd85dc8-5f30-4810-890a-a8ef0ee8c377_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bcd85dc8-5f30-4810-890a-a8ef0ee8c377
        :alt: Chart with point data labels border set
        :figclass: align-center
        :width: 450px

        Chart with point data labels border set

Getting Style
~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        # ... other code
        f_style = dp.style_label_border_line_get()
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_series_series_borders`
        - :py:class:`~ooodev.utils.lo.Lo`