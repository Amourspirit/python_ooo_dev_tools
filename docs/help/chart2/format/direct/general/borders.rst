.. _help_chart2_format_direct_general_borders:

Chart2 Direct General Borders
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

As seen in :ref:`help_chart2_format_direct_general_area` the border are set using ``style_border_line()`` method of :ref:`ooodev.calc.chart2.chart_doc.ChartDoc`

The ``style_border_line()`` method gives the same options as the Chart Area Borders dialog
as seen in :numref:`236867528-6d4b235a-7239-486b-9c61-63e9a5877e81_1`.


.. cssclass:: screen_shot

    .. _236867528-6d4b235a-7239-486b-9c61-63e9a5877e81_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236867528-6d4b235a-7239-486b-9c61-63e9a5877e81.png
        :alt: Chart Area Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Borders Default Dialog

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from pathlib import Path
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.utils.color import StandardColor
        from ooodev.loader.lo import Lo


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
                _ = chart_doc.style_border_line(color=StandardColor.GREEN_DARK3, width=2.2)

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

The ``style_area_color()`` method is used to set to set the border line properties.

Before setting the border line properties the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

.. tabs::

    .. code-tab:: python

        _ = chart_doc.style_border_line(color=StandardColor.GREEN_DARK3, width=2.2)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`236869888-8057b9ea-cc8a-4e65-bcd5-24a36fd67d8c_1` and :numref:`236869677-f1d98fb1-4d71-4197-b13b-26e3ef6b35f1_1`


.. cssclass:: screen_shot

    .. _236869888-8057b9ea-cc8a-4e65-bcd5-24a36fd67d8c_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236869888-8057b9ea-cc8a-4e65-bcd5-24a36fd67d8c.png
        :alt: Chart with border set to green
        :figclass: align-center
        :width: 450px

        Chart with border set to green

.. cssclass:: screen_shot

    .. _236869677-f1d98fb1-4d71-4197-b13b-26e3ef6b35f1_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236869677-f1d98fb1-4d71-4197-b13b-26e3ef6b35f1.png
        :alt: Chart Area Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Borders Default Dialog

Getting the border style
------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = chart_doc.style_border_line_get()
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
        - :ref:`help_chart2_format_direct_wall_floor_borders`
        - :ref:`ooodev.calc.chart2.chart_doc.ChartDoc`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.general.borders.LineProperties`