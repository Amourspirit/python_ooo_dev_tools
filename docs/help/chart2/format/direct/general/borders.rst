.. _help_chart2_format_direct_general_borders:

Chart2 Direct General Borders
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.general.borders.LineProperties` class gives the same options as the Chart Area Borders dialog
as seen in :numref:`236867528-6d4b235a-7239-486b-9c61-63e9a5877e81`.


.. cssclass:: screen_shot

    .. _236867528-6d4b235a-7239-486b-9c61-63e9a5877e81:

    .. figure:: https://user-images.githubusercontent.com/4193389/236867528-6d4b235a-7239-486b-9c61-63e9a5877e81.png
        :alt: Chart Area Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Borders Default Dialog

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 22, 23

        import uno
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
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

                chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=2.2)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_bdr_line])

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

The :py:class:`~ooodev.format.chart2.direct.general.borders.LineProperties` class is used to set the border line properties.

Before setting the border line properties the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

.. tabs::

    .. code-tab:: python

        chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=2.2)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_bdr_line])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`236869888-8057b9ea-cc8a-4e65-bcd5-24a36fd67d8c` and :numref:`236869677-f1d98fb1-4d71-4197-b13b-26e3ef6b35f1`


.. cssclass:: screen_shot

    .. _236869888-8057b9ea-cc8a-4e65-bcd5-24a36fd67d8c:

    .. figure:: https://user-images.githubusercontent.com/4193389/236869888-8057b9ea-cc8a-4e65-bcd5-24a36fd67d8c.png
        :alt: Chart with border set to green
        :figclass: align-center
        :width: 450px

        Chart with border set to green

.. cssclass:: screen_shot

    .. _236869677-f1d98fb1-4d71-4197-b13b-26e3ef6b35f1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236869677-f1d98fb1-4d71-4197-b13b-26e3ef6b35f1.png
        :alt: Chart Area Borders Default Dialog Modified
        :figclass: align-center
        :width: 450px

        Chart Area Borders Default Dialog Modified

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_wall_floor_borders`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.general.borders.LineProperties`