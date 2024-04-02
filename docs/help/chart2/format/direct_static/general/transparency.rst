.. _help_chart2_format_direct_static_general_transparency:

Chart2 Direct General Transparency (Static)
===========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Classes in the :py:mod:`ooodev.format.chart2.direct.general.transparency` module can be used to set the transparency of a chart.

.. seealso::

    - :ref:`help_chart2_format_direct_general_transparency`

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.chart2.direct.general.area import Color as ChartColor
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.transparency import Transparency as ChartTransparency
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.gui import GUI
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

                chart_color = ChartColor(color=StandardColor.GREEN_LIGHT2)
                chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
                chart_transparency = ChartTransparency(value=50)
                Chart2.style_background(
                    chart_doc=chart_doc, styles=[chart_color, chart_bdr_line, chart_transparency]
                )

                f_style = ChartTransparency.from_obj(chart_doc.getPageBackground())
                assert f_style is not None

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Transparency
------------

Before setting the background transparency of the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Setting Transparency
^^^^^^^^^^^^^^^^^^^^

The :py:class:`ooodev.format.chart2.direct.general.transparency.Transparency` class can be used to set the transparency of a chart.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 3,4,5,6

        chart_color = ChartColor(color=StandardColor.GREEN_LIGHT2)
        chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
        chart_transparency = ChartTransparency(value=50)
        Chart2.style_background(
            chart_doc=chart_doc, styles=[chart_color, chart_bdr_line, chart_transparency]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`236953627-38d0f2c5-e19a-402e-942c-d7f1c1a27c11` and :numref:`236953723-96edce28-2476-4abb-af3d-223723c4fd1a`.

.. cssclass:: screen_shot

    .. _236953627-38d0f2c5-e19a-402e-942c-d7f1c1a27c11:

    .. figure:: https://user-images.githubusercontent.com/4193389/236953627-38d0f2c5-e19a-402e-942c-d7f1c1a27c11.png
        :alt: Chart with border, color and  transparency
        :figclass: align-center
        :width: 450px

        Chart with border, color and  transparency

.. cssclass:: screen_shot

    .. _236953723-96edce28-2476-4abb-af3d-223723c4fd1a:

    .. figure:: https://user-images.githubusercontent.com/4193389/236953723-96edce28-2476-4abb-af3d-223723c4fd1a.png
        :alt: Chart Area Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Transparency Dialog

Getting the transparency from a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = ChartTransparency.from_obj(chart_doc.getPageBackground())
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Gradient Transparency
---------------------

Before setting the background gradient transparency of the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Setting Gradient
^^^^^^^^^^^^^^^^

The :py:class:`ooodev.format.chart2.direct.general.transparency.Gradient` class can be used to set the gradient transparency of a chart.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 8,9,10,11,12,13,14

        from ooodev.format.chart2.direct.general.transparency import Gradient as ChartGradientTransparency
        from ooodev.format.chart2.direct.general.transparency import IntensityRange
        from ooodev.utils.data_type.angle import Angle
        # ... other code

        chart_color = ChartColor(color=StandardColor.GREEN_LIGHT2)
        chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
        chart_grad_transparent = ChartGradientTransparency(
            chart_doc=chart_doc, angle=Angle(30), grad_intensity=IntensityRange(0, 100)
        )
        Chart2.style_background(
            chart_doc=chart_doc,
            styles=[chart_color, chart_bdr_line, chart_grad_transparent]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`236955053-0dba0e1e-6bbf-4b22-921b-5e19e2131baf` and :numref:`236955121-cad9d1e7-c86d-435f-920c-02e0bb451c84`.

.. cssclass:: screen_shot

    .. _236955053-0dba0e1e-6bbf-4b22-921b-5e19e2131baf:

    .. figure:: https://user-images.githubusercontent.com/4193389/236955053-0dba0e1e-6bbf-4b22-921b-5e19e2131baf.png
        :alt: Chart with border, color and  transparency
        :figclass: align-center
        :width: 450px

        Chart with border, color and  transparency

.. cssclass:: screen_shot

    .. _236955121-cad9d1e7-c86d-435f-920c-02e0bb451c84:

    .. figure:: https://user-images.githubusercontent.com/4193389/236955121-cad9d1e7-c86d-435f-920c-02e0bb451c84.png
        :alt: Chart Area Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Transparency Dialog

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_chart2_format_direct_general_transparency`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_wall_floor_transparency`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.general.transparency.Transparency`
        - :py:class:`ooodev.format.chart2.direct.general.transparency.Gradient`
