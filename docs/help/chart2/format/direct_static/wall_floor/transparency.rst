.. _help_chart2_format_direct_static_wall_floor_transparency:

Chart2 Direct Wall/Floor Transparency
=====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Classes in the :py:mod:`ooodev.format.chart2.direct.wall.transparency` module can be used to set the chart wall and floor transparency.

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.chart2.direct.wall.transparency import Transparency as WallTransparency
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient, PresetGradientKind
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("col_chart3d.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="col_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.MAGENTA, width=0.7)
                chart_grad = ChartGradient.from_preset(chart_doc, PresetGradientKind.MAHOGANY)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                wall_transparency = WallTransparency(value=40)
                Chart2.style_wall(chart_doc=chart_doc, styles=[wall_transparency])

                floor_transparency = WallTransparency(value=30)
                Chart2.style_floor(chart_doc=chart_doc, styles=[floor_transparency])

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

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.

Setting Transparency
^^^^^^^^^^^^^^^^^^^^

The :py:class:`ooodev.format.chart2.direct.wall.transparency.Transparency` class can be used to set the transparency of a chart wall and floor.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.wall.transparency import Transparency as WallTransparency

        # ... other code
        wall_transparency = WallTransparency(value=40)
        Chart2.style_wall(chart_doc=chart_doc, styles=[wall_transparency])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        floor_transparency = WallTransparency(value=30)
        Chart2.style_floor(chart_doc=chart_doc, styles=[floor_transparency])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`eceee99c-fa50-400f-a55c-343ee6966e6d` and :numref:`c21fd97d-d12c-4779-8cba-c45f49ad03be`.

.. cssclass:: screen_shot

    .. _eceee99c-fa50-400f-a55c-343ee6966e6d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/eceee99c-fa50-400f-a55c-343ee6966e6d
        :alt: Chart with transparency applied to wall and floor
        :figclass: align-center
        :width: 450px

        Chart with transparency applied to wall and floor

.. cssclass:: screen_shot

    .. _c21fd97d-d12c-4779-8cba-c45f49ad03be:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/c21fd97d-d12c-4779-8cba-c45f49ad03be
        :alt: Chart Area Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Transparency Dialog

Gradient Transparency
---------------------

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.

Setting Gradient
^^^^^^^^^^^^^^^^

The :py:class:`ooodev.format.chart2.direct.wall.transparency.Gradient` class can be used to set the gradient transparency of a chart.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.wall.transparency import Gradient as WallGradientTransparency
        from ooodev.format.chart2.direct.wall.transparency import IntensityRange
        from ooodev.utils.data_type.angle import Angle
        # ... other code

        wall_grad_transparent = WallGradientTransparency(
            chart_doc=chart_doc, angle=Angle(30), grad_intensity=IntensityRange(0, 100)
        )
        Chart2.style_wall(chart_doc=chart_doc, styles=[wall_grad_transparent])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to Floor.

.. tabs::

    .. code-tab:: python

        floor_grad_transparent = WallGradientTransparency(
            chart_doc=chart_doc, angle=Angle(120), grad_intensity=IntensityRange(0, 100)
        )
        Chart2.style_floor(chart_doc=chart_doc, styles=[floor_grad_transparent])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`689bf589-8de2-49a0-b260-9f94244aacde` and :numref:`0f8ac32f-e2d2-41c1-b0ad-a3ead8371ee9`.

.. cssclass:: screen_shot

    .. _689bf589-8de2-49a0-b260-9f94244aacde:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/689bf589-8de2-49a0-b260-9f94244aacde
        :alt: Chart with wall and floor gradient transparency
        :figclass: align-center
        :width: 450px

        Chart with wall and floor gradient transparency

.. cssclass:: screen_shot

    .. _0f8ac32f-e2d2-41c1-b0ad-a3ead8371ee9:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/0f8ac32f-e2d2-41c1-b0ad-a3ead8371ee9
        :alt: Chart Wall Gradient Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Wall Gradient Transparency Dialog

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_general_transparency`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_wall() <ooodev.office.chart2.Chart2.style_wall>`
        - :py:meth:`Chart2.style_floor() <ooodev.office.chart2.Chart2.style_floor>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.wall.transparency.Transparency`
        - :py:class:`ooodev.format.chart2.direct.wall.transparency.Gradient`