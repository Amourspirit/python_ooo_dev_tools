.. _help_chart2_format_direct_wall_floor_transparency:

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
        :emphasize-lines: 30

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor
        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "col_chart3d.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_border_line(
                    color=StandardColor.MAGENTA,
                    width=0.7,
                )

                _ = chart_doc.style_area_gradient_from_preset(
                    preset=PresetGradientKind.MAHOGANY,
                )

                wall = chart_doc.first_diagram.wall
                wall.style_area_transparency_transparency(30)

                Lo.delay(1_000)
                doc.close()
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

The ``style_area_transparency_transparency()`` method is called to set the transparency of a chart wall and floor.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

        # ... other code
        wall = chart_doc.first_diagram.wall
        wall.style_area_transparency_transparency(30)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

        # ... other code
        floor = chart_doc.first_diagram.floor
        floor.style_area_transparency_transparency(30)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`eceee99c-fa50-400f-a55c-343ee6966e6d_1` and :numref:`c21fd97d-d12c-4779-8cba-c45f49ad03be_1`.

.. cssclass:: screen_shot

    .. _eceee99c-fa50-400f-a55c-343ee6966e6d_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/eceee99c-fa50-400f-a55c-343ee6966e6d
        :alt: Chart with transparency applied to wall and floor
        :figclass: align-center
        :width: 450px

        Chart with transparency applied to wall and floor

.. cssclass:: screen_shot

    .. _c21fd97d-d12c-4779-8cba-c45f49ad03be_1:

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

The ``style_area_transparency_gradient()`` method is called to set the gradient transparency of a chart.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.data_type.intensity_range import IntensityRange

        # ... other code
        wall = chart_doc.first_diagram.wall
        wall.style_area_transparency_gradient(
            angle=30,
            grad_intensity=IntensityRange(0, 100),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to Floor.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.data_type.intensity_range import IntensityRange

        # ... other code
        floor = chart_doc.first_diagram.floor
        floor.style_area_transparency_gradient(
            angle=30,
            grad_intensity=IntensityRange(0, 100),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`689bf589-8de2-49a0-b260-9f94244aacde_1` and :numref:`0f8ac32f-e2d2-41c1-b0ad-a3ead8371ee9_1`.

.. cssclass:: screen_shot

    .. _689bf589-8de2-49a0-b260-9f94244aacde_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/689bf589-8de2-49a0-b260-9f94244aacde
        :alt: Chart with wall and floor gradient transparency
        :figclass: align-center
        :width: 450px

        Chart with wall and floor gradient transparency

.. cssclass:: screen_shot

    .. _0f8ac32f-e2d2-41c1-b0ad-a3ead8371ee9_1:

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
        - :py:class:`~ooodev.loader.Lo`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`