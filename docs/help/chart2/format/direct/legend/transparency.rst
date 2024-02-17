.. _help_chart2_format_direct_legend_transparency:

Chart2 Direct Legend Transparency
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The Legend parts of a Chart can be styled using the various ``style_*`` methods of the :py:class:`~ooodev.calc.chart2.chart_legend.ChartLegend` class.

Here we will see how the ``style_area_transparency_transparency()`` method can be used to set the transparency of the legend.

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 38,39

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor
        from ooodev.utils.data_type.color_range import ColorRange

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "piechart.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_border_line(
                    color=StandardColor.BRICK,
                    width=1,
                )
                _ = chart_doc.style_area_gradient(
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(
                        StandardColor.GREEN_DARK4,
                        StandardColor.TEAL_LIGHT2,
                    ),
                )
                legend = chart_doc.first_diagram.get_legend()
                if legend is None:
                    raise ValueError("Legend is None")

                _ = legend.style_area_color(StandardColor.GREEN_LIGHT2)
                _ = legend.style_area_transparency_transparency(50)

                f_style = legend.style_area_transparency_transparency_get()
                assert f_style is not None

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

Before formatting the chart is seen in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

Setting Transparency
^^^^^^^^^^^^^^^^^^^^

The ``style_area_transparency_transparency()`` method can be used to set the transparency of a chart legend.

The Transparency needs a background color in order to view the transparency. See: :ref:`help_chart2_format_direct_legend_area`.

.. tabs::

    .. code-tab:: python

        # ... other code
        _ = legend.style_area_color(StandardColor.GREEN_LIGHT2)
        _ = legend.style_area_transparency_transparency(50)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`cd864a77-de1d-45d6-b74c-56914b2ffb99_1` and :numref:`de4b284c-8e3f-4b55-9d61-7e23344e01f5_1`.

.. cssclass:: screen_shot

    .. _cd864a77-de1d-45d6-b74c-56914b2ffb99_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/cd864a77-de1d-45d6-b74c-56914b2ffb99
        :alt: Chart with transparency applied to legend
        :figclass: align-center
        :width: 450px

        Chart with transparency applied to legend

.. cssclass:: screen_shot

    .. _de4b284c-8e3f-4b55-9d61-7e23344e01f5_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/de4b284c-8e3f-4b55-9d61-7e23344e01f5
        :alt: Chart Legend Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Legend Transparency Dialog

Getting Transparency
^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        f_style = legend.style_area_transparency_transparency_get()
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Gradient Transparency
---------------------

Before formatting the chart is seen in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

Setting Gradient
^^^^^^^^^^^^^^^^

The :py:class:`ooodev.format.chart2.direct.legend.transparency.Gradient` class can be used to set the gradient transparency of a legend.

Like the Transparency the Gradient Transparency needs a background color in order to view the transparency. See: :ref:`help_chart2_format_direct_legend_area`.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 2,3,4,5,9,10,11

        from ooodev.utils.data_type.intensity_range import IntensityRange
        # ... other code

        _ = legend.style_area_color(StandardColor.GREEN_LIGHT2)
        _ = legend.style_area_transparency_gradient(
            angle=90, grad_intensity=IntensityRange(0, 100)
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The results can bee seen in :numref:`a84c06d4-33b7-4edf-b171-4d9f65cc38ad_1` and :numref:`37e8d8b9-3aa5-48ac-97ba-880d80489d85_1`.

.. cssclass:: screen_shot

    .. _a84c06d4-33b7-4edf-b171-4d9f65cc38ad_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a84c06d4-33b7-4edf-b171-4d9f65cc38ad
        :alt: Chart with legend gradient transparency
        :figclass: align-center
        :width: 450px

        Chart with legend gradient transparency

.. cssclass:: screen_shot

    .. _37e8d8b9-3aa5-48ac-97ba-880d80489d85_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/37e8d8b9-3aa5-48ac-97ba-880d80489d85
        :alt: Chart Legend Gradient Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Legend Gradient Transparency Dialog

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_legend_area`
        - :py:class:`~ooodev.loader.Lo`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.legend.transparency.Transparency`
        - :py:class:`ooodev.format.chart2.direct.legend.transparency.Gradient`