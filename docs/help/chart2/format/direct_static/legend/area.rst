.. _help_chart2_format_direct_static_legend_area:

Chart2 Direct Legend Area
=========================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

The :py:mod:`ooodev.format.chart2.direct.legend.area` module is used to format the Legend parts of a Chart.

Calls to the :py:meth:`Chart2.style_legend() <ooodev.office.chart2.Chart2.style_legend>` and method is used to style legend.

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.chart2.direct.legend.area import Color as LegendAreaColor
        from ooodev.format.chart2.direct.legend.transparency import Transparency as LegendTransparency
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
        from ooodev.format.chart2.direct.general.area import GradientStyle, ColorRange
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.kind.zoom_kind import ZoomKind
        from ooodev.loader.lo import Lo


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("pie_chart.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, ZoomKind.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="pie_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.BRICK, width=1)
                chart_grad = ChartGradient(
                    chart_doc=chart_doc,
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(StandardColor.GREEN_DARK4, StandardColor.TEAL_LIGHT2),
                )
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                legend_color_style = LegendAreaColor(color=StandardColor.GREEN_LIGHT2)
                legend_bg_transparency_style = LegendTransparency(0)
                Chart2.style_legend(
                    chart_doc=chart_doc, styles=[legend_bg_transparency_style, legend_color_style]
                )

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

Color
-----

The :py:class:`ooodev.format.chart2.direct.legend.area.Color` class is used to set the background color of a Chart legend.

Before formatting the chart is visible in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

In order for the background color to be visible the transparency of the legend must be set. See also: :ref:`help_chart2_format_direct_legend_transparency`.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 1,5

        from ooodev.format.chart2.direct.legend.area import Color as LegendAreaColor
        from ooodev.format.chart2.direct.legend.transparency import Transparency as LegendTransparency
        # ... other code

        legend_color_style = LegendAreaColor(color=StandardColor.GREEN_LIGHT2)
        legend_bg_transparency_style = LegendTransparency(0)
        Chart2.style_legend(
            chart_doc=chart_doc, styles=[legend_bg_transparency_style, legend_color_style]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`36dc662a-dc69-4873-a2d4-1dc8ecb38874` and :numref:`a9f41364-cf39-4f26-b00c-60e96870f6b5`.


.. cssclass:: screen_shot

    .. _36dc662a-dc69-4873-a2d4-1dc8ecb38874:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/36dc662a-dc69-4873-a2d4-1dc8ecb38874
        :alt: Chart with Legend Area Color set
        :figclass: align-center
        :width: 450px

        Chart with Legend Area Color set

.. cssclass:: screen_shot

    .. _a9f41364-cf39-4f26-b00c-60e96870f6b5:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a9f41364-cf39-4f26-b00c-60e96870f6b5
        :alt: Chart Legend Area Color Dialog
        :figclass: align-center
        :width: 450px

        Chart Legend Area Color Dialog

Gradient
--------

The :py:class:`ooodev.format.chart2.direct.legend.area.Gradient` class is used to set the Legend gradient of a Chart.

Before formatting the chart is visible in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

In order for the gradient to be visible the transparency of the legend must be set. See also: :ref:`help_chart2_format_direct_legend_transparency`.

Gradient from preset
^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` enum is used to select the preset gradient.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 1,2,6,7,8

        from ooodev.format.chart2.direct.legend.area import Gradient as LegendAreaGradient
        from ooodev.format.chart2.direct.legend.area import Gradient as PresetGradientKind
        from ooodev.format.chart2.direct.legend.transparency import Transparency as LegendTransparency
        # ... other code

        legend_area_gradient_style = LegendAreaGradient.from_preset(
            chart_doc=chart_doc, preset=PresetGradientKind.NEON_LIGHT
        )
        legend_bg_transparency_style = LegendTransparency(0)
        Chart2.style_legend(
            chart_doc=chart_doc, styles=[legend_bg_transparency_style, legend_area_gradient_style]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`21a484ad-3105-4074-b4f3-449eff0febfc` and :numref:`a7742330-73d4-4f3d-9385-5c03b115f63f`.


.. cssclass:: screen_shot

    .. _21a484ad-3105-4074-b4f3-449eff0febfc:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/21a484ad-3105-4074-b4f3-449eff0febfc
        :alt: Chart with gradient Legend
        :figclass: align-center
        :width: 450px

        Chart with gradient Legend

.. cssclass:: screen_shot

    .. _a7742330-73d4-4f3d-9385-5c03b115f63f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a7742330-73d4-4f3d-9385-5c03b115f63f
        :alt: Chart Area Legend Gradient Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Legend Gradient Dialog


Apply a custom Gradient
^^^^^^^^^^^^^^^^^^^^^^^

Demonstrates how to create a custom gradient.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 1,5,6,7,8,9,10,11

        from ooodev.format.chart2.direct.legend.area import Gradient as LegendAreaGradient
        from ooodev.format.chart2.direct.legend.transparency import Transparency as LegendTransparency
        # ... other code

        legend_area_gradient_style = LegendAreaGradient(
            chart_doc=chart_doc,
            step_count=64,
            style=GradientStyle.SQUARE,
            angle=45,
            grad_color=ColorRange(StandardColor.BRICK_LIGHT1, StandardColor.TEAL_DARK1),
        )
        legend_bg_transparency_style = LegendTransparency(0)
        Chart2.style_legend(
            chart_doc=chart_doc, styles=[legend_bg_transparency_style, legend_area_gradient_style]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`ffd75758-f6c7-4363-8042-8e8bf6687ab5` and :numref:`71ec18f9-e8a8-43ca-98c7-61a7afa470cf`.


.. cssclass:: screen_shot

    .. _ffd75758-f6c7-4363-8042-8e8bf6687ab5:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ffd75758-f6c7-4363-8042-8e8bf6687ab5
        :alt: Chart Legend area with gradient Legend
        :figclass: align-center
        :width: 450px

        Chart Legend area with gradient Legend

.. cssclass:: screen_shot

    .. _71ec18f9-e8a8-43ca-98c7-61a7afa470cf:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/71ec18f9-e8a8-43ca-98c7-61a7afa470cf
        :alt: Chart Legend Area Gradient Dialog
        :figclass: align-center
        :width: 450px

        Chart Legend Area Gradient Dialog

Image
-----

The :py:class:`ooodev.format.chart2.direct.legend.area.Img` class is used to set the background image of the Legend.

Before formatting the chart is visible in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

In order for the image to be visible the transparency of the legend must be set. See also: :ref:`help_chart2_format_direct_legend_transparency`.

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` enum is used to select an image preset.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 1,5,6,7

        from ooodev.format.chart2.direct.legend.area import Img as LegendAreaImg, PresetImageKind
        from ooodev.format.chart2.direct.legend.transparency import Transparency as LegendTransparency
        # ... other code

        legend_img_style = LegendAreaImg.from_preset(
            chart_doc=chart_doc, preset=PresetImageKind.PARCHMENT_PAPER
        )
        legend_bg_transparency_style = LegendTransparency(0)
        Chart2.style_legend(
            chart_doc=chart_doc, styles=[legend_bg_transparency_style, legend_img_style]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`3558a0c0-627f-41a5-979e-0b173557dd8f` and :numref:`7dc81f18-c208-454c-b0cc-0a83397a8076`.

.. cssclass:: screen_shot

    .. _3558a0c0-627f-41a5-979e-0b173557dd8f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3558a0c0-627f-41a5-979e-0b173557dd8f
        :alt: Chart Legend with background image
        :figclass: align-center
        :width: 450px

        Chart Legend with background image

.. cssclass:: screen_shot

    .. _7dc81f18-c208-454c-b0cc-0a83397a8076:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7dc81f18-c208-454c-b0cc-0a83397a8076
        :alt: Chart Area Legend Image Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Legend Image Dialog

Pattern
-------

The :py:class:`ooodev.format.chart2.direct.legend.area.Pattern` class is used to set the background pattern of a Chart.

Before formatting the chart is visible in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` enum is used to select a pattern preset.

In order for the pattern to be visible the transparency of the legend must be set. See also: :ref:`help_chart2_format_direct_legend_transparency`.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 1,2,6,7,8

        from ooodev.format.chart2.direct.legend.area import Pattern as LegendAreaPattern
        from ooodev.format.chart2.direct.legend.area import PresetPatternKind
        from ooodev.format.chart2.direct.legend.transparency import Transparency as LegendTransparency
        # ... other code

        legend_pattern_style = LegendAreaPattern.from_preset(
            chart_doc=chart_doc, preset=PresetPatternKind.HORIZONTAL_BRICK
        )
        legend_bg_transparency_style = LegendTransparency(0)
        Chart2.style_legend(
            chart_doc=chart_doc, styles=[legend_bg_transparency_style, legend_pattern_style]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`4870f30b-be4d-493a-87eb-d1195621a12e` and :numref:`7c634953-b9e0-4892-bd80-8bc93f854a71`.


.. cssclass:: screen_shot

    .. _4870f30b-be4d-493a-87eb-d1195621a12e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4870f30b-be4d-493a-87eb-d1195621a12e
        :alt: Chart Legend with pattern
        :figclass: align-center
        :width: 450px

        Chart Legend with pattern

.. cssclass:: screen_shot

    .. _7c634953-b9e0-4892-bd80-8bc93f854a71:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7c634953-b9e0-4892-bd80-8bc93f854a71
        :alt: Chart Area Legend Pattern Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Legend Pattern Dialog

Hatch
-----

The :py:class:`ooodev.format.chart2.direct.legend.area.Hatch` class is used to set the Title and Subtitle hatch of a Chart.

Before formatting the chart is visible in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` enum is used to select a hatch preset.

In order for the hatch to be visible the transparency of the legend must be set. See also: :ref:`help_chart2_format_direct_legend_transparency`.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 1,5,6,7

        from ooodev.format.chart2.direct.legend.area import Hatch as LegendAreaHatch, PresetHatchKind
        from ooodev.format.chart2.direct.legend.transparency import Transparency as LegendTransparency
        # ... other code

        legend_hatch_style = LegendAreaHatch.from_preset(
            chart_doc=chart_doc, preset=PresetHatchKind.YELLOW_45_DEGREES_CROSSED
        )
        legend_bg_transparency_style = LegendTransparency(0)
        Chart2.style_legend(
            chart_doc=chart_doc, styles=[legend_bg_transparency_style, legend_hatch_style]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`acad9e8e-bdb9-4ac1-b6a8-007d2c7ad250` and :numref:`b7362646-e286-485e-8b9f-ca115be3d1ff`.

.. cssclass:: screen_shot

    .. _acad9e8e-bdb9-4ac1-b6a8-007d2c7ad250:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/acad9e8e-bdb9-4ac1-b6a8-007d2c7ad250
        :alt: Chart Legend with hatch
        :figclass: align-center
        :width: 450px

        Chart Legend with hatch

.. cssclass:: screen_shot

    .. _b7362646-e286-485e-8b9f-ca115be3d1ff:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/b7362646-e286-485e-8b9f-ca115be3d1ff
        :alt: Chart Title Hatch Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Hatch Dialog

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_general_area`
        - :ref:`help_chart2_format_direct_legend_transparency`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.legend.area.Color`
        - :py:class:`ooodev.format.chart2.direct.legend.area.Gradient`
        - :py:class:`ooodev.format.chart2.direct.legend.area.Img`
        - :py:class:`ooodev.format.chart2.direct.legend.area.Pattern`
        - :py:class:`ooodev.format.chart2.direct.legend.area.Hatch`