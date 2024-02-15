.. _help_chart2_format_direct_legend_area:

Chart2 Direct Legend Area
=========================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

The Legend parts of a Chart can be styled using the various ``style_*`` methods of the :py:class:`~ooodev.calc.chart2.chart_legend.ChartLegend` class.

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 37,38

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
                _ = legend.style_area_transparency_transparency(0)
                _ = legend.style_area_color(color=StandardColor.GREEN_LIGHT2)

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())


Color
-----

The ``style_area_color()`` method is used to set the background color of the legend.
The ``style_area_transparency_transparency()`` method is used to set the transparency of the legend.

Before formatting the chart is visible in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

In order for the background color to be visible the transparency of the legend must be set. See also: :ref:`help_chart2_format_direct_legend_transparency`.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 1,5

        # ... other code
        # set the transparency of the legend to 0 and the color to green light2
        _ = legend.style_area_transparency_transparency(0)
        _ = legend.style_area_color(color=StandardColor.GREEN_LIGHT2)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`36dc662a-dc69-4873-a2d4-1dc8ecb38874_1` and :numref:`a9f41364-cf39-4f26-b00c-60e96870f6b5_1`.


.. cssclass:: screen_shot

    .. _36dc662a-dc69-4873-a2d4-1dc8ecb38874_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/36dc662a-dc69-4873-a2d4-1dc8ecb38874
        :alt: Chart with Legend Area Color set
        :figclass: align-center
        :width: 450px

        Chart with Legend Area Color set

.. cssclass:: screen_shot

    .. _a9f41364-cf39-4f26-b00c-60e96870f6b5_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a9f41364-cf39-4f26-b00c-60e96870f6b5
        :alt: Chart Legend Area Color Dialog
        :figclass: align-center
        :width: 450px

        Chart Legend Area Color Dialog

Gradient
--------

The ``style_area_gradient_from_preset()`` method is used to set the Legend gradient of a Chart.

Before formatting the chart is visible in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

In order for the gradient to be visible the transparency of the legend must be set. See also: :ref:`help_chart2_format_direct_legend_transparency`.

Gradient from preset
^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` enum is used to select the preset gradient.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind
        # ... other code

        _ = legend.style_area_transparency_transparency(0)
        _ = legend.style_area_gradient_from_preset(
            preset=PresetGradientKind.NEON_LIGHT,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`21a484ad-3105-4074-b4f3-449eff0febfc_1` and :numref:`a7742330-73d4-4f3d-9385-5c03b115f63f_1`.


.. cssclass:: screen_shot

    .. _21a484ad-3105-4074-b4f3-449eff0febfc_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/21a484ad-3105-4074-b4f3-449eff0febfc
        :alt: Chart with gradient Legend
        :figclass: align-center
        :width: 450px

        Chart with gradient Legend

.. cssclass:: screen_shot

    .. _a7742330-73d4-4f3d-9385-5c03b115f63f_1:

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

        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.utils.data_type.color_range import ColorRange
        # ... other code

        _ = legend.style_area_transparency_transparency(0)
        _ = legend.style_area_gradient(
            step_count=64,
            style=GradientStyle.SQUARE,
            angle=45,
            grad_color=ColorRange(StandardColor.BRICK_LIGHT1, StandardColor.TEAL_DARK1),
        )


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`ffd75758-f6c7-4363-8042-8e8bf6687ab5_1` and :numref:`71ec18f9-e8a8-43ca-98c7-61a7afa470cf_1`.


.. cssclass:: screen_shot

    .. _ffd75758-f6c7-4363-8042-8e8bf6687ab5_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ffd75758-f6c7-4363-8042-8e8bf6687ab5
        :alt: Chart Legend area with gradient Legend
        :figclass: align-center
        :width: 450px

        Chart Legend area with gradient Legend

.. cssclass:: screen_shot

    .. _71ec18f9-e8a8-43ca-98c7-61a7afa470cf_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/71ec18f9-e8a8-43ca-98c7-61a7afa470cf
        :alt: Chart Legend Area Gradient Dialog
        :figclass: align-center
        :width: 450px

        Chart Legend Area Gradient Dialog

Image
-----

The ``style_area_image_from_preset()`` method is used to set the background image of the Legend.

Before formatting the chart is visible in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

In order for the image to be visible the transparency of the legend must be set. See also: :ref:`help_chart2_format_direct_legend_transparency`.

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` enum is used to select an image preset.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_image import PresetImageKind
        # ... other code

        _ = legend.style_area_transparency_transparency(0)
        _ = legend.style_area_image_from_preset(
            preset=PresetImageKind.PARCHMENT_PAPER,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`3558a0c0-627f-41a5-979e-0b173557dd8f_1` and :numref:`7dc81f18-c208-454c-b0cc-0a83397a8076_1`.

.. cssclass:: screen_shot

    .. _3558a0c0-627f-41a5-979e-0b173557dd8f_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3558a0c0-627f-41a5-979e-0b173557dd8f
        :alt: Chart Legend with background image
        :figclass: align-center
        :width: 450px

        Chart Legend with background image

.. cssclass:: screen_shot

    .. _7dc81f18-c208-454c-b0cc-0a83397a8076_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7dc81f18-c208-454c-b0cc-0a83397a8076
        :alt: Chart Area Legend Image Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Legend Image Dialog

Pattern
-------

The ``style_area_pattern_from_preset()`` method is used to set the background pattern of a Chart.

Before formatting the chart is visible in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` enum is used to select a pattern preset.

In order for the pattern to be visible the transparency of the legend must be set. See also: :ref:`help_chart2_format_direct_legend_transparency`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_pattern import PresetPatternKind
        # ... other code

        _ = legend.style_area_transparency_transparency(0)
        _ = legend.style_area_pattern_from_preset(
            preset=PresetPatternKind.HORIZONTAL_BRICK,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`4870f30b-be4d-493a-87eb-d1195621a12e_1` and :numref:`7c634953-b9e0-4892-bd80-8bc93f854a71_1`.


.. cssclass:: screen_shot

    .. _4870f30b-be4d-493a-87eb-d1195621a12e_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4870f30b-be4d-493a-87eb-d1195621a12e
        :alt: Chart Legend with pattern
        :figclass: align-center
        :width: 450px

        Chart Legend with pattern

.. cssclass:: screen_shot

    .. _7c634953-b9e0-4892-bd80-8bc93f854a71_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7c634953-b9e0-4892-bd80-8bc93f854a71
        :alt: Chart Area Legend Pattern Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Legend Pattern Dialog

Hatch
-----

The ``style_area_hatch_from_preset()`` method is used to set the Title and Subtitle hatch of a Chart.

Before formatting the chart is visible in :numref:`ce52cea5-2b22-4d2a-a158-9e22364d4544`.

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` enum is used to select a hatch preset.

In order for the hatch to be visible the transparency of the legend must be set. See also: :ref:`help_chart2_format_direct_legend_transparency`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_hatch import PresetHatchKind
        # ... other code

        _ = legend.style_area_transparency_transparency(0)
        _ = legend.style_area_hatch_from_preset(
            preset=PresetHatchKind.YELLOW_45_DEGREES_CROSSED,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are visible in :numref:`acad9e8e-bdb9-4ac1-b6a8-007d2c7ad250_1` and :numref:`b7362646-e286-485e-8b9f-ca115be3d1ff_1`.

.. cssclass:: screen_shot

    .. _acad9e8e-bdb9-4ac1-b6a8-007d2c7ad250_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/acad9e8e-bdb9-4ac1-b6a8-007d2c7ad250
        :alt: Chart Legend with hatch
        :figclass: align-center
        :width: 450px

        Chart Legend with hatch

.. cssclass:: screen_shot

    .. _b7362646-e286-485e-8b9f-ca115be3d1ff_1:

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
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`