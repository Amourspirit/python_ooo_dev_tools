.. _help_chart2_format_direct_series_series_area:

Chart2 Direct series Data Series Area
=====================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Overview
--------

The Data Series and Data Point of a Chart can be styled using the various ``style_*`` methods of
the :py:class:`~ooodev.calc.chart2.chart_data_series.ChartDataSeries` and :py:class:`~ooodev.calc.chart2.chart_data_point.ChartDataPoint` classes.

Setup
-----

General setup for examples.

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
                ds.style_area_color(StandardColor.TEAL_DARK2)

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Color
-----

The ``style_area_color()`` method is called to set the background color of a data series in Chart.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Apply the background color to a data series
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        ds = chart_doc.get_data_series()[0]
        ds.style_area_color(StandardColor.TEAL_DARK2)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`4f8d241f-a6d7-49b7-9fce-e5a801329163_1` and :numref:`29ec9307-2ddb-4b85-8865-aa99f216c2bc_1`.


.. cssclass:: screen_shot

    .. _4f8d241f-a6d7-49b7-9fce-e5a801329163_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4f8d241f-a6d7-49b7-9fce-e5a801329163
        :alt: Chart with data series color set to green
        :figclass: align-center
        :width: 450px

        Chart with data series color set to green

.. cssclass:: screen_shot

    .. _29ec9307-2ddb-4b85-8865-aa99f216c2bc_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/29ec9307-2ddb-4b85-8865-aa99f216c2bc
        :alt: Chart Area Color Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Color Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[2]
        dp.style_area_color(StandardColor.TEAL_DARK2)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`4f6bd16f-440c-4bac-a774-1909bac08e7d_1`.


.. cssclass:: screen_shot

    .. _4f6bd16f-440c-4bac-a774-1909bac08e7d_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4f6bd16f-440c-4bac-a774-1909bac08e7d
        :alt: Chart with data point color set to green
        :figclass: align-center
        :width: 450px

        Chart with data point color set to green

Gradient
--------

The ``style_area_gradient_from_preset()`` method is called to set the background gradient of a Chart.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Gradient from preset
^^^^^^^^^^^^^^^^^^^^

Apply the preset gradient to a data series
""""""""""""""""""""""""""""""""""""""""""

The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` enum is used to select the preset gradient.

Style Data Series
~~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

        # ... other code
        ds = chart_doc.get_data_series()[0]
        ds.style_area_gradient_from_preset(preset=PresetGradientKind.DEEP_OCEAN)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`90acf78e-9cd0-4c27-bfe7-67f18cde61ba_1` and :numref:`79a1ab8e-b004-42be-ad3d-fe99f20e565c_1`.


.. cssclass:: screen_shot

    .. _90acf78e-9cd0-4c27-bfe7-67f18cde61ba_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/90acf78e-9cd0-4c27-bfe7-67f18cde61ba
        :alt: Chart with gradient data series modified
        :figclass: align-center
        :width: 450px

        Chart with gradient data series modified

.. cssclass:: screen_shot

    .. _79a1ab8e-b004-42be-ad3d-fe99f20e565c_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/79a1ab8e-b004-42be-ad3d-fe99f20e565c
        :alt: Chart Data Series Area Gradient Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Area Gradient Dialog

Style Data Point
~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[-1]
        dp.style_area_gradient_from_preset(preset=PresetGradientKind.DEEP_OCEAN)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`97f6969e-7db1-455c-b259-faed13a83c21_1`.


.. cssclass:: screen_shot

    .. _97f6969e-7db1-455c-b259-faed13a83c21_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/97f6969e-7db1-455c-b259-faed13a83c21
        :alt: Chart with gradient data point modified
        :figclass: align-center
        :width: 450px

        Chart with gradient data point modified


Apply a custom Gradient
^^^^^^^^^^^^^^^^^^^^^^^

Demonstrates how to create a custom gradient.

Apply the preset gradient to a data series
""""""""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.utils.data_type.color_range import ColorRange
        # ... other code

        ds = chart_doc.get_data_series()[0]
        ds.style_area_gradient(
            style=GradientStyle.LINEAR,
            angle=215,
            grad_color=ColorRange(StandardColor.TEAL_DARK3, StandardColor.BLUE_LIGHT2),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`20125632-2842-4ab6-8264-7db8d4f69a14_1`


.. cssclass:: screen_shot

    .. _20125632-2842-4ab6-8264-7db8d4f69a14_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/20125632-2842-4ab6-8264-7db8d4f69a14
        :alt: Chart with custom gradient data series formatting
        :figclass: align-center
        :width: 450px

        Chart with custom gradient data series formatting


Image
-----

The ``style_area_image_from_preset()`` method is called to set the data series background image of a Chart.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.


Apply background image of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` enum is used to select an image preset.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_image import PresetImageKind
        # ... other code

        dds = chart_doc.get_data_series()[0]
        ds.style_area_image_from_preset(preset=PresetImageKind.POOL)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9bc504c1-7b59-4405-be2f-5a25bbcb46cf_1` and :numref:`f4bb389f-71fb-40a7-9d53-3608780135f4_1`.


.. cssclass:: screen_shot

    .. _9bc504c1-7b59-4405-be2f-5a25bbcb46cf_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9bc504c1-7b59-4405-be2f-5a25bbcb46cf
        :alt: Chart with data series background image
        :figclass: align-center
        :width: 450px

        Chart with data series background image

.. cssclass:: screen_shot

    .. _f4bb389f-71fb-40a7-9d53-3608780135f4_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f4bb389f-71fb-40a7-9d53-3608780135f4
        :alt: Chart Data Series Area Image Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Area Image Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_image import PresetImageKind

        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[0]
        dp.style_area_image_from_preset(preset=PresetImageKind.POOL)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9aa0b92d-7686-4e59-bc1b-cea1eeb10e34_1`.


.. cssclass:: screen_shot

    .. _9aa0b92d-7686-4e59-bc1b-cea1eeb10e34_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9aa0b92d-7686-4e59-bc1b-cea1eeb10e34
        :alt: Chart with data point background image
        :figclass: align-center
        :width: 450px

        Chart with data point background image

Pattern
-------

The ``style_area_pattern_from_preset()`` method is called to set the background pattern of a Chart.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.


Apply background pattern of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` enum is used to select a pattern preset.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.area import Pattern as SeriesPattern
        from ooodev.format.inner.preset.preset_pattern import PresetPatternKind
        # ... other code

        ds = chart_doc.get_data_series()[0]
        ds.style_area_pattern_from_preset(
            preset=PresetPatternKind.ZIG_ZAG
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`38b5b471-17e3-462e-8e8f-57ea193c77fd_1` and :numref:`66d5b091-a31f-4291-a51e-ac14f66f80e8_1`.


.. cssclass:: screen_shot

    .. _38b5b471-17e3-462e-8e8f-57ea193c77fd_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/38b5b471-17e3-462e-8e8f-57ea193c77fd
        :alt: Chart data series with background pattern
        :figclass: align-center
        :width: 450px

        Chart data series with background pattern

.. cssclass:: screen_shot

    .. _66d5b091-a31f-4291-a51e-ac14f66f80e8_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/66d5b091-a31f-4291-a51e-ac14f66f80e8
        :alt: Chart Data Series Area Pattern Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Area Pattern Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_pattern import PresetPatternKind

        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[4]
        dp.style_area_pattern_from_preset(preset=PresetPatternKind.ZIG_ZAG)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`256a9572-ba48-4810-847d-d09f2f7f558d_1`.


.. cssclass:: screen_shot

    .. _256a9572-ba48-4810-847d-d09f2f7f558d_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/256a9572-ba48-4810-847d-d09f2f7f558d
        :alt: Chart data point with background pattern
        :figclass: align-center
        :width: 450px

        Chart data point with background pattern


Hatch
-----

The ``style_area_hatch_from_preset()`` method is called to set the background hatch of a Chart.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.


Apply background hatch of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` enum is used to select a hatch preset.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_hatch import PresetHatchKind

        # ... other code

        ds = chart_doc.get_data_series()[0]
        ds.style_area_hatch_from_preset(
            preset=PresetHatchKind.BLUE_45_DEGREES_CROSSED,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`331e5a64-f4d3-4eab-a375-6c6df880eed0_1` and :numref:`7c2912b1-69dd-4342-aa8b-5c8873bc3be8_1`.


.. cssclass:: screen_shot

    .. _331e5a64-f4d3-4eab-a375-6c6df880eed0_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/331e5a64-f4d3-4eab-a375-6c6df880eed0
        :alt: Chart with data series background hatch
        :figclass: align-center
        :width: 450px

        Chart with data series background hatch

.. cssclass:: screen_shot

    .. _7c2912b1-69dd-4342-aa8b-5c8873bc3be8_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7c2912b1-69dd-4342-aa8b-5c8873bc3be8
        :alt: Chart Data Series Area Hatch Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Area Hatch Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_hatch import PresetHatchKind

        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[-1]
        dp.style_area_hatch_from_preset(
            preset=PresetHatchKind.BLUE_45_DEGREES_CROSSED,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`87f3fa77-903c-4f37-bbb0-b7692e33bffa_1`.


.. cssclass:: screen_shot

    .. _87f3fa77-903c-4f37-bbb0-b7692e33bffa_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/87f3fa77-903c-4f37-bbb0-b7692e33bffa
        :alt: Chart with data point background hatch
        :figclass: align-center
        :width: 450px

        Chart with data point background hatch


Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_wall_floor_area`
        - :py:class:`~ooodev.loader.Lo`