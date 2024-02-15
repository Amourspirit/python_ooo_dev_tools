.. _help_chart2_format_direct_static_series_series_area:

Chart2 Direct series Data Series Area (Static)
==============================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Overview
--------

The :py:mod:`ooodev.format.chart2.direct.series.data_series.area` module contains classes that are used to set the data series area of a Chart.

Calls to the :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
and :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>` methods are used to set the data series area of a Chart.

.. seealso::

    - :ref:`help_chart2_format_direct_series_series_area`

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 27,28

        import uno
        from ooodev.format.chart2.direct.series.data_series.area import Color as DataSeriesColor
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient, PresetGradientKind
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
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

                chart_bdr_line = ChartLineProperties(color=StandardColor.BLUE_LIGHT3, width=0.7)
                chart_grad = ChartGradient.from_preset(chart_doc, PresetGradientKind.TEAL_BLUE)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                data_series_color = DataSeriesColor(StandardColor.TEAL_DARK2)
                Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_color])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Color
-----

The :py:class:`ooodev.format.chart2.direct.series.data_series.area.Color` class is used to set the background color of a data series in Chart.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Apply the background color to a data series
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.area import Color as DataSeriesColor
        # ... other code

        data_series_color = DataSeriesColor(StandardColor.TEAL_DARK2)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_color])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`4f8d241f-a6d7-49b7-9fce-e5a801329163` and :numref:`29ec9307-2ddb-4b85-8865-aa99f216c2bc`.


.. cssclass:: screen_shot

    .. _4f8d241f-a6d7-49b7-9fce-e5a801329163:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4f8d241f-a6d7-49b7-9fce-e5a801329163
        :alt: Chart with data series color set to green
        :figclass: align-center
        :width: 450px

        Chart with data series color set to green

.. cssclass:: screen_shot

    .. _29ec9307-2ddb-4b85-8865-aa99f216c2bc:

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
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=2, styles=[data_series_color])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`4f6bd16f-440c-4bac-a774-1909bac08e7d`.


.. cssclass:: screen_shot

    .. _4f6bd16f-440c-4bac-a774-1909bac08e7d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4f6bd16f-440c-4bac-a774-1909bac08e7d
        :alt: Chart with data point color set to green
        :figclass: align-center
        :width: 450px

        Chart with data point color set to green

Gradient
--------

The :py:class:`ooodev.format.chart2.direct.series.data_series.area.Gradient` class is used to set the background gradient of a Chart.

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

        from ooodev.format.chart2.direct.series.data_series.area import Gradient as DataSeriesGradient

        # ... other code
        data_series_grad = DataSeriesGradient.from_preset(chart_doc, PresetGradientKind.DEEP_OCEAN)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_grad])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`90acf78e-9cd0-4c27-bfe7-67f18cde61ba` and :numref:`79a1ab8e-b004-42be-ad3d-fe99f20e565c`.


.. cssclass:: screen_shot

    .. _90acf78e-9cd0-4c27-bfe7-67f18cde61ba:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/90acf78e-9cd0-4c27-bfe7-67f18cde61ba
        :alt: Chart with gradient data series modified
        :figclass: align-center
        :width: 450px

        Chart with gradient data series modified

.. cssclass:: screen_shot

    .. _79a1ab8e-b004-42be-ad3d-fe99f20e565c:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/79a1ab8e-b004-42be-ad3d-fe99f20e565c
        :alt: Chart Data Series Area Gradient Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Area Gradient Dialog

Style Data Point
~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=-1, styles=[data_series_grad])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`97f6969e-7db1-455c-b259-faed13a83c21`.


.. cssclass:: screen_shot

    .. _97f6969e-7db1-455c-b259-faed13a83c21:

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

        from ooodev.format.chart2.direct.series.data_series.area import Gradient as DataSeriesGradient
        from ooodev.format.chart2.direct.series.data_series.area import GradientStyle, ColorRange
        # ... other code

        data_series_grad = DataSeriesGradient(
            chart_doc=chart_doc,
            style=GradientStyle.LINEAR,
            angle=215,
            grad_color=ColorRange(StandardColor.TEAL_DARK3, StandardColor.BLUE_LIGHT2),
        )
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_grad])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`20125632-2842-4ab6-8264-7db8d4f69a14`


.. cssclass:: screen_shot

    .. _20125632-2842-4ab6-8264-7db8d4f69a14:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/20125632-2842-4ab6-8264-7db8d4f69a14
        :alt: Chart with custom gradient data series formatting
        :figclass: align-center
        :width: 450px

        Chart with custom gradient data series formatting


Image
-----

The :py:class:`ooodev.format.chart2.direct.series.data_series.area.Img` class is used to set the data series background image of a Chart.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.


Apply background image of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` enum is used to select an image preset.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.area import Img as SeriesImg
        from ooodev.format.chart2.direct.series.data_series.area import PresetImageKind
        # ... other code

        data_series_img = SeriesImg.from_preset(chart_doc, PresetImageKind.POOL)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_img])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9bc504c1-7b59-4405-be2f-5a25bbcb46cf` and :numref:`f4bb389f-71fb-40a7-9d53-3608780135f4`.


.. cssclass:: screen_shot

    .. _9bc504c1-7b59-4405-be2f-5a25bbcb46cf:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9bc504c1-7b59-4405-be2f-5a25bbcb46cf
        :alt: Chart with data series background image
        :figclass: align-center
        :width: 450px

        Chart with data series background image

.. cssclass:: screen_shot

    .. _f4bb389f-71fb-40a7-9d53-3608780135f4:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f4bb389f-71fb-40a7-9d53-3608780135f4
        :alt: Chart Data Series Area Image Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Area Image Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=0, styles=[data_series_img])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9aa0b92d-7686-4e59-bc1b-cea1eeb10e34`.


.. cssclass:: screen_shot

    .. _9aa0b92d-7686-4e59-bc1b-cea1eeb10e34:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9aa0b92d-7686-4e59-bc1b-cea1eeb10e34
        :alt: Chart with data point background image
        :figclass: align-center
        :width: 450px

        Chart with data point background image

Pattern
-------

The :py:class:`ooodev.format.chart2.direct.series.data_series.area.Pattern` class is used to set the background pattern of a Chart.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.


Apply background pattern of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` enum is used to select a pattern preset.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.area import Pattern as SeriesPattern
        from ooodev.format.chart2.direct.series.data_series.area import PresetPatternKind
        # ... other code

        data_series_pattern = SeriesPattern.from_preset(chart_doc, PresetPatternKind.ZIG_ZAG)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_pattern])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`38b5b471-17e3-462e-8e8f-57ea193c77fd` and :numref:`66d5b091-a31f-4291-a51e-ac14f66f80e8`.


.. cssclass:: screen_shot

    .. _38b5b471-17e3-462e-8e8f-57ea193c77fd:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/38b5b471-17e3-462e-8e8f-57ea193c77fd
        :alt: Chart data series with background pattern
        :figclass: align-center
        :width: 450px

        Chart data series with background pattern

.. cssclass:: screen_shot

    .. _66d5b091-a31f-4291-a51e-ac14f66f80e8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/66d5b091-a31f-4291-a51e-ac14f66f80e8
        :alt: Chart Data Series Area Pattern Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Area Pattern Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=4, styles=[data_series_pattern])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`256a9572-ba48-4810-847d-d09f2f7f558d`.


.. cssclass:: screen_shot

    .. _256a9572-ba48-4810-847d-d09f2f7f558d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/256a9572-ba48-4810-847d-d09f2f7f558d
        :alt: Chart data point with background pattern
        :figclass: align-center
        :width: 450px

        Chart data point with background pattern


Hatch
-----

The :py:class:`ooodev.format.chart2.direct.series.data_series.area.Hatch` class is used to set the background hatch of a Chart.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.


Apply background hatch of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` enum is used to select a hatch preset.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_series.area import Hatch as SeriesHatch
        from ooodev.format.chart2.direct.series.data_series.area import PresetHatchKind
        # ... other code

        data_series_hatch = SeriesHatch.from_preset(chart_doc, PresetHatchKind.BLUE_45_DEGREES_CROSSED)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[data_series_hatch])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`331e5a64-f4d3-4eab-a375-6c6df880eed0` and :numref:`7c2912b1-69dd-4342-aa8b-5c8873bc3be8`.


.. cssclass:: screen_shot

    .. _331e5a64-f4d3-4eab-a375-6c6df880eed0:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/331e5a64-f4d3-4eab-a375-6c6df880eed0
        :alt: Chart with data series background hatch
        :figclass: align-center
        :width: 450px

        Chart with data series background hatch

.. cssclass:: screen_shot

    .. _7c2912b1-69dd-4342-aa8b-5c8873bc3be8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7c2912b1-69dd-4342-aa8b-5c8873bc3be8
        :alt: Chart Data Series Area Hatch Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Area Hatch Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=-1, styles=[data_series_hatch])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`87f3fa77-903c-4f37-bbb0-b7692e33bffa`.


.. cssclass:: screen_shot

    .. _87f3fa77-903c-4f37-bbb0-b7692e33bffa:

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
        - :ref:`help_chart2_format_direct_series_series_area`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_wall_floor_area`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
        - :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`