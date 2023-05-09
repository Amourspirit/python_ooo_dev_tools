.. _help_chart2_format_direct_general_area:

Chart2 Direct General Area
==========================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Overview
--------

The :py:mod:`ooodev.format.chart2.direct.general.area` module is used to set the background of a Chart.


Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.chart2.direct.general.area import Color as ChartColor
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

                chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
                chart_color = ChartColor(color=StandardColor.GREEN_LIGHT2)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_color, chart_bdr_line])

                f_style = ChartColor.from_obj(chart_doc.getPageBackground())
                assert f_style is not None

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

The :py:class:`ooodev.format.chart2.direct.general.area.Color` class is used to set the background color of a Chart.

Before setting the background color of the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Apply the background color to a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 5,6

        from ooodev.format.chart2.direct.general.area import Color as ChartColor
        # ... other code

        chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
        chart_color = ChartColor(color=StandardColor.GREEN_LIGHT2)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_color, chart_bdr_line])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`236884919-28fb1be6-5cbd-4bdf-95e1-5cacb75a65ef` and :numref:`236885274-e11f0494-063b-4035-a2d1-0482a10470c4`


.. cssclass:: screen_shot

    .. _236884919-28fb1be6-5cbd-4bdf-95e1-5cacb75a65ef:

    .. figure:: https://user-images.githubusercontent.com/4193389/236884919-28fb1be6-5cbd-4bdf-95e1-5cacb75a65ef.png
        :alt: Chart with color set to green
        :figclass: align-center
        :width: 450px

        Chart with color set to green

.. cssclass:: screen_shot

    .. _236885274-e11f0494-063b-4035-a2d1-0482a10470c4:

    .. figure:: https://user-images.githubusercontent.com/4193389/236885274-e11f0494-063b-4035-a2d1-0482a10470c4.png
        :alt: Chart Area Color Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Color Dialog

Getting the color From a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = ChartColor.from_obj(chart_doc.getPageBackground())
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Gradient
--------

The :py:class:`ooodev.format.chart2.direct.general.area.Gradient` class is used to set the background gradient of a Chart.

Before setting the background color of the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Gradient from preset
^^^^^^^^^^^^^^^^^^^^

Apply the preset gradient to a Chart
""""""""""""""""""""""""""""""""""""

The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` enum is used to select the preset gradient.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 5,6

        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient, PresetGradientKind

        # ... other code
        chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
        chart_grad = ChartGradient.from_preset(chart_doc, PresetGradientKind.NEON_LIGHT)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`236910715-fbeaae07-9b55-49e9-8e75-318bf28c72ab` and :numref:`236910913-c636dd2b-29b2-47d4-9cb6-d38b7afd75f1`


.. cssclass:: screen_shot

    .. _236910715-fbeaae07-9b55-49e9-8e75-318bf28c72ab:

    .. figure:: https://user-images.githubusercontent.com/4193389/236910715-fbeaae07-9b55-49e9-8e75-318bf28c72ab.png
        :alt: Chart with gradient background
        :figclass: align-center
        :width: 450px

        Chart with gradient background

.. cssclass:: screen_shot

    .. _236910913-c636dd2b-29b2-47d4-9cb6-d38b7afd75f1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236910913-c636dd2b-29b2-47d4-9cb6-d38b7afd75f1.png
        :alt: Chart Area Gradient Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Gradient Dialog


Apply a custom Gradient
^^^^^^^^^^^^^^^^^^^^^^^

Demonstrates how to create a custom gradient.

Apply the preset gradient to a Chart
""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 6,7,8,9,10,11,12

        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient, GradientStyle
        from ooodev.format.chart2.direct.general.area import ColorRange

        # ... other code
        chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
        chart_grad = ChartGradient(
            chart_doc=chart_doc,
            style=GradientStyle.LINEAR,
            angle=45,
            grad_color=ColorRange(StandardColor.GREEN_DARK3, StandardColor.GREEN_LIGHT2),
        )
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`236915417-68679799-feaf-4574-a0c4-6ace0fd4eb6a`


.. cssclass:: screen_shot

    .. _236915417-68679799-feaf-4574-a0c4-6ace0fd4eb6a:

    .. figure:: https://user-images.githubusercontent.com/4193389/236915417-68679799-feaf-4574-a0c4-6ace0fd4eb6a.png
        :alt: Chart with custom gradient background
        :figclass: align-center
        :width: 450px

        Chart with custom gradient background


Getting the gradient from a chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = ChartGradient.from_obj(chart_doc, chart_doc.getPageBackground())
        assert f_style.prop_name == chart_grad.prop_name

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Image
-----

The :py:class:`ooodev.format.chart2.direct.general.area.Img` class is used to set the background image of a Chart.

Before setting the background image of the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.


Apply background image of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` enum is used to select an image preset.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 5,6

        from ooodev.format.chart2.direct.general.area import Img as ChartImg, PresetImageKind
        # ... other code

        chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
        chart_img = ChartImg.from_preset(chart_doc, PresetImageKind.ICE_LIGHT)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_img, chart_bdr_line])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`236939959-33e79374-1504-473e-b2ef-66fa9d9c452d` and :numref:`236940111-f9621402-a9bb-42c4-99bf-e557704344e0`


.. cssclass:: screen_shot

    .. _236939959-33e79374-1504-473e-b2ef-66fa9d9c452d:

    .. figure:: https://user-images.githubusercontent.com/4193389/236939959-33e79374-1504-473e-b2ef-66fa9d9c452d.png
        :alt: Chart with background image
        :figclass: align-center
        :width: 450px

        Chart with background image

.. cssclass:: screen_shot

    .. _236940111-f9621402-a9bb-42c4-99bf-e557704344e0:

    .. figure:: https://user-images.githubusercontent.com/4193389/236940111-f9621402-a9bb-42c4-99bf-e557704344e0.png
        :alt: Chart Area Image Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Image Dialog

Getting the image from a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = ChartImg.from_obj(chart_doc, chart_doc.getPageBackground())
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Pattern
-------

The :py:class:`ooodev.format.chart2.direct.general.area.Pattern` class is used to set the background pattern of a Chart.

Before setting the background pattern of the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.


Apply background pattern of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` enum is used to select a pattern preset.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 5,6

        from ooodev.format.chart2.direct.general.area import Pattern as ChartPattern, PresetPatternKind
        # ... other code

        chart_bdr_line = ChartLineProperties(color=StandardColor.BLUE_LIGHT2, width=0.7)
        chart_pattern = ChartPattern.from_preset(chart_doc, PresetPatternKind.ZIG_ZAG)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_pattern, chart_bdr_line])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`236945010-c70afbc5-3916-4c0c-b67f-2c5a8824e0ae` and :numref:`236945582-b028fc8f-7d40-4384-872d-ca4cdeda1f9e`


.. cssclass:: screen_shot

    .. _236945010-c70afbc5-3916-4c0c-b67f-2c5a8824e0ae:

    .. figure:: https://user-images.githubusercontent.com/4193389/236945010-c70afbc5-3916-4c0c-b67f-2c5a8824e0ae.png
        :alt: Chart with background pattern
        :figclass: align-center
        :width: 450px

        Chart with background pattern

.. cssclass:: screen_shot

    .. _236945582-b028fc8f-7d40-4384-872d-ca4cdeda1f9e:

    .. figure:: https://user-images.githubusercontent.com/4193389/236945582-b028fc8f-7d40-4384-872d-ca4cdeda1f9e.png
        :alt: Chart Area Pattern Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Pattern Dialog

Getting the pattern from a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = ChartPattern.from_obj(chart_doc, chart_doc.getPageBackground())
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Hatch
-----

The :py:class:`ooodev.format.chart2.direct.general.area.Hatch` class is used to set the background hatch of a Chart.

Before setting the background hatch of the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.


Apply background hatch of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` enum is used to select a hatch preset.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 5,6

        from ooodev.format.chart2.direct.general.area import Hatch as ChartHatch, PresetHatchKind
        # ... other code

        chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
        chart_hatch = ChartHatch.from_preset(chart_doc, PresetHatchKind.GREEN_30_DEGREES)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_hatch, chart_bdr_line])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`236948256-33f0c206-6d96-40ee-a8ec-e78e3a59cc91` and :numref:`236948325-4c411e94-2b41-4542-9c4b-185edcc8f828`


.. cssclass:: screen_shot

    .. _236948256-33f0c206-6d96-40ee-a8ec-e78e3a59cc91:

    .. figure:: https://user-images.githubusercontent.com/4193389/236948256-33f0c206-6d96-40ee-a8ec-e78e3a59cc91.png
        :alt: Chart with background hatch
        :figclass: align-center
        :width: 450px

        Chart with background hatch

.. cssclass:: screen_shot

    .. _236948325-4c411e94-2b41-4542-9c4b-185edcc8f828:

    .. figure:: https://user-images.githubusercontent.com/4193389/236948325-4c411e94-2b41-4542-9c4b-185edcc8f828.png
        :alt: Chart Area Hatch Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Hatch Dialog


Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_wall_floor_area`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`