.. _help_chart2_format_direct_wall_floor_area:

Chart2 Direct Wall/Floor Area
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Overview
--------

The :py:mod:`ooodev.format.chart2.direct.wall` module is used to format the wall/floor of a Chart.


Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.chart2.direct.wall.area import Color as WallColor
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("col_chart3d.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="col_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.BLUE_DARK1, width=1.0)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_bdr_line])

                wall_color = WallColor(color=StandardColor.DEFAULT_BLUE)
                Chart2.style_wall(chart_doc=chart_doc, styles=[wall_color])

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

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.

Apply the background color to a wall and floor
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.wall.area import Color as WallColor
        # ... other code

        # wall color
        wall_color = WallColor(color=StandardColor.DEFAULT_BLUE)
        Chart2.style_wall(chart_doc=chart_doc, styles=[wall_color])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        # floor color
        floor_color = WallColor(color=StandardColor.BLUE_DARK1)
        Chart2.style_floor(chart_doc=chart_doc, styles=[floor_color])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`181c5c98-e4e1-4519-b91a-ffc39f5fa430` and :numref:`21408192-4819-4557-beba-b48d881312ee`


.. cssclass:: screen_shot

    .. _181c5c98-e4e1-4519-b91a-ffc39f5fa430:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/181c5c98-e4e1-4519-b91a-ffc39f5fa430
        :alt: Chart with Wall and Floor color set
        :figclass: align-center
        :width: 450px

        Chart with Wall and Floor color set

.. cssclass:: screen_shot

    .. _21408192-4819-4557-beba-b48d881312ee:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/21408192-4819-4557-beba-b48d881312ee
        :alt: Chart Wall Color Dialog Modified
        :figclass: align-center
        :width: 450px

        Chart Wall Color Dialog Modified

Gradient
--------

The :py:class:`ooodev.format.chart2.direct.wall.area.Gradient` class is used to set the background gradient of a Chart.

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.

Gradient from preset
^^^^^^^^^^^^^^^^^^^^

Apply the preset gradient to wall and floor
"""""""""""""""""""""""""""""""""""""""""""

The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` enum is used to select the preset gradient.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.wall.area import Gradient as WallGradient, PresetGradientKind

        # ... other code
        wall_grad = WallGradient.from_preset(chart_doc, PresetGradientKind.DEEP_OCEAN)
        Chart2.style_wall(chart_doc=chart_doc, styles=[wall_grad])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to Floor.

.. tabs::

    .. code-tab:: python

        floor_grad = WallGradient.from_preset(chart_doc, PresetGradientKind.MIDNIGHT)
        Chart2.style_floor(chart_doc=chart_doc, styles=[floor_grad])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`3f60aea8-ab07-4831-9f2c-ba13c69bef55` and :numref:`a1ca65eb-8f71-4113-b5d9-57f11e1a88d3`


.. cssclass:: screen_shot

    .. _3f60aea8-ab07-4831-9f2c-ba13c69bef55:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3f60aea8-ab07-4831-9f2c-ba13c69bef55
        :alt: Chart with gradient wall and floor
        :figclass: align-center
        :width: 450px

        Chart with gradient wall and floor

.. cssclass:: screen_shot

    .. _a1ca65eb-8f71-4113-b5d9-57f11e1a88d3:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a1ca65eb-8f71-4113-b5d9-57f11e1a88d3
        :alt: Chart Wall Gradient Dialog Modified
        :figclass: align-center
        :width: 450px

        Chart Wall Gradient Dialog Modified


Apply a custom Gradient
^^^^^^^^^^^^^^^^^^^^^^^

Demonstrates how to create a custom gradient.

Apply the preset gradient to wall and floor
"""""""""""""""""""""""""""""""""""""""""""

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.wall.area import Gradient as WallGradient, GradientStyle
        from ooodev.format.chart2.direct.wall.area import ColorRange

        # ... other code
        wall_grad = WallGradient(
            chart_doc=chart_doc,
            style=GradientStyle.LINEAR,
            angle=45,
            grad_color=ColorRange(StandardColor.BLUE_DARK3, StandardColor.BLUE_LIGHT2),
        )
        Chart2.style_wall(chart_doc=chart_doc, styles=[wall_grad])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        floor_grad = WallGradient(
            chart_doc=chart_doc,
            style=GradientStyle.LINEAR,
            angle=-10,
            grad_color=ColorRange(StandardColor.BLUE_DARK4, StandardColor.BLUE),
        )
        Chart2.style_floor(chart_doc=chart_doc, styles=[floor_grad])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`1790240c-ce82-4e42-b115-1a41bff70db7`


.. cssclass:: screen_shot

    .. _1790240c-ce82-4e42-b115-1a41bff70db7:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1790240c-ce82-4e42-b115-1a41bff70db7
        :alt: Chart with custom gradient background
        :figclass: align-center
        :width: 450px

        Chart with custom gradient background


Image
-----

The :py:class:`ooodev.format.chart2.direct.wall.area.Img` class is used to set the background image of the wall and floor.

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.


Apply image to wall and floor
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` enum is used to select an image preset.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.wall.area import Img as WallImg, PresetImageKind
        # ... other code

        wall_img = WallImg.from_preset(chart_doc, PresetImageKind.ICE_LIGHT)
        Chart2.style_wall(chart_doc=chart_doc, styles=[wall_img])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        floor_img = WallImg.from_preset(chart_doc, PresetImageKind.MARBLE)
        Chart2.style_floor(chart_doc=chart_doc, styles=[floor_img])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`17e4da98-46c0-47a0-84e0-6d5ba1f13a57` and :numref:`7db6608b-e2bc-4c75-a41a-39d3ebf4e05c`


.. cssclass:: screen_shot

    .. _17e4da98-46c0-47a0-84e0-6d5ba1f13a57:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/17e4da98-46c0-47a0-84e0-6d5ba1f13a57
        :alt: Chart with wall and floor image
        :figclass: align-center
        :width: 450px

        Chart with wall and floor image

.. cssclass:: screen_shot

    .. _7db6608b-e2bc-4c75-a41a-39d3ebf4e05c:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7db6608b-e2bc-4c75-a41a-39d3ebf4e05c
        :alt: Chart Area Image Dialog Modified
        :figclass: align-center
        :width: 450px

        Chart Area Image Dialog Modified

Pattern
-------

The :py:class:`ooodev.format.chart2.direct.wall.area.Pattern` class is used to set the background pattern of a Chart.

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.


Apply background pattern of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` enum is used to select a pattern preset.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.wall.area import Pattern as WallPattern, PresetPatternKind
        # ... other code

        wall_pattern = WallPattern.from_preset(chart_doc, PresetPatternKind.ZIG_ZAG)
        Chart2.style_wall(chart_doc=chart_doc, styles=[wall_pattern])


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        floor_pattern = WallPattern.from_preset(chart_doc, PresetPatternKind.PERCENT_20)
        Chart2.style_floor(chart_doc=chart_doc, styles=[floor_pattern])


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9cc6eeae-d204-4f6d-b10d-18d7434fe156` and :numref:`8468ed6a-228d-4ec7-8e21-dda0d70dc7ea`


.. cssclass:: screen_shot

    .. _9cc6eeae-d204-4f6d-b10d-18d7434fe156:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9cc6eeae-d204-4f6d-b10d-18d7434fe156
        :alt: Chart with wall and floor pattern
        :figclass: align-center
        :width: 450px

        Chart with wall and floor pattern

.. cssclass:: screen_shot

    .. _8468ed6a-228d-4ec7-8e21-dda0d70dc7ea:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/8468ed6a-228d-4ec7-8e21-dda0d70dc7ea
        :alt: Chart Wall Pattern Dialog Modified
        :figclass: align-center
        :width: 450px

        Chart Wall Pattern Dialog Modified


Hatch
-----

The :py:class:`ooodev.format.chart2.direct.wall.area.Hatch` class is used to set the background hatch of a Chart.

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.


Apply background hatch of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` enum is used to select a hatch preset.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.wall.area import Hatch as WallHatch, PresetHatchKind
        # ... other code

        wall_hatch = WallHatch.from_preset(chart_doc, PresetHatchKind.BLUE_45_DEGREES_CROSSED)
        Chart2.style_wall(chart_doc=chart_doc, styles=[wall_hatch])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        floor_hatch = WallHatch.from_preset(chart_doc, PresetHatchKind.BLUE_45_DEGREES)
        Chart2.style_floor(chart_doc=chart_doc, styles=[floor_hatch])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`cec9bb9e-9edb-46dc-96c1-5fc57069973a` and :numref:`92b60156-00b7-4c75-bbb2-a7fa829992b3`


.. cssclass:: screen_shot

    .. _cec9bb9e-9edb-46dc-96c1-5fc57069973a:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/cec9bb9e-9edb-46dc-96c1-5fc57069973a
        :alt: Chart with wall and floor hatch
        :figclass: align-center
        :width: 450px

        Chart with wall and floor hatch

.. cssclass:: screen_shot

    .. _92b60156-00b7-4c75-bbb2-a7fa829992b3:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/92b60156-00b7-4c75-bbb2-a7fa829992b3
        :alt: Chart Area Hatch Dialog Modified
        :figclass: align-center
        :width: 450px

        Chart Area Hatch Dialog Modified


Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_general_area`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_wall() <ooodev.office.chart2.Chart2.style_wall>`
        - :py:meth:`Chart2.style_floor() <ooodev.office.chart2.Chart2.style_floor>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:mod:`ooodev.format.chart2.direct.wall`