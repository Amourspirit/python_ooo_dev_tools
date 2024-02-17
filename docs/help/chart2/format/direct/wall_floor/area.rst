.. _help_chart2_format_direct_wall_floor_area:

Chart2 Direct Wall/Floor Area
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Overview
--------

The various ``style_*`` methods are used to format the wall/floor of a Chart.


Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 25

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor

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
                    color=StandardColor.PURPLE_DARK1,
                    width=0.7,
                )

                wall = chart_doc.first_diagram.wall
                wall.style_area_color(StandardColor.DEFAULT_BLUE)

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

The ``style_area_color()`` method is used to set the background color the chart wall and floor.

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.

Apply the background color to a wall and floor
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.color import StandardColor

        # ... other code
        wall = chart_doc.first_diagram.wall
        wall.style_area_color(StandardColor.DEFAULT_BLUE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.color import StandardColor

        # ... other code
        floor = chart_doc.first_diagram.floor
        floor.style_area_color(StandardColor.DEFAULT_BLUE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`181c5c98-e4e1-4519-b91a-ffc39f5fa430_` and :numref:`21408192-4819-4557-beba-b48d881312ee_1`


.. cssclass:: screen_shot

    .. _181c5c98-e4e1-4519-b91a-ffc39f5fa430_:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/181c5c98-e4e1-4519-b91a-ffc39f5fa430
        :alt: Chart with Wall and Floor color set
        :figclass: align-center
        :width: 450px

        Chart with Wall and Floor color set

.. cssclass:: screen_shot

    .. _21408192-4819-4557-beba-b48d881312ee_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/21408192-4819-4557-beba-b48d881312ee
        :alt: Chart Wall Color Dialog
        :figclass: align-center
        :width: 450px

        Chart Wall Color Dialog

Gradient
--------

The ``style_area_gradient_from_preset()`` method is called to set the background gradient of a Chart Wall/Floor.

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.

Gradient from preset
^^^^^^^^^^^^^^^^^^^^

Apply the preset gradient to wall and floor
"""""""""""""""""""""""""""""""""""""""""""

The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` enum is used to select the preset gradient.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

        # ... other code
        wall = chart_doc.first_diagram.wall
        wall.style_area_gradient_from_preset(
            preset=PresetGradientKind.DEEP_OCEAN,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to Floor.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

        # ... other code
        floor = chart_doc.first_diagram.floor
        floor.style_area_gradient_from_preset(
            preset=PresetGradientKind.DEEP_OCEAN,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`3f60aea8-ab07-4831-9f2c-ba13c69bef55_1` and :numref:`a1ca65eb-8f71-4113-b5d9-57f11e1a88d3_1`


.. cssclass:: screen_shot

    .. _3f60aea8-ab07-4831-9f2c-ba13c69bef55_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3f60aea8-ab07-4831-9f2c-ba13c69bef55
        :alt: Chart with gradient wall and floor
        :figclass: align-center
        :width: 450px

        Chart with gradient wall and floor

.. cssclass:: screen_shot

    .. _a1ca65eb-8f71-4113-b5d9-57f11e1a88d3_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a1ca65eb-8f71-4113-b5d9-57f11e1a88d3
        :alt: Chart Wall Gradient Dialog
        :figclass: align-center
        :width: 450px

        Chart Wall Gradient Dialog


Apply a custom Gradient
^^^^^^^^^^^^^^^^^^^^^^^

The ``style_area_gradient()`` method is called to set the background gradient of a Chart Wall/Floor.

Apply the preset gradient to wall and floor
"""""""""""""""""""""""""""""""""""""""""""

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.utils.data_type.color_range import ColorRange
        from ooodev.utils.color import StandardColor

        # ... other code
        wall = chart_doc.first_diagram.wall
        wall.style_area_gradient(
            style=GradientStyle.LINEAR,
            angle=45,
            grad_color=ColorRange(StandardColor.BLUE_DARK3, StandardColor.BLUE_LIGHT2),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.utils.data_type.color_range import ColorRange
        from ooodev.utils.color import StandardColor

        # ... other code
        floor = chart_doc.first_diagram.floor
        floor.style_area_gradient(
            style=GradientStyle.LINEAR,
            angle=45,
            grad_color=ColorRange(StandardColor.BLUE_DARK3, StandardColor.BLUE_LIGHT2),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`1790240c-ce82-4e42-b115-1a41bff70db7_1`


.. cssclass:: screen_shot

    .. _1790240c-ce82-4e42-b115-1a41bff70db7_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1790240c-ce82-4e42-b115-1a41bff70db7
        :alt: Chart with custom gradient background
        :figclass: align-center
        :width: 450px

        Chart with custom gradient background


Image
-----

The ``style_area_image_from_preset()`` or ``style_area_image()`` methods are called to set the background image of the Chart Wall/Floor.

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.


Apply image to wall and floor
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` enum is used to select an image preset.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_image import PresetImageKind

        # ... other code
        wall = chart_doc.first_diagram.wall
        wall.style_area_image_from_preset(preset=PresetImageKind.ICE_LIGHT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_image import PresetImageKind

        # ... other code
        floor = chart_doc.first_diagram.floor
        floor.style_area_image_from_preset(preset=PresetImageKind.ICE_LIGHT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`17e4da98-46c0-47a0-84e0-6d5ba1f13a57_1` and :numref:`7db6608b-e2bc-4c75-a41a-39d3ebf4e05c_1`


.. cssclass:: screen_shot

    .. _17e4da98-46c0-47a0-84e0-6d5ba1f13a57_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/17e4da98-46c0-47a0-84e0-6d5ba1f13a57
        :alt: Chart with wall and floor image
        :figclass: align-center
        :width: 450px

        Chart with wall and floor image

.. cssclass:: screen_shot

    .. _7db6608b-e2bc-4c75-a41a-39d3ebf4e05c_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7db6608b-e2bc-4c75-a41a-39d3ebf4e05c
        :alt: Chart Area Image Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Image Dialog

Pattern
-------

The ``style_area_pattern_from_preset()`` or ``style_area_pattern()`` methods are called to set the background pattern of a Chart Wall/Floor.

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.


Apply background pattern of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` enum is used to select a pattern preset.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_pattern import PresetPatternKind

        # ... other code
        wall = chart_doc.first_diagram.wall
        wall.style_area_pattern_from_preset(preset=PresetPatternKind.ZIG_ZAG)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_pattern import PresetPatternKind

        # ... other code
        floor = chart_doc.first_diagram.floor
        floor.style_area_pattern_from_preset(preset=PresetPatternKind.ZIG_ZAG)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9cc6eeae-d204-4f6d-b10d-18d7434fe156_1` and :numref:`8468ed6a-228d-4ec7-8e21-dda0d70dc7ea_1`


.. cssclass:: screen_shot

    .. _9cc6eeae-d204-4f6d-b10d-18d7434fe156_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9cc6eeae-d204-4f6d-b10d-18d7434fe156
        :alt: Chart with wall and floor pattern
        :figclass: align-center
        :width: 450px

        Chart with wall and floor pattern

.. cssclass:: screen_shot

    .. _8468ed6a-228d-4ec7-8e21-dda0d70dc7ea_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/8468ed6a-228d-4ec7-8e21-dda0d70dc7ea
        :alt: Chart Wall Pattern Dialog
        :figclass: align-center
        :width: 450px

        Chart Wall Pattern Dialog


Hatch
-----

The ``style_area_hatch_from_preset()`` or ``style_area_hatch()`` methods are called to set the background hatch of a Chart Wall/Floor.

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.


Apply background hatch of a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` enum is used to select a hatch preset.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_hatch import PresetHatchKind

        # ... other code
        wall = chart_doc.first_diagram.wall
        wall.style_area_hatch_from_preset(preset=PresetHatchKind.BLUE_45_DEGREES)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_hatch import PresetHatchKind

        # ... other code
        floor = chart_doc.first_diagram.floor
        floor.style_area_hatch_from_preset(preset=PresetHatchKind.BLUE_45_DEGREES)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`cec9bb9e-9edb-46dc-96c1-5fc57069973a_1` and :numref:`92b60156-00b7-4c75-bbb2-a7fa829992b3_1`


.. cssclass:: screen_shot

    .. _cec9bb9e-9edb-46dc-96c1-5fc57069973a_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/cec9bb9e-9edb-46dc-96c1-5fc57069973a
        :alt: Chart with wall and floor hatch
        :figclass: align-center
        :width: 450px

        Chart with wall and floor hatch

.. cssclass:: screen_shot

    .. _92b60156-00b7-4c75-bbb2-a7fa829992b3_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/92b60156-00b7-4c75-bbb2-a7fa829992b3
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
        - :ref:`help_chart2_format_direct_general_area`
        - :py:class:`~ooodev.loader.Lo`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`