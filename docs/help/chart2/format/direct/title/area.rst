.. _help_chart2_format_direct_title_area:

Chart2 Direct Title/Subtitle Area
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Overview
--------

The various ``style_*`` method are used to format the Title/Subtitle parts of a Chart.

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 39

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
                    color=StandardColor.PURPLE_DARK1,
                    width=0.7,
                )
                _ = chart_doc.style_area_gradient(
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(
                        StandardColor.BLUE_DARK1,
                        StandardColor.PURPLE_LIGHT2,
                    ),
                )

                title = chart_doc.get_title()
                if title is None:
                    raise ValueError("Title not found")

                title.style_area_color(StandardColor.DEFAULT_BLUE)

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

The ``style_area_color()`` method is called to set the background color of a Chart Title or Subtitle.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Apply to a Title
^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        title.style_area_color(StandardColor.DEFAULT_BLUE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`d0c12498-bb40-45a8-a0f4-66dae7141333_1` and :numref:`e8faa720-a716-4bd4-8205-fe9e8178d813_1`.


.. cssclass:: screen_shot

    .. _d0c12498-bb40-45a8-a0f4-66dae7141333_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d0c12498-bb40-45a8-a0f4-66dae7141333
        :alt: Chart with Title Area Color set
        :figclass: align-center
        :width: 450px

        Chart with Title Area Color set

.. cssclass:: screen_shot

    .. _e8faa720-a716-4bd4-8205-fe9e8178d813_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/e8faa720-a716-4bd4-8205-fe9e8178d813
        :alt: Chart Title Color Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Color Dialog

Apply to a Subtitle
^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")

        sub_title.style_area_color(StandardColor.DEFAULT_BLUE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`da29f9f9-aa75-4285-85b1-1ff820346a8a_1`.

.. cssclass:: screen_shot

    .. _da29f9f9-aa75-4285-85b1-1ff820346a8a_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/da29f9f9-aa75-4285-85b1-1ff820346a8a
        :alt: Chart with Subtitle Area Color set
        :figclass: align-center
        :width: 450px

        Chart with Subtitle Area Color set

Gradient
--------

The :py:class:`ooodev.format.chart2.direct.title.area.Gradient` class is used to set the Title/Subtitle gradient of a Chart.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Gradient from preset
^^^^^^^^^^^^^^^^^^^^

Apply the preset gradient
"""""""""""""""""""""""""

The ``style_area_gradient_from_preset()`` method is called to set the background gradient of a Chart Title or Subtitle.

The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` enum is used to select the preset gradient.

Apply to Title
~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

        # ... other code
        title = chart_doc.get_title()
        if title is None:
            raise ValueError("Title not found")

        title.style_area_gradient_from_preset(
            preset=PresetGradientKind.PASTEL_DREAM,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`ee250577-bcac-4e60-8659-22f562bbc3c6_1` and :numref:`7ec766cd-3a72-47a7-af96-b24cf444f6c1_1`.


.. cssclass:: screen_shot

    .. _ee250577-bcac-4e60-8659-22f562bbc3c6_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ee250577-bcac-4e60-8659-22f562bbc3c6
        :alt: Chart with gradient Title
        :figclass: align-center
        :width: 450px

        Chart with gradient Title

.. cssclass:: screen_shot

    .. _7ec766cd-3a72-47a7-af96-b24cf444f6c1_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7ec766cd-3a72-47a7-af96-b24cf444f6c1
        :alt: Chart Title Gradient Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Gradient Dialog

Apply to Subtitle
~~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")

        sub_title.style_area_gradient_from_preset(
            preset=PresetGradientKind.PASTEL_DREAM,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`09f2fffe-81ed-4ccb-ae36-f1efa8b3fdb4_1`.


.. cssclass:: screen_shot

    .. _09f2fffe-81ed-4ccb-ae36-f1efa8b3fdb4_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/09f2fffe-81ed-4ccb-ae36-f1efa8b3fdb4
        :alt: Chart with gradient Title
        :figclass: align-center
        :width: 450px

        Chart with gradient Title

Apply a custom Gradient
^^^^^^^^^^^^^^^^^^^^^^^

The ``style_area_gradient()`` method is called to set the background gradient of a Chart Title or Subtitle.

Apply to Title
""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.utils.data_type.color_range import ColorRange
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.utils.color import StandardColor

        # ... other code
        title = chart_doc.get_title()
        if title is None:
            raise ValueError("Title not found")

        title.style_area_gradient(
            step_count=64,
            style=GradientStyle.SQUARE,
            angle=45,
            grad_color=ColorRange(StandardColor.PURPLE_LIGHT2, StandardColor.BLUE_DARK1),
        )


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`21e3cc9d-2847-4118-9191-f85efb21c3eb_1` and :numref:`f8109112-6d3f-4852-ad7c-5fcbb22db31d_1`.


.. cssclass:: screen_shot

    .. _21e3cc9d-2847-4118-9191-f85efb21c3eb_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/21e3cc9d-2847-4118-9191-f85efb21c3eb
        :alt: Chart with gradient Title
        :figclass: align-center
        :width: 450px

        Chart with gradient Title

.. cssclass:: screen_shot

    .. _f8109112-6d3f-4852-ad7c-5fcbb22db31d_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f8109112-6d3f-4852-ad7c-5fcbb22db31d
        :alt: Chart Title Gradient Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Gradient Dialog

Apply to Subtitle
""""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.utils.data_type.color_range import ColorRange
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.utils.color import StandardColor

        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")

        sub_title.style_area_gradient(
            step_count=64,
            style=GradientStyle.SQUARE,
            angle=45,
            grad_color=ColorRange(StandardColor.PURPLE_LIGHT2, StandardColor.BLUE_DARK1),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`d3e8fa2f-ba8e-4786-8518-2d958214cc55_1`.


.. cssclass:: screen_shot

    .. _d3e8fa2f-ba8e-4786-8518-2d958214cc55_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d3e8fa2f-ba8e-4786-8518-2d958214cc55
        :alt: Chart with gradient Title
        :figclass: align-center
        :width: 450px

        Chart with gradient Title

Image
-----

The ``style_area_image_from_preset()`` method is called to set the background image of a Chart Title or Subtitle.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Apply image
^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` enum is used to select an image preset.

Apply to Title
""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_image import PresetImageKind
        # ... other code

        title = chart_doc.get_title()
        if title is None:
            raise ValueError("Title not found")

        title.style_area_image_from_preset(
            preset=PresetImageKind.SPACE,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`2aa029b8-a3b6-4b66-a069-b8585cedab3d_1` and :numref:`5dc5ef1b-229f-4256-9b3a-902a911bd7bf_1`.

.. cssclass:: screen_shot

    .. _2aa029b8-a3b6-4b66-a069-b8585cedab3d_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2aa029b8-a3b6-4b66-a069-b8585cedab3d
        :alt: Chart with Title Image
        :figclass: align-center
        :width: 450px

        Chart with Title Image

.. cssclass:: screen_shot

    .. _5dc5ef1b-229f-4256-9b3a-902a911bd7bf_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/5dc5ef1b-229f-4256-9b3a-902a911bd7bf
        :alt: Chart Title Image Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Image Dialog

Apply to Subtitle
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_image import PresetImageKind

        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")

        sub_title.style_area_image_from_preset(
            preset=PresetImageKind.SPACE,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`6d7e1b11-40e2-4e36-9eec-e9c97c716eca_1`.

.. cssclass:: screen_shot

    .. _6d7e1b11-40e2-4e36-9eec-e9c97c716eca_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/6d7e1b11-40e2-4e36-9eec-e9c97c716eca
        :alt: Chart with Title Image
        :figclass: align-center
        :width: 450px

        Chart with Title Image

Pattern
-------

The ``style_area_pattern_from_preset()`` method is called to set the background pattern of a Chart Title or Subtitle.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` enum is used to select a pattern preset.

Apply to Title
^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_pattern import PresetPatternKind

        # ... other code
        title = chart_doc.get_title()
        if title is None:
            raise ValueError("Title not found")

        title.style_area_pattern_from_preset(
            preset=PresetPatternKind.HORIZONTAL_BRICK,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`4cd109ba-6d3e-4dca-9754-82f6f24ce175_1` and :numref:`dd96600c-8960-423c-92ad-94b0bebb89c8_1`.


.. cssclass:: screen_shot

    .. _4cd109ba-6d3e-4dca-9754-82f6f24ce175_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4cd109ba-6d3e-4dca-9754-82f6f24ce175
        :alt: Chart with Title pattern
        :figclass: align-center
        :width: 450px

        Chart with Title pattern

.. cssclass:: screen_shot

    .. _dd96600c-8960-423c-92ad-94b0bebb89c8_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/dd96600c-8960-423c-92ad-94b0bebb89c8
        :alt: Chart Title Pattern Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Pattern Dialog

Apply to Subtitle
^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_pattern import PresetPatternKind

        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")

        sub_title.style_area_pattern_from_preset(
            preset=PresetPatternKind.HORIZONTAL_BRICK,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`eb494866-fb9c-46de-9231-a720a258cca2_1`.


.. cssclass:: screen_shot

    .. _eb494866-fb9c-46de-9231-a720a258cca2_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/eb494866-fb9c-46de-9231-a720a258cca2
        :alt: Chart with Subtitle pattern
        :figclass: align-center
        :width: 450px

        Chart with Subtitle pattern


Hatch
-----

The ``style_area_hatch_from_preset()`` method is called to set the background hatch of a Chart Title or Subtitle.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` enum is used to select a hatch preset.

Apply to Title
^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_hatch import PresetHatchKind

        # ... other code
        title = chart_doc.get_title()
        if title is None:
            raise ValueError("Title not found")

        title.style_area_hatch_from_preset(
            preset=PresetHatchKind.YELLOW_45_DEGREES_CROSSED,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9876f022-0e42-4b5d-b07d-86c83f31e30c_1` and :numref:`dc35bca9-d365-4a04-a1e3-8e9c2db96d28_1`.

.. cssclass:: screen_shot

    .. _9876f022-0e42-4b5d-b07d-86c83f31e30c_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9876f022-0e42-4b5d-b07d-86c83f31e30c
        :alt: Chart with Title hatch
        :figclass: align-center
        :width: 450px

        Chart with Title hatch

.. cssclass:: screen_shot

    .. _dc35bca9-d365-4a04-a1e3-8e9c2db96d28_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/dc35bca9-d365-4a04-a1e3-8e9c2db96d28
        :alt: Chart Title Hatch Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Hatch Dialog

Apply to Subtitle
^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.preset.preset_hatch import PresetHatchKind

        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")

        sub_title.style_area_hatch_from_preset(
            preset=PresetHatchKind.YELLOW_45_DEGREES_CROSSED,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`7254600c-c53c-4219-91e5-aaa486709dae_1`.

.. cssclass:: screen_shot

    .. _7254600c-c53c-4219-91e5-aaa486709dae_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7254600c-c53c-4219-91e5-aaa486709dae
        :alt: Chart with Title hatch
        :figclass: align-center
        :width: 450px

        Chart with Title hatch


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