.. _help_chart2_format_direct_static_title_area:

Chart2 Direct Title/Subtitle Area (Static)
==========================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Overview
--------

The :py:mod:`ooodev.format.chart2.direct.title.area` module is used to format the Title/Subtitle parts of a Chart.

Calls to the :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>` and
:py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>` methods are used to set the Title and Subtitle formatting of a Chart.

.. seealso::

    - :ref:`help_chart2_format_direct_title_area`

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 34,35

        import uno
        from ooodev.format.chart2.direct.title.area import Color as TitleBgColor
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
        from ooodev.format.chart2.direct.general.area import GradientStyle, ColorRange
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("pie_flat_chart.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="pie_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.PURPLE_DARK1, width=0.7)
                chart_grad = ChartGradient(
                    chart_doc=chart_doc,
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(StandardColor.BLUE_DARK1, StandardColor.PURPLE_LIGHT2),
                )
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                title_color = TitleBgColor(color=StandardColor.DEFAULT_BLUE)
                Chart2.style_title(chart_doc=chart_doc, styles=[title_color])

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

The :py:class:`ooodev.format.chart2.direct.title.area.Color` class is used to set the background color of a Chart Title or Subtitle.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Apply to a Title
^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.area import Color as TitleBgColor
        # ... other code

        title_color = TitleBgColor(color=StandardColor.DEFAULT_BLUE)
        Chart2.style_title(chart_doc=chart_doc, styles=[title_color])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`d0c12498-bb40-45a8-a0f4-66dae7141333` and :numref:`e8faa720-a716-4bd4-8205-fe9e8178d813`.


.. cssclass:: screen_shot

    .. _d0c12498-bb40-45a8-a0f4-66dae7141333:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d0c12498-bb40-45a8-a0f4-66dae7141333
        :alt: Chart with Title Area Color set
        :figclass: align-center
        :width: 450px

        Chart with Title Area Color set

.. cssclass:: screen_shot

    .. _e8faa720-a716-4bd4-8205-fe9e8178d813:

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
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[title_color])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`da29f9f9-aa75-4285-85b1-1ff820346a8a`.

.. cssclass:: screen_shot

    .. _da29f9f9-aa75-4285-85b1-1ff820346a8a:

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


The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` enum is used to select the preset gradient.

Apply to Title
~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.area import Gradient as TitleGrad, PresetGradientKind

        # ... other code
        title_grad = TitleGrad.from_preset(chart_doc, PresetGradientKind.PASTEL_DREAM)
        Chart2.style_title(chart_doc=chart_doc, styles=[title_grad])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`ee250577-bcac-4e60-8659-22f562bbc3c6` and :numref:`7ec766cd-3a72-47a7-af96-b24cf444f6c1`.


.. cssclass:: screen_shot

    .. _ee250577-bcac-4e60-8659-22f562bbc3c6:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ee250577-bcac-4e60-8659-22f562bbc3c6
        :alt: Chart with gradient Title
        :figclass: align-center
        :width: 450px

        Chart with gradient Title

.. cssclass:: screen_shot

    .. _7ec766cd-3a72-47a7-af96-b24cf444f6c1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7ec766cd-3a72-47a7-af96-b24cf444f6c1
        :alt: Chart Title Gradient Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Gradient Dialog

Apply to Subtitle
~~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[title_grad])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`09f2fffe-81ed-4ccb-ae36-f1efa8b3fdb4`.


.. cssclass:: screen_shot

    .. _09f2fffe-81ed-4ccb-ae36-f1efa8b3fdb4:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/09f2fffe-81ed-4ccb-ae36-f1efa8b3fdb4
        :alt: Chart with gradient Title
        :figclass: align-center
        :width: 450px

        Chart with gradient Title

Apply a custom Gradient
^^^^^^^^^^^^^^^^^^^^^^^

Demonstrates how to create a custom gradient.

Apply to Title
""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.area import Gradient as TitleGradient
        from ooodev.format.chart2.direct.title.area import ColorRange

        # ... other code
        title_grad = TitleGradient(
            chart_doc=chart_doc,
            step_count=64,
            style=GradientStyle.SQUARE,
            angle=45,
            grad_color=ColorRange(StandardColor.PURPLE_LIGHT2, StandardColor.BLUE_DARK1),
        )
        Chart2.style_title(chart_doc=chart_doc, styles=[title_grad])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`21e3cc9d-2847-4118-9191-f85efb21c3eb` and :numref:`f8109112-6d3f-4852-ad7c-5fcbb22db31d`.


.. cssclass:: screen_shot

    .. _21e3cc9d-2847-4118-9191-f85efb21c3eb:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/21e3cc9d-2847-4118-9191-f85efb21c3eb
        :alt: Chart with gradient Title
        :figclass: align-center
        :width: 450px

        Chart with gradient Title

.. cssclass:: screen_shot

    .. _f8109112-6d3f-4852-ad7c-5fcbb22db31d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f8109112-6d3f-4852-ad7c-5fcbb22db31d
        :alt: Chart Title Gradient Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Gradient Dialog

Apply to Subtitle
""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[title_grad])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`d3e8fa2f-ba8e-4786-8518-2d958214cc55`.


.. cssclass:: screen_shot

    .. _d3e8fa2f-ba8e-4786-8518-2d958214cc55:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d3e8fa2f-ba8e-4786-8518-2d958214cc55
        :alt: Chart with gradient Title
        :figclass: align-center
        :width: 450px

        Chart with gradient Title

Image
-----

The :py:class:`ooodev.format.chart2.direct.title.area.Img` class is used to set the background image of the Title and Subtitle.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Apply image
^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` enum is used to select an image preset.

Apply to Title
""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.area import Img as TitleImg, PresetImageKind
        # ... other code

        title_img = TitleImg.from_preset(chart_doc, PresetImageKind.SPACE)
        Chart2.style_title(chart_doc=chart_doc, styles=[title_img])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`2aa029b8-a3b6-4b66-a069-b8585cedab3d` and :numref:`5dc5ef1b-229f-4256-9b3a-902a911bd7bf`.

.. cssclass:: screen_shot

    .. _2aa029b8-a3b6-4b66-a069-b8585cedab3d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2aa029b8-a3b6-4b66-a069-b8585cedab3d
        :alt: Chart with Title Image
        :figclass: align-center
        :width: 450px

        Chart with Title Image

.. cssclass:: screen_shot

    .. _5dc5ef1b-229f-4256-9b3a-902a911bd7bf:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/5dc5ef1b-229f-4256-9b3a-902a911bd7bf
        :alt: Chart Title Image Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Image Dialog

Apply to Subtitle
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[title_img])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`6d7e1b11-40e2-4e36-9eec-e9c97c716eca`.

.. cssclass:: screen_shot

    .. _6d7e1b11-40e2-4e36-9eec-e9c97c716eca:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/6d7e1b11-40e2-4e36-9eec-e9c97c716eca
        :alt: Chart with Title Image
        :figclass: align-center
        :width: 450px

        Chart with Title Image

Pattern
-------

The :py:class:`ooodev.format.chart2.direct.title.area.Pattern` class is used to set the background pattern of a Chart.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` enum is used to select a pattern preset.

Apply to Title
^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.area import Pattern as TitlePattern, PresetPatternKind
        # ... other code

        title_pattern = TitlePattern.from_preset(chart_doc, PresetPatternKind.HORIZONTAL_BRICK)
        Chart2.style_title(chart_doc=chart_doc, styles=[title_pattern])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`4cd109ba-6d3e-4dca-9754-82f6f24ce175` and :numref:`dd96600c-8960-423c-92ad-94b0bebb89c8`.


.. cssclass:: screen_shot

    .. _4cd109ba-6d3e-4dca-9754-82f6f24ce175:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4cd109ba-6d3e-4dca-9754-82f6f24ce175
        :alt: Chart with Title pattern
        :figclass: align-center
        :width: 450px

        Chart with Title pattern

.. cssclass:: screen_shot

    .. _dd96600c-8960-423c-92ad-94b0bebb89c8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/dd96600c-8960-423c-92ad-94b0bebb89c8
        :alt: Chart Title Pattern Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Pattern Dialog

Apply to Subtitle
^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.area import Pattern as TitlePattern, PresetPatternKind
        # ... other code

        title_pattern = TitlePattern.from_preset(chart_doc, PresetPatternKind.HORIZONTAL_BRICK)
        Chart2.style_title(chart_doc=chart_doc, styles=[title_pattern])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`eb494866-fb9c-46de-9231-a720a258cca2`.


.. cssclass:: screen_shot

    .. _eb494866-fb9c-46de-9231-a720a258cca2:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/eb494866-fb9c-46de-9231-a720a258cca2
        :alt: Chart with Subtitle pattern
        :figclass: align-center
        :width: 450px

        Chart with Subtitle pattern


Hatch
-----

The :py:class:`ooodev.format.chart2.direct.title.area.Hatch` class is used to set the Title and Subtitle hatch of a Chart.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` enum is used to select a hatch preset.

Apply to Title
^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.area import Hatch as TitleHatch, PresetHatchKind
        # ... other code

        title_hatch = TitleHatch.from_preset(chart_doc, PresetHatchKind.YELLOW_45_DEGREES_CROSSED)
        Chart2.style_title(chart_doc=chart_doc, styles=[title_hatch])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9876f022-0e42-4b5d-b07d-86c83f31e30c` and :numref:`dc35bca9-d365-4a04-a1e3-8e9c2db96d28`.

.. cssclass:: screen_shot

    .. _9876f022-0e42-4b5d-b07d-86c83f31e30c:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9876f022-0e42-4b5d-b07d-86c83f31e30c
        :alt: Chart with Title hatch
        :figclass: align-center
        :width: 450px

        Chart with Title hatch

.. cssclass:: screen_shot

    .. _dc35bca9-d365-4a04-a1e3-8e9c2db96d28:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/dc35bca9-d365-4a04-a1e3-8e9c2db96d28
        :alt: Chart Title Hatch Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Hatch Dialog

Apply to Subtitle
^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        Chart2.style_subtitle(chart_doc=chart_doc, styles=[title_hatch])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`7254600c-c53c-4219-91e5-aaa486709dae`.

.. cssclass:: screen_shot

    .. _7254600c-c53c-4219-91e5-aaa486709dae:

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
        - :ref:`help_chart2_format_direct_title_area`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_general_area`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>`
        - :py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.title.area.Color`
        - :py:class:`ooodev.format.chart2.direct.title.area.Gradient`
        - :py:class:`ooodev.format.chart2.direct.title.area.Img`
        - :py:class:`ooodev.format.chart2.direct.title.area.Pattern`
        - :py:class:`ooodev.format.chart2.direct.title.area.Hatch`