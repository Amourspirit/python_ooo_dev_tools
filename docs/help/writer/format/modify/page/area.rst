.. _help_writer_format_modify_page_area:

Write Modify Page Area
======================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The following classes are used to modify the Area style values seen in :numref:`235159060-e1c81443-a0c3-4ddf-a1f7-171ae633db89` of a Page style.

- :py:class:`ooodev.format.writer.modify.page.area.Color`
- :py:class:`ooodev.format.writer.modify.page.area.Gradient`
- :py:class:`ooodev.format.writer.modify.page.area.Img`
- :py:class:`ooodev.format.writer.modify.page.area.Pattern`
- :py:class:`ooodev.format.writer.modify.page.area.Hatch`

.. cssclass:: screen_shot

    .. _235159060-e1c81443-a0c3-4ddf-a1f7-171ae633db89:
    .. figure:: https://user-images.githubusercontent.com/4193389/235159060-e1c81443-a0c3-4ddf-a1f7-171ae633db89.png
        :alt: Writer dialog Area default
        :figclass: align-center
        :width: 450px

        Writer dialog Area default

Default Page Style Dialog

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.area import Color as PageAreaColor, WriterStylePageKind
        from ooodev.utils.color import StandardColor
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
           with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                color_style = PageAreaColor(
                    color=StandardColor.BLUE_LIGHT3, style_name=WriterStylePageKind.STANDARD
                )
                color_style.apply(doc)

                style_obj = PageAreaColor.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
                assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Area Color
----------

The :py:class:`~ooodev.format.writer.modify.page.area.Color` class is used to modify the area color of a page style.
The result are seen in :numref:`235160627-5e2c7367-481d-4465-9402-408f204f0156`.

Setting Area Color
^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.area import Color as PageAreaColor, WriterStylePageKind
        # ... other code

        color_style = PageAreaColor(color=StandardColor.BLUE_LIGHT3, style_name=WriterStylePageKind.STANDARD)
        color_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235160627-5e2c7367-481d-4465-9402-408f204f0156:
    .. figure:: https://user-images.githubusercontent.com/4193389/235160627-5e2c7367-481d-4465-9402-408f204f0156.png
        :alt: Writer dialog Area style color set
        :figclass: align-center
        :width: 450px

        Writer dialog Area style color set

Getting color from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = PageAreaColor.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Area Gradient
-------------

Setting Area Gradient
^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.area.Gradient` class is used to modify the area gradient of a page style.
The result are seen in :numref:`235162481-6df8e5aa-99d6-4271-bf41-6ebb76bd0dcf`.

The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` class is used to look up the presets of gradient for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.area import Gradient, PresetGradientKind, WriterStylePageKind
        # ... other code

        gradient_style = Gradient.from_preset(
            preset=PresetGradientKind.DEEP_OCEAN, style_name=WriterStylePageKind.STANDARD
        )
        gradient_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235162481-6df8e5aa-99d6-4271-bf41-6ebb76bd0dcf:
    .. figure:: https://user-images.githubusercontent.com/4193389/235162481-6df8e5aa-99d6-4271-bf41-6ebb76bd0dcf.png
        :alt: Writer dialog Area style gradient set
        :figclass: align-center
        :width: 450px

        Writer dialog Area style gradient set

Getting gradient from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Gradient.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Area Image
----------

Setting Area Image
^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.area.Img` class is used to modify the area image of a page style.
The result are seen in :numref:`235177592-45f5c000-3a01-4ab7-922c-baa0406efebd`.

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` class is used to look up the presets of image for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.area import Img as PageAreaImg, PresetImageKind, WriterStylePageKind
        # ... other code

        img_style = PageAreaImg.from_preset(
            preset=PresetImageKind.COLOR_STRIPES, style_name=WriterStylePageKind.STANDARD
        )
        img_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235177592-45f5c000-3a01-4ab7-922c-baa0406efebd:
    .. figure:: https://user-images.githubusercontent.com/4193389/235177592-45f5c000-3a01-4ab7-922c-baa0406efebd.png
        :alt: Writer dialog Area style image set
        :figclass: align-center
        :width: 450px

        Writer dialog Area style image set

Getting image from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = PageAreaImg .from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Area Pattern
------------

Setting Area Pattern
^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.area.Pattern` class is used to modify the area pattern of a page style.
The result are seen in :numref:`235178928-a1f82ee8-1224-4cbc-abee-de843c11c639`.

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` class is used to look up the presets of pattern for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.area import Pattern as PageStylePattern
        from ooodev.format.writer.modify.page.area import PresetPatternKind, WriterStylePageKind
        # ... other code

        pattern_style = PageStylePattern.from_preset(
            preset=PresetPatternKind.HORIZONTAL_BRICK, style_name=WriterStylePageKind.STANDARD
        )
        pattern_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235178928-a1f82ee8-1224-4cbc-abee-de843c11c639:
    .. figure:: https://user-images.githubusercontent.com/4193389/235178928-a1f82ee8-1224-4cbc-abee-de843c11c639.png
        :alt: Writer dialog Area style pattern set
        :figclass: align-center
        :width: 450px

        Writer dialog Area style pattern set

Getting pattern from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = PageStylePattern .from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Area Hatch
----------

Setting Area Hatch
^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.area.Hatch` class is used to modify the area hatch of a page style.
The result are seen in :numref:`235180945-3fdba1f7-8065-4cfa-8dfc-34ceeed0623a`.

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` class is used to look up the presets of hatch for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.area import Hatch as PageStyleHatch
        from ooodev.format.writer.modify.page.area import PresetHatchKind, WriterStylePageKind
        # ... other code

        hatch_style = PageStyleHatch.from_preset(
            preset=PresetHatchKind.RED_45_DEGREES_NEG_TRIPLE, style_name=WriterStylePageKind.STANDARD
        )
        hatch_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235180945-3fdba1f7-8065-4cfa-8dfc-34ceeed0623a:
    .. figure:: https://user-images.githubusercontent.com/4193389/235180945-3fdba1f7-8065-4cfa-8dfc-34ceeed0623a.png
        :alt: Writer dialog Area style hatch set
        :figclass: align-center
        :width: 450px

        Writer dialog Area style hatch set

Getting hatch from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = PageStyleHatch .from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.writer.modify.page.area.Color`
        - :py:class:`ooodev.format.writer.modify.page.area.Gradient`
        - :py:class:`ooodev.format.writer.modify.page.area.Img`
        - :py:class:`ooodev.format.writer.modify.page.area.Pattern`
        - :py:class:`ooodev.format.writer.modify.page.area.Hatch`