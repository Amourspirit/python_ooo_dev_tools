.. _help_writer_format_modify_page_footer_area:

Write Modify Page Footer Area
=============================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The following classes are used to modify the Area style values seen in :numref:`ss_page_footer_default_dialog` of a Page style.

- :py:class:`ooodev.format.writer.modify.page.footer.area.Color`
- :py:class:`ooodev.format.writer.modify.page.footer.area.Gradient`
- :py:class:`ooodev.format.writer.modify.page.footer.area.Img`
- :py:class:`ooodev.format.writer.modify.page.footer.area.Pattern`
- :py:class:`ooodev.format.writer.modify.page.footer.area.Hatch`

Default Page Style Dialog

Setup
-----

General function used to run these examples.

Note that in order to apply a style, the document footer must be turned on as seen in :ref:`help_writer_format_modify_page_footer_footer`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.area import Color as PageAreaColor, WriterStylePageKind
        from ooodev.format import Styler
        from ooodev.office.write import Write
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
           with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                header_style = Footer(
                    on=True,
                    shared_first=True,
                    shared=True,
                    height=10.0,
                    spacing=3.0,
                    spacing_dyn=True,
                    margin_left=1.5,
                    margin_right=2.0,
                    style_name=WriterStylePageKind.STANDARD,
                )
                header_color_style = HeaderAreaColor(
                    color=StandardColor.GOLD_LIGHT2, style_name=header_style.prop_style_name
                )
                Styler.apply(doc, header_style, header_color_style)

                style_obj = HeaderAreaColor.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
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

The :py:class:`~ooodev.format.writer.modify.page.footer.area.Color` class is used to modify the footer area color of a page style.
The result are seen in :numref:`235279409-ef77a3a6-239b-475f-9b66-a97684538b53` and :numref:`235274417-3f4ed6c4-bc94-4f06-a15d-c4007af86332_2`.

Setting Area Color
^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.footer.area import Color as FooterAreaColor, WriterStylePageKind
        # ... other code

        footer_color_style = FooterAreaColor(
            color=StandardColor.GOLD_LIGHT2, style_name=footer_style.prop_style_name
        )
        Styler.apply(doc, footer_style, footer_color_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235279409-ef77a3a6-239b-475f-9b66-a97684538b53:
    .. figure:: https://user-images.githubusercontent.com/4193389/235279409-ef77a3a6-239b-475f-9b66-a97684538b53.png
        :alt: Writer Page Footer
        :figclass: align-center
        :width: 520px

        Writer Page Footer
    
    .. _235274417-3f4ed6c4-bc94-4f06-a15d-c4007af86332_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235274417-3f4ed6c4-bc94-4f06-a15d-c4007af86332.png
        :alt: Writer dialog Footer Area style color set
        :figclass: align-center
        :width: 450px

        Writer dialog Footer Area style color set

Getting color from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = FooterAreaColor.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Area Gradient
-------------

Setting Area Gradient
^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.footer.area.Gradient` class is used to modify the footer area gradient of a page style.
The result are seen in :numref:`235279508-8549c510-ecc1-465f-a37d-3af99209ba95` and :numref:`235276638-bfd94db4-3f13-424f-acb0-e23d7ba5521d_2`.

The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` class is used to look up the presets of gradient for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.footer.area import Gradient, PresetGradientKind
        # ... other code

        gradient_style = Gradient.from_preset(
            preset=PresetGradientKind.DEEP_OCEAN, style_name=WriterStylePageKind.STANDARD
        )
        Styler.apply(doc, footer_style, gradient_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235279508-8549c510-ecc1-465f-a37d-3af99209ba95:
    .. figure:: https://user-images.githubusercontent.com/4193389/235279508-8549c510-ecc1-465f-a37d-3af99209ba95.png
        :alt: Writer Page Footer
        :figclass: align-center
        :width: 520px

        Writer Page Footer

    .. _235276638-bfd94db4-3f13-424f-acb0-e23d7ba5521d_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235276638-bfd94db4-3f13-424f-acb0-e23d7ba5521d.png
        :alt: Writer dialog Footer Area style gradient set
        :figclass: align-center
        :width: 450px

        Writer dialog Footer Area style gradient set

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

The :py:class:`~ooodev.format.writer.modify.page.area.Img` class is used to modify the footer area image of a page style.
The result are seen in :numref:`235279902-8b66dc92-f204-4ca3-9749-faad730ff368` and :numref:`235276967-1409f709-7725-44fa-a290-cb719d6f5850_2`.

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` class is used to look up the presets of image for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.footer.area import Img as FooterAreaImg, PresetImageKind
        # ... other code

        img_style = FooterAreaImg.from_preset(
            preset=PresetImageKind.COLOR_STRIPES, style_name=WriterStylePageKind.STANDARD
        )
        Styler.apply(doc, footer_style, img_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235279902-8b66dc92-f204-4ca3-9749-faad730ff368:
    .. figure:: https://user-images.githubusercontent.com/4193389/235279902-8b66dc92-f204-4ca3-9749-faad730ff368.png
        :alt: Writer Page Footer
        :figclass: align-center
        :width: 520px

        Writer Page Footer

    .. _235276967-1409f709-7725-44fa-a290-cb719d6f5850_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235276967-1409f709-7725-44fa-a290-cb719d6f5850.png
        :alt: Writer dialog Footer Area style image set
        :figclass: align-center
        :width: 450px

        Writer dialog Footer Area style image set

Getting image from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = FooterAreaImg.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Area Pattern
------------

Setting Area Pattern
^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.area.Pattern` class is used to modify the footer area pattern of a page style.
The result are seen in :numref:`235280087-5e384ced-5620-4ca1-9c56-635e48db6059` and :numref:`235277323-cbefe390-bd71-4b3c-97c8-29db5ecf45d5_2`.

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` class is used to look up the presets of pattern for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.footer.area import Pattern as FooterStylePattern, PresetPatternKind
        # ... other code

        pattern_style = FooterStylePattern.from_preset(
            preset=PresetPatternKind.HORIZONTAL_BRICK, style_name=WriterStylePageKind.STANDARD
        )
        Styler.apply(doc, footer_style, pattern_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235280087-5e384ced-5620-4ca1-9c56-635e48db6059:
    .. figure:: https://user-images.githubusercontent.com/4193389/235280087-5e384ced-5620-4ca1-9c56-635e48db6059.png
        :alt: Writer Page Footer
        :figclass: align-center
        :width: 520px

        Writer Page Footer

    .. _235277323-cbefe390-bd71-4b3c-97c8-29db5ecf45d5_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235277323-cbefe390-bd71-4b3c-97c8-29db5ecf45d5.png
        :alt: Writer dialog Footer Area style pattern set
        :figclass: align-center
        :width: 450px

        Writer dialog Footer Area style pattern set

Getting pattern from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = FooterStylePattern.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Area Hatch
----------

Setting Area Hatch
^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.area.Hatch` class is used to modify the footer area hatch of a page style.
The result are seen in :numref:`235279706-08675945-3de2-4510-ab35-44ea3c8d8023` and :numref:`235277592-c150738e-6fae-43c8-89f0-a43ae19eb99a_2`.

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` class is used to look up the presets of hatch for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.footer.area import Hatch as FooterStyleHatch, PresetHatchKind
        # ... other code

        hatch_style = FooterStyleHatch.from_preset(
            preset=PresetHatchKind.RED_45_DEGREES_NEG_TRIPLE, style_name=WriterStylePageKind.STANDARD
        )
        Styler.apply(doc, footer_style, hatch_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235279706-08675945-3de2-4510-ab35-44ea3c8d8023:
    .. figure:: https://user-images.githubusercontent.com/4193389/235279706-08675945-3de2-4510-ab35-44ea3c8d8023.png
        :alt: Writer Page Footer
        :figclass: align-center
        :width: 520px

        Writer Page Footer

    .. _235277592-c150738e-6fae-43c8-89f0-a43ae19eb99a_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235277592-c150738e-6fae-43c8-89f0-a43ae19eb99a.png
        :alt: Writer dialog Footer Area style hatch set
        :figclass: align-center
        :width: 450px

        Writer dialog Footer Area style hatch set

Getting hatch from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = FooterStyleHatch.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.page.footer.area.Color`
        - :py:class:`ooodev.format.writer.modify.page.footer.area.Gradient`
        - :py:class:`ooodev.format.writer.modify.page.footer.area.Img`
        - :py:class:`ooodev.format.writer.modify.page.footer.area.Pattern`
        - :py:class:`ooodev.format.writer.modify.page.footer.area.Hatch`