.. _help_writer_format_modify_page_header_area:

Write Modify Page Header Area
=============================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The following classes are used to modify the Area style values seen in :numref:`ss_page_header_default_dialog` of a Page style.

- :py:class:`ooodev.format.writer.modify.page.header.area.Color`
- :py:class:`ooodev.format.writer.modify.page.header.area.Gradient`
- :py:class:`ooodev.format.writer.modify.page.header.area.Img`
- :py:class:`ooodev.format.writer.modify.page.header.area.Pattern`
- :py:class:`ooodev.format.writer.modify.page.header.area.Hatch`

Default Page Style Dialog

Setup
-----

General function used to run these examples.

Note that in order to apply a style, the document header must be turned on as seen in :ref:`help_writer_format_modify_page_header_header`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.header import Header, WriterStylePageKind
        from ooodev.format.writer.modify.page.header.area import Color as PageAreaColor
        from ooodev.format import Styler
        from ooodev.office.write import Write
        from ooodev.utils.color import StandardColor
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
           with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                header_style = Header(
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

The :py:class:`~ooodev.format.writer.modify.page.header.area.Color` class is used to modify the header area color of a page style.
The result are seen in :numref:`235274358-2ee74e38-d41c-44b1-bb47-b9a3b9dca5b1` and :numref:`235274417-3f4ed6c4-bc94-4f06-a15d-c4007af86332`.

Setting Area Color
^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.header.area import Color as HeaderAreaColor
        # ... other code

        header_color_style = HeaderAreaColor(
            color=StandardColor.GOLD_LIGHT2, style_name=header_style.prop_style_name
        )
        Styler.apply(doc, header_style, header_color_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235274358-2ee74e38-d41c-44b1-bb47-b9a3b9dca5b1:
    .. figure:: https://user-images.githubusercontent.com/4193389/235274358-2ee74e38-d41c-44b1-bb47-b9a3b9dca5b1.png
        :alt: Writer Page Header
        :figclass: align-center
        :width: 520px

        Writer Page Header
    
    .. _235274417-3f4ed6c4-bc94-4f06-a15d-c4007af86332:
    .. figure:: https://user-images.githubusercontent.com/4193389/235274417-3f4ed6c4-bc94-4f06-a15d-c4007af86332.png
        :alt: Writer dialog Header Area style color set
        :figclass: align-center
        :width: 450px

        Writer dialog Header Area style color set

Getting color from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = HeaderAreaColor.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Area Gradient
-------------

Setting Area Gradient
^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.header.area.Gradient` class is used to modify the header area gradient of a page style.
The result are seen in :numref:`235276610-b48373b4-19ad-4716-8432-e1424d524ca0` and :numref:`235276638-bfd94db4-3f13-424f-acb0-e23d7ba5521d`.

The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` class is used to look up the presets of gradient for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.header.area import Gradient, PresetGradientKind
        # ... other code

        gradient_style = Gradient.from_preset(
            preset=PresetGradientKind.DEEP_OCEAN, style_name=WriterStylePageKind.STANDARD
        )
        Styler.apply(doc, header_style, gradient_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235276610-b48373b4-19ad-4716-8432-e1424d524ca0:
    .. figure:: https://user-images.githubusercontent.com/4193389/235276610-b48373b4-19ad-4716-8432-e1424d524ca0.png
        :alt: Writer Page Header
        :figclass: align-center
        :width: 520px

        Writer Page Header

    .. _235276638-bfd94db4-3f13-424f-acb0-e23d7ba5521d:
    .. figure:: https://user-images.githubusercontent.com/4193389/235276638-bfd94db4-3f13-424f-acb0-e23d7ba5521d.png
        :alt: Writer dialog Header Area style gradient set
        :figclass: align-center
        :width: 450px

        Writer dialog Header Area style gradient set

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

The :py:class:`~ooodev.format.writer.modify.page.area.Img` class is used to modify the header area image of a page style.
The result are seen in :numref:`235276938-69941d73-2edd-43e7-bc2a-047ca31d73fa` and :numref:`235276967-1409f709-7725-44fa-a290-cb719d6f5850`.

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` class is used to look up the presets of image for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.header.area import Img as HeaderAreaImg, PresetImageKind
        # ... other code

        img_style = HeaderAreaImg.from_preset(
            preset=PresetImageKind.COLOR_STRIPES, style_name=WriterStylePageKind.STANDARD
        )
        Styler.apply(doc, header_style, img_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235276938-69941d73-2edd-43e7-bc2a-047ca31d73fa:
    .. figure:: https://user-images.githubusercontent.com/4193389/235276938-69941d73-2edd-43e7-bc2a-047ca31d73fa.png
        :alt: Writer Page Header
        :figclass: align-center
        :width: 520px

        Writer Page Header

    .. _235276967-1409f709-7725-44fa-a290-cb719d6f5850:
    .. figure:: https://user-images.githubusercontent.com/4193389/235276967-1409f709-7725-44fa-a290-cb719d6f5850.png
        :alt: Writer dialog Header Area style image set
        :figclass: align-center
        :width: 450px

        Writer dialog Header Area style image set

Getting image from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = HeaderAreaImg.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Area Pattern
------------

Setting Area Pattern
^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.area.Pattern` class is used to modify the header area pattern of a page style.
The result are seen in :numref:`235277296-0de4eda4-41aa-403d-9c6f-649dbdea3af7` and :numref:`235277323-cbefe390-bd71-4b3c-97c8-29db5ecf45d5`.

The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` class is used to look up the presets of pattern for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.header.area import Pattern as HeaderStylePattern, PresetPatternKind
        # ... other code

        pattern_style = HeaderStylePattern.from_preset(
            preset=PresetPatternKind.HORIZONTAL_BRICK, style_name=WriterStylePageKind.STANDARD
        )
        Styler.apply(doc, header_style, pattern_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235277296-0de4eda4-41aa-403d-9c6f-649dbdea3af7:
    .. figure:: https://user-images.githubusercontent.com/4193389/235277296-0de4eda4-41aa-403d-9c6f-649dbdea3af7.png
        :alt: Writer Page Header
        :figclass: align-center
        :width: 520px

        Writer Page Header

    .. _235277323-cbefe390-bd71-4b3c-97c8-29db5ecf45d5:
    .. figure:: https://user-images.githubusercontent.com/4193389/235277323-cbefe390-bd71-4b3c-97c8-29db5ecf45d5.png
        :alt: Writer dialog Header Area style pattern set
        :figclass: align-center
        :width: 450px

        Writer dialog Header Area style pattern set

Getting pattern from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = HeaderStylePattern.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Area Hatch
----------

Setting Area Hatch
^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.area.Hatch` class is used to modify the header area hatch of a page style.
The result are seen in :numref:`235277562-f68ac8b1-22a5-4474-8ba9-1c7b7b03c68c` and :numref:`235277592-c150738e-6fae-43c8-89f0-a43ae19eb99a`.

The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` class is used to look up the presets of hatch for convenience.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.header.area import Hatch as HeaderStyleHatch, PresetHatchKind
        # ... other code

        hatch_style = HeaderStyleHatch.from_preset(
            preset=PresetHatchKind.RED_45_DEGREES_NEG_TRIPLE, style_name=WriterStylePageKind.STANDARD
        )
        Styler.apply(doc, header_style, hatch_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235277562-f68ac8b1-22a5-4474-8ba9-1c7b7b03c68c:
    .. figure:: https://user-images.githubusercontent.com/4193389/235277562-f68ac8b1-22a5-4474-8ba9-1c7b7b03c68c.png
        :alt: Writer Page Header
        :figclass: align-center
        :width: 520px

        Writer Page Header

    .. _235277592-c150738e-6fae-43c8-89f0-a43ae19eb99a:
    .. figure:: https://user-images.githubusercontent.com/4193389/235277592-c150738e-6fae-43c8-89f0-a43ae19eb99a.png
        :alt: Writer dialog Header Area style hatch set
        :figclass: align-center
        :width: 450px

        Writer dialog Header Area style hatch set

Getting hatch from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = HeaderStyleHatch.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
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
        - :ref:`help_writer_format_modify_page_footer_area`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.writer.modify.page.header.area.Color`
        - :py:class:`ooodev.format.writer.modify.page.header.area.Gradient`
        - :py:class:`ooodev.format.writer.modify.page.header.area.Img`
        - :py:class:`ooodev.format.writer.modify.page.header.area.Pattern`
        - :py:class:`ooodev.format.writer.modify.page.header.area.Hatch`