.. _help_writer_format_modify_page_footer_transparency:

Write Modify Page Footer Transparency
=====================================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.page.footer.transparency.Transparency` and :py:class:`ooodev.format.writer.modify.page.footer.transparency.Gradient` classes are used to modify the Area style values seen in :numref:`235732260-413041b5-2822-4512-9775-bd97688b87e1_2` of a Page style.


.. cssclass:: screen_shot

    .. _235732260-413041b5-2822-4512-9775-bd97688b87e1_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235732260-413041b5-2822-4512-9775-bd97688b87e1.png
        :alt: Writer dialog Footer Transparency default
        :figclass: align-center
        :width: 450px

        Writer dialog Footer Transparency default

Default Page Style Dialog

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.footer import Footer, WriterStylePageKind
        from ooodev.format.writer.modify.page.footer.transparency import (
            Transparency,
            Gradient,
            IntensityRange,
            GradientStyle,
        )
        from ooodev.format.writer.modify.page.footer.area import Color as HeaderAreaColor
        from ooodev.format import Styler
        from ooodev.utils.color import StandardColor
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
           with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                footer_style = Footer(
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
                page_footer_style_kind = WriterStylePageKind.STANDARD
                color_style = HeaderAreaColor(color=StandardColor.RED, style_name=page_footer_style_kind)
                transparency_style = Transparency(value=85, style_name=page_footer_style_kind)
                Styler.apply(doc, footer_style, color_style, transparency_style)

                style_obj = Transparency.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
                assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Transparency
------------

The :py:class:`~ooodev.format.writer.modify.page.footer.transparency.Transparency` class is used to modify the transparency of a page footer style.
The result are seen in :numref:`235745123-48b822ef-09f6-47da-b3e1-f23b7a8cd019` and :numref:`235739497-aed8fad2-ba01-4bbc-abfa-5996d0d0ea71_2`.

Setting Transparency
^^^^^^^^^^^^^^^^^^^^

In this example we will apply a transparency to the page footer style background color.
The transparency needs to be applied after the page footer style color as the transparency is applied to the color.
This means the order ``Styler.apply(doc, footer_style, color_style, transparency_style)`` is important.
The transparency is set to 85% in this example.

.. tabs::

    .. code-tab:: python

        # ... other code

        page_footer_style_kind = WriterStylePageKind.STANDARD
        color_style = HeaderAreaColor(color=StandardColor.RED, style_name=page_footer_style_kind)
        transparency_style = Transparency(value=85, style_name=page_footer_style_kind)
        Styler.apply(doc, footer_style, color_style, transparency_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235745123-48b822ef-09f6-47da-b3e1-f23b7a8cd019:
    .. figure:: https://user-images.githubusercontent.com/4193389/235745123-48b822ef-09f6-47da-b3e1-f23b7a8cd019.png
        :alt: Writer Page Footer
        :figclass: align-center
        :width: 520px

        Writer Page Footer

    .. _235739497-aed8fad2-ba01-4bbc-abfa-5996d0d0ea71_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235739497-aed8fad2-ba01-4bbc-abfa-5996d0d0ea71.png
        :alt: Writer dialog Page Footer Transparency style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page Footer Transparency style changed

Getting transparency from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Transparency.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Transparency Gradient
---------------------

Setting Transparency Gradient
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.footer.transparency.Gradient` class is used to modify the area gradient of a page footer style.
The result are seen in :numref:`235745677-609ead3c-03eb-4ce4-af18-c7dda60b913c` and :numref:`235742293-942b97ad-2455-4c12-9749-529658010add_2`.

In this example we will apply a transparency to the page footer style background color.
The transparency needs to be applied after the page footer style color as the transparency is applied to the color.
This means the order ``Styler.apply(doc, footer_style, color_style, footer_gradient_style)`` is important.

.. tabs::

    .. code-tab:: python

        # ... other code

        page_footer_style_kind = WriterStylePageKind.STANDARD
        color_style = HeaderAreaColor(color=StandardColor.GREEN_DARK1, style_name=page_footer_style_kind)
        footer_gradient_style = Gradient(
            style=GradientStyle.LINEAR,
            angle=45,
            border=22,
            grad_intensity=IntensityRange(0, 100),
            style_name=page_footer_style_kind,
        )
        Styler.apply(doc, footer_style, color_style, footer_gradient_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235745677-609ead3c-03eb-4ce4-af18-c7dda60b913c:
    .. figure:: https://user-images.githubusercontent.com/4193389/235745677-609ead3c-03eb-4ce4-af18-c7dda60b913c.png
        :alt: Writer Page Footer
        :figclass: align-center
        :width: 520px

        Writer Page Footer

    .. _235742293-942b97ad-2455-4c12-9749-529658010add_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235742293-942b97ad-2455-4c12-9749-529658010add.png
        :alt: Writer dialog Page Footer Transparency style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page Footer Transparency style changed

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

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_page_header_transparency`
        - :ref:`help_writer_format_modify_page_transparency`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.page.footer.transparency.Transparency`
        - :py:class:`ooodev.format.writer.modify.page.footer.transparency.Gradient`