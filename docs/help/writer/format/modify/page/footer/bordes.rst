.. _help_writer_format_modify_page_footer_borders:

Write Modify Page Footer Borders
================================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.page.footer.borders.Sides`, :py:class:`ooodev.format.writer.modify.page.footer.borders.Padding`, and :py:class:`ooodev.format.writer.modify.page.footer.borders.Shadow`
classes are used to modify the border values seen in :numref:`235311718-e3290dbc-73b5-4555-ad52-5c9be611a7a9_2` of a character border style.


Default Page Footer Borders Style Dialog

.. cssclass:: screen_shot

    .. _235311718-e3290dbc-73b5-4555-ad52-5c9be611a7a9_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235311718-e3290dbc-73b5-4555-ad52-5c9be611a7a9.png
        :alt: Writer dialog Page Footer Borders default
        :figclass: align-center
        :width: 450px

        Writer dialog Page Footer Borders default


Setup
-----

General function used to run these examples.

Note that in order to apply a style, the document footer must be turned on as seen in :ref:`help_writer_format_modify_page_footer_footer`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.footer import Footer, WriterStylePageKind
        from ooodev.format.writer.modify.page.footer.borders import Padding, Shadow, Sides
        from ooodev.format.writer.modify.page.footer.borders import BorderLineKind, LineSize, Side
        from ooodev.format import Styler
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor

        def main() -> int:
           with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
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

                side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED, width=LineSize.MEDIUM)
                sides_style = Sides(all=side, style_name=WriterStylePageKind.STANDARD)
                Styler.apply(doc, footer_style, sides_style)

                style_obj = Sides.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
                assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Border Sides
------------

Setting Border Sides
^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED, width=LineSize.MEDIUM)
        sides_style = Sides(all=side, style_name=WriterStylePageKind.STANDARD)
        Styler.apply(doc, footer_style, sides_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235313652-3c8e5262-a004-4847-8eae-a52f746cbd8b:
    .. figure:: https://user-images.githubusercontent.com/4193389/235313652-3c8e5262-a004-4847-8eae-a52f746cbd8b.png
        :alt: Writer Page Footer
        :figclass: align-center
        :width: 520px

        Writer Page Footer

    .. _235312429-45a5fcdc-7edc-4d4c-8958-83242e999fce_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235312429-45a5fcdc-7edc-4d4c-8958-83242e999fce.png
        :alt: Writer dialog Page Footer Borders style sides changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page Footer Borders style sides changed


Getting border sides from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Sides.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Border Padding
--------------

Setting Border Padding
^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        padding_style = Padding(
            left=5, right=5, top=3, bottom=3, style_name=WriterStylePageKind.STANDARD
        )
        Styler.apply(doc, footer_style, padding_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235312730-65f07bdc-e354-46d7-816e-bd9826d55ab5_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235312730-65f07bdc-e354-46d7-816e-bd9826d55ab5.png
        :alt: Writer dialog Page Footer Borders style padding changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page Footer Borders style padding changed

Getting border padding from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border padding from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Padding.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Border Shadow
-------------

Setting Border Shadow
^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        shadow_style = Shadow(
            color=StandardColor.BLUE_DARK2, width=1.5, style_name=WriterStylePageKind.STANDARD
        )
        Styler.apply(doc, footer_style, shadow_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235312951-d5ce72eb-2993-4857-b7d0-8d19f003b574_2:
    .. figure:: https://user-images.githubusercontent.com/4193389/235312951-d5ce72eb-2993-4857-b7d0-8d19f003b574.png
        :alt: Writer dialog Page Footer Borders style shadow changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page Footer Borders style shadow changed

Getting border shadow from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border shadow from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Shadow.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
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
        - :ref:`help_writer_format_direct_para_borders`
        - :ref:`help_writer_format_modify_para_borders`
        - :ref:`help_writer_format_modify_page_borders`
        - :ref:`help_writer_format_modify_page_header_borders`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.page.footer.borders.Padding`
        - :py:class:`ooodev.format.writer.modify.page.footer.borders.Sides`
        - :py:class:`ooodev.format.writer.modify.page.footer.borders.Shadow`