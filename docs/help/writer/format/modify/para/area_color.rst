.. _help_writer_format_modify_para_color:

Write Modify Paragraph Area Color
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.area.Color` class is used to modify the values seen in :numref:`234414996-29658481-e0e6-4d6f-9f46-bd645e21da64` of a paragraph style.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 13, 14, 15, 15, 16

        from ooodev.format.writer.modify.para.area import Color as ParaStyleColor, StyleParaKind
        from ooodev.office.write import Write
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_area_color_style = ParaStyleColor(
                    color=StandardColor.GREEN_LIGHT2, style_name=StyleParaKind.STANDARD
                )
                para_area_color_style.apply(doc)

                style_obj = ParaStyleColor.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply alignment to a style
--------------------------

Before applying Style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234414996-29658481-e0e6-4d6f-9f46-bd645e21da64:

    .. figure:: https://user-images.githubusercontent.com/4193389/234414996-29658481-e0e6-4d6f-9f46-bd645e21da64.png
        :alt: Writer dialog Paragraph Area Color style default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Area Color style default

Apply style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_area_color_style = ParaStyleColor(
            color=StandardColor.GREEN_LIGHT2, style_name=StyleParaKind.STANDARD
        )
        para_area_color_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


After applying style
^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234415852-4f17c6b9-0379-445f-83a5-d8c3c184beeb:

    .. figure:: https://user-images.githubusercontent.com/4193389/234415852-4f17c6b9-0379-445f-83a5-d8c3c184beeb.png
        :alt: Writer dialog Paragraph Area Color style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Area Color style changed


Getting the area color from a style
-----------------------------------

We can get the area color from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = ParaStyleColor.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_para_area_color`
        - :ref:`help_writer_format_modify_page_area`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.writer.modify.para.area.Color`