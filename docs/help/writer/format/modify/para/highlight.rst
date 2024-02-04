.. _help_writer_format_modify_para_highlight:

Write Modify Paragraph Highlight
================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.highlight.Highlight` class is used to modify the values seen in :numref:`234734093-1997fb9d-1b50-4c0e-bd75-5f2f2b04409a` of a paragraph style.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 14,15

        from ooodev.format.writer.modify.para.highlight import Highlight, StyleParaKind
        from ooodev.utils.color import StandardColor
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_hl_style = Highlight(color=StandardColor.YELLOW_LIGHT3, style_name=StyleParaKind.STANDARD)
                para_hl_style.apply(doc)

                style_obj = Highlight.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply Highlight to a style
--------------------------

Before applying Style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234734093-1997fb9d-1b50-4c0e-bd75-5f2f2b04409a:

    .. figure:: https://user-images.githubusercontent.com/4193389/234734093-1997fb9d-1b50-4c0e-bd75-5f2f2b04409a.png
        :alt: Writer dialog Paragraph Area Color style default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Highlighting style default

Apply style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_hl_style = Highlight(color=StandardColor.YELLOW_LIGHT3, style_name=StyleParaKind.STANDARD)
        para_hl_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


After applying style
^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234734961-a056b02d-56e1-4da0-8854-e9bf002b001f:

    .. figure:: https://user-images.githubusercontent.com/4193389/234734961-a056b02d-56e1-4da0-8854-e9bf002b001f.png
        :alt: Writer dialog Paragraph Highlight Color style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Highlight Color style changed


Getting the highlight color from a style
----------------------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Highlight.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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
        - :ref:`help_writer_format_direct_char_highlight`
        - :ref:`help_writer_format_modify_char_highlight`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.highlight.Highlight`