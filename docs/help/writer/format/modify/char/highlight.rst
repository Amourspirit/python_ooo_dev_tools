.. _help_writer_format_modify_char_highlight:

Write Modify Character Highlight Class
======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.char.highlight.Highlight` class is used to modify the highlight value.
As seen in :numref:`234257852-37671950-e1eb-4f16-a87c-0bd00bb11248` by default no color is set for the ``Example`` style.


Before Settings

.. cssclass:: screen_shot

    .. _234257852-37671950-e1eb-4f16-a87c-0bd00bb11248:
    .. figure:: https://user-images.githubusercontent.com/4193389/234257852-37671950-e1eb-4f16-a87c-0bd00bb11248.png
        :alt: Writer dialog Character highlight default
        :figclass: align-center

        Writer dialog Character highlight default


Setting the font effects
------------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 14, 15

        from ooodev.format.writer.modify.char.highlight import Highlight, StyleCharKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.utils.color import StandardColor

        def main() -> int:
           with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                font_style = Highlight(color=StandardColor.YELLOW_LIGHT2, style_name=StyleCharKind.EXAMPLE)
                font_style.apply(doc)

                style_obj = Highlight.from_style(doc=doc, style_name=StyleCharKind.EXAMPLE)
                assert style_obj.prop_style_name == str(StyleCharKind.EXAMPLE)
                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

After applying the font highlight.

.. cssclass:: screen_shot

    .. _234259234-f3a57507-bd10-4256-8e08-090a8c4cdc6d:
    .. figure:: https://user-images.githubusercontent.com/4193389/234259234-f3a57507-bd10-4256-8e08-090a8c4cdc6d.png
        :alt: Writer dialog character highlight style changed
        :figclass: align-center

        Writer dialog character highlight style changed


Getting the highlight from a style
----------------------------------

We can get the highlight from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Highlight.from_style(doc=doc, style_name=StyleCharKind.EXAMPLE)
        assert style_obj.prop_style_name == str(StyleCharKind.EXAMPLE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None



.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_char_highlight`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.char.highlight.Highlight`