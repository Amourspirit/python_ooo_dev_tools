.. _help_writer_format_modify_para_alignment:

Write Modify Paragraph Alignment
================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.alignment.Alignment` class is used to modify the values seen in :numref:`234373498-bc618e71-6d0d-42fe-abac-d8c2437165bc` of a paragraph style.

Setup
-----

General function used to run these examples:

.. must be before the tabs directive
.. include:: ../../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 16, 17, 18, 19, 20, 21, 22, 23, 24

        from ooodev.format.writer.modify.para.alignment import Alignment, StyleParaKind
        from ooodev.format.writer.modify.para.alignment import ParagraphAdjust, ParagraphVertAlignEnum
        from ooodev.format.writer.modify.para.alignment import WritingMode, WritingMode2Enum
        from ooodev.format.writer.modify.para.alignment import LastLineKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                shadow_style = Alignment(
                    align=ParagraphAdjust.BLOCK,
                    align_vert=ParagraphVertAlignEnum.TOP,
                    txt_direction=WritingMode(WritingMode2Enum.LR_TB),
                    expand_single_word=True,
                    align_last=LastLineKind.JUSTIFY,
                    style_name=StyleParaKind.STANDARD,
                )
                shadow_style.apply(doc)

                cursor = Write.get_cursor(doc)
                Write.append_para(cursor=cursor, text=p_txt)

                style_obj = Alignment.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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

    .. _234373498-bc618e71-6d0d-42fe-abac-d8c2437165bc:

    .. figure:: https://user-images.githubusercontent.com/4193389/234373498-bc618e71-6d0d-42fe-abac-d8c2437165bc.png
        :alt: Writer dialog Paragraph Alignmet style default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Alignmet style default

Apply style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        shadow_style = Alignment(
            align=ParagraphAdjust.BLOCK,
            align_vert=ParagraphVertAlignEnum.TOP,
            txt_direction=WritingMode(WritingMode2Enum.LR_TB),
            expand_single_word=True,
            align_last=LastLineKind.JUSTIFY,
            style_name=StyleParaKind.STANDARD,
        )
        shadow_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


After appling style
^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234381289-24adf761-cf5b-4bbb-ba53-4f855b1a3232:

    .. figure:: https://user-images.githubusercontent.com/4193389/234381289-24adf761-cf5b-4bbb-ba53-4f855b1a3232.png
        :alt: Writer dialog Paragraph Alignmet style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Alignmet style changed


Getting the alignment from a style
----------------------------------

We can get the alignment from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Alignment.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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
        - :ref:`help_writer_format_direct_para_alignment`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.alignment.Alignment`