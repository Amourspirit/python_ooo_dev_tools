.. _help_writer_format_direct_para_transparency:

Write Direct Paragraph Transparency
===================================

.. This file can also include Gradient if and when it is implemented.

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Writer has an Transparency dialog section.

The :py:class:`ooodev.format.writer.direct.para.transparency.Transparency` class is used to set the paragraph background transparency.

.. cssclass:: screen_shot

    .. _ss_writer_dialog_para_transparency:
    .. figure:: https://user-images.githubusercontent.com/4193389/233992459-220ac38d-cb9e-4c5f-8321-82a65224e839.png
        :alt: Writer Paragraph Transparency dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Transparency dialog


Setup
-----

General function used to run these examples:

.. must be before the tabs directive
.. include:: ../../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            from ooodev.format.writer.direct.para.area import Color
            from ooodev.format.writer.direct.para.transparency import Transparency
            from ooodev.office.write import Write
            from ooodev.utils.gui import GUI
            from ooodev.utils.lo import Lo
            from ooodev.utils.color import StandardColor


            def main() -> int:
                p_txt = (
                    |short_ptext|
                )

                with Lo.Loader(Lo.ConnectPipe()):
                    doc = Write.create_doc()
                    GUI.set_visible(doc)
                    Lo.delay(300)
                    GUI.zoom(GUI.ZoomEnum.ZOOM_100_PERCENT)
                    cursor = Write.get_cursor(doc)

                    color_style = Color(StandardColor.LIME)
                    t_style = Transparency(52)

                    Write.append_para(cursor=cursor, text=p_txt, styles=[color_style, t_style])

                    para_cursor = Write.get_paragraph_cursor(cursor)
                    para_cursor.gotoPreviousParagraph(False)
                    para_cursor.gotoEndOfParagraph(True)

                    text_para = cast("TextRangeContentProperties", para_cursor)

                    para_t = Transparency.from_obj(text_para.TextParagraph)
                    assert para_t is not None

                    para_cursor.gotoEnd(False)
                    Lo.delay(1_000)
                    Lo.close_doc(doc)
                return 0


            if __name__ == "__main__":
                SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Examples
--------

Apply Transparency to Paragraph
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

In this example we will apply a transparency to a paragraph background color.
The transparency needs to be applied after the paragraph color as the transparency is applied to the color.
This means the order ``[color_style, t_style]`` is important.
The transparency is set to 52% in this example.

.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)

        color_style = Color(StandardColor.LIME)
        t_style = Transparency(52)

        Write.append_para(cursor=cursor, text=p_txt, styles=[color_style, t_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _233993401-2201b7eb-b28e-4f57-923c-ca0bc30d28b4:
    .. figure:: https://user-images.githubusercontent.com/4193389/233993401-2201b7eb-b28e-4f57-923c-ca0bc30d28b4.png
        :alt: Writer Paragraph Color and Transparency
        :figclass: align-center
        :width: 520px

        Writer Paragraph Color and Transparency

Get Transparency from Paragraph
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        # ... other code

        para_cursor = Write.get_paragraph_cursor(cursor)
        para_cursor.gotoPreviousParagraph(False)
        para_cursor.gotoEndOfParagraph(True)

        text_para = cast("TextRangeContentProperties", para_cursor)

        para_t = Transparency.from_obj(text_para.TextParagraph)
        assert para_t is not None

        para_cursor.gotoEnd(False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

   .. cssclass:: ul-list

        - :ref:`help_writer_format_style_para_reset_default`
        - :ref:`help_writer_format_style_para`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_para_transparency`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.transparency.Transparency`