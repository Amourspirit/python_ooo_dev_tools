.. _help_writer_format_direct_para_area_hatch:

Write Direct Paragraph Area Hatch Class
=======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Writer has an Area Hatch dialog section.

The :py:class:`ooodev.format.writer.direct.para.area.Hatch` class is used to set the paragraph background hatch.

.. cssclass:: screen_shot

    .. _ss_writer_dialog_para_area_hatch:
    .. figure:: https://user-images.githubusercontent.com/4193389/233857800-c87cbe3f-3228-48ed-b3ab-122c6f0302d0.png
        :alt: Writer Paragraph Area Hatch dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Area Hatch dialog

There are many presets for hatch available.
Using the :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` enum, you can set the hatch to one of the presets.

These examples use the :py:class:`ooodev.theme.ThemeTextDoc` class and the :py:class:`ooodev.utils.color.RGB` to determine the font color for the paragraph text.
This is done to ensure that the text is readable.

Setup
-----

General function used to run these examples:

.. must be before the tabs directive
.. include:: ../../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            from typing import TYPE_CHECKING, cast
            from ooodev.format.writer.direct.char.font import Font
            from ooodev.format.writer.direct.para.area import Hatch, PresetHatchKind
            from ooodev.office.write import Write
            from ooodev.theme import ThemeTextDoc
            from ooodev.utils.color import StandardColor, RGB
            from ooodev.utils.gui import GUI
            from ooodev.utils.lo import Lo
            
            if TYPE_CHECKING:
                from com.sun.star.text import TextRangeContentProperties  # service


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

                    theme_doc = ThemeTextDoc()
                    doc_color = RGB.from_int(theme_doc.doc_color)
                    if doc_color.is_dark():
                        font_color = StandardColor.WHITE
                    else:
                        font_color = StandardColor.BLACK

                    doc_font_style = Font(color=font_color)
                    doc_font_style.apply(cursor)

                    hatch_style = Hatch.from_preset(PresetHatchKind.YELLOW_45_DEGREES_CROSSED)
                    Write.append_para(cursor=cursor, text=p_txt, styles=[hatch_style])
                return 0


            if __name__ == "__main__":
                SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Examples
--------

Apply Hatch
^^^^^^^^^^^

Create a hatch for the paragraph background.

.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)
        # ... other code
        hatch_style = Hatch.from_preset(PresetHatchKind.YELLOW_45_DEGREES_CROSSED)
        Write.append_para(cursor=cursor, text=p_txt, styles=[hatch_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _233862027-c9b28ec0-4022-4592-984a-0ee8011bfde6:
    .. figure:: https://user-images.githubusercontent.com/4193389/233862027-c9b28ec0-4022-4592-984a-0ee8011bfde6.png
        :alt: Writer Paragraph Area Hatch
        :figclass: align-center
        :width: 520px

        Writer Paragraph Area Hatch

Get Hatch from Paragraph
^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        # ... other code

        para_cursor = Write.get_paragraph_cursor(cursor)
        para_cursor.gotoPreviousParagraph(False)
        para_cursor.gotoEndOfParagraph(True)

        text_para = cast("TextRangeContentProperties", para_cursor)
        para_hatch = Hatch.from_obj(text_para.TextParagraph)
        assert para_hatch.prop_name == str(PresetHatchKind.YELLOW_45_DEGREES_CROSSED)

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
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.area.Hatch`