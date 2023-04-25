.. _help_writer_format_direct_para_area_gradient:

Write Direct Paragraph Area Gradient Class
==========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Writer has an Area Gradient dialog section.

The :py:class:`ooodev.format.writer.direct.para.area.Gradient` class is used to set the paragraph background gradient.

.. cssclass:: screen_shot

    .. _ss_writer_dialog_para_area_gradient:
    .. figure:: https://user-images.githubusercontent.com/4193389/233850124-2fc250c2-efb9-40b5-b6d9-278e389bfe58.png
        :alt: Writer Paragraph Area Gradient dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Area Gradient dialog

There are many presets for gradient available.
Using the :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` enum, you can set the gradient to one of the presets.

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
            from ooodev.format.writer.direct.para.area import Gradient, PresetGradientKind
            from ooodev.office.write import Write
            from ooodev.utils.gui import GUI
            from ooodev.utils.lo import Lo
            
            if TYPE_CHECKING:
                from com.sun.star.text import TextRangeContentProperties  # service


            def main() -> int:
                p_txt = (
                    |short_ptext|
                )

                with Lo.Loader(Lo.ConnectSocket()):
                    doc = Write.create_doc()
                    GUI.set_visible(doc)
                    Lo.delay(300)
                    GUI.zoom(GUI.ZoomEnum.ZOOM_100_PERCENT)
                    cursor = Write.get_cursor(doc)

                    gradient_style = Gradient.from_preset(PresetGradientKind.MAHOGANY)
                    Write.append_para(cursor=cursor, text=p_txt, styles=[gradient_style])

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

Apply Gradient Color
^^^^^^^^^^^^^^^^^^^^

Create a gradient for the paragraph background.

.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)
        gradient_style = Gradient.from_preset(PresetGradientKind.MAHOGANY)
        Write.append_para(cursor=cursor, text=p_txt, styles=[gradient_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _233850291-c58c312c-8d6e-4a7f-b6ca-51d1c307622d:
    .. figure:: https://user-images.githubusercontent.com/4193389/233850291-c58c312c-8d6e-4a7f-b6ca-51d1c307622d.png
        :alt: Writer Paragraph Area Gradient
        :figclass: align-center
        :width: 520px

        Writer Paragraph Area Gradient

Get Gradient from Paragraph
^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        gradient_style = Gradient.from_preset(PresetGradientKind.MAHOGANY)
        Write.append_para(cursor=cursor, text=p_txt, styles=[gradient_style])

        para_cursor = Write.get_paragraph_cursor(cursor)
        para_cursor.gotoPreviousParagraph(False)
        para_cursor.gotoEndOfParagraph(True)

        text_para = cast("TextRangeContentProperties", para_cursor)

        para_gradient = Gradient.from_obj(text_para.TextParagraph)
        para_gradient.prop_name == str(PresetGradientKind.MAHOGANY)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

   .. cssclass:: ul-list

        - :ref:`help_writer_format_style_para_reset_default`
        - :ref:`help_writer_format_style_para`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_para_gradient`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.area.Gradient`