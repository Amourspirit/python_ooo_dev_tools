.. _help_writer_format_direct_para_area_img:

Write Direct Paragraph Area Img Class
=====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Writer has an Area Hatch dialog section.

The :py:class:`ooodev.format.writer.direct.para.area.Img` class is used to set the paragraph background image.

.. cssclass:: screen_shot

    .. _ss_writer_dialog_para_area_image:
    .. figure:: https://user-images.githubusercontent.com/4193389/233866224-5c75d4bd-6916-46f6-bfa8-a9f8244f7778.png
        :alt: Writer Paragraph Area Image dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Area Image dialog

There are many presets for image available.
Using the :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` enum, you can set the image to one of the presets.


Setup
-----

General function used to run these examples:

.. must be before the tabs directive
.. include:: ../../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            from ooodev.format.writer.direct.para.area import Img, PresetImageKind
            from ooodev.office.write import Write
            from ooodev.utils.gui import GUI
            from ooodev.utils.lo import Lo


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

                    image_style = Img.from_preset(PresetImageKind.FENCE)
                    Write.append_para(cursor=cursor, text=p_txt, styles=[image_style])

                    para_cursor = Write.get_paragraph_cursor(cursor)
                    para_cursor.gotoPreviousParagraph(False)
                    para_cursor.gotoEndOfParagraph(True)

                    text_para = cast("TextRangeContentProperties", para_cursor)

                    para_img = Img.from_obj(text_para.TextParagraph)
                    assert para_img is not None

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

Apply Image
^^^^^^^^^^^

Create a image for the paragraph background.

.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)

        image_style = Img.from_preset(PresetImageKind.FENCE)
        Write.append_para(cursor=cursor, text=p_txt, styles=[image_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _233866774-2acf6d1b-5e5b-4e4c-8018-b72cc07c7970:
    .. figure:: https://user-images.githubusercontent.com/4193389/233866774-2acf6d1b-5e5b-4e4c-8018-b72cc07c7970.png
        :alt: Writer Paragraph Area Image
        :figclass: align-center
        :width: 520px

        Writer Paragraph Area Image

Get Image from Paragraph
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

        para_img = Img.from_obj(text_para.TextParagraph)
        assert para_img is not None

        para_cursor.gotoEnd(False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

   .. cssclass:: ul-list

        - :ref:`help_writer_format_style_para_reset_default`
        - :ref:`help_writer_format_style_para`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_para_image`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.area.Img`