.. _help_writer_format_direct_para_area_pattern:

Write Direct Paragraph Area Pattern Class
=========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Writer has an Area Pattern dialog section.

The :py:class:`ooodev.format.writer.direct.para.area.Pattern` class is used to set the paragraph background pattern.

.. cssclass:: screen_shot

    .. _ss_writer_dialog_para_area_pattern:
    .. figure:: https://user-images.githubusercontent.com/4193389/233867486-6746cf77-c164-4add-819a-bec376920402.png
        :alt: Writer Paragraph Area Pattern dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Area Pattern dialog

There are many presets for image available.
Using the :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` enum, you can set the pattern to one of the presets.


Setup
-----

General function used to run these examples:

.. must be before the tabs directive
.. include:: ../../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            from ooodev.format.writer.direct.para.area import Pattern, PresetPatternKind
            from ooodev.office.write import Write
            from ooodev.utils.gui import GUI
            from ooodev.loader.lo import Lo


            def main() -> int:
                p_txt = (
                    |short_ptext|
                )

                with Lo.Loader(Lo.ConnectPipe()):
                    doc = Write.create_doc()
                    GUI.set_visible(doc=doc)
                    Lo.delay(300)
                    GUI.zoom(GUI.ZoomEnum.ZOOM_100_PERCENT)
                    cursor = Write.get_cursor(doc)

                    pattern_style = Pattern.from_preset(PresetPatternKind.HORIZONTAL_BRICK)
                    Write.append_para(cursor=cursor, text=p_txt, styles=[pattern_style])

                    para_cursor = Write.get_paragraph_cursor(cursor)
                    para_cursor.gotoPreviousParagraph(False)
                    para_cursor.gotoEndOfParagraph(True)

                    text_para = cast("TextRangeContentProperties", para_cursor)

                    para_img = Pattern.from_obj(text_para.TextParagraph)
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

Apply Pattern
^^^^^^^^^^^^^

Create a pattern for the paragraph background.

.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)

        pattern_style = Pattern.from_preset(PresetPatternKind.HORIZONTAL_BRICK)
        Write.append_para(cursor=cursor, text=p_txt, styles=[pattern_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _233867593-a1a78f33-b099-4e1b-a50d-3a960dd67b55:
    .. figure:: https://user-images.githubusercontent.com/4193389/233867593-a1a78f33-b099-4e1b-a50d-3a960dd67b55.png
        :alt: Writer Paragraph Area Pattern
        :figclass: align-center
        :width: 520px

        Writer Paragraph Area Pattern

Get Pattern from Paragraph
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        # ... other code

        para_cursor = Write.get_paragraph_cursor(cursor)
        para_cursor.gotoPreviousParagraph(False)
        para_cursor.gotoEndOfParagraph(True)

        text_para = cast("TextRangeContentProperties", para_cursor)

        para_img = Pattern.from_obj(text_para.TextParagraph)
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
        - :ref:`help_writer_format_modify_para_pattern`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.area.Pattern`