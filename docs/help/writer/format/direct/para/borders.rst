.. _help_writer_format_direct_para_borders:

Write Direct Paragraph Borders
==============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Writer has an Borders dialog tab.

The :py:class:`ooodev.format.writer.direct.para.borders.Borders` class is used to set the paragraph borders.


.. cssclass:: screen_shot

    .. _ss_writer_dialog_para_borders:
    .. figure:: https://user-images.githubusercontent.com/4193389/230152464-f444a863-45ed-4a98-9eba-27dc95a0031c.png
        :alt: Writer Paragraph Borders dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Borders dialog.

Setup
-----

General function used to run these examples:

.. must be before the tabs directive
.. include:: ../../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            from ooodev.format.writer.style import Para as StylePara
            from ooodev.office.write import Write
            from ooodev.utils.color import CommonColor
            from ooodev.utils.gui import GUI
            from ooodev.loader.lo import Lo
            from ooodev.format.writer.direct.para.borders import (
                Borders,
                BorderLineKind,
                Padding,
                Side,
                Shadow,
            )
            
            def main() -> int:
                p_txt = (
                    |short_ptext|
                )

                with Lo.Loader(Lo.ConnectSocket()):
                    doc = Write.create_doc()
                    GUI.set_visible(True, doc)
                    Lo.delay(500)
                    GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                    cursor = Write.get_cursor(doc)
                    bdr = Borders(
                        all=Side(line=BorderLineKind.DASH_DOT, color=CommonColor.BLUE_VIOLET),
                        shadow=Shadow(),
                        padding=Padding(all=1.7),
                        merge=False,
                    )
                    Write.append_para(cursor=cursor, text=p_txt, styles=[bdr])
                    StylePara.default.apply(cursor)
                    Write.append_para(cursor=cursor, text=p_txt)
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

After applying borders to a paragraph, the default paragraph style is applied to the paragraph using :ref:`help_writer_format_style_para_reset_default`.

Apply Border
^^^^^^^^^^^^

Create a border around a paragraph.

Note that next paragraph has the same border properties because merge is the default behaviour.

.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)
        bdr = Borders(all=Side(line=BorderLineKind.DASH_DOT, color=CommonColor.BLUE_VIOLET))
        Write.append_para(cursor=cursor, text=p_txt, styles=[bdr])
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _230156822-0e7feccc-5a23-448d-8aa5-f47fad4a8e74:
    .. figure:: https://user-images.githubusercontent.com/4193389/230156822-0e7feccc-5a23-448d-8aa5-f47fad4a8e74.png
        :alt: Writer Paragraph Borders
        :figclass: align-center

        Writer Paragraph Borders.


.. cssclass:: screen_shot

    .. _230157493-f609d4dc-415e-415e-807a-e4ff09abc5af:
    .. figure:: https://user-images.githubusercontent.com/4193389/230157493-f609d4dc-a9fc-415e-807a-e4ff09abc5af.png
        :alt: Writer Paragraph Borders Dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Borders Dialog.


When ``merge=False`` the paragraphs have the same border properties but are not merged.

.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)
        bdr = Borders(all=Side(line=BorderLineKind.DASH_DOT, color=CommonColor.BLUE_VIOLET), merge=False)
        Write.append_para(cursor=cursor, text=p_txt, styles=[bdr])
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230158435-ea58957f-933c-44c1-a88d-fdce0b6b99a9:
    .. figure:: https://user-images.githubusercontent.com/4193389/230158435-ea58957f-933c-44c1-a88d-fdce0b6b99a9.png
        :alt: Writer Paragraph Borders
        :figclass: align-center
        :width: 450px

        Writer Paragraph Borders.


Resetting for the next paragraph can be done by applying :py:attr:`Para.default <ooodev.format.writer.style.Para.default>` to the cursor.


.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)
        bdr = Borders(all=Side(line=BorderLineKind.DASH_DOT, color=CommonColor.BLUE_VIOLET), merge=False)
        Write.append_para(cursor=cursor, text=p_txt, styles=[bdr])
        StylePara.default.apply(cursor)
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230159618-3a44c8f8-a96f-4889-80ba-9a02c07054b5:
    .. figure:: https://user-images.githubusercontent.com/4193389/230159618-3a44c8f8-a96f-4889-80ba-9a02c07054b5.png
        :alt: Writer Paragraph Borders
        :figclass: align-center
        :width: 450px

        Writer Paragraph Borders.


Apply Border With Shadow
^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)
        bdr = Borders(
            all=Side(line=BorderLineKind.DASH_DOT, color=CommonColor.BLUE_VIOLET),
            shadow=Shadow(),
            merge=False,
        )
        Write.append_para(cursor=cursor, text=p_txt, styles=[bdr])
        StylePara.default.apply(cursor)
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _230161085-e2642f76-5558-40ef-b5b2-b169598fea38:
    .. figure:: https://user-images.githubusercontent.com/4193389/230161085-e2642f76-5558-40ef-b5b2-b169598fea38.png
        :alt: Writer Paragraph Borders
        :figclass: align-center

        Writer Paragraph Borders.


.. cssclass:: screen_shot

    .. _230161408-c1e66410-f88d-4fc0-8124-f8db8c8a3b45:
    .. figure:: https://user-images.githubusercontent.com/4193389/230161408-c1e66410-f88d-4fc0-8124-f8db8c8a3b45.png
        :alt: Writer Paragraph Borders Dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Borders Dialog.


Apply Border With Shadow and Padding
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)
        bdr = Borders(
            all=Side(line=BorderLineKind.DASH_DOT, color=CommonColor.BLUE_VIOLET),
            padding=Padding(all=7.7),
            shadow=Shadow(),
            merge=False,
        )
        Write.append_para(cursor=cursor, text=p_txt, styles=[bdr])
        StylePara.default.apply(cursor)
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _230162520-de38f579-6a6d-492a-847a-19ee0ee6314f:
    .. figure:: https://user-images.githubusercontent.com/4193389/230162520-de38f579-6a6d-492a-847a-19ee0ee6314f.png
        :alt: Writer Paragraph Borders
        :figclass: align-center

        Writer Paragraph Borders.


.. cssclass:: screen_shot

    .. _230162825-b4fa9c13-9a3f-4633-a36a-7c804a21adc1:
    .. figure:: https://user-images.githubusercontent.com/4193389/230162825-b4fa9c13-9a3f-4633-a36a-7c804a21adc1.png
        :alt: Writer Paragraph Borders Dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Borders Dialog.

Getting Borders from Paragraph
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

A paragraph cursor object is used to select the first paragraph in the document.
The paragraph cursor is then used to get the style.

.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)
        bdr = Borders(
            all=Side(line=BorderLineKind.DASH_DOT, color=CommonColor.BLUE_VIOLET),
            padding=Padding(all=7.7),
            shadow=Shadow(),
            merge=False,
        )
        Write.append_para(cursor=cursor, text=p_txt, styles=[bdr])

        para_cursor = Write.get_paragraph_cursor(cursor)
        para_cursor.gotoPreviousParagraph(False)
        para_cursor.gotoEndOfParagraph(True)

        para_bdr = Borders.from_obj(para_cursor)
        assert para_bdr.prop_inner_sides.prop_left.prop_line == BorderLineKind.DASH_DOT
        assert para_bdr.prop_inner_sides.prop_right.prop_color == CommonColor.BLUE_VIOLET

        StylePara.default.apply(cursor)
        Write.append_para(cursor=cursor, text=p_txt)

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
        - :ref:`help_calc_format_direct_cell_borders`
        - :ref:`help_writer_format_modify_para_borders`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.borders.Borders`