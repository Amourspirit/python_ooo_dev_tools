.. _help_writer_format_direct_para_outline_and_list:

Write Direct Paragraph Outline & List
=====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Writer has a Outlines & List dialog tab.

The :py:class:`ooodev.format.writer.direct.para.outline_list.Outline`, :py:class:`ooodev.format.writer.direct.para.outline_list.LineNum`,
and :py:class:`ooodev.format.writer.direct.para.outline_list.ListStyle` classes are used to set the outline, list style, and line numbering of a paragraph.


.. cssclass:: screen_shot

    .. _ss_writer_dialog_para_outline_and_list:
    .. figure:: https://user-images.githubusercontent.com/4193389/229916123-f708d23e-ead9-4b7a-ab95-59e5fe703ded.png
        :alt: Writer Paragraph Outline & List dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Outline & List dialog

Setup
-----

.. must be before the tabs directive
.. include:: ../../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            from ooodev.office.write import Write
            from ooodev.utils.gui import GUI
            from ooodev.utils.lo import Lo
            from ooodev.format.writer.direct.para.outline_list import LevelKind, Outline, LineNum, ListStyle
            
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
                    Write.append_para(
                        cursor=cursor, text=p_txt, styles=[Outline(LevelKind.LEVEL_01)]
                    )
                    Write.append_para(
                        cursor=cursor, text=p_txt, styles=[Outline(LevelKind.TEXT_BODY)]
                    )
                    Lo.delay(1_000)
                    Lo.close_doc(doc)
                return 0


            if __name__ == "__main__":
                SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Outline Class
-------------

The :py:class:`ooodev.format.writer.direct.para.outline_list.Outline` class is used to set the outline level of a paragraph.

Example
^^^^^^^

In this example the first paragraph is given a outline level of ``1`` using :py:class:`~ooodev.format.writer.direct.para.outline_list.Outline` class.

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        Write.append_para(
            cursor=cursor, text=p_txt, styles=[Outline(LevelKind.LEVEL_01)]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229931097-35f40c72-22c9-452a-9dab-f15427afa2eb:
    .. figure:: https://user-images.githubusercontent.com/4193389/229931937-f7c1d45b-6646-42b3-91a5-b7faee1b7780.png
        :alt: Outline, Level 1
        :figclass: align-center
        :width: 550px

        Outline, Level 1.

ListStyle Class
---------------

Example
^^^^^^^

When creating a list each paragraph becomes the next list item.
For this reason it is best to create a :py:class:`~ooodev.format.writer.direct.para.outline_list.ListStyle` and apply it directly to the cursor.
After list is written the cursor can be reset to a default and seen in the examples below.

Number List
"""""""""""

In this example the list is set to a Numbered List using ``Numbering 123`` style.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 9

        with Lo.Loader(Lo.ConnectSocket()):
            doc = Write.create_doc()
            GUI.set_visible(True, doc)
            Lo.delay(500)
            GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

            cursor = Write.get_cursor(doc)

            ol = ListStyle(StyleListKind.NUM_123)
            # apply numbered list directly to cursor
            ol.apply(cursor)
            for i in range(1, 6):
                Write.append_para(cursor=cursor, text=f"Num Point {i}")
            # reset cursor for next paragraph
            ol.default.apply(cursor)
            Write.append_para(cursor=cursor, text=p_txt)
            Lo.delay(1_000)
            Lo.close_doc(doc)
        return 0

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229934543-7ea97d0d-3ed2-4d96-9e9f-7ca698f0ea59:
    .. figure:: https://user-images.githubusercontent.com/4193389/229934543-7ea97d0d-3ed2-4d96-9e9f-7ca698f0ea59.png
        :alt: ListStyle, Numbered List 123
        :figclass: align-center

        ListStyle, Numbered List 123.

In this example the list is set to a Numbered List using ``Numbering ivx`` style.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 9

        with Lo.Loader(Lo.ConnectSocket()):
            doc = Write.create_doc()
            GUI.set_visible(True, doc)
            Lo.delay(500)
            GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

            cursor = Write.get_cursor(doc)

            ol = ListStyle(StyleListKind.NUM_ivx)
            # apply numbered list directly to cursor
            ol.apply(cursor)
            for i in range(1, 6):
                Write.append_para(cursor=cursor, text=f"Num Point {i}")
            # reset cursor for next paragraph
            ol.default.apply(cursor)
            Write.append_para(cursor=cursor, text=p_txt)
            Lo.delay(1_000)
            Lo.close_doc(doc)
        return 0

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229935125-97f7ce40-163e-4f6d-805b-98d3d427d6d4:
    .. figure:: https://user-images.githubusercontent.com/4193389/229935125-97f7ce40-163e-4f6d-805b-98d3d427d6d4.png
        :alt: ListStyle, Numbered List ivx
        :figclass: align-center

        ListStyle, Numbered List ivx.

Number List Reset
"""""""""""""""""

Number styles can also be reset.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 16

        with Lo.Loader(Lo.ConnectSocket()):
            doc = Write.create_doc()
            GUI.set_visible(True, doc)
            Lo.delay(500)
            GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

            cursor = Write.get_cursor(doc)
            # set num_start -2 to force number restart.
            ol = ListStyle(list_style=StyleListKind.NUM_123, num_start=-2)
            # apply numbered list directly to cursor
            ol.apply(cursor)
            for i in range(1, 6):
                Write.append_para(cursor=cursor, text=f"Num Point {i}")

            # include line number to reset list number
            ol = ListStyle(list_style=StyleListKind.NUM_123, num_start=1)
            ol.apply(cursor)
            for i in range(1, 6):
                Write.append_para(cursor=cursor, text=f"Num Point {i}")

            # reset cursor for next paragraph
            ol.default.apply(cursor)
            Write.append_para(cursor=cursor, text=p_txt)
            Lo.delay(1_000)
            Lo.close_doc(doc)
        return 0

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229936617-b27093de-bd3c-40f7-854d-543fa60042ac:
    .. figure:: https://user-images.githubusercontent.com/4193389/229936617-b27093de-bd3c-40f7-854d-543fa60042ac.png
        :alt: ListStyle, Numbered List with number reset
        :figclass: align-center

        ListStyle, Numbered List with number reset.



Other List Styles
"""""""""""""""""

Set style using ``List 3`` style.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 4

        cursor = Write.get_cursor(doc)

        # set num_start -2 to force number restart.
        ol = ListStyle(list_style=StyleListKind.LIST_03, num_start=-2)
        # apply numbered list directly to cursor
        ol.apply(cursor)
        for i in range(1, 6):
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        # reset cursor for next paragraph
        ol.default.apply(cursor)
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229937895-59818f82-9bd7-4337-92fa-e6c857c04ff1:
    .. figure:: https://user-images.githubusercontent.com/4193389/229937895-59818f82-9bd7-4337-92fa-e6c857c04ff1.png
        :alt: ListStyle, Numbered List 3
        :figclass: align-center

        ListStyle, Numbered List 3.

Set style using ``List 5`` style.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 4

        cursor = Write.get_cursor(doc)

        # set num_start -2 to force number restart.
        ol = ListStyle(list_style=StyleListKind.LIST_05, num_start=-2)
        # apply numbered list directly to cursor
        ol.apply(cursor)
        for i in range(1, 6):
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        # reset cursor for next paragraph
        ol.default.apply(cursor)
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229938370-b22952bc-d85b-4b00-b8b8-06b2cb473270:
    .. figure:: https://user-images.githubusercontent.com/4193389/229938370-b22952bc-d85b-4b00-b8b8-06b2cb473270.png
        :alt: ListStyle, Numbered List 5
        :figclass: align-center

        ListStyle, Numbered List 5.


LineNum Class
-------------

The :py:class:`~ooodev.format.writer.direct.para.outline_list.LineNum` class is used to set line numbering for a paragraph.

If ``num=0`` then this paragraph is include in line numbering.
If ``num=-1`` then this paragraph is excluded in line numbering.
If greater then zero then this paragraph is included in line numbering and the numbering is restarted with value of ``num``.

Example
^^^^^^^

Set to 3
""""""""

In this example the paragraph line number start value is set to ``3``.

.. tabs::

    .. code-tab:: python

        # ... other code
        ln = LineNum(3)
        Write.append_para(cursor=cursor, text=p_txt, styles=[ln])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229939917-1289d74d-db02-4e55-a4b1-8e8c23c7912c:
    .. figure:: https://user-images.githubusercontent.com/4193389/229939917-1289d74d-db02-4e55-a4b1-8e8c23c7912c.png
        :alt: Paragraph Line Numbering
        :figclass: align-center
        :width: 550px

        Paragraph Line Numbering


.. cssclass:: screen_shot

    .. _229940097-5d63d58f-592f-4927-b3f2-5f6302b3345d:
    .. figure:: https://user-images.githubusercontent.com/4193389/229940097-5d63d58f-592f-4927-b3f2-5f6302b3345d.png
        :alt: Paragraph Outline & List dialog
        :figclass: align-center
        :width: 450px

        Paragraph Outline & List dialog.


Exclude
"""""""

In this example the paragraph line number start value is set to ``-1`` (exclude).

.. tabs::

    .. code-tab:: python

        # ... other code
        ln = LineNum(-1)
        Write.append_para(cursor=cursor, text=p_txt, styles=[ln])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229940898-f5082c21-4b90-45a3-a027-edde6cb9c03e:
    .. figure:: https://user-images.githubusercontent.com/4193389/229940898-f5082c21-4b90-45a3-a027-edde6cb9c03e.png
        :alt: Paragraph Line Numbering
        :figclass: align-center
        :width: 550px

        Paragraph Line Numbering


.. cssclass:: screen_shot

    .. _229941010-dd5753ea-b35f-4708-8067-1c1966d60d85:
    .. figure:: https://user-images.githubusercontent.com/4193389/229941010-dd5753ea-b35f-4708-8067-1c1966d60d85.png
        :alt: Paragraph Outline & List dialog
        :figclass: align-center
        :width: 450px

        Paragraph Outline & List dialog.

Include
"""""""

In this example the paragraph line number start value is set to ``0`` (include).

.. tabs::

    .. code-tab:: python

        # ... other code
        ln = LineNum(0)
        Write.append_para(cursor=cursor, text=p_txt, styles=[ln])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229941793-b8326483-d3b9-4228-97e1-33d1396d9f49:
    .. figure:: https://user-images.githubusercontent.com/4193389/229941793-b8326483-d3b9-4228-97e1-33d1396d9f49.png
        :alt: Paragraph Line Numbering
        :figclass: align-center
        :width: 550px

        Paragraph Line Numbering


.. cssclass:: screen_shot

    .. _229941914-b3651119-9b5c-4d97-94f4-c83f11744d35:
    .. figure:: https://user-images.githubusercontent.com/4193389/229941914-b3651119-9b5c-4d97-94f4-c83f11744d35.png
        :alt: Paragraph Outline & List dialog
        :figclass: align-center
        :width: 450px

        Paragraph Outline & List dialog.

.. seealso::

   .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_para_outline_and_list`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.outline_list.Outline`
        - :py:class:`ooodev.format.writer.direct.para.outline_list.LineNum`
        - :py:class:`ooodev.format.writer.direct.para.outline_list.ListStyle`
