.. _help_writer_format_direct_para_tabs:

Write Direct Paragraph Tabs
===========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 3

Overview
--------

Writer has a Tabs dialog tab.

The :py:class:`ooodev.format.writer.direct.para.tabs.Tabs` class is used to set the paragraph tabs.


.. cssclass:: screen_shot

    .. _ss_writer_dialog_para_tabs:
    .. figure:: https://user-images.githubusercontent.com/4193389/230204230-0edaf74a-c3e1-4249-9521-cbcb2b4a8894.png
        :alt: Writer Paragraph Tabs dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Tabs dialog.

Setup
-----

General function used to run these examples:

.. tabs::

    .. code-tab:: python

        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.writer.direct.para.tabs import Tabs, TabAlign, FillCharKind


        def main() -> int:
            with Lo.Loader(Lo.ConnectSocket()):
                doc = Write.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                cursor = Write.get_cursor(doc)
                tb = Tabs(position=11.3, align=TabAlign.LEFT, fill_char=FillCharKind.UNDER_SCORE)
                Write.append_para(cursor=cursor, text="Some Paragraph", styles=[tb])

                tb = Tabs(position=12.0, align=TabAlign.DECIMAL)
                tb.apply(cursor)

                tb = Tabs(position=6.5, align=TabAlign.CENTER, fill_char="*")
                tb.apply(cursor)

                tb = Tabs.find(cursor, 6.5)
                tb.prop_align = TabAlign.RIGHT
                tb.prop_fill_char = FillCharKind.DASH
                tb.apply(cursor)

                Tabs.remove_by_pos(cursor, 12.0)

                Tabs.remove_all(cursor)

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

Tabs in Writer are determined by ``Position``. When adding a Tab with the same ``Position`` value
as another existing Tab it results the existing Tab's values being updated.

Adding Tabs
^^^^^^^^^^^

Add via creating a paragraph
""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        tb = Tabs(position=11.3, align=TabAlign.LEFT, fill_char=FillCharKind.UNDER_SCORE)
        Write.append_para(cursor=cursor, text="Some Paragraph", styles=[tb])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230206247-f350e985-83af-44aa-bdf2-67c56bdeb17f:
    .. figure:: https://user-images.githubusercontent.com/4193389/230206247-f350e985-83af-44aa-bdf2-67c56bdeb17f.png
        :alt: Writer Paragraph Tabs dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Tabs dialog.

Add via applying directly to Cursor
"""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        tb = Tabs(position=11.3, align=TabAlign.LEFT, fill_char=FillCharKind.UNDER_SCORE)
        Write.append_para(cursor=cursor, text="Some Paragraph", styles=[tb])

        tb = Tabs(position=12.0, align=TabAlign.DECIMAL)
        tb.apply(cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230207212-8bc9ca1c-307c-4161-85ec-bb36673a9a89:
    .. figure:: https://user-images.githubusercontent.com/4193389/230207212-8bc9ca1c-307c-4161-85ec-bb36673a9a89.png
        :alt: Writer Paragraph Tabs dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Tabs dialog.

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        tb = Tabs(position=11.3, align=TabAlign.LEFT, fill_char=FillCharKind.UNDER_SCORE)
        Write.append_para(cursor=cursor, text="Some Paragraph", styles=[tb])

        tb = Tabs(position=12.0, align=TabAlign.DECIMAL)
        tb.apply(cursor)

        tb = Tabs(position=6.5, align=TabAlign.CENTER, fill_char="*")
        tb.apply(cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230208002-74b26b94-b1c6-4274-874a-ae21d7b268e3:
    .. figure:: https://user-images.githubusercontent.com/4193389/230208002-74b26b94-b1c6-4274-874a-ae21d7b268e3.png
        :alt: Writer Paragraph Tabs dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Tabs dialog.

Updating an existing tab
""""""""""""""""""""""""

Finds the tab that was initially set with a position of ``6.5``, updates is value and applies it to the cursor.

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        tb = Tabs(position=11.3, align=TabAlign.LEFT, fill_char=FillCharKind.UNDER_SCORE)
        Write.append_para(cursor=cursor, text="Some Paragraph", styles=[tb])

        tb = Tabs(position=12.0, align=TabAlign.DECIMAL)
        tb.apply(cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The result is the value are now updated.

.. cssclass:: screen_shot

    .. _230208703-04eb0210-4c55-41f8-8b8d-c02afbafef4d:
    .. figure:: https://user-images.githubusercontent.com/4193389/230208703-04eb0210-4c55-41f8-8b8d-c02afbafef4d.png
        :alt: Writer Paragraph Tabs dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Tabs dialog.

Removing Tabs
^^^^^^^^^^^^^

Removing a Tab

Remove a Tab can be done via :py:meth:`Tabs.remove_by_pos <ooodev.format.writer.direct.para.tabs.Tabs.remove_by_pos>`, which removes a tab with it position as input.
Or :py:meth:`Tabs.remove <ooodev.format.writer.direct.para.tabs.Tabs.remove>` which can take a ``Tab`` or ``TabStop`` as input (``Tabs`` inherits from ``Tab``).

.. tabs::

    .. code-tab:: python

        # ... other code
        Tabs.remove_by_pos(cursor, 12.0)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _230209932-ac363e4d-7e21-4a18-8c68-3d8a7691ce6c:
    .. figure:: https://user-images.githubusercontent.com/4193389/230209932-ac363e4d-7e21-4a18-8c68-3d8a7691ce6c.png
        :alt: Writer Paragraph Tabs dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Tabs dialog.


.. tabs::

    .. code-tab:: python

        # ... other code
        Tabs.remove_all(cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

..
    copy of figure _ss_writer_dialog_para_tabs fom top of page

.. cssclass:: screen_shot

    .. figure:: https://user-images.githubusercontent.com/4193389/230204230-0edaf74a-c3e1-4249-9521-cbcb2b4a8894.png
        :alt: Writer Paragraph Tabs dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Tabs dialog.