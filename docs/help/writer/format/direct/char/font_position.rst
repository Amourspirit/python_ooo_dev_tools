.. _help_writer_format_direct_char_font_position:

Write Direct Character FontPosition Class
=========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 3


The :py:class:`ooodev.format.writer.direct.char.font.FontPosition` class is used to set the position of the character.
The class gives you the same options as Writer's Positon Dialog, but without the dialog. as seen in :numref:`ss_writer_dialog_char_font_position`.

.. cssclass:: screen_shot

    .. _ss_writer_dialog_char_font_position:
    .. figure:: https://user-images.githubusercontent.com/4193389/233837139-43ed8233-6957-4683-bd9e-beb5370f8dc2.png
        :alt: Font Position Dialog
        :figclass: align-center
        :width: 450px

        Font Position Dialog


Setup
-----

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.char.font import FontPosition
        from ooodev.format.writer.direct.char.font import CharSpacingKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo


        with Lo.Loader(Lo.ConnectPipe()):
            doc = Write.create_doc()
            GUI.set_visible(doc)
            Lo.delay(300)
            GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
            cursor = Write.get_cursor(doc)

            fp_style = FontPosition().superscript
            cursor = Write.get_cursor(doc)
            Write.append(cursor, "hello")
            Write.style(pos=0, length=1, styles=[fp_style], cursor=cursor)

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

Set Character to Superscript
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Setting a character to superscript can be done by using the superscript property of the FontPosition class.

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)

        fp_style = FontPosition().superscript
        cursor = Write.get_cursor(doc)
        Write.append(cursor, "hello")
        Write.style(pos=0, length=1, styles=[fp_style], cursor=cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _233836784-15e63ad6-5d70-446a-af89-77de03b07603:
    .. figure:: https://user-images.githubusercontent.com/4193389/233836784-15e63ad6-5d70-446a-af89-77de03b07603.png
        :alt: Superscript
        :figclass: align-center

        Superscript

Set Character to Subscript
^^^^^^^^^^^^^^^^^^^^^^^^^^

Setting a character to subscript can be done by using the subscript property of the FontPosition class.

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)

        fp_style = FontPosition().subscript
        cursor = Write.get_cursor(doc)
        Write.append(cursor, "hello")
        Write.style(pos=4, length=1, styles=[fp_style], cursor=cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _233837490-b176ead4-4e7d-4d24-ae36-e8f73fe7b652:
    .. figure:: https://user-images.githubusercontent.com/4193389/233837490-b176ead4-4e7d-4d24-ae36-e8f73fe7b652.png
        :alt: Subscript
        :figclass: align-center

        Subscript

Set Character to Normal
^^^^^^^^^^^^^^^^^^^^^^^

Setting a character to normal is done by using the normal property of the FontPosition class.

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)

        fp_style = FontPosition().subscript
        cursor = Write.get_cursor(doc)
        Write.append(cursor, "hello")
        Write.style(pos=4, length=1, styles=[fp_style], cursor=cursor)

        # set back to normal position by using the normal property
        Write.style(pos=4, length=1, styles=[fp_style.normal], cursor=cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set Character Rotation
^^^^^^^^^^^^^^^^^^^^^^

Setting characters rotation can be done by using one of the rotation properties of the FontPosition class or by using the rotation argument in the constructor.

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)

        fp_style = FontPosition().rotation_270
        # alternative
        # fp_style = FontPosition(rotation=270)
        cursor = Write.get_cursor(doc)
        Write.append(cursor, "Hello", styles=[fp_style])
        Write.append(cursor, "World", styles=[fp_style.rotation_90])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _233838116-01aaa779-dff0-4614-bd0c-54fae7dd80e0:
    .. figure:: https://user-images.githubusercontent.com/4193389/233838116-01aaa779-dff0-4614-bd0c-54fae7dd80e0.png
        :alt: Character Rotation
        :figclass: align-center

        Character Rotation

Set Character Spacing
^^^^^^^^^^^^^^^^^^^^^

Character Spacing Tight
"""""""""""""""""""""""


.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)

        fp_style = FontPosition(spacing=CharSpacingKind.TIGHT, pair=False)
        cursor = Write.get_cursor(doc)
        Write.append(cursor, "Hello", styles=[fp_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _233838540-fbd62095-67cd-4c81-8974-b26b989aa61b:
    .. figure:: https://user-images.githubusercontent.com/4193389/233838540-fbd62095-67cd-4c81-8974-b26b989aa61b.png
        :alt: Character Space Tight
        :figclass: align-center

        Character Space Tight

Character Spacing Very Loose
""""""""""""""""""""""""""""


.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)

        fp_style = FontPosition(spacing=CharSpacingKind.VERY_LOOSE, pair=True)
        cursor = Write.get_cursor(doc)
        Write.append(cursor, "Hello", styles=[fp_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _233838685-3941d255-5d66-4382-80d7-202f6ddf9ee8:
    .. figure:: https://user-images.githubusercontent.com/4193389/233838685-3941d255-5d66-4382-80d7-202f6ddf9ee8.png
        :alt: Character Space Very Loose
        :figclass: align-center

        Character Space Very Loose

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_char_font_only`
        - :ref:`help_writer_format_direct_char_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.char.font.FontPosition`