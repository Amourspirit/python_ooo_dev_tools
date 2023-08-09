.. _help_writer_format_direct_char_borders:

Write Direct Character Borders Class
====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.char.borders.Borders` class gives you the same options
as Writer's Borders Dialog, but without the dialog. as seen in :numref:`ss_writer_dialog_char_borders`.

.. cssclass:: screen_shot

    .. _ss_writer_dialog_char_borders:
    .. figure:: https://user-images.githubusercontent.com/4193389/233844568-3688fa44-bd94-468c-970a-d2b0f5261983.png
        :alt: Writer dialog Character Borders
        :figclass: align-center
        :width: 450px

        Writer dialog Character Borders

Setting the character borders
-----------------------------

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.char.borders import Borders, Side
        from ooodev.format.writer.direct.char.borders import Padding
        from ooodev.format.writer.direct.char.borders import Shadow
        from ooodev.format import StandardColor
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
                cursor = Write.get_cursor(doc)

                Write.append(cursor, "Hello ")

                side = Side(color=StandardColor.GREEN_LIGHT2)
                border_style = Borders(all=side)

                Write.append(cursor, "World", styles=[border_style])
                Write.end_paragraph(cursor)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`233844781-5680f568-666a-4e3c-97e0-d6430dbc4104`.

.. cssclass:: screen_shot

    .. _233844781-5680f568-666a-4e3c-97e0-d6430dbc4104:
    .. figure:: https://user-images.githubusercontent.com/4193389/233844781-5680f568-666a-4e3c-97e0-d6430dbc4104.png
        :alt: Character Borders
        :figclass: align-center

        Character Borders

Setting the character borders With padding
------------------------------------------

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)

        Write.append(cursor, "Hello ")
    
        side = Side(color=StandardColor.GREEN_LIGHT2)
        border_style = Borders(all=side)
        # create a padding of 3 mm on all sides
        padding_style = Padding(all=3)
        Write.append(cursor, "World", styles=[border_style, padding_style])
        Write.end_paragraph(cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`233845741-f8284145-f350-4f8b-9521-91689af629b9`.

.. cssclass:: screen_shot

    .. _233845741-f8284145-f350-4f8b-9521-91689af629b9:
    .. figure:: https://user-images.githubusercontent.com/4193389/233845741-f8284145-f350-4f8b-9521-91689af629b9.png
        :alt: Character Borders with shadow
        :figclass: align-center

        Character Borders with shadow

Setting the character borders With Shadow
-----------------------------------------

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)

        side = Side(color=StandardColor.GREEN_LIGHT2)
        border_style = Borders(all=side)

        Write.append(cursor, "Hello ")
        # create shadow
        shadow_style = Shadow(color=StandardColor.GREEN_DARK2, width=1.0)
        Write.append(cursor, "World", styles=[border_style, shadow_style])
        Write.end_paragraph(cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`233846091-f9a38f33-3cde-4428-b056-d9c0dc6a1251`.

.. cssclass:: screen_shot

    .. _233846091-f9a38f33-3cde-4428-b056-d9c0dc6a1251:
    .. figure:: https://user-images.githubusercontent.com/4193389/233846091-f9a38f33-3cde-4428-b056-d9c0dc6a1251.png
        :alt: Character Borders with shadow
        :figclass: align-center

        Character Borders with shadow

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_char_borders`
        - :ref:`help_calc_format_direct_cell_borders`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.char.borders.Borders`
        - :py:class:`ooodev.format.writer.direct.char.borders.Padding`
        - :py:class:`ooodev.format.writer.direct.char.borders.Shadow`