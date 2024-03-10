.. _help_writer_format_direct_table_borders:

Write Direct Table Borders
==========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 3

The :py:class:`ooodev.format.writer.direct.table.borders.Borders` class can be used to set the borders of a table.
There are also other ways to set the borders of a table as noted in the examples below.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.table.borders import (
            Borders,
            Shadow,
            Side,
            BorderLineKind,
            Padding,
            LineSize,
        )
        from ooodev.write import WriteDoc
        from ooodev.utils.color import CommonColor
        from ooodev.utils.color import StandardColor
        from ooodev.loader import Lo
        from ooodev.utils.table_helper import TableHelper


        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = WriteDoc.create_doc(visible=True)
                Lo.delay(300)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)
                cursor = doc.get_cursor()

                tbl_data = TableHelper.make_2d_array(num_rows=5, num_cols=5)
                # bg_img_style = TblBgImg.from_preset(PresetImageKind.PAPER_TEXTURE)
                table = cursor.add_table(
                    table_data=tbl_data,
                    tbl_bg_color=CommonColor.LIGHT_BLUE,
                    tbl_fg_color=CommonColor.BLACK,
                    first_row_header=False,
                )

                blue_side = Side(line=BorderLineKind.SOLID, color=StandardColor.BLUE_DARK2, width=LineSize.THICK)
                green_side = Side(line=BorderLineKind.SOLID, color=StandardColor.GREEN_DARK1, width=LineSize.THIN)
                bdr_style = Borders(
                    border_side=blue_side,
                    vertical=green_side,
                    horizontal=green_side,
                    shadow=Shadow(color=StandardColor.BLUE_DARK2),
                )

                bdr_style.apply(table.component)

                # getting the table properties
                tbl_bdr_style = Borders.from_obj(table.component)
                assert tbl_bdr_style is not None

                Lo.delay(1_000)
                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Examples
--------

Borders
+++++++

Borders Simple
^^^^^^^^^^^^^^

Set using style_direct
""""""""""""""""""""""

Table color can also be set using ``style_direct`` property of the table.

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )
        table.style_direct.style_borders_side()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Set using styles
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        bdr_style = Borders(border_side=Side())
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
            styles=[bdr_style],
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234038179-f02a8294-fa98-4c6d-b897-e50b2a509c0c:

    .. figure:: https://user-images.githubusercontent.com/4193389/234038179-f02a8294-fa98-4c6d-b897-e50b2a509c0c.png
        :alt: Border simple
        :figclass: align-center
        :width: 520px

        Border simple


.. cssclass:: screen_shot

    .. _234038394-d5b35e6f-1b84-4972-b990-5741fd9c19c6:

    .. figure:: https://user-images.githubusercontent.com/4193389/234038394-d5b35e6f-1b84-4972-b990-5741fd9c19c6.png
        :alt: Table Borders Dialog
        :figclass: align-center
        :width: 450px

        Table Borders Dialog


Set using table_border2
""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )

        table.table_border2.left_line.color = StandardColor.RED_DARK1
        table.table_border2.left_line.line_width = UnitPT(float(LineSize.MEDIUM))
        table.table_border2.left_line.line_style = BorderLineKind.SOLID
        table.table_border2.right_line = table.table_border2.left_line


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set using style_direct
""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )

        default_side = Side()
        red_side = Side(
            line=BorderLineKind.SOLID, color=StandardColor.RED_DARK1, width=LineSize.MEDIUM
        )
        table.style_direct.style_borders(
            left=red_side, right=red_side, top=default_side, bottom=default_side
        )


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set using styles
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )

        default_side = Side()
        red_side = Side(
            line=BorderLineKind.SOLID, color=StandardColor.RED_DARK1, width=LineSize.MEDIUM
        )
        bdr_style = Borders(left=red_side, top=default_side, bottom=default_side, right=red_side)

        bdr_style.apply(table.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234112245-28e7d85f-19dd-438d-a701-d5f32a5929e7:

    .. figure:: https://user-images.githubusercontent.com/4193389/234112245-28e7d85f-19dd-438d-a701-d5f32a5929e7.png
        :alt: Border Red Sides
        :figclass: align-center
        :width: 520px

        Border Red Sides


.. cssclass:: screen_shot

    .. _234112467-e8549bf9-62c6-4442-84ed-5e2e2b00477a:

    .. figure:: https://user-images.githubusercontent.com/4193389/234112467-e8549bf9-62c6-4442-84ed-5e2e2b00477a.png
        :alt: Table Borders Dialog
        :figclass: align-center
        :width: 450px

        Table Borders Dialog


Borders Set Horizontal & Vertical
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Set using table_border2
""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )

        table.table_border2.left_line.color = StandardColor.BLUE_DARK2
        table.table_border2.left_line.line_width = UnitPT(float(LineSize.THICK))
        table.table_border2.left_line.line_style = BorderLineKind.SOLID
        table.table_border2.right_line = table.table_border2.left_line
        table.table_border2.top_line = table.table_border2.left_line
        table.table_border2.bottom_line = table.table_border2.left_line

        table.table_border2.vertical_line.color = StandardColor.GREEN_DARK1
        table.table_border2.vertical_line.line_width = UnitPT(float(LineSize.THIN))
        table.table_border2.vertical_line.line_style = BorderLineKind.SOLID
        table.table_border2.horizontal_line = table.table_border2.vertical_line


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Set using style_direct
""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )

        blue_side = Side(
            line=BorderLineKind.SOLID, color=StandardColor.BLUE_DARK2, width=LineSize.THICK
        )
        green_side = Side(
            line=BorderLineKind.SOLID, color=StandardColor.GREEN_DARK1, width=LineSize.THIN
        )
        table.style_direct.style_borders(
            border_side=blue_side, vertical=green_side, horizontal=green_side
        )


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set using styles
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )

        blue_side = Side(
            line=BorderLineKind.SOLID, color=StandardColor.BLUE_DARK2, width=LineSize.THICK
        )
        green_side = Side(
            line=BorderLineKind.SOLID, color=StandardColor.GREEN_DARK1, width=LineSize.THIN
        )
        bdr_style = Borders(
            border_side=blue_side,
            vertical=green_side,
            horizontal=green_side,
        )

        bdr_style.apply(table.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234114135-189451ce-e25f-43ba-bce7-70506d2c03f3:

    .. figure:: https://user-images.githubusercontent.com/4193389/234114135-189451ce-e25f-43ba-bce7-70506d2c03f3.png
        :alt: Borders Set Horizontal & Vertical
        :figclass: align-center
        :width: 520px

        Borders Set Horizontal & Vertical


.. cssclass:: screen_shot

    .. _234114333-7d0889d5-c80a-4fc9-b30e-460afeb57de0:

    .. figure:: https://user-images.githubusercontent.com/4193389/234114333-7d0889d5-c80a-4fc9-b30e-460afeb57de0.png
        :alt: Table Borders Dialog
        :figclass: align-center
        :width: 450px

        Table Borders Dialog

Padding
+++++++

Borders & Padding
^^^^^^^^^^^^^^^^^

Set using table_border2
""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.units import UnitMM

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )

        table.table_border2.left_line.color = StandardColor.BLUE_DARK2
        table.table_border2.left_line.line_width = UnitPT(float(LineSize.THICK))
        table.table_border2.left_line.line_style = BorderLineKind.SOLID
        table.table_border2.right_line = table.table_border2.left_line
        table.table_border2.top_line = table.table_border2.left_line
        table.table_border2.bottom_line = table.table_border2.left_line

        table.table_border2.vertical_line.color = StandardColor.GREEN_DARK1
        table.table_border2.vertical_line.line_width = UnitPT(float(LineSize.THIN))
        table.table_border2.vertical_line.line_style = BorderLineKind.SOLID
        table.table_border2.horizontal_line = table.table_border2.vertical_line

        table.table_border2.distance = UnitMM(5)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set using style_direct
""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )

        blue_side = Side(line=BorderLineKind.SOLID, color=StandardColor.BLUE_DARK2, width=LineSize.THICK)
        green_side = Side(line=BorderLineKind.SOLID, color=StandardColor.GREEN_DARK1, width=LineSize.THIN)
        padding = Padding(all=5)
        table.style_direct.style_borders(
            border_side=blue_side,
            vertical=green_side,
            horizontal=green_side,
            padding=padding,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Set using styles
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )

        blue_side = Side(line=BorderLineKind.SOLID, color=StandardColor.BLUE_DARK2, width=LineSize.THICK)
        green_side = Side(line=BorderLineKind.SOLID, color=StandardColor.GREEN_DARK1, width=LineSize.THIN)
        padding = Padding(all=5)
        bdr_style = Borders(
            border_side=blue_side,
            vertical=green_side,
            horizontal=green_side,
            padding=padding,
        )

        bdr_style.apply(table.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234115517-22704ec3-b3f5-4972-95d4-12a491ea85ce:

    .. figure:: https://user-images.githubusercontent.com/4193389/234115517-22704ec3-b3f5-4972-95d4-12a491ea85ce.png
        :alt: Borders and Padding
        :figclass: align-center
        :width: 520px

        Borders and Padding


.. cssclass:: screen_shot

    .. _234115698-6fb07d18-5472-4010-8ec6-6f514b1c4b6d:

    .. figure:: https://user-images.githubusercontent.com/4193389/234115698-6fb07d18-5472-4010-8ec6-6f514b1c4b6d.png
        :alt: Table Borders Dialog
        :figclass: align-center
        :width: 450px

        Table Borders Dialog

Shadow
+++++++

Borders & Shadow
^^^^^^^^^^^^^^^^

Set using shadow_format
"""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.units import UnitMM
        from ooo.dyn.table.shadow_location import ShadowLocation

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )

        table.table_border2.left_line.color = StandardColor.BLUE_DARK2
        table.table_border2.left_line.line_width = UnitPT(float(LineSize.THICK))
        table.table_border2.left_line.line_style = BorderLineKind.SOLID
        table.table_border2.right_line = table.table_border2.left_line
        table.table_border2.top_line = table.table_border2.left_line
        table.table_border2.bottom_line = table.table_border2.left_line

        table.table_border2.vertical_line.color = StandardColor.GREEN_DARK1
        table.table_border2.vertical_line.line_width = UnitPT(float(LineSize.THIN))
        table.table_border2.vertical_line.line_style = BorderLineKind.SOLID
        table.table_border2.horizontal_line = table.table_border2.vertical_line

        table.shadow_format.color = StandardColor.BLUE_DARK2
        table.shadow_format.location = ShadowLocation.BOTTOM_RIGHT
        table.shadow_format.is_transparent = False
        table.shadow_format.shadow_width = UnitMM(1.76)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set using style_direct
""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            tbl_bg_color=CommonColor.LIGHT_BLUE,
            tbl_fg_color=CommonColor.BLACK,
        )

        blue_side = Side(line=BorderLineKind.SOLID, color=StandardColor.BLUE_DARK2, width=LineSize.THICK)
        green_side = Side(line=BorderLineKind.SOLID, color=StandardColor.GREEN_DARK1, width=LineSize.THIN)
        table.style_direct.style_borders(
            border_side=blue_side,
            vertical=green_side,
            horizontal=green_side,
            shadow=Shadow(color=StandardColor.BLUE_DARK2),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set using styles
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
        )

        blue_side = Side(line=BorderLineKind.SOLID, color=StandardColor.BLUE_DARK2, width=LineSize.THICK)
        green_side = Side(line=BorderLineKind.SOLID, color=StandardColor.GREEN_DARK1, width=LineSize.THIN)
        bdr_style = Borders(
            border_side=blue_side,
            vertical=green_side,
            horizontal=green_side,
            shadow=Shadow(color=StandardColor.BLUE_DARK2),
        )

        bdr_style.apply(table.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234117019-78fc20c0-6885-4ce9-a2ba-a09170a93bdb:

    .. figure:: https://user-images.githubusercontent.com/4193389/234117019-78fc20c0-6885-4ce9-a2ba-a09170a93bdb.png
        :alt: Borders and Shadow
        :figclass: align-center
        :width: 520px

        Borders and Shadow


.. cssclass:: screen_shot

    .. _234117150-01fdbad2-4590-47a1-a94d-5dbfba646f94:

    .. figure:: https://user-images.githubusercontent.com/4193389/234117150-01fdbad2-4590-47a1-a94d-5dbfba646f94.png
        :alt: Table Borders Dialog
        :figclass: align-center
        :width: 450px

        Table Borders Dialog

Getting the Borders from the table
++++++++++++++++++++++++++++++++++

.. tabs::

    .. code-tab:: python

        # ... other code
        # getting the table properties
        tbl_bdr_style = Borders.from_obj(table.component)
        assert tbl_bdr_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_table_properties`
        - :ref:`help_writer_format_direct_table_background` 
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_direct_cell_borders`
        - :py:meth:`Write.add_table() <ooodev.office.write.Write.add_table>`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.writer.direct.table.borders.Borders`