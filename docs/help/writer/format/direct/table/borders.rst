.. _help_writer_format_direct_table_borders:

Write Direct Table Borders
==========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 3

The :py:class:`ooodev.format.writer.direct.table.borders.Borders` c;ass is used to set the borders of a table.

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
        from ooodev.office.write import Write
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.utils.table_helper import TableHelper


        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_100_PERCENT)
                cursor = Write.get_cursor(doc)

                tbl_data = TableHelper.make_2d_array(num_rows=5, num_cols=5)
                table = Write.add_table(
                    cursor=cursor,
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

                bdr_style.apply(table)

                # getting the table properties
                tbl_bdr_style = Borders.from_obj(table)
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

.. tabs::

    .. code-tab:: python

        # ... other code
        bdr_style = Borders(border_side=Side())

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
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


Borders Red Sides
^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
        )

        default_side = Side()
        red_side = Side(line=BorderLineKind.SOLID, color=StandardColor.RED_DARK1, width=LineSize.MEDIUM)
        bdr_style = Borders(left=red_side, top=default_side, bottom=default_side, right=red_side)

        bdr_style.apply(table)

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

.. tabs::

    .. code-tab:: python

        # ... other code
        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
        )

        blue_side = Side(line=BorderLineKind.SOLID, color=StandardColor.BLUE_DARK2, width=LineSize.THICK)
        green_side = Side(line=BorderLineKind.SOLID, color=StandardColor.GREEN_DARK1, width=LineSize.THIN)
        bdr_style = Borders(
            border_side=blue_side,
            vertical=green_side,
            horizontal=green_side,
        )

        bdr_style.apply(table)

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

.. tabs::

    .. code-tab:: python

        # ... other code
        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
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

        bdr_style.apply(table)

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

.. tabs::

    .. code-tab:: python

        # ... other code
        table = Write.add_table(
            cursor=cursor,
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

        bdr_style.apply(table)

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
        tbl_bdr_style = Borders.from_obj(table)
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
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.table.borders.Borders`