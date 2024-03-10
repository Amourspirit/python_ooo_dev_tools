.. _help_writer_format_direct_table_background:

Write Direct Table Background
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 3

The :py:class:`ooodev.format.writer.direct.table.background.Color` and :py:class:`ooodev.format.writer.direct.table.background.Img` classes can used to set the background of a table.
There are also other ways to set the background of a table as noted in the examples below.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import sys
        import uno
        from ooodev.write import WriteDoc
        from ooodev.format.writer.direct.table.background import Color as TblBgColor
        from ooodev.format.writer.direct.table.background import Img as TblBgImg, PresetImageKind
        from ooodev.utils.color import StandardColor
        from ooodev.loader import Lo
        from ooodev.utils.kind.zoom_kind import ZoomKind
        from ooodev.utils.table_helper import TableHelper


        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = WriteDoc.create_doc(visible=True)
                Lo.delay(300)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)
                cursor = doc.get_cursor()

                tbl_data = TableHelper.make_2d_array(num_rows=5, num_cols=5)
                bg_color_style = TblBgColor(StandardColor.LIME_LIGHT3)
                table = cursor.add_table(
                    table_data=tbl_data,
                    first_row_header=False,
                    styles=[bg_color_style],
                )

                # getting the color object
                tbl_bg_color_style = TblBgColor.from_obj(table.component)
                assert tbl_bg_color_style is not None

                Lo.delay(1_000)
                doc.close()

            return 0


        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Examples
--------

Background Color
++++++++++++++++

Setting background color
^^^^^^^^^^^^^^^^^^^^^^^^

Set using back_color
""""""""""""""""""""

Table color can be set using ``back_color`` property of the table.

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
        )
        table.back_color = StandardColor.LIME_LIGHT3

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Set using style_direct
""""""""""""""""""""""

Table color can also be set using ``style_direct`` property of the table.

.. tabs::

    .. code-tab:: python

        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
        )
        table.style_direct.style_area_color(StandardColor.LIME_LIGHT3)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Set using styles
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        bg_color_style = TblBgColor(StandardColor.LIME_LIGHT3)
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            styles=[bg_color_style],
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None



.. cssclass:: screen_shot

    .. _234120927-65db58b6-2d26-4af9-bfdf-77a998a7eae3:

    .. figure:: https://user-images.githubusercontent.com/4193389/234120927-65db58b6-2d26-4af9-bfdf-77a998a7eae3.png
        :alt: Table Background Color
        :figclass: align-center
        :width: 520px

        Table Background Color

.. cssclass:: screen_shot

    .. _234121141-869acb01-ce86-47b0-a3c2-bcb3ef5faa46:

    .. figure:: https://user-images.githubusercontent.com/4193389/234121141-869acb01-ce86-47b0-a3c2-bcb3ef5faa46.png
        :alt: Table Background Color Dialog
        :figclass: align-center
        :width: 450px

        Table Background Color Dialog

Getting Background Color Object
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Get using style_direct
""""""""""""""""""""""

Table color can also be set using ``style_direct`` property of the table.

.. tabs::

    .. code-tab:: python

        # ... other code
        # getting the table properties
        tbl_bg_color_style = table.style_direct.style_area_color_get()
        assert tbl_bg_color_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Get using styles
""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        # getting the table properties
        tbl_bg_color_style = TblBgColor.from_obj(table.component)
        assert tbl_bg_color_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Background Image
++++++++++++++++

Setting background Image
^^^^^^^^^^^^^^^^^^^^^^^^

Background image has many options. The following example shows how to set a background image from a preset image.

The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` has many preset images to choose from.

Set using style_direct
""""""""""""""""""""""

Table background image can be set using ``style_direct`` property of the table.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.table.background import PresetImageKind
        # ... other code
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
        )
        table.style_direct.style_area_image_from_preset(PresetImageKind.PAPER_TEXTURE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set using styles
""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.table.background import PresetImageKind
        # ... other code
        bg_img_style = TblBgImg.from_preset(PresetImageKind.PAPER_TEXTURE)
        table = cursor.add_table(
            table_data=tbl_data,
            first_row_header=False,
            styles=[bg_img_style],
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234122154-552b5eb8-94fe-480c-a1fd-868c95ad240b:

    .. figure:: https://user-images.githubusercontent.com/4193389/234122154-552b5eb8-94fe-480c-a1fd-868c95ad240b.png
        :alt: Table Background Color
        :figclass: align-center
        :width: 520px

        Table Background Color

.. cssclass:: screen_shot

    .. _234122267-fc3697ec-5759-4ea1-bdad-2a71c776df06:

    .. figure:: https://user-images.githubusercontent.com/4193389/234122267-fc3697ec-5759-4ea1-bdad-2a71c776df06.png
        :alt: Table Background Image Dialog
        :figclass: align-center
        :width: 450px

        Table Background Image Dialog


Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_table_borders`
        - :ref:`help_writer_format_direct_table_properties`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_direct_cell_borders`
        - :py:meth: `TextCursorPartial.add_table() <ooodev.write.partial.text_cursor_partial.add_table>`
        - :py:meth:`Write.add_table() <ooodev.office.write.Write.add_table>`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.writer.direct.table.properties.TableProperties`