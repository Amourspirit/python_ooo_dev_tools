.. _help_writer_format_direct_table_background:

Write Direct Table Background
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 3

The :py:class:`ooodev.format.writer.direct.table.background.Color` and :py:class:`ooodev.format.writer.direct.table.background.Img` classes are used to set the background of a table.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.table.background import Color as TblBgColor
        from ooodev.format.writer.direct.table.background import Img as TblBgImg, PresetImageKind
        from ooodev.office.write import Write
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.utils.table_helper import TableHelper


        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_100_PERCENT)
                cursor = Write.get_cursor(doc)

                tbl_data = TableHelper.make_2d_array(num_rows=5, num_cols=5)
                bg_color_style = TblBgColor(StandardColor.LIME_LIGHT3)
                table = Write.add_table(
                    cursor=cursor,
                    table_data=tbl_data,
                    first_row_header=False,
                    styles=[bg_color_style],
                )

                # getting the color object
                tbl_bg_color_style = TblBgColor.from_obj(table)
                assert tbl_bg_color_style is not None

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

Background Color
++++++++++++++++

Setting background color
^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        bg_color_style = TblBgColor(StandardColor.LIME_LIGHT3)
        table = Write.add_table(
            cursor=cursor,
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

.. tabs::

    .. code-tab:: python

        # ... other code
        # getting the table properties
        tbl_bg_color_style = TblBgColor.from_obj(table)
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

.. tabs::

    .. code-tab:: python

        # ... other code
        bg_img_style = TblBgImg.from_preset(PresetImageKind.PAPER_TEXTURE)
        table = Write.add_table(
            cursor=cursor,
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

Getting Background Image Object
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        # getting the table properties
        tbl_bg_color_style = TblBgColor.from_obj(table)
        assert tbl_bg_color_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_table_borders`
        - :ref:`help_writer_format_direct_table_properties`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_direct_cell_borders`
        - :py:meth:`Write.add_table() <ooodev.office.write.Write.add_table>`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.table.properties.TableProperties`