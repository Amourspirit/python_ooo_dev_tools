.. _help_writer_format_style_page:

Write Style Page Class
======================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2


Applying Page Styles can be accomplished using the :py:class:`ooodev.format.writer.style.Page` class.

Setup
-----

General function used to run these examples.

.. must be before the tabs directive
.. include:: ../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            import uno
            from ooodev.format.writer.style import Page, WriterStylePageKind
            from ooodev.format.writer.modify.page.area import Color as PageAreaColor
            from ooodev.gui import GUI
            from ooodev.loader.lo import Lo
            from ooodev.office.write import Write
            from ooodev.utils.color import StandardColor


            def main() -> int:
                p_txt = (
                    |short_ptext|
                )

                with Lo.Loader(Lo.ConnectPipe()):
                    doc = Write.create_doc()
                    GUI.set_visible(doc=doc)
                    Lo.delay(300)
                    GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                    pg_cursor = Write.get_page_cursor(doc)
                    style = Page(name=WriterStylePageKind.FIRST_PAGE)
                    style.apply(pg_cursor)

                    color_style = PageAreaColor(
                        color=StandardColor.GREEN_DARK2, style_name=WriterStylePageKind.FIRST_PAGE
                    )
                    color_style.apply(doc)

                    cursor = Write.get_cursor(doc)
                    Write.append_para(cursor=cursor, text=p_txt)

                    Lo.delay(1_000)

                    Lo.close_doc(doc)

                return 0



            if __name__ == "__main__":
                SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply Style
-----------

In this example we will apply the style ``First Page`` to the page cursor.
Then we will apply the area color ``Green Dark 2`` to the First page Style.
See the :ref:`help_writer_format_modify_page_area` for more information on modifying page area styles.

.. tabs::

    .. code-tab:: python

        # ... other code
        # get the page cursor
        pg_cursor = Write.get_page_cursor(doc)
        style = Page(name=WriterStylePageKind.FIRST_PAGE)
        # apply the style to the page cursor, changing the page style to "First Page"
        style.apply(pg_cursor)

        # create a page area color style to modify the ``First Page`` style with the color ``Green Dark 2``
        color_style = PageAreaColor(
            color=StandardColor.GREEN_DARK2, style_name=WriterStylePageKind.FIRST_PAGE
        )
        color_style.apply(doc)

        # write the paragraph
        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _235803885-06a25a40-fa71-4704-92bd-8e9332a6fd77:
    .. figure:: https://user-images.githubusercontent.com/4193389/235803885-06a25a40-fa71-4704-92bd-8e9332a6fd77.png
        :alt: Styles applied to First Page
        :figclass: align-center
        :width: 550px

        Styles applied to First Page

Get Style from Cursor
---------------------

.. tabs::

    .. code-tab:: python

        # ... other code
        f_style = Page.from_obj(pg_cursor)
        assert f_style.prop_name == style.prop_name

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_page_area`
        - :py:class:`~ooodev.office.write.Write`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.writer.style.Page`
