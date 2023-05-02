.. _help_writer_format_modify_page_header_header:

Write Modify Page Header
========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.page.header.Header` class is used to modify the page header values seen in :numref:`ss_page_header_default_dialog` of a Writer document.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25

        from ooodev.format.writer.modify.page.header import Header, WriterStylePageKind
        from ooodev.utils.color import StandardColor
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
           with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                header_style = Header(
                    on=True,
                    shared_first=True,
                    shared=True,
                    height=10.0,
                    spacing=3.0,
                    spacing_dyn=True,
                    margin_left=1.5,
                    margin_right=2.0,
                    style_name=WriterStylePageKind.STANDARD,
                )
                header_style.apply(doc)

                style_obj = Header.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
                assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Applying Header Style
---------------------

Before applying style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _ss_page_header_default_dialog:
    .. figure:: https://user-images.githubusercontent.com/4193389/235272913-3b6b62ef-b1bf-4b1e-bc7c-ea8afea7e7d4.png
        :alt: Writer dialog Page Header default
        :figclass: align-center
        :width: 450px

        Writer dialog Page Header default

Apply Style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        header_style = Header(
            on=True,
            shared_first=True,
            shared=True,
            height=10.0,
            spacing=3.0,
            spacing_dyn=True,
            margin_left=1.5,
            margin_right=2.0,
            style_name=WriterStylePageKind.STANDARD,
        )
        header_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

After applying style
^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _ss_page_header_dialog_set_with_header_class:
    .. figure:: https://user-images.githubusercontent.com/4193389/235273524-461f2cf4-a48f-4e93-9308-f7283456daf0.png
        :alt: Writer dialog Page Header set with Header class
        :figclass: align-center
        :width: 450px

        Writer dialog Page Header set with Header class


Getting the Header from a style
-------------------------------

.. tabs::

    .. code-tab:: python

        style_obj = Header.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.page.header.Header`