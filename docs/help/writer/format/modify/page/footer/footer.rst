.. _help_writer_format_modify_page_footer_footer:

Write Modify Page Footer
========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.page.footer.Footer` class is used to modify the page footer values seen in :numref:`ss_page_footer_default_dialog` of a Writer document.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25

        from ooodev.format.writer.modify.page.footer import Footer, WriterStylePageKind
        from ooodev.utils.color import StandardColor
        from ooodev.office.write import Write
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                footer_style = Footer(
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
                footer_style.apply(doc)

                style_obj = Footer.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
                assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Applying Footer Style
---------------------

Before applying style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _ss_page_footer_default_dialog:
    .. figure:: https://user-images.githubusercontent.com/4193389/235278066-fcd9edde-360c-4010-8fed-6e14812d95b4.png
        :alt: Writer dialog Page Footer default
        :figclass: align-center
        :width: 450px

        Writer dialog Page Footer default

Apply Style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        footer_style = Footer(
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
        footer_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

After applying style
^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _ss_page_footer_dialog_set_with_footer_class:
    .. figure:: https://user-images.githubusercontent.com/4193389/235278243-902386fe-8c10-40e3-b596-451d2c290160.png
        :alt: Writer dialog Page Footer set with Footer class
        :figclass: align-center
        :width: 450px

        Writer dialog Page Footer set with Footer class


Getting the Footer from a style
-------------------------------

.. tabs::

    .. code-tab:: python

        style_obj = Footer.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.writer.modify.page.footer.Footer`