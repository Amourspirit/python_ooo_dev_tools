.. _help_writer_format_modify_para_tabs:

Write Modify Paragraph Tabs
===========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.tabs.Tabs`, class is used to set the paragraph style font :numref:`234718068-f6527b7c-49e3-4384-af05-ab5b65fe9158`.

.. cssclass:: screen_shot

    .. _234718068-f6527b7c-49e3-4384-af05-ab5b65fe9158:

    .. figure:: https://user-images.githubusercontent.com/4193389/234718068-f6527b7c-49e3-4384-af05-ab5b65fe9158.png
        :alt: Writer dialog Paragraph Tabs default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Tabs default


Setting Tabs
------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 15, 16, 17, 18

        from ooodev.format.writer.modify.para.tabs import Tabs, TabAlign
        from ooodev.format.writer.modify.para.tabs import FillCharKind, StyleParaKind
        from ooodev.format import Styler
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                tb1 = Tabs(position=12.0, align=TabAlign.DECIMAL)
                tb2 = Tabs(position=6.5, align=TabAlign.CENTER, fill_char="*")
                tb3 = Tabs(position=11.3, align=TabAlign.LEFT, fill_char=FillCharKind.UNDER_SCORE)
                Styler.apply(doc, tb1, tb2, tb3)

                style_obj = Tabs.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following results in the Writer dialog.


.. cssclass:: screen_shot

    .. _234718550-a1907764-4f3c-419d-ac98-42801eb881a7:

    .. figure:: https://user-images.githubusercontent.com/4193389/234718550-a1907764-4f3c-419d-ac98-42801eb881a7.png
        :alt: Writer dialog Paragraph Tabs changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Tabs changed


Getting tabs from a style
-------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Tabs.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_para_tabs`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.tabs.Tabs`