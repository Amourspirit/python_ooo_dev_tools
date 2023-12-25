.. _help_writer_format_modify_para_outline_and_list:

Write Modify Paragraph Outline & List
=====================================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.outline_list.Outline`, :py:class:`ooodev.format.writer.modify.para.outline_list.LineNum`,
and :py:class:`ooodev.format.writer.modify.para.outline_list.ListStyle` classes are used to set the outline, list style, and line numbering of a paragraph.


Default Paragraph Outline & List Style Dialog

.. cssclass:: screen_shot

    .. _234433333-c024a6cf-71fe-47f7-be4c-a3590a24499b:

    .. figure:: https://user-images.githubusercontent.com/4193389/234433333-c024a6cf-71fe-47f7-be4c-a3590a24499b.png
        :alt: Writer dialog Paragraph Outline & List default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Outline & List default


Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.para.outline_list import Outline, LineNum, ListStyle
        from ooodev.format.writer.modify.para.outline_list import LevelKind, StyleParaKind, StyleListKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_outline_style = Outline(level=LevelKind.LEVEL_01, style_name=StyleParaKind.STANDARD)
                para_outline_style.apply(doc)

                style_obj = Outline.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Outline Class
-------------

Setting Outline
^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_outline_style = Outline(level=LevelKind.LEVEL_01, style_name=StyleParaKind.STANDARD)
        para_outline_style.apply(doc)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234434393-f7c5b8fd-8dd4-4c59-a93e-fa1b22305d16:

    .. figure:: https://user-images.githubusercontent.com/4193389/234434393-f7c5b8fd-8dd4-4c59-a93e-fa1b22305d16.png
        :alt: Writer dialog Paragraph Outline & List style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Outline & List style changed


Getting outlne from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Outline.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

ListStyle Class
---------------

Setting List Style
^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_liststyle_style = ListStyle(
            list_style=StyleListKind.NUM_123, style_name=StyleParaKind.STANDARD
        )
        para_liststyle_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234434962-ccc0d8ee-ac17-4314-b7fd-6ed51b433a6a:

    .. figure:: https://user-images.githubusercontent.com/4193389/234434962-ccc0d8ee-ac17-4314-b7fd-6ed51b433a6a.png
        :alt: Writer dialog Paragraph Outline & List style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Outline & List style changed

Getting list style from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border padding from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = ListStyle.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

LineNum Class
-------------

Setting Line Number Style
^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_linenum_style = LineNum(num_start=3, style_name=StyleParaKind.STANDARD)
        para_linenum_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234435651-fb052287-3f55-42ae-8e0f-b53a35499565:

    .. figure:: https://user-images.githubusercontent.com/4193389/234435651-fb052287-3f55-42ae-8e0f-b53a35499565.png
        :alt: Writer dialog Paragraph Outline & List style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Outline & List style changed

Getting list number from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border shadow from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = LineNum.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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
        - :ref:`help_writer_format_direct_para_outline_and_list`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.outline_list.Outline`
        - :py:class:`ooodev.format.writer.modify.para.outline_list.LineNum`
        - :py:class:`ooodev.format.writer.modify.para.outline_list.ListStyle`