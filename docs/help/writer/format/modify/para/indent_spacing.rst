.. _help_writer_format_modify_para_indent_spacing:

Write Modify Paragraph Indent & Spacing
=======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2


The :py:class:`ooodev.format.writer.modify.para.indent_space.Indent`, :py:class:`ooodev.format.writer.modify.para.indent_space.LineSpacing`,
and :py:class:`ooodev.format.writer.modify.para.indent_space.Spacing` classes are used to set the Indent & Spacing style of a paragraph.


Default Paragraph Indent & Spacing Style Dialog

.. cssclass:: screen_shot

    .. _234613266-c0e6eb88-a924-4b80-b6fb-3db7e52d859c:
    .. figure:: https://user-images.githubusercontent.com/4193389/234613266-c0e6eb88-a924-4b80-b6fb-3db7e52d859c.png
        :alt: Writer dialog Paragraph Indent & Spacing default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Indent & Spacing default


Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.para.indent_space import Indent, Spacing, LineSpacing
        from ooodev.format.writer.modify.para.indent_space import StyleParaKind, ModeKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_indent_style = Indent(
                    before=22.0, after=20.0, first=8.0, style_name=StyleParaKind.STANDARD
                )
                para_indent_style.apply(doc)

                style_obj = Indent.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Indent Class
------------

Setting Indent
^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_indent_style = Indent(before=22.0, after=20.0, first=8.0, style_name=StyleParaKind.STANDARD)
        para_indent_style.apply(doc)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234614646-ca9a3c9f-dfa5-41be-81ac-0e811300ed80:
    .. figure:: https://user-images.githubusercontent.com/4193389/234614646-ca9a3c9f-dfa5-41be-81ac-0e811300ed80.png
        :alt: Writer dialog Paragraph Indent & Spacing style changed indent
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Indent & Spacing style changed indent


Getting indent from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Indent.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Spacing Class
-------------

Setting Spacing
^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_spacing_style = Spacing(above=8.0, below=10.0, style_name=StyleParaKind.STANDARD)
        para_spacing_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234616355-8c595049-ac4b-4b27-a3b6-c9cbff24b6c4:
    .. figure:: https://user-images.githubusercontent.com/4193389/234616355-8c595049-ac4b-4b27-a3b6-c9cbff24b6c4.png
        :alt: Writer dialog Paragraph Indent & Spacing style changed spacing
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Indent & Spacing style changed spacing

Getting spacing from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Spacing.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

LineSpacing Class
-----------------

Setting Line Spacing Style
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_ln_spacing_style = LineSpacing(
            mode=ModeKind.PROPORTIONAL, value=85, style_name=StyleParaKind.STANDARD
        )
        para_ln_spacing_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234617906-3211917c-f926-455d-813f-f39fee06df20:
    .. figure:: https://user-images.githubusercontent.com/4193389/234617906-3211917c-f926-455d-813f-f39fee06df20.png
        :alt: Writer dialog Paragraph Indent & Spacing style changed line spacing
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Indent & Spacing style changed line spacing

Getting line spacing from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = LineSpacing.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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
        - :ref:`help_writer_format_direct_para_indent_spacing`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.indent_space.Indent`
        - :py:class:`ooodev.format.writer.modify.para.indent_space.LineSpacing`
        - :py:class:`ooodev.format.writer.modify.para.indent_space.Spacing`