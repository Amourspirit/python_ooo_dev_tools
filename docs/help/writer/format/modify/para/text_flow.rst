.. _help_writer_format_modify_para_text_flow:

Write Modify Paragraph Text Flow
================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Specify hyphenation and pagination options.

The :py:class:`ooodev.format.writer.modify.para.text_flow.Breaks`, :py:class:`ooodev.format.writer.modify.para.text_flow.FlowOptions`
and :py:class:`ooodev.format.writer.modify.para.text_flow.Hyphenation` classes is used to set the paragraph text flow.


Default Paragraph Text Flow Style Dialog

.. cssclass:: screen_shot

    .. _234623240-f02a99f3-83e5-49d9-86ad-ea641bfd336d:

    .. figure:: https://user-images.githubusercontent.com/4193389/234623240-f02a99f3-83e5-49d9-86ad-ea641bfd336d.png
        :alt: Writer dialog Paragraph Text Flow default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Text Flow default


Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.para.text_flow import Breaks, FlowOptions, Hyphenation
        from ooodev.format.writer.modify.para.text_flow import StyleParaKind, BreakType
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_hy_style = Hyphenation(
                    auto=True, start_chars=3, end_chars=3, style_name=StyleParaKind.STANDARD
                )
                para_hy_style.apply(doc)

                style_obj = Hyphenation.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Hyphenation Class
-----------------

Setting Hyphenation
^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_hy_style = Hyphenation(
            auto=True, start_chars=3, end_chars=3, style_name=StyleParaKind.STANDARD
        )
        para_hy_style.apply(doc)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234625239-bc127505-5d69-4c3a-8957-4924c524b1c2:

    .. figure:: https://user-images.githubusercontent.com/4193389/234625239-bc127505-5d69-4c3a-8957-4924c524b1c2.png
        :alt: Writer dialog Paragraph Text Flow style changed hyphenation
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Text Flow style changed hyphenation


Getting hyphenation from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Hyphenation.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

FlowOptions Class
-----------------

Setting Options
^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_flow_style = FlowOptions(orphans=3, widows=4, keep=True, style_name=StyleParaKind.STANDARD)
        para_flow_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234626344-4a168449-92a5-4e70-b6e2-97926f1c8c91:

    .. figure:: https://user-images.githubusercontent.com/4193389/234626344-4a168449-92a5-4e70-b6e2-97926f1c8c91.png
        :alt: Writer dialog Paragraph Text Flow style changed flow options
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Text Flow style changed flow options

Getting options from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = FlowOptions.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Breaks Class
-----------------

Setting Breaks Style
^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_break_style = Breaks(
            type=BreakType.PAGE_BEFORE, style="Right Page", style_name=StyleParaKind.STANDARD
        )
        para_break_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234628622-684bba5e-0256-4591-9b69-92dd92da4a7a:
    .. figure:: https://user-images.githubusercontent.com/4193389/234628622-684bba5e-0256-4591-9b69-92dd92da4a7a.png
        :alt: Writer dialog Paragraph Text Flow style changed breaks
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Text Flow style changed breaks

Getting breaks from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Breaks.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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
        - :ref:`help_writer_format_direct_para_text_flow`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.writer.modify.para.text_flow.Breaks`
        - :py:class:`ooodev.format.writer.modify.para.text_flow.FlowOptions`
        - :py:class:`ooodev.format.writer.modify.para.text_flow.Hyphenation`