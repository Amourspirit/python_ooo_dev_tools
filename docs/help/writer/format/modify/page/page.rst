.. _help_writer_format_modify_page_page:

Write Modify Page Page
======================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.page.page.PaperFormat`, :py:class:`ooodev.format.writer.modify.page.page.Margins`,
and :py:class:`ooodev.format.writer.modify.page.page.LayoutSettings` classes are used to modify the Page style values
seen in :numref:`234897302-ed14b309-d65b-42e3-841a-955a583c3d13` of a Page style.

.. cssclass:: screen_shot

    .. _234897302-ed14b309-d65b-42e3-841a-955a583c3d13:
    .. figure:: https://user-images.githubusercontent.com/4193389/234897302-ed14b309-d65b-42e3-841a-955a583c3d13.png
        :alt: Writer dialog Page default
        :figclass: align-center
        :width: 450px

        Writer dialog Page default

Default Page Style Dialog

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.page import LayoutSettings, Margins, PaperFormat
        from ooodev.format.writer.modify.page.page import PageStyleLayout, NumberingTypeEnum
        from ooodev.format.writer.modify.page.page import PaperFormatKind, WriterStylePageKind
        from ooodev.utils.data_type.size_mm import SizeMM
        from ooodev.units import UnitInch
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                layout_style = LayoutSettings(
                    layout=PageStyleLayout.MIRRORED,
                    numbers=NumberingTypeEnum.CHARS_UPPER_LETTER,
                    ref_style=WriterStylePageKind.TEXT_BODY,
                    style_name=WriterStylePageKind.STANDARD,
                )
                layout_style.apply(doc)

                style_obj = LayoutSettings.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
                assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Page Paper Format
-----------------

The :py:class:`~ooodev.format.writer.modify.page.page.PaperFormat` class is used to modify the paper format of a page style.

Setting Paper Format from a preset
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.inner.preset.preset_paper_format.PaperFormatKind` class is used to look up the preset of paper format for convenience.

.. tabs::

    .. code-tab:: python

        # ... other code

        paper_fmt_style = PaperFormat.from_preset(
            preset=PaperFormatKind.A3, landscape=True, style_name=WriterStylePageKind.STANDARD
        )
        paper_fmt_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234900373-52f17983-34cc-459e-a064-6b0f828b26ef:
    .. figure:: https://user-images.githubusercontent.com/4193389/234900373-52f17983-34cc-459e-a064-6b0f828b26ef.png
        :alt: Writer dialog Page style Paper Format changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page style Paper Format changed

Setting Paper to Custom Format
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

It is possible to set the page to a custom format by using the :py:class:`~ooodev.format.writer.modify.page.page.PaperFormat` class.

The constructor takes a :py:class:`~ooodev.utils.data_type.size_mm.SizeMM` object which can also take :ref:`proto_unit_obj` object for ``width`` and ``height``.
The :py:class:`~ooodev.units.UnitInch` supports ``UnitT`` and is used to set the page size in inches.

If the ``width`` is greater than the ``height`` then the page is set to landscape; Otherwise, the page is set to portrait.

.. tabs::

    .. code-tab:: python

        # ... other code

        paper_fmt_style = PaperFormat(
            size=SizeMM(width=UnitInch(11), height=UnitInch(8.5)),
            style_name=WriterStylePageKind.STANDARD,
        )
        paper_fmt_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234911812-3c0ec32e-35f9-4c45-b38b-950525703d2a:
    .. figure:: https://user-images.githubusercontent.com/4193389/234911812-3c0ec32e-35f9-4c45-b38b-950525703d2a.png
        :alt: Writer dialog Page style Paper Format changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page style Paper Format changed


Getting paper format from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border sides from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = PaperFormat.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Page Margins
------------

Setting Page Margins using MM units
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.page.Margins` class is used to modify the margins of a page style.
In this case the margins are set to ``mm`` values which is the default unit of the class.
The result are seen in :numref:`234916023-c12a16b9-02d2-420f-8da5-c4a6a5bb597b`.

.. tabs::

    .. code-tab:: python

        # ... other code

        margin_style = Margins(
            left=30,
            right=30,
            top=35,
            bottom=15,
            gutter=10,
            style_name=WriterStylePageKind.STANDARD,
        )
        margin_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234916023-c12a16b9-02d2-420f-8da5-c4a6a5bb597b:
    .. figure:: https://user-images.githubusercontent.com/4193389/234916023-c12a16b9-02d2-420f-8da5-c4a6a5bb597b.png
        :alt: Writer dialog Page margins style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page margins style changed

Setting Page Margins using other units
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The margins can be set using a different unit. The parameters used to set the margin size also support :ref:`proto_unit_obj` objects.
The :py:class:`~ooodev.units.UnitInch` supports ``UnitT`` and is used to set the page margin in inches.
The result are seen in :numref:`234917591-f9e4deb2-e4b0-4f42-832f-fb43222c7635`.

.. tabs::

    .. code-tab:: python

        # ... other code

        margin_style = Margins(
            left=UnitInch(1.0),
            right=UnitInch(1.0),
            top=UnitInch(1.5),
            bottom=UnitInch(0.75),
            gutter=UnitInch(0.5),
            style_name=WriterStylePageKind.STANDARD,
        )
        margin_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234917591-f9e4deb2-e4b0-4f42-832f-fb43222c7635:
    .. figure:: https://user-images.githubusercontent.com/4193389/234917591-f9e4deb2-e4b0-4f42-832f-fb43222c7635.png
        :alt: Writer dialog Page margins style set using inches
        :figclass: align-center
        :width: 450px

        Writer dialog Page margins style set using inches

Getting margins from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Margins.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Page Layout
-----------

Setting Page Layout
^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.page.LayoutSettings` class is used to modify the layout of a page style.
The result are seen in :numref:`235153674-56569ad7-6e77-4e42-9ef4-ab362582eda5`.

.. tabs::

    .. code-tab:: python

        # ... other code

        layout_style = LayoutSettings(
            layout=PageStyleLayout.MIRRORED,
            numbers=NumberingTypeEnum.CHARS_UPPER_LETTER,
            ref_style=WriterStylePageKind.SUBTITLE,
            right_gutter=True,
            gutter_pos_left=False,
            style_name=WriterStylePageKind.STANDARD,
        )
        layout_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235153674-56569ad7-6e77-4e42-9ef4-ab362582eda5:
    .. figure:: https://user-images.githubusercontent.com/4193389/235153674-56569ad7-6e77-4e42-9ef4-ab362582eda5.png
        :alt: Writer dialog Page Layout style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page Layout style changed

Getting layout from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^


.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = LayoutSettings.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
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
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.page.page.PaperFormat`
        - :py:class:`ooodev.format.writer.modify.page.page.Margins`
        - :py:class:`ooodev.format.writer.modify.page.page.LayoutSettings`