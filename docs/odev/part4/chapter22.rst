.. _ch22:

******************
Chapter 22. Styles
******************

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

.. topic:: Overview

    Obtaining Style Information: the TableCellStyle_ and TablePageStyle_ Services; Creating and Using a New Style; Adding Borders

    Examples: |stles_info|_ and |build_tbl|_.

This chapter looks at how spreadsheet styles are stored, how they can be examined, and how new styles can be instantiated and used.

.. _ch22_get_style_info:

22.1 Obtaining Style Information
================================

Calc uses the same style API as Writer, Draw, and Impress documents.
:numref:`ch22fig_calc_style_familes_props` shows its structure.

..
    figure 1

.. cssclass:: diagram invert

    .. _ch22fig_calc_style_familes_props:
    .. figure:: https://user-images.githubusercontent.com/4193389/203393811-9fb81f81-15c4-4ab8-9bd0-7c2a234532bd.png
        :alt: Calc Style Families and their Property Sets
        :figclass: align-center

        :Calc Style Families and their Property Sets

The Calc API only has two style families, ``CellStyles`` and ``PageStyles``.
A cell style can be applied to a cell, a cell range, or a spreadsheet (which is a very big cell range).
A page style can be applied only to a spreadsheet.

Each style family consists of styles, which are collection of property sets.
The default cell style is called ``Default``, as is the default page style.

The |stles_info_py|_ example prints out the style family names and the style names associated with the input document:

.. tabs::

    .. code-tab:: python

        # in styles_all_info.py
        from __future__ import annotations

        import uno
        from com.sun.star.sheet import XSpreadsheetDocument

        from ooodev.office.calc import Calc
        from ooodev.utils.file_io import FileIO
        from ooodev.utils.info import Info
        from ooodev.utils.lo import Lo
        from ooodev.utils.props import Props
        from ooodev.utils.type_var import PathOrStr


        class StylesAllInfo:
            def __init__(self, fnm: PathOrStr, rpt_cell_styles: bool) -> None:
                _ = FileIO.is_exist_file(fnm, True)
                self._fnm = FileIO.get_absolute_path(fnm)
                self._rpt_cell_styles = rpt_cell_styles

            def main(self) -> None:
                with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:

                    doc = Calc.open_doc(fnm=self._fnm, loader=loader)
                    try:

                        # get all the style families for this document
                        style_families = Info.get_style_family_names(doc)
                        print(f"Style Family Names ({len(style_families)})")
                        for style_family in style_families:
                            print(f"  {style_family}")
                        print()

                        # list all the style names for each style family
                        for i, style_family in enumerate(style_families):
                            print(f'{i + 1}. "{style_family}" Family Styles:')
                            style_names = Info.get_style_names(
                                doc=doc, family_style_name=style_family
                            )
                            Lo.print_names(style_names)

                        if self._rpt_cell_styles:
                            print()
                            self._report_cell_styles(doc)

                    except Exception:
                        raise

                    finally:
                        Lo.close_doc(doc=doc, deliver_ownership=True)

            def _report_cell_styles(self, doc: XSpreadsheetDocument) -> None:
                Props.show_props(
                    "CellStyles Default", Info.get_style_props(
                        doc=doc, family_style_name="CellStyles", prop_set_nm="Default"
                    )
                )

                Props.show_props(
                    "PageStyles Default", Info.get_style_props(
                        doc=doc, family_style_name="PageStyles", prop_set_nm="Default"
                    )
                )


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


This code uses the :py:meth:`.Info.get_style_family_names` and :py:meth:`.Info.get_style_names` functions that is utilized in earlier chapters, so won't explain their implementation again.
The output for a simple spreadsheet is:

.. cssclass:: rst-collapse

    .. collapse:: Output:
        :open:

        .. code::

            Style Family Names (2)
              CellStyles
              PageStyles

            1. "CellStyles" Family Styles:
            No. of names: 20
              ----------|-----------|-----------|-----------
              Accent    | Accent 1  | Accent 2  | Accent 3
              Bad       | Default   | Error     | Footnote
              Good      | Heading   | Heading 1 | Heading 2
              Good      | Heading   | Heading 1 | Heading 2
              Hyperlink | Neutral   | Note      | Result
              Result2   | Status    | Text      | Warning



            2. "PageStyles" Family Styles:
            No. of names: 3
              ------------------------|-------------------------|-------------------------
              Default                 | PageStyle_ACPT (Python) | Report


.. _ch22_finding_style_info:

Finding Style Information
-------------------------

From a programming point of view, the main difficult with styles is finding documentation about their properties, so that a value can be correctly read or changed.

One approach is to use :py:meth:`.Info.get_style_props` method to list the properties for a given style family and style name.
For example, the ``_report_cell_styles()`` from above from  displays all the properties for the default cell and page styles:

The problem is that the output for ``_report_cell_styles()`` is extremely long, and some property names are less descriptive/understandable than others.

It's probably better to search the online documentation for properties. Cell styles are managed by the TableCellStyle_ service (see :numref:`ch22fig_table_cell_style_srv`) and
page styles by the TablePageStyle_ service (:numref:`ch22fig_table_page_style_srv`).

The properties managed by TableCellStyle_ are inherited from a number of places, as summarized by :numref:`ch22fig_table_cell_style_srv`.

..
    figure 2

.. cssclass:: diagram invert

    .. _ch22fig_table_cell_style_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/203403383-9444c075-1b7f-4a98-938c-a04b022d8515.png
        :alt: The Table Cell Style Service
        :figclass: align-center

        :The TableCellStyle_ Service

By far the most important source of cell style properties is the CellProperties_ class in the table module.
However, if a property relates to the text in a cell then it's more likely to originate from the CharacterProperties_ or
ParagraphProperties_ classes in the style module.

The properties managed by TablePageStyle_ are also inherited from a few places, as summarized by :numref:`ch22fig_table_page_style_srv`.

..
    figure 3

.. cssclass:: diagram invert

    .. _ch22fig_table_page_style_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/203404158-3d603d27-34b2-4db8-bb70-a8434a5cde65.png
        :alt: The Table Page Style Service.
        :figclass: align-center

        :The TablePageStyle_ Service.


The main place to look for page properties is the PageProperties_ class in the style module.
The properties relate to things such as page margins, headers, and footers, which become important when printing a sheet.

.. _ch22_create_new_styles:

22.2 Creating and Using New Styles
==================================

The steps required in creating and using a new style are illustrated by |build_tbl_py|_, in ``_create_styles()`` and ``_apply_styles()``:

.. tabs::

    .. code-tab:: python

        # in build_table.py
        class BuildTable:
            HEADER_STYLE_NAME = "My HeaderStyle"
            DATA_STYLE_NAME = "My DataStyle"

            def main(self) -> None:
                loader = Lo.load_office(Lo.ConnectSocket())

                try:
                    doc = Calc.create_doc(loader)
                    GUI.set_visible(is_visible=True, odoc=doc)
                    sheet = Calc.get_sheet(doc=doc, index=0)
                    self._convert_addresses(sheet)

                    self._build_array(sheet)

                    # ...

                    if self._add_style:
                        self._create_styles(doc)
                        self._apply_styles(sheet)
                # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``_create_styles()`` creates two cell styles called ``My HeaderStyle`` and ``My DataStyle``, which are applied to the spreadsheet by ``_apply_styles()``.
The result is shown in :numref:`ch22fig_styles_sheet_cells`.

..
    figure 4

.. cssclass:: screen_shot

    .. _ch22fig_styles_sheet_cells:
    .. figure:: https://user-images.githubusercontent.com/4193389/203407362-37312fdb-5e51-4e1a-ac54-1af08acecf42.png
        :alt: Styled Spreadsheet Cells
        :figclass: align-center

        :Styled Spreadsheet Cells.

The ``My HeaderStyle`` style is applied to the top row and the first column: the cells are colored blue, and the text made white and centered.
The ``My DataStyle`` is used for the numerical data and formulae cells: the background color is made a light blue, and the text is right-justified.
``_apply_styles()`` also changes the border properties of the bottom edges of the cells in the last row to be thick and blue.

If the resulting spreadsheet is saved and this document is examined by the |stles_info|_ program, it lists the new styles in the ``CellStyles`` family:

.. cssclass:: rst-collapse

    .. collapse:: Output:
        :open:

        ::

            Style Family Names (2)
              CellStyles
              PageStyles

            1. "CellStyles" Family Styles:
            No. of names: 21
              ---------------|----------------|----------------|----------------
              Accent         | Accent 1       | Accent 2       | Accent 3
              Bad            | Default        | Error          | Footnote
              Good           | Heading        | Heading 1      | Heading 2
              Hyperlink      | My DataStyle   | My HeaderStyle | Neutral
              Note           | Result         | Status         | Text
              Warning



            2. "PageStyles" Family Styles:
            No. of names: 2
              --------|---------
              Default | Report

.. _ch22_creating_new_style:

22.2.1 Creating a New Style
---------------------------

|build_tbl_py|_ calls ``_create_styles()`` to create two styles:

.. tabs::

    .. code-tab:: python

        # in build_table.py
        def _create_styles(self, doc: XSpreadsheetDocument) -> None:
            try:
                # create a style using Calc
                header_style = Calc.create_cell_style(
                    doc=doc, style_name=BuildTable.HEADER_STYLE_NAME
                )

                # create formats to apply to header_style
                header_bg_color_style = BgColor(
                    color=CommonColor.ROYAL_BLUE, style_name=BuildTable.HEADER_STYLE_NAME
                )
                effects_style = FontEffects(
                    color=CommonColor.WHITE, style_name=BuildTable.HEADER_STYLE_NAME
                )
                txt_align_style = TextAlign(
                    hori_align=HoriAlignKind.CENTER,
                    vert_align=VertAlignKind.MIDDLE,
                    style_name=BuildTable.HEADER_STYLE_NAME,
                )
                # Apply formatting to header_style
                Styler.apply(
                    header_style, header_bg_color_style, effects_style, txt_align_style
                )

                # create style
                data_style = Calc.create_cell_style(doc=doc, style_name=BuildTable.DATA_STYLE_NAME)

                # create formats to apply to data_style
                footer_bg_color_style = BgColor(
                    color=CommonColor.LIGHT_BLUE, style_name=BuildTable.DATA_STYLE_NAME
                )
                bdr_style = modify_borders.Borders(
                    padding=modify_borders.Padding(left=UnitMM(5))
                )

                # Apply formatting to data_style
                Styler.apply(
                    data_style, footer_bg_color_style, bdr_style, txt_align_style
                )

            except Exception as e:
                print(e)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The styles are created by two calls to :py:meth:`.Calc.create_cell_style`, which stores them in the ``CellStyles`` family:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @staticmethod
        def create_cell_style(doc: XSpreadsheetDocument, style_name: str) -> XStyle:
            comp_doc = Lo.qi(XComponent, doc, raise_err=True)
            style_families = Info.get_style_container(doc=comp_doc, family_style_name="CellStyles")
            style = Lo.create_instance_msf(XStyle, "com.sun.star.style.CellStyle", raise_err=True)

            try:
                style_families.insertByName(style_name, style)
                return style
            except Exception as e:
                raise Exception(f"Unable to create style: {style_name}") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.create_cell_style` calls :py:meth:`.Info.get_style_container` to return a reference to the ``CellStyles`` family as an XNameContainer_.
A new cell style is created by calling :py:meth:`.Lo.create_instance_msf`, and referred to using the XStyle_ interface.
This style is added to the family by calling ``XNameContainer.insertByName()`` with the name passed to the function.

A new style is automatically derived from the ``Default`` style, so the rest of the ``_create_styles()`` method involves the changing of properties.
Five are adjusted in the ``My HeaderStyle`` style, and three in ``My DataStyle``.

The header properties are ``IsCellBackgroundTransparent``, ``CellBackColor``, ``CharColor``, ``HoriJustify``, and ``VertJustify``,
which are all defined in the CellProperties_ class (see :numref:`ch22fig_table_cell_style_srv`).

The data properties are ``IsCellBackgroundTransparent``, ``CellBackColor``, and ``ParaRightMargin``.
Although ``IsCellBackgroundTransparent`` and ``CellBackColor`` are from the CellProperties_ class,
``ParaRightMargin`` is inherited from the ParagraphProperties_ class in the style module (also in :numref:`ch22fig_table_cell_style_srv`).

.. _ch22_appling_new_style:

22.2.2 Applying a New Style
---------------------------

The new styles, ``My HeaderStyle`` and ``My DataStyle``, are applied to the spreadsheet by the |build_tbl_py|_ method ``_apply_styles()``:

.. tabs::

    .. code-tab:: python

        # in build_table.py
        def _apply_styles(self, sheet: XSpreadsheet) -> None:

            Calc.change_style(
                sheet=sheet, style_name=BuildTable.HEADER_STYLE_NAME, range_name="B1:N1"
            )
            Calc.change_style(
                sheet=sheet, style_name=BuildTable.HEADER_STYLE_NAME, range_name="A2:A4"
            )
            Calc.change_style(
                sheet=sheet, style_name=BuildTable.DATA_STYLE_NAME, range_name="B2:N4"
            )

            # ... other code

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The header style is applied to two cell ranges: ``B1:N1`` is the top row containing the months (see :numref:`ch22fig_styles_sheet_cells`),
and ``A2:A4`` is the first column. The data style is applied to ``B2:N4`` which spans the numerical data and formulae.

.. tabs::

    .. code-tab:: python

        # in Calc class (overload method, simplified)
        @classmethod
        def change_style(cls, sheet: XSpreadsheet, style_name: str, range_name: str) -> bool:
            cell_range = cls.get_cell_range(sheet=sheet, range_name=range_name)
            Props.set(cell_range, CellStyle=style_name)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`change_style`

:py:meth:`.Calc.change_styles` manipulates the styles via the cell range.
The ``cell_range`` variable refers to a SheetCellRange_ service which inherits many properties, including those from CellProperties_.
Its ``CellStyle`` property holds the style name used by that cell range.

.. _ch22_adding_borders:

22.2.3 Adding Borders
---------------------

The :py:meth:`.Calc.add_border` method highlights borders for a given range of cells.
The two calls in ``_apply_styles()`` draw a blue line along the bottom edge of the ``A4:N4`` cell range,
and two lines on either side of the ``SUM`` column (the ``N1:N4`` range), as shown in :numref:`ch22fig_borders_and_data`.

..
    figure 5

.. cssclass:: screen_shot

    .. _ch22fig_borders_and_data:
    .. figure:: https://user-images.githubusercontent.com/4193389/203636217-7e487405-0a05-4642-86fc-dae32137708f.png
        :alt: Borders around the Data spreadsheet screen shot.
        :figclass: align-center

        :Borders around the Data.

The border style is applied to the bottom row of the table, and the right column.

Using the :py:mod:`ooodev.format.calc.direct.cell.borders` module (imported as ``direct_borders`` in the code), the border style is created by calling
:py:class:`~ooodev.format.inner.direct.structs.side.Side` class. The side has a width of ``2.85`` points, and a color of ``CommonColor.DARK_BLUE``.

The side is applied to the bottom of the ``A4:N4`` range by creating a :py:class:`~ooodev.format.calc.direct.cell.borders.Borders` object,
and to the left and right of the ``N1:N4`` range by creating a second :py:class:`~ooodev.format.calc.direct.cell.borders.Borders` object.

.. tabs::

    .. code-tab:: python

        # in build_table.py
        from ooodev.format.calc.direct.cell import borders as direct_borders
        # ... other imports

        def _apply_styles(self, sheet: XSpreadsheet) -> None:

            # ... other code

            # create a border side, default width units are points
            side = direct_borders.Side(width=2.85, color=CommonColor.DARK_BLUE)
            # create a border setting bottom side
            bdr = direct_borders.Borders(bottom=side)
            # Apply border to range
            Calc.set_style_range(sheet=sheet, range_name="A4:N4", styles=[bdr])

            # create a border with left and right
            bdr = direct_borders.Borders(left=side, right=side)
            # Apply border to range
            Calc.set_style_range(sheet=sheet, range_name="N1:N4", styles=[bdr])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. |stles_info| replace:: All Styles Info
.. _stles_info: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_styles_all_info

.. |stles_info_py| replace:: styles_all_info.py
.. _stles_info_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_styles_all_info/styles_all_info.py

.. |build_tbl| replace:: Build Table
.. _build_tbl: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_build_table

.. |build_tbl_py| replace:: build_table.py
.. _build_tbl_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_build_table/build_table.py

.. _BorderLine2: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1table_1_1BorderLine2.html
.. _BorderLine2: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1table_1_1BorderLine2.html
.. _CellProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1CellProperties.html
.. _CharacterProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
.. _PageProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1PageProperties.html
.. _ParagraphProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties.html
.. _SheetCellRange: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SheetCellRange.html
.. _TableBorder2: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1table_1_1TableBorder2.html
.. _TableCellStyle: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1TableCellStyle.html
.. _TablePageStyle: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1TablePageStyle.html
.. _XCellRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCellRange.html
.. _XNameContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameContainer.html
.. _XStyle: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1style_1_1XStyle.html
