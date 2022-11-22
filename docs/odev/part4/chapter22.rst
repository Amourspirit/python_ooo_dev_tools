.. _ch22:

******************
Chapter 22. Styles
******************

.. topic:: Overview

    Obtaining Style Information: the TableCellStyle and TablePageStyle Services; Creating and Using a New Style; Adding Borders

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

        ::

            Style Family Names (2)
              CellStyles
              PageStyles

            1. "CellStyles" Family Styles:
            No. of names: 20
              'Accent'  'Accent 1'  'Accent 2'  'Accent 3'
              'Bad'  'Default'  'Error'  'Footnote'
              'Good'  'Heading'  'Heading 1'  'Heading 2'
              'Hyperlink'  'Neutral'  'Note'  'Result'
              'Result2'  'Status'  'Text'  'Warning'

            2. "PageStyles" Family Styles:
            No. of names: 3
              'Default'  'PageStyle_ACPT (Python)'  'Report'


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

The resulting spreadsheet is saved and this document is examined by the |stles_info|_ program, it lists the new styles in the ``CellStyles`` family:

.. cssclass:: rst-collapse

    .. collapse:: Output:
        :open:

        ::

            Style Family Names (2)
              CellStyles
              PageStyles

            1. "CellStyles" Family Styles:
            No. of names: 21
              'Accent'  'Accent 1'  'Accent 2'  'Accent 3'
              'Bad'  'Default'  'Error'  'Footnote'
              'Good'  'Heading'  'Heading 1'  'Heading 2'
              'Hyperlink'  'My DataStyle'  'My HeaderStyle'  'Neutral'
              'Note'  'Result'  'Status'  'Text'
              'Warning'


            2. "PageStyles" Family Styles:
            No. of names: 2
              'Default'  'Report'

Work in progress ...

.. |stles_info| replace:: All Styles Info
.. _stles_info: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_styles_all_info

.. |stles_info_py| replace:: styles_all_info.py
.. _stles_info_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_styles_all_info/styles_all_info.py

.. |build_tbl| replace:: Build Table
.. _build_tbl: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_build_table

.. |build_tbl_py| replace:: build_table.py
.. _build_tbl_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_build_table/build_table.py

.. _TableCellStyle: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1TableCellStyle.html
.. _TablePageStyle: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1TablePageStyle.html
.. _CellProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1CellProperties.html
.. _CharacterProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
.. _ParagraphProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties.html
.. _PageProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1PageProperties.html
