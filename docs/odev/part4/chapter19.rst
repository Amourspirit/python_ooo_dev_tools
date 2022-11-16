.. _ch19:

*****************************
Chapter 19. Calc API Overview
*****************************

.. topic:: Overview

    The Spreadsheet Document; Document Spreadsheets; Spreadsheet Data; The Spreadsheet Service; Cell Range Services; Cell Services; Sheet Cell Ranges

This chapter gives an overview of the main services and interfaces used in the Calc parts of the Office API, illustrated with small code fragments.
We'll revisit these topics in greater details (and with larger examples) in subsequent chapters.

If you're unfamiliar with Calc, then a good starting point is its user guide, available from https://libreoffice.org/get-help/documentation.
Chapter 8 of the Developer's Guide looks at spreadsheet programming, and is available from https://wiki.openoffice.org/w/images/d/d9/DevelopersGuide_OOo3.1.0.pdf.
Alternatively, you can access the chapter online, starting at https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Spreadsheet_Documents (or use ``loguide calc``).
The guide's examples can be found at https://api.libreoffice.org/examples/DevelopersGuide/examples.html#Spreadsheet.
There's also a few examples in the "Spreadsheet Document Examples" section of https://api.libreoffice.org/examples/examples.html#Java_examples.

19.1 The Spreadsheet Document
=============================

..
    figure 1

.. cssclass:: diagram invert

    .. _ch19_some_spreadsheet_serv_interface:
    .. figure:: https://user-images.githubusercontent.com/4193389/186757939-05ed9dda-f593-4db7-bc34-9d742036d962.png
        :alt: Diagram of Some Spreadsheet Services and Interfaces
        :figclass: align-center

        :Some Spreadsheet Services and Interfaces.

Calc's functionality is mostly divided between two Java packages (modules), sheet and table,
which are documented at https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet.html and https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1table.html.
Alternatively, you can try ``lodoc star sheet module`` and ``lodoc star table module``, but these only get you 'close' to the right pages.

The reason for this module division is Office's support for three types of 'table': text tables, database tables, and spreadsheets.
A spreadsheet is a table with formulae added into the mix.

19.2 Document Spreadsheets
==========================

A spreadsheet document (i.e. a Calc file) can consist of multiple spreadsheets (or sheets).
This is implemented using two services – called Spreadsheets (note the 's') and Spreadsheet, as in :numref:`ch19_sheet_doc_hierarchy`.

..
    figure 2

.. cssclass:: diagram invert

    .. _ch19_sheet_doc_hierarchy:
    .. figure:: https://user-images.githubusercontent.com/4193389/186759549-e2851e66-0a52-4d34-abb3-df6f6a1c2bdc.png
        :alt: Diagram of A Spreadsheet Document Hierarchy
        :figclass: align-center

        :A Spreadsheet Document Hierarchy.

The sheets stored in a Spreadsheets_ object can be accessed by index or by name.
A newly created document always contains a blank spreadsheet in index position 0.

The following code fragment shows how the first sheet in the ``test.odt`` document is accessed:


.. tabs::

    .. code-tab:: python

        loader = Lo.load_office(Lo.ConnectSocket())
        compdoc = Lo.open_doc("test.odt", loader)
        doc = Lo.qi(XSpreadsheetDocument, compdoc)
        sheets = doc.getSheets()
        sheets_idx = Lo.qi(XIndexAccess, sheets)
        sheet = Lo.qi(XSpreadsheet, sheets_idx.getByIndex(0))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

These steps are hidden by methods in the :py:class:`~.calc.Calc` utility class, so the programmer can write:

.. tabs::

    .. code-tab:: python

        loader = Lo.load_office(Lo.ConnectSocket())
        Calc.open_doc(doc_path, loader)
        sheet = Calc.get_sheet(doc, 0)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Some Casting Required
---------------------

Surprisingly, XSpreadsheetDocument_ doesn't subclass XComponent_.
This means that it's not possible to pass an XSpreadsheetDocument_ reference to a method expecting an XComponent_ argument:

Text documents can be passed to methods that expect XComponent_ because XTextDocument_ does subclass XComponent_.
The same is possible for Draw and Impress documents.

It's possible to manipulate a spreadsheet document as an XComponent_, but it must be cast first:


.. tabs::

    .. code-tab:: python

        xc = Lo.qi(XComponent, doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This is why casting to XComponent_ is done automatically in  :py:meth:`.GUI.set_visible`.
For example, the ``odoc`` arg of :py:meth:`.GUI.set_visible` assumes that it is of type Object:

.. tabs::

    .. code-tab:: python

        # in GUI class
        @classmethod
        def set_visible(cls, is_visible: bool, odoc: object = None) -> None:
            if odoc is None:
                xwindow = cls.get_window()
            else:
                doc = Lo.qi(XComponent, odoc)
                if doc is None:
                    return
                xwindow = cls.get_frame(doc).getContainerWindow()

            if xwindow is not None:
                xwindow.setVisible(is_visible)
                xwindow.setFocus()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This :py:meth:`.GUI.set_visible` can be called with a XSpreadsheet_ reference: ``GUI.set_visible(True, doc)``.
The document is cast to XComponent_ inside :py:meth:`~.GUI.set_visible` and then processed.

19.3 Spreadsheet Data
=====================

The data in a spreadsheet can be accessed in many ways:
for example, as individual cells, cell ranges, collections of cell ranges, rows, and columns.
These ways of viewing data are supported by different services which are used as labels in :numref:`ch19_sheet_services_data`.

..
    figure 3

.. cssclass:: diagram invert

    .. _ch19_sheet_services_data:
    .. figure:: https://user-images.githubusercontent.com/4193389/186767178-3366a5d1-e0e8-4a81-8928-c9c1904d602c.png
        :alt: Diagram of Services used with Spreadsheet Data.
        :figclass: align-center

        :Services used with Spreadsheet Data.

The simplest spreadsheet unit is a cell, which can be located by its (column, row) coordinate/position or by its name, as in :numref:`ch19_addressing_cells`.

..
    figure 4

.. cssclass:: diagram invert

    .. _ch19_addressing_cells:
    .. figure:: https://user-images.githubusercontent.com/4193389/186767510-244d630f-b2ec-4bbe-aa23-5b0bbb61d77f.png
        :alt: Diagram of Addressing Cells
        :figclass: align-center

        :Addressing Cells.

For instance, the cell named ``C5`` in :numref:`ch19_addressing_cells` is at coordinate ``(2,4)``.
Note that row names start at ``1`` but row positions begin at ``0``.
A cell range is defined by the position of the top-left and bottom-right cells in the range's rectangle, and can use the same dual naming scheme. For example,
the cell range ``B2:D3`` is the rectangle between the cells ``(1,1)`` and ``(3,2)``.

A spreadsheet document may contain multiple sheets, so a cell address can include a sheet name.
The first sheet is called ``Sheet1``, the second ``Sheet2``, and so on.
For example, ``Sheet1.A3:Sheet3.D4`` refers to a cube of 24 cells consisting of 3 sheets of 8 cells between ``A3`` and ``D4``.
Sheets can be assigned more informative names, if you wish.

A collection of cell ranges is defined using ``~`` (the tilde) as the concatenation operator.
For example, ``A1:C3~B2:D2`` is a group of two ranges, ``A1:C3`` and ``B2:D2``.
The comma, ``,``, can be used as an alternative concatenation symbol, at least in some Calc functions.

There's also an intersection operator, ``!``, for calculating the intersection of two ranges.

Cell references can be relative or absolute, which mainly affect how formulae are copied between cells.
For example, a formula ``(=A1*3)`` in cell ``C3`` becomes ``(=B1*3)`` when copied one cell to the right into ``D3``.
However, an absolute reference (which uses ``\`` ( as a prefix) is unaffected when moved.
For instance ``(=\)A$1*3)`` stops the ``A`` and ``1`` from being changed by a move.

The :py:class:`~.calc.Calc` support class includes methods for converting between simple cell names and positions;
they don't handle ``~``, ``!``, or absolute references using ``$``.

19.4 The Spreadsheet Service
============================

The Spreadsheet_ service is a subclass of SheetCellRange_, as shown in :numref:`ch19_spreadsheet_service`, which means that a sheet can be treated as a very big cell range.

..
    figure 5

.. cssclass:: diagram invert

    .. _ch19_spreadsheet_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/186772291-17097766-8fae-42b4-bde3-5e5184ce108d.png
        :alt: Diagram of The Spreadsheet Service
        :figclass: align-center

        :The Spreadsheet Service.

A great deal of spreadsheet-related functionality is implemented as interfaces belonging to the Spreadsheet_ service.
The most important is probably XSpreadsheet_ (see ``lodoc xspreadsheet reference`` ), which gives the programmer access
to a sheet's cells and cell ranges via ``getCellByPosition()``, ``getCellRangeByPosition()``, and ``getCellRangeByName()``. For example:

.. tabs::

    .. code-tab:: python

        sheet = Calc.get_sheet(doc, 0)
        cell = sheet.getCellByPosition(2, 4) # (column,row)

        # startColumn, startRow, endColumn, endRow
        cell_range1 = sheet.getCellRangeByPosition(1, 1, 3, 2)

        cell_range2 = sheet.getCellRangeByName("B2:D3")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Oddly enough there's no ``getCellByName()`` method, but the :py:meth:`.Calc.get_cell` has an overload that takes a name.

19.5 Cell Range Services
========================

The main service for cell ranges is SheetCellRange_, which inherits the CellRange_ service from the table
module and several property-based classes, as indicated in :numref:`ch19_cell_range_service`.

..
    figure 6

.. cssclass:: diagram invert

    .. _ch19_cell_range_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/186776296-3d499331-ded9-4232-bc73-e0eaad08ae33.png
        :alt: Diagram of The Cell Range Services
        :figclass: align-center

        :The Cell Range Services.

SheetCellRange_ supports an XSheetCellRange_ interface, but that interface gets most of its functionality by inheriting XSheetCellRange_ from the table module.
Most programs that manipulate cell ranges tend to use XCellRange_ rather than XSheetCellRange_.

XCellRange_ is where the useful cell and cell range access methods are defined, as shown in the class diagram in :numref:`ch19_cell_range_class`.

..
    figure 7

.. cssclass:: screen_shot invert

    .. _ch19_cell_range_class:
    .. figure:: https://user-images.githubusercontent.com/4193389/186776991-7e4433fb-aee5-4ea8-996f-cae1ec212756.png
        :alt: Screen shot of The Cell Range Class Diagram
        :figclass: align-center

        :The CellRange_ Class Diagram.

You can access the documentation using ``lodoc XCellRange``.

What's missing from XCellRange_ is a way to set the values in a cell range.
This is supported by the XCellRangeData_ interface (see :numref:`ch19_cell_range_service`) which offers a ``setDataArray()`` method (and a ``getDataArray()``).

``CellProperties`` in the table module is frequently accessed to adjust cell styling, such as color, borders, and the justification and
orientation of data inside a cell. However, styling for a cell's text is handled by properties in the ``CharacterProperties`` or ``ParagraphProperties``
classes (see :numref:`ch19_cell_range_service`).

Rows and columns of cells can be accessed using the TableRows_ and TableColumns_ services
(and their corresponding XTableRows_ and XTableColumns_ interfaces).
They're accessed through the XColumnRowRange_ interface shown in :numref:`ch19_cell_range_service`.
Code for obtaining the first row of a sheet is:

.. tabs::

    .. code-tab:: python

        # get the XColumnRowRange interface for the sheet
        cr_range = Lo.qi(XColumnRowRange, sheet)

        # get all the rows
        rows = cr_range.getRows()

        # treat the rows as an indexed container
        con = Lo.qi(XIndexAccess, rows)

        # access the first row as a cell range
        row_range = Lo.qi(XCellRange, con.getByIndex(0));

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

XTableRows_ is an indexed container containing a sequence of XCellRange_ objects.
The TableRow_ services and interfaces are shown in :numref:`ch19_tbl_row_services`:

..
    figure 8

.. cssclass:: diagram invert

    .. _ch19_tbl_row_services:
    .. figure:: https://user-images.githubusercontent.com/4193389/186781411-de179a21-62d6-4e3d-9484-6b4f57a1fd34.png
        :alt: Diagram of The TableRow Services and Interfaces
        :figclass: align-center

        :The TableRow_ Services and Interfaces.

Similar coding is used to retrieve a column: ``XColumnRowRange.getColumns()`` gets all the columns.
:numref:`ch19_tbl_col_services` shows the TableColumn_ services and interfaces.

..
    figure 9

.. cssclass:: diagram invert

    .. _ch19_tbl_col_services:
    .. figure:: https://user-images.githubusercontent.com/4193389/186781802-3180fcea-6c72-483e-89b6-eff0257dd8e2.png
        :alt: Diagram of The TableColumn Services and Interfaces.
        :figclass: align-center

        :The TableColumn_ Services and Interfaces.

:py:class:`~.calc.Calc` class includes methods that hide these details, so the accessing the first row of the sheet becomes:

.. tabs::

    .. code-tab:: python

        row_range = Calc.get_row_range(sheet, 0);

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

19.6 Cell Services
==================

``XCellRange.getCellByPosition()`` returns a single cell from a given cell range.
However, this method can also be applied to a sheet because the API considers a sheet to be a very big cell range.
For example:

.. tabs::

    .. code-tab:: python

        cell = sheet.getCellByPosition(2, 4)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The SheetCell_ service manages properties related to cell formulae and cell input validation.
However, most cell functionality comes from inheriting the Cell service in the table module, and its XCell_ interface.
This arrangement is shown in :numref:`ch19_sheet_cell_services`.

..
    figure 10

.. cssclass:: diagram invert

    .. _ch19_sheet_cell_services:
    .. figure:: https://user-images.githubusercontent.com/4193389/186782922-85e8d39a-bdf9-4dc9-91dc-8623fff1b417.png
        :alt: Diagram of The The SheetCell Services and Interfaces.
        :figclass: align-center

        :The SheetCell_ Services and Interfaces.

SheetCell_ doesn't support an ``XSheetCell`` interface; instead most programming is done using XCell_.
``XCell`` contains useful methods for getting and setting the values in a cell (which may be numbers, text, or formulae).
For example, the following stores the number 9 in the cell at coordinate ``(2, 4)`` (the ``C5`` cell):

.. tabs::

    .. code-tab:: python

        sheet = Calc.get_sheet(doc, 0)
        cell = sheet.getCellByPosition(2, 4) # (column,row)
        cell.setValue(9)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

SheetCell_ inherits the same properties as SheetCellRange_.
For example, ``CellProperties`` stores cell formatting properties, while text styling properties are supported by
``CharacterProperties`` and ``ParagraphProperties`` (see :numref:`ch19_sheet_cell_services`).

The Cell_ service supports both the XCell_ and XText_ interfaces.
Via the XText_ interface, it's possible to manipulate cell text in the same way that text is handled in a text document.
However, for most purposes, it’s enough to use ``XCell's`` ``setFormula()`` which, despite its name,
can be used to assign plain text to a cell. For instance:

.. tabs::

    .. code-tab:: python

        cell.setFormula("hello") # put "hello" text in the cell

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Calc differentiates between ordinary text and formulae by expecting a formula to begin with ``=``.

The XCell_ class diagram is shown in :numref:`ch19_xcell_class`.

..
    figure 11

.. cssclass:: diagram invert

    .. _ch19_xcell_class:
    .. figure:: https://user-images.githubusercontent.com/4193389/186784216-ab5cdd95-df13-4714-960a-83a3102664f3.png
        :alt: Diagram of The XCell Class
        :figclass: align-center

        :The XCell_ Class Diagram.

The documentation for XCell can be found using ``lodoc xcell``.

19.7 Sheet Cell Ranges
======================

A collection of cell ranges has its own service, SheetCellRanges_, shown in :numref:`ch19_sheet_cell_ranges_service`.

..
    figure 12

.. cssclass:: diagram invert

    .. _ch19_sheet_cell_ranges_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/186784624-04ce1f9a-4366-4881-9cb8-ca34cd5405d5.png
        :alt: Diagram of The SheetCellRanges Services and Interfaces.
        :figclass: align-center

        :The SheetCellRanges_ Services and Interfaces.

SheetCellRanges_ doesn't turn up much when programming since it's easy to access multiple cell ranges by accessing them one at a time inside a loop.

.. todo::

    Chapter 19.7, add link to chapter 26.

One major use for SheetCellRanges_ are in sheet searches which return the matching cell ranges in a
XSheetCellRangeContainer_ object. There are examples in Chapter 26.

.. _Cell: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1Cell.html
.. _CellRange: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1CellRange.html
.. _SheetCell: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SheetCell.html
.. _SheetCellRange: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SheetCellRange.html
.. _SheetCellRanges: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SheetCellRanges.html
.. _Spreadsheet: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1Spreadsheet.html
.. _Spreadsheets: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1Spreadsheets.html
.. _TableColumn: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableColumn.html
.. _TableColumns: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableColumns.html
.. _TableRow: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableRow.html
.. _TableRows: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableRows.html
.. _XCell: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCell.html
.. _XCellRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCellRange.html
.. _XCellRangeData: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XCellRangeData.html
.. _XColumnRowRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XColumnRowRange.html
.. _XComponent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XComponent.html
.. _XSheetCellRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSheetCellRange.html
.. _XSheetCellRangeContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSheetCellRangeContainer.html
.. _XSpreadsheet: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSpreadsheet.html
.. _XSpreadsheetDocument: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSpreadsheetDocument.html
.. _XTableColumns: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XTableColumns.html
.. _XTableRows: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XTableRows.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html
.. _XTextDocument: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextDocument.html