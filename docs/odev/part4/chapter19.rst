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
        sheetsIdx = Lo.qi(XIndexAccess, sheets)
        sheet = Lo.qi(XSpreadsheet, sheetsIdx.getByIndex(0))

These steps are hidden by methods in the :py:class:`~.calc.Calc` utility class, so the programmer can write:

.. tabs::

    .. code-tab:: python

        loader = Lo.load_office(Lo.ConnectSocket())
        Calc.open_doc(doc_path, loader)
        sheet = Calc.get_sheet(doc, 0)

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

This :py:meth:`.GUI.set_visible` can be called with a XSpreadsheet_ reference: ``GUI.set_visible(True, doc)``.
The document is cast to XComponent_ inside :py:meth:`~.GUI.set_visible` and then processed.

19.3 Spreadsheet Data
=====================

The data in a spreadsheet can be accessed in many ways:
for example, as individual cells, cell ranges, collections of cell ranges, rows, and columns.
These ways of viewing data are supported by different services which are used as labels in :numref:`ch19_sheet_services_data`.

.. cssclass:: diagram invert

    .. _ch19_sheet_services_data:
    .. figure:: https://user-images.githubusercontent.com/4193389/186767178-3366a5d1-e0e8-4a81-8928-c9c1904d602c.png
        :alt: Diagram of Services used with Spreadsheet Data.
        :figclass: align-center

        :Services used with Spreadsheet Data.

The simplest spreadsheet unit is a cell, which can be located by its (column, row) coordinate/position or by its name, as in :numref:`ch19_addressing_cells`.

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

Work in progress ...

.. _Spreadsheets: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1Spreadsheets.html
.. _XComponent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XComponent.html
.. _XSpreadsheetDocument: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSpreadsheetDocument.html
.. _XTextDocument: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextDocument.html
.. _XSpreadsheet: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSpreadsheet.html