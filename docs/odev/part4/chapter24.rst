.. _ch24:

*************************************
Chapter 24. Complex Data Manipulation
*************************************

.. topic:: Overview

    Sorting Data; Generating Data: Automatic, LINEAR Mode, DATE Mode, GROWTH Mode; Fancy Text: borders, headlines, hyperlinks, annotations

    Examples: |c_txt|_, |filler|_, and |d_sort|_

This chapter looks at a variety of less common text manipulation techniques, including
the sorting of data, generating data based on examples, and the use of borders, headlines, hyperlinks, and annotations in cells.

.. _ch24_sorting_data:

24.1 Sorting Data
=================

Sorting is available through :spelling:word:`SheetCellRange's` XSortable_ interface. There are four basic steps required for sorting a block of cells:

1. Obtain an XSortable_ interface for the cell range;
2. Specify the sorting criteria as a TableSortField_ sequence;
3. Create a sort descriptor;
4. Execute the sort.

These steps are illustrated by the |d_sort_py|_ example, which begins by building a small table:

.. tabs::

    .. code-tab:: python

        # in data_sort.py
        def main(self) -> None:
            loader = Lo.load_office(Lo.ConnectSocket())

            try:
                doc = Calc.create_doc(loader)

                GUI.set_visible(is_visible=True, odoc=doc)

                sheet = Calc.get_sheet(doc=doc, index=0)
                
                # create the table that needs sorting
                vals = (
                    ("Level", "Code", "No.", "Team", "Name"),
                    ("BS", 20, 4, "B", "Elle"),
                    ("BS", 20, 6, "C", "Sweet"),
                    ("BS", 20, 2, "A", "Chcomic"),
                    ("CS", 30, 5, "A", "Ally"),
                    ("MS", 10, 1, "A", "Joker"),
                    ("MS", 10, 3, "B", "Kevin"),
                    ("CS", 30, 7, "C", "Tom"),
                )
                Calc.set_array(values=vals, sheet=sheet, name="A1:E8")  # or just "A1"

                # 1. obtain an XSortable interface for the cell range
                source_range = Calc.get_cell_range(sheet=sheet, range_name="A1:E8")
                xsort = Lo.qi(XSortable, source_range, True)

                # 2. specify the sorting criteria as a TableSortField array
                sort_fields = (self._make_sort_asc(1, True), self._make_sort_asc(2, True))

                # 3. define a sort descriptor
                props = Props.make_props(SortFields=Props.any(*sort_fields), ContainsHeader=True)

                Lo.wait(2_000)  # wait so user can see original before it is sorted
                # 4. do the sort
                print("Sorting...")
                xsort.sort(props)

                if self._out_fnm:
                    Lo.save_doc(doc=doc, fnm=self._out_fnm)

                msg_result = MsgBox.msgbox(
                    "Do you wish to close document?",
                    "All done",
                    boxtype=MessageBoxType.QUERYBOX,
                    buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
                )
                if msg_result == MessageBoxResultsEnum.YES:
                    Lo.close_doc(doc=doc, deliver_ownership=True)
                    Lo.close_office()
                else:
                    print("Keeping document open")

            except Exception:
                Lo.close_office()
                raise

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The unsorted table is shown in :numref:`ch24fig_tbl_unsorted`.

..
    figure 1

.. cssclass:: screen_shot invert

    .. _ch24fig_tbl_unsorted:
    .. figure:: https://user-images.githubusercontent.com/4193389/204056363-ee551ac7-e25d-4909-bcb3-be4d60486ffa.png
        :alt: An Unsorted Table.
        :figclass: align-center

        :An Unsorted Table.

The table is sorted so that its rows are in ascending order depending on their "Code" column values.
When two rows have the same code number then the sort uses the "No." column.
:numref:`ch24fig_tbl_sorted` shows the result of applying these two sorting criteria:

..
    figure 2

.. cssclass:: screen_shot invert

    .. _ch24fig_tbl_sorted:
    .. figure:: https://user-images.githubusercontent.com/4193389/204056473-b3cc541a-5631-4fb7-b075-38e679291965.png
        :alt: The Sorted Table, Using Two Sort Criteria.
        :figclass: align-center

        :The Sorted Table, Using Two Sort Criteria.

The four sorting steps mentioned above are implemented like so:

.. tabs::

    .. code-tab:: python

        # in data_sort.py
        # ...
        # 1. obtain an XSortable interface for the cell range
        source_range = Calc.get_cell_range(sheet=sheet, range_name="A1:E8")
        xsort = Lo.qi(XSortable, source_range, True)

        # 2. specify the sorting criteria as a TableSortField array
        sort_fields = (self._make_sort_asc(1, True), self._make_sort_asc(2, True))

        # 3. define a sort descriptor
        props = Props.make_props(SortFields=Props.any(*sort_fields), ContainsHeader=True)

        Lo.wait(2_000)  # wait so user can see original before it is sorted
        # 4. do the sort
        print("Sorting...")
        xsort.sort(props)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The ``A1:E8`` cell range referenced using the XCellRange_ interface is converted to XSortable_.
This interface is defined in Office's util module, not in sheet or table, probably because it's also used in text documents for sorting tables.

The two sorting criteria are represented by two TableSortField_ objects in tuple.
The ``_make_sort_asc()`` function is defined in |d_sort_py| as:

.. tabs::

    .. code-tab:: python

        # in data_sort.py
        def _make_sort_asc(self, index: int, is_ascending: bool) -> TableSortField:
            return TableSortField(Field=index, IsAscending=is_ascending, IsCaseSensitive=False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. note::

    Because TableSortField_ is imported with |ooouno|_ (``from ooo.dyn.table.table_sort_field import TableSortField``)
    ``TableSortField`` can be created using Keyword arguments. This a feature added by |ooouno|_ for ``uno structs``.
    Normally ``uno`` objects only take positional only arguments.

A sort descriptor is an array of PropertyValue_ objects which affect how ``XSortable.sort()`` executes.
The most commonly used properties are ``SortFields`` and ``ContainsHeader``. ``SortFields`` is assigned the sorting criteria (:abbreviation:`i.e.` the TableSortField_ tuple),
and the ``ContainsHeader`` boolean specifies whether the sort should exclude the first row because it contains header text.

The sort descriptor properties are defined in a number of classes (SortDescriptor2_, TableSortDescriptor2_, and TextSortDescriptor2_),
which are most easily accessed from the XSortable_ documentation page.

.. _ch24_generating_data:

24.2 Generating Data
====================

Cell data is generated by supplying numbers to a function which treats them as the initial values in a arithmetic (or geometric) series.
The function employs the series to churn out as many more numbers as are needed to fill a given cell range.

A series is created by the XCellSeries_ interface, which is part of the SheetCellRange_ service (see :numref:`ch24fig_cell_rng_services`).

..
    figure 3

.. cssclass:: diagram invert

    .. _ch24fig_cell_rng_services:
    .. figure:: https://user-images.githubusercontent.com/4193389/204058012-b3dc13c8-1fa9-40d7-8e8f-6a271ba60fcc.png
        :alt: The Cell Range Services
        :figclass: align-center

        :The Cell Range Services.

Several examples of how to use ``XCellSeries'`` two methods, ``fillAuto()`` and ``fillSeries()``, are contained in the |filler_py|_ example described next.

|filler_py|_ starts by filling a blank sheet with an assortment of data, which will be used by the XCellSeries_ methods to initialize several series.
The original sheet is shown in :numref:`ch24fig_filler_py_sheet_default`.

..
    figure 4

.. cssclass:: screen_shot invert

    .. _ch24fig_filler_py_sheet_default:
    .. figure:: https://user-images.githubusercontent.com/4193389/204058288-e0853694-ed56-4b88-8804-4dba1b5fb18b.png
        :alt: The filler.py Sheet before Data Generation
        :figclass: align-center

        :The |filler_py|_ Sheet before Data Generation.

The simpler of the two XCellSeries_ methods, ``XCellSeries.fillAuto()``, requires a cell range, fill direction, and how many cells should be examined as 'seeds'.
For example, rows ``7``, ``8``, and ``9`` of :numref:`ch24fig_filler_py_sheet_default` are filled using:

.. tabs::

    .. code-tab:: python

        # in Filler._fill_series() of filler.py
        # set first two values of three rows

        # ascending integers: 1, 2
        Calc.set_val(sheet=sheet, cell_name="B7", value=2)
        Calc.set_val(sheet=sheet, cell_name="A7", value=1)

        # dates, decreasing by month
        Calc.set_date(sheet=sheet, cell_name="A8", day=28, month=2, year=2015)
        Calc.set_date(sheet=sheet, cell_name="B8", day=28, month=1, year=2015)

        # descending integers: 6, 4
        Calc.set_val(sheet=sheet, cell_name="A9", value=6)
        Calc.set_val(sheet=sheet, cell_name="B9", value=4)

        # get cell range series
        series = Calc.get_cell_series(sheet=sheet, range_name="A7:G9")

        # use first 2 cells for series, and fill to the right
        series.fillAuto(FillDirection.TO_RIGHT, 2)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The supplied cell range (``A7:G9``) includes the seed values, and the cells to be filled.

It's converted into an XCellSeries_ interface by Calc.getCellSeries(), which is defined as:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @staticmethod
        def get_cell_series(sheet: XSpreadsheet, range_name: str) -> XCellSeries:
            cell_range = sheet.getCellRangeByName(range_name)
            series = Lo.qi(XCellSeries, cell_range, True)
            return series

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``XCellSeries.fillAuto()`` can be supplied with four possible fill directions (``TO_BOTTOM``, ``TO_RIGHT``, ``TO_TOP``, and ``TO_LEFT``) which also dictate which cells are examined for seeds.
By setting the direction to be ``TO_RIGHT``, seed cells in the left-hand parts of the rows are examined.
The numerical (:t_red:`2`) in the call to ``fillAuto()`` shown above specifies how many of those cells will be considered in order to automatically determine the series used for the generated cell values.

:numref:`ch24fig_x_cell_series_fill` shows the result of filling rows ``7``, ``8``, and ``9``.

..
    figure 5

.. cssclass:: screen_shot invert

    .. _ch24fig_x_cell_series_fill:
    .. figure:: https://user-images.githubusercontent.com/4193389/204059144-a8fe7af2-9c86-4987-8fba-e7ec6f5c21f1.png
        :alt: Row Filling Using X Cell Series fill Auto method
        :figclass: align-center

        :Row Filling Using ``XCellSeries.fillAuto()``.

If ``XCellSeries.fillAuto()`` doesn't guess the correct series for the data generation, then ``XCellSeries.fillSeries()`` offers finer control over the process.
It supports five modes: ``SIMPLE``, ``LINEAR`` , ``GROWTH``, ``DATE``, and ``AUTO``.

``SIMPLE`` switches off the series generator, and the seed data is copied unchanged to the other blank cells.
``AUTO`` makes Office generate its data series automatically, so performs in the same way as fillAuto().
``LINEAR`` , ``GROWTH``, and ``DATE`` give more control to the programmer.

24.2.1 Using the LINEAR Mode
----------------------------

Rows ``2`` and ``3`` of the spreadsheet contain the numbers ``1`` and ``4`` (see :numref:`ch24fig_filler_py_sheet_default`).
By using the ``LINEAR`` mode, a step, and a stopping value, it's possible to specify an arithmetic series.
For example:

.. tabs::

    .. code-tab:: python

        # in Filler._fill_series() of filler.py
        # ...
        Calc.set_val(sheet=sheet, cell_name="A2", value=1)
        Calc.set_val(sheet=sheet, cell_name="A3", value=4)

        # Fill 2 rows; the 2nd row is not filled completely since
        # the end value is reached
        series = Calc.get_cell_series(sheet=sheet, range_name="A2:E3")
        series.fillSeries(FillDirection.TO_RIGHT, FillMode.LINEAR, Calc.NO_DATE, 2, 9)
                        # ignore date mode; step == 2; end at 9

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The :py:attr:`.Calc.NO_DATE` argument means that dates are not being generated. The ``2`` value is the step, and ``9`` is the maximum.
The resulting rows ``2`` and ``3`` are shown in :numref:`ch24fig_data_end_linear`.

..
    figure 6

.. cssclass:: screen_shot invert

    .. _ch24fig_data_end_linear:
    .. figure:: https://user-images.githubusercontent.com/4193389/204059586-c228dcc2-8217-4bd7-b527-507456990d2b.png
        :alt: Data Generation Using the LINEAR Mode
        :figclass: align-center

        :Data Generation Using the LINEAR Mode.

Note that the second row is incomplete since the generated values for those cells (``10`` and ``12``) exceeded the stopping value.

If no stopping value is required, then the last argument can be replaced with :py:attr:`.Calc.MAX_VALUE`.

24.2.2 Using the DATE Mode
--------------------------

If ``XCellSeries.fillSeries()`` is called using the ``DATE`` mode then it's possible to specify whether the day, weekday, month, or year parts of the seed date are changed by the series.
For example, the seed date at the start of row ``4`` (``20th Nov. 2015``) can be incremented one month at a time with the code:

.. tabs::

    .. code-tab:: python

        # in Filler._fill_series() of filler.py
        # ...
        Calc.set_date(sheet=sheet, cell_name="A4", day=20, month=11, year=2015)

        # fill by adding one month to date
        series = Calc.get_cell_series(sheet=sheet, range_name="A4:E4")
        series.fillSeries(FillDirection.TO_RIGHT, FillMode.DATE, FillDateMode.FILL_DATE_MONTH, 1, Calc.MAX_VALUE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The result is shown in :numref:`ch24fig_data_gen_date_mode`.

..
    figure 7

.. cssclass:: screen_shot invert

    .. _ch24fig_data_gen_date_mode:
    .. figure:: https://user-images.githubusercontent.com/4193389/204059924-eefbb861-24ea-4ea2-a9c8-71c034468952.png
        :alt: Data Generation Using the DATE Mode
        :figclass: align-center

        :Data Generation Using the DATE Mode.

When the month is incremented past ``12``, it resets to ``1``, and the year is incremented.

24.2.3 Using the GROWTH Mode
----------------------------

Whereas the ``LINEAR`` mode is for creating arithmetic series (:abbreviation:`i.e.` ones incrementing or decrementing in steps),
``GROWTH`` mode is for geometric progressions where the 'step' value is repeatedly multiplied to the seed.

In the following example, the seed in ``G6`` (:t_red:`10`; see :numref:`ch24fig_data_gen_date_mode`) is used in a geometric progression using multiples of ``2``.
The series is placed in cells going up the sheet starting from ``G6``.
The code:

.. tabs::

    .. code-tab:: python

        # in Filler._fill_series() of filler.py
        # ...
        Calc.set_val(sheet=sheet, cell_name="G6", value=10)

        # Fill from  bottom to top with a geometric series (*2)
        series = Calc.get_cell_series(sheet=sheet, range_name="G2:G6")
        series.fillSeries(FillDirection.TO_TOP, FillMode.GROWTH, Calc.NO_DATE, 2, Calc.MAX_VALUE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The resulting sheet is shown in :numref:`ch24fig_data_gen_growth_mode`.

..
    figure 8

.. cssclass:: screen_shot invert

    .. _ch24fig_data_gen_growth_mode:
    .. figure:: https://user-images.githubusercontent.com/4193389/204060246-c6b9b4e6-3e54-4c61-9171-874e24ecad34.png
        :alt:Data Generation Using the GROWTH Mode.
        :figclass: align-center

        :Data Generation Using the GROWTH Mode.

.. _ch24_cells_fancy_txt:

24.3 Cells with Fancy Text
==========================

The |c_txt_py|_ example brings together a few techniques for manipulating text in cells, namely the addition of borders, headlines, hyperlinks, and annotations.
The sheet ends up looking like :numref:`ch24fig_text_manipulaton_sht`.

..
    figure 9

.. cssclass:: screen_shot invert

    .. _ch24fig_text_manipulaton_sht:
    .. figure:: https://user-images.githubusercontent.com/4193389/204060401-9529598c-684d-407b-9db2-87bb22f6243d.png
        :alt: Text manipulation in a Sheet.
        :figclass: align-center

        :Text manipulation in a Sheet.

.. _ch24_creating_boder_headline:

24.3.1 Creating a Border and Headline
-------------------------------------

|c_txt_py|_ draws a decorative border and headline by calling:

.. tabs::

    .. code-tab:: python

        # in cell_texts.py
        Calc.highlight_range(
            sheet=sheet, range_name="A2:C7", headline="Cells and Cell Ranges"
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.highlight_range` adds a light blue border around the specified cell range (``A2:C7``), and the string argument is added to the top-left cell of the range.
It's intended to be a headline, so is drawn in dark blue, and the entire top row is made light blue to match the border.
The method is implemented as:

.. tabs::

    .. code-tab:: python

        # in Calc class (simplified)
        @classmethod
        def highlight_range(cls, sheet: XSpreadsheet, headline: str, range_name: str) -> XCell:
            cls.add_border(sheet=sheet, range_name=range_name, color=CommonColor.LIGHT_BLUE)
            addr = cls..get_address(sheet=sheet, range_name=range_name)
            header_range = Calc.get_cell_range(
                sheet=sheet,
                col_start=addr.StartColumn,
                row_start=addr.StartRow,
                col_end=addr.EndColumn,
                row_end=addr.StartRow
                )
            Props.set(header_range, CellBackColor=CommonColor.LIGHT_BLUE)

            # add headline text to the first cell of the row
            first_cell = cls.get_cell(cell_range=headerRange, col=0, row=0)
            cls.set_val(value=headline, cell=first_cell)

            # make text dark blue and bold
            Props.set(first_cell, CharColor=CommonColor.DARK_BLUE, CharWeight=FontWeight.BOLD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The three-argument :py:meth:`~.Calc.add_border` method calls the four-argument version which was described back in :ref:`ch22_adding_borders`.
It passes it a bitwise composition of all the border constants:

The cell range for the top row is extracted from the larger range supplied to :py:meth:`.Calc.highlight_range`.
The easiest way of doing this is to get the address of the larger range as a CellRangeAddress_ object, and use its row and column positions.
The header cell range uses the same row index for its starting and finishing rows:

.. tabs::

    .. code-tab:: python

        # part of Calc.highlight_range() (simplified)
        addr = Calc.get_address(sheet=sheet, range_name=range_name)
        header_range = Calc.getCellRange(
            sheet=sheet,
            col_start=addr.StartColumn,
            row_start=addr.StartRow,
            col_end=addr.EndColumn,
            row_end=addr.StartRow
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`highlight_range`

Perhaps the most confusing part of :py:meth:`.Calc.highlight_range` is how the first cell of the header range is referenced:

.. tabs::

    .. code-tab:: python

        first_cell = cls.get_cell(cell_range=headerRange, col=0, row=0)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This is a somewhat different use of :py:meth:`~.Calc.get_cell` than previous examples, which have always found a cell within a sheet.
For instance:

.. tabs::

    .. code-tab:: python

        cell = Calc.getCell(sheet=sheet, col=0, row=0)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The definition for this version of :py:meth:`~.Calc.get_cell` is:

.. tabs::

    .. code-tab:: python

        # in Calc class (overload method, simpilified)
        @classmethod
        def get_cell(cls, cell_range: XCellRange, col: int, row: int) -> XCell:
            return cell_range.getCellByPosition(col, row)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

A position in a cell range (:abbreviation:`i.e.` a (column, row) coordinate) is defined relative to the cell range.
This means that the call: ``first_cell = cls.get_cell(cell_range=headerRange, col=0, row=0)`` is requesting the top-left cell in ``headerRange``.
Since the ``headerRange`` covers ``A2:C2``, (``0``, ``0``) means the ``A2`` cell.

.. _ch24_adding_hyperlink:

24.3.2 Adding Hyperlink Text
----------------------------

:numref:`ch24fig_text_manipulaton_sht` shows that the ``B4`` cell contains two paragraphs.
The second ends with a hyperlink, which means that if the user control-clicks on the "hypertext" text,
then the URL `<https://github.com/Amourspirit/python_ooo_dev_tools>`__ is opened in the default web browser.

The SheetCell_ service inherits the Cell service which allows a cell to be manipulated with the XCell_ or the XText_ interfaces (see :numref:`ch24fig_sheet_cell_srv_interfaces`).

..
    figure 10

.. cssclass:: diagram invert

    .. _ch24fig_sheet_cell_srv_interfaces:
    .. figure:: https://user-images.githubusercontent.com/4193389/204062785-25d5c46a-b122-4346-b0c4-59bcc5976254.png
        :alt: The Sheet Cell Services and Interfaces
        :figclass: align-center

        :The SheetCell_ Services and Interfaces.

Once the cell is converted into XText_, many of my Writer support methods can be utilized.
For example:

.. tabs::

    .. code-tab:: python

        # in cell_texts.py
        # ...
        # Insert two text paragraphs and a hyperlink into the cell
        xtext = Lo.qi(XText, xcell, True)
        cursor = xtext.createTextCursor()
        Write.append_para(cursor=cursor, text="Text in first line.")
        Write.append(cursor=cursor, text="And a ")
        Write.add_hyperlink(
            cursor=cursor,
            label="hyperlink",
            url_str="https://github.com/Amourspirit/python_ooo_dev_tools"
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

A text cursor is created for the cell, and used to add the two paragraphs and the hyperlink.

Cell formatting is done through its properties.
As :numref:`ch24fig_sheet_cell_srv_interfaces` shows, the SheetCell_ service inherits the CharacterProperties_ and ParagraphProperties_ classes,
which contain the properties related to cell text:

.. tabs::

    .. code-tab:: python

        # in cell_texts.py
        # ...
        # beautify the cell
        Props.set(
            xcell,
            CharColor=CommonColor.DARK_BLUE,  # from styles.CharacterProperties
            CharHeight=18.0,  # from styles.CharacterProperties
            ParaLeftMargin=500,  # from styles.ParagraphProperties
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

24.3.3 Printing the Cell's Text
-------------------------------

The cell's text is accessed via its XText_ interface:

.. tabs::

    .. code-tab:: python

        # in cell_texts.py
        def _print_cell_text(self, cell: XCell) -> None:
            txt = Lo.qi(XText, cell, True)
            print(f'Cell Text: "{txt.getString()}"')
            # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The call to ``XText.getString()`` returns all the text, which is printed as:

::

    Cell Text: "Text in first line.
    And a hypertext"

The text can also be examined by moving a text cursor through it:

.. tabs::

    .. code-tab:: python

        cursor = txt.createTextCursor()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

However, it was surprising to discover that this text cursor can not be converted into a sentence or paragraph cursor.
Both the following calls return ``None``:

.. tabs::

    .. code-tab:: python

        sent_cursor = Lo.qi(XSentenceCursor, cursor)
        para_cursor = Lo.qi(XParagraphCursor, cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch24_adding_annotation:

24.3.4 Adding an Annotation
---------------------------

Cells can be annotated, which causes a little yellow text box to appear near the cell, linked to the cell by an arrow (as in :numref:`ch24fig_text_manipulaton_sht`).
Creating a new annotation is a two-step process: the XSheetAnnotationsSupplier_ interface is used to access the collection of existing annotations,
and a new one is added by supplying the annotation text and the address of the cell where its arrow will point. These steps are performed by the first half of :py:meth:`.Calc.add_annotation`:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @classmethod
        def add_annotation(cls, sheet: XSpreadsheet, cell_name: str, msg: str, is_visible=True) -> XSheetAnnotation:
            # add the annotation
            addr = cls.get_cell_address(sheet=sheet, cell_name=cell_name)
            anns_supp = Lo.qi(XSheetAnnotationsSupplier, sheet, True)
            anns = anns_supp.getAnnotations()
            anns.insertNew(addr, msg)

            # get a reference to the annotation
            xcell = cls.get_cell(sheet=sheet, cell_name=cell_name)
            ann_anchor = Lo.qi(XSheetAnnotationAnchor, xcell, True)
            ann = ann_anchor.getAnnotation()
            ann.setIsVisible(is_visible)
            return ann

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Annotation creation doesn't return a reference to the new annotation object.
For that it's necessary to examine the cell pointed to by the annotation.
XCell_ is converted into a XSheetAnnotationAnchor_, which has a ``getAnnotation()`` method for returning the annotation (if one exists).

XSheetAnnotation_ has several methods for obtaining information about the position, author, and modification date of the annotation.
``setIsVisible()`` allows its visibility to be switched on and off.

Work in progress ...

.. |c_txt| replace:: Cell Texts
.. _c_txt: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_cell_texts

.. |c_txt_py| replace:: cell_texts.py
.. _c_txt_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_cell_texts/cell_texts.py

.. |filler| replace:: Fuller
.. _filler: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_filler

.. |filler_py| replace:: filler.py
.. _filler_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_filler/filler.py

.. |d_sort| replace:: Data Sort
.. _d_sort: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_data_sort

.. |d_sort_py| replace:: data_sort.py
.. _d_sort_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_data_sort/data_sort.py

.. _CellRangeAddress: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1table_1_1CellRangeAddress.html
.. _CharacterProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
.. _ParagraphProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties.html
.. _PropertyValue: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1beans_1_1PropertyValue.html
.. _SheetCell: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SheetCell.html
.. _SheetCellRange: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SheetCellRange.html
.. _SortDescriptor2: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1SortDescriptor2.html
.. _TableSortDescriptor2: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableSortDescriptor2.html
.. _TableSortField: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1table_1_1TableSortField.html
.. _TextSortDescriptor2: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextSortDescriptor2.html
.. _XCell: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCell.html
.. _XCellRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCellRange.html
.. _XCellSeries: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XCellSeries.html
.. _XSheetAnnotation: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSheetAnnotation.html
.. _XSheetAnnotationAnchor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSheetAnnotationAnchor.html
.. _XSheetAnnotationsSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSheetAnnotationsSupplier.html
.. _XSortable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XSortable.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html
