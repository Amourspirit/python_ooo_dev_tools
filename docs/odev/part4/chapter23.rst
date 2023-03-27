.. _ch23:

**************************
Chapter 23. Garlic Secrets
**************************

.. topic:: Overview

    Freezing Rows; GeneralFunctions; Sheet Searching; Sheet Range Queries; Hidden Cells; Cell Merging; Splitting Windows; View Panes; View State Data; Active Panes; Inserting Rows and Columns; Shifting Cells

    Examples: |g_secrets|_

This chapter's |g_secrets_py|_ example illustrates how data can be extracted from an existing spreadsheet (``produceSales.xlsx``) using ``general`` functions, sheet searching, and sheet range queries.
It also has more examples of cell styling, and demonstrates sheet freezing, view pane splitting, pane activation, and the insertion of new rows into a sheet.

The idea for this chapter, and the data, comes from the Excel example in "Automate the Boring Stuff with Python" by Al :spelling:word:`Sweigart`, `chapter 13 <https://automatetheboringstuff.com/2e/chapter13/#calibre_link-437>`__. However, he utilized the Python library OpenPyXL to manipulate his file.

The beginning of the spreadsheet is shown in :numref:`ch23fig_part_product_sheet`.

..
    figure 1

.. cssclass:: screen_shot invert

    .. _ch23fig_part_product_sheet:
    .. figure:: https://user-images.githubusercontent.com/4193389/203836683-8033e670-791d-48e2-b3f6-ec61d2476154.png
        :alt: Part of the produce Sales Spreadsheet
        :figclass: align-center

        :Part of the ``produceSales.xlsx`` Spreadsheet.


Each row holds information about the sale of fruit or vegetables.
The columns are the type of produce sold ``column A``, the cost per pound of that produce ``B``, the number of pounds sold ``C``, and the total revenue from the sale ``D``.
The TOTAL column is calculated using a formula which multiplies the cost per pound by the number of pounds sold, and rounds the result to the nearest cent.
For example, cell ``D2`` contains ``=ROUND(B2*C2, 2)``.

Most of the ``main()`` function for |g_secrets_py|_ is shown below.
I'll explain the commented out parts in later sections:

.. tabs::

    .. code-tab:: python

        # garlic_secrets.py
        def main(self) -> None:
            loader = Lo.load_office(Lo.ConnectSocket())

            doc = Calc.open_doc(fnm=self._fnm, loader=loader)

            GUI.set_visible(is_visible=True, odoc=doc)

            sheet = Calc.get_sheet(doc=doc, index=0)
            Calc.goto_cell(cell_name="A1", doc=doc)

            # freeze one row of view
            # Calc.freeze_rows(doc=doc, num_rows=1)

            # find total for the "Total" column
            total_range = Calc.get_col_range(sheet=sheet, idx=3)
            total = Calc.compute_function(fn=GeneralFunction.SUM, cell_range=total_range)
            print(f"Total before change: {total:.2f}")
            print()

            self._increase_garlic_cost(doc=doc, sheet=sheet)  # takes several seconds

            # recalculate total
            total = Calc.compute_function(fn=GeneralFunction.SUM, cell_range=total_range)
            print(f"Total after change: {total:.2f}")
            print()

            # add a label at the bottom of the data, and hide it

            # split window into 2 view panes

            # access panes; make top pane show first row

            # display view properties

            # show view data

            # show sheet states

            # make top pane the active one in the first sheet

            # show revised sheet states

            # add a new first row, and label that as at the bottom

            # Save doc

            # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch23_freezing_rows:

23.1 Freezing Rows
==================

:py:meth:`.Calc.freeze_rows` specifies the number of rows that should be ``frozen`` on-screen as Office's view of the spreadsheet changes (:abbreviation:`i.e.` when the user scrolls downwards).
The function's argument is the number of rows to freeze, not a row index, and the choice of which rows are frozen depends on which row is currently selected (active) in the application window when the function is called.

The earlier call to :abbreviation:`i.e.` :py:meth:`.Calc.goto_cell` in ``main()`` means that ``A1`` is the active cell in the spreadsheet,
and so row ``1`` is the active row (and ``A`` the active column).
For that reason, ``Calc.freeze_rows(doc=doc, num_rows=1)`` will freeze rows ``1``, ``2``, and ``3`` so they remain in view when the spreadsheet is scrolled up or down.

:py:meth:`.Calc.freeze_rows` and :py:meth:`.Calc.freeze_cols` are implemented using :py:meth:`.Calc.freeze`:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @classmethod
        def freeze_rows(cls, doc: XSpreadsheetDocument, num_rows: int) -> None:
            cls.freeze(doc=doc, num_cols=0, num_rows=num_rows)

        @classmethod
        def freeze_cols(cls, doc: XSpreadsheetDocument, num_cols: int) -> None:
            cls.freeze(doc=doc, num_cols=num_cols, num_rows=0)

        @classmethod
        def freeze(cls, doc: XSpreadsheetDocument, num_cols: int, num_rows: int) -> None:
            ctrl = cls.get_controller(doc)
            if ctrl is None:
                return
            if num_cols < 0 or num_rows < 0:
                return
            xfreeze = Lo.qi(XViewFreezable, ctrl)
            xfreeze.freezeAtPosition(num_cols, num_rows)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.freeze` accesses the SpreadsheetView_ service (see :numref:`ch23fig_spread_sheet_vivew_srv_interface`) via the document's controller, and utilizes its XViewFreezable_ interface to call ``freezeAtPosition()``.

..
    figure 2

.. cssclass:: diagram invert

    .. _ch23fig_spread_sheet_vivew_srv_interface:
    .. figure:: https://user-images.githubusercontent.com/4193389/203843659-f617e223-1146-4ca4-8373-e2b0dbbb76e5.png
        :alt: The SpreadsheetView Services and Interfaces.
        :figclass: align-center

        :The SpreadsheetView_ Services and Interfaces.

.. _ch23_gen_func:

23.2 General Functions
======================

Office has a small set of functions (called GeneralFunction_) which can be applied to cell ranges without the need for formula strings (:abbreviation:`i.e.` a string starting with ``=``).

The more important functions are shown in :numref:`ch23tbl_some_general_functions`.

..
    Table 1

.. _ch23tbl_some_general_functions:

.. table:: Some GeneralFunctions.
    :name: some_general_functions

    ======================= =========================================================
    GeneralFunction Name    Calculation Applied to the Cell Values                   
    ======================= =========================================================
     SUM                     Sum the numerical values.
     COUNT                   Count all the values, including the non-numerical ones.
     COUNTNUMS               Count only the numerical values.
     AVERAGE                 Average all the numerical values.
     MAX                     Find the maximum of all the numerical values.
     MIN                     Find the minimum of all the numerical values.
     PRODUCT                 Return the product of all the numerical values.
     STDEV                   Standard deviation is calculated based on a sample.
    ======================= =========================================================

``GeneralFunction.SUM`` is used in ``main()``, to sum the ``TOTALS`` column of the spreadsheet:

.. tabs::

    .. code-tab:: python

        # in garlic_secrets.py
        total_range = Calc.get_col_range(sheet=sheet, idx=3)
        total = Calc.compute_function(fn=GeneralFunction.SUM, cell_range=total_range)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Cal.get_col_range` utilizes the XColumnRowRange_ interface to access the sheet as a series of columns.
The required column is extracted from the series via its index position:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @staticmethod
        def get_col_range(sheet: XSpreadsheet, idx: int) -> XCellRange:
            cr_range = Lo.qi(XColumnRowRange, sheet)
            if cr_range is None:
                raise MissingInterfaceError(XColumnRowRange)
            cols = cr_range.getColumns()
            con = Lo.qi(XIndexAccess, cols)
            if con is None:
                raise MissingInterfaceError(XIndexAccess)
            cell_range = Lo.qi(XCellRange, con.getByIndex(idx))
            if cell_range is None:
                raise MissingInterfaceError(
                    XCellRange, f"Could not access range for column position: {idx}"
                )
            return cell_range

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The sheet can also be treated as a series of rows by calling ``XColumnRowRange.getRows()``, as in :py:meth:`.Calc.get_row_range`:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @staticmethod
        def get_row_range(sheet: XSpreadsheet, idx: int) -> XCellRange:
            cr_range = Lo.qi(XColumnRowRange, sheet)
            if cr_range is None:
                raise MissingInterfaceError(XColumnRowRange)
            rows = cr_range.getRows()
            con = con = Lo.qi(XIndexAccess, rows)
            if con is None:
                raise MissingInterfaceError(XIndexAccess)
            cell_range = Lo.qi(XCellRange, con.getByIndex(idx))
            if cell_range is None:
                raise MissingInterfaceError(XCellRange, f"Could not access range for row position: {idx}")
            return cell_range

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The column returned by ``Calc.get_col_range(sheet=sheet, idx=3)`` includes the cell containing the word ``TOTALS``, but ``GeneralFunction.SUM`` only sums cells holding numerical data.

:py:meth:`.Calc.compute_function` obtains the XSheetOperation_ interface for the cell range, and calls ``XSheetOperation.computeFunction()`` to apply a GeneralFunction_:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @classmethod
        def compute_function(cls, fn: GeneralFunction | str, cell_range: XCellRange) -> float:
            try:
                sheet_op = Lo.qi(XSheetOperation, cell_range, raise_err=True)
                func = GeneralFunction(fn)  # convert to enum value if str
                if not isinstance(fn, uno.Enum):
                    Lo.print("Arg fn is invalid, returning 0.0")
                    return 0.0
                return sheet_op.computeFunction(func)
            except Exception as e:
                Lo.print("Compute function failed. Returning 0.0")
                Lo.print(f"    {e}")
            return 0.0

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch23_increase_garlic_cost:

23.3 Searching for the Cost of Garlic, and Increasing it
========================================================

|g_secrets_py|_ increases the ``Code per Pound`` value for every garlic entry.
The source document uses $1.19 (:abbreviation:`i.e.` see row 6 of :numref:`ch23fig_part_product_sheet`).
Due to a worldwide garlic shortage, this must be increased by 5% to $1.2495.

``_increase_garlic_cost()`` in |g_secrets_py|_ scans every used row in the sheet, examining the ``Produce`` cell to see if it contains the string ``Garlic``.
When the string is found, the corresponding ``Cost per Pound`` entry on that row is changed.
When the scanning reaches an empty cell, the end of the data has been reached, and the function returns.

.. tabs::

    .. code-tab:: python

        # in garlic_secrets.py
        def _increase_garlic_cost(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> int:
            row = 0
            prod_cell = Calc.get_cell(sheet=sheet, col=0, row=row)  # produce column
            # iterate down produce column until an empty cell is reached
            while prod_cell.getType() != CellContentType.EMPTY:
                if prod_cell.getFormula() == "Garlic":
                    # show the cell in-screen
                    Calc.goto_cell(doc=doc, cell_name=Calc.get_cell_str(col=0, row=row))
                    # change cost/pound column
                    cost_cell = Calc.get_cell(sheet=sheet, col=1, row=row)
                    cost_cell.setValue(1.05 * cost_cell.getValue())
                    Props.set(cost_cell, CharWeight=FontWeight.BOLD, CharColor=CommonColor.RED)
                row += 1
                prod_cell = Calc.get_cell(sheet=sheet, col=0, row=row)
            return row

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

To help the user see that changes have been made to the sheet, the text of each updated ``Cost per Pound`` cell is made **bold** and :red:`red`.
The cell properties being altered come from the CharacterProperties class.

The progression of the function is also highlighted by calling :py:meth:`.Calc.goto_cell` inside the loop.
This causes the spreadsheet to scroll down, to follow the changes.

Back in ``main()`` after the updates, the ``Totals`` column is summed once again, and the new value reported:

::

    Total before change: 231353.27
    Total after change: 231488.35

.. _ch23_hidden_msg:

23.4 Adding a Secret, Hidden Message
====================================

The change made by ``_increase_garlic_cost()`` are of a top-secret nature, and so the code adds an invisible message to the end of the sheet:

.. tabs::

    .. code-tab:: python

        # in GarlicSecrets.main() of garlic_secrets.py
        # ...
        empty_row_num = self._find_empty_row(sheet=sheet)
        self._add_garlic_label(doc=doc, sheet=sheet, empty_row_num=empty_row_num)
        Lo.delay(2_000)  # wait a bit before hiding last row

        row_range = Calc.get_row_range(sheet=sheet, idx=empty_row_num)
        Props.set(row_range, IsVisible=False)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.find_empty_row` returns the index of the first empty row in the sheet, which happens to be the first row after the end of the data.
It passes the index to ``_add_garlic_label()`` which inserts the large red text ``Top Secret Garlic Changes`` into the first cell on the row.
The message is so big that several cells are merged together to make enough space; the row's height is also increased.
The result is shown in :numref:`ch32fig_msg_end_sheet`.

..
    figure 3

.. cssclass:: screen_shot

    .. _ch32fig_msg_end_sheet:
    .. figure:: https://user-images.githubusercontent.com/4193389/203852280-ab987804-cda9-4566-8d54-182b8c3aff4a.png
        :alt: The Message at the end of the Sheet
        :figclass: align-center

        :The Message at the end of the Sheet.


This message is visible for about ``2`` seconds before it's hidden by setting the height of the row to ``0``.

This results in :numref:`ch32fig_hidden_msg_end_sheet`.

..
    figure 4

.. cssclass:: screen_shot

    .. _ch32fig_hidden_msg_end_sheet:
    .. figure:: https://user-images.githubusercontent.com/4193389/203852523-0615a4e3-39db-4551-85ee-58c6ed444f23.png
        :alt: The Hidden Message at the end of the Sheet
        :figclass: align-center

        :The Hidden Message at the end of the Sheet.

``_find_empty_row()`` utilizes a sheet ranges query to find all the empty cell ranges in the first column (``XCellRangesQuery.queryEmptyCells()``).
Then it extracts the smallest row index from those ranges:

.. tabs::

    .. code-tab:: python

        # in garlic_secrets.py
        def _find_empty_row(self, sheet: XSpreadsheet) -> int:
            # create a ranges query for the first column of the sheet
            cell_range = Calc.get_col_range(sheet=sheet, idx=0)
            Calc.print_address(cell_range=cell_range)
            cr_query = Lo.qi(XCellRangesQuery, cell_range)
            sc_ranges = cr_query.queryEmptyCells()
            addrs = sc_ranges.getRangeAddresses()
            Calc.print_addresses(*addrs)

            # find smallest row index
            row = -1
            if addrs is not None and len(addrs) > 0:
                row = addrs[0].StartRow
                for addr in addrs:
                    if row < addr.StartRow:
                        row = addr.StartRow
                print(f"First empty row is at position: {row}")
            else:
                print("Could not find an empty row")
            return row

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The XCellRangesQuery_ interface needs a cell range to search, which is obtained by calling :py:meth:`.Calc.get_col_range` to get the first column.
The cell range is printed by :py:meth:`.Calc.print_address`:

::

    Range: Sheet1.A5001:A1048576

There's only one empty cell range in the column, starting at row position ``5001`` and extending to the bottom of the sheet.
This is correct because the produce data is made up of ``5000`` records.

``_find_empty_row()`` returns the smallest start row (:abbreviation:`i.e.` ``5001``).

.. _ch23_adding_lbl:

23.5 Adding the Label
=====================

``_add_garlic_label()`` adds the large text string ``Top Secret Garlic Changes`` to the first cell in the supplied row.
The cell is made wider by merging a few cells together, made taller by adjusting the row height, and turned bright :red:`red`.

.. tabs::

    .. code-tab:: python

        # in garlic_secrets.py
        def _add_garlic_label(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet, empty_row_num: int) -> None:

            Calc.goto_cell(cell_name=Calc.get_cell_str(col=0, row=empty_row_num), doc=doc)

            # Merge first few cells of the last row
            rng_obj = Calc.get_range_obj(
                col_start=0, row_start=empty_row_num, col_end=3, row_end=empty_row_num
            )

            # merge and center range
            Calc.merge_cells(sheet=sheet, range_obj=rng_obj, center=True)

            # make the row taller
            Calc.set_row_height(sheet=sheet, height=18, idx=empty_row_num)
            # get the cell from the range cell start
            cell = Calc.get_cell(sheet=sheet, cell_obj=rng_obj.cell_start)
            cell.setFormula("Top Secret Garlic Changes")
            Props.set(
                cell,
                CharWeight=FontWeight.BOLD,
                CharHeight=24,
                CellBackColor=CommonColor.RED
                )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Cell merging requires a cell range, which is obtained by calling the version of :py:meth:`.Calc.get_range_obj` that employs start and end cell positions in (column, row) order.

The range spans the first four cells of the empty row, making it wide enough for the large text.

The XMergeable_ interface is supported by the SheetCellRange_ service and uses ``merge()`` with a boolean argument to ``merge`` or ``unmerge`` a cell range.

:py:meth:`.Calc.merge_cells` makes use of XMergeable_ and SheetCellRange_ to merge and center the range into a single cell range.

Changing the cell height affects the entire row, not just the merged cells, and so :py:meth:`.Calc.set_row_height` manipulates a cell range representing the row:

.. tabs::

    .. code-tab:: python

        # in Calc class (simplified)
        @classmethod
        def set_row_height(
            cls, sheet: XSpreadsheet, height: int, idx: int
            ) -> XCellRange:

            if height <= 0:
                Lo.print("Height must be greater then 0")
                return None
            cell_range = cls.get_row_range(sheet=sheet, idx=idx)
            # Info.show_services(obj_name="Cell range for a row", obj=cell_range)
            Props.set(cell_range, Height=(height * 100))
            return cell_range

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`set_row_height`

:py:meth:`~.Calc.set_row_height` illustrates the difficulties of finding property documentation.
The first line obtains an XCellRange_ interface for the row, and the second line changes a property in the cell range's service.
Pretend for a moment, that we don't know the name of this height property (``Height``). How could we find it?

That depends on finding the cell range's service.
First turn to the online documentation for the XCellRange_ class, which includes the class diagram shown in :numref:`ch23fig_xcellrange_children`.

..
    figure 5

.. cssclass:: diagram invert

    .. _ch23fig_xcellrange_children:
    .. figure:: https://user-images.githubusercontent.com/4193389/203855085-f450a3b2-3741-4929-8d2d-6ffc0de3cc4d.png
        :alt: Classes that Inherit XCellRange.
        :figclass: align-center

        :Classes that Inherit XCellRange_

This diagram combines the service and interface inheritance hierarchies.
Ignoring the interfaces that inherit XCellRange_ is easy because their names always begin with ``X``.
The remaining names mean that XCellRange_ is implemented by almost every service in the Calc API: ``CellRange`` (twice), ``TableColumn``, ``TableRow``, and ``TextTable``.
``CellRange`` appears twice because the blue triangle in the bottom-right corner of the first ``CellRange`` box means that there's more subclass hierarchy that's not shown;
in this case, ``SheetCellRange``, ``SheetCellCursor``, and ``Spreadsheet``.
The ``height`` property must be in one of these services, or one of their super-classes.

The correct choice is TableRow_ because the cell range is representing a spreadsheet row.
TableRow_ contains four properties, one of which is ``Height``.

Another approach for finding the service is to call :py:meth:`.Info.show_services`.
For example, by adding the following line to :py:meth:`.Calc.set_row_height`:

.. tabs::

    .. code-tab:: python

        Info.show_services("Cell range for a row", cell_range)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The following is printed in console:

::

    Cell range for a row Supported Services (1)
      "com.sun.star.table.TableRow"

Back in ``_add_garlic_label()``, text is stored in the cell, and its properties set.
Although XMergeable_ changes a cell range into a cell, it doesn't return a reference to that cell.
It can be accessed by calling :py:meth:`.Calc.get_cell`:

.. tabs::

    .. code-tab:: python

        Calc.get_cell(sheet=sheet, col=0, row=empty_row_num)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The various cell properties changed in ``add_garlic_label()`` are inherited from different classes shown in :numref:`ch23fig_sheetcell_serv_interface`.

..
    figure 6

.. cssclass:: diagram invert

    .. _ch23fig_sheetcell_serv_interface:
    .. figure:: https://user-images.githubusercontent.com/4193389/203856109-669f529b-f081-4ca8-8e6c-d7ac65240a02.png
        :alt: The Sheet Cell Services and Interfaces.
        :figclass: align-center

        :The SheetCell_ Services and Interfaces.

``CharWeight`` and ``CharHeight`` come from CharacterProperties_, and ``CellBackColor``, ``HoriJustify``, and ``VertJustify`` from CellProperties_.

Back in`` main()``, the newly created label is hidden after an interval of ``2`` seconds:

.. tabs::

    .. code-tab:: python

        # in GarlicSecrets.main() of garlic_secrets.py
        Lo.delay(2_000)  # wait a bit before hiding last row

        row_range = Calc.get_row_range(sheet=sheet, idx=empty_row_num)
        Props.set(row_range, IsVisible=False)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Row invisibility requires a property change to the row.
The row's cell range is obtained by calling :py:meth:`.Calc.get_row_range`, and then the ``IsVisible`` property is switched off.
Finding the name of this property involves the same investigative skills as the search for ``Height`` in :py:meth:`.Calc.set_row_height`.
As with ``Height``, ``IsVisible`` is defined in the TableRow_ service.

.. _ch23_splitting_panes:

23.6 Splitting a Window into Two View Panes
===========================================

The produce sales data is quite lengthy, so it's useful to split the window into two view panes: one showing the modified rows at the end, and the other the first few rows at the top of the sheet.
The first attempt at splitting the sheet is shown in :numref:`ch23fig_two_views`.
The division occurs just above row ``4999``, drawn as a thick dark gray line.

..
    figure 7

.. cssclass:: screen_shot invert

    .. _ch23fig_two_views:
    .. figure:: https://user-images.githubusercontent.com/4193389/203882408-41955c25-f03a-43a4-aee5-b2bab7bf31aa.png
        :alt: Two Views of the Sheet.
        :figclass: align-center

        :Two Views of the Sheet.

The code in ``main()`` of |g_secrets_py|_ for this:

.. tabs::

    .. code-tab:: python

        # in garlic_secrets.py
        # ...
        # split window into 2 view panes
        cell_name = Calc.get_cell_str(col=0, row=empty_row_num - 2)
        print(f"Splitting at: {cell_name}")
        # doesn't work with Calc.freeze()
        Calc.split_window(doc=doc, cell_name=cell_name)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.split_window` can utilize the SpreadsheetView_ service (see :numref:`ch23fig_spread_sheet_vivew_srv_interface`), and its XViewSplitable_ interface:

.. tabs::

    .. code-tab:: python

        controller = Calc.get_controller(doc)
        viewSplit = Lo.qi(XViewSplitable, controller);

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Unfortunately, XViewSplitable_ only offers a ``splitAtPosition()`` method which specifies the split location in terms of pixels.
In addition, the interface is deprecated.

A better alternative is to employ the ``splitWindow`` dispatch command, which has a ``ToPoint`` property argument for a cell name (:abbreviation:`i.e.` ``A4999``) where the split will occur.
Therefore, :py:meth:`.Calc.split_window` is coded as:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @classmethod
        def split_window(cls, doc: XSpreadsheetDocument, cell_name: str) -> None:
            frame = cls.get_controller(doc).getFrame()
            cls.goto_cell(cell_name=cell_name, frame=frame)
            props = Props.make_props(ToPoint=cell_name)
            Lo.dispatch_cmd(cmd="SplitWindow", props=props, frame=frame)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The call to :py:meth:`.Calc.goto_cell` changes the on-screen active cell.
If it's left out then the ``SplitWindow`` dispatch creates a split at the currently selected cell rather than the one stored in the ``ToPoint`` property.
In other words, it appears that the ``SplitWindow`` dispatch ignores the property.

:numref:`ch23fig_two_views` shows another problem with the split - the top pane still shows the rows of data immediately above the split line.
The preference is for the top view to show the first rows at the start of the sheet.

One way of changing the displayed rows is via the view's XViewPane_ interface (see :numref:`ch23fig_spread_sheet_vivew_srv_interface`).
Each view (or pane) created by a split is represented by an XViewPane_ object, and a collection of all the current panes can be accessed through the SpreadsheetView_ service.
This approach is implemented in :py:meth:`.Calc.get_view_panes`, which returns the collection as an array:

.. tabs::

    .. code-tab:: python

        # in Calc class (simplified)
        @classmethod
        def get_view_panes(cls, doc: XSpreadsheetDocument) -> List[XViewPane] | None:
            con = Lo.qi(XIndexAccess, cls.get_controller(doc))
            if con is None:
                raise MissingInterfaceError(XIndexAccess, "Could not access the view pane container")

            panes = []
            for i in range(con.getCount()):
                try:
                    panes.append(Lo.qi(XViewPane, con.getByIndex(i)))
                except UnoException:
                    Lo.print(f"Could not get view pane {i}")
            if len(panes) == 0:
                Lo.print("No view panes found")
                return None
            return panes

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.get_view_panes` is called like so:

.. tabs::

    .. code-tab:: python

        panes = Calc.get_view_panes(doc)
        print(f'No of panes: {len(panes)}')

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The XViewPane_ interface has methods for setting and getting the visible row and column in the view.
For example, the first pane can be made to show the first row, by calling:

.. tabs::

    .. code-tab:: python

        panes[0].setFirstVisibleRow(0)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch23_view_states_top_pane:

23.7 View States, and Making the Top Pane Active
================================================

The previous section split the window into two panes, and changed the view in the top pane to show the first rows of the sheet. But there's still a problem which
can be seen in :numref:`ch23fig_two_views` - the active cell is still in the bottom pane, and I want it to be in the first row of the top pane.
More coding is required.

Hidden away in the XController_ interface are the methods ``getViewData()`` and ``restoreViewData()``.
They allow a programmer to access and change the view details of all the sheets in the document.
For example, the following retrieval of the view data for a document:

.. tabs::

    .. code-tab:: python

        ctrl = Calc.get_controller(doc) # XController
        print(ctrl.getViewData())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Prints:

::

    100/60/0;0;tw:270;3/13/0/0/0/0/2/0/0/0/1;5/15/0/0/0/0/2/0/0/0/0;0/0/0
    /0/0/0/2/0/0/0/0

This can be better understood by separating the data according to the ``;``'s, producing:

::

    100/60/0
    0
    tw:270
    3/13/0/0/0/0/2/0/0/0/1
    5/15/0/0/0/0/2/0/0/0/0
    0/0/0/0/0/0/2/0/0/0/0

The first three lines refer to the document's zoom settings, the active sheet index, and the position of the scrollbar.
The fourth line and below give the view state information for each sheet.
In the example document, there are three sheets, so three view state lines.

Each view state consists of ``11`` values, separated by ``/``'s. Their meaning, based on their index positions:

.. cssclass:: ul-list

    - indices ``0`` and ``1`` contain the current cursor position in terms of column and row positions;
    - ``2``: this records if column split mode is being used (``0`` or ``1``);
    - ``3``: is row split mode being used? (``0`` or ``1``);
    - ``4``: the vertical split position (in pixels);
    - ``5``: the horizontal split position (in pixels);
    - ``6``: the active/focused pane number for this sheet;
    - ``7``: the left column index of the left-hand panes;
    - ``8``: the left column index of the right-hand panes;
    - ``9``: the top row index of the upper panes;
    - ``10``: the top row index of the lower panes.

A sheet can be split horizontal and/or vertically, which can generate a maximum of four panes, which are numbered as in :numref:`ch23fig_four_panes_window`.

..
    figure 8

.. cssclass:: screen_shot

    .. _ch23fig_four_panes_window:
    .. figure:: https://user-images.githubusercontent.com/4193389/203885930-aa162cc7-397c-4882-87b7-cd698bb0236c.png
        :alt: The Four Panes in a Split Window
        :figclass: align-center

        :The Four Panes in a Split Window.

If a window is split only horizontally, then numbers ``0`` and ``1`` are used. If the split is only vertical, then the numbers used are ``0`` and ``2``.

Only one pane can be active (:abbreviation:`i.e.` have keyboard focus) at a time.
For example, in :numref:`ch23fig_four_panes_window`, pane ``1`` is active.
The active pane number is stored in view state index ``6``.

The view state information at index positions ``7`` to ``10`` define the top-left corners of each pane.
For example, since pane ``1`` is in the top-right of the sheet, its top-left corner is obtained by combining the values in view state index positions ``8`` and ``9``.
Position ``8`` supplies the column index of the two right-hand panes, and position ``9`` the row index of the top two panes.

If a sheet is not split at all, then its top-left corner is reconstructed by accessing index positions ``7`` and ``10``.

Although it's possible for a programmer to extract all this information from the view data string by themselves,
|odev| implemented a support class called :py:class:`~.view_state.ViewState` which stores the data in a more easily accessible form.
:py:meth:`.Calc.get_view_states` parses the view data string, creating an array of ViewState objects, one object for each sheet in the document.
For example, the following code is in |g_secrets_py|_:

.. tabs::

    .. code-tab:: python

        # in garlic_secrets.py
        # ...
        # show sheet states
            states = Calc.get_view_states(doc=doc)
            for s in states:
                s.report()
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

When it's executed after the sheet has been split as shown in :numref:`ch23fig_two_views`, the following is printed:

::

    Sheet View State
      Cursor pos (column, row): (0, 4998) or 'A4999'
      Sheet is split horizontally at 259
      Number of focused pane: 2
      Left column indicies of left/right panes: 0 / 0
      Top row indicies of upper/lower panes: 0 / 4998

One view state is reported since the document only contains one sheet.
The output says that the sheet is split vertically, and the lower pane is active (in focus).

:py:meth:`.Calc.get_view_states` is implemented as:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @classmethod
        def get_view_states(cls, doc: XSpreadsheetDocument) -> List[mViewState.ViewState] | None:
            ctrl = cls.get_controller(doc)

            view_data = str(ctrl.getViewData())
            view_parts = view_data.split(";")
            p_len = len(view_parts)
            if p_len < 4:
                Lo.print("No sheet view states found in view data")
                return None
            states = []
            for i in range(3, p_len):
                states.append(mViewState.ViewState(view_parts[i]))
            return states

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The first three entries in the view data (:abbreviation:`i.e.` the document's zoom, active sheet, and scrollbar position) are discarded, so only the document's view states are stored.

Paired with :py:meth:`.Calc.get_view_states` is :py:meth:`.Calc.set_view_states` which uses an array of :py:class:`~.view_state.ViewState` objects to update the view states of a document.
It is coded as:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @classmethod
        def set_view_states(
            cls, doc: XSpreadsheetDocument, states: Sequence[mViewState.ViewState]
        ) -> None:
            ctrl = cls.get_controller(doc)
            if ctrl is None:
                return
            view_data = str(ctrl.getViewData())
            view_parts = view_data.split(";")
            p_len = len(view_parts)
            if p_len < 4:
                Lo.print("No sheet view states found in view data")
                return None

            vd_new = []
            for i in range(3):
                vd_new.append(view_parts[i])

            for state in states:
                vd_new.append(str(state))
            s_data = ";".join(vd_new)
            Lo.print(s_data)
            ctrl.restoreViewData(s_data)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

A new view data string is constructed, and loaded into the document by calling ``XController.restoreViewData()``.
The string is composed from view state strings obtained by calling :py:meth:`.ViewState.to_string` for each :py:class:`~.view_state.ViewState` object.
Also, the existing values for the document's zoom, active sheet, and scrollbar position are copied over unchanged by extracting their current values from a call to ``XController.getViewData()``.

Finally the active pane is able to be changed to be the top view.
Also move the view in that newly activated pane to the top of the sheet:

.. tabs::

    .. code-tab:: python

        # in garlic_secrets.py
        # ...
        states = Calc.get_view_states(doc=doc)

        # make top pane the active one in the first sheet
        states[0].move_pane_focus(dir=ViewState.PaneEnum.MOVE_UP)
        Calc.set_view_states(doc=doc, states=states)
        # move selection to top cell
        Calc.goto_cell(cell_name="A1", doc=doc)

         # show revised sheet states
        states = Calc.get_view_states(doc=doc)
        for s in states:
            s.report()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The view states are obtained by calling :py:meth:`.Calc.get_view_states`.
The ``states`` list will hold one :py:class:`~.view_state.ViewState` object for each sheet in the document, so by using ``states[0]`` the panes in the first sheet will be affected.
:py:meth:`.ViewState.move_pane_focus`, which is described shortly, changes the focus to the top pane.
Finally, the modified view states are written back to the document by :py:meth:`.Calc.set_view_states`.

:numref:`ch23fig_changed_pane` shows the outcome of this code: the active cell is now in the top pane, at cell ``A1``.

..
    figure 9

.. cssclass:: screen_shot invert

    .. _ch23fig_changed_pane:
    .. figure:: https://user-images.githubusercontent.com/4193389/203888425-32b2d539-caf8-46e8-96c8-e5d8965404ca.png
        :alt: A Changed Active Cell and Pane
        :figclass: align-center

        :A Changed Active Cell and Pane.

The code fragment above also prints out the revised view state, which is:

::

    Sheet View State
      Cursor pos (column, row): (0, 0) or 'A1'
      Sheet is split horizontally at 259
      Number of focused pane: 0
      Left column indicies of left/right panes: 0 / 0
      Top row indicies of upper/lower panes: 0 / 4998

:py:meth:`.ViewState.move_pane_focus` changes one value in the view state - the focused pane number (index no. ``6`` in the list given earlier).
|odev| does not to implement this by having the programmer supply a pane number (:abbreviation:`i.e.` ``0``, ``1``, ``2``, or ``3`` as shown in :numref:`ch23fig_four_panes_window`)
since these numbers may not all be used in a given split. Instead the focus change is specified in terms of a direction, as shown in the code:

.. tabs::

    .. code-tab:: python

        # in viewState class
        def move_pane_focus(self, dir: int | ViewState.PaneEnum) -> bool:
            try:
                d = ViewState.PaneEnum(dir)
            except Exception:
                raise ValueError("Unknown move direction")

            if d == ViewState.PaneEnum.MOVE_UP:
                if self._pane_focus_num == 3:
                    self._pane_focus_num = 1
                elif self._pane_focus_num == 2:
                    self._pane_focus_num = 0
                else:
                    Lo.print("cannot move up")
                    return False
            elif d == ViewState.PaneEnum.MOVE_DOWN:
                if self._pane_focus_num == 1:
                    self._pane_focus_num = 3
                elif self._pane_focus_num == 0:
                    self._pane_focus_num = 2
                else:
                    Lo.print("cannot move down")
                    return False
            elif d == ViewState.PaneEnum.MOVE_LEFT:
                if self._pane_focus_num == 1:
                    self._pane_focus_num = 0
                elif self._pane_focus_num == 3:
                    self._pane_focus_num = 2
                else:
                    Lo.print("cannot move left")
                    return False
            elif d == ViewState.PaneEnum.MOVE_RIGHT:
                if self._pane_focus_num == 0:
                    self._pane_focus_num = 1
                elif self._pane_focus_num == 2:
                    self._pane_focus_num = 3
                else:
                    Lo.print("cannot move right")
                    return False
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    :py:class:`~.view_state.ViewState.PaneEnum`

.. _ch23_adding_new_first:

23.8 Adding a New First Row and Shifting Cells
==============================================

The final task in |g_secrets_py|_ is to add the ``Top Secret Garlic Changes`` text to the sheet again, this time as a visible title for the spreadsheet.
The only new API feature used is the insertion of a row. This is done with:

.. tabs::

    .. code-tab:: python

        # in garlic_secrets.py
        # ...
        # add a new first row, and label that as at the bottom
        Calc.insert_row(sheet=sheet, idx=0)
        self._add_garlic_label(doc=doc, sheet=sheet, empty_row_num=0)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The ``_add_garlic_label()`` method is unchanged from earlier, but is now passed row index ``0`` rather than the last row.
The result is shown in :numref:`ch23fig_sheet_new_title_row`.

..
    figure 10

.. cssclass:: screen_shot

    .. _ch23fig_sheet_new_title_row:
    .. figure:: https://user-images.githubusercontent.com/4193389/203891472-5d7dffe9-1099-4d1d-b998-d35111bb7226.png
        :alt: The Sheet with a New Title Row
        :figclass: align-center

        :The Sheet with a New Title Row.

:py:meth:`.Calc.insert_row` manipulates a row as a cell range, so it's once again necessary to access the sheet's XColumnRowRange_ interface, to retrieve a TableRows_ object.
The XTableRows_ interface supports the adding and removal of rows at specified index positions.
This allows :py:meth:`.Calc.insert_row` to be coded as:

.. tabs::

    .. code-tab:: python

        # in Calc class (simplified)
        @staticmethod
        def insert_row(sheet: XSpreadsheet, idx: int) -> bool:
            cr_range = Lo.qi(XColumnRowRange, sheet, True)
            rows = cr_range.getRows()
            rows.insertByIndex(idx, 1)  # add 1 row at idx position
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`insert_row`

There's a similar :py:meth:`.Calc.insert_cols` method that utilizes the XTableColumns_ interface:

.. tabs::

    .. code-tab:: python

        # in Calc class (simplified)
        @staticmethod
        def insert_column(sheet: XSpreadsheet, idx: int) -> bool:
            cr_range = mLo.Lo.qi(XColumnRowRange, sheet, True)
            cols = cr_range.getColumns()
            cols.insertByIndex(idx, 1)  # add 1 column at idx position
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`insert_column`

The insertion of an arbitrary number of blank cells into a sheet is a bit more complicated because existing cells must be 'moved' out of the way, and this can be done by moving them downwards or to the right.
The shift-able cells are specified as a cell range, and the sheet's XCellRangeMovement_ interface moves them in a specific direction. XCellRangeMovement_ is supported by the Spreadsheet_ service.

The :py:meth:`.Calc.insert_cells` method implements this approach:

.. tabs::

    .. code-tab:: python

        # in Calc class (simplified)
        @classmethod
        def insert_cells(cls, sheet: XSpreadsheet, cell_range: XCellRange, is_shift_right: bool) -> bool:
            mover = mLo.Lo.qi(XCellRangeMovement, sheet, True)
            addr = cls.get_address(cell_range)
            if is_shift_right:
                mover.insertCells(addr, CellInsertMode.RIGHT)
            else:
                mover.insertCells(addr, CellInsertMode.DOWN)
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        - :odev_src_calc_meth:`insert_cells`

An example call:

.. tabs::

    .. code-tab:: python

        blanks = Calc.get_cell_range(sheet=sheet, range_name="A4999:B5001")
        Calc.insert_cells(sheet=sheet, cell_range=blanks, is_shift_right=True)  # shift right

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This shifts the last three rows of the produce sheet ``A4999:B5001`` to the right by two cells, producing :numref:`ch23fig_shifted_cells`.

..
    figure 11

.. cssclass:: screen_shot

    .. _ch23fig_shifted_cells:
    .. figure:: https://user-images.githubusercontent.com/4193389/203893256-eade70ad-dead-48a7-bb14-491f8056cbb5.png
        :alt: Shifted Cells at the end of the Produce Sheet
        :figclass: align-center

        :Shifted Cells at the end of the Produce Sheet.

.. |g_secrets| replace::  Garlic Secrets
.. _g_secrets: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_garlic_secrets

.. |g_secrets_py| replace:: garlic_secrets.py
.. _g_secrets_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_garlic_secrets/garlic_secrets.py

.. _CellProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1CellProperties.html
.. _CharacterProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
.. _GeneralFunction: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet.html#ad184d5bd9055f3b4fd57ce72c781758d
.. _SheetCell: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SheetCell.html
.. _SheetCellRange: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SheetCellRange.html
.. _Spreadsheet: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1Spreadsheet.html
.. _SpreadsheetView: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SpreadsheetView.html
.. _TableRow: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableRow.html
.. _TableRows: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableRows.html
.. _XCellRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCellRange.html
.. _XCellRangeMovement: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XCellRangeMovement.html
.. _XCellRangesQuery: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XCellRangesQuery.html
.. _XColumnRowRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XColumnRowRange.html
.. _XController: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XController.html
.. _XMergeable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XMergeable.html
.. _XSheetOperation: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSheetOperation.html
.. _XTableColumns: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XTableColumns.html
.. _XTableRows: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XTableRows.html
.. _XViewFreezable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XViewFreezable.html
.. _XViewPane: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XViewPane.html
.. _XViewSplitable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XViewSplitable.html
