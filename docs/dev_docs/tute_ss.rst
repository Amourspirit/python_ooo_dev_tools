.. _tute_ss:

LIBREOFFICE AND PYTHON TO WORK WITH EXCEL SPREADSHEETS 
******************************************************

.. cssclass:: screen_shot invert

    .. _tute_ss_fig_main_img:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299691-79a5b476-dd5c-42e4-927b-982c1213d43b.png
        :alt: Tutorial Image
        :figclass: align-center

.. warning:: 

    This document is a work in progress and many of the examples are not yet in full working order.

Although we don’t often think of spreadsheets as programming tools, almost everyone uses them to organize information into two-dimensional data structures, perform calculations with formulas, and produce output as charts. In this tutorial, we’ll integrate Python into the popular spreadsheet application LibreOffice.

Excel is a popular and powerful spreadsheet application for Windows. The |odev| package allows your Python programs to read and modify Excel spreadsheet files. For example, you might have the boring task of copying certain data from one spreadsheet and pasting it into another one. Or you might have to go through thousands of rows and pick out just a handful of them to make small edits based on some criteria. Or you might have to look through hundreds of spreadsheets of department budgets, searching for any that are in the red. These are exactly the sort of boring, mindless spreadsheet tasks that Python can do for you.

Although Excel is proprietary software from Microsoft, there are free alternatives that run on Windows, macOS, and Linux. Both LibreOffice Calc and OpenOffice Calc work with Excel’s .xlsx file format for spreadsheets, which means the |odev| module can work on spreadsheets from these applications as well. You can download the software from https://www.libreoffice.org/ and https://www.openoffice.org/, respectively. Even if you already have Excel installed on your computer, you may find these programs easier to use. The screenshots in this chapter, however, are all from Excel 2010 on Windows 10.

.. _tute_ss_excel_docs:

Excel Documents
---------------

Let’s go over some basic definitions: an Excel spreadsheet document is called a workbook. A single workbook is saved in a file with the .xlsx extension. Each workbook can contain multiple sheets (also called worksheets). The sheet the user is currently viewing (or last viewed before closing Excel) is called the active sheet.

Each sheet has columns (addressed by letters starting at A) and rows (addressed by numbers starting at 1). A box at a particular column and row is called a cell. Each cell can contain a number or text value. The grid of cells with data makes up a sheet.

.. _tute_ss_install_odev:

Installing |odev|
---------------

Python does not come with |odev|, so you’ll have to install it. Follow the instructions in the |odev| :ref:`dev_doc` for installing the :ref:`dev_doc_virtulal_env`.

.. _tute_ss_python_libreoffice:

Working with Python and LibreOffice
-----------------------------------

Note: Python is normally used by running script files, but it is an interpretive language executing line by line. This tutorial uses the REPL, an interactive python shell, so the user can execute a single Python command and get the result. New modules are normally loaded in the tutorial as they are needed but later they may appear at the top of the listing as normally used.

Note: Code in a section often required code earlier in the section to be executed beforehand but all the required code should be within the section. The names used for objects like the workbook and worksheet vary throughout the sections and a mix of plain arguments and keyword pairs are used.

Firstly, let us understand how python works with Office. An office instance is required before python can interact with the objects. When the python program is finished it is important to close any document and the Office instance or it will continue to run in the computer stopping other interfaces from starting it. This initialisation and finalisation code is required even if it is not shown in the examples.

Once |odev| is installed, start up a python shell and enter the following code into the REPL:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> loader = Lo.load_office(Lo.ConnectSocket(headless=True))
        >>> # use the Office API...
        >>> Lo.close_doc(wb)
        >>> # generates an error if wb not open
        >>> Lo.close_office()
        True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None




.. cssclass:: bg_light_gray, blue

   As a comparison, elsewhere this might be done in a script with similar code to the following to close the loader and context manager automatically after it runs, even if there is an error:

.. tabs::

    .. code-tab:: python

        def main() -> int:
            with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:
            doc = Calc.create_doc(loader=loader)
            sheet = Calc.get_sheet(doc=doc, index=0)
            # do some work
            Lo.close_doc(doc=doc)
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Note: Similar commands are used to open with GUI:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> from ooodev.utils.gui import GUI
        >>> _ = Lo.load_office(Lo.ConnectSocket())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. _tute_ss_reading_excel_docs:

Reading Excel Documents
=======================

The examples in this section will use a spreadsheet named ``example.xlsx`` stored in the root folder.
You can either create the spreadsheet yourself or download it from `<https://nostarch.com/automatestuff2/>`__.
:numref:`tute_ss_fig_office_timeline` shows the tabs for the three default sheets named ``Sheet1``, ``Sheet2``, and ``Sheet3`` that Excel automatically provides for new workbooks.
(The number of default sheets created may vary between operating systems and spreadsheet programs.)

.. cssclass:: diagram invert

    .. _tute_ss_fig_office_timeline:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299710-3cfbd875-1d13-43f2-8e62-d93af56fa5f1.png
        :alt: OpenOffice Timeline Image
        :figclass: align-center

        The tabs for a workbook’s sheets are in the lower-left corner of Excel

Sheet 1 in the example file should look like :numref:`tute_ss_tbl_sheet_data`
(If you didn’t download ``example.xlsx`` from the website, you should enter this data into the sheet yourself).

:numref:`tute_ss_tbl_sheet_data`: The ``example.xlsx`` Spreadsheet

.. _tute_ss_tbl_sheet_data:

.. table:: Sheet Data.
    :name: sheet_data

    +--+-----------------+--------------+----+
    |  | A               | B            | C  |
    +==+=================+==============+====+
    | 1|5/04/2015 13:34  |Apples        |  73|
    +--+-----------------+--------------+----+
    | 2|5/04/2015 3:41   |Cherries      |  85|
    +--+-----------------+--------------+----+
    | 3|6/04/2015 12:46  |Pears         |  14|
    +--+-----------------+--------------+----+
    | 4|8/04/2015 8:59   |Oranges       |  52|
    +--+-----------------+--------------+----+
    | 5|10/04/2015 2:07  |Apples        | 152|
    +--+-----------------+--------------+----+
    | 6|10/04/2015 18:10 |Bananas       |  23|
    +--+-----------------+--------------+----+
    | 7|10/04/2015 2:40  |Strawberries  |  98|
    +--+-----------------+--------------+----+

Now that we have our example spreadsheet, let’s see how we can manipulate it with the |odev| package.

.. _tute_ss_open_excel_doc_odev:

Opening Excel Documents with |odev|
---------------------------------

Once you’ve installed the |odev| package, you’ll be able to use the Calc class. Enter the following into a new interactive shell:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> loader = Lo.load_office(Lo.ConnectSocket(headless=True, soffice="C:\\Program Files\\LibreOfficeDev 7\\program\\soffice.exe"))
        >>>
        >>> from ooodev.office.calc import Calc
        >>> wb = Calc.open_doc('example.xlsx', loader)
        >>> type(wb)
        <class 'pyuno'>

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The Calc.open_doc() class takes in the filename and loader, and returns a value of the workbook data type.
This Workbook object represents the Excel file, a bit like how a File object represents an opened text file.

Remember that example.xlsx needs to be in the current working directory in order for you to work with it.
You can find out what the current working directory is by importing os and using os.getcwd(), and you can change the current working directory using os.chdir().

.. _tute_ss_get_sheet_wb:

Getting Sheets from the Workbook
--------------------------------

You can get a list of all the sheet names in the workbook by accessing the sheetnames property.
Enter the following into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> Calc.get_sheet_names(wb)
        ('Sheet1', 'Sheet2', 'Sheet3')
        >>> ws = Calc.get_sheet(doc=wb, sheet_name='Sheet3')
        >>> Calc.get_sheet_name(ws)
        'Sheet3'
        >>> ws2 = Calc.get_active_sheet(wb)
        >>> Calc.get_sheet_name(ws2)
        'Sheet1'

        >>> Lo.close_doc(wb)
        >>> Lo.close_office()
        True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Each sheet is represented by a Worksheet object and you can use the Calc class to return it's properties.
:py:meth:`~.Calc.get_sheet_names` will return all sheets in the workbook given as an argument.
A particular Worksheet object is returned using :py:meth:`~.Calc.get_sheet` with the Workbook and sheet name string as arguments, and :py:meth:`~.Calc/get_sheet_name` with a Worksheet object argument returns teh Worksheet name.
Finally, you can use :py:meth:`~.Calc.get_active_sheet` of a Workbook object to get the workbook’s active sheet, and from there the name.
The active sheet is the sheet that is displayed when the workbook is opened on your computer.

.. _tute_ss_get_sheet_cells:

Getting Cells from the Sheets
-----------------------------

Once you have a Worksheet object, you can access a Cell object using the Calc class. Enter the following into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> from ooodev.office.calc import Calc
        >>> from ooodev.utils.gui import GUI
        >>> from ooodev.utils.date_time_util import DateUtil
        >>>
        >>> _ = Lo.load_office(Lo.ConnectSocket(soffice="C:\\Program Files\\LibreOfficeDev 7\\program\\soffice.exe"))
        >>> wb = Calc.open_doc('example.xlsx')
        >>> GUI.set_visible(is_visible=True, doc=wb)
        >>>
        >>> ws = Calc.get_sheet(doc=wb, sheet_name='Sheet1')
        >>>
        >>> Calc.get_val(sheet=ws, cell_name="A1")
        42099.565300925926
        >>> DateUtil.date_from_number(Calc.get_val(sheet=ws, cell_name="A1"))
        datetime.datetime(2015, 4, 5, 13, 34, 2, tzinfo=datetime.timezone.utc)
        >>> str(DateUtil.date_from_number(Calc.get_val(sheet=ws, cell_name="A1")))
        '2015-04-05 13:34:02+00:00'
        >>>
        >>> Calc.get_val(sheet=ws, cell_name="B1")
        'Apples'
        >>>
        >>> c = Calc.get_cell(ws, "B1")
        >>> 'Row %s, Column %s is %s' % (Calc.get_cell_address(c).Row, Calc.get_cell_address(c).Column, Calc.get_val(c))
        'Row 0, Column 1 is Apples'
        >>>
        >>> Calc.get_val(sheet=ws, cell_name="C1")
        73.0

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None




The Cell object has a value property that contains, unsurprisingly, the value stored in that cell.
There are many ways of referencing Cell objects, using the cell object, or the sheet with: cell address, cell name also have row, column, and coordinate properties that provide location information for the cell.

|odev| returns dates as float so they need to be formatted to display the date in the required format.

Here, accessing the value property of our Cell object for cell ``B1`` gives us the string ``Apples``.
The row property gives us the integer ``1``, the column property gives us ``B``, and the coordinate property gives us ``B1``.

Specifying a column by letter can be tricky to program, especially because after column ``Z``, the columns start by using two letters: ``AA``, ``AB``, ``AC``, and so on.
As an alternative, you can also get a cell using :py:meth:`.Calc.get_cell` method and passing integers for its row and column keyword arguments.
The first row or column integer is ``0``, not ``1``.
Continue the interactive shell example by entering the following:

.. tabs::

    .. code-tab:: python

        >>> Calc.get_val(Calc.get_cell(ws, "B1"))
        'Apples'
        >>> Calc.get_val(Calc.get_cell(ws, 1,0))
        'Apples'
        >>> for i in range(0, 7, 2): # Go through every other row:
        ...     print(i+1, Calc.get_val(Calc.get_cell(ws, 1,i)))
        ...
        1 Apples
        3 Pears
        5 Apples
        7 Strawberries

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

As you can see, using :py:meth:`.Calc.get_cell` method and passing it ``column=1`` and ``row=0`` gets you a Cell object for cell ``B1``, just like specifying :py:meth:`~.Calc.get_cell` with 'B1' did.
Then, using the :py:meth:`~.Calc.get_val` method and its keyword arguments, you can write a for loop to print the values of a series of cells.

Say you want to go down column ``B`` and print the value in every cell with an odd row number.
By passing ``2`` for the ``range()`` function’s “step” parameter, you can get cells from every second row (in this case, all the odd-numbered rows).
The for loop’s ``i`` variable is passed for the row keyword argument to the ``cell()`` method, while ``2`` is always passed for the column keyword argument.
Note that the integer ``2``, not the string ``B``, is passed.

You can determine the size of the sheet with the Worksheet object’s max_row and max_column properties.
Enter the following into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> range = Calc.find_used_range(ws)
        >>> Calc.get_range_str(range)
        'A1:C7'
        >>> Calc.get_address(range)
        (com.sun.star.table.CellRangeAddress){ Sheet = (short)0x0, StartColumn = (long)0x0, StartRow = (long)0x0, EndColumn = (long)0x2, EndRow = (long)0x6 }
        >>> Calc.get_address(range).EndRow
        6
        >>> Calc.get_address(range).EndColumn
        2

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Note that the max_column property is an integer rather than the letter that appears in Excel.

.. _tute_ss_letter_number:

Converting Between Column Letters and Numbers
---------------------------------------------

To convert from letters to numbers, use the :py:class:`.TableHelper` class with the :py:meth:`~.TableHelper.col_name_to_int` method.
To convert from numbers to letters, use the :py:meth:`~.TableHelper.make_column_name` method.
Enter the following into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.table_helper import TableHelper
        >>> TableHelper.col_name_to_int('A') # Get A's number.
        1
        >>> TableHelper.col_name_to_int('AA')
        27
        >>> TableHelper.make_column_name(85)
        'CG'

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


After you import the :py:class:`.TableHelper` class from |odev| , you can use :py:meth:`~.Calc.make_column_name` and pass it an integer like ``27`` to figure out what the letter name of the ``27th`` column is.
The function :py:meth:`~.Calc.column_index_string` does the reverse: you pass it the letter name of a column, and it tells you what number that column is. You don’t need to have a workbook loaded to use these functions. If you want, you can load a workbook, get a Worksheet object, and use a Worksheet property like max_column to get an integer. Then, you can pass that integer to get_column_letter().

.. _tute_ss_rows_cols_sheet:

Getting Rows and Columns from the Sheets
----------------------------------------

You can slice Worksheet objects to get all the Cell objects in a row, column, or rectangular area of the spreadsheet.
Then you can loop over all the cells in the slice. Enter the following into the interactive shell:


.. tabs::

    .. code-tab:: python

        >>> data = Calc.get_array(sheet=ws, range_name="A1:C3")
        >>> tuple(data)
        ((42099.565300925926, 'Apples', 73.0), (42099.15373842593, 'Cherries', 85.0), (42100.532534722224, 'Pears', 14.0))
        >>> for i, r in enumerate(data):
        ...     for j, c in enumerate(r):
        ...         print(Calc.column_number_str(j)+str(i+1), c)
        ...     print('--- END OF ROW ---')
        ...
        A1 42099.565300925926
        B1 Apples
        C1 73.0
        --- END OF ROW ---
        A2 42099.15373842593
        B2 Cherries
        C2 85.0
        --- END OF ROW ---
        A3 42100.532534722224
        B3 Pears
        C3 14.0
        --- END OF ROW ---

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Here, we specify that we want the Cell objects in the rectangular area from ``A1`` to ``C3``, and we get a Generator object containing the Cell objects in that area.
To help us visualize this Generator object, we can use ``tuple()`` on it to display its Cell objects in a tuple, alternatively use the :py:meth:`.Calc.print_array`.

This tuple contains three tuples: one for each row, from the top of the desired area to the bottom.
Each of these three inner tuples contains the Cell objects in one row of our desired area, from the leftmost cell to the right.
So overall, our slice of the sheet contains all the Cell objects in the area from ``A1`` to ``C3``, starting from the top-left cell and ending with the bottom-right cell.

To print the values of each cell in the area, we use two for loops.
The outer for loop goes over each row in the slice.
Then, for each row, the nested for loop goes through each cell in that row.

To access the values of cells in a particular row or column, you can also use a Worksheet object’s rows and columns interface.
These properties must be converted to lists with the ``list()`` function before you can use the square brackets and an index with them.
Enter the following into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> list(Calc.get_col(ws,1))
        ['Apples', 'Cherries', 'Pears', 'Oranges', 'Apples', 'Bananas', 'Strawberries']

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Using the rows property on a Worksheet object will give you a tuple of tuples.
Each of these inner tuples represents a row, and contains the Cell objects in that row.
The columns property also gives you a tuple of tuples, with each of the inner tuples containing the Cell objects in a particular column.
For ``example.xlsx``, since there are ``7`` rows and ``3`` columns, rows gives us a tuple of ``7`` tuples (each containing ``3`` Cell objects), and columns gives us a tuple of ``3`` tuples (each containing ``7`` Cell objects).

To access one particular tuple, you can refer to it by its index in the larger tuple.
For example, to get the tuple that represents column ``B``, you use ``list(sheet.columns)[1]``.
To get the tuple containing the Cell objects in column A, you’d use ``list(sheet.columns)[0]``.
Once you have a tuple representing one row or column, you can loop through its Cell objects and print their values.

.. _tute_ss_wb_sheet_cells:

Workbooks, Sheets, Cells
------------------------

As a quick review, here’s a rundown of all the functions, methods, and data types involved in reading a cell out of a spreadsheet file:


| Import the |odev| modules.
| Get a Workbook object.
| Use the active or sheetnames properties.
| Get a Worksheet object.
| Use indexing or the cell() sheet method with row and column keyword arguments.
| Get a Cell object.
| Read the Cell object’s value property.

This section is finished so close the doc and office:

.. tabs::

    .. code-tab:: python

        >>> Lo.close_doc(wb)
        >>> Lo.close_office()
        True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _tute_ss_proj_read_data_sheet:

Project: Reading Data from a Spreadsheet
========================================

Say you have a spreadsheet of data from the 2010 US Census and you have the boring task of going through its thousands of rows to count both the total population and the number of census tracts for each county.
(A census tract is simply a geographic area defined for the purposes of the census.)
Each row represents a single census tract. We’ll name the spreadsheet file ``censuspopdata.xlsx``, and you can download it from `<https://nostarch.com/automatestuff2/>`__.
Its contents look like :numref:`tute_ss_fig_censuspopdata_sht`.

.. cssclass:: diagram invert

    .. _tute_ss_fig_censuspopdata_sht:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299730-026a12e8-1105-4637-ad7b-13914a247fc7.png
        :alt: The censuspopdata.xlsx spreadsheet
        :figclass: align-center

        :The ``censuspopdata.xlsx`` spreadsheet

Even though Excel can calculate the sum of multiple selected cells, you’d still have to select the cells for each of the 3,000-plus counties.
Even if it takes just a few seconds to calculate a county’s population by hand, this would take hours to do for the whole spreadsheet.

In this project, you’ll write a script that can read from the census spreadsheet file and calculate statistics for each county in a matter of seconds.

This is what your program does:

.. cssclass:: ul-list

    - Reads the data from the Excel spreadsheet
    - Counts the number of census tracts in each county
    - Counts the total population of each county
    - Prints the results

This means your code will need to do the following:

.. cssclass:: ul-list

    - Open and read the cells of an Excel document with |odev| modules
    - Calculate all the tract and population data and store it in a data structure
    - Write the data structure to a text file with the ``.py`` extension using the pprint module

.. _tute_ss_step_read_sheet_data:

Step 1: Read the Spreadsheet Data
---------------------------------

There is just one sheet in the ``censuspopdata.xlsx`` spreadsheet, named 'Population by Census Tract', and each row holds the data for a single census tract.
The columns are the tract number ``A``, the state abbreviation ``B``, the county name ``C``, and the population of the tract ``D``.

Open a new file editor tab and enter the following code. Save the file as ``readCensusExcel.py``.

.. tabs::

    .. code-tab:: python

        #! python3
        # readCensusExcel.py - Tabulates population and number of census tracts for
        # each county.

        import pprint
        from ooodev.utils.lo import Lo
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.date_time_util import DateUtil

        _ = Lo.load_office(Lo.ConnectSocket())
        print('Opening workbook...')
        wb = Calc.open_doc('censuspopdata.xlsx')
        GUI.set_visible(is_visible=True, doc=wb)

        sheet = Calc.get_sheet(doc=wb, sheet_name='Population by Census Tract')
        county_data = {}

        # TODO: Fill in county_data with each county's population and tracts.

        print('Reading rows...')
        for row in range(2, Calc.get_row_used_last_index(sheet) + 2):
            # Each row in the spreadsheet has data for one census tract.
            state  = Calc.get_val(sheet, 'B' + str(row))
            county = Calc.get_val(sheet, 'C' + str(row))
            pop    = Calc.get_val(sheet, 'D' + str(row))

        # TODO: Open a new text file and write the contents of county_data to it.

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This code imports the |odev| modules, as well as the ``pprint`` module that you’ll use to print the final county data.
Then it opens the ``censuspopdata.xlsx`` file, gets the sheet with the census data, and begins iterating over its rows.

Note that you’ve also created a variable named ``county_data``, which will contain the populations and number of tracts you calculate for each county.
Before you can store anything in it, though, you should determine exactly how you’ll structure the data inside it.

.. _tute_ss_step_pop_data_structure:

Step 2: Populate the Data Structure
-----------------------------------

The data structure stored in ``county_data`` will be a dictionary with state abbreviations as its keys.
Each state abbreviation will map to another dictionary, whose keys are strings of the county names in that state.
Each county name will in turn map to a dictionary with just two keys, ``tracts`` and ``pop``.
These keys map to the number of census tracts and population for the county.
For example, the dictionary will look similar to this:

.. tabs::

    .. code-tab:: python

        {'AK': {'Aleutians East': {'pop': 3141, 'tracts': 1},
                'Aleutians West': {'pop': 5561, 'tracts': 2},
                'Anchorage': {'pop': 291826, 'tracts': 55},
                'Bethel': {'pop': 17013, 'tracts': 3},
                'Bristol Bay': {'pop': 997, 'tracts': 1},

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None



If the previous dictionary were stored in ``county_data``, the following expressions would evaluate like this:

.. tabs::

    .. code-tab:: python

        >>> county_data['AK']['Anchorage']['pop']
        291826
        >>> county_data['AK']['Anchorage']['tracts']
        55

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


More generally, the ``county_data`` dictionary’s keys will look like this:

.. tabs::

    .. code-tab:: python

        county_data[state abbrev][county]['tracts']
        county_data[state abbrev][county]['pop']

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Now that you know how ``county_data`` will be structured, you can write the code that will fill it with the county data.
Add the following code to the bottom of your program:

.. tabs::

    .. code-tab:: python

        #! python 3
        # readCensusExcel.py - Tabulates population and number of census tracts for
        # each county.

        print('Reading rows...')
        for row in range(2, Calc.get_row_used_last_index(sheet) + 2):
            # Each row in the spreadsheet has data for one census tract.
            state  = Calc.get_val(sheet, 'B' + str(row))
            county = Calc.get_val(sheet, 'C' + str(row))
            pop    = Calc.get_val(sheet, 'D' + str(row))
            # Make sure the key for this state exists.
            _ = county_data.setdefault(state, {})
            # Make sure the key for this county in this state exists.
            _ = county_data[state].setdefault(county, {'tracts': 0, 'pop': 0})
            # Each row represents one census tract, so increment by one.
            county_data[state][county]['tracts'] += 1
            # Increase the county pop by the pop in this census tract.
            county_data[state][county]['pop'] += int(pop)

        # TODO: Open a new text file and write the contents of county_data to it.

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The last two lines of code perform the actual calculation work, incrementing the value for tracts and increasing the value for pop for the current county on each iteration of the for loop.

The other code is there because you cannot add a county dictionary as the value for a state abbreviation key until the key itself exists in ``county_data``
(that is, ``county_data['AK']['Anchorage']['tracts'] += 1`` will cause an error if the ``AK`` key doesn’t exist yet).
To make sure the state abbreviation key exists in your data structure, you need to call the ``setdefault()`` method to set a value if one does not already exist for state.

Just as the county_data dictionary needs a dictionary as the value for each state abbreviation key, each of those dictionaries will need its own dictionary as the value for each county key.
And each of those dictionaries in turn will need keys ``tracts`` and ``pop`` that start with the integer value ``0``
(if you ever lose track of the dictionary structure, look back at the example dictionary at the start of this section).

Since ``setdefault()`` will do nothing if the key already exists, you can call it on every iteration of the for loop without a problem.

.. _tute_ss_step_write_file:

Step 3: Write the Results to a File
-----------------------------------

After the for loop has finished, the ``county_data`` dictionary will contain all of the population and tract information keyed by county and state.
At this point, you could program more code to write this to a text file or another Excel spreadsheet.
For now, let’s just use the ``pprint.pformat()`` function to write the ``county_data`` dictionary value as a massive string to a file named ``census2010.py``.
Add the following code to the bottom of your program (making sure to keep it unindented so that it stays outside the for loop):

.. tabs::

    .. code-tab:: python

        #! python 3
        # readCensusExcel.py - Tabulates population and number of census tracts for
        # each county.

        # --snip--

        for row in range(2, Calc.get_row_used_last_index(sheet)-1):
        # --snip--

        # Open a new text file and write the contents of county_data to it.
        print('Writing results...')
        result_file = open('census2010.py', 'w')
        result_file.write('allData = ' + pprint.pformat(county_data))
        result_file.close()
        print('Done.')

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


        

The ``pprint.pformat()`` function produces a string that itself is formatted as valid Python code.
By outputting it to a text file named ``census2010.py``, you’ve generated a Python program from your Python program!
This may seem complicated, but the advantage is that you can now import ``census2010.py`` just like any other Python module.
In the interactive shell, change the current working directory to the folder with your newly created ``census2010.py`` file and then import it:

.. tabs::

    .. code-tab:: python

        >>> import os
        >>> import census2010
        >>> census2010.allData['AK']['Anchorage']
        {'pop': 291826, 'tracts': 55}
        >>> anchoragePop = census2010.allData['AK']['Anchorage']['pop']
        >>> print('The 2010 population of Anchorage was ' + str(anchoragePop))
        The 2010 population of Anchorage was 291826

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The ``readCensusExcel.py`` program was throwaway code: once you have its results saved to ``census2010.py``, you won’t need to run the program again.
Whenever you need the county data, you can just run ``import census2010``.

Calculating this data by hand would have taken hours; this program did it in a few seconds.
Using |odev|, you will have no trouble extracting information that is saved to an Excel spreadsheet and performing calculations on it.
You can download the complete program from `<https://nostarch.com/automatestuff2/>`__.

.. tabs::

    .. code-tab:: python

        >>> #! python3
        >>> # readCensusExcel.py - Tabulates population and number of census tracts for
        >>> # each county.
        >>>
        >>> import pprint
        >>> from ooodev.utils.lo import Lo
        >>> from ooodev.office.calc import Calc
        >>> from ooodev.utils.gui import GUI
        >>>
        >>> _ = Lo.load_office(Lo.ConnectSocket(soffice="C:\\Program Files\\LibreOfficeDev 7\\program\\soffice.exe"))
        >>> print('Opening workbook...')
        Opening workbook...
        >>> wb = Calc.open_doc('censuspopdata.xlsx')
        >>> GUI.set_visible(is_visible=True, doc=wb)
        >>>
        >>> sheet = Calc.get_sheet(doc=wb, sheet_name='Population by Census Tract')
        >>> county_data = {}
        >>>
        >>> range_name = 'B2:D' + str(Calc.get_row_used_last_index(sheet)+1)
        >>> # print(range_name)
        >>> data = Calc.get_array(sheet=sheet, range_name=range_name)
        >>>
        >>> print('Reading rows...')
        Reading rows...
        >>> for i, row in enumerate(data):
        ...     # Each row in the spreadsheet has data for one census tract.
        ...     state, county, pop = row
        ...     # Make sure the key for this state exists.
        ...     _ = county_data.setdefault(state, {})
        ...     # Make sure the key for this county in this state exists.
        ...     _ = county_data[state].setdefault(county, {'tracts': 0, 'pop': 0})
        ...     # Each row represents one census tract, so increment by one.
        ...     county_data[state][county]['tracts'] += 1
        ...     # Increase the county pop by the pop in this census tract.
        ...     county_data[state][county]['pop'] += int(pop)
        ...
        >>>
        >>> # Open a new text file and write the contents of county_data to it.
        >>> print('Writing results...')
        Writing results...
        >>> result_file = open('census2010B.py', 'w')
        >>> result_file.write('allData = ' + pprint.pformat(county_data))
        152237
        >>> result_file.close()
        >>> print('Done.')
        Done.
        >>>
        >>> import os
        >>> import census2010
        >>> census2010.allData['AK']['Anchorage']
        {'pop': 291826, 'tracts': 55}
        >>> anchoragePop = census2010.allData['AK']['Anchorage']['pop']
        >>> print('The 2010 population of Anchorage was ' + str(anchoragePop))
        The 2010 population of Anchorage was 291826
        >>>
        >>> Lo.close_doc(wb)
        >>> Lo.close_office()
        True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _tute_ss_ideas_programs:

Ideas for Similar Programs
--------------------------

Many businesses and offices use Excel to store various types of data, and it’s not uncommon for spreadsheets to become large and unwieldy.
Any program that parses an Excel spreadsheet has a similar structure: it loads the spreadsheet file, preps some variables or data structures, and then loops through each of the rows in the spreadsheet.
Such a program could do the following:

.. cssclass:: ul-list

    - Compare data across multiple rows in a spreadsheet.
    - Open multiple Excel files and compare data between spreadsheets.
    - Check whether a spreadsheet has blank rows or invalid data in any cells and alert the user if it does.
    - Read data from a spreadsheet and use it as the input for your Python programs.

.. _tute_ss_writing_sheet_docs:

Writing Spreadsheet Documents
=============================

|odev| also provides ways of writing data, meaning that your programs can create and edit spreadsheet files.
With Python, it’s simple to create spreadsheets with thousands of rows of data.

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> loader = Lo.load_office(Lo.ConnectSocket(headless=True))
        >>> # use the Office API... NOTE: Following lines raise an error
        >>> Lo.close_doc(wb)
        Closing the document
        >>> Lo.close_office()
        Closing Office
        Office has already been requested to terminate
        True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _tute_ss_create_save_sheet_docs:

Creating and Saving Spreadsheet Documents
-----------------------------------------

Start a lo instance and use the Calc create_doc class to create a new, blank Workbook object.
Enter the following into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> loader = Lo.load_office(Lo.ConnectSocket(headless=True))
        >>>
        >>> from ooodev.office.calc import Calc
        >>> wb = Calc.create_doc(loader=loader)
        >>> ws = Calc.get_sheet(doc=wb, index=0)
        >>> Calc.get_sheet_name(ws)
        'Sheet1'
        >>> Calc.set_sheet_name(ws, 'Spam Bacon Eggs Sheet')
        True
        >>> Calc.get_sheet_name(ws)
        'Spam Bacon Eggs Sheet'
        >>> Calc.get_sheet_names(wb)
        ('Spam Bacon Eggs Sheet',)
        >>> Calc.save_doc(wb, "foo.ods")
        >>>
        >>> Lo.close_doc(wb)
        >>> Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The workbook will start off with a single sheet named Sheet.
You can change the name of the sheet using the :py:meth:`.Calc.set_sheet_name` method which stores a new string in its title property.

Any time you modify the Workbook object or its sheets and cells, the spreadsheet file will not be saved until you call the :py:meth:`.Calc.save_doc` workbook method.
Enter the following into the interactive shell (with ``example.xlsx`` in the current working directory):

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> loader = Lo.load_office(Lo.ConnectSocket(headless=True))
        >>>
        >>> from ooodev.office.calc import Calc
        >>> wb = Calc.open_doc('example.ods', loader)
        >>> ws = Calc.get_sheet(wb, 0)
        >>> Calc.set_sheet_name(ws, 'Spam Spam Spam')
        True
        >>> Calc.save_doc(wb, 'example_copy.ods')
        >>>
        >>> Lo.close_doc(wb)
        >>> Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Here, we change the name of our sheet. To save our changes, we pass a filename as a string to the :py:meth:`.Calc.save_doc` method.
Passing a different filename than the original, such as ``example_copy.xlsx``, saves the changes to a copy of the spreadsheet.

Whenever you edit a spreadsheet you’ve loaded from a file, you should always save the new, edited spreadsheet to a different filename than the original.
That way, you’ll still have the original spreadsheet file to work with in case a bug in your code caused the new, saved file to have incorrect or corrupt data.

.. _tute_ss_create_remove_shts:

Creating and Removing Sheets
----------------------------

Sheets can be added to and removed from a workbook with the :py:meth:`.Calc.insert_sheet` and :py:meth:`.Calc.remove_sheet` methods.
Enter the following into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> loader = Lo.load_office(Lo.ConnectSocket(headless=True))
        >>>
        >>> from ooodev.office.calc import Calc
        >>> wb = Calc.create_doc(loader=loader)
        Creating Office document scalc
        >>> ws = Calc.get_sheet(doc=wb, index=0)
        >>> Calc.get_sheet_names(wb)
        ('Sheet1',)
        >>> Calc.insert_sheet(wb, 'Sheet2', 1)
        >>> Calc.get_sheet_names(wb)
        ('Sheet1', 'Sheet2')
        >>> Calc.insert_sheet(wb, 'First Sheet', 0)
        >>> Calc.get_sheet_names(wb)
        ('First Sheet', 'Sheet1', 'Sheet2')
        >>> Calc.insert_sheet(wb, 'Middle Sheet', 2)
        >>> Calc.get_sheet_names(wb)
        ('First Sheet', 'Sheet1', 'Middle Sheet', 'Sheet2')

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The :py:meth:`.Calc.insert_sheet` method returns a new Worksheet object named ``SheetX``, which by default is set to be the last sheet in the workbook.
Optionally, the name and index of the new sheet can be specified with the name and index keyword arguments.

Continue the previous example by entering the following:

.. tabs::

    .. code-tab:: python

        >>> Calc.get_sheet_names(wb)
        ('First Sheet', 'Sheet1', 'Middle Sheet', 'Sheet2')
        >>> Calc.remove_sheet(wb, 'Middle Sheet')
        True
        >>> Calc.remove_sheet(wb, 'Sheet2')
        True
        >>> Calc.get_sheet_names(wb)
        ('First Sheet', 'Sheet1')
        >>> Lo.close_doc(wb)
        >>> Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


You can use the :py:meth:`.Calc.remove_sheet` method to remove a sheet from a workbook, similarly to deleting a key-value pair from a dictionary.

Remember to call the :py:meth:`.Calc.save_doc` method to save the changes after adding sheets to or removing sheets from the workbook.

.. _tute_ss_vals_cells:

Writing Values to Cells
-----------------------

Writing values to cells is much like writing values to keys in a dictionary.
Enter this into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> loader = Lo.load_office(Lo.ConnectSocket(headless=True))
        >>>
        >>> from ooodev.office.calc import Calc
        >>> wb = Calc.create_doc(loader=loader)
        Creating Office document scalc
        >>> ws = Calc.get_sheet(doc=wb, index=0)
        >>> Calc.set_val('Hello, world!', ws, 'A1')
        >>> Calc.get_string(ws, 'A1')
        'Hello, world!'
        >>> Lo.close_doc(wb)
        >>> Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

If you have the cell’s coordinate as a string, you can use it just like a dictionary key on the Worksheet object to specify which cell to write to.

.. _tute_ss_updating_sheet:

Project: Updating a Spreadsheet
===============================

In this project, you’ll write a program to update cells in a spreadsheet of produce sales.
Your program will look through the spreadsheet, find specific kinds of produce, and update their prices.
Download this spreadsheet from `<https://nostarch.com/automatestuff2/>`__. :numref:`tute_ss_fig_produce_sht`  shows what the spreadsheet looks like.

.. cssclass:: diagram invert

    .. _tute_ss_fig_produce_sht:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299752-dd9cdbe8-7171-4312-a578-c3e1b699b042.png
        :alt: A spreadsheet of produce sales
        :figclass: align-center

        :A spreadsheet of produce sales

Each row represents an individual sale.
The columns are the type of produce sold ``A``, the cost per pound of that produce ``B``, the number of pounds sold ``C``, and the total revenue from the sale ``D``.
The TOTAL column is set to the Excel formula`` =ROUND(B3*C3, 2)``, which multiplies the cost per pound by the number of pounds sold and rounds the result to the nearest cent.
With this formula, the cells in the TOTAL column will automatically update themselves if there is a change in column ``B`` or ``C``.

Now imagine that the prices of garlic, celery, and lemons were entered incorrectly,
leaving you with the boring task of going through thousands of rows in this spreadsheet to update the cost per pound for any garlic, celery, and lemon rows.
You can’t do a simple find-and-replace for the price, because there might be other items with the same price that you don’t want to mistakenly “correct.” For thousands of rows, this would take hours to do by hand.
But you can write a program that can accomplish this in seconds.

See Also: :ref:`ch23`

Your program does the following:

.. cssclass:: ul-list

    - Loops over all the rows
    - If the row is for garlic, celery, or lemons, changes the price

This means your code will need to do the following:

.. cssclass:: ul-list

    - Open the spreadsheet file.
    - For each row, check whether the value in column A is Celery, Garlic, or Lemon.
    - If it is, update the price in column B.
    - Save the spreadsheet to a new file (so that you don’t lose the old spreadsheet, just in case).

.. _tute_ss_step_set_data_structure:

Step 1: Set Up a Data Structure with the Update Information
-----------------------------------------------------------

The prices that you need to update are as follows:

::

    Celery         1.19
    Garlic         3.07
    Lemon          1.27

You could write code like this:

.. tabs::

    .. code-tab:: python

        if produceName == 'Celery':
            cellObj = 1.19
        if produceName == 'Garlic':
            cellObj = 3.07
        if produceName == 'Lemon':
            cellObj = 1.27

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Having the produce and updated price data hardcoded like this is a bit inelegant. If you needed to update the spreadsheet again with different prices or different produce, you would have to change a lot of the code. Every time you change code, you risk introducing bugs.

A more flexible solution is to store the corrected price information in a dictionary and write your code to use this data structure. In a new file editor tab, enter the following code:

.. todo::

    Tute ss. This section seems to be half pseudocode but openpyxl needs to go to odev

    Re fix this. Needs to be referred back to original doc for context.
    Formatting is really screwy in this section too

    [*** FIX THIS ***

    #! python3
    # updateProduce.py - Corrects costs in produce sales spreadsheet.

    import ***openpyxl***************************************************************************************

    wb = ***openpyxl***.load_workbook('produceSales.xlsx')
    sheet = wb['Sheet']

.. tabs::

    .. code-tab:: python

        # The produce types and their updated prices
        PRICE_UPDATES = {'Garlic': 3.07,
                        'Celery': 1.19,
                        'Lemon': 1.27}

        # TODO: Loop through the rows and update the prices.

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Save this as ``updateProduce.py``.
If you need to update the spreadsheet again, you’ll need to update only the ``PRICE_UPDATES`` dictionary, not any other code.

.. _tute_ss_step_update_row_prices:

Step 2: Check All Rows and Update Incorrect Prices
--------------------------------------------------

The next part of the program will loop through all the rows in the spreadsheet.
Add the following code to the bottom of ``updateProduce.py``:

.. todo:: 
    Tute SS, fix code section below: Loop through the rows and update the prices.

.. tabs::

    .. code-tab:: python

        #! python3
        # updateProduce.py - Corrects costs in produce sales spreadsheet.

        # --snip--

        # Loop through the rows and update the prices.
        for rowNum in range(2, sheet.max_row):    # skip the first row
            produceName = sheet.cell(row=rowNum, column=1).value
            if produceName in PRICE_UPDATES:
                sheet.cell(row=rowNum, column=2).value = PRICE_UPDATES[produceName]

        wb.save('updatedProduceSales.xlsx')

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


We loop through the rows starting at row ``2``, since row 1 is just the header ➊.
The cell in column ``1`` (that is, column ``A``) will be stored in the variable produceName ➋.
If produceName exists as a key in the ``PRICE_UPDATES`` dictionary ➌, then you know this is a row that must have its price corrected.
The correct price will be in ``PRICE_UPDATES[produceName]``.

Notice how clean using ``PRICE_UPDATES`` makes the code.
Only one if statement, rather than code like if ``produceName == 'Garlic'``: , is necessary for every type of produce to update.
And since the code uses the ``PRICE_UPDATES`` dictionary instead of hardcoding the produce names and updated costs into the for loop,
you modify only the ``PRICE_UPDATES`` dictionary and not the code if the produce sales spreadsheet needs additional changes.

After going through the entire spreadsheet and making changes, the code saves the Workbook object to ``updatedProduceSales.xlsx`` ➍.
It doesn’t overwrite the old spreadsheet just in case there’s a bug in your program and the updated spreadsheet is wrong.
After checking that the updated spreadsheet looks right, you can delete the old spreadsheet.

You can download the complete source code for this program from `<https://nostarch.com/automatestuff2/>`__.

.. _tute_ss_ideas_simalar_programs:

Ideas for Similar Programs
--------------------------

Since many office workers use Excel spreadsheets all the time, a program that can automatically edit and write Excel files could be really useful.
Such a program could do the following:

Read data from one spreadsheet and write it to parts of other spreadsheets.
Read data from websites, text files, or the clipboard and write it to a spreadsheet.
Automatically “clean up” data in spreadsheets.
For example, it could use regular expressions to read multiple formats of phone numbers and edit them to a single, standard format.

.. _tute_ss_set_cell_font_style:

Setting the Font Style of Cells
===============================

Styling certain cells, rows, or columns can help you emphasize important areas in your spreadsheet.
In the produce spreadsheet, for example, your program could apply bold text to the potato, garlic, and parsnip rows.
Or perhaps you want to italicize every row with a cost per pound greater than ``$5``.
Styling parts of a large spreadsheet by hand would be tedious, but your programs can do it instantly.

To customize font styles in cells the |odev| Props class and two ``ooo.dyn.awt`` import from  |ooouno|_ classes, ``FontSlant`` and ``FontWeight``, must be imported.

Note that an alias has been used on the classes to make them easier to recognise.

Here’s an example that creates a new workbook and sets cell ``A1`` to have an italicized, bold, 24-point font.
Enter the following into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> from ooodev.office.calc import Calc
        >>> from ooodev.utils.gui import GUI
        >>>
        >>> loader = Lo.load_office(Lo.ConnectSocket())
        >>> doc = Calc.create_doc()
        >>> GUI.set_visible(is_visible=True, doc=doc)
        >>> sheet = Calc.get_sheet(doc=doc)
        >>> for i in range(1, 6): # create some data in column A
        ...     Calc.set_val(i, sheet, 'A'+str(i))
        ...
        >>> from ooodev.utils.props import Props
        >>> from ooo.dyn.awt.font_slant import FontSlant
        >>> from ooo.dyn.awt.font_weight import FontWeight
        >>>
        >>> cell = Calc.get_cell(sheet, 'A1')
        >>> Props.set(cell, CharPosture=FontSlant.ITALIC, CharWeight=FontWeight.BOLD, CharHeight=24,)
        >>> _ = Calc.save_doc(doc, "sampleChart.xlsx")
        >>> # check file
        >>> Lo.close_doc(doc=doc)
        >>> _ = Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

In this example, :py:meth:`.Calc.get_cell` returns an XCell_ type with is used to reference the cell in :py:meth:`.Props.set` and set the properties directly.
``CharPosture`` and ``CharWeight`` use the ``FontSlat`` and ``FontWeight`` classes respectively as previously imported.
``CharHeight`` is set directly. The effect is shown in the saved file.

.. _tute_ss_font_objects:

Font Objects
============

A number of |odev| classes have methods to change font properties.
:numref:`tute_ss_tbl_props_for_font_objects` shows key properties for Font objects.

..
    Table 13-2

.. _tute_ss_tbl_props_for_font_objects:

.. table:: Properties for Font Objects.
    :name: props_for_font_objects

    +-----------------+-----------+---------------------------------+
    | Property        | Data type | Description                     |
    +=================+===========+=================================+
    |name             +String     + The font name, such as 'Calibri'|
    |                 +           + or 'Times New Roman'            |
    +-----------------+-----------+---------------------------------+
    |size             +Integer    +The point size                   |
    +-----------------+-----------+---------------------------------+
    |bold             +Boolean    +True, for bold font              |
    +-----------------+-----------+---------------------------------+
    |italic           +Boolean    +True, for italic font            |
    +-----------------+-----------+---------------------------------+

The best way of setting font attributes is to define a style and apply it to the required objects.
In this example a spreadsheet is created the a style is; named, created, properties set, and applied to a cell object.
The cell value is then set which demonstrates the new style, and the process is repeated again.

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> from ooodev.office.calc import Calc
        >>> from ooodev.utils.gui import GUI
        >>>
        >>> loader = Lo.load_office(Lo.ConnectSocket())
        >>> doc = Calc.create_doc()
        >>> GUI.set_visible(is_visible=True, doc=doc)
        >>> sheet = Calc.get_sheet(doc=doc)

        >>> from ooodev.utils.props import Props
        >>> from ooo.dyn.awt.font_slant import FontSlant
        >>> from ooo.dyn.awt.font_weight import FontWeight
        >>>
        >>> # Name style
        >>> HEADER_STYLE_NAME = "My HeaderStyle"
        >>> # Create style
        >>> style1 = Calc.create_cell_style(doc=doc, style_name=HEADER_STYLE_NAME)
        >>> # Set style properties
        >>> Props.set(style1, CharWeight=FontWeight.BOLD, CharHeight=14,)
        >>> # Apply style
        >>> Calc.change_style(sheet=sheet, style_name=HEADER_STYLE_NAME, range_name="A1")
        >>> # Set cell value
        >>> Calc.set_val('Bold Times New Roman', sheet, 'A1')
        >>> # Repeat for data
        >>> DATA_STYLE_NAME = "My DataStyle"
        >>> style2 = Calc.create_cell_style(doc=doc, style_name=DATA_STYLE_NAME)
        >>> Props.set(style2, CharPosture=FontSlant.ITALIC, CharHeight=24,)
        >>> Calc.change_style(sheet=sheet, style_name=DATA_STYLE_NAME, range_name="B3")
        >>> Calc.set_val('24 pt Italic', sheet, 'B3')
        >>> _ = Calc.save_doc(doc, "styles.xlsx")

        >>> # check file
        >>> Lo.close_doc(doc=doc)
        >>> _ = Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Here, we store a style name in a ``STYLE_NAME`` constant, create a style with :py:meth:`.Calc.create_cell_style` method,
use :py:meth:`.Props.set` method to set the style properties, then set the cell value with the :py:meth:`.Calc.set_val` method.
We repeat the process with another style for a second cell.
After you run this code, the styles of the ``A1`` and ``B3`` cells in the spreadsheet will be set to custom character styles, as shown in :numref:`tute_ss_fig_custom_font_styles`.

..
    Figure 13-4

.. cssclass:: diagram invert

    .. _tute_ss_fig_custom_font_styles:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299766-0bfc9ef8-9675-4266-80b8-c8c57059f2ea.png
        :alt: A spreadsheet with custom font styles
        :figclass: align-center

        :A spreadsheet with custom font styles

.. todo:: 

    Tute ss: Correct how to set a font for a cell.

For cell A1, we set the font name to ``Times New Roman`` and set bold to true, so our text appears in bold Times New Roman.
We didn’t specify a size, so the default is used.
In cell ``B3``, our text is italic, with a size of ``24``; we didn’t specify a font name, so the default, ``Calibri``, is used.

.. _tute_ss_formulas:

Formulas
========

Spreadsheet formulas, which begin with an equal sign, can configure cells to contain values calculated from other cells.
In this section, you’ll use :py:meth:`.Calc.set_val` to set a formula on a cell, just like any normal value.
For example:

.. tabs::

    .. code-tab:: python

        >>> Calc.set_val(sheet=sheet, cell_name="B9", value="=SUM(B1:B8)")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


This will store ``=SUM(B1:B8)`` as the value in cell ``B9``. This sets the ``B9`` cell to a formula that calculates the sum of values in cells ``B1`` to ``B8``.
You can see this in action in :numref:`tute_ss_figb9_b1_b8`.

.. cssclass:: diagram invert

    .. _tute_ss_figb9_b1_b8:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299779-ff5d2bfa-8e36-4606-8bd3-e48a0704a80d.png
        :alt: :Cell B9 contains the formula =SUM(B1:B8), which adds the cells B1 to B8
        :figclass: align-center

        :Cell ``B9`` contains the formula ``=SUM(B1:B8)``, which adds the cells ``B1`` to ``B8``

A formula is set just like any other text value in a cell. Enter the following into the interactive shell:

See also :ref:`ch20_storing_2d_arrays`.

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> from ooodev.office.calc import Calc
        >>> from ooodev.utils.gui import GUI
        >>>
        >>> loader = Lo.load_office(Lo.ConnectSocket())
        >>> doc = Calc.create_doc()
        >>> GUI.set_visible(is_visible=True, doc=doc)
        >>> sheet = Calc.get_sheet(doc=doc)
        >>> Calc.set_val(sheet=sheet, cell_name='A1', value=200)
        >>> Calc.set_val(sheet=sheet, cell_name='A2', value=300)
        >>> Calc.set_val(sheet=sheet, cell_name="A3", value="=SUM(A1:A2)") # Set the formula
        >>> _ = Calc.save_doc(doc, "writeFormula.xlsx")
        >>> # check file
        >>> Lo.close_doc(doc=doc)
        >>> _ = Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None




The cells in ``A1`` and ``A2`` are set to ``200`` and ``300``, respectively with the :py:meth:`.Calc.set_val` method.
The value in cell ``A3`` is set to a formula that sums the values in ``A1`` and ``A2``.
When the spreadsheet is opened, ``A3`` will display its value as ``500``.

Formulas offer a level of programmability for spreadsheets but can quickly become unmanageable for complicated tasks.
For example, even if you’re deeply familiar with formulas, it’s a headache to try to decipher what the following actually does:

::

    =IFERROR(TRIM(IF(LEN(VLOOKUP(F7, Sheet2!$A$1:$B$10000, 2, FALSE))>0,SUBSTITUTE(VLOOKUP(F7, Sheet2!$A$1:$B$10000, 2, FALSE), " ", ""),"")), "")

Python code is much more readable.

.. _tute_ss_adjusting_rows_cols:

Adjusting Rows and Columns
==========================

Adjusting the sizes of rows and columns is as easy as clicking and dragging the edges of a row or column header.
But if you need to set a row or column’s size based on its cells’ contents or if you want to set sizes in a large number of spreadsheet files, it will be much quicker to write a Python program to do it.

Rows and columns can also be hidden entirely from view.
Or they can be “frozen” in place so that they are always visible on the screen and appear on every page when the spreadsheet is printed (which is handy for headers).

.. _tute_ss_setting_row_col_width:

Setting Row Height and Column Width
-----------------------------------

.. todo::

    Tute ss, setting row and height seection needs serious review and updates.

Worksheet objects have row_dimensions and ``column_dimensions`` properties that control row heights and column widths.
Enter this into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> from ooodev.office.calc import Calc
        >>> from ooodev.utils.gui import GUI
        >>>
        >>> loader = Lo.load_office(Lo.ConnectSocket())
        >>> doc = Calc.create_doc()
        >>> GUI.set_visible(is_visible=True, doc=doc)
        >>> sheet = Calc.get_sheet(doc=doc)
        >>> Calc.set_val(sheet=sheet, cell_name='A1', value='Tall row')
        >>> Calc.set_val(sheet=sheet, cell_name='B2', value='Wide column',)
        >>> # Set the height and width:
        >>> _ = Calc.set_row_height(sheet=sheet, height=70, idx=0)
        >>> _ = Calc.set_col_width(sheet=sheet, width=40, idx=1)
        >>> _ = Calc.save_doc(doc, 'dimensions.xlsx')
        >>> # check file
        >>> Lo.close_doc(doc=doc)
        >>> _ = Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


A sheet’s row_dimensions and ``column_dimensions`` are dictionary-like values; ``row_dimensions`` contains ``RowDimension`` objects and ``column_dimensions`` contains ``ColumnDimension`` objects.
In ``row_dimensions``, you can access one of the objects using the number of the row (in this case, ``1`` or ``2``).
In ``column_dimensions``, you can access one of the objects using the letter of the column (in this case, ``A`` or ``B``).

The ``dimensions.xlsx`` spreadsheet looks like :numref:`tute_ss_fig_rot1b_larger`.

..
    Figure 13-6

.. cssclass:: diagram invert

    .. _tute_ss_fig_rot1b_larger:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299789-682e72d3-b7f5-44c2-b941-96bc0854b41c.png
        :alt: Row 1 and column B set to larger heights and widths
        :figclass: align-center

        :Row ``1`` and column ``B`` set to larger heights and widths

Once you have the RowDimension object, you can set its height.
Once you have the ``ColumnDimension`` object, you can set its width.
The row height can be set to an integer or float value between 0 and 409.
This value represents the height measured in points, where one point equals 1/72 of an inch.
The default row height is 12.75. The column width can be set to an integer or float value between 0 and 255.
This value represents the number of characters at the default font size (11 point) that can be displayed in the cell.
The default column width is 8.43 characters.
Columns with widths of 0 or rows with heights of 0 are hidden from the user.

Merging and Unmerging Cells
---------------------------

.. todo::

    Tute ss, Merging and Unmerging Cells section.
    Calc will be getting a merge_cells() method and this section needs to reflect that.

A rectangular area of cells can be merged into a single cell with the ``merge_cells()`` sheet method.
Enter the following into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> from ooodev.office.calc import Calc
        >>> from ooodev.utils.gui import GUI
        >>>
        >>> loader = Lo.load_office(Lo.ConnectSocket())
        >>> doc = Calc.create_doc()
        >>> GUI.set_visible(is_visible=True, doc=doc)
        >>> sheet = Calc.get_sheet(doc=doc)
        >>> 
        >>> # Merge first few cells of the last row
        >>> cell_range = Calc.get_cell_range(sheet, 'A1:D3')
        >>> from com.sun.star.util import XMergeable
        >>> xmerge = Lo.qi(XMergeable, cell_range, True)
        >>> xmerge.merge(True)
        >>> Calc.set_val('Twelve cells merged together.', sheet, 'A1')
        >>> cell_range = Calc.get_cell_range(sheet, 'C5:D5')
        >>> xmerge = Lo.qi(XMergeable, cell_range, True)
        >>> xmerge.merge(True)
        >>> Calc.set_val('Two merged cells.', sheet, 'C5')
        >>> _ = Calc.save_doc(doc, 'merged.xlsx')

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The argument to ``merge_cells()`` is a single string of the top-left and bottom-right cells of the rectangular area to be merged: ``A1:D3`` merges ``12`` cells into a single cell.
To set the value of these merged cells, simply set the value of the top-left cell of the merged group.

When you run this code, merged.xlsx will look like :numref:`tute_ss_fig_merged_cells`.

..
    Figure 13-7

.. cssclass:: diagram invert

    .. _tute_ss_fig_merged_cells:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299799-b8b51ce7-8f6c-46f0-8aec-e62bc571c609.png
        :alt: Merged cells in a spreadsheet
        :figclass: align-center

        :Merged cells in a spreadsheet

To unmerge cells, call the ``unmerge_cells()`` sheet method.
Enter this into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> xmerge.merge(False) # Split up last merged cells
        >>> cell_range = Calc.get_cell_range(sheet, 'A1:D3')
        >>> Lo.qi(XMergeable, cell_range, True).merge(False)
        >>> _ = Calc.save_doc(doc, 'merged.xlsx')
        >>> # check file
        >>> Lo.close_doc(doc=doc)
        >>> _ = Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


If you save your changes and then take a look at the spreadsheet, you’ll see that the merged cells have gone back to being individual cells.

Freezing Panes
--------------

For spreadsheets too large to be displayed all at once, it’s helpful to “freeze” a few of the top rows or leftmost columns onscreen.
Frozen column or row headers, for example, are always visible to the user even as they scroll through the spreadsheet.
These are known as freeze panes. In OpenPyXL, each Worksheet object has a freeze_panes property that can be set to a Cell object or a string of a cell’s coordinates.
Note that all rows above and all columns to the left of this cell will be frozen, but the row and column of the cell itself will not be frozen.

See Also: :ref:`ch23_freezing_rows`

To unfreeze all panes, set freeze_panes to None or ``A1``. :numref:`tute_ss_tbl_frozen_pane_ex` shows which rows and columns will be frozen for some example settings of ``freeze_panes``.

.. todo::

    Tute ss, Frozen Pane Examples table needs to be completly redone.

..
    Table 13-3

.. _tute_ss_tbl_frozen_pane_ex:

.. table:: Frozen Pane Examples.
    :name: tbl_frozen_pane_ex

    +----------------------------------------+---------------------------+
    |freeze_panes setting                    |Rows and columns frozen    |
    +========================================+===========================+
    |sheet.freeze_panes = 'A2'               |Row 1                      |
    +----------------------------------------+---------------------------+
    |sheet.freeze_panes = 'B1'               |Column A                   |
    +----------------------------------------+---------------------------+
    |sheet.freeze_panes = 'C1'               |Columns A and B            |
    +----------------------------------------+---------------------------+
    |sheet.freeze_panes = 'C2'               |Row 1 and columns A and B  |
    +----------------------------------------+---------------------------+
    |sheet.freeze_panes = 'A1'               |No frozen panes            |
    | or sheet.freeze_panes = None           |                           |
    +----------------------------------------+---------------------------+

Make sure you have the produce sales spreadsheet from `<https://nostarch.com/automatestuff2/>`__.
Then enter the following into the interactive shell:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.utils.lo import Lo
        >>> from ooodev.office.calc import Calc
        >>> from ooodev.utils.gui import GUI
        >>>
        >>> loader = Lo.load_office(Lo.ConnectSocket())
        >>> doc = Calc.open_doc('produceSales.xlsx')
        >>> GUI.set_visible(is_visible=True, doc=doc)
        >>> sheet = Calc.get_sheet(doc=doc)
        >>> Calc.goto_cell(cell_name="A1", doc=doc) # activate reference row
        >>> Calc.freeze_rows(doc=doc, num_rows=1)   # freeze one row before reference
        >>> _ = Calc.save_doc(doc, 'freezeExample.xlsx')

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

If you set the freeze_panes property to ``A2``, row ``1`` will always be viewable, no matter where the user scrolls in the spreadsheet.
You can see this in :numref:`tute_ss_fig_freeze_a2`.

..
    Figure 13-8

.. cssclass:: diagram invert

    .. _tute_ss_fig_freeze_a2:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299812-13dd64f0-5dca-4906-af52-5cf4e90e6622.png
        :alt: With freeze_panes set to A2, row 1 is always visible, even as the user scrolls down
        :figclass: align-center

        :With freeze_panes set to ``A2``, row ``1`` is always visible, even as the user scrolls down

.. _tute_ss_charts:

Charts
======

|odev| supports creating many charts including bar, line, scatter, and pie charts using the data in a sheet’s cells. To make a chart, you need to do the following:

.. cssclass:: ul-list

    - Create a Reference object from a rectangular selection of cells.
    - Create a Series object by passing in the Reference object.
    - Create a Chart object.
    - Append the Series object to the Chart object.
    - Add the Chart object to the Worksheet object, optionally specifying which cell should be the top-left corner of the chart.

The Reference object requires some explaining.
You create Reference objects by calling the ***openpyxl***.chart.Reference() function and passing three arguments:

The Worksheet object containing your chart data.
A tuple of two integers, representing the top-left cell of the rectangular selection of cells containing your chart data: the first integer in the tuple is the row, and the second is the column. Note that 1 is the first row, not 0.
A tuple of two integers, representing the bottom-right cell of the rectangular selection of cells containing your chart data: the first integer in the tuple is the row, and the second is the column.
:numref:`tute_ss_fig_tuple_vals` shows some sample coordinate arguments.

..
    Figure 13-9

.. cssclass:: diagram invert

    .. _tute_ss_fig_tuple_vals:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299822-1620a00b-f148-4ff3-9086-8f4b55c60273.png
        :alt: tuple values
        :figclass: align-center

        From left to right: (1, 1), (10, 1); (3, 2), (6, 4); (5, 3), (5, 3)

Enter this interactive shell example to create a bar chart and add it to the spreadsheet:

.. tabs::

    .. code-tab:: python

        >>> from ooodev.office.calc import Calc
        >>> from ooodev.office.chart2 import Chart2, Angle
        >>> from ooodev.utils.gui import GUI
        >>> from ooodev.utils.lo import Lo
        >>>
        >>> _ = Lo.load_office(connector=Lo.ConnectPipe(soffice="C:\\Program Files\\LibreOfficeDev 7\\program\\soffice.exe"))
        >>> doc = Calc.create_doc()
        >>> GUI.set_visible(is_visible=True, doc=doc)
        >>> sheet = Calc.get_sheet(doc=doc)
        >>> for i in range(1, 11): # create some data in column A
        ...     Calc.set_val(i, sheet, 'A' + str(i))
        ...
        >>> range_addr = Calc.get_address(sheet=sheet, range_name="A1:A10")
        >>> chart_doc = Chart2.insert_chart(
        ...     sheet=sheet,
        ...     cells_range=range_addr,
        ...     cell_name="C5",
        ... )
        >>> Calc.goto_cell(cell_name="A1", doc=doc)
        >>> _ = Chart2.set_title(chart_doc=chart_doc, title="My Chart")
        >>> Chart2
        <class 'ooodev.office.chart2.Chart2'>
        >>> Calc.save_doc(doc, "sampleChart.xlsx")
        True
        >>> Lo.close_doc(doc)
        >>> Lo.close_office()
        True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This produces a spreadsheet that looks like Figure 13-10.

.. cssclass:: diagram invert

    .. _ch01fig_timeline:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299968-9fdc7c59-b2ca-4369-bb9a-364c41c67f5a.png
        :alt: OpenOffice Timeline Image
        :figclass: align-center

        :A spreadsheet with a chart added

We’ve created a bar chart by using :py:meth:`.Calc.get_address` method to set a range to ``A1:A10``, then using :py:meth:`.Chart2.insert_chart` method to insert the chart at ``C5``.
The default insert a column chart with no row or column values and default colours.
You can create many chart types including: line charts, scatter charts, and pie charts.

.. _tute_ss_summary:

Summary
=======

Often the hard part of processing information isn’t the processing itself but simply getting the data in the right format for your program.
But once you have your spreadsheet loaded into Python, you can extract and manipulate its data much faster than you could by hand.

You can also generate spreadsheets as output from your programs.
So if colleagues need your text file or PDF of thousands of sales contacts transferred to a spreadsheet file, you won’t have to tediously copy and paste it all into spreadsheets.

Equipped with |odev| module and some programming knowledge, you’ll find processing even the biggest spreadsheets a piece of cake.

.. _tute_ss_practice_questions:

Practice Questions
==================

For the following questions, imagine you have a Workbook object in the variable wb, a Worksheet object in sheet, a Cell object in cell, a Comment object in comm, and an Image object in ``img``.

.. todo::

    Tute ss, Practice questions most all need revamped.

1. What does the ***openpyxl***.load_workbook() function return?
2. What does the wb.sheetnames workbook property contain?
3. How would you retrieve the Worksheet object for a sheet named 'Sheet1'?
4. How would you retrieve the Worksheet object for the workbook’s active sheet?
5. How would you retrieve the value in the cell C5?
6. How would you set the value in the cell C5 to "Hello"?
7. How would you retrieve the cell’s row and column as integers?
8. What do the sheet.max_column and sheet.max_row sheet properties hold, and what is the data type of these properties?
9. If you needed to get the integer index for column 'M', what function would you need to call?
10. If you needed to get the string name for column 14, what function would you need to call?
11. How can you retrieve a tuple of all the Cell objects from A1 to F1?
12. How would you save the workbook to the filename example.xlsx?
13. How do you set a formula in a cell?
14. If you want to retrieve the result of a cell’s formula instead of the cell’s formula itself, what must you do first?
15. How would you set the height of row 5 to 100?
16. How would you hide column C?
17. What is a freeze pane?
18. What five functions and methods do you have to call to create a bar chart?

.. _tute_ss_practice_projects:

Practice Projects
=================

For practice, write programs that perform the following tasks.

.. _tute_ss_multiplicaton_tbl:

Multiplication Table Maker
--------------------------

Create a program ``multiplicationTable.py`` that takes a number ``N`` from the command line and creates an ``NxN`` multiplication table in a spreadsheet.
For example, when the program is run like this:

::

    py multiplicationTable.py 6

. . . it should create a spreadsheet that looks like :numref:`tute_ss_fig_multiplication_tbl`.

..
    Figure 13-11

.. cssclass:: diagram invert

    .. _tute_ss_fig_multiplication_tbl:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299901-74f36232-747a-4803-adfa-ae6d66fab93d.png
        :alt: A multiplication table generated in a spreadsheet
        :figclass: align-center

        :A multiplication table generated in a spreadsheet

Row ``1`` and column ``A`` should be used for labels and should be in bold.

.. _tute_ss_blank_row_inserter:

Blank Row Inserter
------------------

Create a program ``blankRowInserter.py`` that takes two integers and a filename string as command line arguments.
Let’s call the first integer ``N`` and the second integer ``M``.
Starting at row ``N``, the program should insert ``M`` blank rows into the spreadsheet.
For example, when the program is run like this:

::

    python blankRowInserter.py 3 2 myProduce.xlsx

. . . the “before” and “after” spreadsheets should look like :numref:`tute_ss_fig_ex_inserted_row_3`.

..
    Figure 13-12

.. cssclass:: diagram invert

    .. _tute_ss_fig_ex_inserted_row_3:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299859-486ca40a-0bbf-46e4-add9-5fa101781563.png
        :alt: Before (left) and after (right) the two blank rows are inserted at row 3
        :figclass: align-center

        Before (left) and after (right) the two blank rows are inserted at row 3

You can write this program by reading in the contents of the spreadsheet.
Then, when writing out the new spreadsheet, use a for loop to copy the first N lines.
For the remaining lines, add M to the row number in the output spreadsheet.

.. _tute_ss_sht_cell_invert:

Spreadsheet Cell Inverter
-------------------------

Write a program to invert the row and column of the cells in the spreadsheet.
For example, the value at row ``5``, column ``3`` will be at row ``3``, column ``5`` (and vice versa).
This should be done for all cells in the spreadsheet.
For example, the “before” and “after” spreadsheets would look something like :numref:`tute_ss_fig_sht_before_after_top_btm`.

..
    Figure 13-13

.. cssclass:: diagram invert

    .. _tute_ss_fig_sht_before_after_top_btm:
    .. figure:: https://user-images.githubusercontent.com/11288701/208299872-1d3fec93-a74f-4660-a6af-fde3ad9ae33d.png
        :alt: The spreadsheet before (top) and after (bottom) inversion
        :figclass: align-center

        The spreadsheet before (top) and after (bottom) inversion

You can write this program by using nested for loops to read the spreadsheet’s data into a list of lists data structure.
This data structure could have ``sheet_data[x][y]`` for the cell at column x and row y.
Then, when writing out the new spreadsheet, use ``sheet_data[y][x]`` for the cell at column ``x`` and row ``y``.

.. _text_ss_text_file_sht:

Text Files to Spreadsheet
-------------------------

Write a program to read in the contents of several text files (you can make the text files yourself) and insert those contents into a spreadsheet, with one line of text per row.
The lines of the first text file will be in the cells of column ``A``, the lines of the second text file will be in the cells of column ``B``, and so on.

Use the ``readlines()`` File object method to return a list of strings, one string per line in the file.
For the first file, output the first line to column ``1``, row ``1``.
The second line should be written to column ``1``, row ``2``, and so on.
The next file that is read with ``readlines()`` will be written to column ``2``, the next file to column 3``, and so on.

.. _tute_ss_sht_to_txt_file:

Spreadsheet to Text Files
-------------------------

Write a program that performs the tasks of the previous program in reverse order: the program should open a spreadsheet and write the cells of column ``A`` into one text file,
the cells of column B into another text file, and so on.

.. _XCell: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCell.html
