.. _ch20:

***********************************************
Chapter 20. Spreadsheet Displaying and Creation
***********************************************

.. topic:: Overview

    Displaying a Document; Read-only and Protected Viewing; Active Sheets; Sheet Names; Zooming; Creating a Document;
    Cell Names and Ranges; Cell Values; Data Arrays; Rows and Columns of Data;  Adding a Picture and a Chart

    Examples: |build_tbl|_ and |show_sheet|_.


This chapter looks at two main topics: the display of an existing spreadsheet document, and the creation of a new document, based around two examples.

As part of displaying a document, we'll look at describing how to use read-only and protected viewing, change the active sheet, use sheet names, and adjust the window view size by zooming.

Document creation illustrates the use of cell names and ranges, the setting of cell data using arrays, rows, and columns, and adding a picture and a chart to a sheet.

.. _ch20_display_doc:

20.1 Displaying a Document
==========================

The |show_sheet_py|_ example shows how to open a spreadsheet document, and display its first sheet.
If the program is called with a filename argument, then the document is saved to that file before the program closes.
The extension of the output filename is used to determine the exported file type.
For example:

::

    python start.py --show --file "totals.ods" --out "tmp/totals.pdf"

displays the ``totals.ods`` spreadsheet, and saves it as a PDF file. Alternatively:

::

    python start.py --show --file "sorted.csv" --out "totals.html"

opens the CSV file as a Calc document, and saves it as HTML.

The ``main()`` function of |show_sheet_py|_:

.. tabs::

    .. code-tab:: python

        # ShowSheet.main() function of show_sheet.py
        def main(self) -> None:
            loader = Lo.load_office(Lo.ConnectSocket())

            try:
                doc = Calc.open_doc(fnm=self._input_fnm, loader=loader)

                # doc = Lo.open_readonly_doc(fnm=self._input_fnm, loader=loader)
                # doc = Calc.get_ss_doc(doc)

                if self._visible:
                    GUI.set_visible(is_visible=True, odoc=doc)

                Calc.goto_cell(cell_name="A1", doc=doc)
                sheet_names = Calc.get_sheet_names(doc=doc)
                print(f"Names of Sheets ({len(sheet_names)}):")
                for name in sheet_names:
                    print(f"  {name}")

                sheet = Calc.get_sheet(doc=doc, index=0)
                Calc.set_active_sheet(doc=doc, sheet=sheet)
                pro = Lo.qi(XProtectable, sheet, True)
                pro.protect("foobar")
                print(f"Is protected: {pro.isProtected()}")

                Lo.delay(2000)
                # query the user for the password
                pwd = GUI.get_password("Password", "Enter sheet Password")
                if pwd == "foobar":
                    pro.unprotect(pwd)
                    MsgBox.msgbox("Password is Correct", "Password", boxtype=MessageBoxType.INFOBOX)
                else:
                    MsgBox.msgbox("Password is incorrect", "Password", boxtype=MessageBoxType.ERRORBOX)

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

:py:meth:`.Calc.open_doc` opens the document, returning an XSpreadsheetDocument_ reference:

.. tabs::

    .. code-tab:: python

        # in Calc class (simplified)
        @classmethod
        def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> XSpreadsheetDocument:
            doc = Lo.open_doc(fnm=fnm, loader=loader)
            if doc is None:
                raise Exception("Document is null")
            return cls.get_ss_doc(doc)

        @staticmethod
        def get_ss_doc(doc: XComponent) -> XSpreadsheetDocument:
            if not Info.is_doc_type(doc_type=mLo.Lo.Service.CALC, obj=doc):
                if not Lo.is_macro_mode:
                    Lo.close_doc(doc=doc)
                raise Exception("Not a spreadsheet doc")

            ss_doc = Lo.qi(XSpreadsheetDocument, doc)
            if ss_doc is None:
                if not Lo.is_macro_mode:
                    Lo.close_doc(doc=doc)
                raise MissingInterfaceError(XSpreadsheetDocument)
            return ss_doc

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        - :odev_src_calc_meth:`open_doc`
        - :odev_src_calc_meth:`get_ss_doc`

First :py:meth:`.Lo.open_doc` returns an XComponent_ reference, then :py:meth:`~.Calc.get_ss_doc` converts it to XSpreadsheetDocument_.
This conversion will fail if the input document isn't a spreadsheet.

``GUI.set_visible(is_visible=True, odoc=doc)`` causes Office to display the spreadsheet's active sheet, which is the one that was being worked on when the file was previously saved.
In addition, the application will display the cell or cells selected in the sheet at that time. The selection can be changed by calling :py:meth:`.Calc.goto_cell`:

.. tabs::

    .. code-tab:: python

        # in the Calc class
        @staticmethod
        def get_controller(doc: XSpreadsheetDocument) -> XController:
            model = Lo.qi(XModel, doc, True)
            return model.getCurrentController()

        # overload method, simplified
        @classmethod
        def goto_cell(cls, cell_name: str, doc: XSpreadsheetDocument) -> None:
            frame = cls.get_controller(doc).getFrame()
            cls.goto_cell(cell_name=cell_name, frame=frame)
    
        # overload method, simplified
        @classmethod
        def goto_cell(cls, cell_name: str, frame: XFrame) -> None:
            props = Props.make_props(ToPoint=kargs[1])
            Lo.dispatch_cmd(cmd="GoToCell", props=props, frame=frame)
    
    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        - :odev_src_calc_meth:`get_controller`
        - :odev_src_calc_meth:`goto_cell`

Any changes to the document's view requires a reference to its controller.
The active cell is changed by sending a ``GoToCell`` dispatch to the controller's frame.
``GoToCell`` requires a cell name argument, which is encoded as a property array containing a single ``ToPoint`` property.

:py:meth:`.Calc.get_sheet` returns a reference to the currently active sheet.
:py:meth:`~.Calc.get_sheet` is implemented using code similar to that described in the previous chapter:

.. tabs::

    .. code-tab:: python

        # in Calc class (overload method, simplified)
        @staticmethod
        def get_sheet(doc: XSpreadsheetDocument, index: int) -> XSpreadsheet:
            try:
                sheets = doc.getSheets()
                xsheets_idx = Lo.qi(XIndexAccess, sheets, True)
                sheet = Lo.qi(XSpreadsheet, xsheets_idx.getByIndex(index), raise_err=True)
                return sheet
            except Exception as e:
                raise Exception(f"Could not access spreadsheet: {index}") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`get_sheet`

If the user calls |show_sheet_py|_ with a ``--out`` argument, then :py:meth:`.Lo.save_doc` performs a save to that file.
:py:meth:`~.Lo.save_doc` utilizes :py:meth:`.Lo.ext_to_format` to map the filename's extension (:abbreviation:`i.e.` ``pdf``, ``html``, ``xhtml``)
and the document type (in this case, a spreadsheet) to a suitable output format.
The function consists of a long else if statement which decides on the most suitable mapping, as illustrated by the code fragment:

:py:class:`~.lo.Lo.DocType` is an enum that provides the type of document.

.. tabs::

    .. code-tab:: python

        # in Lo class
        @classmethod
        def ext_to_format(cls, ext: str, doc_type: Lo.DocType = DocType.UNKNOWN) -> str:
            dtype = cls.DocType(doc_type)
            s = ext.lower()
            if s == "doc":
                return "MS Word 97"
            elif s == "docx":
                return "Office Open XML Text"  # MS Word 2007 XML
            elif s == "rtf":
                if dtype == cls.DocType.CALC:
                    return "Rich Text Format (StarCalc)"
                else:
                    return "Rich Text Format"
            elif s == "odt":
                return "writer8"
            elif s == "ott":
                return "writer8_template"
            elif s == "pdf":
                if dtype == cls.DocType.WRITER:
                    return "writer_pdf_Export"
                elif dtype == cls.DocType.IMPRESS:
                    return "impress_pdf_Export"
                elif dtype == cls.DocType.DRAW:
                    return "draw_pdf_Export"
                elif dtype == cls.DocType.CALC:
                    return "calc_pdf_Export"
                elif dtype == cls.DocType.MATH:
                    return "math_pdf_Export"
                else:
                    return "writer_pdf_Export"  # assume we are saving a writer doc
            
            # and many more cases ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The ``pdf`` case is selected when the output file extension is ``pdf``, but the export format also depends on the Office document.
For |show_sheet_py|_, the ``doc_type`` value will be :py:attr:`.Lo.DocType.CALC`, which causes :py:meth:`~.Lo.ext_to_format` to return ``calc_pdf_Export``.

:py:meth:`.Lo.ext_to_format` is very far from comprehensive, but understands Office and MS spreadsheet formats, ``CSV``, ``RTF``, ``text``, ``HTML``, ``XHTML``, and ``PDF``.
Other mappings can be added as required.

.. _ch20_read_only_protect_view:

20.1.1 Read-only and Protected Viewing
--------------------------------------

One variant of |show_sheet_py|_ prevents a user from changing the spreadsheet. 
Unfortunately, implementing this by opening the document read-only isn't particularly robust because
Office always displays a message asking if the user wants to override the read-only setting.
Nevertheless, the code is short:

.. tabs::

    .. code-tab:: python

        # Commeted out in show_sheet.py
        doc = Lo.open_readonly_doc(fnm=self._input_fnm, loader=loader)
        doc = Calc.get_ss_doc(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


:py:meth:`.open_readonly_doc` calls :py:meth:`.Lo.open_doc` with the ``ReadOnly`` property set to ``True``:

.. tabs::

    .. code-tab:: python

        # in the Lo Class
        @classmethod
        def open_readonly_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> XComponent:
            return cls.open_doc(fnm, loader, Props.make_props(Hidden=True, ReadOnly=True))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

If you want to actually stop the user from changing the spreadsheet, then it must be protected, using the XProtectable_ interface:

.. tabs::

    .. code-tab:: python

        # in ShoWSheet.main() of show_sheet.py
        pro = Lo.qi(XProtectable, sheet, True)
        pro.protect("foobar")


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``XProtectable.protect()`` assigns a password to the sheet (in this example, ``foobar``), which the user must supply in order to change any data.

Document-level protection isn't supported in the current version of Office.
The best we can do is to apply protection to individual sheets. Namely:

.. tabs::

    .. code-tab:: python

        # 
        pro = Lo.qi(XProtectable, sheet, True)
        pro.protect("foobar")

        # query the user for the password
        pwd = GUI.get_password("Password", "Enter sheet Password")
        if pwd == "foobar":
            pro.unprotect(pwd)
            MsgBox.msgbox("Password is Correct", "Password", boxtype=MessageBoxType.INFOBOX)
        else:
            MsgBox.msgbox("Password is incorrect", "Password", boxtype=MessageBoxType.ERRORBOX)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The code fragment above shows how to query the user for the password. :py:meth:`.GUI.get_password` displays a dialog window which includes a Password Field:

As a fallback :py:meth:`.GUI.get_password` will attempt to build a dialog using ``tkinter`` if for any reason Office dialog cannot be built. 

.. tabs::

    .. code-tab:: python

        # in GUI class
        @staticmethod
        def get_password(title: str, input_msg: str) -> str:
            try:
                result = Input.get_input(title=title, msg=input_msg, is_password=True)
                return result
            except Exception:
                # may not be in a LibreOffice window
                pass

            # try a tkinter dialog. Not available in macro mode.
            # this also means may not work on windows when virtual environment
            # is set to LibreOffice python.exe
            try:
                from ..dialog.tk_input import Window

                pass_inst = Window(title=title, input_msg=input_msg, is_password=True)
                return pass_inst.get_input()
            except ImportError:
                pass
            raise Exception("Unable to access a GUI to create a password dialog box")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. seealso::

    .. cssclass:: ul-list

        - :ref:`class_msg_box`
        - :ref:`class_dialog_input`
        - :ref:`dialog_tk_input`

.. _ch20_change_active_sheet:

20.1.2 Changing the Active Sheet
--------------------------------

Another variation of |show_sheet_py|_ allows the user to specify which sheet to make active, and so be displayed in Office.
It's not enough to execute :py:meth:`.Calc.get_sheet` with a sheet index; :py:meth:`.Calc.set_active_sheet` must also be called to make that sheet active:

.. tabs::

    .. code-tab:: python

        # in the Calc class (simplified)
        @classmethod
        def set_active_sheet(cls, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
            ss_view = cls.get_view(doc)
            if ss_view is None:
                return
            ss_view.setActiveSheet(sheet)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`set_active_sheet`

:py:meth:`.Calc.set_active_sheet` converts the controller interface for the document into an XSpreadsheetView_,
which is the main interface for the SpreadsheetView_ service (see :numref:`ch20fig_spreadsheetview_interfaces`).

..
    figure 1

.. cssclass:: diagram invert

    .. _ch20fig_spreadsheetview_interfaces:
    .. figure:: https://user-images.githubusercontent.com/4193389/202597547-984bd253-57ff-4096-a2d3-4b78ae35cb23.png
        :alt: The Spread sheet View Services and Interfaces.
        :figclass: align-center

        :The SpreadsheetView_ Services and Interfaces.

SpreadsheetView_ implements a number of interfaces for affecting the user's view of a document.
View-related properties are stored in the inherited SpreadsheetViewSettings_ class, which can be listed by calling :py:meth:`.Props.show_obj_props`:

.. _ch20_sheet_names:

20.1.3 Sheet Names
------------------

Default spreadsheet names use ``Sheet`` followed by a positive integer.
For example, a new document will name its first sheet ``Sheet1``.
:py:meth:`.Calc.get_sheet` can accept a sheet name, as in the following code which makes ``Sheet1`` active:

.. tabs::

    .. code-tab:: python

        sheet = Calc.get_sheet(doc=doc, sheet_name="Sheet1")
        Calc.set_active_sheet(doc=doc, sheet=sheet)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

All the sheet names in a document can be accessed via :py:meth:`.Calc.get_sheet_names`, and a sheet's name can be changed by :py:meth:`.Calc.set_sheet_name`:

.. tabs::

    .. code-tab:: python

        # in the Calc class
        @staticmethod
        def get_sheet_names(doc: XSpreadsheetDocument) -> Tuple[str, ...]:
            sheets = doc.getSheets()
            return sheets.getElementNames()

        @staticmethod
        def set_sheet_name(sheet: XSpreadsheet, name: str) -> bool:
            xnamed = Lo.qi(XNamed, sheet)
            if xnamed is None:
                Lo.print("Could not access spreadsheet")
                return False
            xnamed.setName(name)
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch20_zooming:

20.1.4 Zooming
--------------

Zooming the view of a document is done by adjusting properties in SpreadsheetViewSettings_ (see :numref:`ch20fig_spreadsheetview_interfaces`).
The ``ZoomType`` property specifies the kind of zooming, which may be a size percentage or a constant indicating a particular zooming amount.
The constants are defined in :py:class:`GUI.ZoomEnum`:

The constants are understood by :py:meth:`.Calc.zoom`:

.. tabs::

    .. code-tab:: python

        # in the Calc class
        @classmethod
        def zoom(cls, doc: XSpreadsheetDocument, type: GUI.ZoomEnum) -> None:

            ctrl = cls.get_controller(doc)
            if ctrl is None:
                return

            def zoom_val(value: int) -> None:
                Props.set(ctrl, ZoomType=GUI.ZoomEnum.BY_VALUE.value, ZoomValue=value)

            if (
                type == GUI.ZoomEnum.ENTIRE_PAGE
                or type == GUI.ZoomEnum.OPTIMAL
                or type == GUI.ZoomEnum.PAGE_WIDTH
                or type == GUI.ZoomEnum.PAGE_WIDTH_EXACT
            ):
                Props.set(ctrl, ZoomType=type.value)
            elif type == GUI.ZoomEnum.ZOOM_200_PERCENT:
                zoom_val(200)
            elif type == GUI.ZoomEnum.ZOOM_150_PERCENT:
                zoom_val(150)
            elif type == GUI.ZoomEnum.ZOOM_100_PERCENT:
                zoom_val(100)
            elif type == GUI.ZoomEnum.ZOOM_75_PERCENT:
                zoom_val(75)
            elif type == GUI.ZoomEnum.ZOOM_50_PERCENT:
                zoom_val(50)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


For example, the entire sheet can be made visible by calling:

.. tabs::

    .. code-tab:: python

        Calc.Zoom(doc=doc, type=GUI.ZoomEnum.ENTIRE_PAGE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

For percentage zooming, the value must be assigned to the ``ZoomValue`` property.
This is handled by :py:meth:`.Calc.zoom_value`:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @classmethod
        def zoom_value(cls, doc: XSpreadsheetDocument, value: int) -> None:
            ctrl = cls.get_controller(doc)
            if ctrl is None:
                return
            Props.set(ctrl, ZoomType=GUI.ZoomEnum.BY_VALUE.value, ZoomValue=value)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch20_creating_doc:

20.2 Creating a Document
========================

The |build_tbl|_ example shows how to create a new spreadsheet document, populate it with data, apply cell styling, and save it to a file.
We'll look at styling in the next chapter, and will concentrate here on the different ways to add data to a sheet.

The ``main()`` method of |build_tbl_py|_ is:

.. tabs::

    .. code-tab:: python

        # BuildTable.main() of build_table.py
        def main(self) -> None:
            loader = Lo.load_office(Lo.ConnectSocket())

            try:
                doc = Calc.create_doc(loader)

                GUI.set_visible(is_visible=True, odoc=doc)

                sheet = Calc.get_sheet(doc=doc, index=0)

                self._convert_addresses(sheet)

                # other possible build methods
                # self._buld_cells(sheet)
                # self._build_rows(sheet)
                # self._build_cols(sheet)

                self._build_array(sheet)

                if self._add_pic:
                    self._add_picture(sheet=sheet, doc=doc)

                # add a chart
                if self._add_chart and Chart2:
                    # assumes _build_array() has filled the spreadsheet with data
                    chart_cell = "B6" if self._add_pic else "D6"
                    rng_addr = Calc.get_address(sheet=sheet, range_name="B2:M4")
                    Chart2.insert_chart(
                        sheet=sheet, cells_range=rng_addr, cell_name=chart_cell,
                        width=21, height=11, diagram_name="Column"
                    )

                if self._add_style:
                    self._create_styles(doc)
                    self._apply_styles(sheet)

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


``main()`` can call one of four different build methods to demonstrate various :py:class:`~calc.Calc` methods for filling cells and cell ranges.
``_convert_addresses()`` illustrates the :py:class:`~calc.Calc` methods for converting between cell names and positions, and between cell range names and position intervals.

.. _ch20_switch_name_range_pos:

20.2.1 Switching between Cell Names, Cell Ranges, and Positions
---------------------------------------------------------------

Although the :py:class:`~calc.Calc` get/set methods for cells, columns, rows, and cell ranges support both name and position based addressing (:abbreviation:`i.e.` ``D5`` and (``3``, ``4``)),
it's still sometimes necessary to convert between the different formats. ``_convert_addresses()`` demonstrates those methods:

.. tabs::

    .. code-tab:: python

        # in build_table.py
        def _convert_addresses(self, sheet: XSpreadsheet) -> None:
            # cell name <--> position
            pos = Calc.get_cell_position(cell_name="AA2")
            print(f"Positon of AA2: ({pos.X}, {pos.Y})")

            cell = Calc.get_cell(sheet=sheet, col=pos.X, row=pos.Y)
            Calc.print_cell_address(cell)

            print(f"AA2: {Calc.get_cell_str(col=pos.X, row=pos.Y)}")
            print()

            # cell range name <--> position
            rng = Calc.get_cell_range_positions("A1:D5")
            print(f"Range of A1:D5: ({rng[0].X}, {rng[0].Y}) -- ({rng[1].X}, {rng[1].Y})")

            cell_rng = Calc.get_cell_range(
                sheet=sheet, col_start=rng[0].X, row_start=rng[0].Y, col_end=rng[1].X, row_end=rng[1].Y
            )
            Calc.print_address(cell_rng)
            print(
                "A1:D5: " + Calc.get_range_str(
                                col_start=rng[0].X, row_start=rng[0].Y, col_end=rng[1].X, row_end=rng[1].Y
                            )
            )
            print()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


``_convert_addresses()`` prints the following:

::

    Positon of AA2: (26, 1)
    Cell: Sheet1.AA2
    AA2: AA2

    Range of A1:D5: (0, 0) -- (3, 4)
    Range: Sheet1.A1:D5
    A1:D5: A1:D5

.. _ch20_name_manipulation:

Cell Name Manipulation
^^^^^^^^^^^^^^^^^^^^^^

:py:meth:`.Calc.get_cell_position` converts a cell name, such as ``AA2``, into a (column, row) position coordinate, which it returns as a Point_ object.
For ``AA2`` the result is ``(26, 1)``, since the column labeled ``AA`` follows ``Z`` in a spreadsheet.
The implementation uses regular expression parsing of the input string to separate out the alphabetic and numerical parts before processing them:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @classmethod
        def get_cell_position(cls, cell_name: str) -> Point:
            #  _rx_cell = re.compile(r"([a-zA-Z]+)([0-9]+)")
            m = cls._rx_cell.match(cell_name)
            if m:
                ncolumn = cls.column_string_to_number(str(m.group(1)).upper())
                nrow = cls.row_string_to_number(m.group(2))
                return Point(ncolumn, nrow)
            else:
                raise ValueError("Not a valid cell name")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


:py:meth:`.Calc.get_cell` converts a position into an XCell_ reference to the cell:

.. tabs::

    .. code-tab:: python

        cell = Calc.get_cell(sheet=sheet, col=26, row=1);

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The function is a wrapper around ``XCellRange.getCellByPosition()``:

.. tabs::

    .. code-tab:: python

        # in Calc class (overloads method, simplified)
        @classmethod
        def get_cell(cls, sheet: XSpreadsheet, col: int, row: int) -> XCell:
            return sheet.getCellByPosition(col, row)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


A second overload of :py:meth:`~.Calc.get_cell` refers to a cell by name:

.. tabs::

    .. code-tab:: python

        cell = Calc.get_cell(sheet=sheet, cell_name="AA2");

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The code:

.. tabs::

    .. code-tab:: python

        # in Calc class (overloads method, simplified)
        @classmethod
        def get_cell(cls, sheet: XSpreadsheet, cell_name: str) -> XCell:
            cell_range = sheet.getCellRangeByName(cell_name)
            return cls.get_cell(cell_range=cell_range, col=0, row=0)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`get_cell`

The call to ``XCellRange.getCellRangeByName()`` with a single cell name returns a cell range made up of one cell.
This XCellRange_ reference can be passed to :py:meth:`.Calc.get_cell` since XCellRange_ is a superclass of XSpreadsheet_.
The ``get_cell(sheet: XSpreadsheet, col: int, row: int)`` overload of :py:meth:`~.Calc.get_cell` is called,
and ``XCellRange.getCellByPosition()`` treats ``(0, 0)`` as a position relative to the cell range.
There's only one cell in this range, so ``getCellByPosition()`` returns a reference to the ``AA2`` cell.

.. _ch20_range_manipulation:

Cell Range Manipulation
^^^^^^^^^^^^^^^^^^^^^^^

The second half of ``_convert_addresses()`` shows off some of the cell range addressing methods.
:py:meth:`.Calc.get_cell_range_positions` returns a tuple of two Point_ objects corresponding to the top-left and bottom-right cells of the range:

.. tabs::

    .. code-tab:: python

        # in BuildTable._convert_addresses() of build_table.py
        pos = Calc.get_cell_position(cell_name="AA2")
        print(f"Positon of AA2: ({pos.X}, {pos.Y})")
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Only simple cell range names of the form ``:`` are understood by :py:class:`~.calc.Calc` methods.
Range operators, such as ``~``, ``!``, and absolute references using ``$`` are **not** supported.

:py:meth:`.Calc.get_cell_range` converts a range address into an XCellRange_ reference:


.. tabs::

    .. code-tab:: python

        cell = Calc.get_cell_range(sheet=sheet, range_name="A1:D5");

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This method wraps a call to ``XCellRange.getCellRangeByName()``:

.. tabs::

    .. code-tab:: python

        # in Calc class (overload method, simplified)
        @staticmethod
        def get_cell_range(sheet: XSpreadsheet, range_name: str) -> XCellRange:
            cell_range = sheet.getCellRangeByName(range_name)
            if cell_range is None:
                raise Exception(f"Could not access cell range: {range_name}")
            return cell_range

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`get_cell_range`

.. _ch20_changing_cell_values:

20.2.2 Changing Cell Values
---------------------------

Back in |build_tbl_py|_, the ``_build_cells()`` methods shows how individual cells can be assigned values.
The code uses two versions of :py:meth:`.Calc.set_val`, one that accepts a cell position, the other a cell name.
For example:

.. tabs::

    .. code-tab:: python

        # in _build_cells() of build_table.py
        # ...
        for i, val in enumerate(header_vals):
            # set by name
            Calc.set_val(value=val, sheet=sheet, col=i + 1, row=0)

        # ...
        for i, val in enumerate(vals):
            # set by row, column
            cell_name = TableHelper.make_cell_name(row=2, col=i + 2)
            Calc.set_val(value=val, sheet=sheet, cell_name=cell_name)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Both methods store a number or a string in a cell, by processing the input value as an Object:

.. tabs::

    .. code-tab:: python

        # in Calc class (overload methods, simpilified)
        @classmethod
        def set_val(cls, value: object, sheet: XSpreadsheet, cell_name: str) -> None:
            pos = cls.get_cell_position(cell_name)
            cls.set_val(value=value, sheet=sheet, col=pos.X, row=pos.Y)

        @classmethod
        def set_val(cls, value: object, sheet: XSpreadsheet, col: int, row: int) -> None:
            cell = cls.get_cell(sheet=sheet, col=col, row=row)
            cls.set_val(value=value, cell=cell)

        @classmethod
        def set_val(cls, value: object, cell: XCell) -> None:
            if isinstance(value, numbers.Number):
                cell.setValue(float(value))
            elif isinstance(value, str):
                cell.setFormula(str(value))
            else:
                Lo.print(f"Value is not a number or string: {value}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`set_val`

The ``set_val(cls, value: object, cell: XCell) -> None`` method examines the type of the value to decide whether to call ``XCell.setValue()`` or ``XCell.setFormula()``.

.. _ch20_storing_2d_arrays:

20.2.3 Storing 2D Arrays of Data
--------------------------------

The ``_build_array()`` method in |build_tbl_py|_ shows how a block of data can be stored by :py:meth:`.Calc.set_array`:

.. tabs::

    .. code-tab:: python

        # in build_table.py
        def _build_array(self, sheet: XSpreadsheet) -> None:
            vals = (
                ("", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"),
                ("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3, -67.3, 30.5, 23.2, -97.3, 22.4, 23.5),
                ("Jones", 21, 40.9, -57.5, -23.4, 34.5, 59.3, 27.3, -38.5, 43.2, 57.3, 25.4, 28.5),
                ("Brown", 31.45, -20.9, -117.5, 23.4, -114.5, 115.3, -171.3, 89.5, 41.2, 71.3, 25.4, 38.5),
            )
            Calc.set_array(values=vals, sheet=sheet, name="A1:M4")  # or just A1

            Calc.set_val(sheet=sheet, cell_name="N1", value="SUM")
            Calc.set_val(sheet=sheet, cell_name="N2", value="=SUM(B2:M2)")
            Calc.set_val(sheet=sheet, cell_name="N3", value="=SUM(B3:M3)")
            Calc.set_val(sheet=sheet, cell_name="N4", value="=SUM(B4:M4)")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.set_array` accepts a 2D array of Object values (which means it can contain a mix of strings and doubles) with the data arranged in row-order.
For example, the data shown above is stored in the sheet as in :numref:`ch20fig_bt_block_data`.

..
    figure 2

.. cssclass:: screen_shot invert

    .. _ch20fig_bt_block_data:
    .. figure:: https://user-images.githubusercontent.com/4193389/202787908-45294533-f8be-444f-b7bb-e25f087fe622.png
        :alt: A Block of Data Added to a Sheet
        :figclass: align-center

        :A Block of Data Added to a Sheet.

The second argument of :py:meth:`.Calc.set_array` can be a cell range or a single cell name representing the top-left corner of the range.
In the latter case, the cell range is calculated from the size of the array.
This means that the call used above could be rewritten as:

.. tabs::

    .. code-tab:: python

        # in BuildTable._build_array() of build_table.py
        Calc.set_array(values=vals, sheet=sheet, name="A1:M4")  # or just A1

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.set_array` is defined as:

.. tabs::

    .. code-tab:: python

        # in Calc class (overload methid, simpilified)
        @classmethod
        def set_array(cls, values: Table, sheet: XSpreadsheet, name: str) -> None:
                # set_array(values: Sequence[Sequence[object]], sheet: XSpreadsheet, name: str)
            if cls.is_cell_range_name(name):
                cls.set_array_range(sheet=sheet, range_name=name, values=values)
            else:
                cls.set_array_cell(sheet=sheet, cell_name=name, values=values)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`set_array`

.. tabs::

    .. code-tab:: python

        # in Calc class (simplified)
        @classmethod
        def set_array_range(cls, sheet: XSpreadsheet, range_name: str, values: Table) -> None:
            v_len = len(values)
            if v_len == 0:
                Lo.print("Values has not data")
                return
            cell_range = cls.get_cell_range(sheet=sheet, range_name=range_name)
            cls.set_cell_range_array(cell_range=cell_range, values=values)

        @classmethod
        def set_array_cell(cls, sheet: XSpreadsheet, cell_name: str, values: Table) -> None:
            v_len = len(values)
            if v_len == 0:
                Lo.print("Values has not data")
                return
            pos = cls.get_cell_position(cell_name)
            col_end = pos.X + (len(values[0]) - 1)
            row_end = pos.Y + (v_len - 1)
            cell_range = cls._get_cell_range_col_row(
                sheet=sheet, start_col=pos.X, start_row=pos.Y, end_col=col_end, end_row=row_end
            )
            cls.set_cell_range_array(cell_range=cell_range, values=values)

        @staticmethod
        def set_cell_range_array(cell_range: XCellRange, values: Table) -> None:
            v_len = len(values)
            if v_len == 0:
                Lo.print("Values has not data")
                return
            cr_data = mLo.Lo.qi(XCellRangeData, cell_range)
            if cr_data is None:
                return
            cr_data.setDataArray(values)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. seealso::

    .. cssclass:: src-link

        - :odev_src_calc_meth:`set_array_range`
        - :odev_src_calc_meth:`set_array_cell`
        - :odev_src_calc_meth:`set_cell_range_array`

The storage of the array is performed by :py:meth:`.Calc.set_cell_range_array` which is passed an XCellRange_ object and a 2D array.
XCellRange_ is converted into XCellRangeData_ which has a ``setDataArray()`` method.

.. _ch20_storing_rows_data:

20.2.4 Storing Rows of Data
---------------------------

:py:meth:`.Calc.set_row` lets the programmer store a 1D array as a row of data:

.. tabs::

    .. code-tab:: python

        # in BuildTable._build_rows() of build_table.py
        vals = (42, 58.9, -66.5, 43.4, 44.5, 45.3, -67.3, 30.5, 23.2, -97.3, 22.4, 23.5)
        Calc.set_row(sheet=sheet, values=vals, cell_name="B2")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.set_row` employs ``XCellRangeData.setDataArray()``, which requires an XCellRange_ object and a 2D array:

.. tabs::

    .. code-tab:: python

        # in Calc class (overload method, simpilified)
        @classmethod
        def set_row(cls, sheet: XSpreadsheet, values: Row, cell_name: str) -> None:
            pos = cls.get_cell_position(cell_name)
            # column row
            cls.set_row(sheet=sheet, values=values, col_start=pos.X, ros_start=pos.Y)

        @classmethod
        def set_row(cls, sheet: XSpreadsheet, values: Row, col_start: int, row_start: int) -> None:
            try:
                cell_range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
                if cell_range is None:
                    raise Exception
                return cell_range
            except Exception as e:
                raise Exception(
                    f"Could not access cell range : ({start_col}, {start_row}, {end_col}, {end_row})"
                ) from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`set_row`

.. _ch20_storing_col_data:

20.2.5 Storing Columns of Data
------------------------------

:py:meth:`.Calc.set_col` lets the programmer store a column of data, as shown in |build_tbl_py|_ in its ``_build_cols()`` method:

.. tabs::

    .. code-tab:: python

        # in BuildTable._build_cols() of build_table.py
        def _build_cols(self, sheet: XSpreadsheet) -> None:
            vals = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
            Calc.set_col(sheet=sheet, values=vals, cell_name="A2")
            Calc.set_val(value="SUM", sheet=sheet, cell_name="A14")

            Calc.set_val(value="Smith", sheet=sheet, cell_name="B1")
            vals = (42, 58.9, -66.5, 43.4, 44.5, 45.3, -67.3, 30.5, 23.2, -97.3, 22.4, 23.5)
            Calc.set_col(sheet=sheet, values=vals, cell_name="B2")
            Calc.set_val(value="=SUM(B2:M2)", sheet=sheet, cell_name="B14")

            Calc.set_val(value="Jones", sheet=sheet, col=2, row=0)
            vals = (21, 40.9, -57.5, -23.4, 34.5, 59.3, 27.3, -38.5, 43.2, 57.3, 25.4, 28.5)
            Calc.set_col(sheet=sheet, values=vals, col_start=2, row_start=1)
            Calc.set_val(value="=SUM(B3:M3)", sheet=sheet, col=2, row=13)

            Calc.set_val(value="Brown", sheet=sheet, col=3, row=0)
            vals = (31.45, -20.9, -117.5, 23.4, -114.5, 115.3, -171.3, 89.5, 41.2, 71.3, 25.4, 38.5)
            Calc.set_col(sheet=sheet, values=vals, col_start=3, row_start=1)
            Calc.set_val(value="=SUM(A4:L4)", sheet=sheet, col=3, row=13)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``_build_cols()`` creates the spreadsheet shown in :numref:`ch20fig_bt_col_data`.

..
    figure 3

.. cssclass:: screen_shot invert

    .. _ch20fig_bt_col_data:
    .. figure:: https://user-images.githubusercontent.com/4193389/202793984-770d3e98-50a8-4613-b964-34951ab2aaeb.png
        :alt: Columns of Data in a Sheet
        :figclass: align-center

        :Columns of Data in a Sheet.

Column creation is a little harder than row building since it's not possible to use ``XCellRangeData.setDataArray()`` which assumes that data is row-ordered.
Instead :py:meth:`.Calc.set_col` calls :py:meth:`.Calc.set_val` in a loop:

.. tabs::

    .. code-tab:: python

        # in Calc class (overload method, simplified)
        @classmethod
        def set_col(cls, sheet: XSpreadsheet, values: Column, cell_name: str) -> None:
            pos = cls.get_cell_position(cell_name)
            cls.set_col(sheet=sheet, value=values, col_start=pos.X, row_start=pos.Y)

        @classmethod
        def set_col(cls, sheet: XSpreadsheet, values: Column, col_start: int, row_start: int) -> None:
            cell_range = cls.get_cell_range(
                sheet=sheet, col_start=col_start, row_start=y, col_end=x, row_end=y + val_len - 1
            )
            xcell: XCell = None
            for val in range(val_len):
                xcell = cls.get_cell(cell_range=cell_range, col=0, row=val)
                cls.set_val(cell=xcell, value=values[val])


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        :odev_src_calc_meth:`set_col`

.. _ch20_adding_pic:

20.2.6 Adding a Picture
-----------------------

Adding an image to a spreadsheet is straightforward since every sheet is also a draw page.
The Spreadsheet_ service has an XDrawPageSupplier_ interface, which has a ``getDrawPage()`` method.
The returned XDrawPage_ reference points to a transparent drawing surface that lies over the top of the sheet.

Adding a picture is done by calling :py:meth:`.Draw.draw_image`:

.. tabs::

    .. code-tab:: python

        # in BuildTable._add_picture() of build_table.py
        # ...
        dp_sup = Lo.qi(XDrawPageSupplier, sheet, True)
        page = dp_sup.getDrawPage()
        x = 230 if self._add_chart else 125
        Draw.draw_image(slide=page, fnm=self._im_fnm, x=x, y=32)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The ``(125, 32)`` or ``(230, 32)`` passed to :py:meth:`.Draw.draw_image` is the ``(x, y)`` offset from the top-left corner of the sheet, specified in millimeters.
This method comes from my :py:class:`~.draw.Draw` class, explained in :ref:`part03`.

.. _ch20_draw_warn:

Warning when Drawing
^^^^^^^^^^^^^^^^^^^^

Many of the :py:class:`~.draw.Draw` methods take a document argument, such as :py:meth:`.Draw.get_slides_count` which returns the number of draw pages in the document:

.. tabs::

    .. code-tab:: python

        print(f'No of draw pages: {Draw.get_slides_count(doc)}')

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

These methods assume that the document argument can be cast to XComponent_.
For instance, the function prototype for :py:meth:`.Draw.get_slides_count` is:

.. tabs::

    .. code-tab:: python

        def get_slides_count(cls, doc: XComponent) -> int:
            ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Unfortunately, casting via :py:meth:`.Lo.qi` will not work with spreadsheet documents because XSpreadsheetDocument_ doesn't inherit XComponent_.
Instead the XSpreadsheetDocument_ interface must be explicitly converted to XComponent_ first, as in:

.. tabs::

    .. code-tab:: python

        # in BuildTable._add_picture() of build_table.py
        # ...
        comp_doc = Lo.qi(XComponent, doc, True)
        print(f"2. No. of draw pages: {Draw.get_slides_count(comp_doc)}")
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch20_adding_chart:

20.2.7 Adding a Chart
---------------------

.. todo::

    Chapter 20.2.7 Add reference to Part 5

Charting is discussed at length in Part 5, but for now here is a taster of it here since a CellRangeAddress_ object is used to pass data to the charting methods.
For example, the cell range for ``A1:N4`` is passed to :py:meth:`.Chart2.insert_chart`:

.. tabs::

    .. code-tab:: python

        # in BuildTable.main() of build_table.py
        # assumes _build_array() has filled the spreadsheet with data
        rng_addr = Calc.get_address(sheet=sheet, range_name="B2:M4")
        chart_cell = "B6" if self._add_pic else "D6"
        Chart2.insert_chart(
            sheet=sheet, cells_range=rng_addr, cell_name=chart_cell, width=21, height=11, diagram_name="Column"
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The other arguments used by :py:meth:`.Chart2.insert_chart` are a cell name, the millimeter width and height of the generated chart, and a chart type string.
The named cell acts as an anchor point for the top-left corner of the chart. :numref:`ch20fig_bt_column_chart` shows what the resulting chart looks like.

..
    figure 4

.. cssclass:: screen_shot invert

    .. _ch20fig_bt_column_chart:
    .. figure:: https://user-images.githubusercontent.com/4193389/202811720-a7374f7b-8c8e-4f61-960d-ef482891479d.png
        :alt: A Column Chart in a Spreadsheet
        :width: 550px
        :figclass: align-center

        :A Column Chart in a Spreadsheet.

.. |show_sheet| replace:: Show Sheet
.. _show_sheet: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_show_sheet

.. |show_sheet_py| replace:: show_sheet.py
.. _show_sheet_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_show_sheet/show_sheet.py

.. |build_tbl| replace:: Build Table
.. _build_tbl: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_build_table

.. |build_tbl_py| replace:: build_table.py
.. _build_tbl_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_build_table/build_table.py

.. _CellRangeAddress: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1table_1_1CellRangeAddress.html
.. _Point: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1awt_1_1Point.html
.. _Spreadsheet: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1Spreadsheet.html
.. _SpreadsheetView: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SpreadsheetView.html
.. _SpreadsheetViewSettings: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SpreadsheetViewSettings.html
.. _XCell: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCell.html
.. _XCellRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCellRange.html
.. _XCellRangeData: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XCellRangeData.html
.. _XComponent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XComponent.html
.. _XDrawPage: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPage.html
.. _XDrawPageSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPageSupplier.html
.. _XProtectable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XProtectable.html
.. _XSpreadsheet: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSpreadsheet.html
.. _XSpreadsheetDocument: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSpreadsheetDocument.html
.. _XSpreadsheetView: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSpreadsheetView.html
