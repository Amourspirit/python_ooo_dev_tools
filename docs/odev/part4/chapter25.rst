.. _ch25:

*****************************
Chapter 25. Monitoring Sheets
*****************************

.. topic:: Overview

    Listening for Document Modifications (XModifyListener_); Listening for Application Closing (XTopWindowListener_); Listening for Cell Selection (XSelectionChangeListener_)

    Examples: |mod_list|_ and |sel_list|_.

The chapter looks at three kinds of listeners for spreadsheets: document modification, application closing, and cell selection listeners.

Office's support for listeners was first described back in :ref:`ch04`.

.. _ch25_listenf_for_mods:

25.1 Listening for the User's Modifications
===========================================

A common requirement for spreadsheet programming is controlling how the user interacts with the sheet's data.
In the extreme case, this might mean preventing the user from changing anything, which is possible through the XProtectable_ interface discussed in :ref:`ch20_read_only_protect_view`.
But often we want to let the user edit the sheet, but monitor what is being changed.

One way of doing this is to attach a XModifyListener_ interface to the open document so that its ``modified()`` method will be triggered whenever a cell is changed.

|mod_list|_ illustrates this approach:

.. tabs::

    .. group-tab:: Python

        .. tabs::

            .. tab:: ModifyListenerAdapter

                .. code-block:: python

                    class ModifyListenerAdapter:
                        def __init__(self, out_fnm: PathOrStr) -> None:
                            super().__init__()
                            if out_fnm:
                                outf = FileIO.get_absolute_path(out_fnm)
                                _ = FileIO.make_directory(outf)
                                self._out_fnm = outf
                            else:
                                self._out_fnm = ""
                            self.closed = False
                            loader = Lo.load_office(Lo.ConnectPipe())
                            self._doc = Calc.create_doc(loader)

                            GUI.set_visible(is_visible=True, odoc=self._doc)
                            self._sheet = Calc.get_sheet(doc=self._doc, index=0)

                            # insert some data
                            Calc.set_col(
                                sheet=self._sheet,
                                cell_name="A1",
                                values=("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3)
                            )

                            # pass GenericArgs with listener arg of self.
                            # this will allow for this instance to be passed to events.
                            # pass doc to constructor,
                            # this will allow listener to be automatically attached to document.
                            self._m_listener = ModifyListener(
                                trigger_args=GenericArgs(listener=self), doc=self._doc
                            )
                            self._m_listener.on("modified", ModifyListenerAdapter.on_modified)
                            self._m_listener.on("disposing", ModifyListenerAdapter.on_disposing)

                            # close down when window closes
                            self._twl = TopWindowListener(trigger_args=GenericArgs(listener=self))
                            self._twl.on("windowClosing", ModifyListenerAdapter.on_window_closing)

                        @staticmethod
                        def on_window_closing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
                            print("Closing")
                            try:
                                ml = cast(ModifyListenerAdapter, kwargs.get("listener", None))
                                if ml:
                                    Lo.close_doc(ml._doc)
                                    Lo.close_office()
                                    ml.closed = True
                            except Exception as e:
                                print(f"  {e}")

                        @staticmethod
                        def on_modified(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
                            print("Modified")
                            try:
                                event = cast("EventObject", event_args.event_data)
                                ml = cast(ModifyListenerAdapter, kwargs["listener"])
                                doc = Lo.qi(XSpreadsheetDocument, event.Source, True)
                                addr = Calc.get_selected_cell_addr(doc)
                                print(
                                    f"  {Calc.get_cell_str(addr=addr)} = {Calc.get_val(sheet=ml._sheet, addr=addr)}"
                                )
                            except Exception as e:
                                print(e)

                        @staticmethod
                        def on_disposing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
                            print("Disposing")

            .. tab:: ModifyListener

                .. code-block:: python

                    class ModifyListener(unohelper.Base, XModifyListener):
                        def __init__(self, out_fnm: PathOrStr) -> None:
                            super().__init__()
                            if out_fnm:
                                outf = FileIO.get_absolute_path(out_fnm)
                                _ = FileIO.make_directory(outf)
                                self._out_fnm = outf
                            else:
                                self._out_fnm = ""
                            self.closed = False
                            loader = Lo.load_office(Lo.ConnectPipe())
                            self._doc = Calc.create_doc(loader)

                            GUI.set_visible(is_visible=True, odoc=self._doc)
                            self._sheet = Calc.get_sheet(doc=self._doc, index=0)

                            # insert some data
                            Calc.set_col(
                                sheet=self._sheet,
                                cell_name="A1",
                                values=("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3)
                            )

                            mb = Lo.qi(XModifyBroadcaster, self._doc, True)
                            mb.addModifyListener(self)

                            # close down when window closes
                            self._twl = TopWindowListener(trigger_args=GenericArgs(listener=self))
                            self._twl.on("windowClosing", ModifyListener.on_window_closing)

                        @staticmethod
                        def on_window_closing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
                            print("Closing")
                            try:
                                ml = cast(ModifyListener, kwargs.get("listener", None))
                                if ml:
                                    Lo.close_doc(ml._doc)
                                    Lo.close_office()
                                    ml.closed = True
                            except Exception as e:
                                print(f"  {e}")

                        def modified(self, event: EventObject) -> None:
                            print("Modified")
                            doc = Lo.qi(XSpreadsheetDocument, event.Source, True)
                            addr = Calc.get_selected_cell_addr(doc)
                            print(
                                f"  {Calc.get_cell_str(addr=addr)} = {Calc.get_val(sheet=self._sheet, addr=addr)}"
                            )

                        def disposing(self, event: EventObject) -> None:
                            print("Disposing")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

|mod_list|_ example utilizes one of two classes, ``ModifyListenerAdapter`` of |mod_list_adapter_py|_
or ``ModifyListener`` of |mod_list_py|_. These classes are functionally the same.
These two class are interchangeable and are for example purposes. We also seen this in :ref:`ch04_listen_win`.

We will focus on ``ModifyListenerAdapter`` here.

.. _ch25_listening_close_box:

25.1.1 Listening to the Close Box
---------------------------------

``__init__()`` creates a ModifyListener object and then terminates, which means that the object must deal with the closing of the spreadsheet and the termination of Office.


This is done by employing another listener: an adapter for XTopWindowListener_, |top_window_listener|, attached to the Calc application's close box:

.. tabs::

    .. code-tab:: python

        # in modify_listener_adapter.py
        # close down when window closes
        def __init__(self, out_fnm: PathOrStr) -> None:
            # ...
            self._twl = TopWindowListener(trigger_args=GenericArgs(listener=self))
            self._twl.on("windowClosing", ModifyListenerAdapter.on_window_closing)
            # ...


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

XTopWindowListener_ was described in :ref:`ch04_listen_win`, but |top_window_listener| is an |odev| support class in the :ref:`adapter` namespace.

XTopWindowListener_ defines eight methods, called when the application window is in different states: ``opened``, ``activated``, ``deactivated``, ``minimized``, ``normalized``, ``closing``, ``closed``, and ``disposed``.
|top_window_listener| supplies empty implementations for those methods:

.. tabs::

    .. code-tab:: python

        class TopWindowListener(AdapterBase, XTopWindowListener):

            def __init__(
                self, trigger_args: GenericArgs | None = None, add_listener: bool = True
            ) -> None:
                super().__init__(trigger_args=trigger_args)
                if add_listener:
                    self._tk = mLo.Lo.create_instance_mcf(
                        XExtendedToolkit, "com.sun.star.awt.Toolkit", raise_err=True
                    )
                    if self._tk is not None:
                        self._tk.addTopWindowListener(self)

            def windowOpened(self, event: EventObject) -> None:
                self._trigger_event("windowOpened", event)

            def windowActivated(self, event: EventObject) -> None:
                self._trigger_event("windowActivated", event)

            def windowDeactivated(self, event: EventObject) -> None:
                """Is invoked when a window is deactivated."""
                self._trigger_event("windowDeactivated", event)

            def windowMinimized(self, event: EventObject) -> None:
                self._trigger_event("windowMinimized", event)

            def windowNormalized(self, event: EventObject) -> None:
                self._trigger_event("windowNormalized", event)

            def windowClosing(self, event: EventObject) -> None:
                self._trigger_event("windowClosing", event)

            def windowClosed(self, event: EventObject) -> None:
                self._trigger_event("windowClosed", event)

            def disposing(self, event: EventObject) -> None:
                self._trigger_event("disposing", event)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

|mod_list_adapter_py|_ subscribes to ``windowClosing()``, and ignores the other methods. ``windowClosing()`` is triggered when the application's close box is clicked,
and it responds by closing the document and Office:

.. tabs::

    .. code-tab:: python

        # in modify_listener_adapter.py
        @staticmethod
        def on_window_closing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            print("Closing")
            try:
                ml = cast(ModifyListenerAdapter, kwargs.get("listener", None))
                if ml:
                    Lo.close_doc(ml._doc)
                    Lo.close_office()
                    ml.closed = True
            except Exception as e:
                print(f"  {e}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch25_listening_for_modifications:

25.1.2 Listening for Modifications
----------------------------------

|modify_listener| is notified of document changes by attaching itself to the document's XModifyBroadcaster_:

.. tabs::

    .. code-tab:: python

        # in ModifyListener class
        def __init__(self, trigger_args: GenericArgs | None = None, doc: XComponent | None = None) -> None:
            super().__init__(trigger_args=trigger_args)
            if doc is None:
                return

            mb = Lo.qi(XModifyBroadcaster, doc, True)
            mb.addModifyListener(self)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

|mod_list_adapter_py|_ Creates an instance of |modify_listener| and subscribes to ``modified`` event.

.. tabs::

    .. code-tab:: python

        # in modify_listener_adapter.py
        def __init__(self, out_fnm: PathOrStr) -> None:
            # ...
            self._m_listener = ModifyListener(trigger_args=GenericArgs(listener=self), doc=self._doc)
            self._m_listener.on("modified", ModifyListenerAdapter.on_modified)
            # ...

        @staticmethod
        def on_modified(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            print("Modified")
            try:
                event = cast("EventObject", event_args.event_data)
                # print("Modified: " + event.Source)
                # Info.show_services("Event source", event.Source)

                ml = cast(ModifyListenerAdapter, kwargs["listener"])
                # doc = Lo.qi(XSpreadsheetDocument, event.Source, True)
                doc  = ml._doc
                addr = Calc.get_selected_cell_addr(doc)
                print(f"  {Calc.get_cell_str(addr=addr)} = {Calc.get_val(sheet=ml._sheet, addr=addr)}")
            except Exception as e:
                print(e)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


An :py:class:`~.events.args.event_args.EventArgs` object arriving at ``modified()`` has an ``event_data`` property that is an EventObject_ with a ``Source`` field of type XInterface_.
Every Office interface inherits XInterface_ so it's difficult to know what the source really is.
The simplest solution is to print the names of the source's supported services, by calling :py:meth:`.Info.show_services`, as seen in the commented-out code above.

In this case, the ``Source`` field is supported by the SpreadsheetDocument_ service, which means that it can be converted into an XSpreadsheetDocument_ interface.
Lots of useful things can be accessed through this interface, but that's also commented-out because ``ml._doc`` field points to the ``doc``.

.. _ch25_examining_changed_cells:

25.1.3 Examining the Changed Cell (or Cells)
--------------------------------------------

While ``modified()`` is being executed, the modified cell in the document is still selected (or active), and so can be retrieved:

.. tabs::

    .. code-tab:: python

        # in modify_listener_adapter.py
        addr = Calc.get_selected_cell_addr(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.get_selected_cell_addr` needs the XModel_ interface for the document so that ``XModel.getCurrentSelection()`` can be called.
It also has to handle the possibility that a cell range is currently selected rather than a single cell:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @classmethod
        def get_selected_cell_addr(cls, doc: XSpreadsheetDocument) -> CellAddress:
            cr_addr = cls.get_selected_addr(doc=doc)
            if cls.is_single_cell_range(cr_addr):
                sheet = cls.get_active_sheet(doc)
                cell = cls.get_cell(sheet=sheet, col=cr_addr.StartColumn, row=cr_addr.StartRow)
                return cls.get_cell_address(cell)
            else:
                raise CellError("Selected address is not a single cell")

        @overload
        @classmethod
        def get_selected_addr(cls, doc: XSpreadsheetDocument) -> CellRangeAddress:
            model = Lo.qi(XModel, doc)
            return cls.get_selected_addr(model)


        @overload
        @classmethod
        def get_selected_addr(cls, model: XModel) -> CellRangeAddress:
            ra = Lo.qi(XCellRangeAddressable, model.getCurrentSelection(), raise_err=True)
            return ra.getRangeAddress()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        - :odev_src_calc_meth:`get_selected_cell_addr`
        - :odev_src_calc_meth:`get_selected_addr`

:py:meth:`.Calc.get_selected_cell_addr` utilizes :py:meth:`.Calc.get_selected_addr`, which returns the address of the selected cell range.
:py:meth:`.Calc.get_selected_cell_addr` examines this cell range to see if it's really just a single cell by calling :py:meth:`.Calc.is_single_cell_range`:


.. tabs::

    .. code-tab:: python

        # in Calc class
        @staticmethod
        def is_single_cell_range(cr_addr: CellRangeAddress) -> bool:
            return cr_addr.StartColumn == cr_addr.EndColumn and cr_addr.StartRow == cr_addr.EndRow

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

If the cell range is referencing a cell then the cell range address position is used to directly access the cell in the sheet:

.. tabs::

    .. code-tab:: python

        # in Calc.get_selected_cell_addr()
        sheet = cls.get_active_sheet(doc)
        cell = cls.get_cell(sheet=sheet, col=cr_addr.StartColumn, row=cr_addr.StartRow)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This requires the current active sheet, which is obtained through :py:meth:`.Calc.get_active_sheet`.

.. _ch25_problems_with_modify:

25.1.4 Problems with the modified() Method
------------------------------------------

After all this coding, the bad news is that ``modified()`` is still lacking in functionality.

One minor problem is that ``modified()`` is called twice when the user finishes editing a cell.
This occurs when the user presses enter, or tab, or an arrow key, and for unknown reasons.
It could be fixed with some judicious hacking: :abbreviation:`i.e.` by using a counter to control when the code is executed.

A more important concern is that ``modified()`` only has access to the new value in the cell, but doesn't know what was overwritten,
which would be very useful for implementing data validation.
This led to investigation of another form of listening, based on cell selection, which is described next.

.. _ch25_listen_cell_select:

25.2 Listening for Cell Selections
==================================

Listening to cell selections on the sheet has the drawback of generating a lot of events, but this abundance of data turns out to be useful;
It can be used to report more about cell modifications.

The |sel_list|_ example is similar to |mod_list|_ except that it uses |selection_change_listener| rather than |modify_listener|:

.. tabs::

    .. code-tab:: python

        # in select_listener.py
        class SelectionListener:
            def __init__(self) -> None:
                super().__init__()
                self.closed = False
                loader = Lo.load_office(Lo.ConnectSocket())
                self._doc = Calc.create_doc(loader)

                GUI.set_visible(is_visible=True, odoc=self._doc)
                self.sheet = Calc.get_sheet(doc=self._doc, index=0)

                self.curr_addr = Calc.get_selected_cell_addr(self._doc)
                self.curr_val = self._get_cell_float(self.curr_addr)  # may be None

                self._attach_listener()

                # insert some data
                Calc.set_col(sheet=self.sheet, cell_name="A1", values=(
                    "Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3)
                )

                # close down when window closes
                self._twl = TopWindowListener(trigger_args=GenericArgs(listener=self))
                self._twl.on("windowClosing", SelectionListener.on_window_closing)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

|sel_list_py|_ also keeps track of variables  holding the address of the currently selected cell (``self.curr_addr``) and its numerical value (``self.curr_val``).
If the cell doesn't contain a float then ``self.curr_val`` is assigned ``None``. ``self.curr_addr`` and ``self.curr_val`` are initialized after the document is first created, and are updated whenever the user changes a cell.

``_attach_listener()`` is called to attach the listener to the document:

.. tabs::

    .. code-tab:: python

        # in select_listener.py
        def _attach_listener(self) -> None:
            self._s_listener = SelectionChangeListener(
                trigger_args=GenericArgs(listener=self), doc=self._doc
            )
            self._s_listener.on("selectionChanged", SelectionListener.on_selection_changed)
            self._s_listener.on("disposing", SelectionListener.on_disposing)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The current document is passed to |selection_change_listener| which handles setting up the XSelectionSupplier_.

``on_selection_changed()`` listens for three kinds of changes in the sheet:

1. it reports when the selected cell changes by printing the name of the previous cell and the newly selected one;
2. it reports whether the cell that has just lost focus now has a value different from when it was selected;
3. it reports if the newly selected cell contains a numerical value.

For example, :numref:`ch25fig_selection_sheet_data` shows the initial sheet of data created by |sel_list_py|_:

..
    figure 1

.. cssclass:: screen_shot invert

    .. _ch25fig_selection_sheet_data:
    .. figure:: https://user-images.githubusercontent.com/4193389/205182487-b1796a72-ec04-4bdc-8a8c-26acdf72039e.png
        :alt: The Sheet of Data in SelectListener
        :figclass: align-center

        :|sel_list|_ Sheet Data.

Note that the selected cell when the sheet is first created is ``A1``.

If the user carries out the following operations:

.. cssclass:: ul-list

    - click in cell ``B2``
    - click in cell ``A4``
    - click in ``A5``
    - change ``A5`` to ``4`` and press tab

then the sheet will end up looking like :numref:`ch25fig_selection_sheet_modified_data`, with ``B5`` being the selected cell.

..
    figure 2

.. cssclass:: screen_shot invert

    .. _ch25fig_selection_sheet_modified_data:
    .. figure:: https://user-images.githubusercontent.com/4193389/205191488-3df39fa0-2fdc-424f-b42a-2c9cd9039c56.png
        :alt: SelectListener modified data
        :figclass: align-center

        :|sel_list|_ Modified Sheet.

During these changes, ``on_selection_changed()`` will report:

::

    A2 value: 42.0
    A3 value: 58.9
    A4 value: -66.5
    A5 value: 43.4
    A5 value: 43.4
    A5 has changed from 43.40 to 4.00

The "value" lines state the value of a cell when it's first selected, and the "changed" lines report whether the cell was left changed when the focus moved to another cell.

The output from ``on_selection_changed()`` shown above shows how the user moved around the spreadsheet, and changed the ``A5`` cell's contents from ``43.4`` to ``4``.

``on_selection_changed()`` is defined as:

.. tabs::

    .. code-tab:: python

        # in select_listener.py
        @staticmethod
        def on_selection_changed(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            event = cast("EventObject", event_args.event_data)
            ctrl = Lo.qi(XController, event.Source)
            if ctrl is None:
                print("No ctrl for event source")
                return
            sl = cast(SelectionListener, kwargs["listener"])
            addr = Calc.get_selected_cell_addr(sl._doc)
            if addr is None:
                return
            try:
                # better to wrap in try block.
                # otherwise errors crahses office
                if not Calc.is_equal_addresses(addr, sl.curr_addr):
                    flt = sl._get_cell_float(sl.curr_addr)
                    if flt is not None:
                        if sl.curr_val is None:  # so previously stored value was null
                            print(f"{Calc.get_cell_str(sl.curr_addr)} new value: {flt:.2f}")
                        else:
                            if sl.curr_val != flt:
                                print(
                                    f"{Calc.get_cell_str(sl.curr_addr)} has changed from {sl.curr_val:.2f} to {flt:.2f}"
                                )

                # update current address and value
                sl.curr_addr = addr
                sl.curr_val = sl._get_cell_float(addr)
                if sl.curr_val is not None:
                    print(f"{Calc.get_cell_str(sl.curr_addr)} value: {sl.curr_val}")
            except Exception as e:
                print(e)

        def _get_cell_float(self, addr: CellAddress) -> float | None:
            # get_val returns floats as float intance
            obj = Calc.get_val(sheet=self.sheet, addr=addr)
            if isinstance(obj, float):
                return obj
            return None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


``on_selection_changed()`` is called whenever the user selects a new cell.
The address of this new cell is obtained by :py:meth:`.Calc.get_selected_cell_addr`, which returns null if the user has selected a cell range.

If the new selection is a cell then a series of comparisons are carried out between the previously selected cell address and
value (stored in ``self.curr_addr`` and ``self.curr_val``) and the new address and its possible numerical value (stored in ``addr`` and ``flt``).
At the end of the method the current address and value are updated with the new ones.

XSelectionChangeListener_ shares a similar problem to XModifyListener_ in that a single user selection triggers multiple calls to ``selectionChanged()``.
Clicking once inside a cell causes four calls, and an arrow key press may trigger two calls depending on how it's entered from the keyboard.


.. |mod_list| replace:: Calc Modify Listener
.. _mod_list: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_modify_listener

.. |mod_list_py| replace:: modify_listener.py
.. _mod_list_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/calc/odev_modify_listener/modify_listener.py

.. |mod_list_adapter_py| replace:: modify_listener_adapter.py
.. _mod_list_adapter_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/calc/odev_modify_listener/modify_listener_adapter.py

.. |sel_list| replace:: Calc Select Listener
.. _sel_list: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_select_listener

.. |sel_list_py| replace:: select_listener.py
.. _sel_list_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/calc/odev_select_listener/select_listener.py

.. |top_window_listener| replace:: :ref:`TopWindowListener <adapter_awt_top_window_listener>`
.. |modify_listener| replace:: :ref:`ModifyListener <adapter_util_modify_listener>`
.. |selection_change_listener| replace:: :ref:`SelectionChangeListener <adapter_view_selection_change_listener>`

.. _EventObject: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1lang_1_1EventObject.html
.. _SpreadsheetDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SpreadsheetDocument.html
.. _XInterface: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1uno_1_1XInterface.html
.. _XModel: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XModel.html
.. _XModifyBroadcaster: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XModifyBroadcaster.html
.. _XModifyListener: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XModifyListener.html
.. _XProtectable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XProtectable.html
.. _XSelectionChangeListener: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1view_1_1XSelectionChangeListener.html
.. _XSelectionSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1view_1_1XSelectionSupplier.html
.. _XSpreadsheetDocument: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSpreadsheetDocument.html
.. _XTopWindowListener: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XTopWindowListener.html
