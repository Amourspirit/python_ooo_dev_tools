***************
Version History
***************

Version 0.44.2
==============

Added ``ooodev.io.sfa.Sfa`` class for working with Simple File Access. This class can be used to read/write/copy and delete files embedded in the document.
This class can bridge from the document to the file system.

Version 0.44.1
==============

Added ``ooodev.calc.CalcSheet.code_name`` and  ``ooodev.calc.CalcSheet.unique_id`` that is used to access sheet code name and unique id respectively.

Added ``get_sheet_name_from_code_name()`` and ``get_sheet_name_from_unique_id()`` to ``ooodev.calc.CalcDoc``
that can be used to look up the current sheet name from the sheet code name or the sheet unique id.

Version 0.44.0
==============

Several new classes in the ``ooodev.adapter`` module for working with LibreOffice objects.

Other minor updates and additions.

Subprocess
----------

Now a subprocess can be used when needed.

Main script

.. code-block:: python

    from __future__ import annotations
    import logging
    import sys
    import os
    from pathlib import Path
    import subprocess
    import uno

    from ooodev.calc import CalcDoc
    from ooodev.loader import Lo
    from ooodev.loader.inst.options import Options


    def main():

        loader = Lo.load_office(connector=Lo.ConnectPipe(), opt=Options(log_level=logging.DEBUG))
        doc = CalcDoc.create_doc(loader=loader, visible=True)
        try:
            # Start the subprocess
            script_path = Path(__file__).parent / "myscript.py"
            env = os.environ.copy()
            env["PYTHONPATH"] = get_paths()
            proc = subprocess.Popen(
                [sys.executable, str(script_path)],
                stdin=subprocess.PIPE,
                env=env,
            )

        finally:
            doc.close()
            Lo.close_office()


    def get_paths() -> str:
        pypath = ""
        p_sep = ";" if os.name == "nt" else ":"
        for d in sys.path:
            pypath = pypath + d + p_sep
        return pypath


    if __name__ == "__main__":
        main()


``myscript.py``

.. code-block:: python

    from __future__ import annotations
    import sys
    import os
    from ooodev.calc import CalcDoc
    from ooodev.utils.string.str_list import StrList
    from ooodev.loader import Lo
    from ooodev.conn import conn_factory
    from ooodev.loader.inst.options import Options


    def main():
        conn_str = os.environ.get("ODEV_CURRENT_CONNECTION", "")
        conn_opt = os.environ.get("ODEV_CURRENT_CONNECTION_OPTIONS", None)

        conn = conn_factory.get_from_json(conn_str)
        if conn_opt:
            opt = Options.deserialize(conn_opt)
        else:
            opt = Options()
    
        loader = Lo.load_office(connector=conn, opt=opt)  # type: ignore
        doc = CalcDoc.from_current_doc()
        sheet = doc.get_active_sheet()
        sheet[0, 0].value = "Hello World!"
        # ...


Breaking changes
----------------

``doc.python_script.write_file()`` method longer has a ``allow_override`` arg. Now has a ``mode`` arg that can be ``a`` (append), ``w`` (overwrite if existing, default) or ``x`` (error if exist).


Version 0.43.2
==============

Update Dialog Controls to have a static ``create()`` method that can be used to create controls for a Top Window.

Version 0.43.1
==============

Auto Load Disabled. Auto Load is currently causing issue when OooDev is being using in an Extension.


Version 0.43.0
==============

Read and Write Python Macro Code
--------------------------------

Now it is possible to read and write Python macro code to documents.

This example writes a Python script to a document and then reads it back.

The python macros are persisted when the document is saved and re-opened.

.. code-block:: python

    from __future__ import annotations
    import logging
    import uno

    from ooodev.calc import CalcDoc
    from ooodev.loader import Lo
    from ooodev.loader.inst.options import Options
    from ooodev.utils.string.str_list import StrList


    def main():

        loader = Lo.load_office(connector=Lo.ConnectPipe(), opt=Options(log_level=logging.DEBUG))
        doc = CalcDoc.create_doc(loader=loader, visible=True)
        try:
            psa = doc.python_script
            assert psa is not None
            code = StrList(sep="\n")
            code.append("from __future__ import annotations")
            code.append()
            code.append("def say_hello() -> None:")
            with code.indented():
                code.append('print("Hello World!")')
            code.append()
            code_str = str(code)
            assert psa.is_valid_python(code_str)
            psa.write_file("MyFile", code_str, allow_override=True)
            psa_code = psa.read_file("MyFile")
            assert psa_code == code_str

        finally:
            doc.close()
            Lo.close_office()


    if __name__ == "__main__":
        main()



Write Basic code
----------------

Now it is possible to write and add ``basic`` scripts to documents.

This example shows how to add a basic script to a Calc document.

The basic macro is persisted when the document is saved and re-opened.


.. code-block:: python

    from __future__ import annotations
    import logging
    import uno

    from ooodev.calc import CalcDoc
    from ooodev.loader import Lo
    from ooodev.loader.inst.options import Options
    from ooodev.utils.string.str_list import StrList
    from ooodev.adapter.container.name_container_comp import NameContainerComp
    from ooodev.macro.script.macro_script import MacroScript


    def main():
        loader = Lo.load_office(connector=Lo.ConnectPipe(), opt=Options(log_level=logging.DEBUG))
        doc = CalcDoc.create_doc(loader=loader, visible=True)
        try:
            inst = doc.basic_libraries
            mod_name = "MyModule"
            lib_name = "MyLib"
            clean = True
            added_lib = False

            if not inst.has_by_name(lib_name):
                added_lib = True
                inst.create_library(lib_name)

            inst.load_library(lib_name)

            lib = NameContainerComp(inst.get_by_name(lib_name))  # type: ignore
            if lib.has_by_name(mod_name):
                lib.remove_by_name(mod_name)

            code = StrList(sep="\n")
            code.append("Option Explicit")
            code.append("Sub Main")
            with code.indented():
                code.append('MsgBox "Hello World"')
            code.append("End Sub")
            lib.insert_by_name(mod_name, code.to_string())

            MacroScript.call(
                name="Main",
                library=lib_name,
                module=mod_name,
                location="document",
            )
            print("Macro Executed")
            if clean:
                lib.remove_by_name(mod_name)
                if added_lib:
                    inst.remove_library(lib_name)

            print("Done")
        finally:
            doc.close()
            Lo.close_office()

    if __name__ == "__main__":
        main()

Auto loader
-----------

A new Auto load for the ``ooodev`` library has been added. Now the library attempts to automatically load the ``Lo`` class with ``from ooodev.loader import Lo``.
This should eliminate the need to manually call ``Lo.current_doc`` or use the ``MacroLoader`` before using the library.
Note this only for when the library is used in a macro. In a script the ``Lo`` class will still need to be loaded manually.

StrList/IndexAccessImplement
----------------------------

``ooodev.utils.string.str_list.StrList`` has been updated and now  support slicing.

``ooodev.adapter.container.index_access_implement.IndexAccessImplement`` has been updated and now supports slicing, iteration, reversed iteration, and length.

Hidden Controls
---------------

Update for Hidden Controls. Now hidden controls can be added to documents and are persisted when the document is saved and re-opened.

.. code-block:: python

    from __future__ import annotations
    from pathlib import Path
    import uno
    from ooo.dyn.beans.property_attribute import PropertyAttributeEnum
    from ooodev.calc import CalcDoc

    doc = CalcDoc.from_current_doc()

    sheet = doc.sheets[0]
    if len(sheet.draw_page.forms) == 0:
        frm = sheet.draw_page.forms.add_form("MyForm")
    else:
        frm = sheet.draw_page.forms[0]
    ctl = frm.insert_control_hidden(name="MyHidden")
    ctl.hidden_value = "Hello World"
    ctl.add_property("Special", PropertyAttributeEnum.CONSTRAINED, "Special Data")
    fnm = Path.cwd() / "tmp" / "hidden.ods"
    doc.save_doc(fnm)

Breaking Changes
----------------

The ``insert_control_hidden()`` method args have changed. Some args have been removed.
This should not affect preexisting code as the hidden control was not properly implemented before.

Version 0.42.1
==============

Added ``ooodev.io.zip.ZIP`` class for working with zip files.

Version 0.42.0
==============

Added :ref:`ooodev.io.xml.XML` for working with XML files.

Added ``ooodev.utils.string.text_steam.TextStream`` class for working Text Streams.

Add classes to ``ooodev.adapter.io`` module for working with Streams.

Added classes to ``ooodev.adapter.ucb`` module for working with Files.

Added classes to ``ooodev.adapter.packages.zip`` for working with zip files.

Global events
-------------

Global document events can be temporarily disabled via built in context manager.

.. code-block:: python

    from ooodev.write import WriteDoc

    doc = WriteDoc.from_current_doc()
    with doc.lo_inst.global_event_broadcaster:
        # do work. Global document events are disabled here.
        pass
    # global events are working again

Version 0.41.2
==============

Fix for ``Lo.kill_office()`` method. Was not closing Office on Linux and Mac. Note ``Lo.kill_office()`` forces close without saving.
Normally ``doc.close()`` with ``Lo.close_office()`` would be used.

Version 0.41.1
==============

Minor fix for embedding into a oooscript file.

Version 0.41.0
==============

Menus
-----

Many updates for working with menus. Now menus can be created and modified in a much easier way including importing an exporting json files.

See :ref:`help_common_menus`.

Global
------

Added ``ooodev.global`` module that contains global classes for the library.

The ``ooodev.global.GTC`` class is a global timed cache that can be used to cache objects for a set amount of time.

The ``ooodev.global.GblEvents`` class is a global event broadcaster that can be used to broadcast events to all listeners.

Caching
-------

Added ``ooodev.utils.cache.file_cache.PickleCache`` and ``ooodev.utils.cache.file_cache.TextCache`` cache classes.
These classes can be used to cache objects to disk in the LibreOffice Temp folder.
Optionally an expire time can be set for the cache.

ThePathSettingsComp
-------------------

Added ``ooodev.adapter.util.the_path_settings_comp.ThePathSettingsComp`` class.
This class gets access to the LibreOffice paths such as the Temp folder and the User folder.

.. code-block:: python

    >>> from ooodev.adapter.util.the_path_settings_comp import ThePathSettingsComp
    >>> path_settings = ThePathSettingsComp.from_lo()
    >>> print(str(path_settings.temp))
    file:///tmp

Lo Updates
----------

Now the ``Lo`` class not has a ``tmp_dir`` property that returns a python ``pathlib.Path`` object of the LibreOffice Temp folder.

.. code-block:: python

    >>> from ooodev.loader import Lo
    >>> print(str(Lo.tmp_dir))
    /tmp/


Version 0.40.1
==============

``LRUCache`` moved to ``ooodev.utils.cache`` module.

Added ``TimeCache`` and ``TLRUCache`` (Time and Least Recently used) to ``ooodev.utils.cache`` module.

Version 0.40.0
==============

Menu
----

New menu options have been added to the library for working with the menu system and menu shortcuts.
A lot of work has been done in this area.

See :ref:`help_common_menus` for more information.

GUI
---

The ``gui`` module has been moved from the ``ooodev.utils`` to the ``ooodev.gui`` module.

The old imports still work but are deprecated.

New proper usage:

.. code-block:: python

    from ooodev.gui import Gui
    # ...

New ``ooodev.macro.MacroScript`` class tha can be used to invoke python or basic macro scripts.

Many new enhancements to the underlying dynamic construction of components that implement services.
Now classes can be implemented based upon the services they support at runtime.

Caching
-------

Added a new caching class that can be used to cache objects.

The ``ooodev.utils.lru_cache.LRUCache`` class can be used to cache objects.

The an instance ``LRUCache`` is used in the ``Lo`` class and can be accessed via the ``Lo.cache`` property.
The ``Lo.cache`` can be used to cache objects that are used often.

The size of the cache can be set in the options if needed. The default size is ``200``.


.. code-block:: python

    from ooodev.loader import Lo
    from ooodev.loader.inst import Options

    loader = Lo.load_office(
        connector=Lo.ConnectPipe(),
        opt=Options(log_level=logging.DEBUG, lo_cache_size=400)
    )
    # ...
    Lo.cache["my_key"] = "my_value"
    assert Lo.cache["my_key"] == "my_value"

Logging
-------

A new logger has been added to the library.

The default logging level is ``logging.INFO``.

Currently there is only logging to the console.

The |odev| Library uses is currently using this logging in a limited way.
This will change in subsequent versions.

Logging Module
^^^^^^^^^^^^^^

This logger is a singleton and can be accessed via the ``ooodev.logger`` module.

To use the logger simply import the module and use th logging methods:

Logging Date format is in the format ``"%d/%m/%Y %H:%M:%S"`` (Day, Month, Year, Hour, Minute, Second).

.. code-block:: python

    from ooodev.io.log import logging as logger
    logger.info("Hello World")
    logger.error("Error has occured")

Named Logger
^^^^^^^^^^^^

For convenience a named logger has been added to the library.
It is a wrapper around the logger that allows for a name to be added to the log output.

.. code-block:: python

    from ooodev.io.log import NamedLogger

    class MyClass:
        def __init__(self):
            # ...
            self._logger = NamedLogger(name=f"{self.__class__.__name__} - {self._implementation_name}")

        def _process_import(self, arg) -> None:
            # ...
            clz = self._get_class(arg)
            self._add_base(clz, arg)
            self._logger.debug(f"Added: {arg.ooodev_name}")
            # ...

The log output might look like this:

.. code-block::

    09/04/2024 10:15:45 - DEBUG - MyClass - ScTabViewObj: Added: ooodev.utils.partial.service_partial.ServicePartial

Logging Options
^^^^^^^^^^^^^^^

``Options`` now has a new ``log_level`` property that can be set to control the logging level of the library.

.. code-block:: python

    import logging
    from ooodev.loader.inst.options import Options

    loader = Lo.load_office(connector=Lo.ConnectPipe(), opt=Options(log_level=logging.DEBUG))
    # ...

Also the log level can be set via the logging module.

.. code-block:: python

    import logging
    from ooodev.io.log import logging as logger

    logger.set_log_level(logging.DEBUG)
    assert logger.get_log_level() == logging.DEBUG

Bug Fixes
---------

Fixed bug in ``ooodev.units.UnitMM10.from_unit_val()`` that was not converting the value correctly.

Version 0.39.1
==============

Update Form Controls to allow for better access to the control properties.
Form controls are now also context managers.

Using ``ctl.set_property()`` will automatically toggle control design  mode if needed.

Example of using a control as a context manager:

The width block will toggle design mode on and off.

.. code-block:: python

    with ctl:
        ctl.model.Width = 200   


Version 0.39.0
==============

Add dozens of new classes to support Extended view on controls.

Version 0.38.2
==============

Extended ``ooodev.adapter.sheet.spreadsheet_view_comp.SpreadsheetViewComp`` to include ``ooodev.adapter.view.form_layer_access_partial.FormLayerAccessPartial``.

Now checking of a Calc document in in design mode can be done as follows:

.. code-block:: python

    >>> from ooodev.calc import CalcDoc
    >>> doc = CalcDoc.from_current_doc()
    >>> view = doc.get_view()
    >>> view.is_form_design_mode()
    False

    >>> view.set_form_design_mode(True)
    >>> view.is_form_design_mode()
    True

Version 0.38.1
==============

Added new :ref:`ooodev.utils.context.dispatch_context.DispatchContext`.

Now Message box ``boxtype`` can also accept an ``int`` value.

Minor updates and bug fixes.

Version 0.38.0
==============

Cell and Range Controls
-----------------------

Add a new property to ``CalcCell`` and ``CalcCellRange`` called ``controls``.
This property returns a ``CalcCellControls`` and ``CalcCellRangeControls`` class respectively.
These classes can be used to access and manipulate the form controls in a cell or range.
In other words this makes it super simple to add controls to a cell or a range.

.. code-block:: python

    from ooodev.calc import CalcDoc
    doc = CalcDoc.create_doc(visible=True)
    sheet = doc.sheets[0]

    cell = sheet["A1"]
    chk = cell.control.insert_control_check_box("My CheckBox", tri_state=False)
    assert chk is not None

    cell = sheet["A1"]
    chk = cell.control.current_control
    assert chk is not None

    cell = sheet["B3"]
    btn = cell.control.insert_control_button("My Button")
    assert btn is not None

    cell = sheet["B3"]
    btn = cell.control.current_control

    rng = sheet.get_range(range_name="b10:c12")
    list_box = rng.control.insert_control_list_box(entries=["D", "E", "F"], drop_down=False)


Basic Script Access
-------------------

Add a new Basic script manager that can be used to access basic scripts.

.. code-block:: python

    ooodev.macro.script.basic import Basic
    def r_trim(input: str, remove: str = " ") -> str:
        script = Basic.get_basic_script(macro="RTrimStr", module="Strings", library="Tools", embedded=False)
        res = script.invoke((input, remove), (), ())
        return res[0]
    result = r_trim("hello ")
    assert result == "hello"

Forms
-----

Now it is possible to Find a shape in a Draw Page with the ``Form.find_shape_for_control()`` method.

Also a new ``Form.find_cell_with_control()`` method has been added that can be used to find a cell that contains a form control.

Version 0.37.0
==============

Added new reflect class ``ooodev.utils.reflection.reflect.Reflect`` that can be used to reflect UNO objects.

Added new ``ooodev.utils.kind.enum_helper.EnumHelper`` class that can be used to get the enum values of a UNO object and can create dynamic enums.

Breaking Changes
----------------

Dialog controls now use ``UnitPX`` and ``AppFont*`` classes for measurements.
Int values can still be used to set measurements as before and still default to Pixels.
Now the default is ``UnitPX`` for measurements.
Dialog UNO controls by default use pixels for View measurements and App Font measurements for Model measurements.

This change should not affect most users as the default is still pixels. But now reading pixels will return a ``UnitPX`` object which.
Hint: ``int(my_unit_px)`` will return the pixel value as in int.

Version 0.36.3
==============

Added new App Font Classes:

- ``ooodev.units.AppFontSize``
- ``ooodev.units.AppFontPos``
- ``ooodev.units.UnitAppFontWidth``
- ``ooodev.units.UnitAppFontHeight``.

Version 0.36.2
==============

Fix for ``ooodev.units.UnitAppFont`` Now ``UnitAppFont`` is ``UnitAppFontX``. Added a new ``UnitAppFontY`` class.

LibreOffice Office uses different ``AppFont`` values for X and Y.

Version 0.36.1
==============

Minor adjustment for ``ooodev.dialog.dl_control.CtlGrid`` properties ``row_header_width``,  and ``row_height``.

Version 0.36.0
==============

Added ``ooodev.units.UnitAppFont`` class that can be used where App Font Measurements are used.
``UnitAppFont`` units may change value on different systems. This class is used for measurements that are based on the current system.

``ooodev.dialog.dl_control.CtlGrid`` now uses ``UnitAppFont`` for ``column_header_height``, ``row_header_width``,  and ``row_height`` properties.

Version 0.35.0
==============

Added all the same conversions found in `CONVERT function <https://help.libreoffice.org/latest/en-US/text/scalc/01/func_convert.html?&DbPAR=CALC&System=UNIX>`__
to :ref:`ns_units_convert`. There are enum for all the conversions.

Version 0.34.3
==============

Update for Draw Shapes. Now can access many more properties on various shapes.

Added ``ooodev.draw.shapes.shape_factory.ShapeFactory`` class that can be used to Convert ``XShape`` into ``OooDev`` Shapes.   

``ooodev.adapter.text.graphic_crop_struct_comp.GraphicCropStructComp`` Now is a Generic for Unit measurements.

``ooodev.adapter.drawing.rotation_descriptor_properties_partial.RotationDescriptorPropertiesPartial.shear_angle`` property is not optional.


Version 0.34.1
==============

Add a unit factory for converting units to other units. The module is ``ooodev.units.unit_factory``.

Draw shapes now have better support when selecting Group Shapes.

Shapes can now set size and position directly by setting the ``x``, ``y``, ``width`` and ``height`` properties of ``size`` and ``position`` properties.

New Generic ``ooodev.adapter.awt.size_struct_generic_comp.SizeStructGenericComp`` for working with sizes and generic Unit Sizes.
New Generic ``ooodev.adapter.awt.point_struct_generic_comp.PointStructGenericComp`` for working with positions and generic Unit Sizes.

Version 0.34.0
==============

Customs shapes are much more dynamic. when selecting shapes the list of the selected shapes have access to many more properties.
Many properties are added to shapes based upon the services they support.

Index containers in ``ooodev.adapter.container`` package are now generic. This allow for better tying support when working with elements in the container.

Created ``ooodev.office.partial.office_document_prop_partial.OfficeDocumentPropPartial`` and implement this class. It has bee implemented into all Documents and many other classes.

For instance Draw shapes implement ``OfficeDocumentPropPartial`` and this gives access to the document that the shape is in.

``DrawDoc`` class and ``ImpressDoc`` class now have a common base class ``DocPartial``.

Version 0.33.0
==============

Now there is a ``get_selected_shapes()`` method for ``DrawDoc`` and ``ImpressDoc`` that returns a list of the current selected shapes.

Many updates for Draw Shapes. Now can access many more properties on various shapes.

Now DrawDoc has a ``current_controller`` property that returns a ``DrawDocView`` instance.
``DrawDocView`` is a new class that represents a Draw document view.

Angles can now be added, subtracted, multiplied and divided to each other and the conversion is automatic.

.. code-block:: python

    from ooodev.units import Angle, Angle10
    a1 = Angle(90)
    a2 = Angle10(110) # 10 degrees
    a3 = a1 + a2
    assert isinstance(a3, Angle)
    assert a3 == 101


Version 0.32.2
==============

Added Table Border 2 for Writer Tables.

Version 0.32.1
==============

Added new formatting options to Write Tables.


Version 0.32.0
==============

Many classes added for working with Writer Tables. See :ref:`ns_write_table` namespace.

Other additions to Write to make accessing various parts of the document easier.

Other minor updates and bug fixes.

RangeObj Updates
----------------

Fix for ``RangeObj.get_row()`` returning the wrong row.

Update for ``RangeObj``. Now you can iterate over the cells in a range.

The iteration is done in a row-major order, meaning that the cells are iterated over by row, then by column.

.. code-block:: python

    >>> rng = RangeObj.from_range("A1:C4")
    >>> for cell in rng:
    >>>     print(cell)
    A1
    B1
    C1
    A2
    B2
    C2
    A3
    B3
    C3
    A4
    B4
    C4

The iteration can be especially useful when you want iterate over a row or a column in a range.

Iterating over a row in a range:

.. code-block:: python

    >>> rng = RangeObj.from_range("A1:C4")
    >>> for cell in rng.get_row(1):
    >>>     print(cell)
    A2
    B2
    C2

Iterating over a column in a range:

.. code-block:: python

    >>> rng = RangeObj.from_range("A1:C4")
    >>> for cell in rng.get_col("B"):
    >>>     print(cell)
    B1
    B2
    B3
    B4

Checking if range contains a cell This is functionally the same as the ``RangeObj.contains()`` method.

.. code-block:: python

    >>> rng = RangeObj.from_range("A1:C4")
    >>> assert "B2" in rng
    True

Getting a ``CellObj`` from a ``RangeObj``:

.. code-block:: python

    >>> rng = RangeObj.from_range("A1:C4")
    >>> cell = rng["B2"] # gets a CellObj instance
    >>> assert str(cell) == "B2"
    True

Version 0.31.0
==============

Massive refactoring of imports.
Inspired by `André Menck - Avoiding Circular Imports in Python <https://medium.com/brexeng/avoiding-circular-imports-in-python-7c35ec8145ed>`__ article.
This version saw then entire library refactored to help avoid circular imports. Over ``1,000`` modules were updated.
Now objects are always imported from the files where they are defined in. Test have be constructed to ensure this rule.

Version 0.30.4
==============

Minor updates. Better support for compiled script via the ``oooscript`` command line tool.

Version 0.30.3
==============

Minor updates and bug fixes.

Breaking Changes
----------------

``ooodev.write.WriteText.create_text_cursor()`` now return instance of ``ooodev.write.WriteTextCursor[WriteText]`` instead
of  ``XTextCursor``. Direct access to  can be done via ``WriteTextCursor.component``. or ``WriteText.component.createTextCursor()``.

``ooodev.write.WriteText.create_text_cursor_by_range()`` now return instance of ``ooodev.write.WriteTextCursor[WriteText]`` instead
of  ``XTextCursor``. Direct access to  can be done via ``WriteTextCursor.component``. or ``WriteText.component.create_text_cursor_by_range()``.



Version 0.30.2
==============

Added ``WriteTextViewCursor.style_direct_char`` that allows for direct character styling.

Same as changes for ``WriteTextCursor`` in version ``0.30.1``.

Version 0.30.1
==============

Added ``WriteTextCursor.style_direct_char`` that allows for direct character styling.

Example:

.. code-block:: python

    doc = WriteDoc.create_doc(visible=True)

    cursor = doc.get_cursor()
    cursor.append("hello")
    cursor.go_left(5, True)
    # font 30, bold, italic, underline, blue
    cursor.style_direct_char.style_font_general(
        size=30.0,
        b=True,
        i=True,
        u=True,
        color=StandardColor.BLUE,
    )
    cursor.goto_end()
    # reset the style before adding more text
    cursor.style_direct_char.clear()


Version 0.30.0
==============

Added search and replace methods to ``WriteDoc`` and ``WriteTextViewCursor``.


Version 0.29.0
==============

Added new ``CtlSpinButton`` class for working with Spin Button controls.
Update all controls to make formatting (font, text color, etc) easier.
This includes the ability to set font and text color for all controls that support it.

Version 0.28.4
==============

Added ``CalcCellTextCursor`` class that can be used to get the text of a cell. Cursor can be accessed via ``CalcCell.create_text_cursor()``.

.. code-block:: python

    from ooodev.calc import CalcDoc
    doc = CalcDoc.create_doc(visible=True)
    sheet = doc.sheets[0]
    cell = sheet["A1"]
    cursor = cell.create_text_cursor()
    cursor.append_para("Text in first line.")
    cursor.append("And a ")
    cursor.add_hyperlink(
        label="hyperlink",
        url_str="https://github.com/Amourspirit/python_ooo_dev_tools",
    )

Version 0.28.3
==============

``CalcCell`` and ``CalcCellRange`` now have ``style_by_name()`` methods that allow setting a cell or range style by name.

Version 0.28.2
==============

Added many style methods to Calc classes such as ``CalcCell`` and ``CalcCellRange``.

Version 0.28.1
==============

Minor fix for getting the current document in multi document usage.

Version 0.28.0
==============

Added :ref:`the_global_event_broadcaster`.

Added ``style_*get`` methods to many partial style classes.

Updated help docs for Chart2 styling.

Version 0.27.1
==============

Update documentation for Chart2 Calc related classes.

Other minor updates and bug fixes.

Version 0.27.0
==============

Big update for charts. Now charts can be created and modified in a much easier way.

Charts are now accessible via ``CalcSheet`` and ``CalcDoc`` classes.

Styling of most all chart objects is built into the chart objects themselves.

.. code-block:: python

    sheet = doc.sheets[0] # get the first sheet from the Calc doc
    range_addr = sheet.rng("A2:B8")
    tbl_chart = sheet.charts.insert_chart(
        rng_obj=range_addr,
        cell_name="C3",
        width=15,
        height=11,
        diagram_name=ChartTypes.Column.TEMPLATE_STACKED.COLUMN,
    )
    sheet["A1"].goto()

    chart_doc = tbl_chart.chart_doc
    _ = chart_doc.set_title(sheet["A1"].value)
    _ = chart_doc.axis_x.set_title(sheet["A2"].value)
    y_axis_title = chart_doc.axis_y.set_title(sheet["B2"].value)
    y_axis_title.style_orientation(angle=90)
    chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)

See :ref:`ns_calc_chart2`

Other minor updates and bug fixes.

Version 0.26.0
==============

The ``Lo`` class and other loader classes ahve been moved into ``ooodev.loader`` namespace.

Now ``Lo`` is imported as follows. ``from ooodev.loader import Lo``. This should not be a breaking change as the old import should still work.
Previous import was ``from ooodev.utils.lo import Lo``.

``Lo`` is basically the context manager for the entire library. It is used to connect to LibreOffice, manage the connection and communitate with Documents.
In this version the ``Lo`` and related classes have been update to have much better multi-document support.

``Lo`` class now has a ``desktop`` property that is an instance of the new ``ooodev.loader.comp.the_desktop.TheDesktop`` class.

Now in macro mode there are multiple ways to get the current document. The ``Lo`` class has a ``current_doc`` property that returns the current document.
In Macro Mode it is not necessary to use ``ooodev.macro.MacroLoader`` to access the document in the following mannor.

.. code-block:: python

    from ooodev.loader import Lo
    doc = Lo.current_doc
    doc.msgbox("Hello World")

or for know more specific document types such as ``ooodev.write.WriteDoc`` or ``ooodev.calc.CalcDoc``.

.. code-block:: python

    from ooodev.write import WriteDoc
    doc = WriteDoc.from_current_doc()
    doc.msgbox("Hello World")

.. code-block:: python

    from ooodev.calc import CalcDoc
    doc = CalcDoc.from_current_doc()
    doc.msgbox("Hello World")

Version 0.25.2
==============

Added the ability for Document classes to dispatch commands via the ``dispatch_cmd()``. This allows for dispatching to be done to the correct document in multi-document usage.

Other minor fixes and updates.

Breaking Changes
----------------

Removed redundant ``ooodev.calc.calc_cell_range.set_style()`` method. This method was not needed and was redundant with ``ooodev.calc.calc_cell_range.apply_styles()``.

Version 0.25.1
==============

Better support for `ooodev.utils.lo.Lo.current_doc` in macros. Now it is possible to use `ooodev.utils.lo.Lo.current_doc` in macros to get the current document without needing to use ``ooodev.macro.MacroLoader``.

.. code-block:: python

    from ooodev.loader.lo import Lo

    # get the current document
    doc = Lo.current_doc

Added ``ooodev.utils.partial.doc_io_partial.from_current_doc()`` method.
This method load a document from the current context and applies to all document classes such as a ``ooodev.write.WriteDoc`` or ``ooodev.calc.CalcDoc``.
This will also work in macros without needing to use ``ooodev.macro.MacroLoader``.

.. code-block:: python

    from ooodev.calc import CalcDoc
    doc = CalcDoc.from_current_doc()
    doc.sheets[0]["A1"].Value = "Hello World"

Version 0.25.0
==============

Added ``ooodev.utils.lo.Lo.current_doc`` static property. This property returns the current document that is being worked on such as a ``ooodev.write.WriteDoc`` or ``ooodev.calc.CalcDoc``.

Type support for a more general Document via ``ooodev.proto.office_document_t.OfficeDocumentT``. This is the type returned by ``ooodev.utils.lo.Lo.current_doc``.

Other Type enhancements and protocols.

Version 0.24.0
==============

Update for Dialogs and Multi-document support. Now Dialogs can be created from document classes such as ``ooodev.write.WriteDoc`` and ``ooodev.calc.CalcDoc``.
This ensures that the Dialog is created in the same context as the document and this supports multi-document usage.

Other minor bug fixes and updates.

Version 0.23.1
==============

Minor updates for form controls.

Version 0.23.0
==============

Document classes can now create instances of themselves and open documents.

``ooodev.Calc.CalcCellRange`` now has a ``highlight()`` method.

``ooodev.Calc.CalcCell`` now has a ``make_constraint()`` method.

Updates for event related classes.

Other Misc updates.

Version 0.22.1
==============

Added ``ooodev.write.WriteDoc.text_frames`` property. This property returns a ``ooodev.write.WriteTextFrames`` class for working with text frames.

Marked many methods in ``ooodev.office`` class as safe for multi-document usage or not. When no the ``LoContext`` manager can be used.

Better support for multi documents. Now classes ``ooodev.draw``, ``ooodev.calc`` and ``ooodev.write`` can be used with multiple documents at the same time.

Version 0.22.0
==============

Added ``ooodev.draw.ImpressPages`` class. Handles working with Impress pages via ``ooodev.Draw.ImpressDoc``.

Add a Content manager, ``ooodev.utils.context.lo_context.Locontext``. This class can be used to manage the context of a LibreOffice instance.
Now it is possible to have multiple LibreOffice document running at the same time. Implemented for ``ooodev.draw.ImpressDoc`` and ``ooodev.draw.DrawDoc``
and ``ooodev.write.WriteDoc`` so far.

Example of create two Draw documents at the same time.

.. code-block:: python

    from ooodev.draw import DrawDoc
    from ooodev.loader.lo import Lo

    # create first doc normally
    doc_first = DrawDoc.create_doc()
    doc.set_visible()

    # for a second doc create a new LoInst to open an new document with.
    lo_inst = Lo.create_lo_instance()
    # create a new DrawDoc and pass it the new instance context.
    second_doc = DrawDoc.create_doc(lo_inst=lo_inst)
    second_doc.set_visible()


Version 0.21.3
==============

Now shapes in the ``ooodev.draw.shapes`` namespace can cloned using the ``clone()`` method.

Added Create Document methods to ``WriteDoc``, ``DrawDoc``, ``ImpressDoc``.

Version 0.21.1
==============

Added LO Instance to Writer Classes. This will allow for better support of multiple Writer documents.

Implement a shape factory, ``ooodev.draw.shapes.partial.shape_factory_partial.ShapeFactoryPartial``.
Now various Draw pages can return know shapes as objects such as ``ooodev.draw.shapes.Rectangle`` and ``ooodev.draw.shapes.Ellipse``.

.. code-block:: python

    # doc is a DrawDoc instance in this case
    # The first shape added to the first slide of the document is a rectangle
    >>> shape = doc.slides[0][0]
    >>> shape.get_shape_type()
    "com.sun.star.drawing.RectangleShape"
    >>> shape
    <ooodev.draw.shapes.rectangle_shape.RectangleShape object at 0x7f9f87133ac0>


Version 0.21.0
==============

``DrawDoc`` and all of the related classes now can use a seperate instance of ``Lo`` to connect to LibreOffice.
In short this means it is now possible have mulitiple instanes of Draw Doucment open at the same time.

``DrawPage`` can now export the page as a ``png`` or ``jpg`` image using the ``export_page_png()`` and ``export_page_jpg()`` methods.
See ``tests/test_draw/test_draw_ns/test_draw_page_export_img.py`` for examples.

All Shapes in the ``ooodev.draw.shapes`` namespace now can export the shape as a ``png`` or ``jpg`` image using the ``export_shape_png()`` and ``export_shape_jpg()`` methods.

Calc Range can now export the range as a ``png`` or ``jpg`` image using the ``export_range_png()`` and ``export_range_jpg()`` methods that can alos set the image resolution.

Breaking Changes
----------------

``ooodev.events.event_data.img_export_t.ImgExportT`` has been removed. No longer needed now that ``CalcCellRange.export_png()`` and ``CalcCellRange.export_jpg()`` have been implemented.

Version 0.20.4
==============

Now ``ooodev.write.WriteTextViewCursor`` export Writer document pages as images (png or jpg) files.
See ``./tests/test_write/test_write_ns/test_export_image.py`` file for examples.

Version 0.20.3
==============

Now a Calc spreadsheet range can be exported to an image (png or jpg) file.
Exported is done via `` class.

Example of saving range as image.

.. code-block:: python

    sheet = doc.sheets[0]
    rng = sheet.get_range(range_name="A1:M4")
    rng.export_as_image("./my_image.png")

Version 0.20.2
==============

Updated ``ooodev.draw.DrawDoc``. Now has a ``save_doc`` method for saving the document.

Updated ``ooodev.draw.ImpressDoc``. Now has a ``save_doc`` method for saving the document.

Other minor bug fixes and updates.

Version 0.20.1
==============

``ooodev.calc.CalcCell`` Now has a ``value`` attribute that can get or set the value of the cell.

Breaking Changes
----------------

``ooodev.calc.CalcCell.position`` now returns :ref:`generic_unit_point` instead of a UNO ``Point``.
UNO ``Point`` can still be accessed via ``ooodev.calc.CalcCell.component.Position``.


Version 0.20.0
==============

Many new classes for working with Calc Spreadsheet view added to ``ooodev.adapter`` module.

Added ``ooodev.adapter.calc.CalcDoc.current_controller`` property.

Other minor bug fixes and updates.

Version 0.19.0
==============

``ooodev.draw.DrawPage`` now has a ``forms`` property that returns a ``ooodev.Draw.DrawForms`` class for working with and accessing forms.

Breaking Changes
----------------

``ooodev.form.control.*`` controls no longer have ``width``, ``height``, ``x``, ``y`` properties. They were not reporting the correct value from the draw page.
They can still be accessed via the controls ``ctl.get_view().getPosSize()`` method.

Now there are ``size`` and ``position`` properties that return the expected values as ``UnitMM`` objects.

Version 0.18.2
==============

Added ``ooodev.calc.SpreadsheetDrawPages`` class. Handles working with Calc Draw Pages.
Added ``ooodev.calc.SpreadsheetDrawPage`` class. Handles working with Calc Sheet Draw Page.

Added ``ooodev.calc.CalcForms`` class. Handles working with Calc Sheet Forms.
Added ``ooodev.calc.CalcForm`` class. Handles working with Calc Sheet Form.

Version 0.18.1
==============


Added ``ooodev.draw.GenericDrawPage`` class. Handles generic draw page such as ``ooodev.write.WriteDoc.get_draw_page()``.
Added ``ooodev.draw.GenericDrawPages`` class. Handles generic draw pages such as ``ooodev.write.WriteDoc.get_draw_pages()``.
Added ``ooodev.calc.SpreadsheetDrawPages`` class.
Added ``ooodev.calc.SpreadsheetDrawPage`` class.

``ooodev.calc.CalcDoc`` now have has a ``draw_pages`` property that returns a ``ooodev.calc.SpreadsheetDrawPages`` class.
``ooodev.calc.CalcSheet`` now have has a ``draw_page`` property that returns a ``ooodev.calc.SpreadsheetDrawPage`` class.

Breaking Changes
----------------

``ooodev.write.WriteDrawPage`` has been removed. Now ``ooodev.write.WriteDoc.get_draw_page()`` returns a ``ooodev.draw.GenericDrawPage[WriteDoc]``:


Version 0.18.0
==============

Now many Draw shape will accept -1 as a value for ``width``, ``height``, ``x``, ``y``.
This will usually mean that the shape size and/or position will not be set when created.

Now the Units in the ``ooodev.units`` can do math such has ``+``, ``-``, ``*``, ``/``, ``+-``, ``-+``.
Eg:

.. code-block:: python

    from ooodev.units import UnitCM, UnitInch
    u1 = UnitCM(0.44)
    u1 = += 1 # 1.44 cm
    u2 = UnitInch(2)
    u3 = u1 + u2
    assert u3 == 6.52

Version 0.17.13
===============

Added ``ooodev.draw.DrawPages`` class that is accessed via ``DrawDoc.slides`` property.

Breaking changes
----------------

``CalcDoc.get_by_index()`` Now returns a ``CalcSheet`` instance instead of ``com.sun.star.sheet.Spreadsheet`` service.
The ``CalcSheet.component`` will return the ``com.sun.star.sheet.Spreadsheet`` service.

``CalcDoc.get_by_name()`` Now returns a ``CalcSheet`` instance instead of ``com.sun.star.sheet.Spreadsheet`` service.
The ``CalcSheet.component`` will return the ``com.sun.star.sheet.Spreadsheet`` service.

Version 0.17.12
===============

Added support for modifying Draw Style Indent and Spacing.

Version 0.17.11
===============

Added ``ooodev.calc.CalcSheets`` class that is accessed via ``CalcDoc.sheets`` property.

Version 0.17.10
===============

Calc Sheets now can use ``sheet["A1"]`` to get a cell. This is a shortcut for ``sheet.get_cell("A1")``.
Any single parameter method of ``get_cell()`` can now use this shortcut such as ``cell_range``, ``cell_name``, ``cell_obj``, ``cell`` and ``addr``.

Version 0.17.9
==============

Add support for modifying Draw Style Area Image.

Add support for modifying Draw Style Area Gradient.
Add support for modifying Draw Style Area Transparency.
Add support for modifying Draw Style Font.
Add support for modifying Draw Style Font Effects.

Version 0.17.8
==============

Add support for formatting Draw Shape Text columns and Text Alignment.

Version 0.17.7
==============

Added ``get_write_text()`` to ``WriteTextCursor()`` that allows for easier access to the ``XText`` of a cursor.

Version 0.17.6
==============

Add text animation support to Draw Shapes.

Version 0.17.5
==============

Add ``get_shape_text_cursor()`` to Draw Shapes that allows for getting the text cursor of a shape.
This allows editing and formatting of the text in a shape.

More new formats for Draw Shapes.

Version 0.17.4
==============

More new formats for Draw Shapes.

Version 0.17.3
==============

Added new formats for Draw Shapes.

Version 0.17.2
==============

Fix to allow ``com.sun.star.presentation.Shape`` as a ShapeComp.

Version 0.17.1
==============

Added support for ``LineCursor`` and ``ScreenCursor`` on ``ooodev.write.WriteTextViewCursor``.

Version 0.17.0
==============

Added ``ooodev.draw`` module. This module contains classes for working with Draw and Impress documents.
Many new classes which make working with Draw and Impress documents much easier.

Version 0.16.0
==============

Added ``ooodev.write`` module. This module contains classes for working with Writer documents.
Many new classes which make working with Writer documents much easier.

Version 0.15.1
==============

Extended ``ooodev.calc`` classes with new methods

Version 0.15.0
==============

Added ``ooodev.calc`` Which contains classes for working with Calc documents.
Now Calc documents are much easier to work with.

Version 0.14.2
==============

Updates for ``Forms`` and ``Dialogs``.

Version 0.14.1
==============

Minor updates for ``Forms`` and ``Dialogs``.

Version 0.14.0
==============

Added Form Controls and Form Database Controls. More then 30 new classes for working with forms.

Add several new classes in the ``adapter`` module.

Other minor bug fixes and updates.

Version 0.13.7
==============

Added Form Controls and Form Database Controls

Added many new classes in the ``adapter`` module.

Renamed ``StyleObj`` to ``StyleT``

Renamed ``UnitObj`` to ``UnitT``

Version 0.13.7
==============

Added dozens of new classes in the ``adapter`` module.

Version 0.13.6
==============

Added subscriber to constructor of many classes in the ``adapter`` module.

Version 0.13.5
==============

Event classes now implement dispose method in the ``adapter`` module classes.

Version 0.13.4
==============

New options for event classes in the ``adapter`` module.

Version 0.13.3
==============

Update to ``CtlTree`` for better flat list loading of data.

Other minor bug fixes and updates.

Version 0.13.2
==============

Add new properties to several Dialog control classes.

Add new classes in ``adapters`` module.

Version 0.13.0
==============

Dialog Module added. Many new classes for working with dialogs.
Many new adapters added into the adapter module.

Other minor bug fixes and updates.

Version 0.12.1
==============

Doc updates, minor bug fixes and updates.

Add guide for installing OooDev as a LibreOffice `Extension <https://github.com/Amourspirit/libreoffice_ooodev_ext/tree/main>`__.

Version 0.12.0
==============

This version saw the removal of ``lxml`` as a dependency. Now the Library has no external binary dependencies.

The ``ooodev.utils.xml_util`` module was removed and all methods were moved to `Ooo Dev Xml <https://pypi.org/project/ooo-dev-xml/>`__ package.

If you were using the class directly from the ``ooodev.utils.xml_util`` module, you can now use the class from the ``ooodev_xml.odxml`` module.

Version 0.11.14
===============

Added ``FileIO.expand_macro()`` method that can be used to expand macro paths.

Version 0.11.13
===============

Updates for better support of ``Lo.this_component`` in and ``Lo.XSCRIPTCONTEXT``.

Version 0.11.12
===============

Fix bug in ``Calc.set_sheet_name()`` that was not working correctly.

Add new parameter to ``Calc.get_sheet_name()`` that allows for wrapping of the sheet name in single quotes if it is needed.

Version 0.11.11
===============

Now there is a context manager for macros that set the proper context for the document and |odev|.
See :ref:`ch02_macro_load`.

Version 0.11.10
===============

Updated connection to LibreOffice be more robust. Remote connections have been tested and work.

Version 0.11.9
==============

Fix for potential bug when connection to LibreOffice instance.

Version 0.11.8
==============

Update to allow connections to LibreOffice Snap and Flatpak versions on Linux.

Version 0.11.7
==============

Added ``env_vars`` options to Bridge base connectors. Now Environment variables can be passed to the subprocess that connects to LibreOffice.
This makes it possible to connect to a snap instance of LibreOffice and pass in ``PYTHONPATH`` and other environment variables.

Version 0.11.6
==============

Add environment check to ``ooodev.utils.paths.get_soffice_path`` to  ``ODEV_CONN_SOFFICE`` environment variable is set to LibreOffice soffice.

Update for better support of ``Lo.this_component`` in macros.

Version 0.11.5
==============

Remove unused module ``ooodev.utils.images``.

Remove unused dependency ``Pillow``.

Version 0.11.4
==============

Fix for ``Write.get_cursor()`` not working correctly in Snap version of LibreOffice in macros.

Version 0.11.3
==============

Fix for ``Lo.this_component`` in macros.

Version 0.11.2
==============

Added ``Calc.print_sheet()`` method that allows printing of a specified cell range directly to a printer.

Version 0.11.1
==============

Minor tweaks and dependency updates.

Version 0.11.0
==============

Major Refactoring of entire Library. Much improved typing support.

This version now has complete type support. Entire code base has been refactored to support type hints and type checking.

Test have been preformed with ``pyright`` to ensure type hints are correct.

Version 0.10.3
==============

Minor bug fixes and updates.

Version 0.10.2
==============

Fix for Chart2 Gradient Fill class.

Version 0.10.1
==============

Minor style bug fixes.


Version 0.10.0
==============

Support dropped for Python ``3.7``. Now supporting Python ``3.8`` and up.

Added Calc methods ``get_cell_protection()``, ``is_cell_protected()``, ``protect_sheet()``, ``unprotect_sheet()`` and ``is_sheet_protected()``. 

Other Minor Calc tweaks.

Version 0.9.8
=============

Created :ref:`ns_inst_lo` that also contains ``LoInst`` class. This class can create a new LibreOffice instance and connect to it and/or
connect to an existing LibreOffice instance and be used for sub-components. This class is for advanced usage.
The ``Lo`` class is still the recommended way to connect to LibreOffice and under the hood it uses ``LoInst``.
See :ref:`ch02_multiple_docs`.


Version 0.9.7
=============

Fix ``ooodev.utils.data_type.size_mm.SizeMM`` constructor to accept ``UnitObj`` as well as ``float``.

Minor bug fixes and updates.

Version 0.9.6
=============

Minor bug fixes and updates.

Version 0.9.5
=============

Minor bug fixes and updates.

Version 0.9.4
=============

Added more than five dozen new classes in ``ooodev.format.chart2.direct`` for formatting ``Chart2`` charts.

Added ``ooodev.office.chart2.Chart2ControllerLock`` class that can be used to lock and unlock ``Chart2`` charts for faster updating.

Added ``ooodev.format.calc.direct.cell.numbers.Numbers`` class that can be used to format numbers styles in ``Calc`` cells and ranges.

Added new event to ``ooodev.utils.props.Props.set()``. Now subscribers can be notified when a property set error occurs and handle the error if needed.

Added new event to ``ooodev.utils.props.Props.set_default()``. Now subscribers can be notified when a property set default error occurs and handle the error if needed.

Version 0.9.3
=============

Minor revisions and updates.

Version 0.9.2
=============

Added style options to ``from ooodev.utils.forms.Forms`` module methods.

Added ``Write.create_style_para()`` that creates new paragraph styles and adds the them to the document paragraph styles.

Added ``Write.create_style_char()`` that creates new character styles and adds the them to the document character styles.

Added ``Write.set_footer()`` that sets the footer text and style.

Added ``styles`` to ``Write.set_header()`` that also sets the header style.

Added ``ooodev.format.writer.direct.page`` module that contains classes for page header and footer styles that can be
used with ``Write.set_header()`` and ``Write.set_footer()``.

Version 0.9.1
=============

Added :ref:`ns_theme` that access LibreOffice theme properties.

Added ``Info.get_office_theme()`` That gets the current LibreOffice theme name.

Added overloads to several Calc methods to allow styles to be applied when setting sheet values.


Version 0.9.0
=============

Added :ref:`ns_format` module with hundreds of new classes for applying styles and formatting to documents and sheets.

Added :ref:`ns_units` module that contains classes for many of the LibreOffice units such as ``mm`` units, ``px`` units and ``pt`` units (and more).

Renamed method ``GUI.show_memu_bar()`` to ``GUI.show_menu_bar()``

Fixed issue with ``Calc.get_sheet_names()`` when overload with no args was used.

Rename ``CellObj.col_info`` to ``CellObj.col_obj``

Rename ``CellObj.row_info`` to ``CellObj.row_obj``

All events now can have key value pairs of data added or removed

Added ``Calc.get_safe_rng_str()`` method.

Added ``Info.is_uno()`` method.

Added ``Write.style()`` method.

Added ``Write.get_cursor()`` overload.

Added ``Write.append(cursor: XTextCursor, text: str, styles: Iterable[StyleObj])`` overload.

Added ``Write.style_left(cursor: XTextCursor, pos: int, styles: Iterable[StyleObj])`` overload.

Added ``Write.style_prev_paragraph(cursor: XTextCursor, styles: Iterable[StyleObj])`` overload.

Added ``Write.append_line(cursor: XTextCursor, text: str, styles: Iterable[StyleObj])`` overload.

Added ``Write.append_para(cursor: XTextCursor, text: str, styles: Iterable[StyleObj])`` overload.

Added ``Chart2.style_background()`` Method.

Added ``Chart2.style_wall()`` Method.

Added ``Chart2.style_data_point()`` Method.

Version 0.8.6
=============

Added Styles namespace.

Extended Cell Objects with focus on ``CellValues`` Class.

Added overload to ``GUI.set_visible()``

Added overload to ``GUI.set_visible()``

Added overload to ``Calc.get_sheet_names()``

Added overload to ``Calc.set_sheet_name()``

Changed ``Calc.get_sheet(doc: XSpreadsheetDocument, index: int)`` to ``Calc.get_sheet(doc: XSpreadsheetDocument, idx: int)``.
``index`` will still work but is not documented.

Changed ``Calc.remove_sheet(doc: XSpreadsheetDocument, index: int)`` to ``Calc.remove_sheet(doc: XSpreadsheetDocument, idx: int)``
``index`` will still work but is not documented.

Version 0.8.5
=============

Fix for Some ``Calc`` related method getting a new doc with the existing doc was expected.

Version 0.8.4
=============

Added methods, ``Calc.merge_cells()``, ``Calc.unmerge_cells()``, and ``Calc.is_merged_cells()``

Version 0.8.3
=============

Many new Overloads in ``Calc`` module for range objects.

Several enhancements for range objects.

Version 0.8.2
=============

Added ``Calc.is_single_column_range()``.

Added ``Calc.is_single_row_range()``.

Added ``Calc.get_range_size()``

Added ``Calc.get_range_obj()``

Added ``Calc.get_selected_range()``

Added ``Calc.get_selected_cell()``

Many enhancements for working with sheet ranges.

Version 0.8.1
=============

``Chart2.insert_chart()`` all parameters made optional, added ``chart_name`` parameter.

Added ``Chart2.remove_chart()``.

Added ``Calc.set_selected_addr()``.

Updated ``Angle`` to accept any integer value, positive or negative.

Version 0.8.0
=============

Added ``Calc.get_col_first_used_index()`` method.

Added ``Calc.get_col_last_used_index()`` method.

Added ``Calc.get_row_first_used_index()`` method.

Added ``Calc.get_row_last_used_index()`` method.

Added overloads to ``Calc.get_col()``.

Added overloads to ``Calc.get_row()``.

``Calc.get_col()`` now returns an empty list like ``Calc.get_row()`` if no values are found.
In previous version it it returned ``None`` When no values were found.

``Calc.extract_col()`` now returns an empty list if no values are found.
In previous version it it returned ``None`` When no values were found.

Version 0.7.1
=============

Minor updates to ``chart2_types`` module.

Version 0.7.0
=============

Added ``Lo.loader_current``. Now after ``Lo.load_office()`` is called the ``Lo.loader_current`` property will contain the same loader that is returned by ``Lo.load_office()``

All methods that are using ``loader`` now have a overload to make ``loader`` optional.

``Calc.open_doc()`` has new overloads. Now if a file is not passed to open then a new spreadsheet document is returned.

``Write.open_doc()`` has new overloads. Now if a file is not passed to open then a new Writer document is returned.

Version 0.6.10
==============

Now ``Lo.load_load()`` has extra options that allow for turning on or off of verbose via the loader.
Going forward verbose is off by default.

Added overload to ``Calc.get_sheet()``

Update ``Props.show_props()`` to support extra formatting.

Fix bug in ``Calc.get_function_names()``

Removed unnecessary events from

.. cssclass:: ul-list

    - ``Calc.print_addresses()``
    - ``Calc.print_array()``
    - ``Calc.print_cell_address()``
    - ``Calc.print_fun_arguments()``
    - ``Calc.print_function_info()``
    - ``Calc.print_head_foot.print_address``
    - ``Calc.print_head_foot``

Version 0.6.9
=============

Added ``FileIO.uri_absolute()``

Added overload to ``props.get()``.

``FileIO.uri_to_path()`` now raises ``ConvertPathError`` if unable to convert.

Added an enum lookup option to ``Info.get_paths()``.

Added ``utils.Gallery`` module.

Version 0.6.8
=============

Added ``utils.adapter`` namespace and classes.

Version 0.6.7
=============

Add new methods ``convert_1d_to_2d``, ``get_smallest_str``, ``get_largest_str``, ``get_smallest_int``, ``get_largest_int`` to ``TableHelper`` Class.

Added overload method ``Lo.print_table(name: str, table: Table, format_opt: FormatterTable)``

Updated ``Lo.print_names()`` to print output in a table format.

Version 0.6.6
=============

Add overload to ``Calc.convert_to_floats``

Add ``formatters`` module for formatting console output.

Added overload method ``Calc.print_array(vals: Table, format_opt: FormatterTable)``

Version 0.6.5
=============

Added overload to ``FileIo.make_directory`` that handles creating directory from file path.

Fix for ``FileIo.url_to_path`` on windows sometimes not converting correctly.

Other ``FileIo`` Minor updates.

Fix bug in ``Chart2.set_template`` when ``diagram_name`` was passed as string.

Fix bug in ``Draw.warns_position`` when no Slide size is available.

Renamed ``Calc.get_range_str`` args from ``start_col``, ``start_row``, ``end_col``, ``end_row`` to ``col_start``, ``row_start``, ``col_end``, ``row_end`` respectively.
Change is backwards compatible.

Renamed ``Calc.get_cell_range`` args from ``start_col``, ``start_row``, ``end_col``, ``end_row`` to ``col_start``, ``row_start``, ``col_end``, ``row_end`` respectively.
Change is backwards compatible.

Version 0.6.4
=============

Fix for ``Draw.report_pos_size``. Now handles when a shape does not have a ``Name`` property an other errors.

Version 0.6.3
=============

Overloads for ``GUI.get_window_handle()``

Removed unused ``*titles`` arg from ``Draw.add_dispatch_shape()`` method.

Removed unused ``*titles`` arg from ``Draw.create_dispatch_shape()`` method.

``GUI.get_title_bar()`` method now returns empty string when not able to get title bar text.

Version 0.6.2
=============

Rename private enum ``_LayoutKind`` to public ``LayoutKind`` to make available for public use.

Added new Fast Lookup methods to ``Props`` class.

New Exceptions ``PropertyGeneralError``

Version 0.6.1
=============

Added ``Draw.add_dispatch_shape()`` method.

Added ``Draw.create_dispatch_shape()`` method.

Added Dispatch Lookup ``ShapeDispatchKind`` Enum.

Added None to ``GraphicArrowStyleKind`` Enum.

Added classes ``WindowTitle`` and ``DialogTitle`` for working with GUI packages.

Version 0.6.0
=============

Breaking changes.

``Write.ControlCharacter`` was an alias of ``ooo.dyn.text.control_character.ControlCharacterEnum``.
Now ``ControlCharacterEnum`` must be used instead of ``Write.ControlCharacter``.
``ControlCharacterEnum`` can be imported from ``Write``.
:abbreviation:`e.g.` ``from ooodev.office.write import Write, ControlCharacterEnum``

``Write.DictionaryType`` was an alias of ``ooo.dyn.linguistic2.dictionary_type.DictionaryType``.
Now ``DictionaryType`` must be used instead of ``Write.DictionaryType``.
``DictionaryType`` can be imported from ``Write``.
:abbreviation:`e.g.` ``from ooodev.office.write import Write, DictionaryType``

``Calc.CellFlags`` was an alias of ``ooo.dyn.sheet.cell_flags.CellFlagsEnum``.
Now ``CellFlagsEnum`` must be used instead of ``Calc.CellFlags``.
``CellFlagsEnum`` can be imported from ``Calc``.
:abbreviation:`e.g.` ``from ooodev.office.calc import Calc, CellFlagsEnum``

``Calc.GeneralFunction`` was an alias of ``ooo.dyn.sheet.general_function.GeneralFunction``.
Now ``GeneralFunction`` must be used instead of ``Calc.GeneralFunction``.
``GeneralFunction`` can be imported from ``Calc``.
:abbreviation:`e.g.` ``from ooodev.office.calc import Calc, GeneralFunction``

``Calc.SolverConstraintOperator`` was an alias of ``ooo.dyn.sheet.solver_constraint_operator.SolverConstraintOperator``.
Now ``SolverConstraintOperator`` must be used instead of ``Calc.SolverConstraintOperator``.
``SolverConstraintOperator`` can be imported from ``Calc``.
:abbreviation:`e.g.` ``from ooodev.office.calc import Calc, SolverConstraintOperator``


``Calc.FillDateMode`` was an alias of ``ooo.dyn.sheet.fill_date_mode.FillDateMode``.
Now ``FillDateMode`` must be used instead of ``Calc.FillDateMode``.
``FillDateMode`` can be imported from ``Calc``.
:abbreviation:`e.g.` ``from ooodev.office.calc import Calc, FillDateMode``

Version 0.5.3
=============

``Lo.dispatch_cmd`` Now returns the result of the dispatch command if any.
Formerly a ``bool`` was returned.

``Lo.dispatch_cmd`` Now raises ``DispatchError`` if an error occurs.

Version 0.5.2
=============

Chart Samples and tests

Misc code tweaks.

Version 0.5.1
=============

Chart 2 Samples and tests.

Version 0.5.0
=============

New modules

- Draw
- Chart
- Chart2

Added ``utils.dispatch`` which as several new classes for looking up dispatch values.

Misc bug fixes.

Version 0.4.19
==============

Fix bug in setup.py

Version 0.4.17
==============

Update to Write:

- new method ``split_paragraph_into_sentences``
- new overloads for ``print_meaning``
- new overloads for ``print_services_info``
- new overloads for ``proof_sentence``
- new overloads for ``spell_sentence``
- new overloads for ``spell_word``
- ``load_spell_checker`` now load spell checker from ``com.sun.star.linguistic2.SpellChecker``


Version 0.4.16
==============

Fixes for Write spell checking


Version 0.4.15
==============

Update Graphic methods to move away from ``GraphicURL``

Other minor bug fixes.

Version 0.4.14
==============

Minor fix in ``Write.set_page_numbers``

Version 0.4.13
==============

Fix for  ``Write.add_text_frame()`` events.

Version 0.4.12
==============

Add defaults for cfg in case config.json is not available.

Version 0.4.11
==============

Fix bug in ``Lo.print_names()``

Remove internal events from some print functions that should not have had them.

Fix bug that did copy config.json during setup.

Version 0.4.10
==============

Add new event_source property to internal event classes.

Version 0.4.9
=============

| Added a Bridge Connector :py:attr:`.Lo.bridge`
| See also: :ref:`ch04_bridge_stop`
| See example: `Office Window Monitor <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_monitor>`_

Added Session class for registering and importing.
See example: `Shared Library Access <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_share_lib>`_

Version 0.4.8
=============

New listeners in ooodev.listeners namespace

Fix For Lo.XSCRIPTCONTEXT

Version 0.4.7
=============

Added ``minimize()``, ``maximize()`` and ``activate()`` methods to :py:class:`~.gui.GUI` class.

Version 0.4.6
=============

Updates and fixes for :py:class:`~.utils.info.Info` class.


Version 0.4.5
=============

Added :py:class:`~.break_context.BreakContext` class.

Version 0.4.4
=============

Bug fix reading document properties.

Version 0.4.2
=============

Fix bug in windows connections

Version 0.4.1
=============

Fix bug in :py:class:`~.utils.info.Info`.
Some methods were expecting string but got Path object.

Version 0.4.0
=============

New more flexible and robust way of connecting to office.

This update change :py:meth:`.Lo.load_office` method

Paths used internally now automatically resolve to absolute paths.

Version 0.3.0
=============

Write module released

Version 0.2.0
=============

Initial release with full support for calc.