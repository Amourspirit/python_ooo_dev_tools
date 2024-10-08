.. _ch02:

********************************
Chapter 2. Starting and Stopping
********************************

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

.. topic:: Starting Office

    Starting Office; Closing Down/Killing Office; Opening a Document; Creating a Document; Saving; Closing; Document Conversion; Bug Detection and Reporting

:ref:`ch01` introduced some of the core ideas of Office.
Now it's time to show how these data structures and relationships (e.g. service, interfaces, FCM, inheritance)
are programmed in |app_name_bold| (|odev|) API.

This chapter will focus on the most fundamental tasks: starting Office,
loading (or creating) a document, saving and closing the document, and shutting down Office.

Some the examples come from the `LibreOffice Python UNO Examples <https://github.com/Amourspirit/python-ooouno-ex>`_ project.

The aim of |odev| is to hide some of the verbiage of Office.
When (if?) a programmer feels ready for more detail, then my code is documented.
Here, only explaining functions that illustrate Office ideas, such as service managers and components.

.. todo:: 

    Link to sections 8

This is the first chapter with code, and so the first where programs could crash! Section 8 gives a few tips on bug detection and reporting.

.. _ch02_starting_office:

2.1 Starting Office
===================

.. _ch02_1_running_as_macro:

2.1.1 Introduction
------------------

Every program must load Office before working with a document (unless run in a macro), and shut it down before exiting.
These tasks are handled by :py:meth:`.Lo.load_office` and :py:meth:`.Lo.close_office` from the :py:class:`.Lo`.
A typical program will look like the following:

.. tabs::

    .. code-tab:: python

        from ooodev.loader.lo import Lo

        def main() -> None:
            loader = Lo.load_office(Lo.ConnectSocket(headless=True)) # XComponentLoader

            # load, manipulate and close a document

            Lo.close_office()

        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``Lo.load_office(Lo.ConnectSocket(headless=True))`` invokes Office and sets up a UNO bridge using named pipes with a headless connection.
If not using the Graphic User Interface (GUI) of LibreOffice then ``headless=True`` is recommended.
See example |extract_graphics_py|_.

There is also ``Lo.load_office(Lo.ConnectPipes(headless=True))`` which uses which uses pipes instead of sockets.
See example |extract_txt_py|_.

For convenience ``Lo.ConnectPipe`` is an alias of :py:class:`~.conn.connectors.ConnectPipe`
and ``Lo.ConnectSocket`` is an alias of :py:class:`~.conn.connectors.ConnectSocket`.

In both cases, a remote component context is created (see Chapter 1, :numref:`ch01fig_python_using_office`) and then a service manager,
Desktop object, and component loader are initialized.
Below is a simplified version of :py:meth:`.Lo.load_office`, that show the principle of connecting to LibreOffice.

.. tabs::

    .. code-tab:: python

        # in Lo class (simplified)
        @classmethod
        def load_office(
            cls, connector: ConnectPipe | ConnectSocket | None = None, cache_obj: Cache | None = None
        ) -> XComponentLoader:
    
            Lo.print("Loading Office...")
            if connector is None:
                try:
                    cls._lo_inst = LoDirectStart()
                    cls._lo_inst.connect()
                except Exception as e:
                    Lo.print((
                        "Office context could not be created."
                        " A connector must be supplied if not running as a macro"
                    ))
                    Lo.print(f"    {e}")
                    raise SystemExit(1)
            elif isinstance(connector, ConnectPipe):
                try:
                    cls._lo_inst = LoPipeStart(connector=connector, cache_obj=cache_obj)
                    cls._lo_inst.connect()
                except Exception as e:
                    Lo.print("Office context could not be created")
                    Lo.print(f"    {e}")
                    raise SystemExit(1)
            elif isinstance(connector, ConnectSocket):
                try:
                    cls._lo_inst = LoSocketStart(connector=connector, cache_obj=cache_obj)
                    cls._lo_inst.connect()
                except Exception as e:
                    Lo.print("Office context could not be created")
                    Lo.print(f"    {e}")
                    raise SystemExit(1)
            else:
                Lo.print("Invalid Connector type. Fatal Error.")
                raise SystemExit(1)

            cls._xcc = cls._lo_inst.ctx
            cls._mc_factory = cls._xcc.getServiceManager()
            if cls._mc_factory is None:
                Lo.print("Office Service Manager is unavailable")
                raise SystemExit(1)
            cls._xdesktop = cls.create_instance_mcf(XDesktop, "com.sun.star.frame.Desktop")
            if cls._xdesktop is None:
                Lo.print("Could not create a desktop service")
                raise SystemExit(1)
            loader = cls.qi(XComponentLoader, cls._xdesktop)
            if loader is None:
                Lo.print("Unable to access XComponentLoader")
                SystemExit(1)
            return loader

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        - :odev_src_loinst_meth:`load_office`


There is also :py:class:`.Lo.Loader` context manager that allows for automatic closing of office.
See |convert_doc|_ for an example.


It is also simple to start LibreOffice from the command line automate tasks and leave it open for user input.
See `Calc Add Range of Data Automation <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_add_range_data>`_ for an example.

.. _ch02_1_1_running_as_macro:

2.1.2 Running as a Macro
------------------------

There is of course running as a macro as well.

Running a project with several modules can be a bit daunting task.
For this reason oooscript_ was created, which can easily pack several scripts into one script and embed the result into a LibreOffice Document.

The easiest way to run a several module/class project in LibreOffice is to pack into a single script first.
Many examples can be found on |lo_uno_ex|_,
such as |lo_uno_calc_ex|_, |lo_uno_sudoku|_, |lo_uno_tab_control|_.

.. seealso::

    :py:class:`~.utils.session.Session` and |lo_uno_shared_acc|_

Macros only need use use ``Lo.ThisComponent`` as show below.

.. tabs::

    .. code-tab:: python

        from ooodev.loader.lo import Lo
        from ooodev.office.calc import Calc

        def main():
            # get access to current Calc Document
            doc = Calc.get_ss_doc(Lo.ThisComponent)

            # get access to current spreadsheet
            sheet = Calc.get_active_sheet(doc=doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch02_1_3_static_classes:

2.1.3 Using Static classes
--------------------------

:py:meth:`.Lo.load_office` probably illustrates a significant coding decisions – the use of global static variables inside the :py:class:`~.lo.Lo` class.
In particular, the XComponentContext_, XDesktop_, and XMultiComponentFactory_ objects created by :py:meth:`~.lo.Lo.load_office` are stored globally for later use.
This approach is chosen since it allows other support functions to be called with simpler arguments because the objects can be accessed without
the user having to explicitly pass around references to them.
The main drawback is that if ``load_office()`` is called more than once all previous :py:class:`~.lo.Lo` class globals are overwritten.
See :ref:`ch02_multiple_docs` for a solution to this problem.

The creation of the XDesktop_ interface object uses :py:meth:`.Lo.create_instance_mcf`:

.. tabs::

    .. code-tab:: python

        # in Lo class
        @classmethod
        def create_instance_mcf(
            cls,
            atype: Type[T], service_name: str,
            args: Tuple[Any, ...] | None = None,
            raise_err: bool = False
        ) -> T | None:

            if cls._xcc is None or cls._mc_factory is None:
                raise Exception("No office connection found")
            try:
                if args is not None:
                    obj = cls._mc_factory.createInstanceWithArgumentsAndContext(
                        service_name, args, cls._xcc
                    )
                else:
                    obj = cls._mc_factory.createInstanceWithContext(service_name, cls._xcc)
                if raise_err is True and obj is None:
                    CreateInstanceMcfError(atype, service_name)
                interface_obj = cls.qi(atype=atype, obj=obj)
                if raise_err is True and interface_obj is None:
                    raise MissingInterfaceError(atype)
                return interface_obj
            except CreateInstanceMcfError:
                raise
            except MissingInterfaceError:
                raise
            except Exception as e:
                raise Exception(f"Couldn't create interface for '{service_name}'") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

If you ignore the error-checking, :py:meth:`.Lo.create_instance_mcf` does two things.
The call to `XMultiComponentFactory.createInstanceWithContext() <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XMultiComponentFactory.html#ac62a80213fcf269e7a881abc6fa3e6d2>`_
asks the service manager (``cls._mc_factory``) to create a service object inside the remote component context (``cls._xcc``). Then the call to ``uno_obj.queryInterface()``
via :py:meth:`.Lo.qi` looks inside the service instance for the specified interface (``atype``), returning an instance of the interface as its result.

The :py:meth:`.Lo.qi` function's reduces programmer typing, since calls to ``uno_obj.queryInterface()`` are very common in this frame work.
Querying for the interface has the huge advantage of providing typing :numref:`ch02fig_lo_qi_auto_demo` (autocomplete, static type checking) support thanks to types-unopy_.

.. cssclass:: rst-collapse bg-transparent

    .. collapse:: Demo
        :open:

        .. cssclass:: a_gif

            .. _ch02fig_lo_qi_auto_demo:
            .. figure:: https://user-images.githubusercontent.com/4193389/178285134-70b9aa56-5eaa-43c8-aa59-c19f2b495336.gif
                :alt: Lo.qi autocomplete demo image
                :figclass: align-center

                : :py:meth:`.Lo.qi` autocomplete demo


The use of generics makes :py:meth:`.Lo.create_instance_mcf` useful for creating any type of interface object.
Unfortunately, generics aren't utilized in the Office API, which relies instead on Object, Office's Any class, or the XInterface class which is inherited by all interfaces.

.. _ch02_1_cache_options:

2.1.4 Cache Options
-------------------

The :py:meth:`.Lo.load_office` method has an optional ``cache_obj`` parameter.
This parameter is used to cache the control the environment used to connect to LibreOffice.

When using the :py:meth:`.Lo.load_office` method with the ``cache_obj`` parameter, the cache object is used to store the connection environment.
The ``cache_obj`` parameter is an instance of the :py:class:`~.conn.cache.Cache` class.

What is the purpose of the cache object?

The cache object is used to copy a profile director to a temporary directory or pre-defined directory.
The cache object is used to set an empty profile directory, which is useful when you want to start LibreOffice with a clean profile.
The cache object can be used to disable shared extension.

This can be very useful when you are running LibreOffice in a headless mode and there is an extension that has a dialog popup for a new profile.
I found this to be the case with the APSO extension. On one environment the extension was installed for Shared Users.
When running LibreOffice in headless mode on a new profile the extension would try popup a dialog. This would result in LibreOffice hanging headless mode.
The solution was to disable the shared extension with the cache object.

How to disable shared extensions?

In this example the shares extensions are disabled by setting the ``no_shared_ext`` parameter to ``True``.
Under the hood the cache object will set environment var ``UNO_SHARED_PACKAGES_CACHE`` for LibreOffice to an empty temporary directory.

.. tabs::

    .. code-tab:: python

        from ooodev.loader.lo import Lo
        from ooodev.conn.cache import Cache

        def main():
            with Lo.Loader(
                Lo.ConnectSocket(),
                cache_obj=Cache(no_shared_ext=True),
            ) as loader:
                # do work and then loader will be closed automatically.
                pass


        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

How to copy default profile directory to a temporary directory?

In this example by creating an instance of the :py:class:`~.conn.cache.Cache` class and not setting any parameters the users profile directory (if found) is copied to a temporary directory.
Then when LibreOffice is started its profile directory is set to the temporary directory. This gives a way to make changes for the session without affecting the users main profile.

.. tabs::

    .. code-tab:: python

        from ooodev.loader.lo import Lo
        from ooodev.conn.cache import Cache

        def main():
            with Lo.Loader(
                Lo.ConnectSocket(headless=True),
                cache_obj=Cache(),
            ) as loader:
                # do work and then loader will be closed automatically.
                pass


        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


How to copy custom profile directory to a temporary directory?

By default the profile is copied to a temporary directory. If you want to copy a custom profile to be copied then you can set the profile path.
On Linux for example the profile path is typically ``~/.config/libreoffice/4``.

With this example the profile directory is copied to a temporary directory from ``/path/to/profile`` and used when LibreOffice is started.


.. tabs::

    .. code-tab:: python

        from ooodev.loader.lo import Lo
        from ooodev.conn.cache import Cache

        def main():
            with Lo.Loader(
                Lo.ConnectSocket(headless=True),
                cache_obj=Cache(profile_path="/path/to/profile"),
            ) as loader:
                # do work and then loader will be closed automatically.
                pass


        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

How to use a clean profile directory?

If the ``profile_path`` is set to an empty string then a clean profile is generated and used. This is favorable in most cases.


.. tabs::

    .. code-tab:: python

        from ooodev.loader.lo import Lo
        from ooodev.conn.cache import Cache

        def main():
            with Lo.Loader(
                Lo.ConnectSocket(headless=True),
                cache_obj=Cache(profile_path=""),
            ) as loader:
                # do work and then loader will be closed automatically.
                pass


        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


What is the recommended setting for headless?

Setting the ``profile_path`` to an empty string and ``no_shared_ext`` to ``True`` is the recommended setting for headless mode.
This will disable shared extensions and create a clean profile directory.


.. tabs::

    .. code-tab:: python

        from ooodev.loader.lo import Lo
        from ooodev.conn.cache import Cache

        def main():
            with Lo.Loader(
                Lo.ConnectSocket(headless=True),
                cache_obj=Cache(profile_path="", no_shared_ext=True),
            ) as loader:
                # do work and then loader will be closed automatically.
                pass


        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. _ch02_clossing_office:

2.2 Closing Down/Killing Office
===============================

:py:meth:`Lo.close_office` shuts down Office by calling ``terminate()`` on the XDesktop_ instance created inside :py:meth:`.Lo.load_office`:
``boolean isDead = xDesktop.terminate()`` This is usually sufficient but occasionally it necessary to delay the ``terminate()`` call for a
few milliseconds in order to give Office components time to finish.
As a consequence, :py:meth:`Lo.close_office` may actually call ``terminate()`` a few times, until it returns ``True``.


While developing/debugging code, it's quite easy to inadvertently trigger a runtime exception in the Office API.
In the worst case, this can cause your program to exit without calling :py:meth:`Lo.close_office`.
This will leave an extraneous Office process running in the OS, which should be killed. The easiest way is with |dsearch|_
``loproc --kill``. As of version `0.47.19` internally |odev| watches the subprocess that LibreOffice is started on and terminates it when the main process is terminated.

.. _ch02_2_context_manager:

Using a Context Manager
-----------------------

When using the context manager it is not necessary to call :py:meth:`Lo.close_office`. This context manager will take care of closing office.

.. tabs::

    .. code-tab:: python

        from ooodev.loader.lo import Lo
        from ooodev.calc import CalcDoc


        def main():
            with Lo.Loader(Lo.ConnectPipe()) as loader:
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                sheet["A1"].value = 10
                doc.msgbox("All done")
                doc.close()


        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. _ch02_open_doc:

2.3 Opening a Document
======================

.. _ch02_3_open_doc_using_lo:


2.3.1 Using Lo Methods
----------------------

The general format of a program that opens a document, manipulates it in some way, and then saves it, is:

.. tabs::

    .. code-tab:: python
    
        def main() -> None:
            fnm = sys.argv[1:] # get file from first args

            loader = Lo.load_office(Lo.ConnectSocket(headless=True))
            doc = Lo.open_doc(fnm=fnm, loader=loader)

            # use the Office API to manipulate doc...
            Lo.save_doc(doc, "foo.docx") # save as a Word file
            Lo.close_doc(doc)
            lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

See |convert_doc|_ for an example.

The new methods are :py:meth:`.Lo.open_doc`, :py:meth:`.Lo.save_doc`, and :py:meth:`.Lo.close_doc`.

:py:meth:`.Lo.open_doc` calls `XComponentLoader.loadComponentFromURL() <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XComponentLoader.html#a6f89db7d45da267af47d1acf01cd986d>`_,
which requires a document URL, the type of Office frame used to display the document, optional search flags, and an array of document properties.

For example:

.. tabs::

    .. code-tab:: python

        file_url = FileIO.fnm_to_url(fnm)
        props = Props.make_props(Hidden=True)
        doc = loader.loadComponentFromURL(file_url, "_blank", 0, props)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The frame type is almost always "_blank" which indicates that a new window will be created for the newly loaded document.
(Other possibilities are listed in the XComponentLoader_ documentation which you can access with ``lodoc XComponentLoader``.)
The search flags are usually set to 0, and document properties are stored in the PropertyValue_ tuple, props.

``loadComponentFromURL()``'s return type is XComponent_, which refers to the document.

:py:meth:`.FileIO.fnm_to_url` converts an ordinary filename (e.g. “foo.doc”) into a URL (a full path prefixed with ``file:///``).

:py:meth:`.Props.make_props` takes a property name and value and returns a PropertyValue_ tuple; there are several variants which accept different numbers of property name - value pairs.

A complete list of document properties can be found in the MediaDescriptor documentation (accessed with ``lodoc MediaDescriptor service``),
but some of the important ones are listed in :numref:`ch02tbl_some_doc_prop`

.. _ch02tbl_some_doc_prop:

.. table:: Some Document Properties.
    :name: md_common_srv

    ==================== =============================================================================
    Property Name        Use                                                                          
    ==================== =============================================================================
    AsTemplate           Creates a new document using a specified template                            
    Hidden               Determines if the document is invisible after being loaded                   
    ReadOnly             Opens the document read-only                                                 
    StartPresentation    Starts showing a slide presentation immediately after loading the document   
    ==================== =============================================================================

.. _ch02_3_open_doc_using_classes:

2.3.2 Using the Classes such as CalcDoc
---------------------------------------

The :py:class:`~.calc.CalcDoc`, :py:class:`~.write.WriteDoc`, :py:class:`~.draw.DrawDoc`, :py:class:`~.draw.ImpressDoc` classes can be used to open a document.
These classes all implement the same methods for creating, opening, saving, and closing a documents.

For this example we will use the :py:class:`~.calc.CalcDoc` class to open a Calc document.


.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        from ooodev.loader.lo import Lo
        from ooodev.calc import CalcDoc
        from pathlib import Path


        def main():
            pth = Path.cwd() / "example.ods"
            doc = None
            try:
                loader = Lo.load_office(connector=Lo.ConnectPipe())
                doc = CalcDoc.open_doc(fnm=pth, visible=True)
                doc.msgbox("All done")
                doc.close()
            except Exception as e:
                print(f"Error: {e}")
            finally:
                if doc:
                    doc.close()
                Lo.close_office()


        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None



.. _ch02_create_doc:

2.4 Creating a Document
=======================


.. _ch02_4_create_doc_using_lo:

2.4.1 Using Lo Methods
----------------------

The general format of a program that creates a new document, manipulates it in some way, and then saves it, is:

.. include:: ../../resources/odev/02/create_save_tab.rst

A new document is created by calling ``XComponentLoader.loadComponentFromURL()`` with a special URL string for the document type.
The possible strings are listed in :numref:`ch02tbl_new_doc_url`.

.. _ch02tbl_new_doc_url:

.. table:: URLs for Creating New Documents.
    :name: new_doc_type

    =========================================== ============================== 
    URL String                                  Document Type                 
    =========================================== ============================== 
    ``private:factory/swriter``                 Writer                        
    ``private:factory/sdraw``                   Draw                          
    ``private:factory/simpress``                Impress                       
    ``private:factory/scalc``                   Calc                          
    ``private:factory/sdatabase``               Base                          
    ``private:factory/swriter/web``             HTML document in Writer       
    ``private:factory/swriter/GlobalDocument``  A Master document in Writer   
    ``private:factory/schart``                  Chart                         
    ``private:factory/smath``                   Math Formulae                 
    ``.component:Bibliography/View1``           Bibliography Entries          
    ``.component:DB/QueryDesign``               Database User Interfaces      
    ``.component:DB/TableDesign``                                             
    ``.component:DB/RelationDesign``                                          
    ``.component:DB/DataSourceBrowser``                                       
    ``.component:DB/FormGridView``                                            
    =========================================== ============================== 


For instance, a Writer document is created by:

.. tabs::

    .. code-tab:: python

        doc = loader.loadComponentFromURL("private:factory/swriter", "_blank", 0, props)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The office classes include code for simplifying the creation of Writer, Draw, Impress, Calc, and Base documents, which I'll be looking at in later chapters.

.. _ch02_second_srv_mgr:

A Second Service Manager
------------------------

:py:meth:`.Lo.open_doc` and :py:meth:`.Lo.create_doc` do a bit of additional work after document loading/creation – they instantiate a
XMultiServiceFactory_ service manager which is stored in the :py:class:`~.lo.Lo` class. This is done by applying :py:meth:`.Lo.qi` to the document:

.. tabs::

    .. code-tab:: python

        # _ms_factory global in Lo
        doc = loader.loadComponentFromURL("private:factory/swriter", "_blank", 0, props)
        Lo._ms_factory =  Lo.qi(XMultiServiceFactory, doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

First :py:meth:`.Lo.qi` is employed in :py:meth:`~.lo.Lo.create_instance_mcf` to access an interface inside a service.
This time :py:meth:`~.lo.Lo.qi` is casting one interface (XComponent_) to another (XMultiServiceFactory_).

The XMultiServiceFactory_ object is the second service manager we've encountered; the first was an XMultiComponentFactory_ instance, created during Office's loading.

The reasons for Office having two service managers are historical:
the XMultiServiceFactory_ manager is older, and creates a service object without the need for an explicit reference to the remote component context.

As Office developed, it was decided that service object creation should always be relative to an explicit component context,
and so the newer XMultiComponentFactory_ service manager came into being.
A lot of older code still uses the XMultiServiceFactory_ service manager, so both are supported in the :py:class:`~.lo.Lo` class.

Another difference between the managers is that the XMultiComponentFactory_ manager is available as soon as Office is loaded,
while the XMultiServiceFactory_ manager is initialized only when a document is loaded or created.


.. _ch02_4_create_doc_using_classes:

2.4.2 Using the Classes such as CalcDoc
---------------------------------------

The :py:class:`~.calc.CalcDoc`, :py:class:`~.write.WriteDoc`, :py:class:`~.draw.DrawDoc`, :py:class:`~.draw.ImpressDoc` classes can be used to create a document.

These classes all implement the same methods for creating, opening, saving, and closing a documents.

For this example we will use the :py:class:`~.calc.CalcDoc` class to create a Calc document.

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        from ooodev.loader.lo import Lo
        from ooodev.calc import CalcDoc
        from pathlib import Path


        def main():
            pth = Path.cwd()  / "example.ods"
            doc = None
            try:
                loader = Lo.load_office(connector=Lo.ConnectPipe())
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                sheet["A1"].value = 10
                doc.msgbox("All done")
                doc.save_doc(pth)
                doc.close()
            except Exception as e:
                print(f"Error: {e}")
            finally:
                if doc:
                    doc.close()
                Lo.close_office()


        if __name__ == "__main__":
            main()


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. _ch02_save_doc:

2.5 Saving a Document
=====================

The general format of a program that creates a new document, manipulates it in some way, and then saves it, is:

.. include:: ../../resources/odev/02/create_save_tab.rst

One of the great strengths of Office is that it can export a document in a vast number of formats,
but the programmer must specify the output format (which is called a filter in the Office documentation).

|storeToURL|_ takes the name of the output file (in URL format), and an array of properties, one of which should be "FilterName".
Two other useful output properties are "Overwrite" and "Password".
Input and output document properties are listed in the MediaDescriptor service documentation (``lodoc MediaDescriptor service``).

If "Overwrite" is set to true then the file will be saved without prompting the user if the file already exists.
The "Password" property contains a string which must be entered into an Office dialog by the user before the file can be opened again.

The steps in saving a file are:

.. tabs::

    .. code-tab:: python

        save_file_url = FileIO.fnm_to_url(fnm)
        store_props = Props.make_props(Overwrite=True, FilterName=format, Password=password)

        store = Lo.qi(XStorable, doc)
        store.storeToURL(save_file_url, store_props);

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

If you don't want a password, then the third property should be left out.
:py:meth:`.Lo.qi` is used again to cast an interface, this time from XComponent_ to |XStorable|_.

:numref:`ch01fig_office_doc_serv` in :ref:`Chapter 1 <ch01>` shows that |XStorable|_ is part of the OfficeDocument service,
which means that it's inherited by all Office document types.

.. _ch02_what_filter_name:

What's a Filter Name?
---------------------

|storeToURL|_ needs a "FilterName" property value, but what should the string be to export the document in Word format for example?

:py:meth:`.Info.get_filter_names` returns an array of all the filter names supported by Office.

Rather than force a programmer to search through this list for the correct name, :py:meth:`.Lo.save_doc`
allows him to supply just the name and extension of the output file. For example,
in :ref:`ch02_open_doc`, :py:meth:`.Lo.save_doc` was called like so:

.. tabs::

    .. code-tab:: python
    
        Lo.save_doc(doc, "foo.docx") # save as a Word file

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`~.Lo.save_doc` extracts the file extension (i.e. "docx") and maps it to a corresponding filter name in Office
(in this case, "Office Open XML Text"). One concern is that it's not always clear which extension-to-filter mapping should be utilized.
For instance, another suitable filter name for "docx" is "MS Word 2007 XML".
This problem is essentially ignored, by hard wiring a fixed selection into :py:meth:`~.Lo.save_doc`.

Another issue is that the choice of filter sometimes depends on the extension and the document type.
For example, a Writer document saved as a PDF file should use the filter ``writer_pdf_Export``,
but if the document is a spreadsheet then ``calc_pdf_Export`` is the correct choice.

:py:meth:`~.Lo.save_doc` get document type from :py:meth:`.Info.report_doc_type` that calls :py:meth:`.Info.is_doc_type`
to examine the document's service name which is accessed via the XServiceInfo_ interface:

.. tabs::

    .. code-tab:: python
    
        xinfo = Lo.qi(XServiceInfo, doc)
        is_writer = xinfo.supportsService("com.sun.star.text.TextDocument")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Then :py:meth:`~.Lo.save_doc` utilizes :py:meth:`~.Lo.ext_to_format` to get document extension.

The main document service names are listed in :numref:`ch02tbl_doc_service_names`.
For quick access in your scripts use :py:class:`.Lo.Service` where applicable.

.. _ch02tbl_doc_service_names:

.. table:: Document Service Names.
    :name: doc_service_names

    =========== ==================================================
    Document    Type Service Name                                 
    =========== ==================================================
    Writer      ``com.sun.star.text.TextDocument``                
    Draw        ``com.sun.star.drawing.DrawingDocument``          
    Impress     ``com.sun.star.presentation.PresentationDocument``
    Calc        ``com.sun.star.sheet.SpreadsheetDocument``        
    Base        ``com.sun.star.sdb.OfficeDatabaseDocument``       
    =========== ==================================================


We encountered these service names back in :ref:`Chapter 1 <ch01>`, :numref:`ch01fig_office_doc_super` – they're
sub-classes of the ``OfficeDocument`` service.


A third problem is incompleteness; :py:meth:`~.Lo.save_doc` via :py:meth:`~.Lo.ext_to_format` mappings only implements a small subset
of Office's 250+ filter names, so if you try to save a file with an exotic extension then the code will most likely break.
:py:meth:`~.Lo.save_doc` has an overload that takes format as option, that is a filter name.
This overload can be used to if a filter is not implements by :py:meth:`~.Lo.ext_to_format`.

If you want to study the details, start with :py:meth:`~.Lo.save_doc`, and burrow down; the trickiest part is :py:meth:`~.Lo.ext_to_format`.

.. _ch02_close_doc:

2.6 Closing a Document
======================

Closing a document is a pain if you want to check with the user beforehand: should a modified file be saved, thereby overwriting the old version?
|odev|'s solution is not to bother the user, so the file is closed without saving, irrespective of any modifications.
In other words, it's essential to explicitly save a changed document with :py:meth:`.Lo.save_doc` before calling :py:meth:`.Lo.close_doc`.

The code for closing employs :py:meth:`.Lo.qi` to cast the document's XComponent_ interface to XCloseable_:

.. tabs::

    .. code-tab:: python
    
        closeable =  Lo.qi(XCloseable.class, doc)
        closeable.close(false)  # doc. closed without saving

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch02_gen_purpose_convert:

2.7 A General Purpose Converter
===============================

The |convert_doc|_ example in takes two command line arguments: the name of an input file and the extension that should be used when saving the loaded document.
For instance:

.. tabs::

    .. code-tab:: python
    
        python -m doc_convertor --ext 'odp' --file 'points.ppt'

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

will save slides in MS PowerPoint format as an Impress presentation.

The following converts a JPEG image into PNG:

.. tabs::

    .. code-tab:: python
    
        python -m doc_convertor --ext 'png' --file 'skinner.jpg'

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

|convert_doc_src|_ is relatively short. Here is the main section.

.. tabs::

    .. code-tab:: python

        with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:
            # get the absolute path of input file
            p_fnm = FileIO.get_absolute_path(args.file_path)

            name = Info.get_name(p_fnm)  # get name part of file without ext
            if not ext.startswith("."):
                # just in case user did not include . in --ext value
                ext = "." + ext

            p_save = Path(p_fnm.parent, f"{name}{ext}")  # new file, same as old file but different ext

            doc = Lo.open_doc(fnm=p_fnm, loader=loader)
            Lo.save_doc(doc=doc, fnm=p_save)
            Lo.close_doc(doc)

        print(f"All done! converted file: {p_save}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch02_bug_detection:

2.8 Bug Detection and Reporting
===============================

This chapter began our coding with the Office API, and so the possibility of bugs also becomes an issue.
If you find a problem with |odev| classes then please `submit an issue <https://github.com/Amourspirit/python_ooo_dev_tools/issues>`_.
supplying as much detail as possible.

Another source of bugs is the LibreOffice API itself, which is hardly a surprise considering its complexity and age.
If you find a problem, then you should first search LibreOffice's `Bugzilla site <https://bugs.documentfoundation.org/>`_
to see if the problem has been reported previously (it probably has). Various types of search are explained in the
`Bugzilla documentation <https://bugs.documentfoundation.org/docs/en/html/using/>`_.
If you want to report a new bug, then you'll need to set up an account, which is quite simple, and also explained by the documentation.

Often when people report bugs they don't include enough information, perhaps because the error window displayed by Windows is somewhat lacking.
For example, a typical crash report window is show in :numref:`ch02fig_crash_report`.

.. cssclass:: screen_shot invert

    .. _ch02fig_crash_report:
    .. figure:: https://user-images.githubusercontent.com/4193389/178563937-ce136961-4d80-4abf-9a61-f936d30a727b.png
        :alt: The LibreOffice Crash Reported by Windows 7.
        :figclass: align-center

        :The LibreOffice Crash Reported by Windows 7.

If you're going to make an official report, you should first read the article `How to Report Bugs in LibreOffice <https://wiki.documentfoundation.org/QA/BugReport>`__.

Expert forum members and Bugzilla maintainers sometimes point people towards WinDbg for Windows as a tool for producing good debugging details. The wiki has a `detailed explanation <https://wiki.documentfoundation.org/How_to_get_a_backtrace_with_WinDbg>`__
of how to install and use it , which is a bit scary in its complexity.

A much easier alternative is the `WinCrashReport <https://www.nirsoft.net/utils/application_crash_report.html>`__ application from NirSoft_.

It presents the Windows Error Reporting (WER) data generated by a crash in a readable form.

When a crash window appears (like the one in :numref:`ch02fig_crash_report`), start WinCrashReport to examine the automatically-generated error report, as in :numref:`ch02fig_win_crash_rpt_gui`.

.. cssclass:: screen_shot invert

    .. _ch02fig_win_crash_rpt_gui:
    .. figure:: https://user-images.githubusercontent.com/4193389/178566048-95c4d2f5-76c5-4ec5-9ba8-bc7d880b35ef.png
        :alt: Win Crash Report GUI
        :figclass: align-center

        :Win Crash Report GUI

:numref:`ch02fig_win_crash_rpt_gui` indicates that the problem lies inside ``mergedlo.dll``, an access violation (the exception code ``0xC0000005``) to a memory address.

`mergedlo.dll` is part of LibreOffice which probably means that you can find the DLL in /program.
Most Office DLLs are located in that directory.

``WinCrashReport`` generates two alternative call stacks, with slightly more information in the second in this case.
``mergedlo.dll`` is called by the ``uno_getCurrentEnvironment()`` function in ``cppu3.dll``, as indicated in :numref:`ch02fig_sec_call_stack_rpt`.

.. cssclass:: screen_shot invert

    .. _ch02fig_sec_call_stack_rpt:
    .. figure:: https://user-images.githubusercontent.com/4193389/178566993-99c82ab5-ca0b-483b-b42f-8527673aeb09.png
        :alt: The Second Call Stack in WinCrashReport
        :figclass: align-center

        :The Second Call Stack in WinCrashReport.


This narrows the problem to a specific function and two DLLs, which is very helpful.

If you want to better understand the DLLs, they can be examined using `DLL Export Viewer <https://www.nirsoft.net/utils/dll_export_viewer.html>`_
, another NirSoft_ tool, which lists a DLL's exported functions. Running it on `mergedlo.dll` turns up nothing, but the details for `cppu3.dll` are shown in :numref:`ch02fig_dll_export_view`.

.. cssclass:: screen_shot invert

    .. _ch02fig_dll_export_view:
    .. figure:: https://user-images.githubusercontent.com/4193389/178568452-ac8bf39f-a026-4260-bd32-a66ffb6deded.png
        :alt: DLL Export Viewer's view of cppu3.dll
        :figclass: align-center

        :DLL Export Viewer's view of ``cppu3.dll``

``mergedlo.dll`` appears to be empty inside DLL Export Viewer because it exports no functions.
That probably means it's being used as a store for resources, such as icons, cursors, and images.
There's another NirSoft_ tool for looking at DLL resources, called `ResourcesExtract <https://www.nirsoft.net/utils/resources_extract.html>`_
for searching the gigantic code base. :numref:`ch02fig_open_grok_result` shows the results for an ``uno_getCurrentEnvironment`` search.

.. cssclass:: screen_shot invert

    .. _ch02fig_open_grok_result:
    .. figure:: https://user-images.githubusercontent.com/4193389/178569089-5d836f83-458a-49ca-bd90-1614ffc4a86b.png
        :alt: OpenGrok Results for "uno_getCurrentEnvironment"
        :figclass: align-center

        : ``OpenGrok`` Results for ``uno_getCurrentEnvironment``

The function's code is in EnvStack.cxx, which can be examined by clicking on the linked function name shown at the bottom of :numref:`ch02fig_open_grok_result`.

.. _ch02_multiple_docs:

2.9 Opening Multiple Documents
==============================

As of |odev| ``0.9.8``, you can open multiple documents from  a single LibreOffice bridge connection.
In |odev| ``0.9.8`` a new :ref:`ooodev.loader.inst.LoInst` was added. With this new class you can open multiple documents from a single LibreOffice bridge connection.

This is accomplished by creating a new instance of the :ref:`ooodev.loader.inst.LoInst` and then passing the bridge connection to the
:py:meth:`LoInst.load_office <ooodev.utils.inst.lo.lo_inst.LoInst.load_office>` method.

The :ref:`ooodev.loader.inst.LoInst` mirrors the :ref:`ooodev.loader.Lo` in is methods and properties, see :ref:`ooodev.loader.Lo` for any undocumented methods and properties of the :ref:`ooodev.loader.inst.LoInst`.

The following code example demonstrates how to use the :ref:`ooodev.loader.inst.LoInst` to open multiple documents from a single LibreOffice bridge connection.

.. tabs::

    .. code-tab:: python

        from ooodev.loader.lo import Lo
        from ooodev.loader.inst.lo_inst import LoInst
        from ooodev.utils.inst.lo.doc_type import DocTypeStr
        from ooodev.gui import GUI

        def main() -> None:
            # Start LibreOffice using a Socket bridge.
            _ = Lo.load_office(Lo.ConnectSocket())
            # create a new Calc document.
            # Calc.create_doc() uses static bridge connection created by Lo.load_office above by default
            primary_doc= Calc.create_doc()
            # get the first sheet in the primary document
            primary_sheet = Calc.get_sheet(primary_doc, 0)
            # show the primary document
            GUI.set_visible(visible=True, doc=primary_doc)
            # set a value in the primary document sheet
            Calc.set_val(value="LO TEST", sheet=primary_sheet, cell_name="A1")


            # Create a new instance of LoInst and pass the bridge connection from the static Lo class
            # that was created when Lo.load_office was called above.
            inst = LoInst()
            inst.load_office(Lo.bridge_connector)

            # Create a new document from the new instance of LoInst
            secondary_doc = inst.create_doc(DocTypeStr.CALC)
            # show the secondary document
            GUI.set_visible(visible=True, doc=secondary_doc)
            secondary_sheet = Calc.get_sheet(doc, 0)
            # set a value in the secondary document sheet
            Calc.set_val(value="LO INST", sheet=secondary_sheet, cell_name="A1")
            # ... other code

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch02_macro_load:

2.10 Loading for Macro Execution
================================

As of |odev| ``0.11.11``, Macros are now loaded using the :py:class:`~ooodev.macro.MacroLoader` context manager.
This allows for the document context to be manages automatically.

In the example below the ``with MacroLoader()`` context manager is used.
This automatically sets the context for |odev| to the active document.
This allows for the macro to be executed in the context of the active document.

Note that only method that are actually called as macros ``show_hello()`` and ``write_hello()`` require the :py:class:`MacroLoader <ooodev.macro.MacroLoader>` context manager.
The ``write_hello_msg()`` method is call from a macro that already sets the context to the active document and therefore does not require the :py:class:`MacroLoader <ooodev.macro.MacroLoader>` context manager.


.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        from ooodev.office.write import Write
        from ooodev.utils.color import StandardColor
        from ooodev.format.writer.direct.char.font import Font
        from ooodev.dialog.msgbox import MsgBox, MessageBoxButtonsEnum, MessageBoxType
        from ooodev.format.writer.direct.para.alignment import Alignment
        from ooodev.macro import MacroLoader


        def show_hello(*args) -> None:
            with MacroLoader():
                _ = MsgBox.msgbox(
                    "Hello World!",
                    "HI",
                    boxtype=MessageBoxType.INFOBOX,
                    buttons=MessageBoxButtonsEnum.BUTTONS_OK
                )

        def write_hello_msg(msg: str) -> None:
            try:
                cursor = Write.get_cursor(Write.active_doc)
                cursor.gotoEnd(False)
                al = Alignment().align_center
                ft = Font(size=36, u=True, b=True, color=StandardColor.GREEN_DARK2)
                Write.append_para(cursor=cursor, text=msg, styles=[ft, al])
            except Exception as e:
                _ = MsgBox.msgbox(f"This method requires a Writer document.\n{e}")

        def write_hello(*args) -> None:
            with MacroLoader():
                write_hello_msg("Hello World!")


        g_exportedScripts = (show_hello, write_hello)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. |convert_doc| replace:: Write Convert Document Format
.. _convert_doc: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_doc_convert

.. |convert_doc_src| replace:: Write Convert Document Format Source
.. _convert_doc_src: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/general/odev_doc_convert/start.py

.. |extract_txt_py| replace:: extract_text.py
.. _extract_txt_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/impress/odev_extract_text/extract_text.py

.. |extract_graphics_py| replace:: extract_graphics.py
.. _extract_graphics_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/writer/odev_extract_graphics/extract_graphics.py

.. |lo_uno_ex| replace:: LibreOffice Python UNO Examples
.. _lo_uno_ex: https://github.com/Amourspirit/python-ooouno-ex

.. |lo_uno_calc_ex| replace::  Calc Add Range of Data Example
.. _lo_uno_calc_ex: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/calc/odev_add_range_data

.. |lo_uno_sudoku| replace:: LibreOffice Calc Sudoku Example
.. _lo_uno_sudoku: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/calc/sudoku

.. |lo_uno_tab_control| replace:: TAB Control Dialog Box Example
.. _lo_uno_tab_control: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/general/tab_dialog

.. |lo_uno_shared_acc| replace:: Shared Library Access Example
.. _lo_uno_shared_acc: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/general/odev_share_lib

.. |XStorable| replace:: XStorable
.. _XStorable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XStorable.html

.. |storeToURL| replace:: XStorable.storeToURL()
.. _storeToURL: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XStorable.html#af48930bc64a00251aa50915bf087f274

.. _NirSoft: https://www.nirsoft.net/

.. _oooscript: https://pypi.org/project/oooscript/

.. _PropertyValue: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1beans_1_1PropertyValue.html
.. _XCloseable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XCloseable.html
.. _XComponent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XComponent.html
.. _XComponentContext: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1uno_1_1XComponentContext.html
.. _XComponentLoader: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XComponentLoader.html
.. _XDesktop: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XDesktop.html
.. _XMultiComponentFactory: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XMultiComponentFactory.html
.. _XMultiServiceFactory: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XMultiServiceFactory.html
.. _XServiceInfo: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XServiceInfo.html

.. include:: ../../resources/odev/links.rst
