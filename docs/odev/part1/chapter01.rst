.. _ch01:

***********************************
Chapter 1. LibreOffice API Concepts
***********************************

.. topic:: History

    Some History; Help and Examples for the LibreOffice SDK (``lodoc``, ``loguide``);
    Office as a Process; Common Structures (Interface, Property, Service, and Component);
    Service and Interface Inheritance Hierarchies; the Frame- Controller-Model (FCM) Relationship; Extensions; Comparison with Basic

This chapter describes LibreOffice API concepts without resorting to code (that comes along in the next chapter).
These concepts include Office as a (possibly networked) process, the interface, property, service,
and component structures, the two API inheritance hierarchies, and the Frame-Controller-Model (FCM) relationship.

LibreOffice is an open source, cross-platform, office suite, made up of six main applications, and lots of other useful stuff.
The applications are: Writer (a word processor), Draw (vector graphics drawing), Impress (for slide presentations), Calc (spreadsheets),
Base (a database front- end), and Math (for writing formulae).
Some of the lesser- known features include a charting library, spell checker, forms designer, thesaurus, e- mail package,
and support for extensions (e.g. new menu items and libraries). Aside from Open Document Format (ODF) files,
LibreOffice can import, convert, and export a vast number of text, graphic, and other formats,
including Microsoft Office documents, PDF, HTML, SWF (Flash), and SQL databases.

LibreOffice is managed and developed by `The Document Foundation <https://libreoffice.org>`_,
and was first released in 2010. However, earlier Office versions date back to the 1980's,
and traces of this heritage are visible in many parts of its API.
:numref:`ch01fig_timeline` shows a simplified timeline of how StarOffice begat OpenOffice, and so on to LibreOffice.

.. cssclass:: diagram invert

    .. _ch01fig_timeline:
    .. figure:: https://user-images.githubusercontent.com/4193389/177227955-f91f2454-486e-4222-9360-0734b3e50cdf.png
        :alt: OpenOffice Timeline Image

        :Office's Timeline.

This book is not about how to use LibreOffice's GUI (e.g. where to find the menu item for italicizing text).
I'm also not going to discuss how to compile the LibreOffice source, which is a focus of LibreOffice's development
`webpage <https://wiki.documentfoundation.org/Development>`_.
The intention is ot explain how |app_name_bold| (|odev|) can be used to interact with LibreOffice via a console or via macros using python.

.. _ch01_sources_for_api_information:

1.1 Sources for API Information
===============================

This book is an attempt to write a more gradual, modern introduction to the API.

These documents aim to make the more esoteric materials in the developer's guide easier to understand in a python way.
One of the ways will be flattening the learning curve is by hiding parts of the API behind my own collection of utility classes.
This is far from being a novel idea, as it seems that every programmer who has ever written more than a few pages of Office code ends up developing support functions.
Much gratitude to all the intrepid programmers who contributed to this mission in one way or another.

Perhaps the best place for learning about Office macro programming is Andrew Pitonyak's
`website <https://pitonyak.org/>`_, which includes an excellent free-to- download book:
"OpenOffice.org Macros Explained", a macros cookbook, and a document focusing on database macros.

Finding API Documentation Online
--------------------------------

The online API documentation can be time-consuming to search due to its great size.

If you want to have a browse, start at https://api.libreoffice.org/docs/idl/ref/namespaces.html, which takes a while to load.

Each Office application (e.g. Writer, Draw, Impress, Calc, Base, Math) is supported by multiple modules (similar to Java packages).
For example, most of Writer's API is in Office's "text" module, while Impress' functionality comes from the "presentation" and "drawing" modules.
These modules are located in com.sun.star package, which is documented at https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star.html.

Rather than searching manually through a module for a given class, it's a lot quicker to get a search engine to do it for you.
This is the purpose of |dsearch|_.

For instance, at the command line, you can type: ``lodoc xtext`` and the Office API documentation on the XText interface will open in your browser.

``lodoc`` 'almost' always returns the right page, mainly because Office interfaces, and many of its services, have long unique names.
(I'll explain what a service is shortly.) ``lodoc`` can be access by typing `lodoc` in the console.

Service names are less unusual, and so you should probably add the word "service" to your search.
For instance, if you're looking for the Text service, type: ``lodoc text service``

Module names are also quite common words, so add "module" to the search.
If you want to reach the "text" module (which implements most of Writer), search for: ``lodoc text module``

You can call lodoc with Office application names, which are mapped to API module names.
For instance: ``lodoc Impress`` brings up the "presentation" module page.

Searching the Online Developer's Guide
--------------------------------------

The online Developer's Guide can also be time-consuming to search because it's both long (around 1650 pages),
and poorly organized. To help, I've written a ``loguide`` cli which is quite similar to ``lodoc``.
It calls a search engine, limiting the search to the Developer's Guide web pages, and loads the first matching page into your web browser.

The first argument of ``loguide`` must be an Office application name, which restricts the search to the part of the guide focusing on that application's API or otherwise, ``general``.

Type ``loguide -h`` for options.

General example
^^^^^^^^^^^^^^^

.. code-block:: bash

    loguide general Lifetime of UNO Objects

Loads the guide page with that heading into the browser. A less precise query will probably produce the same page, but even when the result is 'wrong' it'll still be somewhere in the guide.

Impress example
^^^^^^^^^^^^^^^

.. code-block:: bash

    loguide impress Page Formatting


Calling ``loguide`` with just an application name, opens the guide at the start of the chapter on that topic.
For example: ``loguide writer`` opens the guide at the start of the "Text Documents" chapter.

loapi
^^^^^

``loapi`` uses a local database to narrow class names and namespaces for a more focused search.

loapi comp
""""""""""
``loapi comp`` can search for a components ``const``, ``enum``, ``exception``, ``interface``, ``singleton``, ``service``, ``struct``, ``typedef`` or ``any``.

Type ``loapi comp -h`` to see options available for ``comp``.


For example:

.. code-block:: bash

    loapi comp --search writer
    Choose an option (default 1):
    [0],  Cancel
    [1],  UnsupportedOverwriteRequest       - com.sun.star.task.UnsupportedOverwriteRequest           - exception
    [2],  LayerWriter                       - com.sun.star.configuration.backend.xml.LayerWriter      - service
    [3],  ManifestWriter                    - com.sun.star.packages.manifest.ManifestWriter           - service
    [4],  Writer                            - com.sun.star.xml.sax.Writer                             - service
    [5],  XCompatWriterDocProperties        - com.sun.star.document.XCompatWriterDocProperties        - interface
    [6],  XManifestWriter                   - com.sun.star.packages.manifest.XManifestWriter          - interface
    [7],  XSVGWriter                        - com.sun.star.svg.XSVGWriter                             - interface
    [8],  XWriter                           - com.sun.star.xml.sax.XWriter                            - interface


Choosing any number greater than ``0`` opens the that components url.
Option ``4`` would open to https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1xml_1_1sax_1_1Writer.html

Search can be narrowed by including ``--component-type`` option.

.. code-block:: bash

    loapi comp --component-type service --search writer
    Choose an option (default 1):
    [0],  Cancel
    [1],  LayerWriter                       - com.sun.star.configuration.backend.xml.LayerWriter      - service
    [2],  ManifestWriter                    - com.sun.star.packages.manifest.ManifestWriter           - service
    [3],  Writer                            - com.sun.star.xml.sax.Writer                             - service

A search parameter can be more that one word.

For Example:

.. code-block:: bash

    loapi comp --component-type exception --search "ill arg"
    Choose an option (default 1):
    [0],  Cancel
    [1],  IllegalArgumentIOException        - com.sun.star.frame.IllegalArgumentIOException           - exception
    [2],  IllegalArgumentException          - com.sun.star.lang.IllegalArgumentException              - exception

searches for all components of type ``exception`` that contain ``ill`` followed by any number of characters and then ``arg``.

loapi ns
""""""""

Similar to ``loapi comp``, ``loapi ns`` search strictly in namespaces.

Type ``loapi ns -h`` to see options available for ``ns``.

For example:

.. code-block:: bash

    loapi ns --search xml
    Choose an option (default 1):
    [0],  Cancel
    [1],  com.sun.star.xml
    [2],  com.sun.star.xml.crypto.sax
    [3],  com.sun.star.xml.dom
    [4],  com.sun.star.xml.crypto
    [5],  com.sun.star.xml.xslt
    [6],  com.sun.star.xml.input
    [7],  com.sun.star.xml.sax
    [8],  com.sun.star.xml.wrapper
    [9],  com.sun.star.xml.xpath
    [10], com.sun.star.xml.dom.views

Choosing any number greater than ``0`` opens the that components url.
Option ``4`` would open to https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1xml_1_1crypto.html

.. tip::

    ``loapi`` can be handy when you are writing code and you have to import LibreOffice components.
    If you know part the name you can quickly find the full import name.

.. _ch01_office_as_process:

1.2 Office as a Process
=======================

Office is started as an OS process, and a Python program communicates with it via a socket or named pipe.
This necessarily complicates the Python/Office link, which is illustrated in :numref:`ch01fig_python_using_office`.

.. cssclass:: diagram invert

    .. _ch01fig_python_using_office:
    .. figure:: https://user-images.githubusercontent.com/4193389/177416327-bb02c050-e7ee-40cd-b1c5-b5b88e9dae78.png
        :alt: Diagram of Python Program Using Office

        :A Python Program Using Office

The invocation of Office and the setup of a named pipe link can be achieved with a single call to the
soffice binary ( ``soffice.exe,  ``soffice.bin`` ).
A call starts the Office executable with several command line arguments, the most important being ``-accept``
which specifies the use of pipes or sockets for the inter-process link.

A call to `XUnoUrlResolver.resolve() <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1bridge_1_1XUnoUrlResolver.html#abaac8ead87dd0ec6dfc1357792cdda3f>`_
creates a remote component context, which acts as proxy for the 'real' component context over in the Office process (see :numref:`ch01fig_python_using_office`).
The context is a container/environment for components and UNO objects which I'll explain below.
When a Python program refers to components and UNO objects in the remote component context, the inter-process bridge maps
those references across the process boundaries to the corresponding components and objects on the Office side.

Underpinning this mapping is the Universal Network Object (UNO) model which links objects in different environments using the UNO remote protocol (URP).
For example, a method call is converted into a byte stream, sent across the bridge and reconstructed. Method results are returned in the same way.

Thankfully, this network communication is hidden by the Office API.
The only place a beginner might encounter UNO mechanisms is when loading or saving documents.

Every document (more generally called a resource) is referred to using a Uniform Resource Identifier (URI);
URIs are employed by Office’s Universal Content Broker (UCB) and Universal Content Providers (UCPs)
to load and save a wide range of data formats.

Connecting to LibreOffice is discussed in the next chapter.

Obtaining a remote component context is not the end of Office’s initialization.

Typically, at least three UNO objects are needed over on the Python side for most programming tasks:
a service manager, a Desktop object, and a component loader.

The service manager is used to load additional services into Office at runtime.
The Desktop object has nothing to do with the OS's desktop – it refers to the top-level of the Office application,
particularly to its GUI. The component loader is used to load or create Office documents.

Other UNO objects might be more useful depending on your programming task.
For example, for historical reasons, Office supports two slightly different service managers
(one that requires an explicit component context argument, and an older one that doesn't).
Both are added to the component context, as a convenience to the programmer;
this detail is hidden by the :py:class:`~.utils.lo.Lo` util class.

.. _ch01_api_data_structures:

1.3 API Data Structures: interface, property, service, and component
====================================================================

There are four main data structures used by the API: **interface**, **property**, **service**, and **component**.

The use of the word 'interface' is obviously influenced by its meaning in Java,
but it's probably best to keep it separate in your mind.
An Office interface is a collection of method prototypes
(i.e. method names, input arguments, and return types)
without any implementation or associated data.
A property is a name-value pair, used to store data.

A service comprises a set of interfaces and properties needed to support an Office feature.

:numref:`ch01fig_service_interface_prop` illustrates how interface, property, and service are related.

.. cssclass:: diagram invert

    .. _ch01fig_service_interface_prop:
    .. figure:: https://user-images.githubusercontent.com/4193389/177419384-0591cdf2-9d4f-4272-8028-4786bed9fc7a.png
        :alt: Diagram of Services, Interfaces, Properties

        :Services, Interfaces, Properties

The Office documentation often talks about property structs (e.g. the Point and KeyEvent structs).

Since interfaces contain no code, a service is a specification for an Office feature.

When a service is implemented (i.e. its interfaces are implemented), it becomes a component.
This distinction means that the Office API can be implemented in different languages (as components)
but always employs the same specifications (services), as represented in :numref:`ch01fig_component_service`.

.. cssclass:: diagram invert

    .. _ch01fig_component_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/177419958-db1061b5-cb33-4056-a7cb-482c72826e0c.png
        :alt: Diagram of Components and Services.

        :Components and Services.

The developer's guide uses a notation like that shown in :numref:`ch01fig_office_doc_serv` to draw a service and its interfaces.

.. cssclass:: diagram invert

    .. _ch01fig_office_doc_serv:
    .. figure:: https://user-images.githubusercontent.com/4193389/177420337-eb786095-1c09-4088-bebb-a4e43d918abe.png
        :alt: Diagram of Office Document service.

        The ``OfficeDocument`` service.

The developer's guide drawing for the SpellChecker service is shown in :numref:`ch01fig_spell_chk_srv`.

.. cssclass:: diagram invert

    .. _ch01fig_spell_chk_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/177420575-08b3122d-1f18-4f97-b4d8-a0807f461c8e.png
        :alt: Diagram of Spell Checker service.

        :The SpellChecker service.

The two figures illustrate a useful naming convention: all interface names start with the letter "X".

The developer's guide notation leaves out information about the properties managed by the services.
Also, the services web pages at the LibreOffice site don't use the guide’s notation.

The URLs for these pages are somewhat difficult to remember.
The best thing is to use my |dsearch|_ tool to find them.
For instance, you can access the office document and spell checker services with:

``lodoc officedocument service``

and

``lodoc spellchecker service``


.. note::

    The "office-document" search result isn't ideal – it takes you to the IDL page for the service.
    You need to click on the "Office-Document" link under the "Classes" heading to get to the actual service details.

The LibreOffice service web pages usually list properties, but sometimes refer to them as 'attributes'.
If the service documentation doesn't describe the properties, then they're probably being managed by a separate “Supplier” interface
(e.g. `XDocumentPropertiesSupplier`_ for OfficeDocument in :numref:`ch01fig_office_doc_serv`).
The supplier will include methods for accessing the properties as an `XPropertySet`_ object.

One great feature of the LibreOffice web pages is the inheritance diagrams on each service and interface page.
Part of the diagram for the `OfficeDocument service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1OfficeDocument.html>`_
is shown in :numref:`ch01fig_inherit_diagram_office_doc`.

.. cssclass:: diagram transparent

    .. _ch01fig_inherit_diagram_office_doc:
    .. figure:: https://user-images.githubusercontent.com/4193389/177428410-a5793eec-3e98-4fc3-ba28-02f9508d5261.png
        :alt: Example Inheritance Diagram for the Office Document

        :Part of the Inheritance Diagram for the OfficeDocument Service.

Each box in the diagram can be clicked upon to jump to the documentation for that subclass or superclass.

.. _ch01_two_inheritance_hierarchies:

1.4 Two Inheritance Hierarchies for Services and interfaces
===========================================================

Services and interfaces both use inheritance, as shown by the UML diagram in :numref:`ch01fig_service_interface_relations`.

.. cssclass:: diagram invert

    .. _ch01fig_service_interface_relations:
    .. figure:: https://user-images.githubusercontent.com/4193389/177429003-eec1bdd0-dadc-4577-9ffa-999570874339.png
        :alt: Diagram of Service and Interface Relationships and Hierarchies.

        :Service and Interface Relationships and Hierarchies.

For example, OfficeDocument is the superclass service of all other document formats, as illustrated in :numref:`ch01fig_office_doc_super`.

.. cssclass:: diagram invert

    .. _ch01fig_office_doc_super:
    .. figure:: https://user-images.githubusercontent.com/4193389/177429219-5cb80ff9-a272-4c9e-a0f9-b8548771384d.png
        :alt: Diagram of Office Document as a Super class Service.

        : ``OfficeDocument`` as a Superclass Service.

Part of this hierarchy can also be seen in :numref:`ch01fig_inherit_diagram_office_doc`.

An interface can also be part of an inheritance hierarchy.
For instance, the `XModel`_ interface inherits XComponent and XInterface, as in :numref:`ch01fig_super_xmodel`.

.. cssclass:: diagram invert

    .. _ch01fig_super_xmodel:
    .. figure:: https://user-images.githubusercontent.com/4193389/177429428-e022d6a0-3302-4f69-bb1d-44379a6aa146.png
        :alt: Diagram of The Super classes of XModel

        :The Superclasses of XModel.


The LibreOffice documentation graphically displays these hierarchies (e.g. see :numref:`ch01fig_inherit_diagram_office_doc`),
but makes no visual distinction between the service and interface hierarchies.
It also represents the "contains" relationship between services and interfaces as inheritance,
rather than as lines with circles as in the developer's guide (e.g. see :numref:`ch01fig_office_doc_serv` and :numref:`ch01fig_spell_chk_srv`).

.. _ch01_fcm_relationship:

1.5 The FCM Relationship
========================

The Frame-Controller-Model (FCM) relationship (or design pattern) is a part of Office
which programmers will encounter frequently.
It appears in the API as connections between the `XFrame`_, `XController`_, and `XModel`_ interfaces,
as shown in :numref:`ch01fig_fcm_relation`.

.. cssclass:: diagram invert

    .. _ch01fig_fcm_relation:
    .. figure:: https://user-images.githubusercontent.com/4193389/177430903-43850d01-c0b5-4352-821b-ca38dfbf9afc.png
        :alt: Diagram of The FCM Relationship

        :The FCM Relationship.

Every Office document inherits the OfficeDocument service (see :numref:`ch01fig_office_doc_super`),
and :numref:`ch01fig_office_doc_serv` shows that OfficeDocument supports the `XModel`_ interface.
This means that every document will include `XModel`_ methods for accessing the document's resources,
such as its URL, file name, type, and meta information.
Via `XModel.getCurrentController() <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XModel.html#a44c3b26a1116ab41654d60357ccda9e1>`_
, a document's controller can be accessed.

A controller manages the visual presentation of a document.
For instance, the Office GUI interacts with the controller to position the cursor in a document,
to control which page is displayed, and to highlight selections.
The `XController`_ interface belongs to the `Controller`_ service, which is a superclass for viewing documents;
subclasses include `TextDocumentView`_, `DrawingDocumentDrawView`_, and `PresentationView`_.

From `XController`_, it's possible to reach `XFrame`_,
which contains information about the document's display window.
A document utilizes two `XWindow`_
objects, called the component and container windows.
The component window represents the rectangular area on screen that displays the document.
It also handles GUI events, such as window activation or minimization. The container window is the component's parent.

For example, a component window displaying a chart might be contained within a spreadsheet window A frame can contain
child frames, allowing the Office GUI to be thought of as a tree of frames.
The root frame of this tree is the Desktop object, which you may recall is one of the first three objects stored in
the remote component context when we start Office. This means that we can move around the frames in the Office GUI starting
from the loaded document, or from the root frame referred to from `XDesktop`_.

For example, `XDesktop`_ provides ``getCurrentFrame()`` to access the currently active frame.

.. _ch01_components_again:

1.6 Components Again
====================

A knowledge of the FCM relationship, and its XFrame, XController, and `XModel`_ interfaces,
lets me give a more detailed definition of a component.
Back in :ref:`Section 3 <ch01_api_data_structures>` section 3 (and in :numref:`ch01fig_component_service`), I said a component was an implemented service. Another way of understanding a component is in terms of how much of the FCM relationship it supports, which allows the 'component' idea to be divided into three:


1. A component that supports both the `XModel`_ and `XController`_ interfaces is usually an Office document.
2. A component with a controller but no model is typically used to implement library functionality that doesn't need to load data. Examples include the spell checker, and Office tools for creating database forms.
3. A component with no model or controller (i.e. just an `XWindow`_ object) is used for simple GUI elements, such as Office's help windows.

Of these three types, the component-as-document (number 1) is the most important for our needs.
In particular, the component loader is used in the remote component context to load Office documents.

.. _ch01_what_is_extension:

1.7 What's an Extension?
========================

.. todo::

    Update cross reference for part 8

The Office developer's guide often uses the words 'extension', 'add-on', and 'add-in'.
There are four chapters on these features in Part 8 (along with macro programming in Python),
but it's worth briefly explaining them now.

An extension is a code library that extends Office's functionality.
Since an extension implements the service, it may also be referred to as a component.

An add-on is an extension with additional XML files defining a GUI for the extension
(e.g. a menu bar, menu item, or toolbar icon).
An add-on is rendered in Office's GUI in the same way as standard Office elements.

An add-in or, to use its full name, a Calc Add-in, is an extension that adds a new function to Calc.

.. _ch01_compare_basic_api:

1.8 A Comparison with the Basic API
===================================

If you start searching the forums, newsgroups, blogs, and web sites for Office examples, it soon becomes clear that
Python is not the language of choice for most Office programmers.
Basic (sometimes called StarBasic, OpenOffice.org Basic, LibreOffice Basic, or even Visual Basic or VB by mistake) is the darling of the coding crowd.

Python is flexable, can run outside of LibreOffice and connect via bridge, and or can be used as a macro.
Python also has an advantage of using the many package on `PYPI <https://pypi.org/>`_.
Python has an advantage in the area of source control and larger projects.

This is understandable since Office (both LibreOffice and OpenOffice) includes an IDE for editing and debugging Basic macros.
Also, there's a lot of good resources on how to utilize these tools
(e.g. start browsing the LibreOffice wiki page `LibreOffice Basic Help <https://help.libreoffice.org/Basic/Basic_Help>`_).
The few books that have been written about programming the Office API have all used Basic
(e.g. Pitonyak's `OpenOffice.org Macros Explained <https://pitonyak.org/book/>`_).

There are two styles of Basic macro programming – scripts can be attached to specific documents, or to the Office application.
In a document, a macro can respond to Office events, such as the loading of the document, or its modification.
The macro can monitor the user's key presses or menu button presses, and can utilize Office dialog.

This isn't the place for a language war between Python and Basic, but it's fair to say that the Basic Office API is more widely used than the Python version!

Unlike Java, Python API and Basic API do not need to use interfaces.
A Python/Basic service directly contains all the interfaces, properties, and methods.
This means that an Office service can be understood as a plain-old object containing methods
and data. One downside of this is no inherent typing_ support.
Well in Basic there is no typing_ support at all; However, this is not the case for Python.

In Python it is possible to cast a to a service go gain typing_ support; However it is tricky because services are not classes
even though ooouno_ and types-unopy_ allow service to be imported as classes. At design time this is fine but at runtime result in an error.
Using ``typing.TYPE_CHECKING`` and ``typing.cast`` we can work around this limitation as show in the following example.

.. collapse:: Example
    :open:

    In this example ``typing.TYPE_CHECKING`` (always ``False`` during runtime) is used
    to ensure the service class is available during design time but not runtime.
    types-unopy_ is require for this example (installs with |odev|)
    This allows for getting full typing support for services.

    .. code-block:: python

        from typing import cast, TYPE_CHECKING
        from ooodev.utils.info import Info
        from ooodev.utils.images_lo import ImagesLo

        if TYPE_CHECKING:
            # only import if design time, will error if runtime.
            from com.sun.star.graphic import Graphic

        def insert_graphic(file_name: str) -> None
            graphic = ImagesLo.load_graphic_file(file_name)
            if Info.support_service(graphic, "com.sun.star.graphic.Graphic")
                # cast type as string as it will not be available during runtime
                img = cast("Graphic", graphic)
                # img now has full typing support in code editor
            else:
                raise ValueError(f"Unable to get service for {file_name}")

            # do work with image here
            ...

The recommended way in |odev| is to use :py:meth:`Lo.qi() <.utils.lo.Lo.qi>` to get access to the desired interface.
This ensures the service has the desired interface and avoids the need for ``typing.cast``.

.. collapse:: Example
    :open:

    Example of querying for interface.

    In this example ``srch`` will automatically have typing support for all the properties and methods XSearchable_ 

    .. code-block:: python
        :emphasize-lines: 3

        from com.sun.star.util import XSearchable
        cell_range = ...
        srch = Lo.qi(XSearchable, cell_range)
        sd = srch.createSearchDescriptor()

Using the basic IDE has has some advantages for simple scripts; However, new tools have emerged and are emerging to make the experience in python desirable in many cases.

types-unopy_ that gives typing_ support for the entire |lo_api|_.

ooouno_ that also contains all |lo_api|_ components in different namespaces. ooouno_ dynamic namespaces automatically gets the appropriate ``uno`` object at runtime, see :numref:`ch01fig_rect_demo`.
The dynamic namespaces give easier access to |lo_api|_ components with full typing_ support and is a real time saver.

.. collapse:: Example

    ooouno_ Example

    At runtime ``ooo.dyn.awt.rectangle.Rectangle`` is actually ``uno.com.sun.star.awt.Rectangle``

    .. code-block:: python

        >>> from ooo.dyn.awt.rectangle import Rectangle
        >>> r = Rectangle(2, 10, 12, 18)
        >>> print(type(r))
        <class 'uno.com.sun.star.awt.Rectangle'>


For ScriptForge there is types-scriptforge_ and for Access2Base there is types-access2base_.

For quicker developer searching there is |dsearch|_.

Then there is this library (|odev|) that takes advantage of some of the aforementioned libraries types-unopy_ and  ooouno_.

Many of these libraries are possible because of `OOO UNO TEMPLATE <https://github.com/Amourspirit/ooo_uno_tmpl>`_ that converts the
entire |lo_api|_ into templates that are converted into ooouno_ and types-unopy_.

.. collapse:: Demo
    :open:

    .. cssclass:: a_gif

        .. _ch01fig_rect_demo:
        .. figure:: https://user-images.githubusercontent.com/4193389/177604603-55660d5d-2aef-4746-a8fe-4365a0dcdaa6.gif
            :alt: OOO Rectangle Demo

            :OOO Rectangle demo


In the Basic API, there's no remote component context since the macros run inside Office or inside a document that is loaded into Office.
In |odev| there is a remote bridge and ``Lo.XSCRIPTCONTEXT`` which implements XScriptContext_.

.. tabs::

    .. code-tab:: vbscript Basic

        Dim oSM, oDesk, oDoc As Object
        Set oSM = CreateObject("com.sun.star.ServiceManager")
        Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
        Set oDoc = oDesk.loadComponentFromURL(
        "file:///C:/tmp/testdoc.odt", "_blank", 0, noArgs())

    .. code-tab:: python
        
        from ooodev.utils.lo import Lo
        from ooodev.office.write import Write

        loader = Lo.load_office()
        doc = Write.open_doc(fnm="file:///C:/tmp/testdoc.odt", loader=loader)


However, if the script is part of a loaded document, then the call to ``loadComponentFromURL()`` isn't needed, reducing the code to:

.. tabs::

    .. code-tab:: vbscript Basic

        Set oSM = CreateObject("com.sun.star.ServiceManager")
        Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
        Set oDoc = oDesk.CurrentComponent

    .. code-tab:: python
        
        from ooodev.utils.lo import Lo
        from ooodev.office.write import Write

        _ = Lo.load_office()
        doc = Write.get_text_doc(Lo.ThisComponent)

Also, Office's Basic runtime environment automatically creates a service manager and Desktop object, so it's unnecessary to create them explicitly.
This reduces the code:

.. tabs::

    .. code-tab:: vbscript Basic

        Set oDoc = StarDesktop.CurrentComponent


    .. code-tab:: python
        
        from ooodev.utils.lo import Lo

        _ = Lo.load_office()
        doc = Lo.ThisComponent


or even:

.. tabs::

    .. code-tab:: vbscript Basic

        Set oDoc = ThisComponent


    .. code-tab:: python
        
        from ooodev.utils.lo import Lo
        doc = Lo.ThisComponent


If other services are needed, Basic programmers call the ``createUnoService()`` function which
transparently requests the named service from the service manager.
Python programmers can call :py:meth:`Lo.create_instance_msf() <.utils.lo.Lo.create_instance_msf>`
For instance:

.. tabs::

    .. code-tab:: vbscript Basic

        set sfAcc = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
        sfAcc.CreateFolder(dirName)
    
    .. code-tab:: python
        
        from com.sun.star.ucb import XSimpleFileAccess

        sf_acc = Lo.create_instance_msf(XSimpleFileAccess, "com.sun.star.ucb.SimpleFileAccess")
        sf_acc.CreateFolder(dir_name)

One of the aims of |odev| is to hide as much of the complexity of Office as the Basic version of the API.

|odev| aims to show how and why python may be a more powerful in many cases.

.. _Controller: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1frame_1_1Controller.html
.. _TextDocumentView: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextDocumentView.html
.. _DrawingDocumentDrawView: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1DrawingDocumentDrawView.html
.. _PresentationView: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1PresentationView.html
.. _XDesktop: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XDesktop.html
.. _XDocumentPropertiesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1document_1_1XDocumentPropertiesSupplier.html
.. _XController: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XController.html
.. _XFrame: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XFrame.html
.. _XModel: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XModel.html
.. _XWindow: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XWindow.html
.. _XPropertySet: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertySet.html
.. _XSearchable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XSearchable.html
.. _XScriptContext: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1script_1_1provider_1_1XScriptContext.html

.. include:: ../../resources/odev/links.rst