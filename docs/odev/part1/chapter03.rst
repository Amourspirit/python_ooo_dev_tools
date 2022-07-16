.. _ch03:

********************
Chapter 3. Examining
********************

.. topic:: Examining Office

    Examining Office; Getting and Setting Document Properties; Examining a Document for API Details; Examining a Document Using |devtools|_

This chapter looks at ways to examine the state of the Office application and a document.
A document will be examined in three different ways: the first retrieves properties about the file, such as its author, keywords,
and when it was last modified. The second and third approaches extract API details, such as what services and interfaces it uses.
This can be done by calling functions in my Utils class or by utilizing the |devtools|_ built into Office.

.. _ch03sec01:

3.1 Examining Office
====================

It's sometimes necessary to examine the state of the Office application, for example to determine its version number or installation directory.
There are two main ways of finding this information, using configuration properties and path settings.

.. _ch03sec01prt01:

3.1.1 Examining Configuration Properties
----------------------------------------

.. todo:: 

    Chapter 3, Link to chapter 15

Configuration management is a complex area, which is explained reasonably well in chapter 15 of the developer's guide and online at
OpenOffice |ooconfigmanage|_; Only basics are explained here.
The easiest way of accessing the relevant online section is by typing: ``loguide "Configuration Management"``.

Office stores a large assortment of XML configuration data as ".xcd" files in the \\share \\registry directory.
They can be programatically accessed in three steps: first a ConfigurationProvider service is created, which represents the configuration database tree.
The tree is examined with a ConfigurationAccess service which is supplied with the path to the node of interest.
Configuration properties can be accessed by name with the XNameAccess interface.


These steps are hidden inside :py:meth:`.Info.get_config` which requires at most two arguments – the path to the required node,
and the name of the property inside that node.

The two most useful paths seem to be ``/org.openoffice.Setup/Product`` and ``/org.openoffice.Setup/L10N``,
which are hardwired as constants in the :py:class:`~.utils.info.Info` class. The simplest version of :py:meth:`~.utils.info.Info.get_config`
looks along both paths by default so the programmer only has to supply a property name when calling the method

This is illustrated in the OfficeInfo.java example in the |oinfo|_:

Many other property names, which don't seem that useful, are documented with the :py:class:`~.info.Info` class.
One way of finding the most current list is to browse `main.xcd` in ``\share\registry``.

.. tabs::

    .. code-tab:: python

        # in demo

        with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:
            print(f"OS Platform: {platform.platform()}")
            print(f"OS Version: {platform.version()}")
            print(f"OS Release: {platform.release()}")
            print(f"OS Architecture: {platform.architecture()}")

            print(f"\nOffice Name: {Info.get_config('ooName')}")
            print(f"\nOffice version (long): {Info.get_config('ooSetupVersionAboutBox')}")
            print(f"Office version (short): {Info.get_config('ooSetupVersion')}")
            print(f"\nOffice language location: {Info.get_config('ooLocale')}")
            print(f"System language location: {Info.get_config('ooSetupSystemLocale')}")

            print(f"\nWorking Dir: {Info.get_paths('Work')}")
            print(f"\nOffice Dir: {Info.get_office_dir()}")
            print(f"\nAddin Dir: {Info.get_paths('Addin')}")
            print(f"\nFilters Dir: {Info.get_paths('Filter')}")
            print(f"\nTemplates Dirs: {Info.get_paths('Template')}")
            print(f"\nGallery Dir: {Info.get_paths('Gallery')}")

Example output
    ::

        OS Platform: Linux-5.15.0-41-generic-x86_64-with-debian-bookworm-sid
        OS Version: #44-Ubuntu SMP Wed Jun 22 14:20:53 UTC 2022
        OS Release: 5.15.0-41-generic
        OS Architecture: ('64bit', 'ELF')

        Office Name: LibreOffice

        Office version (long): 7.3.4.2
        Office version (short): 7.3

        Office language location: en-US
        System language location: 

        Working Dir: file:///home/user/Documents

        Office Dir: /usr/lib/libreoffice

        Addin Dir: file:///usr/lib/libreoffice/program/addin

        Filters Dir: file:///usr/lib/libreoffice/program/filter
        ...

.. _ch03sec01prt02:

3.1.2 Examining Path Settings
-----------------------------

Path settings store directory locations for parts of the Office installation, such as the whereabouts of the gallery and spellchecker files.
A partial list of predefined paths is accessible from within LibreOffice, via the Tools menu: Tools, Options, LibreOffice, Paths.
But the best source of information is the developer's guide, in the "Path Organization" section of chapter 6, or at
OpenOffice |pathorg|_, which can be accessed using: ``loguide "Path Organization"``

One issue is that path settings comes in two forms: a string holding a single directory path, or a string made up of a
``;`` - separated paths. Additionally, the directories are returned in URI format (i.e. they start with ``file:///``).


:py:meth:`.Info.get_paths` hides the creation of a PathSettings service, and the accessing of its properties.

Probably the most common Office forum question about paths is how to determine Office's installation directory.
Unfortunately, that isn't one of the paths stored in the PathSettings service, but the information is accessible
via one of the other paths. It's possible to retrieve the path for AddIns (which is \\program\\addin), and move up
the directory hierarchy two levels. This trick is implemented by :py:meth:`.Info.get_office_dir`.

Examples of using :py:meth:`.Info.get_office_dir` and :py:meth:`.Info.get_paths` appear in |oinfo|_:

.. tabs::

    .. code-tab:: python

        print(f"\nOffice Dir: {Info.get_office_dir()}")
        print(f"\nAddin Dir: {Info.get_paths('Addin')}")
        print(f"\nFilters Dir: {Info.get_paths('Filter')}")
        print(f"\nTemplates Dirs: {Info.get_paths('Template')}")
        print(f"\nGallery Dir: {Info.get_paths('Gallery')}")

.. _ch03sec02:

3.2 Getting and Setting Document Properties
===========================================

Document properties is the information that's displayed when you right-click on a file icon, and select "Properties" from the menu, as in :numref:`ch03fig01`.

.. cssclass:: screen_shot invert

    .. _ch03fig01:
    .. figure:: https://user-images.githubusercontent.com/4193389/179297650-0343ec1e-efb3-4625-9c81-a0589ff9a81f.png
        :alt: A Properties Dialog in Windows 10 for algs.odp

        :A Properties Dialog in Windows 10 for "algs.odp".

If you select the "Details" tab, a list of properties appears like those in :numref:`ch03fig02`.

.. cssclass:: screen_shot invert

    .. _ch03fig02:
    .. figure:: https://user-images.githubusercontent.com/4193389/179298066-7acaa668-7b0b-4a59-bbb8-407ba354bf8a.png
        :alt: Details Properties List for algs.odp

        :Details Properties List for "algs.odp".

An issue with document properties is that the Office API for manipulating them has changed.
The old interfaces were XDocumentInfoSupplier_ and XDocumentInfo_, but these have been deprecated, and replaced by
XDocumentPropertiesSupplier_ and XDocumentProperties_. This wouldn't really matter except that while OpenOffice retains those deprecated interfaces,
LibreOffice has removed them.

.. _ch03sec02prt01:

3.2.1 Reporting OS File Properties
----------------------------------

|doc_props|_ example prints the document properties by calling: ``Info.print_doc_properties(doc)``.

:py:meth:`~.info.Info.print_doc_properties` converts the document to an XDocumentPropertiesSupplier_ interface, and extracts the XDocumentProperties_ object:

.. _ch03_print_doc_properties:

.. tabs::

    .. code-tab:: python

        @classmethod
        def print_doc_properties(cls, doc: object) -> None:
            try:
                doc_props_supp = mLo.Lo.qi(XDocumentPropertiesSupplier, doc, True)
                dps = doc_props_supp.getDocumentProperties()
                cls.print_doc_props(dps=dps)
                ud_props = dps.getUserDefinedProperties()
                mProps.Props.show_obj_props("UserDefined Info", ud_props)
            except Exception as e:
                mLo.Lo.print("Unable to get doc properties")
                mLo.Lo.print(f"    {e}")
            return


Although the XDocumentProperties_ interface belongs to a DocumentProperties_ service, that service does not contain any properties/attributes.
Instead its data is stored inside XDocumentProperties_ and accessed and changed with get/set methods based on the attribute names.
For example, the Author attribute is obtained by calling ``XDocumentProperties.getAuthor()``.

As a consequence, :py:meth:`~.info.Info.print_doc_props` consists of a long list of get method calls inside print statements:

.. tabs::

    .. code-tab:: python

        print("Document Properties Info")
        print("  Author: " + dps.Author)
        print("  Title: " + dps.Title)
        print("  Subject: " + dps.Subject)
        print("  Description: " + dps.Description)
        print("  Generator: " + dps.Generator)

        keys = dps.Keywords
        print("  Keywords: ")
        for keyword in keys:
            print(f"  {keyword}")

        print("  Modified by: " + dps.ModifiedBy)
        print("  Printed by: " + dps.PrintedBy)
        print("  Template Name: " + dps.TemplateName)
        print("  Template URL: " + dps.TemplateURL)
        print("  Autoload URL: " + dps.AutoloadURL)
        print("  Default Target: " + dps.DefaultTarget)
        # and more ...

However, user-defined file properties are accessed with an XPropertyContainer, as can be seen back in :ref:`print_doc_properties() <ch03_print_doc_properties>`.

.. _ch03sec02prt02:

3.2.2 Setting Document Properties
=================================

The setting of document properties is done with set methods, as in :py:meth:`.Info.set_doc_props` which sets the file's subject, title, and author properties:

.. tabs::

    .. code-tab:: python

        @staticmethod
        def set_doc_props(doc: object, subject: str, title: str, author: str) -> None:
            try:
                dp_supplier = mLo.Lo.qi(XDocumentPropertiesSupplier, doc, True)
                doc_props = dp_supplier.getDocumentProperties()
                doc_props.Subject = subject
                doc_props.Title = title
                doc_props.Author = author
            except Exception as e:
                raise mEx.PropertiesError("Unable to set doc properties") from e

This method is called at the end of |doc_props|_:

.. tabs::

    .. code-tab:: python

        Info.set_doc_props(doc, "Example", "Examples", "Amour Spirit")

After the properties are changed, the document must be saved otherwise the changes will be lost when the document is closed.

The changed properties appear in the "Document Statistics" list shown in :numref:`ch03fig03`.

.. cssclass:: screen_shot invert

    .. _ch03fig03:
    .. figure:: https://user-images.githubusercontent.com/4193389/179302791-d8373bd0-7b72-41a3-86b8-dcbd5bac6feb.png
        :alt: "Document Statistics" Properties List for "algs.odp"

        :"Document Statistics" Properties List for "algs.odp".


3.3 Examining a Document for API Information
============================================

After programming with the Office API for a while, you may start to notice that two coding questions keep coming up.
They are:
1. For the service I'm using at the moment, what are its properties?
2. 2. When I need to do something to a document (e.g. close an XComponent instance), which interface should I cast XComponent to by calling :py:meth:`.Lo.qi`?

The first question arose in :ref:`Chapter 2 <ch02>` when set properties in loadComponentFromURL() and storeToURL() were needed.
Unfortunately the LibreOffice documentation for OfficeDocument doesn't list all the properties associated with the service.
Have a look for yourself by typing ``lodoc OfficeDocument service``, which takes you to its IDL Page unfortunately.
You'll then need to click on the OfficeDocument_ link in the "Classes" section to reach the documentation. OfficeDocument's "Public Attributes" section only lists three properties.
Thre is a |odoc_member_list|_ which is a little more helpful but can be challengeing decipher.

The second problem is also only partly addressed by the LibreOffice documentation.
The pages helpfully includes inheritance tree diagrams that can be clicked on to jump to the documentation about other services and interfaces.
But the diagrams don't make a distinction between “contains” relationships (for interfaces in a service) and the two kinds of inheritance (for services and for interfaces).

These complaints have appeared frequently in the Office forums.
Two approaches for easing matters are often suggested. One is to write code to print out details about a loaded document,
which is my approach in the next subsection.
A second technique is to install an Office extension for browsing a document's structure. I'll look at one such extension, MRI, in :ref:`ch03sec03prt02`.

.. _ch03sec03prt01:

3.3.1 Printing Programming Details about a Document
===================================================

The messy job is hidden, the job of collecting service, interface, property, and method information about a document inside the Info and Props utility classes.
The five main methods for retrieving details can be understood by considering their position in :numref:`ch03fig04` Service and Interface Relationship diagram.

.. cssclass:: diagram invert

    .. _ch03fig04:
    .. figure:: https://user-images.githubusercontent.com/4193389/179312606-5628bef3-09e7-4016-be86-545953ed02b8.png
        :alt: Methods to Investigate the Service and Interface Relationships and Hierarchies

        :Methods to Investigate the Service and Interface Relationships and Hierarchies.

The methods are shown in action in the |doc_props|_ example, which loads a document and prints information about its services, interfaces, methods, and properties.
The relevant code fragment:


.. _ch03sec03prt02:

3.3.2 Examining a Document Using Development Tools
==================================================

Work in progress ...

.. |devtools| replace:: Development Tools
.. _devtools: https://help.libreoffice.org/latest/ro/text/shared/guide/dev_tools.html

.. |ooconfigmanage| replace:: Configuration Management
.. _ooconfigmanage: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Config/Configuration_Management

.. |oinfo| replace:: Office Info Demo
.. _oinfo: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_office_info

.. |pathorg| replace:: Path Organization
.. _pathorg: https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Path_Organization

.. |doc_props| replace:: Doc Properties
.. _doc_props: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_doc_prop

.. _OfficeDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1OfficeDocument.html

.. |odoc_member_list| replace:: OfficeDocument Member List
.. _odoc_member_list: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1OfficeDocument-members.html

.. _DocumentProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1DocumentProperties.html
.. _XDocumentProperties: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1document_1_1XDocumentProperties.html
.. _XDocumentPropertiesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1document_1_1XDocumentPropertiesSupplier.html

.. _XDocumentInfoSupplier: https://www.openoffice.org/api/docs/common/ref/com/sun/star/document/XDocumentInfoSupplier.html

.. _XDocumentInfo: https://www.openoffice.org/api/docs/common/ref/com/sun/star/document/XDocumentInfo.html