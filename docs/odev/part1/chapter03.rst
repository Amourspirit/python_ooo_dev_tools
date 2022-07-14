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


Work in progress ...

.. |devtools| replace:: Development Tools
.. _devtools: https://help.libreoffice.org/latest/ro/text/shared/guide/dev_tools.html

.. |ooconfigmanage| replace:: Configuration Management
.. _ooconfigmanage: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Config/Configuration_Management
