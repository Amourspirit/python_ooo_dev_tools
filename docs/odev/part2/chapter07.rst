.. _ch07:

******************************************
Chapter 7. Text Content Other than Strings
******************************************

.. topic:: Overview

    Accessing Text Content; Text Frames; Embedded Objects (Math Formulae); Text Fields; Text Tables; Bookmarks

:ref:`ch05` looked at using text cursors to move around inside text documents, adding or extracting text strings.

That chapter utilized the XText_ inheritance hierarchy, which is shown again in :numref:`ch07fig_xtext_super_classes`.

.. cssclass:: diagram invert

    .. _ch07fig_xtext_super_classes:
    .. figure:: https://user-images.githubusercontent.com/4193389/184925140-3e372a8b-7f8b-4c45-9b77-159d0d7fbb41.png
        :alt: Diagram of XText and its Super-classes
        :figclass: align-center

        :XText and its Super-classes.

The documents manipulated in :ref:`ch05` only contained character-based text, but can be a lot more varied,
including text frames, embedded objects, graphics, shapes, text fields, tables, bookmarks, text sections, footnotes and endnotes, and more.

From the XText_ service its possible to access the XTextContent_ interface (see :numref:`ch07fig_xtext_super_classes`), which belongs to the TextContent_ service.
As :numref:`ch07fig_text_content_super_classes` indicates, that service is the parent of many sub-classes which represent different kinds of text document content.

.. cssclass:: diagram invert

    .. _ch07fig_text_content_super_classes:
    .. figure:: https://user-images.githubusercontent.com/4193389/184926151-3b5ed50e-4d41-4b28-a47e-86c0f00fd3ad.png
        :alt: Diagram of The TextContent Service and Some Sub-classes.
        :figclass: align-center

        :The TextContent Service and Some Sub-classes.

A more complete hierarchy can be found in the documentation for TextContent_ (``lodoc TextContent service``).

The two services highlighted in orange relate to graphical content, which is explained in the next chapter.

Table 1 summarizes content types in terms of their services and access methods.
Most of the methods are in Supplier interfaces which are part of the GenericTextDocument_ or OfficeDocument_ services in :numref:`ch05fig_txt_doc_serv_interfaces`.

.. _ch07tbl_create_access_text_content:

.. table:: Creating and Accessing Text Content.
    :name: create_access_text_content

    +------------------+------------------------------------+---------------------------------------------------+
    | Content Name     | Service for Creating Content       | Access Method in Supplier                         |
    +==================+====================================+===================================================+
    | Text Frame       | **TextFrame**                      | ``XNameAccess XTextFrameSupplier``                |
    |                  |                                    |                                                   |
    |                  |                                    | ``getTextFrames()``                               |
    +------------------+------------------------------------+---------------------------------------------------+
    | Embedded Object  | **TextEmbeddedObject**             | ``XComponent XTextEmbeddedObjectSupplier2``       |
    |                  |                                    |                                                   |
    |                  |                                    | ``getEmbeddedObject()``                           |
    +------------------+------------------------------------+---------------------------------------------------+
    | Graphic Object   | **TextGraphicObject**              | ``XNameAccess XTextGraphicObjectsSupplier``       |
    |                  |                                    |                                                   |
    |                  |                                    | ``getGraphicObjects()``                           |
    +------------------+------------------------------------+---------------------------------------------------+
    | Shape            | **text.Shape**,                    | ``XDrawPage XDrawPageSupplier``                   |
    |                  |                                    |                                                   |
    |                  | drawing.Shape or a subclass        | ``getDrawPage()``                                 |
    +------------------+------------------------------------+---------------------------------------------------+
    | Text Field       | **TextField**                      | ``XEnumerationAccess XTextFieldsSupplier``        |
    |                  |                                    |                                                   |
    |                  |                                    | ``getTextFields()``                               |
    +------------------+------------------------------------+---------------------------------------------------+
    | Text Table       | **TextTable**                      | ``XNameAccess XTextTablesSupplier``               |
    |                  |                                    |                                                   |
    |                  |                                    | ``getTextTables()``                               |
    +------------------+------------------------------------+---------------------------------------------------+
    | Bookmark         | **Bookmark**                       | ``XNameAccess XBookmarksSupplier``                |
    |                  |                                    |                                                   |
    |                  |                                    | ``getBookmarks()``                                |
    +------------------+------------------------------------+---------------------------------------------------+
    | Paragraph        | Paragraph                          | ``XEnumerationAccess on XText``                   |
    +------------------+------------------------------------+---------------------------------------------------+
    | Text Section     | TextSection                        | ``XNameAccess XTextSectionsSupplier``             |
    |                  |                                    |                                                   |
    |                  |                                    | ``getTextSections()``                             |
    +------------------+------------------------------------+---------------------------------------------------+
    | Footnote         | Footnote                           | ``XIndexAccess XFootnotesSupplier``               |
    |                  |                                    |                                                   |
    |                  |                                    | ``getFootnotes()``                                |
    +------------------+------------------------------------+---------------------------------------------------+
    | End Note         | Endnote                            | ``XIndexAccess XEndnotesSupplier.getEndnotes()``  |
    +------------------+------------------------------------+---------------------------------------------------+
    | Reference Mark   | ReferenceMark                      | ``XNameAccess XReferenceMarksSupplier``           |
    |                  |                                    |                                                   |
    |                  |                                    | ``getReferenceMarks()``                           |
    +------------------+------------------------------------+---------------------------------------------------+
    | Index            | DocumentIndex                      | ``XIndexAccess XDocumentIndexesSupplier``         |
    |                  |                                    |                                                   |
    |                  |                                    | ``getDocumentIndexes()``                          |
    +------------------+------------------------------------+---------------------------------------------------+
    | Link Target      | LinkTarget                         | ``XNameAccess XLinkTargetSupplier.getLinks()``    |
    +------------------+------------------------------------+---------------------------------------------------+
    | Redline          | RedlinePortion                     | ``XEnumerationAccess XRedlinesSupplier``          |
    |                  |                                    |                                                   |
    |                  |                                    | ``getRedlines()``                                 |
    +------------------+------------------------------------+---------------------------------------------------+
    | Content Metadata | InContentMetaData                  | ``XDocumentMetadataAccess``                       |
    +------------------+------------------------------------+---------------------------------------------------+



**Graphic** Object and **Shape** are discussed in the next chapter.

7.1 How to Access Text Content
==============================

Most of the examples in this chapter create text document content rather than access it.
This is mainly because the different access functions work in a similar way, so you don’t need many examples to get the general idea.

First the document is converted into a supplier, then its ``getXXX()`` method is called (see column 3 of :numref:`ch07tbl_create_access_text_content`).
For example, accessing the graphic objects in a document (see row 3 of :numref:`ch07tbl_create_access_text_content`) requires:

.. tabs::

    .. code-tab:: python

        # get the graphic objects supplier
        ims_supplier = Lo.qi(XTextGraphicObjectsSupplier, doc)

        # access the graphic objects collection
        xname_access = ims_supplier.getGraphicObjects()

The names associated with the graphic objects in XNameAccess_ can be extracted with ``XNameAccess.getElementNames()``, and printed:

.. tabs::

    .. code-tab:: python

        names = xname_access.getElementNames()
        print(f"Number of graphic names: {len(names)}")

        names.sort() # sort them, if you want
        Lo.print_names(names) # useful for printing long lists

A particular object in an XNameAccess_ collection is retrieved with ``getByName()``:

.. tabs::

    .. code-tab:: python

        # get graphic object called "foo"
        obj_graphic = xname_access.getByName("foo")

A common next step is to convert the object into a property set, which makes it possible to lookup the properties stored in the object's service.
For instance, the graphic object’s filename or URL can be retrieved using:

.. tabs::

    .. code-tab:: python

        props =  Lo.qi(XPropertySet, obj_graphic)
        fnm = props.getPropertyValue("GraphicURL") # string

The graphic object's URL is stored in the ``GraphicURL`` property from looking at the documentation for the TextGraphicObject_ service.
It can be (almost) directly accessed by typing ``lodoc TextGraphicObject service``.

It's possible to call ``setPropertyValue()`` to change a property:

``props.setPropertyValue("Transparency", 50)``

**What About the Text Content tha is not covered?**

:numref:`ch07tbl_create_access_text_content` has many rows without bold entries, which means we won't be looking at them.

Except for the very brief descriptions here; for more please consult the Developer's Guide at
https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Working_with_Text_Documents (or type ``loguide Working with Text Documents``).
All the examples in that section are in TextDocuments.java at https://api.libreoffice.org/examples/DevelopersGuide/examples.html#Text.

**Text Sections**. A text section is a grouping of paragraphs which can be assigned their own style settings.
More usefully, a section may be located in another file, which is the mechanism underlying master documents.
See: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Text_Sections (or type ``loguide Text Sections``).

**Footnotes and Endnotes**. Footnotes and endnotes are blocks of text that appear in the page footers and at the end of a document.
They can be treated as XText_ objects, so manipulated using the same techniques as the main document text.
See: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Footnotes_and_Endnotes (or type ``loguide Footnotes``).

**Reference Marks**. Reference marks can be inserted throughout a document, and then jumped to via GetReference text
fields: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Reference_Marks (or type ``loguide Reference Marks``).

**Indexes and Index Marks**. Index marks, like reference marks, can be inserted anywhere in a document,
but are used to generate indices (collections of information) inside the document.
There are several types of index marks used for generating lists of chapter headings (i.e. a book's index),
lists of key words, illustrations, tables, and a bibliography.
For details see: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Indexes_and_Index_Marks (or type ``loguide Indexes``).

**Link Targets**. A link target (sometimes called a jump mark) labels a location inside a document.
These labels can be included as part of a filename so that the document can be opened at that position.
For information, see: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Link_Targets (or type ``loguide Link Targets``).

**Redlines**. Redlines are the changes recorded when a user edits a document with track changes turned on.
Each of the changes is saved as a text fragment (also called a text portion) inside a redline object.
See: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Redline (or type ``loguide Redline``).

Work in progress ...

.. _GenericTextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1GenericTextDocument.html
.. _OfficeDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1OfficeDocument.html
.. _TextContent: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextContent.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html
.. _XTextContent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextContent.html
.. _XNameAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameAccess.html
.. _TextGraphicObject: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextGraphicObject.html