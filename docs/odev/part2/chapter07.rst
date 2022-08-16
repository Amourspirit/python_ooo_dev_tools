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

7.2 Adding a Text Frame to a Document
=====================================

The TextFrame_ service inherits many of its properties and interfaces, so its inheritance hierarchy is shown in detail in :numref:`ch07fig_text_frame_hiearchy`.

.. cssclass:: diagram invert

    .. _ch07fig_text_frame_hiearchy:
    .. figure:: https://user-images.githubusercontent.com/4193389/184963740-aa2692d1-c7fe-4594-8697-bfb3539d2ea0.png
        :alt: Diagram of The TextFrame Service Hierarchy
        :figclass: align-center

        :The TextFrame Service Hierarchy.

:numref:`ch07fig_text_frame_hiearchy` includes two sibling services of TextFrame_: TextEmbeddedObject_ and TextGraphicObject_,
which is discussed a bit later; in fact, we will only get around to TextGraphicObject_ in the next chapter.

The BaseFrameProperties_ service contains most of the frame size and positional properties, such as "Width", "Height", and margin and border distance settings.

A TextFrame_ interface can be converted into a text content (i.e. XTextContent_) or a shape (i.e. XShape_).
Typically, the former is used when adding text to the frame, the latter when manipulating the shape of the frame.

In the |build_doc|_ example, text frame creation is done by :py:meth:`.Write.add_text_frame`, with |build_doc|_ supplying the frame's y-axis coordinate position for its anchor:

.. tabs::

    .. code-tab:: python

        # code fragment from build doc
        tvc = Write.get_view_cursor(doc)
        ypos = tvc.getPosition().Y

        Write.add_text_frame(
                cursor=cursor,
                ypos=ypos,
                text="This is a newly created text frame.\nWhich is over on the right of the page, next to the code.",
                page_num=pg,
                width=4000,
                height=1500,
            )

An anchor specifies how the text content is positioned relative to the ordinary text around it.
Anchoring can be relative to a character, paragraph, page, or another frame.

:py:meth:`.Write.add_text_frame` uses page anchoring, which means that |build_doc|_ must obtain a view cursor, so that an on-screen page position can be calculated.
As :numref:`ch07fig_build_doc_frame_ss` shows, the text frame is located on the right of the page, with its top edge level with the start of the code listing.

.. cssclass:: screen_shot

    .. _ch07fig_build_doc_frame_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184966954-1f3e8e9f-2694-4fc1-8589-a6042912e879.png
        :alt: Screen shot of Text Frame Position in the Document
        :figclass: align-center

        :Text Frame Position in the Document.

In the code fragment above, :py:meth:`.Write.get_view_cursor` creates the view cursor,
and ``XTextViewCursor.getPosition()`` returns its (x, y) coordinate on the page.
The y-coordinate is stored in ``yPos`` until after the code listing has been inserted into the document, and then passed to :py:meth:`.Write.add_text_frame`.

:py:meth:`.Write.add_text_frame` is defined as:

.. tabs::

    .. code-tab:: python

        @classmethod
        def add_text_frame(
            cls,
            cursor: XTextCursor,
            ypos: int,
            text: str,
            width: int,
            height: int,
            page_num: int = 1,
            border_color: Color | None = CommonColor.RED,
            background_color: Color | None = CommonColor.LIGHT_BLUE,
        ) -> bool:
            cargs = CancelEventArgs(Write.add_text_frame.__qualname__)
            cargs.event_data = {
                "cursor": cursor,
                "ypos": ypos,
                "text": text,
                "width": width,
                "height": height,
                "page_num": page_num,
                "border_color": border_color,
                "background_color": background_color,
            }
            _Events().trigger(WriteNamedEvent.TEXT_FRAME_ADDING, cargs)
            if cargs.cancel:
                return False

            ypos = cargs.event_data["ypos"]
            text = cargs.event_data["text"]
            width = cargs.event_data["width"]
            height = cargs.event_data["height"]
            page_num = cargs.event_data["page_num"]
            border_color = cargs.event_data["border_color"]
            background_color = cargs.event_data["background_color"]

            try:
                xframe = Lo.create_instance_msf(XTextFrame, "com.sun.star.text.TextFrame")
                if xframe is None:
                    raise ValueError("Null value")
            except Exception as e:
                raise CreateInstanceMsfError(XTextFrame, "com.sun.star.text.TextFrame") from e

            try:
                tf_shape = Lo.qi(XShape, xframe, True)

                # set dimensions of the text frame
                tf_shape.setSize(Size(width, height))

                #  anchor the text frame
                frame_props = Lo.qi(XPropertySet, xframe, True)
                frame_props.setPropertyValue("AnchorType", TextContentAnchorType.AT_PAGE)
                frame_props.setPropertyValue("FrameIsAutomaticHeight", True)  # will grow if necessary

                # add a red border around all 4 sides
                border = BorderLine()
                border.OuterLineWidth = 1
                if border_color is not None:
                    border.Color = border_color

                frame_props.setPropertyValue("TopBorder", border)
                frame_props.setPropertyValue("BottomBorder", border)
                frame_props.setPropertyValue("LeftBorder", border)
                frame_props.setPropertyValue("RightBorder", border)

                # make the text frame blue
                if background_color is not None:
                    frame_props.setPropertyValue("BackTransparent", False)  # not transparent
                    frame_props.setPropertyValue("BackColor", background_color)  # light blue

                # Set the horizontal and vertical position
                frame_props.setPropertyValue("HoriOrient", HoriOrientation.RIGHT)
                frame_props.setPropertyValue("VertOrient", VertOrientation.NONE)
                frame_props.setPropertyValue("VertOrientPosition", ypos)  # down from top

                # if page number is Not include for TextContentAnchorType.AT_PAGE
                # then Lo Default so At AT_PARAGRAPH
                frame_props.setPropertyValue("AnchorPageNo", page_num)

                # insert text frame into document (order is important here)
                cls._append_text_content(cursor, xframe)
                cls.end_paragraph(cursor)

                # add text into the text frame
                xframe_text = xframe.getText()
                xtext_range = Lo.qi(XTextRange, xframe_text.createTextCursor(), True)
                xframe_text.insertString(xtext_range, text, False)
            except Exception as e:
                raise Exception("Insertion of text frame failed:") from e
            _Events().trigger(WriteNamedEvent.TEXT_FRAME_ADDED, EventArgs.from_args(cargs))
            return True

:py:meth:`~.Write.add_text_frame` starts by creating a TextFrame_ service, and accessing its XTextFrame_ interface:


.. tabs::

    .. code-tab:: python

        xframe = Lo.create_instance_msf(XTextFrame, "com.sun.star.text.TextFrame")

The service name for a text frame is listed as "TextFrame" in row 1 of :numref:`ch07tbl_create_access_text_content`, but :py:meth:`.Lo.create_instance_msf` requires a fully qualified name.
Almost all the text content services, including TextFrame_, are in the ``com.sun.star.text package``.

The XTextFrame_ interface is converted into XShape_ so the frame's dimensions can be set.
The interface is also cast to XPropertySet_ so that various frame properties can be initialized;
these properties are defined in the TextFrame_ and BaseFrameProperties_ services (see :numref:`ch07fig_text_content_super_classes`).

The "AnchorType" property uses the ``AT_PAGE`` anchor constant to tie the frame to the page.
There are five anchor constants: ``AT_PARAGRAPH``, ``AT_CHARACTER``, ``AS_CHARACTER``, ``AT_PAGE``, and ``AT_FRAME``, which are defined in the TextContentAnchorType_ enumeration.

The difference between ``AT_CHARACTER`` and ``AS_CHARACTER`` relates to how the surrounding text is wrapped around the text content.
"AS" means that the text content is treated as a single (perhaps very large) character inside the text,
while "AT" means that the text frame's upper-left corner is positioned at that character location.

The frame's page position is dealt with a few lines later by the ``HoriOrient`` and ``VertOrient`` properties.
The ``HoriOrientation`` and ``VertOrientation`` constants are a convenient way of positioning a frame at the corners or edges of the page.
However, ``VertOrientPosition`` is used to set the vertical position using the ``yPos`` coordinate, and switch off the ``VertOrient`` vertical orientation.

Towards the end of :py:meth:`.Write.add_text_frame`, the frame is added to the document by calling a version of :py:meth:`.Write.append` that expects an XTextContent_ object:

.. tabs::

    .. code-tab:: python

        # internal method call by Write.append() when adding text
        @classmethod
        def _append_text_content(cls, cursor: XTextCursor, text_content: XTextContent) -> None:
            xtext = cursor.getText()
            xtext.insertTextContent(cursor, text_content, False)
            cursor.gotoEnd(False)

It utilizes the ``XText.insertTextContent()`` method.

The last task of :py:meth:`.Write.add_text_frame`, is to insert some text into the frame.

XTextFrame_ inherits XTextContent_, and so has access to the ``getText()`` method (see :numref:`ch07fig_text_frame_hiearchy`).
This means that all the text manipulations possible in a document are also possible inside a frame.

The ordering of the tasks at the end of :py:meth:`~.Write.add_text_frame` is important.
Office prefers that an empty text content be added to the document, and the data inserted afterwards.

Work in progress ...

.. |build_doc| replace:: Build Doc
.. _build_doc: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_doc

.. _BaseFrameProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1BaseFrameProperties.html
.. _GenericTextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1GenericTextDocument.html
.. _OfficeDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1OfficeDocument.html
.. _TextContent: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextContent.html
.. _TextContentAnchorType: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1text.html#a470b1caeda4ff15fee438c8ff9e3d834
.. _TextEmbeddedObject: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextEmbeddedObject.html
.. _TextFrame: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextFrame.html
.. _TextGraphicObject: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextGraphicObject.html
.. _TextGraphicObject: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextGraphicObject.html
.. _XNameAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameAccess.html
.. _XPropertySet: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertySet.html
.. _XShape: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShape.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html
.. _XTextContent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextContent.html
.. _XTextFrame: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextFrame.html