.. _ch07:

******************************************
Chapter 7. Text Content Other than Strings
******************************************

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

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

The two services highlighted relate to graphical content, which is explained in the next chapter.

:numref:`ch07tbl_create_access_text_content` summarizes content types in terms of their services and access methods.
Most of the methods are in Supplier interfaces which are part of the GenericTextDocument_ or OfficeDocument_ services in :numref:`ch05fig_txt_doc_serv_interfaces`.

.. _ch07tbl_create_access_text_content:

.. table:: Creating and Accessing Text Content.
    :name: create_access_text_content

    +------------------+------------------------------------+---------------------------------------------------+
    | Content Name     | Service for Creating Content       | Access Method in Supplier                         |
    +==================+====================================+===================================================+
    | Text Frame       | **TextFrame**                      | ``XNameAccess XTextFramesSupplier``               |
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

.. _ch07_access_txt_content:

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

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The names associated with the graphic objects in XNameAccess_ can be extracted with ``XNameAccess.getElementNames()``, and printed:

.. tabs::

    .. code-tab:: python

        names = xname_access.getElementNames()
        print(f"Number of graphic names: {len(names)}")

        names.sort() # sort them, if you want
        Lo.print_names(names) # useful for printing long lists

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

A particular object in an XNameAccess_ collection is retrieved with ``getByName()``:

.. tabs::

    .. code-tab:: python

        # get graphic object called "foo"
        obj_graphic = xname_access.getByName("foo")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

A common next step is to convert the object into a property set, which makes it possible to lookup the properties stored in the object's service.
For instance, the graphic object’s filename or URL can be retrieved using:

.. tabs::

    .. code-tab:: python

        props =  Lo.qi(XPropertySet, obj_graphic)
        fnm = props.getPropertyValue("GraphicURL") # string

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

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

.. _ch07_add_txt_frame:

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
        from ooodev.format.writer.direct.frame.area import Color as FrameColor
        from ooodev.format.writer.direct.frame.borders import Side, Sides, BorderLineKind, LineSize
        from ooodev.write import Write, WriteDoc
        # ...
        doc = WriteDoc(Write.create_doc(loader=loader))
        cursor = doc.get_cursor()
        # ...

        cursor.append_para("Here's some code:")
        tvc = doc.get_view_cursor()
        tvc.goto_range(cursor.component.getEnd(), False)

        y_pos = tvc.get_position().Y

        cursor.end_paragraph()
        code_font = Font(name=Info.get_font_mono_name(), size=10)
        code_font.apply(cursor.component)

        cursor.append_line("public class Hello")
        cursor.append_line("{")
        cursor.append_line("  public static void main(String args[]")
        cursor.append_line('  {  System.out.println("Hello World");  }')
        cursor.append_para("}  // end of Hello class")

        # reset the cursor formatting
        ParaStyle.default.apply(cursor.component)

        # Format the background color of the previous paragraph.
        bg_color = ParaBgColor(CommonColor.LIGHT_GRAY)
        cursor.style_prev_paragraph(styles=[bg_color])

        cursor.append_para("A text frame")

        pg = tvc.get_current_page()

        frame_color = FrameColor(CommonColor.DEFAULT_BLUE)
        # create a border
        bdr_sides= Sides(
            all=Side(line=BorderLineKind.SOLID, color=CommonColor.RED, width=LineSize.THIN)
        )

        _ = cursor.add_text_frame(
            text="This is a newly created text frame.\nWhich is over on the right of the page, next to the code.",
            ypos=y_pos,
            page_num=pg,
            width=UnitMM(40),
            height=UnitMM(15),
            styles=[frame_color, bdr_sides],
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

An anchor specifies how the text content is positioned relative to the ordinary text around it.
Anchoring can be relative to a character, paragraph, page, or another frame.

:py:meth:`.WriteTextCursor.add_text_frame` uses page anchoring, which means that |build_doc|_ must obtain a view cursor, so that an on-screen page position can be calculated.
As :numref:`ch07fig_build_doc_frame_ss` shows, the text frame is located on the right of the page, with its top edge level with the start of the code listing.

.. cssclass:: screen_shot

    .. _ch07fig_build_doc_frame_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184966954-1f3e8e9f-2694-4fc1-8589-a6042912e879.png
        :alt: Screen shot of Text Frame Position in the Document
        :figclass: align-center

        :Text Frame Position in the Document.

:py:mod:`ooodev.format.writer.direct.frame.type` module contains size and position classes such as :py:class:`~.writer.direct.frame.type.Anchor` class, which is used to specify the frame's anchor type
that can be passed to :py:meth:`.WriteTextCursor.add_text_frame`.
This creates a rich set of options for positioning the frame.

In the code fragment above, ``doc.get_view_cursor()`` creates the view cursor,
and ``tvc.get_position()`` returns its (x, y) coordinate on the page.
The y-coordinate is stored in ``yPos`` until after the code listing has been inserted into the document, and then passed to :py:meth:`.WriteTextCursor.add_text_frame`.

:py:meth:`.Write.add_text_frame` is defined as:

.. tabs::

    .. code-tab:: python

        # in Write Class
        @classmethod
        def add_text_frame(
            cls,
            *,
            cursor: XTextCursor,
            text: str = "",
            ypos: int | UnitT = 300,
            width: int | UnitT = 5000,
            height: int | UnitT = 5000,
            page_num: int = 1,
            border_color: Color | None = None,
            background_color: Color | None = None,
            styles: Iterable[StyleT] = None,
        ) -> XTextFrame:

            result = None
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

            arg_ypos = cast(Union[int, UnitT], cargs.event_data["ypos"])
            text = cargs.event_data["text"]
            arg_width = cast(Union[int, UnitT], cargs.event_data["width"])
            arg_height = cast(Union[int, UnitT], cargs.event_data["height"])
            page_num = cargs.event_data["page_num"]
            border_color = cargs.event_data["border_color"]
            background_color = cargs.event_data["background_color"]

            try:
                ypos = arg_ypos.get_value_mm100()
            except AttributeError:
                ypos = int(arg_ypos)
            try:
                width = arg_width.get_value_mm100()
            except AttributeError:
                width = int(arg_width)
            try:
                height = arg_height.get_value_mm100()
            except AttributeError:
                height = int(arg_height)

            xframe = mLo.Lo.create_instance_msf(XTextFrame, "com.sun.star.text.TextFrame", raise_err=True)

            try:
                tf_shape = mLo.Lo.qi(XShape, xframe, True)

                # set dimensions of the text frame
                tf_shape.setSize(UnoSize(width, height))

                #  anchor the text frame
                frame_props = mLo.Lo.qi(XPropertySet, xframe, True)
                # if page number is Not include for TextContentAnchorType.AT_PAGE
                # then Lo Default so At AT_PARAGRAPH
                if not page_num or page_num < 1:
                    frame_props.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH)
                else:
                    frame_props.setPropertyValue("AnchorType", TextContentAnchorType.AT_PAGE)
                    frame_props.setPropertyValue("AnchorPageNo", page_num)

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

                # insert text frame into document (order is important here)
                cls._append_text_content(cursor, xframe)
                cls.end_paragraph(cursor)

                if text:
                    xframe_text = xframe.getText()
                    xtext_range = mLo.Lo.qi(XTextRange, xframe_text.createTextCursor(), True)
                    xframe_text.insertString(xtext_range, text, False)
                    result = xframe

                if styles:
                    srv = ("com.sun.star.text.TextFrame", "com.sun.star.text.ChainedTextFrame")
                    for style in styles:
                        if style.support_service(*srv):
                            style.apply(xframe)

            except Exception as e:
                raise Exception("Insertion of text frame failed:") from e
            _Events().trigger(WriteNamedEvent.TEXT_FRAME_ADDED, EventArgs.from_args(cargs))
            return result

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`~.Write.add_text_frame` starts by creating a TextFrame_ service, and accessing its XTextFrame_ interface:


.. tabs::

    .. code-tab:: python

        xframe = Lo.create_instance_msf(XTextFrame, "com.sun.star.text.TextFrame")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

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

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

It utilizes the ``XText.insertTextContent()`` method.

The last task of :py:meth:`.Write.add_text_frame`, is to insert some text into the frame.

XTextFrame_ inherits XTextContent_, and so has access to the ``getText()`` method (see :numref:`ch07fig_text_frame_hiearchy`).
This means that all the text manipulations possible in a document are also possible inside a frame.

The ordering of the tasks at the end of :py:meth:`~.Write.add_text_frame` is important.
Office prefers that an empty text content be added to the document, and the data inserted afterwards.

.. _ch07_add_txt_embedded:

7.3 Adding a Text Embedded Object to a Document
===============================================

.. todo::

    Chapter 7.3. Create a link to chapter 33

Text embedded object content support OLE (Microsoft's Object Linking and Embedding), and is typically used to create a frame linked to an external Office document.
Probably, its most popular use is to link to a chart, but we'll delay looking at that until Chapter 33.

The best way of getting an idea of what OLE objects are available is to go to the Writer application's Insert menu, Object, "OLE Object" dialog.
In my version of Office, it lists Office spreadsheet, chart, drawing, presentation, and formula documents, and a range of Microsoft and PDF types.

Note that text embedded objects aren't utilized for adding graphics to a document.

That's easier to do using the TextGraphicObject_ or GraphicObjectShape_ services, which is described next.

In this section we look at how to insert mathematical formulae into a text document.

The example code is in |math_ques|_, but most of the formula embedding is performed by :py:meth:`.Write.add_formula`
that is invoked when :py:meth:`.WriteTextCursor.add_formula` is called:

.. tabs::

    .. code-tab:: python

        # in Write Class
        @classmethod
        def add_formula(cls, cursor: XTextCursor, formula: str) -> bool:
            cargs = CancelEventArgs(Write.add_formula.__qualname__)
            cargs.event_data = {"cursor": cursor, "formula": formula}
            _Events().trigger(WriteNamedEvent.FORMULA_ADDING, cargs)
            if cargs.cancel:
                return False
            formula = cargs.event_data["formula"]
            embed_content = Lo.create_instance_msf(
                XTextContent, "com.sun.star.text.TextEmbeddedObject", raise_err=True
            )
            try:
                # set class ID for type of object being inserted
                props = Lo.qi(XPropertySet, embed_content, True)
                props.setPropertyValue("CLSID", Lo.CLSID.MATH)
                props.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER)

                # insert object in document
                cls._append_text_content(cursor=cursor, text_content=embed_content)
                cls.end_line(cursor)

                # access object's model
                embed_obj_supplier = Lo.qi(XEmbeddedObjectSupplier2, embed_content, True)
                embed_obj_model = embed_obj_supplier.getEmbeddedObject()

                formula_props = Lo.qi(XPropertySet, embed_obj_model, True)
                formula_props.setPropertyValue("Formula", formula)
                Lo.print(f'Inserted formula "{formula}"')
            except Exception as e:
                raise Exception(f'Insertion fo formula "{formula}" failed:') from e
            _Events().trigger(WriteNamedEvent.FORMULA_ADDED, EventArgs.from_args(cargs))
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

A math formula is passed to :py:meth:`~.Write.add_formula` as a string in a format this is explained shortly.

The method begins by creating a TextEmbeddedObject_ service, and referring to it using the XTextContent_ interface:

.. tabs::

    .. code-tab:: python

        embed_content = Lo.create_instance_msf(
                XTextContent, "com.sun.star.text.TextEmbeddedObject", raise_err=True
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Details about embedded objects are given in row 2 of :numref:`ch07tbl_create_access_text_content`.

Unlike TextFrame_ which has an XTextFrame_ interface, there's no ``XTextEmbeddedObject`` interface for TextEmbeddedObject_.
This can be confirmed by looking at the TextFrame_ inheritance hierarchy in :numref:`ch07fig_text_content_super_classes`.
There is an ``XEmbeddedObjectSuppler``, but that's for accessing objects, not creating them.
Instead XTextContent_ interface is utilized in :py:meth:`.Lo.create_instance_msf` because it's the most specific interface available.

The XTextContent_ interface is converted to XPropertySet_ so the "CLSID" and "AnchorType" properties can be set.
"CLSID" is specific to ``TextEmbeddedObject`` – its value is the OLE class ID for the embedded document.
The :py:class:`.Lo.CLSID` contains the class ID constants for Office's documents.

The "AnchorType" property is set to ``AS_CHARACTER`` so the formula string will be anchored in the document in the same way as a string of characters.

As with the text frame in :py:meth:`.Write.add_text_frame`, an empty text content is added to the document first, then filled with the formula.

The embedded object's content is accessed via the XEmbeddedObjectSupplier2_ interface which has a get method for obtaining the object:

.. tabs::

    .. code-tab:: python

        # access object's model
        embed_obj_supplier = Lo.qi(XEmbeddedObjectSupplier2, embed_content, True)
        embed_obj_model = embed_obj_supplier.getEmbeddedObject()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The properties for this empty object (embed_obj_model) are accessed, and the formula string is assigned to the "Formula" property:

.. tabs::

    .. code-tab:: python

        formula_props = Lo.qi(XPropertySet, embed_obj_model, True)
        formula_props.setPropertyValue("Formula", formula)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch07_what_formula:

7.3.1 What's a Formula String?
------------------------------

Although the working of :py:meth:`.Write.add_formula` has been explained, the format of the formula string that's passed to it has not been explained.
There's a good overview of the notation in the "Commands Reference" appendix of Office's "Math Guide", available at https://libreoffice.org/get-help/documentation
For example, the formula string: "1 {5}over{9} + 3 {5}over{9} = 5 {1}over{9}" is rendered as:

.. math::

   1 \frac{5}{9} + 3 \frac{5}{9} = 5 \frac{1}{9}

.. _ch07_build_formulae:

7.3.2 Building Formulae
-----------------------

|math_ques|_ is mainly a for-loop for randomly generating numbers and constructing simple formulae strings.
Ten formulae are added to the document, which is saved as ``mathQuestions.pdf``. The ``main()`` function:

.. tabs::

    .. code-tab:: python

        def main() -> int:
            delay = 2_000  # delay so users can see changes.

            loader = Lo.load_office(Lo.ConnectPipe())

            doc = WriteDoc(Write.create_doc(loader=loader))

            try:
                doc.set_visible()

                cursor = doc.get_cursor()
                cursor.append_para("Math Questions")
                cursor.style_prev_paragraph("Heading 1")

                cursor.append_para("Solve the following formulae for x:\n")

                # lock screen updating and add formulas
                # locking screen is not strictly necessary but is faster when add lost of input.
                with Lo.ControllerLock():
                    for _ in range(10):  # generate 10 random formulae
                        iA = random.randint(0, 7) + 2
                        iB = random.randint(0, 7) + 2
                        iC = random.randint(0, 8) + 1
                        iD = random.randint(0, 7) + 2
                        iE = random.randint(0, 8) + 1
                        iF1 = random.randint(0, 7) + 2

                        choice = random.randint(0, 2)

                        # formulas should be wrapped in {} but for formatting reasons it is easier to work with [] and replace later.
                        if choice == 0:
                            formula = f"[[[sqrt[{iA}x]] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1} ]]"
                        elif choice == 1:
                            formula = (
                                f"[[[{iA}x] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1}]]"
                            )
                        else:
                            formula = f"[{iA}x + {iB} = {iC}]"

                        # replace [] with {}
                        cursor.add_formula(formula.replace("[", "{").replace("]", "}"))
                        cursor.end_paragraph()

                cursor.append_para(f"Timestamp: {DateUtil.time_stamp()}")

                Lo.delay(delay)
                doc.save_doc(pth / "mathQuestions.pdf")
                doc.close_doc()
                Lo.close_office()

            except Exception:
                Lo.close_office()
                raise

            return 0

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:numref:`ch07fig_math_formula_ss` shows a screenshot of part of ``mathQuestions.pdf``.

.. cssclass:: screen_shot invert

    .. _ch07fig_math_formula_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184988764-6c2891eb-bf2d-4fc5-bc38-1a99b08f06dc.png
        :alt: Screen shot of Math Formulae in a Text Document
        :figclass: align-center

        :Math Formulae in a Text Document.

.. _ch07_txt_fields:

7.4 Text Fields
===============

A text field differs from other text content in that its data is generated dynamically by the document, or by an external source such as a database.
Document-generated text fields include text showing the current date, the page number, the total number of pages in the document, and cross-references to other areas in the text.
We'll look at three examples: the ``DateTime``, ``PageNumber``, and ``PageCount`` text fields.

When a text field depends on an external source, there are two fields to initialize:
the master field representing the external source, and the dependent field for the data used in the document; only the dependent field is visible.
Here we won't be giving any dependent/master field examples, but there's one in the Development Guide section on text fields,
at: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Text_Fields (or type ``loguide Text Fields``).

It utilizes the User master field, which allows the external source to be user-defined data.
The code appears in the TextDocuments.java example at https://api.libreoffice.org/examples/DevelopersGuide/examples.html#Text.

Different kinds of text field are implemented as sub-classes of the TextField_ service.
You can see the complete hierarchy in the online documentation for TextField_.
:numref:`ch07fig_simple_text_field_hiearchy` presents a simplified version.

.. cssclass:: diagram invert

    .. _ch07fig_simple_text_field_hiearchy:
    .. figure:: https://user-images.githubusercontent.com/4193389/184990923-2c7db8e2-5a5d-4a34-be07-a0ff20e0b35e.png
        :alt: Diagram of Simplified Hierarchy for the TextField Service
        :figclass: align-center

        :Simplified Hierarchy for the TextField Service.

.. _ch07_datetime_textfield:

7.4.1 The DateTime TextField
----------------------------

The |build_doc|_ example ends with a few lines that appear to do the same thing twice:

.. tabs::

    .. code-tab:: python

        # code fragment from build doc
        cursor.append_para("\nTimestamp: " + DateUtil.time_stamp() + "\n")
        cursor.append("Time (according to office): ")
        cursor.append_date_time()
        cursor.end_paragraph()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.DateUtil.time_stamp` inserts a timestamp (which includes the date and time), and then :py:meth:`.WriteTextViewCursor.append_date_time` invokes :py:meth:`.Write.append_date_time` which inserts the date and time.
Although these may seem to be the same, :py:meth:`~.DateUtil.time_stamp` adds a string while :py:meth:`~.Write.append_date_time` creates a text field.
The difference becomes apparent if you open the file some time after it was created.

:numref:`ch07fig_time_stamps_ss` shows two screenshots of the time-stamped parts of the document taken after it was first generated, and nearly 50 minutes later.

.. cssclass:: screen_shot invert

    .. _ch07fig_time_stamps_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184992086-499fcafc-e1ad-45ed-b005-f02fccf55339.png
        :alt: Screen shot of the document Timestamps.
        :figclass: align-center

        :Screenshots of the Timestamps.

The text field timestamp is updated each time the file is opened in edit mode (which is the default in Writer).

This dynamic updating occurs in all text fields.
For example, if you add some pages to a document, all the places in the document that use the PageCount text field will be updated to show the new length.

:py:meth:`.Write.append_date_time` creates a DateTime_ service, and returns its XTextField_ interface (see :numref:`ch07fig_simple_text_field_hiearchy`).
The TextField_ service only contains two properties, with most being in the subclass (DateTime in this case).

.. tabs::

    .. code-tab:: python

        # in Write Class

        @classmethod
        def append_date_time(cls, cursor: XTextCursor) -> None:
            dt_field = Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.DateTime")
            Props.set_property(dt_field, "IsDate", True)  # so date is reported
            xtext_content = Lo.qi(XTextContent, dt_field, True)
            cls._append_text_content(cursor, xtext_content)
            cls.append(cursor, "; ")

            dt_field = Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.DateTime")
            Props.set_property(dt_field, "IsDate", False)  # so time is reported
            xtext_content = Lo.qi(XTextContent, dt_field, True)
            cls._append_text_content(cursor, xtext_content)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The method adds two DateTime text fields to the document.
The first has its "IsDate" property set to true, so that the current date is inserted; the second sets "IsDate" to false so the current time is shown.

.. _ch07_pagenumber_pagecount:

7.4.2 The PageNumber and PageCount Text Fields
----------------------------------------------

As discussed most of |story_creator|_ in :ref:`ch06`, but skipped over how page numbers were added to the document's page footer. The footer is shown in :numref:`ch07fig_footer_text_fields_ss`.

.. cssclass:: screen_shot invert

    .. _ch07fig_footer_text_fields_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184993404-97a2d903-9aee-4198-9695-a94b938768b5.png
        :alt: Screen shot of Page Footer using Text Fields
        :figclass: align-center

        :Page Footer using Text Fields.

:py:meth:`.Write.set_page_numbers` inserts the ``PageNumber`` and ``PageCount`` text fields into the footer's text area:

.. tabs::

    .. code-tab:: python

        # in Write Class
        @classmethod
        def set_page_numbers(cls, text_doc: XTextDocument) -> None:
            props = Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
            if props is None:
                raise PropertiesError("Could not access the standard page style")

            try:
                props.setPropertyValue("FooterIsOn", True)
                #   footer must be turned on in the document
                footer_text = Lo.qi(XText, props.getPropertyValue("FooterText"), True)
                footer_cursor = footer_text.createTextCursor()

                Props.set_property(
                    prop_set=footer_cursor, name="CharFontName", value=Info.get_font_general_name()
                )
                Props.set_property(prop_set=footer_cursor, name="CharHeight", value=12.0)
                Props.set_property(prop_set=footer_cursor, name="ParaAdjust", value=ParagraphAdjust.CENTER)

                # add text fields to the footer
                pg_number = cls.get_page_number()
                pg_xcontent = Lo.qi(XTextContent, pg_number)
                if pg_xcontent is None:
                    raise MissingInterfaceError(
                        XTextContent, f"Missing interface for page number. {XTextContent.__pyunointerface__}"
                    )
                cls._append_text_content(cursor=footer_cursor, text_content=pg_xcontent)
                cls._append_text(cursor=footer_cursor, text=" of ")
                pg_count = cls.get_page_count()
                pg_count_xcontent = Lo.qi(XTextContent, pg_count)
                if pg_count_xcontent is None:
                    raise MissingInterfaceError(
                        XTextContent, f"Missing interface for page count. {XTextContent.__pyunointerface__}"
                    )
                cls._append_text_content(cursor=footer_cursor, text_content=pg_count_xcontent)
            except Exception as e:
                raise Exception("Unable to set page numbers") from e

        @staticmethod
        def get_page_number() -> XTextField:
            num_field = Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.PageNumber")
            Props.set_property(prop_set=num_field, name="NumberingType", value=NumberingType.ARABIC)
            Props.set_property(prop_set=num_field, name="SubType", value=PageNumberType.CURRENT)
            return num_field

        @staticmethod
        def get_page_count() -> XTextField:
            pc_field = Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.PageCount")
            Props.set_property(prop_set=pc_field, name="NumberingType", value=NumberingType.ARABIC)
            return pc_field

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Write.set_page_numbers` starts by accessing the "Standard" property set (style) for the page style family.
Via its properties, the method turns on footer functionality and accesses the footer text area as an XText_ object.

An XTextCursor_ is created for the footer text area, and properties are configured:

.. tabs::

    .. code-tab:: python

        footer_text = Lo.qi(XText, props.getPropertyValue("FooterText"), True)
        footer_cursor = footer_text.createTextCursor()
        Props.set_property(
            prop_set=footer_cursor, name="CharFontName", value=Info.get_font_general_name()
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

These properties will be applied to the text and text fields added afterwards:

.. tabs::

    .. code-tab:: python

        Write.append(footer_cursor, Write.get_page_number())
        Write.append(footer_cursor, " of ")
        Write.append(footer_cursor, Write.get_page_count())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`~.Write.get_page_number` and :py:meth:`~.Write.get_page_count` deal with the properties for the PageNumber and PageCount fields.

.. _ch07_add_txt_tbl:

7.5 Adding a Text Table to a Document
=====================================

The |make_table|_ example reads in data about James Bond movies from ``bondMovies.txt`` and stores it as a text table in ``table.odt``.
The first few rows are shown in :numref:`ch07fig_bond_movie_ss`.

.. cssclass:: screen_shot

    .. _ch07fig_bond_movie_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/185215630-734ba335-870e-4f43-8c42-d181be221f06.png
        :alt: Screen shot of A Bond Movies Table
        :figclass: align-center

        :A Bond Movies Table.

The ``bondMovies.txt`` file is read by ``read_table()`` utilizing  Python file processing with pythons ``csv.reader``. It returns a 2D-list:

.. tabs::

    .. code-tab:: python

        # example partial result from read_table()
        [
            ["Title",  "Year", "Actor", "Director"],
            ["Dr. No", "1962", "Sean Connery", "Terence Young"],
            ["From Russia with Love", "1963", "Sean Connery", "Terence Young"],
        ]

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Each line in ``bondMovies.txt`` is converted into a string array by pulling out the sub-strings delimited by tab characters.

``read_table()`` ignores lines in the file that are know not to be csv lines. First valid row in the list contains the table's header text.

The first few lines of ``bondMovies.txt`` are:

.. code-block:: text

    // http://en.wikipedia.org/wiki/James_Bond#Ian_Fleming_novels

    Title Year Actor Director

    Dr. No 1962 Sean Connery Terence Young
    From Russia with Love 1963 Sean Connery Terence Young
    Goldfinger 1964 Sean Connery Guy Hamilton
    Thunderball 1965 Sean Connery Terence Young
    You Only Live Twice 1967 Sean Connery Lewis Gilbert
    On Her Majesty's Secret Service 1969 George Lazenby Peter R. Hunt
    Diamonds Are Forever 1971 Sean Connery Guy Hamilton
    Live and Let Die 1973 Roger Moore Guy Hamilton
    The Man with the Golden Gun 1974 Roger Moore Guy Hamilton
    The Spy Who Loved Me 1977 Roger Moore Lewis Gilbert
        :

The ``main()`` function for |make_table|_ is:

.. tabs::

    .. code-tab:: python

        def main() -> int:

            fnm = Path(__file__).parent / "data" / "bondMovies.txt"  # source csv file

            tbl_data = read_table(fnm)

            delay = 2_000  # delay so users can see changes.

            with Lo.Loader(Lo.ConnectSocket()) as loader:

                doc = Write.create_doc(loader=loader)

                try:
                    GUI.set_visible(is_visible=True, odoc=doc)

                    cursor = Write.get_cursor(doc)

                    Write.append_para(cursor, "Table of Bond Movies")
                    Write.style_prev_paragraph(cursor, "Heading 1")
                    Write.append_para(cursor, 'The following table comes form "bondMovies.txt"\n')

                    # Lock display updating for faster writing of table into document.
                    with Lo.ControllerLock():
                        Write.add_table(cursor=cursor, table_data=tbl_data)
                        Write.end_paragraph(cursor)

                    Lo.delay(delay)
                    Write.append(cursor, f"Timestamp: {DateUtil.time_stamp()}")
                    Lo.delay(delay)
                    Lo.save_doc(doc, "table.odt")

                finally:
                    Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Write.add_table` does the work of converting the list of rows into a text table.

:numref:`ch07fig_text_tabls_hiearchy` shows the hierarchy for the TextTable_ service: it's a subclass of TextContent_ and supports the XTextTable_ interface.

.. cssclass:: diagram invert

    .. _ch07fig_text_tabls_hiearchy:
    .. figure:: https://user-images.githubusercontent.com/4193389/185219547-87a5789e-f06c-40e2-b182-664fec13d8f4.png
        :alt: Diagram of The Text Table Hierarchy
        :figclass: align-center

        :The TextTable Hierarchy.

XTextTable_ contains methods for accessing a table in terms of its rows, columns, and cells.
The cells are referred to using names, based on letters for columns and integers for rows, as in :numref:`ch07fig_cell_name_tbl_ss`.

.. cssclass:: screen_shot invert

    .. _ch07fig_cell_name_tbl_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/185220105-77768947-c6e5-43c5-86d6-b7a3c9ac3f3c.png
        :alt: Screen shot of he Cell Names in a Table
        :figclass: align-center

        :The Cell Names in a Table.

:py:meth:`.Write.add_table` uses this naming scheme in the ``XTextTable.getCellByName()`` method to assign data to cells:

.. tabs::

    .. code-tab:: python

        # in Write Class
        @classmethod
        def add_table(
            cls,
            cursor: XTextCursor,
            table_data: Table,
            header_bg_color: Color | None = CommonColor.DARK_BLUE,
            header_fg_color: Color | None = CommonColor.WHITE,
            tbl_bg_color: Color | None = CommonColor.LIGHT_BLUE,
            tbl_fg_color: Color | None = CommonColor.BLACK,
            first_row_header: bool = True,
            styles: Iterable[StyleT] = None,
        ) -> XTextTable:

            cargs = CancelEventArgs(Write.add_table.__qualname__)
            cargs.event_data = {
                "cursor": cursor,
                "table_data": table_data,
                "header_bg_color": header_bg_color,
                "header_fg_color": header_fg_color,
                "tbl_bg_color": tbl_bg_color,
                "tbl_fg_color": tbl_fg_color,
                "first_row_header": first_row_header,
                "styles": styles,
            }
            _Events().trigger(WriteNamedEvent.TABLE_ADDING, cargs)
            if cargs.cancel:
                return False

            header_bg_color = cargs.event_data["header_bg_color"]
            header_fg_color = cargs.event_data["header_fg_color"]
            tbl_bg_color = cargs.event_data["tbl_bg_color"]
            tbl_fg_color = cargs.event_data["tbl_fg_color"]
            first_row_header = cargs.event_data["first_row_header"]

            def make_cell_name(row: int, col: int) -> str:
                return TableHelper.make_cell_name(row=row + 1, col=col + 1)

            def set_cell_header(cell_name: str, data: str, table: XTextTable) -> None:
                cell_text = mLo.Lo.qi(XText, table.getCellByName(cell_name), True)
                if first_row_header and header_fg_color is not None:
                    text_cursor = cell_text.createTextCursor()
                    mProps.Props.set(text_cursor, CharColor=header_fg_color)

                cell_text.setString(str(data))

            def set_cell_text(cell_name: str, data: str, table: XTextTable) -> None:
                cell_text = mLo.Lo.qi(XText, table.getCellByName(cell_name), True)
                if first_row_header is False or tbl_fg_color is not None:
                    text_cursor = cell_text.createTextCursor()
                    props = {}
                    if not first_row_header:
                        # By default the first row has a style by the name of: Table Heading
                        # Table Contents is the default for cell that are not in the header row.
                        props["ParaStyleName"] = "Table Contents"
                    if tbl_fg_color is not None:
                        props["CharColor"] = tbl_fg_color
                    mProps.Props.set(text_cursor, **props)

                cell_text.setString(str(data))

            num_rows = len(table_data)
            if num_rows == 0:
                raise ValueError("table_data has no values")
            try:
                table = mLo.Lo.create_instance_msf(XTextTable, "com.sun.star.text.TextTable")
                if table is None:
                    raise ValueError("Null Value")
            except Exception as e:
                raise mEx.CreateInstanceMsfError(XTextTable, "com.sun.star.text.TextTable")

            try:
                num_cols = len(table_data[0])
                mLo.Lo.print(f"Creating table rows: {num_rows}, cols: {num_cols}")
                table.initialize(num_rows, num_cols)

                # insert the table into the document
                cls._append_text_content(cursor, table)
                cls.end_paragraph(cursor)

                table_props = mLo.Lo.qi(XPropertySet, table, True)

                # set table properties
                if header_bg_color is not None or tbl_bg_color is not None:
                    table_props.setPropertyValue("BackTransparent", False)  # not transparent
                if tbl_bg_color is not None:
                    table_props.setPropertyValue("BackColor", tbl_bg_color)

                # set color of first row (i.e. the header)
                if first_row_header and header_bg_color is not None:
                    rows = table.getRows()
                    mProps.Props.set(rows.getByIndex(0), BackColor=header_bg_color)

                #  write table header
                if first_row_header:
                    row_data = table_data[0]
                    for x in range(num_cols):
                        set_cell_header(make_cell_name(0, x), row_data[x], table)
                        # e.g. "A1", "B1", "C1", etc

                    # insert table body
                    for y in range(1, num_rows):  # start in 2nd row
                        row_data = table_data[y]
                        for x in range(num_cols):
                            set_cell_text(make_cell_name(y, x), row_data[x], table)
                else:
                    # insert table body
                    for y in range(0, num_rows):  # start in 1st row
                        row_data = table_data[y]
                        for x in range(num_cols):
                            set_cell_text(make_cell_name(y, x), row_data[x], table)

                if styles:
                    srv = ("com.sun.star.text.TextTable",)
                    for style in styles:
                        if style.support_service(*srv):
                            style.apply(table)
            except Exception as e:
                raise Exception("Table insertion failed:") from e
            _Events().trigger(WriteNamedEvent.TABLE_ADDED, EventArgs.from_args(cargs))
            return table

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

A TextTable_ service with an XTextTable_ interface is created at the start of :py:meth:`~.Write.add_table`.
Then the required number of rows and columns is calculated so that ``XTextTable.initialize()`` can be called to specify the table's dimensions.

.. tabs::

    .. code-tab:: python

        num_rows = len(table_data)
        ...

        # use the first row to get the number of column
        num_cols = len(table_data[0])
        Lo.print(f"Creating table rows: {num_rows}, cols: {num_cols}")
        table.initialize(num_rows, num_cols)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Table-wide properties are set (properties are listed in the TextTable_ documentation).
Note that if "BackTransparent" isn't set to false then Office crashes when the program tries to save the document.

The color property of the header row is set to dark blue (:py:attr:`.CommonColor.DARK_BLUE`) by default.
This requires a call to ``XTextTable.getRows()`` to return an XTableRows_ object representing all the rows.
This object inherits XIndexAccess_, so the first row is accessed with index 0.

.. tabs::

    .. code-tab:: python

        # set color of first row (i.e. the header)
        if header_bg_color is not None:
            rows = table.getRows()
            Props.set_property(prop_set=rows.getByIndex(0), name="BackColor", value=header_bg_color)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The filling of the table with data is performed by two loops.
The first deals with adding text to the header row, the second deals with all the other rows.

``make_cell_name()`` converts an (x, y) integer pair into a cell name like those in :numref:`ch07fig_cell_name_tbl_ss`:

``make_cell_name()`` uses :py:class:`~.table_helper.TableHelper` methods to make the conversion.

:py:meth:`.Write.set_cell_header` uses ``TextTable.getCellByName()`` to access a cell, which is of type XCell_.
We'll study XCell_ in :ref:`part04` because it's used for representing cells in a spreadsheet.

The Cell service supports both the XCell_ and XText_ interfaces, as in :numref:`ch07fig_cell_service`.

.. cssclass:: diagram invert

    .. _ch07fig_cell_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/185226758-28d3b90c-32d4-498b-92e7-31a63194c0f2.png
        :alt: Diagram of The Cell Service
        :figclass: align-center

        :The Cell Service.

This means that :py:meth:`.Lo.qi` can convert an XCell_ instance into XText_,
which makes the cell's text and properties accessible to a text cursor.
``set_cell_header()`` implements these features:

.. tabs::

    .. code-tab:: python

        # in Write Class
        def set_cell_header(cell_name: str, data: str, table: XTextTable) -> None:
            cell_text = mLo.Lo.qi(XText, table.getCellByName(cell_name), True)
            if first_row_header and header_fg_color is not None:
                text_cursor = cell_text.createTextCursor()
                mProps.Props.set(text_cursor, CharColor=header_fg_color)

            cell_text.setString(str(data))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The cell's ``CharColor`` property is changed so the inserted text in the header row is white (:py:attr:`.CommonColor.WHITE`) by default, as in :numref:`ch07fig_bond_movie_ss`.

``set_cell_text()`` like ``set_cell_header()`` optionally changes the text's color:

.. tabs::

    .. code-tab:: python

        # in Write Class
        def set_cell_text(cell_name: str, data: str, table: XTextTable) -> None:
            cell_text = mLo.Lo.qi(XText, table.getCellByName(cell_name), True)
            if first_row_header is False or tbl_fg_color is not None:
                text_cursor = cell_text.createTextCursor()
                props = {}
                if not first_row_header:
                    # By default the first row has a style by the name of: Table Heading
                    # Table Contents is the default for cell that are not in the header row.
                    props["ParaStyleName"] = "Table Contents"
                if tbl_fg_color is not None:
                    props["CharColor"] = tbl_fg_color
                mProps.Props.set(text_cursor, **props)

            cell_text.setString(str(data))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch07_add_bookmark:

7.6 Adding a Bookmark to the Document
=====================================

:py:meth:`.Write.add_bookmark` adds a named bookmark at the current cursor position:

.. tabs::

    .. code-tab:: python

        # in Write Class
        @classmethod
        def add_bookmark(cls, cursor: XTextCursor, name: str) -> None:
            cargs = CancelEventArgs(Write.add_bookmark.__qualname__)
            cargs.event_data = {"cursor": cursor, "name": name}
            _Events().trigger(WriteNamedEvent.BOOKMARK_ADDING, cargs)
            if cargs.cancel:
                return False

            # get name from event args in case it has been changed.
            name = cargs.event_data["name"]

            try:
                bmk_content = Lo.create_instance_msf(XTextContent, "com.sun.star.text.Bookmark")
                if bmk_content is None:
                    raise ValueError("Null Value")
            except Exception as e:
                raise CreateInstanceMsfError(XTextContent, "com.sun.star.text.Bookmark") from e
            try:
                bmk_named = Lo.qi(XNamed, bmk_content, True)
                bmk_named.setName(name)

                cls._append_text_content(cursor, bmk_content)
            except Exception as e:
                raise Exception("Unable to add bookmark") from e
            _Events().trigger(WriteNamedEvent.BOOKMARK_ADDED, EventArgs.from_args(cargs))
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The Bookmark_ service doesn't have a specific interface (such as ``XBookmark``), so :py:meth:`.Lo.create_instance_msf` returns an XTextContent_ interface.
These services and interfaces are summarized by :numref:`ch07fig_bookmark_service`.

.. cssclass:: diagram invert

    .. _ch07fig_bookmark_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/185230953-72690b77-d5eb-4c89-80f7-2ddf6be56b5a.png
        :alt: Diagram of The Bookmark Service and Interfaces
        :figclass: align-center

        :The Bookmark Service and Interfaces.

Bookmark_ supports XNamed_, which allows it to be viewed as a named collection of bookmarks (note the plural).
This is useful when searching for a bookmark or adding one, as in the |build_doc|_ example.
It calls :py:meth:`.Write.add_bookmark` to add a bookmark called ``ad-Bookmark`` to the document:

.. tabs::

    .. code-tab:: python

        # code fragment from build doc
        cursor.append("This line ends with a bookmark.")
        cursor.add_bookmark("ad-bookmark")
        cursor.append_line()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Bookmarks, such as ``ad-bookmark``, are not rendered when the document is opened,
which means that nothing appears after the "The line ends with a bookmark." string in "build.odt".

However, bookmarks are listed in Writer's "Navigator" window (press F5), as in :numref:`ch07fig_writer_nav_ss`.

.. cssclass:: screen_shot invert

    .. _ch07fig_writer_nav_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/185232660-d80c79e0-1992-4b45-84d1-e0766f2c6817.png
        :alt: Screen shot The Writer Navigator Window
        :figclass: align-center

        :The Writer Navigator Window.

Clicking on the bookmark causes Writer to jump to its location in the document.

Using Bookmarks; One programming use of bookmarks is for moving a cursor around a document.
Just as with real-world bookmarks, you can add one at some important location in a document and jump to that position at a later time.

:py:meth:`.Write.find_bookmark` finds a bookmark by name, returning it as an XTextContent_ instance:

.. tabs::

    .. code-tab:: python

        # in Write Class
        @staticmethod
        def find_bookmark(text_doc: XTextDocument, bm_name: str) -> XTextContent | None:
            supplier = Lo.qi(XBookmarksSupplier, text_doc, True)

            named_bookmarks = supplier.getBookmarks()
            obookmark = None

            try:
                obookmark = named_bookmarks.getByName(bm_name)
            except Exception:
                Lo.print(f"Bookmark '{bm_name}' not found")
                return None
            return Lo.qi(XTextContent, obookmark)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`~.Write.find_bookmark` can't return an ``XBookmark`` object since there's no such interface (see :numref:`ch07fig_bookmark_service`),
but XTextContent_ is a good alternative. XTextContent_ has a ``getAnchor()`` method which returns an XTextRange_ that can be used for positioning a cursor.
The following code fragment from |build_doc|_ illustrates the idea:

.. tabs::

    .. code-tab:: python

        # code fragment form build doc
        # move view cursor to bookmark position
        bookmark = doc.find_bookmark("ad-bookmark")
        bm_range = bookmark.get_anchor()

        view_cursor = doc.get_view_cursor()
        view_cursor.goto_range(bm_range, False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The call to ``gotoRange()`` moves the view cursor to the ``ad-bookmark`` position, which causes an on-screen change.
``gotoRange()`` can be employed with any type of cursor.

.. |build_doc| replace:: Build Doc
.. _build_doc: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_doc

.. |math_ques| replace:: Math Questions
.. _math_ques: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_math_questions

.. |story_creator| replace:: Story Creator
.. _story_creator: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_story_creator

.. |make_table| replace:: Make Table
.. _make_table: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_make_table

.. _BaseFrameProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1BaseFrameProperties.html
.. _Bookmark: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1Bookmark.html
.. _DateTime: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1textfield_1_1DateTime.html
.. _GenericTextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1GenericTextDocument.html
.. _GraphicObjectShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GraphicObjectShape.html
.. _OfficeDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1OfficeDocument.html
.. _TextContent: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextContent.html
.. _TextContentAnchorType: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1text.html#a470b1caeda4ff15fee438c8ff9e3d834
.. _TextEmbeddedObject: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextEmbeddedObject.html
.. _TextField: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextField.html
.. _TextFrame: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextFrame.html
.. _TextGraphicObject: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextGraphicObject.html
.. _TextTable: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextTable.html
.. _XCell: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCell.html
.. _XEmbeddedObjectSupplier2: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1document_1_1XEmbeddedObjectSupplier2.html
.. _XIndexAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexAccess.html
.. _XNameAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameAccess.html
.. _XNamed: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNamed.html
.. _XPropertySet: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertySet.html
.. _XShape: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShape.html
.. _XTableRows: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XTableRows.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html
.. _XTextContent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextContent.html
.. _XTextCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextCursor.html
.. _XTextField: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextField.html
.. _XTextFrame: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextFrame.html
.. _XTextRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextRange.html
.. _XTextTable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextTable.html
