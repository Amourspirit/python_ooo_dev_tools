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

The example code is in |math_ques|_, but most of the formula embedding is performed by :py:meth:`.Write.add_formula`:

.. tabs::

    .. code-tab:: python

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

A math formula is passed to :py:meth:`~.Write.add_formula` as a string in a format this is explained shortly.

The method begins by creating a TextEmbeddedObject_ service, and referring to it using the XTextContent_ interface:

.. tabs::

    .. code-tab:: python

        embed_content = Lo.create_instance_msf(
                XTextContent, "com.sun.star.text.TextEmbeddedObject", raise_err=True
            )

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

The properties for this empty object (embed_obj_model) are accessed, and the formula string is assigned to the "Formula" property:

.. tabs::

    .. code-tab:: python

        formula_props = Lo.qi(XPropertySet, embed_obj_model, True)
        formula_props.setPropertyValue("Formula", formula)

7.3.1 What's a Formula String?
------------------------------

Although the working of :py:meth:`.Write.add_formula` has been explained, the format of the formula string that's passed to it has not been explained.
There's a good overview of the notation in the "Commands Reference" appendix of Office's "Math Guide", available at https://libreoffice.org/get-help/documentation
For example, the formula string: "1 {5}over{9} + 3 {5}over{9} = 5 {1}over{9}" is rendered as:

.. math::

   1 \frac{5}{9} + 3 \frac{5}{9} = 5 \frac{1}{9}

7.3.2 Building Formulae
-----------------------

|math_ques|_ is mainly a for-loop for randomly generating numbers and constructing simple formulae strings.
Ten formulae are added to the document, which is saved as ``mathQuestions.pdf``. The ``main()`` function:

.. tabs::

    .. code-tab:: python

        def main() -> int:

            delay = 2_000  # delay so users can see changes.

            with Lo.Loader(Lo.ConnectSocket()) as loader:

                doc = Write.create_doc(loader=loader)

                try:
                    GUI.set_visible(is_visible=True, odoc=doc)

                    cursor = Write.get_cursor(doc)
                    Write.append_para(cursor, "Math Questions")
                    Write.style_prev_paragraph(cursor, "Heading 1")

                    Write.append_para(cursor, "Solve the following formulae for x:\n")

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

                            # formulas should be wrapped in {} but for fromatting reasons it is easier to work with [] and replace later.
                            if choice == 0:
                                formula = f"[[[sqrt[{iA}x]] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1} ]]"
                            elif choice == 1:
                                formula = f"[[[{iA}x] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1}]]"
                            else:
                                formula = f"[{iA}x + {iB} = {iC}]"

                            # replace [] with {}
                            Write.add_formula(cursor, formula.replace("[", "{").replace("]", "}"))
                            Write.end_paragraph(cursor)

                    Write.append_para(cursor, f"Timestamp: {DateUtil.time_stamp()}")

                    Lo.delay(delay)
                    Lo.save_doc(doc, "mathQuestions.pdf")

                finally:
                    Lo.close_doc(doc)

            return 0

:numref:`ch07fig_math_formula_ss` shows a screenshot of part of ``mathQuestions.pdf``.

.. cssclass:: screen_shot invert

    .. _ch07fig_math_formula_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184988764-6c2891eb-bf2d-4fc5-bc38-1a99b08f06dc.png
        :alt: Screen shot of Math Formulae in a Text Document
        :figclass: align-center

        :Math Formulae in a Text Document.

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

7.4.1 The DateTime TextField
----------------------------

The build_doc|_ example ends with a few lines that appear to do the same thing twice:

.. tabs::

    .. code-tab:: python

        # code fragment from build doc
        Write.append_para(cursor, "\nTimestamp: " + DateUtil.time_stamp() + "\n")
        Write.append(cursor, "Time (according to office): ")
        Write.append_date_time(cursor=cursor)
        Write.end_paragraph(cursor)

:py:meth:`.DateUtil.time_stamp` inserts a timestamp (which includes the date and time), and then :py:meth:`.Write.append_date_time` inserts the date and time.
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

The method adds two DateTime text fields to the document.
The first has its "IsDate" property set to true, so that the current date is inserted; the second sets "IsDate" to false so the current time is shown.

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

These properties will be applied to the text and text fields added afterwards:

.. tabs::

    .. code-tab:: python

        Write.append(footer_cursor, Write.get_page_number())
        Wirte.append(footer_cursor, " of ")
        Write.append(footer_cursor, Write.get_page_count())

:py:meth:`~.Write.get_page_number` and :py:meth:`~.Write.get_page_count` deal with the properties for the PageNumber and PageCount fields.

Work in progress ...

.. |build_doc| replace:: Build Doc
.. _build_doc: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_doc

.. |math_ques| replace:: Math Questions
.. _math_ques: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_math_quesions

.. |story_creator| replace:: Story Creator
.. _story_creator: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_story_creator

.. _BaseFrameProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1BaseFrameProperties.html
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
.. _XEmbeddedObjectSupplier2: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1document_1_1XEmbeddedObjectSupplier2.html
.. _XNameAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameAccess.html
.. _XPropertySet: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertySet.html
.. _XShape: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShape.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html
.. _XTextContent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextContent.html
.. _XTextCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextCursor.html
.. _XTextField: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextField.html
.. _XTextFrame: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextFrame.html