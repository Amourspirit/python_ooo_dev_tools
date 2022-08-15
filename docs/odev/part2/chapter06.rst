.. _ch06:

**********************
Chapter 6. Text Styles
**********************

.. topic:: Overview

    Five Style Families; Properties; Listing Styles; Creating a Style; Applying Styles;
    Paragraph/Word Styles; Hyperlink Styling; Text Numbering; Headers and Footers

This chapter focuses on how text documents styles can be examined and manipulated.
This revolves around the XStyleFamiliesSupplier_ interface in GenericTextDocument_, which is highlighted in
:numref:`ch06fig_txt_doc_serv_interfaces` (a repeat of :numref:`ch05fig_txt_doc_serv_interfaces` in :ref:`ch05`).

.. cssclass:: diagram invert

    .. _ch06fig_txt_doc_serv_interfaces:
    .. figure:: https://user-images.githubusercontent.com/4193389/181575340-96fb7e21-4e0f-4662-8ed9-92edfb036b0c.png
        :alt: Diagram of The Text Document Services, and some Interfaces
        :figclass: align-center

        :The Text Document Services, and some Interfaces.

XStyleFamiliesSupplier_ has a ``getStyleFamilies()`` method for returning text style families.
All these families are stored in an XNameAccess object, as depicted in :numref:`ch06fig_txt_doc_serv_interfaces`.

XNameAccess_ is one of Office's collection types, and employed when the objects in a collection have names.
There's also an XIndexAccess_ for collections in index order.

XNameContainer_ and XIndexContainer_ add the ability to insert and remove objects from a collection.

.. cssclass:: diagram invert

    .. _ch06fig_style_fam_props:
    .. figure:: https://user-images.githubusercontent.com/4193389/184508794-8a5d0cda-82af-46a1-9da1-125dc73f4c0d.png
        :alt: Diagram of Style Families and their Property Sets
        :figclass: align-center

        :Style Families and their Property Sets.

Five style family names are used by text documents:
``CharacterStyles``, ``FrameStyles``, ``NumberingStyles``, ``PageStyles``, and ``ParagraphStyles``.
The XNameAccess_ families collection can be accessed with one of those names, and returns a style family.
A style family is a modifiable collection of ``PropertySet`` objects, stored in an XNameContainer_ object.

:numref:`ch06fig_style_fam_props` shows that if the ``ParagraphStyles`` family is retrieved, it contains property sets labeled
"Header", "List", "Standard", and quite a few more.
Each property set can format a paragraph, change the text's font, size, and many other attributes.
These property sets are called styles.

The ``CharacterStyles`` family is a container of property sets (styles) which affect selected sentences, words, or characters in the document.
The ``FrameStyles`` container holds property sets (styles) for formatting graphic and text frames.
The ``NumberingStyles`` family is for adding numbers or bullets to paragraphs.
The ``PageStyles`` family is for formatting pages.

The names of the property sets (styles) in the style families can be listed using LibreOffice's GUI.
If you create a new text document in Writer, a "Styles and Formatting" dialog window appears when you press F11
(or click on the brown spanner icon in the "Formatting" toolbar).
Within the window you can switch between five icons representing the five style families. :numref:`ch06fig_writer_style_ss` shows the list
of property set (style) names for the paragraph styles family.
They corresponds to the property set names shown in :numref:`ch06fig_style_fam_props`.

.. cssclass:: screen_shot invert

    .. _ch06fig_writer_style_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184509794-4be89e59-d5fe-4d78-b2f1-db689060f802.png
        :width: 300px
        :alt: Screen shot of Styles and Formatting Window in Writer
        :figclass: align-center

        :Styles and Formatting Window in Writer.

The names listed in the window are the same as the names used in the API, except in two cases:
the "Default Paragraph Style" name that appears in the GUI window for "Paragraph Styles" and "Page Styles" is changed to "Standard" in the API.
Strangely, the "Default Style" name for "Character Styles" in the GUI is called "Default Style" in the API.

Accessing a style (a property set) is a three-step process, shown below.
First the style families, then the style family (:abbreviation:`ex:` "ParagraphStyle"), and then the style (:abbreviation:`ex:` "Standard"):

.. tabs::

    .. code-tab:: python

        # 1. get the style families
        xsupplier = Lo.qi(XStyleFamiliesSupplier, doc)
        name_acc = xsupplier.getStyleFamilies()

        # 2. get the paragraph style family
        para_style_con = Lo.qi(XNameContainer, name_acc.getByName("ParagraphStyles"))

        # 3. get the 'standard' style (property set)
        standard_props = Lo.qi(XPropertySet, para_style_con.getByName("Standard"))

The code that implements this process in the Write utility class is a bit more complicated since the calls to
``getByName()`` may raise exceptions if their string arguments are incorrect.

The calls to :py:meth:`.Lo.qi` cast the object returned from a collection into the correct type.

6.1 What Properties are in a PropertySet?
=========================================

The "Standard" name in the "ParagraphStyles" style family refers to a property set (style).
Each set is a collection of ``name=value`` pairs, and there are get and set methods using a name to get/set its value.
This is simple enough, but what names should the programmer use?
Each property set (style) in the same style family contain the same properties, but with different values.
For instance, in :numref:`ch06fig_style_fam_props` the "Header", "Title", "Standard", "List", and "Table" sets contain the same named properties.

The names of the properties used by the sets in a style family can be found in the documentation for their ``XXXStyle`` service.
:numref:`ch06tbl_syle_prop_info` summarizes the mapping.

.. _ch06tbl_syle_prop_info:

.. table:: Properties Information for Each Style Family.
    :name: syle_prop_info

    ====================== =======================================
    Style Family Name      Service where Properties are Defined
    ====================== =======================================
    ``CharacterStyles``    ``CharacterStyle``
    ``FrameStyles``        ``FrameStyle`` (??)
    ``NumberingStyles``    ``NumberingStyle``
    ``PageStyles``         ``PageStyle``
    ``ParagraphStyles``    ``ParagraphStyle``
    ====================== =======================================

The easiest way of finding Office documentation for the services in the second column of :numref:`ch06tbl_syle_prop_info` is with ``lodoc``.
For example, the page about "CharacterStyle" can be found with ``lodoc CharacterStyle service``.

The ``FrameStyle`` service (full name: ``com.sun.star.style.FrameStyle``) has a "??" against it since there's no online documentation for that service, although such a service exists.

A style's properties are usually defined across several classes in an inheritance hierarchy.
The hierarchies for the five styles are summarized in :numref:`ch06fig_style_inheritance`.

.. cssclass:: diagram invert

    .. _ch06fig_style_inheritance:
    .. figure:: https://user-images.githubusercontent.com/4193389/184510722-272d8e0e-bb4d-4f51-9c97-9b60af40a9d5.png
        :alt: Diagram of The Inheritance Hierarchies for the Style Services.
        :figclass: align-center

        :The Inheritance Hierarchies for the Style Services.

:numref:`ch06fig_style_inheritance` shows the hierarchies for the five style services: ``CharacterStyle``, ``FrameStyle``, ``NumberingStyle``, ``PageStyle``, and ``ParagraphStyle``.
There's clearly a lot of similarities between them, so we are focused on ``CharacterStyle``.

There are three services containing character style properties: ``CharacterStyle``, ``Style``, and ``CharacterProperties``.
If you visit the online documentation for CharacterStyle, the properties are listed under the heading "Public Attributes", which is shown in :numref:`ch06fig_docs_char_style_ss`.

.. cssclass:: screen_shot invert

    .. _ch06fig_docs_char_style_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184510828-8bebec21-aae8-4898-b705-889b5cafb98a.png
        :alt: Screen shot of Styles and Formatting Window in Writer
        :figclass: align-center

        :Part of the Online Documentation for CharacterStyle.

``CharacterStyle`` defines six properties itself, but there are many more inherited from the Style and ``CharacterProperties`` services.
If you click on the triangles next to the "Public Attributes inherited from" lines, the documentation expands to display those properties.

:numref:`ch06fig_style_inheritance` contains two "(??)" strings – one is to indicate that there's no documentation for ``FrameStyle``,
so it is a guess about its inheritance hierarchy.

The other "(??)" is in the ``ParagraphStyle`` hierarchy. The documentation for ``ParagraphStyle``, and the information in the developers guide,
indicate that ParagraphStyle inherits only Style and ParagraphCharacter.
We believe this to be incorrect, based on my coding with ``ParagraphStyle`` (some of which you'll see in the next sections).
ParagraphStyle appears to inherits three services: Style, ParagraphCharacter, and CharacterStyle, as indicated in :numref:`ch06fig_para_serv_supers`.

.. cssclass:: diagram invert

    .. _ch06fig_para_serv_supers:
    .. figure:: https://user-images.githubusercontent.com/4193389/184510955-125605d0-079c-4935-ade4-9d24065ed122.png
        :alt: Diagram of The Paragraph Service and its Superclasses
        :figclass: align-center

        :The Paragraph Service and its Super-classes.

For more information of the styles API, start in the development guide in the "Overall Document Features" section,
online at: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Overall_Document_Features

6.2 Listing Styles Information
==============================

The |styles_info|_ example illustrates some of the Writer and Info utility functions for examining style families and their property sets.
The ``show_styles()`` function starts by listing the style families names:

.. tabs::

    .. code-tab:: python

        def show_styles(doc: XTextDocument) -> None:
            # get all the style families for this document
            style_families = Info.get_style_family_names(doc)
            print(f"No. of Style Family Names: {len(style_families)}")
            for style_family in style_families:
                print(f"  {style_family}")
            print()

            # list all the style names for each style family
            for i, style_family in enumerate(style_families):
                print(f'{i} "{style_family}" Style Family contains containers:')
                style_names = Info.get_style_names(doc, style_family)
                Lo.print_names(style_names)

            # Report the properties for the paragraph styles family under the "Standard" name
            Props.show_props('ParagraphStyles "Standard"', Info.get_style_props(doc, "ParagraphStyles", "Header"))
            print()

Partial output lists the seven family names:

::

    No. of Style Family Names: 7
        CellStyles
        CharacterStyles
        FrameStyles
        NumberingStyles
        PageStyles
        ParagraphStyles
        TableStyles

:py:meth:`.Info.get_style_names` starts by calling :py:meth:`.Info.get_style_container` which in turn calls
:py:meth:`.Info.get_style_families`.
``get_style_families()`` gets XStyleFamiliesSupplier_ that is passed to ``get_style_container()``
which in turn gets XNameContainer_ that is passed to ``get_style_names()``.
The family names in that collection are extracted with ``style_container.getElementNames()``:

.. tabs::

    .. code-tab:: python

        @staticmethod
        def get_style_families(doc: object) -> XNameAccess:
            try:
                xsupplier = Lo.qi(XStyleFamiliesSupplier, doc, True)
                return xsupplier.getStyleFamilies()
            except MissingInterfaceError:
                raise
            except Exception as e:
                raise Exception("Unable to get family style names") from e

        @classmethod
        def get_style_container(cls, doc: object, family_style_name: str) -> XNameContainer:
            name_acc = cls.get_style_families(doc)
            xcontianer = Lo.qi(XNameContainer, name_acc.getByName(family_style_name), True)
            return xcontianer

        @classmethod
        def get_style_names(cls, doc: object, family_style_name: str) -> List[str]:
            try:
                style_container = cls.get_style_container(doc=doc, family_style_name=family_style_name)
                names = style_container.getElementNames()
                lst = list(names)
                lst.sort()
                return lst
            except Exception as e:
                raise Exception("Could not access style names") from e

|styles_info|_ example, the ``show_styles()`` function continues by looping through the list of style family names,
printing all the style (property set) names in each family:

.. tabs::

    .. code-tab:: python

        # list all the style names for each style family
        for i, style_family in enumerate(style_families):
            print(f'{i} "{style_family}" Style Family contains containers:')
            style_names = Info.get_style_names(doc, style_family)
            Lo.print_names(style_names)

The output is lengthy, but informative:

::

    0 "CellStyles" Style Family contains containers:
    No. of names: 0


    1 "CharacterStyles" Style Family contains containers:
    No. of names: 27
      'Bullet Symbols'  'Caption characters'  'Citation'  'Definition'
      'Drop Caps'  'Emphasis'  'Endnote anchor'  'Endnote Symbol'
      'Example'  'Footnote anchor'  'Footnote Symbol'  'Index Link'
      'Internet link'  'Line numbering'  'Main index entry'  'Numbering Symbols'
      'Page Number'  'Placeholder'  'Rubies'  'Source Text'
      'Standard'  'Strong Emphasis'  'Teletype'  'User Entry'
      'Variable'  'Vertical Numbering Symbols'  'Visited Internet Link'

    2 "FrameStyles" Style Family contains containers:
    No. of names: 7
      'Formula'  'Frame'  'Graphics'  'Labels'
      'Marginalia'  'OLE'  'Watermark'

    3 "NumberingStyles" Style Family contains containers:
    No. of names: 11
      'List 1'  'List 2'  'List 3'  'List 4'
      'List 5'  'No List'  'Numbering 123'  'Numbering ABC'
      'Numbering abc'  'Numbering IVX'  'Numbering ivx'

    4 "PageStyles" Style Family contains containers:
    No. of names: 10
      'Endnote'  'Envelope'  'First Page'  'Footnote'
      'HTML'  'Index'  'Landscape'  'Left Page'
      'Right Page'  'Standard'

    5 "ParagraphStyles" Style Family contains containers:
    No. of names: 125
      'Addressee'  'Appendix'  'Bibliography 1'  'Bibliography Heading'
      'Caption'  'Contents 1'  'Contents 10'  'Contents 2'
      'Contents 3'  'Contents 4'  'Contents 5'  'Contents 6'
      'Contents 7'  'Contents 8'  'Contents 9'  'Contents Heading'
      'Drawing'  'Endnote'  'Figure'  'Figure Index 1'
      'Figure Index Heading'  'First line indent'  'Footer'  'Footer left'
      'Footer right'  'Footnote'  'Frame contents'  'Hanging indent'
      'Header'  'Header and Footer'  'Header left'  'Header right'
      'Heading'  'Heading 1'  'Heading 10'  'Heading 2'
      'Heading 3'  'Heading 4'  'Heading 5'  'Heading 6'
      'Heading 7'  'Heading 8'  'Heading 9'  'Horizontal Line'
      'Illustration'  'Index'  'Index 1'  'Index 2'
      'Index 3'  'Index Heading'  'Index Separator'  'List'
      'List 1'  'List 1 Cont.'  'List 1 End'  'List 1 Start'
      'List 2'  'List 2 Cont.'  'List 2 End'  'List 2 Start'
      'List 3'  'List 3 Cont.'  'List 3 End'  'List 3 Start'
      'List 4'  'List 4 Cont.'  'List 4 End'  'List 4 Start'
      'List 5'  'List 5 Cont.'  'List 5 End'  'List 5 Start'
      'List Contents'  'List Heading'  'List Indent'  'Marginalia'
      'Numbering 1'  'Numbering 1 Cont.'  'Numbering 1 End'  'Numbering 1 Start'
      'Numbering 2'  'Numbering 2 Cont.'  'Numbering 2 End'  'Numbering 2 Start'
      'Numbering 3'  'Numbering 3 Cont.'  'Numbering 3 End'  'Numbering 3 Start'
      'Numbering 4'  'Numbering 4 Cont.'  'Numbering 4 End'  'Numbering 4 Start'
      'Numbering 5'  'Numbering 5 Cont.'  'Numbering 5 End'  'Numbering 5 Start'
      'Object index 1'  'Object index heading'  'Preformatted Text'  'Quotations'
      'Salutation'  'Sender'  'Signature'  'Standard'
      'Subtitle'  'Table'  'Table Contents'  'Table Heading'
      'Table index 1'  'Table index heading'  'Text'  'Text body'
      'Text body indent'  'Title'  'User Index 1'  'User Index 10'
      'User Index 2'  'User Index 3'  'User Index 4'  'User Index 5'
      'User Index 6'  'User Index 7'  'User Index 8'  'User Index 9'
      'User Index Heading'

    6 "TableStyles" Style Family contains containers:
    No. of names: 11
      'Academic'  'Box List Blue'  'Box List Green'  'Box List Red'
      'Box List Yellow'  'Default Style'  'Elegant'  'Financial'
      'Simple Grid Columns'  'Simple Grid Rows'  'Simple List Shaded'

:py:meth:`.Info.get_style_names` retrieves the XNameContainer_ object for each style family,
and extracts its style (property set) names using ``getElementNames()``:

.. tabs::

    .. code-tab:: python

        @classmethod
        def get_style_names(cls, doc: object, family_style_name: str) -> List[str]:
            try:
                style_container = cls.get_style_container(doc=doc, family_style_name=family_style_name)
                names = style_container.getElementNames()
                lst = list(names)
                lst.sort()
                return lst
            except Exception as e:
                raise Exception("Could not access style names") from e

The last part of |styles_info|_ lists the properties for a specific property set. :py:meth:`.Info.get_style_props` does that:

.. tabs::

    .. code-tab:: python

        @classmethod
        def get_style_props(cls, doc: object, family_style_name: str, prop_set_nm: str) -> XPropertySet:
            style_container = cls.get_style_container(doc, family_style_name)
            name_props = Lo.qi(XPropertySet, style_container.getByName(prop_set_nm), True)
            return name_props

Its arguments are the document, the style family name, and style (property set) name.

A reference to the property set is returned. Accessing the "Standard" style (property set) of the "ParagraphStyle" family would require:

.. tabs::

    .. code-tab:: python

        props = Info.get_style_props(doc, "ParagraphStyles", "Standard")

The property set can be nicely printed by calling :py:meth:`.Props.show_props`:

.. tabs::

    .. code-tab:: python

        Props.show_props('ParagraphStyles "Standard"', props)

The output is long, but begins and ends like so:

::

    ParagraphStyles "Standard" Properties
        BorderDistance: 0
        BottomBorder: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
        BottomBorderDistance: 0
        BreakType: <Enum instance com.sun.star.style.BreakType ('NONE')>
        Category: 4
        CharAutoKerning: True
        CharBackColor: -1
        CharBackTransparent: True
            :
        Rsid: 0
        SnapToGrid: True
        StyleInteropGrabBag: ()
        TopBorder: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
        TopBorderDistance: 0
        WritingMode: 4

This listing, and in fact any listing of a style from "ParagraphStyles",
shows that the properties are a mixture of those defined in the Style,
ParagraphProperties_, and CharacterProperties_ services.

6.3 Creating a New Style
========================

The |story_creator|_ example adds a new style to the paragraph style family, and uses it to format the document's paragraphs.

The new ParagraphStyle_ service is referenced using one of its interfaces, the usual one being XStyle_ since all the different
style services support it (as shown in :numref:`ch06fig_style_inheritance`).

.. tabs::

    .. code-tab:: python

        xtext_range = doc.getText().getStart()
        Props.set_property(xtext_range, "ParaStyleName", "adParagraph")

        Write.set_header(text_doc=doc, text=f"From: {fnm.name}")
        Write.set_a4_page_format(doc)
        Write.set_page_numbers(doc)

        cursor = Write.get_cursor(doc)

        read_text(fnm=fnm, cursor=cursor)
        Write.end_paragraph(cursor)

``read_text()`` assumes the text file has a certain format. For example, ``scandal.txt`` begins like so:

::

    Title: A Scandal in Bohemia
    Author: Sir Arthur Conan Doyle

    Part I.


    To Sherlock Holmes she is always THE woman. I have seldom heard
    him mention her under any other name. In his eyes she eclipses
    and predominates the whole of her sex.


    It was not that he felt any emotion akin to love for Irene Adler.

    All emotions, and that one particularly, were abhorrent to his
    cold, precise but admirably balanced mind.

A paragraph is a series of text lines followed by a blank line. But there are exceptions: lines that starts with "Title: ", "Author: " or "Part "
are treated as headings, and styled differently. When the text above is processed, the resulting document looks like :numref:`ch06fig_story_creator_out_ss`.

.. cssclass:: screen_shot invert

    .. _ch06fig_story_creator_out_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184560774-db82b140-3f9e-4f10-abd1-031c649bbac8.png
        :alt: Screen Shot of The Output of Story Creator Example
        :figclass: align-center

        :The Output of Story Creator Example.

``read_text()`` is implemented using python's ``with open(fnm, 'r') as file:`` context manager:

.. tabs::

    .. code-tab:: python

        def read_text(fnm: Path, cursor: XTextCursor) -> None:
            sb: List[str] = []
            with open(fnm, 'r') as file:
                i = 0
                for ln in file:
                    line = ln.rstrip() # remove new line \n
                    if len(line) == 0:
                        if len(sb) > 0:
                            Write.append_para(cursor, ' '.join(sb))
                        sb.clear()
                    elif line.startswith("Title: "):
                        Write.append_para(cursor, line[7:])
                        Write.style_prev_paragraph(cursor, "Title")
                    elif line.startswith("Author: "):
                        Write.append_para(cursor, line[8:])
                        Write.style_prev_paragraph(cursor, "Subtitle")
                    elif line.startswith("Part "):
                        Write.append_para(cursor, line)
                        Write.style_prev_paragraph(cursor, "Heading")
                    else:
                        sb.append(line)
                    i += 1
                    # if i > 20:
                    #     break
                if len(sb) > 0:
                    Write.append_para(cursor, ' '.join(sb))

The interesting bits are the calls to :py:meth:`.Write.append_para` and :py:meth:`.Write.style_prev_paragraph` which add a paragraph to the document and apply a style to it.
For instance:

.. tabs::

    .. code-tab:: python

        elif line.startswith("Author: "):
            Write.append_para(cursor, line[8:])
            Write.style_prev_paragraph(cursor, "Subtitle")

:py:meth:`~.Write.append_para` writes the string into the document as a paragraph (the input line without the "Author: " substring).
:py:meth:`~.Write.style_prev_paragraph` changes the paragraph style from ``adParagraph`` to ``Subtitle``.

The hard part of :py:meth:`~.Write.style_prev_paragraph` is making sure that the style change only affects the previous paragraph.
Text appended after this line should use ``adParagraph`` styling.

.. tabs::

    .. code-tab:: python

        @staticmethod
        def style_prev_paragraph(cursor: XTextCursor | XParagraphCursor, prop_val: object, prop_name: str = None) -> None:
            if prop_name is None:
                prop_name = "ParaStyleName"
            old_val = mProps.Props.get_property(cursor, prop_name)

            cursor.gotoPreviousParagraph(True)  # select previous paragraph
            mProps.Props.set_property(prop_set=cursor, name=prop_name, value=prop_val)

            # reset the cursor and property
            cursor.gotoNextParagraph(False)
            mProps.Props.set_property(prop_set=cursor, name=prop_name, value=old_val)

The current ``ParaStyleName`` value is stored before changing its value in the selected range.
Afterwards, that style name is applied back to the cursor.

:py:meth:`~.Write.style_prev_paragraph` changes the XTextCursor_ into a paragraph cursor so that it's easier to move around across paragraphs.

``read_text()`` calls :py:meth:`~.Write.style_prev_paragraph` with three style names ("Title", "Subtitle", and "Heading").
Those names come from looking at the "Paragraph Styles" dialog window in :numref:`ch06fig_writer_style_ss`.

Work in Progress...

.. |styles_info| replace:: Styles Info
.. _styles_info: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_styles_info

.. |story_creator| replace:: Story Creator
.. _story_creator: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_story_creator

.. _CharacterProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
.. _GenericTextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1GenericTextDocument.html
.. _ParagraphProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties.html
.. _ParagraphStyle: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphStyle.html
.. _XIndexAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexAccess.html
.. _XIndexContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexContainer.html
.. _XNameAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameAccess.html
.. _XNameContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameContainer.html
.. _XStyle: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1style_1_1XStyle.html
.. _XStyleFamiliesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1style_1_1XStyleFamiliesSupplier.html
.. _XTextCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextCursor.html