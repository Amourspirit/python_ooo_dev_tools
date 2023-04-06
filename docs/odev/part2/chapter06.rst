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

..
    Figure 1

.. cssclass:: diagram invert

    .. _ch06fig_txt_doc_serv_interfaces:
    .. figure:: https://user-images.githubusercontent.com/4193389/181575340-96fb7e21-4e0f-4662-8ed9-92edfb036b0c.png
        :alt: Diagram of The Text Document Services, and some Interfaces
        :figclass: align-center

        :The Text Document Services, and some Interfaces.

XStyleFamiliesSupplier_ has a ``getStyleFamilies()`` method for returning text style families.
All these families are stored in an XNameAccess_ object, as depicted in :numref:`ch06fig_txt_doc_serv_interfaces`.

XNameAccess_ is one of Office's collection types, and employed when the objects in a collection have names.
There's also an XIndexAccess_ for collections in index order.

XNameContainer_ and XIndexContainer_ add the ability to insert and remove objects from a collection.

..
    Figure 2

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

..
    Figure 3

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

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The code that implements this process in the Write utility class is a bit more complicated since the calls to
``getByName()`` may raise exceptions if their string arguments are incorrect.

The calls to :py:meth:`.Lo.qi` cast the object returned from a collection into the correct type.

.. _ch06_props_propset:

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

..
    Figure 4

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

..
    Figure 5

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

..
    Figure 6

.. cssclass:: diagram invert

    .. _ch06fig_para_serv_supers:
    .. figure:: https://user-images.githubusercontent.com/4193389/184510955-125605d0-079c-4935-ade4-9d24065ed122.png
        :alt: Diagram of The Paragraph Service and its Superclasses
        :figclass: align-center

        :The Paragraph Service and its Super-classes.

For more information of the styles API, start in the development guide in the "Overall Document Features" section,
online at: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Overall_Document_Features

.. _ch06_list_styles:

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

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

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

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

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

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The last part of |styles_info|_ lists the properties for a specific property set. :py:meth:`.Info.get_style_props` does that:

.. tabs::

    .. code-tab:: python

        @classmethod
        def get_style_props(
            cls, doc: object, family_style_name: str, prop_set_nm: str
        ) -> XPropertySet:
            style_container = cls.get_style_container(doc, family_style_name)
            name_props = Lo.qi(XPropertySet, style_container.getByName(prop_set_nm), True)
            return name_props

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Its arguments are the document, the style family name, and style (property set) name.

A reference to the property set is returned. Accessing the "Standard" style (property set) of the "ParagraphStyle" family would require:

.. tabs::

    .. code-tab:: python

        props = Info.get_style_props(doc, "ParagraphStyles", "Standard")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The property set can be nicely printed by calling :py:meth:`.Props.show_props`:

.. tabs::

    .. code-tab:: python

        Props.show_props('ParagraphStyles "Standard"', props)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

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

.. _ch06_create_style:

6.3 Creating a New Style
========================

The |story_creator|_ example adds a new style to the paragraph style family using :py:meth:`~.write.Write.create_style_para` of ``Write``, and uses it to format the document's paragraphs.



.. tabs::

    .. code-tab:: python

        def create_para_style(doc: XTextDocument, style_name: str) -> bool:
            try:
                font = Font(name=Info.get_font_general_name(), size=12.0)
                spc = Spacing(below=UnitMM(4))
                ln_spc = LineSpacing(mode=ModeKind.FIXED, value=UnitMM(6))

                _ = Write.create_style_para(
                    text_doc=doc, style_name=style_name, styles=[font, spc, ln_spc]
                )
                return True
            except Exception as e:
                print("Could not set paragraph style")
                print(f"  {e}")
            return False

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Three style format are created, :py:class:`~.format.writer.direct.char.font.Font`, :py:class:`~.format.writer.direct.para.indent_space.Spacing`,
and :py:class:`~.format.writer.direct.para.indent_space.LineSpacing`.
These formats are passed to :py:meth:`~.write.Write.create_style_para` which results in the style formats being applied to the style as it is being created.

:py:class:`~.units.UnitMM` are imported from :ref:`ns_units` and is a convenient way to pass values into methods that accept :ref:`proto_unit_obj` arguments.

In |story_creator|_, ``create_para_style()`` is called like so:

.. tabs::

    .. code-tab:: python

        doc = Write.create_doc(loader=loader)
        # ...
        if not create_para_style(doc, "adParagraph"):
            raise RuntimeError("Could not create new paragraph style")
        
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

A new style called ``adParagraph`` is added to the paragraph style family.
It uses os dependent font determined by :py:meth:`.Info.get_font_general_name` such as "Liberation Serif" 12pt font, and leaves a 4mm space between paragraphs.

.. _ch06_apply_style_para:

6.4 Applying Styles to Paragraphs (and Characters)
==================================================

An ``adParagraph`` style is added to the paragraph style family, but how to apply that style to some paragraphs in the document?
The easiest way is through the document's XTextRange_ interface.
XTextRange_ is supported by the TextRange_ service, which inherits ParagraphProperties_ and CharacterProperties_ (and several other property classes), as illustrated in :numref:`ch06fig_txt_rng_srvc`.

..
    Figure 8

.. cssclass:: diagram invert

    .. _ch06fig_txt_rng_srvc:
    .. figure:: https://user-images.githubusercontent.com/4193389/184718158-9d8a414c-5682-4df4-9a0f-962f3b360351.png
        :alt: Diagrom of The TextRange Service.
        :figclass: align-center

        :The TextRange Service.

XTextRange_ can be cast to XPropertySet_ to make the properties in ParagraphProperties_ and CharacterProperties_ accessible.
An existing (or new) paragraph style can be applied to a text range by setting its ``ParaStyleName`` property:

However |odev| can simplify this process by using :py:class:`~.format.writer.style.Para` class.



.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.style.para import Para as StylePara
        # ...

        xtext_range = doc.getText().getStart()
        para_style = StylePara("adParagraph")
        para_style.apply(xtext_range)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The code above obtains the text range at the start of the document, and set its paragraph style to ``adParagraph``.
Any text added from this position onward will use that style.

This approach is used in |story_creator|_: the style is set first, then text is added.

.. _ch06_cursors_txt_rng:

6.5 Cursors and Text Ranges
===========================

Another technique for applying styles uses a cursor to select a text range.
Then the text's properties are accessed through the cursor.

All the different kinds of model and view cursor belong to the TextCursor_ service, and this inherits TextRange_.
This allows us to extend :numref:`ch06fig_txt_rng_srvc` to become :numref:`ch06fig_txt_rng_srvc_cursor`.

..
    Figure 9

.. cssclass:: diagram invert

    .. _ch06fig_txt_rng_srvc_cursor:
    .. figure:: https://user-images.githubusercontent.com/4193389/184720203-8147f173-596c-4aae-b7ce-c1e8a3b0e674.png
        :alt: Diagrom of Cursor Access to Text Properties
        :figclass: align-center

        :Cursor Access to Text Properties.

This hierarchy means that a cursor can access the TextRange_ service and its text properties.
The following code fragment demonstrates the idea:

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        cursor.gotoEnd(True) # select the entire document

        props = Lo.qi(XPropertySet, cursor)
        props.setProperty("ParaStyleName", "adParagraph")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Using :py:meth:`.Props.set_property`, simplifies this to:

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        cursor.gotoEnd(True)
        Props.set_property(cursor, "ParaStyleName", "adParagraph")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch06_build_story:

6.6 Building a Story Document
=============================

|story_creator|_  example starts by setting the ``adParagraph`` style, then employs ``read_text()`` to read text from a file and add it to the document:

.. tabs::

    .. code-tab:: python

        xtext_range = doc.getText().getStart()
        para_style = StylePara("adParagraph")
        para_style.apply(xtext_range)

        # ...

        read_text(fnm=fnm, cursor=cursor)
        Write.end_paragraph(cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``read_text()`` assumes the text file has a certain format. For example, ``scandal.txt`` begins like so:

.. .. code-block:: text

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

        sb: List[str] = []
        with open(fnm, "r") as file:
            i = 0
            for ln in file:
                line = ln.rstrip()  # remove new line \n
                if len(line) == 0:
                    if len(sb) > 0:
                        Write.append_para(cursor, " ".join(sb))
                    sb.clear()
                elif line.startswith("Title: "):
                    Write.append_para(cursor, line[7:], styles=[StylePara(StyleParaKind.TITLE)])
                elif line.startswith("Author: "):
                    Write.append_para(cursor, line[8:], styles=[StylePara(StyleParaKind.SUBTITLE)])
                elif line.startswith("Part "):
                    Write.append_para(cursor, line, styles=[StylePara(StyleParaKind.HEADING_1)])
                else:
                    sb.append(line)
                i += 1

            if len(sb) > 0:
                Write.append_para(cursor, " ".join(sb))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The interesting bits are the calls to :py:meth:`.Write.append_para` which add a paragraph to the document and apply a style to it.
For instance:

.. tabs::

    .. code-tab:: python

        elif line.startswith("Author: "):
            Write.append_para(cursor, line[8:], styles=[StylePara(StyleParaKind.SUBTITLE)])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:class:`~.format.writer.style.Para` (imported as ``StylePara``) is used to access ``Subtitle`` style.
:py:class:`~.format.writer.style.StyleParaKind` is an enum contain built in style values.

:py:meth:`.Write.append_para` writes the string into the document as a paragraph and applies the style (the input line without the "Author: " substring).


.. _ch06_style_change:

6.7 Style Changes to Words and Phrases
======================================

Aside from changing paragraph styles, it's useful to apply style changes to words or strings inside a paragraph.
For example, to highlight a word in bold, or write several words in red italics for emphasis.

This is implemented by :py:meth:`.Write.style_left` using a similar approach to :py:meth:`.Write.style_prev_pragraph`.
:py:meth:`~.Write.style_left` is passed an integer position which lies to the left of the current cursor position.
Character style changes are applied to the text range defined by that distance:

.. tabs::

    .. code-tab:: python

        def style_left(cursor: XTextCursor, pos: int, prop_name: str, prop_val: object) -> None:
            old_val = Props.get_property(cursor, prop_name)

            curr_pos = Selection.get_position(cursor)
            cursor.goLeft(curr_pos - pos, True)
            Props.set_property(prop_set=cursor, name=prop_name, value=prop_val)

            cursor.goRight(curr_pos - pos, False)
            Props.set_property(prop_set=cursor, name=prop_name, value=old_val)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


A XTextCursor_ is used to select the range, and the new style is set.
Then the cursor is moved back to its old position, and the previous style reapplied.

There is also an overload (shown below) of :py:meth:`.Write.style_left` that works in s similar manor but takes one or more :py:class:`~.proto.style_obj.StyleObj` styles.

.. tabs::

    .. code-tab:: python

        def style_left(cls, cursor: XTextCursor, pos: int, styles: Iterable[StyleObj]) -> None:
            ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The Write class contain a few support functions that set common styles using :py:meth:`~.Write.style_left`:

.. tabs::

    .. code-tab:: python

        @classmethod
        def style_left_bold(cls, cursor: XTextCursor, pos: int) -> None:
            cls.style_left(cursor, pos, "CharWeight", FontWeight.BOLD)

        @classmethod
        def style_left_italic(cls, cursor: XTextCursor, pos: int) -> None:
            cls.style_left(cursor, pos, "CharPosture", FontSlant.ITALIC)

        @classmethod
        def style_left_color(cls, cursor: XTextCursor, pos: int, color: Color) -> None:
            cls.style_left(cursor, pos, "CharColor", color)

        @classmethod
        def style_left_code(cls, cursor: XTextCursor, pos: int) -> None:
            cls.style_left(cursor, pos, "CharFontName", Info.get_font_mono_name())
            cls.style_left(cursor, pos, "CharHeight", 10)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The position (the pos value) passed to :py:meth:`~.Write.style_left` can be obtained from :py:meth:`.Write.get_position`.

The |build_doc|_ example takes advantage of a few python partial methods to cut down on typing.

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)

        # take advantage of a few partial functions
        nl = partial(Write.append_line, cursor)
        np = partial(Write.end_paragraph, cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The |build_doc|_ example contains several examples of how to use formatting styles from :ref:`ns_format` module.

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        Write.append(cursor, "Some examples of simple text ")
        Write.append_line(cursor=cursor, text="styles.", styles=[Font(b=True)])
        Write.append_para(
            cursor=cursor,
            text="This line is written in red italics.",
            styles=[Font(color=CommonColor.DARK_RED).bold.italic],
        )
        Write.append_para(cursor=cursor, text="Back to old style")
        nl()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The resulting text in the document looks like :numref:`ch06fig_styled_text_ss`.

..
    Figure 10

.. cssclass:: screen_shot

    .. _ch06fig_styled_text_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184726710-b7b94880-723f-4e93-b15d-74477bd7c752.png
        :alt: Screen Shot of Styled Text
        :figclass: align-center

        :Styled Text.

The following fragment from |build_doc|_ applies a 'code' styling to several lines:

.. tabs::

    .. code-tab:: python

        Write.append_para(cursor, "Here's some code:")

        code_font = Font(name=Info.get_font_mono_name(), size=10)
        code_font.apply(cursor)

        nl("public class Hello")
        nl("{")
        nl("  public static void main(String args[]")
        nl('  {  System.out.println("Hello World");  }')
        Write.append_para(cursor, "}  // end of Hello class")

        # reset the cursor formatting
        ParaStyle.default.apply(cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:numref:`ch06fig_styled_text_code_ss` shows the generated document text.

..
    Figure 11

.. cssclass:: screen_shot invert

    .. _ch06fig_styled_text_code_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184730866-6a39e2fd-76a3-4afe-8c32-ccaa8e13633b.png
        :alt: Screen Shot of Text with Code Styling
        :figclass: align-center

        :Text with Code Styling.

:py:class:`ooodev.format.writer.direct.char.font.Font` is imported and used to create mono font values and applied directly to the cursor.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.char.font import Font
        # ...

        code_font = Font(name=Info.get_font_mono_name(), size=10)
        code_font.apply(cursor)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Once the code as been written, :py:attr:`ParaStyle.default <ooodev.format.writer.style.para.para.Para.default>` is applied to the cursor to reset its values.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.style.para import Para as ParaStyle
        # ...

        # reset the cursor formatting
        ParaStyle.default.apply(cursor)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch06_style_hyperlink:

6.8 Hyperlink Styling
=====================

Text hyperlinks are implemented as a style, using :py:class:`ooodev.format.writer.direct.char.hyperlink.Hyperlink`.
|build_doc|_ shows how the hyperlink is constructed and applied.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.char.hyperlink import Hyperlink, TargetKind
        # ...

        # Insert a hyperlink.
        Write.append(cursor, "A link to ")

        hl = Hyperlink(
            name="ODEV_GITHUB",
            url="https://github.com/Amourspirit/python_ooo_dev_tools",
            target=TargetKind.BLANK
        )
        Write.append(cursor, "OOO Development Tools", styles=[hl])

        Write.append_para(cursor, " Website.")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

..
    Figure 12

.. cssclass:: screen_shot invert

    .. _ch06fig_text_hyperlink_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184732547-33adc6b0-7d4a-4d41-9558-1b9f6ae188ea.png
        :alt: Screen Shot of Text Containing a Hypertext Link.
        :figclass: align-center

        :Text Containing a Hypertext Link.

If the user control-clicks on the link, then the URL value is loaded into the browser.

The ``HyperLink`` name property ( ``name="ODEV_GITHUB"`` ) specifies a link name, which can be used when searching a document.

.. _ch06_text_num:

6.9 Text Numbering
==================

It's straightforward to number paragraphs by using  :py:class:`~ooodev.format.writer.direct.para.outline_list.ListStyle`, :py:class:`~ooodev.format.writer.style.lst.style_list_kind.StyleListKind`
and :py:class:`~ooodev.format.writer.style.bullet_list.bullet_list.BulletList`.

The following code from |build_doc|_ , numbers three paragraphs:



.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.para.outline_list import ListStyle, StyleListKind
        # ...

        Write.append_para(cursor, "The following points are important:")

        list_style = ListStyle(list_style=StyleListKind.NUM_123, num_start=-2)
        list_style.apply(cursor)

        Write.append_para(cursor, "Have a good breakfast")
        Write.append_para(cursor, "Have a good lunch")
        Write.append_para(cursor, "Have a good dinner")

        # Reset to default which set cursor to No List Style
        list_style.default.apply(cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The result is shown in :numref:`ch06fig_text_num_para_ss`.

..
    Figure 13

.. cssclass:: screen_shot invert

    .. _ch06fig_text_num_para_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184733566-ce060993-022e-4071-9f6e-b1db5dc3e8b9.png
        :alt: Screen Shot of Numbered Paragraphs.
        :figclass: align-center

        :Numbered Paragraphs.

Letters are drawn instead of numbers by changing the style name to "Numbering abc" (see :numref:`ch06fig_text_letter_para_ss`).

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.para.outline_list import ListStyle, StyleListKind
        from ooodev.format.writer.style.bullet_list import BulletList
        # ...

        Write.append_para(cursor=cursor, text="Breakfast should include:")

        # set cursor style to Number abc
        list_style = ListStyle(list_style=StyleListKind.NUM_abc, num_start=-2)
        list_style.apply(cursor)

        Write.append_para(cursor, "Porridge")
        Write.append_para(cursor, "Orange Juice")
        Write.append_para(cursor, "A Cup of Tea")
        # reset cursor number style
        list_style.default.apply(cursor)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

..
    Figure 14

.. cssclass:: screen_shot invert

    .. _ch06fig_text_letter_para_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/184734264-458598cc-ca43-4e7b-b080-2a5c74b945e5.png
        :alt: Screen Shot of Lettered Paragraphs.
        :figclass: align-center

        :Letter Paragraphs.

One issue with numbered paragraphs is that their default behavior retains the current count when numbering another group of text.
For example, a second group of numbered paragraphs appearing in the document after :numref:`ch06fig_text_num_para_ss` would start at ``4``.
This is fixed by setting ``num_start=-2`` when creating an instance of :py:class:`~ooodev.format.writer.direct.para.outline_list.ListStyle`,
this forces a numbering reset.

.. tabs::

    .. code-tab:: python

        list_style = ListStyle(list_style=StyleListKind.NUM_123, num_start=-2)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

One large topic that is not covered in this document is numbering.
This includes the numbering of chapter headings and lines.
Chapter and line numbering are dealt with differently from most document styles.
Instead of being accessed via XStyleFamiliesSupplier_, they employ XChapterNumberingSupplier_ and XNumberFormatsSupplier_.

For more details, see the development guide: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Line_Numbering_and_Outline_Numbering

.. _ch06_style_oth:

6.10 Other Style Changes
========================

|story_creator|_ example illustrates three other styling effects: the creation of a header, setting the page to A4 format, and employing page numbers in the footer.
The relevant calls are:

.. tabs::

    .. code-tab:: python

        # fragment from story creator
        from ooodev.format.writer.direct.char.font import Font
        from ooodev.format.writer.direct.page.header import Header
        from ooodev.format.writer.direct.page.header.area import Img, PresetImageKind
        # ...

        # header formatting
        # create a header font style with a size of 9 pt, italic and dark green.
        header_font = Font(
            name=Info.get_font_general_name(), size=9.0, i=True, color=CommonColor.DARK_GREEN
        )
        header_format = Header(
            on=True,
            shared_first=True,
            shared=True,
            height=13.0,
            spacing=3.0,
            spacing_dyn=True,
            margin_left=1.5,
            margin_right=2.0,
        )
        # create a header image from a preset
        header_img = Img.from_preset(PresetImageKind.MARBLE)
        # Set header can be passed a list of styles to format the header.
        Write.set_header(
            text_doc=doc, text=f"From: {fnm.name}", styles=[header_font, header_format, header_img]
        )

        # page format A4
        Write.set_a4_page_format(doc)
        Write.set_page_numbers(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Write.set_a4_page_format` sets the page formatting.

Alternatively, the page format can be set via the :py:class:`~ooodev.format.writer.modify.page.page.PaperFormat` class.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.page import PaperFormat, PaperFormatKind
        # ...

        page_size_style = PaperFormat.from_preset(preset=PaperFormatKind.A4)
        page_size_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Write.set_page_numbers` utilizes text fields, which is examined in the "Text Fields" section in :ref:`ch07`.

Changing the header in :py:meth:`.Write.set_header` requires the setting of the ``HeaderIsOn`` boolean in the ``Standard`` page style.
Adding text to the header is done via an XText_ reference.
A simpler version of |odev|'s :py:meth:`.Write.set_header` is shown below.


.. tabs::

    .. code-tab:: python

        @staticmethod
        def set_header(text_doc: XTextDocument, text: str) -> None:
            props = Info.get_style_props(
                doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard"
            )
            if props is None:
                raise PropertiesError("Could not access the standard page style container")
            try:
                props.setPropertyValue("HeaderIsOn", True)
                # header must be turned on in the document
                # props.setPropertyValue("TopMargin", 2200)
                header_text = Lo.qi(XText, props.getPropertyValue("HeaderText"))
                header_cursor = header_text.createTextCursor()
                header_cursor.gotoEnd(False)

                header_props = Lo.qi(XPropertySet, header_cursor, True)
                header_props.setPropertyValue("CharFontName", Info.get_font_general_name())
                header_props.setPropertyValue("CharHeight", 10)
                header_props.setPropertyValue("ParaAdjust", ParagraphAdjust.RIGHT)

                header_text.setString(f"{text}\n")
            except Exception as e:
                raise Exception("Unable to set header text") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

There also is a :py:meth:`.Write.set_footer` method in |odev| that allows for the passing of a list of styles to format the footer.

The header's XText_ reference is retrieved via the page style's ``HeaderText`` property, and a cursor is created local to the header:

.. tabs::

    .. code-tab:: python

        header_cursor = header_text.createTextCursor()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This cursor can only move around inside the header not the entire document.

The properties of the header's XText_ are changed using the cursor, and then the text is adde

.. seealso:: 

    .. cssclass:: ul-list

        - :ref:`help_writer_format_style_para`


.. |styles_info| replace:: Styles Info
.. _styles_info: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_styles_info

.. |story_creator| replace:: Story Creator
.. _story_creator: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_story_creator

.. |build_doc| replace:: Build Doc
.. _build_doc: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_doc

.. _CharacterProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
.. _GenericTextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1GenericTextDocument.html
.. _LineSpacing: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1style_1_1LineSpacing.html
.. _ParagraphProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties.html
.. _ParagraphProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties.html
.. _ParagraphStyle: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphStyle.html
.. _TextCursor: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextCursor.html
.. _TextRange: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextRange.html
.. _XChapterNumberingSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XChapterNumberingSupplier.html
.. _XIndexAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexAccess.html
.. _XIndexContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexContainer.html
.. _XNameAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameAccess.html
.. _XNameContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameContainer.html
.. _XNumberFormatsSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XNumberFormatsSupplier.html
.. _XPropertySet: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertySet.html
.. _XStyle: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1style_1_1XStyle.html
.. _XStyleFamiliesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1style_1_1XStyleFamiliesSupplier.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html
.. _XTextCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextCursor.html
.. _XTextRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextRange.html