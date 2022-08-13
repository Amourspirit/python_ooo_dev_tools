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



Work in Progress...

.. _GenericTextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1GenericTextDocument.html
.. _XIndexAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexAccess.html
.. _XIndexContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexContainer.html
.. _XNameAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameAccess.html
.. _XNameContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameContainer.html
.. _XStyleFamiliesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1style_1_1XStyleFamiliesSupplier.html