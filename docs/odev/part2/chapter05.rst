.. _ch05:

****************************
Chapter 5. Text API Overview
****************************

.. topic:: Overview

    API Overview; Text Cursors; Extracting Text; Cursor Iteration;
    Creating Cursors; Creating a Document; Using and Comparing Text Cursors;
    Inserting/Changing Text in a Document; Text Enumeration; Appending Documents

The next few chapters look at programming with the text document part of the Office API.
This chapter begins with a quick overview of the text API, then a detailed look at text cursors for moving about in a document,
extracting text, and adding/inserting new text.

Text cursors aren't the only way to move around inside a document;
it's also possible to iterate over a document by treating it as a sequence of paragraphs.

The chapter finishes with a look at how two (or more) text documents can be appended.

The online Developer's Guide begins text document programming at
https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Text_Documents (the easiest way of accessing that page is to type ``loguide writer``).
It corresponds to Chapter 7 in the printed guide (available at https://wiki.openoffice.org/w/images/d/d9/DevelopersGuide_OOo3.1.0.pdf),
but the Web material is better structured and formatted.

The guide's text programming examples are in |txt_java|_.

Although the code is long, it's well-organized. Some smaller text processing examples are available at https://api.libreoffice.org/examples/examples.html#Java_examples.

This chapter (and later ones) assume that you're familiar with Writer, including text concepts such as paragraph styles. If you're not, then I recommend the |write_guide|_, user manual.

5.1 An Overview of the Text Document API
========================================

The API is centered around four text document services which sub-class ``OfficeDocument``, as shown in :numref:`ch05fig_txt_doc_service`.

.. cssclass:: diagram invert

    .. _ch05fig_txt_doc_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/181572774-7b5899f5-182c-4fc6-8221-25a2d2ae2b58.png
        :alt: Diagram of The Text Document Services

        :The Text Document Services.


This chapter concentrates on the TextDocument_ service.
Or you can type ``lodoc TextDocument service``.

The ``GlobalDocument`` service in :numref:`ch05fig_txt_doc_service` is employed by master documents, such as a book or thesis.
A master document is typically made up of links to files holding its parts, such as chapters, bibliography, and appendices.

The ``WebDocument`` service in :numref:`ch05fig_txt_doc_service` is for manipulating web pages, although its also possible to generate HTML files with the TextDocument service.

``TextDocument``, ``GlobalDocument``, and ``WebDocument`` are mostly empty because those services don't define any interfaces or properties.
The ``GenericTextDocument`` service is where the action takes place, as summarized in :numref:`ch05fig_txt_doc_serv_interfaces`.

.. cssclass:: diagram invert

    .. _ch05fig_txt_doc_serv_interfaces:
    .. figure:: https://user-images.githubusercontent.com/4193389/181575340-96fb7e21-4e0f-4662-8ed9-92edfb036b0c.png
        :alt: Diagram of The Text Document Services, and some Interfaces

        :The Text Document Services, and some Interfaces.

The numerous 'Supplier' interfaces in :numref:`ch05fig_txt_doc_serv_interfaces` are Office's way of accessing different elements in a document.
For example, ``XStyleFamiliesSupplier`` manages character, paragraph, and other styles, while ``XTextTableSupplier`` deals with tables.


In later chapters we will be looking at these suppliers, which is why they're highlighted,
but for now let's only consider the ``XTextDocument`` interface at the top right of the ``GenericTextDocument`` service box
in :numref:`ch05fig_txt_doc_serv_interfaces` ``XTextDocument`` has a ``getText()`` method for returning an ``XText`` object.
``XText`` supports functionality related to text ranges and positions, cursors, and text contents.

It inherits ``XSimpleText`` and ``XTextRange``, as indicated in :numref:`ch05fig_xtext_supers`.

.. cssclass:: diagram invert

    .. _ch05fig_xtext_supers:
    .. figure:: https://user-images.githubusercontent.com/4193389/181577210-0054e815-2a45-4a86-a782-bd703b1e442a.png
        :alt: Diagram of XText and its Super-classes

        : ``XText`` and its Super-classes.

Text content covers a multitude, such as embedded images, tables, footnotes, and text fields.
Many of the suppliers shown in :numref:`ch05fig_txt_doc_serv_interfaces` (:abbreviation:`e.g.` ``XTextTablesSupplier``)
are for iterating through text content (:abbreviation:`e.g.` accessing the document's tables).

.. todo::

    | Chapte 5, Add link to chapters 7
    | Chapte 5, Add link to chapters 8

This chapter concentrates on ordinary text, chapters 7 and 8 look at more esoteric content forms.

A text document can utilize eight different cursors, which fall into two groups, as in :numref:`ch05fig_cursor_types`.

.. cssclass:: diagram invert

    .. _ch05fig_cursor_types:
    .. figure:: https://user-images.githubusercontent.com/4193389/181580982-4a4c7210-efc2-43a6-b21c-5b9e626d2ff8.png
        :alt: Diagram of Types of Cursor

        :Types of Cursor.

``XTextCursor`` contains methods for moving around the document, and an instance is often called a model cursor
because of its close links to the document's data. A program can create multiple ``XTextCursor`` objects if it wants,
and can convert an ``XTextCursor`` into ``XParagraphCursor``, ``XSentenceCursor``, or ``XWordCursor``.
The differences are that while an ``XTextCursor`` moves through a document character by character, the others travel in units of paragraphs, sentences, and words.

A program may employ a single ``XTextViewCursor`` cursor, to represent the cursor the user sees in the Writer application window;
for this reason, it's often called the view cursor. ``XTextViewCursor`` can be converted into a ``XLineCursor``, ``XPageCursor``, or ``XScreenCursor`` object,
which allows it to move in terms of lines, pages, or screens.

A cursor's location is specified using a text range, which can be the currently selected text, or a position in the document.
A text position is a text range that begins and ends at the same point.

5.2 Extracting Text from a Document
===================================



Work in progress ...

.. |txt_java| replace:: TextDocuments.java
.. _txt_java: https://api.libreoffice.org/examples/DevelopersGuide/examples.html#Text
.. |write_guide| replace:: Writer Guide
.. _write_guide: https://documentation.libreoffice.org/en/english-documentation/

.. _TextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextDocument.html