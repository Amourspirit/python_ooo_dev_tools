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
Many of the suppliers shown in :numref:`ch05fig_txt_doc_serv_interfaces` (:abbreviation:`ex:` ``XTextTablesSupplier``)
are for iterating through text content (:abbreviation:`ex:` accessing the document's tables).

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

The |extract_ex|_ example opens a document using :py:meth:`.Lo.open_doc`, and tries to print its text:

.. tabs::

    .. code-tab:: python

        #!/usr/bin/env python
        # coding: utf-8
        from __future__ import annotations
        import argparse
        from typing import Any, cast

        from ooodev.utils.lo import Lo
        from ooodev.office.write import Write
        from ooodev.utils.info import Info
        from ooodev.wrapper.break_context import BreakContext
        from ooodev.events.gbl_named_event import GblNamedEvent
        from ooodev.events.args.cancel_event_args import CancelEventArgs
        from ooodev.events.lo_events import LoEvents


        def args_add(parser: argparse.ArgumentParser) -> None:
            parser.add_argument(
                "-f",
                "--file",
                help="File path of input file to convert",
                action="store",
                dest="file_path",
                required=True,
            )

        def on_lo_print(source: Any, e: CancelEventArgs) -> None:
            e.cancel = True

        def main() -> int:
            parser = argparse.ArgumentParser(description="main")
            args_add(parser=parser)
            args = parser.parse_args()

            # hook ooodev internal printing event
            LoEvents().on(GblNamedEvent.PRINTING, on_lo_print)

            with BreakContext(Lo.Loader(Lo.ConnectSocket(headless=True))) as loader:
                fnm = cast(str, args.file_path)

                try:
                    doc = Lo.open_doc(fnm=fnm, loader=loader)
                except Exception:
                    print(f"Could not open '{fnm}'")
                    raise BreakContext.Break

                if Info.is_doc_type(obj=doc, doc_type=Lo.Service.WRITER):
                    text_doc = Write.get_text_doc(doc=doc)
                    cursor = Write.get_cursor(text_doc)
                    text = Write.get_all_text(cursor)
                    print("Text Content".center(50, "-"))
                    print(text)
                    print("-" * 50)
                else:
                    print("Extraction unsupported for this doc type")
                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

|extract_ex|_ example also hooks |odev|'s internal events and cancels the printing event.
Thus suppressing any internal printing to console.

.. tabs::

    .. code-tab:: python

        def on_lo_print(source: Any, e: CancelEventArgs) -> None:
            e.cancel = True

        def main() -> int:

            # hook internal printing event
            LoEvents().on(GblNamedEvent.PRINTING, on_lo_print)

If internal printing were not suppressed the output would contains extra
output similar to what is shown here:

.. code-block:: text

    Loading Office...
    Opening /home/user/Python/ooouno_ex/resources/odt/cicero_dummy.odt
    -------------------Text Content-------------------
    Cicero
    Dummy Text
    But I must explain to you how all this mistaken idea of denouncing pleasure and praising ...
    --------------------------------------------------
    Closing the document
    Closing Office
    Office terminated
    Office bridge has gone!!

:py:meth:`.Info.is_doc_type` tests the document's type by casting it into an XServiceInfo_ interface. Then it calls ``XServiceInfo.supportsService()``
to check the document's service capabilities:

.. tabs::

    .. code-tab:: python

        @staticmethod
        def is_doc_type(obj: object, doc_type: Lo.Service) -> bool:
            try:
                si = Lo.qi(XServiceInfo, obj)
                if si is None:
                    return False
                return si.supportsService(str(doc_type))
            except Exception:
                return False

The argument type of the document is Object rather than XComponent_ so that a wider range of objects can be passed to the function for testing.

The service names for documents are hard to remember, so they're defined as an enumeration in the :py:class:`.Lo.Service`.

:py:meth:`.Write.get_text_doc` uses :py:meth:`.Lo.qi` to cast the document's XComponent_ interface into an XTextDocument_:

.. tabs::

    .. code-tab:: python

        text_doc = Lo.qi(XTextDocument, doc, True)


``text_doc = Lo.qi(XTextDocument, doc)`` This may fail (i.e. return None) if the loaded document isn't an instance of the TextDocument_ service.

The casting 'power' of :py:meth:`.Lo.qi` is confusing – it depends on the document's service type.
All text documents are instances of the TextDocument_ service (see :numref:`ch05fig_txt_doc_serv_interfaces`).
This means that :py:meth:`.Lo.qi` can 'switch' between any of the interfaces defined by TextDocument_
or its super-classes (i.e. the interfaces in GenericTextDocument_ or OfficeDocument_).
For instance, the following cast is fine:

.. tabs::

    .. code-tab:: python

        xsupplier = Lo.qi(XStyleFamiliesSupplier, doc)

This changes the instance into an XStyleFamiliesSupplier_, which can access the document's styles.

Alternatively, the following converts the instance into a supplier defined in OfficeDocument_:

.. tabs::

    .. code-tab:: python

        xsupplier = Lo.qi(XDocumentPropertiesSupplier, doc)

Most of the examples in this chapter and the next few cast the document to XTextDocument_ since that interface can access the document's contents as an XText_ object:


.. tabs::

    .. code-tab:: python

        text_doc = Lo.qi(XTextDocument, doc)
        xtext = text_doc.getText()

The XText_ instance can access all the capabilities shown in :numref:`ch05fig_xtext_supers`.

A common next step is to create a cursor for moving around the document.
This is easy since XText_ inherits XSimpleText_ which has a ``createTextCursor()`` method:

.. tabs::

    .. code-tab:: python

        text_cursor = xText.createTextCursor()

These few lines are so useful that they are part of :py:meth:`.Selection.get_cursor` method which :py:class:`~.write.Write` inherits.

An XTextCursor can be converted into other kinds of model cursors (:abbreviation:`eg:`
XParagraphCursor_, XSentenceCursor_, XWordCursor_; see :numref:`ch05fig_cursor_types`).
That's not necessary in for the |extract_ex|_ example; instead, the XTextCursor_ is passed to
:py:meth:`.Write.get_all_text` to access the text as a sequence of characters:

.. tabs::

    .. code-tab:: python

        @staticmethod
        def get_all_text(cursor: XTextCursor) -> str:
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            text = cursor.getString()
            cursor.gotoEnd(False)  # to deselect everything
            return text

All cursor movement operations take a boolean argument which specifies whether the movement should also select the text.
For example, in :py:meth:`~.Write.get_all_text`, ``cursor.gotoStart(False)`` shifts the cursor to the start of the text without selecting anything.
The subsequent call to ``cursor.gotoEnd(True)`` moves the cursor to the end of the text and selects all the text moved over.
The call to ``getString()`` on the third line returns the selection (:abbreviation:`eg:` all the text in the document).

Two other useful XTextCursor_ methods are:

.. tabs::

    .. code-tab:: python

        cursro.goLeft(char_count, is_selected)
        cursor.goRight(char_count, is_selected)

They move the cursor left or right by a given number of characters, and the boolean argument specifies whether the text moved over is selected.

All cursor methods return a boolean result which indicates if the move (and optional selection) was successful.

Another method worth knowing is:

.. tabs::

    .. code-tab:: python

        cursro.gotoRange(text_range, is_selected)

``gotoRange()`` method of XTextCursor_ takes an XTextRange_ argument, which represents a selected region or position where the cursor should be moved to.
For example, it's possible to find a bookmark in a document, extract its text range/position, and move the cursor to that location with ``gotoRange()``.

.. todo::

    Link ch5 to chapter 7

code for this in Chapter 7.

A Problem with Write.get_all_text()
-----------------------------------

:py:meth:`~.Write.get_all_text` may fail if supplied with a very large document because ``XTextCursor.getString()`` might be unable to construct a big enough String object.
For that reason, it's better to iterate over large documents returning a paragraph of text at a time.
These iteration techniques are described next.

Work in progress ...

.. |txt_java| replace:: TextDocuments.java
.. _txt_java: https://api.libreoffice.org/examples/DevelopersGuide/examples.html#Text

.. |write_guide| replace:: Writer Guide
.. _write_guide: https://documentation.libreoffice.org/en/english-documentation/

.. |extract_ex| replace:: Extract Writer Text
.. _extract_ex: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_doc_print_console

.. _GenericTextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1GenericTextDocument.html
.. _OfficeDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1OfficeDocument.html
.. _TextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextDocument.html
.. _XComponent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XComponent.html
.. _XParagraphCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XParagraphCursor.html
.. _XSentenceCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XSentenceCursor.html
.. _XServiceInfo: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XServiceInfo.html
.. _XSimpleText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XSimpleText.html
.. _XStyleFamiliesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1style_1_1XStyleFamiliesSupplier.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html
.. _XTextCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextCursor.html
.. _XTextDocument: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextDocument.html
.. _XTextRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextRange.html
.. _XWordCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XWordCursor.html
