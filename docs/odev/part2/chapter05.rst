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

..
    Figure 1

.. cssclass:: diagram invert

    .. _ch05fig_txt_doc_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/181572774-7b5899f5-182c-4fc6-8221-25a2d2ae2b58.png
        :alt: Diagram of The Text Document Services
        :figclass: align-center

        :The Text Document Services.


This chapter concentrates on the TextDocument_ service.
Or you can type ``lodoc TextDocument service``.

The ``GlobalDocument`` service in :numref:`ch05fig_txt_doc_service` is employed by master documents, such as a book or thesis.
A master document is typically made up of links to files holding its parts, such as chapters, bibliography, and appendices.

The ``WebDocument`` service in :numref:`ch05fig_txt_doc_service` is for manipulating web pages, although its also possible to generate HTML files with the TextDocument service.

``TextDocument``, ``GlobalDocument``, and ``WebDocument`` are mostly empty because those services don't define any interfaces or properties.
The ``GenericTextDocument`` service is where the action takes place, as summarized in :numref:`ch05fig_txt_doc_serv_interfaces`.

..
    Figure 2

.. cssclass:: diagram invert

    .. _ch05fig_txt_doc_serv_interfaces:
    .. figure:: https://user-images.githubusercontent.com/4193389/181575340-96fb7e21-4e0f-4662-8ed9-92edfb036b0c.png
        :alt: Diagram of The Text Document Services, and some Interfaces
        :figclass: align-center

        :The Text Document Services, and some Interfaces.

The numerous 'Supplier' interfaces in :numref:`ch05fig_txt_doc_serv_interfaces` are Office's way of accessing different elements in a document.
For example, ``XStyleFamiliesSupplier`` manages character, paragraph, and other styles, while ``XTextTableSupplier`` deals with tables.


In later chapters we will be looking at these suppliers, which is why they're highlighted,
but for now let's only consider the ``XTextDocument`` interface at the top right of the ``GenericTextDocument`` service box
in :numref:`ch05fig_txt_doc_serv_interfaces` ``XTextDocument`` has a ``getText()`` method for returning an ``XText`` object.
``XText`` supports functionality related to text ranges and positions, cursors, and text contents.

It inherits ``XSimpleText`` and ``XTextRange``, as indicated in :numref:`ch05fig_xtext_supers`.

..
    Figure 3

.. cssclass:: diagram invert

    .. _ch05fig_xtext_supers:
    .. figure:: https://user-images.githubusercontent.com/4193389/181577210-0054e815-2a45-4a86-a782-bd703b1e442a.png
        :alt: Diagram of XText and its Super-classes
        :figclass: align-center

        : ``XText`` and its Super-classes.

Text content covers a multitude, such as embedded images, tables, footnotes, and text fields.
Many of the suppliers shown in :numref:`ch05fig_txt_doc_serv_interfaces` (:abbreviation:`ex:` ``XTextTablesSupplier``)
are for iterating through text content (:abbreviation:`ex:` accessing the document's tables).

.. todo::

    | Chapte 5, Add link to chapters 7
    | Chapte 5, Add link to chapters 8

This chapter concentrates on ordinary text, chapters 7 and 8 look at more esoteric content forms.

A text document can utilize eight different cursors, which fall into two groups, as in :numref:`ch05fig_cursor_types`.

..
    Figure 4

.. cssclass:: diagram invert

    .. _ch05fig_cursor_types:
    .. figure:: https://user-images.githubusercontent.com/4193389/181580982-4a4c7210-efc2-43a6-b21c-5b9e626d2ff8.png
        :alt: Diagram of Types of Cursor
        :figclass: align-center

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

Code for this in :ref:`ch07`.

A Problem with Write.get_all_text()
-----------------------------------

:py:meth:`~.Write.get_all_text` may fail if supplied with a very large document because ``XTextCursor.getString()`` might be unable to construct a big enough String object.
For that reason, it's better to iterate over large documents returning a paragraph of text at a time.
These iteration techniques are described next.

5.3 Cursor Iteration
====================

In the |walk_text|_ example it uses paragraph and word cursors
(:abbreviation:`eg:` the XParagraphCursor_ and XWordCursor_ interfaces in :numref:`ch05fig_cursor_types`).
It also employs the view cursor, an XTextViewCursor_ instance, to control the Writer application's visible cursor.

.. tabs::

    .. code-tab:: python

        def main() -> int:
            parser = argparse.ArgumentParser(description="main")
            args_add(parser=parser)
            args = parser.parse_args()

            with BreakContext(Lo.Loader(Lo.ConnectSocket())) as loader:

                fnm = cast(str, args.file_path)

                try:
                    doc = Write.open_doc(fnm=fnm, loader=loader)
                except Exception:
                    print(f"Could not open '{fnm}'")
                    # office will close and with statement is exited
                    raise BreakContext.Break

                try:
                    GUI.set_visible(is_visible=True, odoc=doc)

                    show_paragraphs(doc)
                    print(f"Word count: {count_words(doc)}")
                    show_lines(doc)
                finally:
                    Lo.close_doc(doc)

            return 0

``main()`` calls :py:meth:`.Write.open_doc` to return the opened document as an XTextDocument_ instance.
If you recall, the previous |extract_ex|_ example started with an XComponent_ instance by calling
:py:meth:`.Lo.open_doc`, and then converted it to XTextDocument_. :py:meth:`.Write.open_doc` returns the XTextDocument_ reference in one go.

``show_paragraphs()`` moves the visible on-screen cursor through the document, highlighting a paragraph at a time.
This requires two cursors – an instance of XTextViewCursor_ and a separate XParagraphCursor_.
The paragraph cursor is capable of moving through the document paragraph-by-paragraph, but it's a model cursor, so invisible to the user
looking at the document on-screen. ``show_paragraphs()`` extracts the start and end positions of each paragraph and uses them to move the view cursor, which is visible.

The code for ``show_paragraphs()``:

.. tabs::

    .. code-tab:: python

        def show_paragraphs(doc: XTextDocument) -> None:
            tvc = Write.get_view_cursor(doc)
            para_cursor = Write.get_paragraph_cursor(doc)
            para_cursor.gotoStart(False)  # go to start test; no selection

            while 1:
                para_cursor.gotoEndOfParagraph(True)  # select all of paragraph
                curr_para = para_cursor.getString()
                if len(curr_para) > 0:
                    tvc.gotoRange(para_cursor.getStart(), False)
                    tvc.gotoRange(para_cursor.getEnd(), True)

                    print(f"P<{curr_para}>")
                    Lo.delay(500)  # delay half a second

                if para_cursor.gotoNextParagraph(False) is False:
                    break

The code utilizes two Write utility functions (:py:meth:`.Write.get_view_cursor` and :py:meth:`.Write.get_paragraph_cursor`) to create the cursors.
The subsequent while loop is a common coding pattern for iterating over a text document:

.. tabs::

    .. code-tab:: python

        para_cursor.gotoStart(False)  # go to start test; no selection

        while 1:
            para_cursor.gotoEndOfParagraph(True)  # select one paragraph
            curr_para = para_cursor.getString()
            # do something with selected text range.

            if para_cursor.gotoNextParagraph(False) is False:
                break

``gotoNextParagraph()`` tries to move the cursor to the beginning of the next paragraph.

If the moves fails (i.e. when the cursor has reached the end of the document), the function returns False, and the loop terminates.

The call to ``gotoEndOfParagraph()`` at the beginning of the loop moves the cursor to the end of the paragraph and selects its text.
Since the cursor was originally at the start of the paragraph, the selection will span that paragraph.

XParagraphCursor_ and the sentence and word cursors inherit XTextCursor_, as shown in :numref:`ch05fig_model_cursor_inherit`.

..
    Figure 5

.. cssclass:: diagram invert

    .. _ch05fig_model_cursor_inherit:
    .. figure:: https://user-images.githubusercontent.com/4193389/181936175-f6086152-0231-4872-a40e-4ade46c63fa6.png
        :alt: Diagram of The Model Cursors Inheritance Hierarchy
        :figclass: align-center

        :The Model Cursors Inheritance Hierarchy.

Since all these cursors also inherit XTextRange_, they can easily access and change their text selections/positions.
In the ``show_paragraphs()`` method above, the two ends of the paragraph are obtained by calling the inherited
``XTextRange.getStart()`` and ``XTextRange.getEnd()``, and the positions are used to move the view cursor:

.. tabs::

    .. code-tab:: python

        para_cursor = Write.get_paragraph_cursor(doc)
        ...
            tvc.gotoRange(para_cursor.getStart(), False)
            tvc.gotoRange(para_cursor.getEnd(), True)

``gotoRange()`` sets the text range/position of the view cursor: the first call moves the cursor to the paragraph's starting position
without selecting anything, and the second moves it to the end position, selecting all the text in between.
Since this is a view cursor, the selection is visible on-screen, as illustrated in :numref:`ch05fig_ss_sel_para`.

..
    Figure 6

.. cssclass:: screen_shot invert

    .. _ch05fig_ss_sel_para:
    .. figure:: https://user-images.githubusercontent.com/4193389/181936346-a4a74a1a-8cce-4e16-88a9-a4a806dce53c.png
        :alt: Screen shot of A Selected Paragraph.
        :figclass: align-center

        :A Selected Paragraph.

Note that ``getStart()`` and ``getEnd()`` do not return integers but collapsed text ranges,
which is Office-lingo for a range that starts and ends at the same cursor position.

Somewhat confusingly, the XTextViewCursor_ interface inherits XTextCursor_ (as shown in :numref:`ch05fig_xtxt_view_inherit`).
This only means that XTextViewCursor supports the same character-based movement and text range operations as the model-based cursor.

..
    Figure 7

.. cssclass:: diagram invert

    .. _ch05fig_xtxt_view_inherit:
    .. figure:: https://user-images.githubusercontent.com/4193389/181936545-b0d970d4-6853-4adb-910c-d2a75150f053.png
        :alt: Diagram of The X Text View Cursor Inheritance Hierarchy.
        :figclass: align-center

        :The ``XTextViewCursor`` Inheritance Hierarchy.

5.4 Creating Cursors
====================

An XTextCursor_ is created by calling :py:meth:`.Write.get_cursor`, which can then be converted into a paragraph, sentence, or word cursor by using
:py:meth:`.Lo.qi`. For example, the :py:class:`~.selection.Selection` utility class defines :py:meth:`~.selection.Selection.get_paragraph_cursor` as:

.. tabs::

    .. code-tab:: python

        @classmethod
        def get_paragraph_cursor(cls, cursor_obj: DocOrCursor) -> XParagraphCursor:
            try:
                if Lo.qi(XTextDocument, cursor_obj) is None:
                    cursor = cursor_obj
                else:
                    cursor = cls.get_cursor(cursor_obj)
                para_cursor = Lo.qi(XParagraphCursor, cursor, True)
                return para_cursor
            except Exception as e:
                raise ParagraphCursorError(str(e)) from e

Obtaining the view cursor is a little more tricky since it's only accessible via the document's controller.

As described in :ref:`ch01_fcm_relationship`, the controller is reached via the document's model, as shown in the first three lines of
:py:meth:`.Selection.get_view_cursor`:

.. tabs::

    .. code-tab:: python

            @staticmethod
            def get_view_cursor(text_doc: XTextDocument) -> XTextViewCursor:
                try:
                    model = Lo.qi(XModel, text_doc, True)
                    xcontroller = model.getCurrentController()
                    supplier = Lo.qi(XTextViewCursorSupplier, xcontroller, True)
                    vc = supplier.getViewCursor()
                    if vc is None:
                        raise Exception("Supplier return null view cursor")
                    return vc
                except Exception as e:
                    raise ViewCursorError(str(e)) from e

The view cursor isn't directly accessible from the controller; a supplier must be queried,
even though there's only one view cursor per document.

5.4.1 Counting Words
--------------------

``count_words()`` in |walk_text|_ shows how to iterate over the document using a word cursor:

.. tabs::

    .. code-tab:: python

        def count_words(doc: XTextDocument) -> int:
            word_cursor = Write.get_word_cursor(doc)
            word_cursor.gotoStart(False)  # go to start of text

            word_count = 0
            while 1:
                word_cursor.gotoEndOfWord(True)
                curr_word = word_cursor.getString()
                if len(curr_word) > 0:
                    word_count += 1
                if word_cursor.gotoNextWord(False) is False:
                    break
            return word_count

This uses the same kind of while loop as ``show_paragraphs()`` except that the XWordCursor_ methods
``gotoEndOfWord()`` and ``gotoNextWord()`` control the iteration.
Also, there's no need for an XTextViewCursor_ instance since the selected words aren't shown on the screen.

5.4.2 Displaying Lines
----------------------

``show_lines()`` in |walk_text|_ iterates over the document highlighting a line at a time.
Don't confuse this with sentence selection because a sentence may consist of several lines on the screen.
A sentence is part of the text's organization (:abbreviation:`eg:` in terms of words, sentences, and paragraphs)
while a line is part of the document view (:abbreviation:`eg:` line, page, screen).
This means that XLineCursor_ is a view cursor, which is obtained by converting XTextViewCursor_ with :py:meth:`.Lo.qi`:

.. tabs::

    .. code-tab:: python

        line_cursor = Lo.qi(XLineCursor, tvc, True)
        tvc = Write.get_view_cursor(doc)

The line cursor has limited functionality compared to the model cursors (paragraph, sentence, word).
In particular, there's no "next' function for moving to the next line (unlike ``gotoNextParagraph()`` or ``gotoNextWord()``).
The screen cursor also lacks this ability, but the page cursor offers ``jumpToNextPage()``.

One way of getting around the absence of a 'next' operation is shown in ``show_lines()``:

.. tabs::

    .. code-tab:: python

        def show_lines(doc: XTextDocument) -> None:
            tvc = Write.get_view_cursor(doc)
            tvc.gotoStart(False)  # go to start of text

            line_cursor = Lo.qi(XLineCursor, tvc, True)
            have_text = True
            while have_text is True:
                line_cursor.gotoStartOfLine(False)
                line_cursor.gotoEndOfLine(True)
                print(f"L<{tvc.getString()}>")
                Lo.delay(500)  # delay half a second
                tvc.collapseToEnd()
                have_text = tvc.goRight(1, True)

The view cursor is manipulated using the XTextViewCursor_ object and the XLineCursor_ line cursor.
This is possible since the two references point to the same on-screen cursor. Either one can move it around the display.

Inside the loop, ``XLineCursor's`` ``gotoStartOfLine()`` and ``gotoEndOfLine()`` highlight a single line.
Then the XTextViewCursor_ instance deselects the line, by moving the cursor to the end of the selection with ``collapseToEnd()``.
At the end of the loop, ``goRight()`` tries to move the cursor one character to the right.
If ``goRight()`` succeeds then the cursor is shifted one position to the first character of the next line. When the loop repeats, this line will be selected.
If ``goRight()`` fails, then there are no more characters to be read from the document, and the loop finishes.

5.5 Creating a Document
=======================

All the examples so far have involved the manipulation of an existing document.
The |hello_save|_ example creates a new text document, containing two short paragraphs, and saves it as "hello.odt".
The main() function is:


.. tabs::

    .. code-tab:: python

        def main() -> int:

            with Lo.Loader(Lo.ConnectSocket()) as loader:

                doc = Write.create_doc(loader)

                GUI.set_visible(is_visible=True, odoc=doc)

                cursor = Write.get_cursor(doc)
                cursor.gotoEnd(False)  # make sure at end of doc before appending
                Write.append_para(cursor=cursor, text="Hello LibreOffice.\n")
                Lo.delay(1_000)  # Slow things down so user can see

                Write.append_para(cursor=cursor, text="How are you?")
                Lo.delay(2_000)
                Write.save_doc(text_doc=doc, fnm="hello.odt")
                Lo.close_doc(doc)

            return 0

:py:meth:`.Write.create_doc` calls :py:meth:`.Lo.create_doc` with the text document service name (the ``Lo.DocTypeStr.WRITER`` enum value is ``swriter``).
Office creates a TextDocument_ service with an XComponent_ interface, which is cast to the XTextDocument_ interface, and returned:

.. tabs::

    .. code-tab:: python

        # simplified version of Write.create_doc
        @staticmethod
        def create_doc(loader: XComponentLoader) -> XTextDocument:
            doc = Lo.qi(
                XTextDocument,
                Lo.create_doc(doc_type=Lo.DocTypeStr.WRITER, loader=loader),
                True,
            )
            return doc

Text documents are saved using :py:meth:`.Write.save_doc` that calls :py:meth:`.Lo.save_doc` which was described in :ref:`ch02_save_doc`.
``save_doc()`` examines the filename's extension to determine its type.
The known extensions include ``doc``, ``docx``, ``rtf``, ``odt``, ``pdf``, and ``txt``.

Back in |hello_save|_, a cursor is needed before text can be added; one is created by calling :py:meth:`.Write.get_cursor`.

The call to ``XTextCursor.gotoEnd()`` isn't really necessary because the new cursor is pointing to an empty document so is already at its end.
It's included to emphasize the assumption by :py:meth:`.Write.append_para` (and other ``Write.appendXXX()`` functions) that the cursor is
positioned at the end of the document before new text is added.

:py:meth:`.Write.append_para` calls :py:meth:`.Write.append` methods:

.. tabs::

    .. code-tab:: python

        # simplified version of Write.append_para
        @classmethod
        def append_para(cls, cursor: XTextCursor, text: str) -> None:
            cls.append(cursor=cursor, text=text)
            cls.append(cursor=cursor, ctl_char=Write.ControlCharacter.PARAGRAPH_BREAK)

The :py:meth:`~.Write.append` name is utilized several times in Write via it overloads:

    - ``append(cursor: XTextCursor, text: str)``
    - ``append(cursor: XTextCursor, ctl_char: ControlCharacter)``
    - ``append(cursor: XTextCursor, text_content: com.sun.star.text.XTextContent)``

``append(cursor: XTextCursor, text: str)`` appends text using ``XTextCursor.setString()`` to add the user-supplied string.

``append(cursor: XTextCursor, ctl_char: ControlCharacter)`` uses ``XTextCursor.insertControlCharacter()``.
After the addition of the text or character, the cursor is moved to the end of the document.

``append(cursor: XTextCursor, text_content: com.sun.star.text.XTextContent)`` appends an object
that is a sub-class of XTextContent_

``ControlCharacter`` is an enumeration of API ControlCharacter_.
Thanks to ooouno_ library that among other things automatically creates enums for LibreOffice Constants.

``Write.ControlCharacter`` is an alias for convenience.

.. tabs::

    .. code-tab:: python

        from ooo.dyn.text.control_character import ControlCharacterEnum

        class Write(Selection):
            ControlCharacter = ControlCharacterEnum

:py:meth:`.Selection.get_position` (inherited by Write) gets the current position if the cursor from the start of the document.
This method is not full optimized and may not be robust on large files.

Office deals with this size issue by using XTextRange_ instances, which encapsulate text ranges and
positions. :py:meth:`.Selection.get_position` returns an integer because its easier to understand when you're first learning to program with Office.
It's better style to use and compare XTextRange_ objects rather than integer positions, an approach demonstrated in the next section.

.. _ch05_txt_cursors:

5.6 Using and Comparing Text Cursors
====================================

|speak_text|_ example utilizes the third-party library text-to-speech_ to convert text into speech.
The inner workings aren't relevant here, so are hidden inside a single method ``speak()``.

|speak_text|_ employs two text cursors: a paragraph cursor that iterates over the paragraphs in the document,
and a sentence cursor that iterates over all the sentences in the current paragraph and passes each sentence to ``speak()``.
text-to-speech_  is capable of speaking long or short sequences of text, but |speak_text|_ processes a sentence at a time since this sounds more natural when spoken.

The crucial function in |speak_text|_ is ``speak_sentences()``:

.. tabs::

    .. code-tab:: python

        def speak_sentences(doc: XTextDocument) -> None:
            tvc = Write.get_view_cursor(doc)
            para_cursor = Write.get_paragraph_cursor(doc)
            para_cursor.gotoStart(False)  # go to start test; no selection

            while 1:
                para_cursor.gotoEndOfParagraph(True)  # select all of paragraph
                end_para = para_cursor.getEnd()
                curr_para_str = para_cursor.getString()
                print(f"P<{curr_para_str}>")

                if len(curr_para_str) > 0:
                    # set sentence cursor pointing to start of this paragraph
                    cursor = para_cursor.getText().createTextCursorByRange(para_cursor.getStart())
                    sc = Lo.qi(XSentenceCursor, cursor)
                    sc.gotoStartOfSentence(False)
                    while 1:
                        sc.gotoEndOfSentence(True)  # select all of sentence

                        # move the text view cursor to highlight the current sentence
                        tvc.gotoRange(sc.getStart(), False)
                        tvc.gotoRange(sc.getEnd(), True)
                        curr_sent_str = strip_non_word_chars(sc.getString())
                        print(f"S<{curr_sent_str}>")
                        if len(curr_sent_str) > 0:
                            speak(
                                curr_sent_str,
                            )
                        if Write.compare_cursor_ends(sc.getEnd(), end_para) >= Write.CompareEnum.EQUAL:
                            print("Sentence cursor passed end of current paragraph")
                            break

                        if sc.gotoNextSentence(False) is False:
                            print("# No next sentence")
                            break

                if para_cursor.gotoNextParagraph(False) is False:
                    break

``speak_sentences()`` comprises two nested loops: the outer loop iterates through the paragraphs, and the inner loop through the sentences in the current paragraph.

The sentence cursor is created like so:

.. tabs::

    .. code-tab:: python

        cursor = para_cursor.getText().createTextCursorByRange(para_cursor.getStart())

        sc = Lo.qi(XSentenceCursor, cursor)

The XText_ reference is returned by ``para_cursor.getText()``, and a text cursor is created.

``createTextCursorByRange()`` allows the start position of the cursor to be specified. The text cursor is converted into a sentence cursor with :py:meth:`.Lo.qi`.

The tricky aspect of this code is the meaning of ``para_cursor.getText()`` which is the XText_ object that ``para_cursor`` utilizes.
This is not a single paragraph but the entire text document.
Remember that the paragraph cursor is created with: ``para_cursor = Write.get_paragraph_cursor(doc)`` This corresponds to:

| ``xtext = doc.getText()``
| ``text_cursor = xtext.createTextCursor()``
| ``para_cursor = Lo.qi(XParagraphCursor, text_cursor)``

Both the paragraph and sentence cursors refer to the entire text document.
This means that it is not possible to code the inner loop using the coding pattern from before.That would result in something like the following:

.. tabs::

    .. code-tab:: python

        # set sentence cursor to point to start of this paragraph
        cursor = para_cursor.getText().createTextCursorByRange(para_cursor.getStart())
        sc = Lo.qi(XSentenceCursor, cursor)
        sc.gotoStartOfSentence(False) # goto start

        while 1:
            sc.gotoEndOfSentence(True) #select 1 sentence

            if sc.gotoNextSentence(False) is False:
                break

.. note::

    To further confuse matters, a ``XText`` object does not always correspond to the entire text document.
    For example, a text frame (e.g. like this one) can return an ``XText`` object for the text only inside the frame.

The problem with the above code fragment is that ``XSentenceCursor.gotoNextSentence()`` will keep moving to the next sentence until it reaches the end of the text document.
This is not the desired behavior – what is needed for the loop to terminate when the last sentence of the current paragraph has been processed.

We need to compare text ranges, in this case the end of the current sentence with the end of the current paragraph.
This capability is handled by the XTextRangeCompare_ interface. A comparer object is created at the beginning of ``speak_sentence()``,
initialized to compare ranges that can span the entire document:

.. tabs::

    .. code-tab:: python

        if Write.compare_cursor_ends(sc.getEnd(), end_para) >= Write.CompareEnum.EQUAL:
            print("Sentence cursor passed end of current paragraph")
            break

:py:meth:`.Selection.compare_cursor_ends` compares cursors ends and returns an enum value.

If the sentence ends after the end of the paragraph then ``compare_cursor_ends()`` returns a value greater or equal to ``Write.CompareEnum.EQUAL``, and the inner loop terminates.

Since there's no string being created by the comparer, there's no way that the instantiating can fail due to the size of the text.


5.7 Inserting/Changing Text in a Document
=========================================

The |shuffle_words|_ example searches a document and changes the words it encounters.
:numref:`ch05fig_word_shuffle` shows the program output. Words longer than three characters are scrambled.

..
    Figure 8

.. cssclass:: screen_shot invert

    .. _ch05fig_word_shuffle:
    .. figure:: https://user-images.githubusercontent.com/4193389/184255719-a3f8a75c-dba3-41b0-bcb4-631fb7b92c0a.png
        :alt: screenshot, Shuffling of Words
        :figclass: align-center

        :Shuffling of Words.

A word shuffle is applied to every word of four letters or more, but only involves the random exchange of the middle letters without changing the first and last characters.

The ``apply_shuffle()`` function which iterates through the words in the input file is similar to ``count_words()`` in |walk_text|_.
One difference is the use of ``XText.insertString()``:

.. tabs::

    .. code-tab:: python

        def apply_shuffle(doc: XTextDocument, delay: int, visible: bool) -> None:
            doc_text = doc.getText()
            if visible:
                cursor = Write.get_view_cursor(doc)
            else:
                cursor = Write.get_cursor(doc)

            word_cursor = Write.get_word_cursor(doc)
            word_cursor.gotoStart(False)  # go to start of text

            while True:
                word_cursor.gotoNextWord(True)

                # move the text view cursor, and highlight the current word
                cursor.gotoRange(word_cursor.getStart(), False)
                cursor.gotoRange(word_cursor.getEnd(), True)
                curr_word = word_cursor.getString()

                # get whitespace padding amounts
                c_len = len(curr_word)
                curr_word = curr_word.lstrip()
                l_pad = c_len - len(curr_word)  # left whitespace padding amount
                curr_word = curr_word.rstrip()
                r_pad = c_len - len(curr_word) - l_pad  # right whitespace padding ammount
                if len(curr_word) > 0:
                    pad_l = " " * l_pad  # recreate left padding
                    pad_r = " " * r_pad  # recreate right padding
                    Lo.delay(delay)
                    mid_shuff = mid_shuffle(curr_word)
                    doc_text.insertString(word_cursor, f"{pad_l}{mid_shuff}{pad_r}", True)

                if word_cursor.gotoNextWord(False) is False:
                    break

            word_cursor.gotoStart(False)  # go to start of text
            cursor.gotoStart(False)

``insertString()`` is located in XSimpleText_:

.. tabs::

    .. code-tab:: python

        def insertString(xRange: XTextRange, aString: str, bAbsorb: bool) -> None:

    .. code-tab:: java

        void insertString(XTextRange xRange, String aString, boolean bAbsorb)

The string s is inserted at the cursor's text range position.
If ``bAbsorb`` is true then the string replaces the current selection (which is the case in ``apply_shuffle()``).

``mid_shuffle()`` shuffles the string in ``curr_word``, returning a new word. It doesn't use the Office API, so no explanation here.


5.8 Treating a Document as Paragraphs and Text Portions
=======================================================

Another approach for moving around a document involves the XEnumerationAccess_ interface which treats the document as a series of Paragraph text contents.

XEnumerationAccess_ is an interface in the Text service, which means that an XText_ reference can be converted into it by using :py:meth:`.Lo.qi`.
These relationships are shown in :numref:`ch05fig_text_service`.

..
    Figure 9

.. cssclass:: diagram invert

    .. _ch05fig_text_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/184417050-ebb948ad-6a4f-4bdd-bc32-cbe90b82b1ab.png
        :alt: Diagram of Text Service and its Interfaces
        :figclass: align-center

        :The Text Service and its Interfaces.

The following code fragment utilizes this technique:

.. tabs::

    .. code-tab:: python

        xtext = doc.getText()
        enum_access = Lo.qi(XEnumerationAccess, xtext);

XEnumerationAccess_ contains a single method, ``createEnumeration()`` which creates an enumerator (an instance of XEnumeration_).
Each element returned from this iterator is a Paragraph text content:

.. tabs::

    .. code-tab:: python

        # create enumerator over the document text
        enum_access = Lo.qi(XEnumerationAccess, doc.getText())
        text_enum = enum_access.createEnumeration()

        while text_enum.hasMoreElements():
            text_con = Lo.qi(XTextContent, text_enum.nextElement())
            # use the Paragraph text content (text_con) in some way...

Paragraph doesn't support its own interface (i.e. there's no ``XParagraph``), so :py:meth:`.Lo.qi` is used to access its XTextContent_ interface,
which belongs to the TextContent_ subclass. The hierarchy is shown in :numref:`ch05fig_text_context_hierarchy`.

..
    Figure 10

.. cssclass:: diagram invert

    .. _ch05fig_text_context_hierarchy:
    .. figure:: https://user-images.githubusercontent.com/4193389/184431023-3a34228a-1e07-4d25-ab3f-a00fc5030085.png
        :alt: Diagram of The Paragraph Text Content Hierarchy
        :figclass: align-center

        :The Paragraph Text Content Hierarchy.

Iterating over a document to access Paragraph text contents doesn't seem much different from iterating over a document using a paragraph cursor,
except that the Paragraph service offers a more structured view of a paragraph.

In particular, you can use another XEnumerationAccess_ instance to iterate over a single paragraph, viewing it as a sequence of text portions.

The following code illustrates the notion, using the ``text_con`` text content from the previous piece of code:

.. tabs::

    .. code-tab:: python

        if not Info.support_service(text_con, "com.sun.star.text.TextTable"):
            para_enum = Write.get_enumeration(text_con)
            while para_enum.hasMoreElements():
                txt_range = Lo.qi(XTextRange, para_enum.nextElement())
                # use the text portion (txt_range) in some way...

The TextTable_ service is a subclass of Paragraph, and cannot be enumerated.

Therefore, the paragraph enumerator is surrounded with an if-test to skip a paragraph if it's really a table.

The paragraph enumerator returns text portions, represented by the TextPortion_ service.
TextPortion_ contains a lot of useful properties which describe the paragraph, but it doesn't have its own interface (such as ``XTextPortion``).
However, TextPortion_ inherits the TextRange_ service, so :py:meth:`.Lo.qi` can be used to obtain its XTextRange_ interface.
This hierarchy is shown in :numref:`ch05fig_text_portion_hierarchy`.

..
    Figure 11

.. cssclass:: diagram invert

    .. _ch05fig_text_portion_hierarchy:
    .. figure:: https://user-images.githubusercontent.com/4193389/184432816-452d8189-652d-4bb8-947e-6147e7754545.png
        :alt: Diagram of The TextPortion Service Hierarchy
        :figclass: align-center

        :The TextPortion Service Hierarchy.

TextPortion_ includes a ``TextPortionType`` property which identifies the type of the portion.
Other properties access different kinds of portion data, such as a text field or footnote.

For instance, the following prints the text portion type and the string inside the ``txt_range`` text portion (``txt_range`` comes from the previous code fragment):

.. tabs::

    .. code-tab:: python

        print(f'  {Props.get_property(txt_range, "TextPortionType")} = "{txt_range.getString()}"')

These code fragments are combined together in the |show_book|_ example.

More details on enumerators and text portions are given in the Developers Guide at https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Iterating_over_Text

5.9 Appending Documents Together
================================

If you need to write a large multi-part document (e.g. a thesis with chapters, appendices, contents page, and an index)
then you should utilize a master document, which acts as a repository of links to documents representing the component parts.
You can find out about master documents in Chapter 13 of the Writers Guide, at https://wiki.documentfoundation.org/Documentation/Publications.

However, the complexity of master documents isn't always needed.
Often the aim is simply to append one document to the end of another.
In that case, the XDocumentInsertable_ interface, and its ``insertDocumentFromURL()`` method is more suitable.

|docs_append|_ example uses ``XDocumentInsertable.insertDocumentFromURL()``.
A list of filenames is read from the command line; the first file is opened, and the other files appended to it by ``append_text_files()``:

.. tabs::

    .. code-tab:: python

        # part of Docs Append example
        def append_text_files(doc: XTextDocument, *args: str) -> None:
            cursor = Write.get_cursor(doc)
            for arg in args:
                try:
                    cursor.gotoEnd(False)
                    print(f"Appending {arg}")
                    inserter = Lo.qi(XDocumentInsertable, cursor)
                    if inserter is None:
                        print("Document inserter could not be created")
                    else:
                        inserter.insertDocumentFromURL(FileIO.fnm_to_url(arg), ())
                except Exception as e:
                    print(f"Could not append {arg} : {e}")

A XDocumentInsertable_ instance is obtained by converting the text cursor with :py:meth:`.Lo.qi`.

``XDocumentInsertable.insertDocumentFromURL()`` requires two arguments – the URL of the file that's being appended, and an empty property value array.

.. |txt_java| replace:: TextDocuments.java
.. _txt_java: https://api.libreoffice.org/examples/DevelopersGuide/examples.html#Text

.. |write_guide| replace:: Writer Guide
.. _write_guide: https://documentation.libreoffice.org/en/english-documentation/

.. |extract_ex| replace:: Extract Writer Text
.. _extract_ex: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_doc_print_console

.. |walk_text| replace:: Walk Text
.. _walk_text: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_walk_text

.. |hello_save| replace:: Hello Save
.. _hello_save: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_hello_save

.. |speak_text| replace:: Speak Text
.. _speak_text: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_speak


.. |shuffle_words| replace:: Shuffle Words
.. _shuffle_words: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_shuffle

.. |show_book| replace:: Show Book
.. _show_book: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_show_book

.. |docs_append| replace:: Docs Append
.. _docs_append: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_docs_append

.. _text-to-speech: https://pypi.org/project/text-to-speech/

.. _ControlCharacter: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1text_1_1ControlCharacter.html
.. _GenericTextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1GenericTextDocument.html
.. _OfficeDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1OfficeDocument.html
.. _TextContent: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextContent.html
.. _TextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextDocument.html
.. _TextPortion: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextPortion.html
.. _TextRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextRange.html
.. _TextTable: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextTable.html
.. _XComponent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XComponent.html
.. _XDocumentInsertable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1document_1_1XDocumentInsertable.html
.. _XEnumeration: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XEnumeration.html
.. _XEnumerationAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XEnumerationAccess.html
.. _XLineCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1view_1_1XLineCursor.html
.. _XParagraphCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XParagraphCursor.html
.. _XSentenceCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XSentenceCursor.html
.. _XServiceInfo: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XServiceInfo.html
.. _XSimpleText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XSimpleText.html
.. _XStyleFamiliesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1style_1_1XStyleFamiliesSupplier.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html
.. _XTextContent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextContent.html
.. _XTextCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextCursor.html
.. _XTextDocument: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextDocument.html
.. _XTextRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextRange.html
.. _XTextRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextRange.html
.. _XTextRangeCompare: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextRangeCompare.html
.. _XTextViewCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextViewCursor.html
.. _XWordCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XWordCursor.html